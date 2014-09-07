//
// IDA Pro plugin to generate symbolic names from a COM 
// component.
//
// V 1.3 for IDA Pro SDK 5.1
//
// Author: Dieter Spaar (spaar@mirider.augusta.de)
//
// The plugin tries to extract the symbol information from
// the typelibrary of the COM component. To get the addresses of 
// the interface methods, the plugin has to create the interface 
// class first. This means that the COM component will be loaded, 
// so be aware of malicious code. Loading of the COM component
// may fail if the component is not already registered or if the 
// component checks if it is licensed to run on the computer.
//
// To register a component, you can use REGSVR32:
//
//   REGSVR32 <component file name>
//
// To unregister the component again:
//
//   REGSVR32 /U <component file name>
//
// 
// For links to some more information about COM and 
// typelibraries see the header of Stringify.cpp and
// CoClassSyms.cpp.
//
// Also, the Microsoft Visual C++ OLEVIEW sample is very
// informative:
//
// http://msdn.microsoft.com/library/devprods/vs6/visualc/vcsample/_sample_mfc_oleview.htm
//
// The basis for this plugin is the PDB plugin sample from 
// the IDA Pro SDK.
//
// Many thanks to Ilfak Guilfanov who provided great support and help
// and also took the time to test the plugin and point out what should
// be improved.
//

/*//////////////////////////////////////////////////////////////////////
                           Necessary Includes
//////////////////////////////////////////////////////////////////////*/

#include <windows.h>
#include <ole2.h>
#include <imagehlp.h>

// IDA SDK includes

#include <ida.hpp>
#include <idp.hpp>
#include <name.hpp>
#include <loader.hpp>
#include <diskio.hpp>
#include <typeinf.hpp>
#include <auto.hpp>
#include <bytes.hpp>

#include "CoClassSyms.h"
#include "stringify.h"
#include "COM.H"

// set to 1 to show some debug messages

#define SHOW_DBG		0

// internal declarations

extern "C" plugin_t PLUGIN;

void my_warning(const char *message,...);
BOOL RelocateVA( PVOID address, DWORD &rva );
__declspec (dllexport) BOOL APIENTRY DllMain(HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved);
BOOL IsFunctionStart(ea_t ea);
BOOL CleanupIfInsideOpcode(ea_t ea);

// globals

unsigned long g_delta = 0; // relocation amount

unsigned long g_nFunctions = 0; // counter for renamed functions
unsigned long g_nNoFunctions = 0; // counter for names, but no functions yet

BOOL g_bGUIVersion = FALSE;
BOOL g_bSilent = FALSE;

inline ssize_t idaapi ExtraGet(ea_t ea, int what, char *buf, size_t bufsize) { return netnode(ea).supstr(what, buf, bufsize); }


/*----------------------------------------------------------------------
Similar to IDA's warning() function, but simulates "SILENT" for GUI version
----------------------------------------------------------------------*/

void my_warning(const char *message,...)
{
	va_list va;
	va_start(va, message);
	if(g_bGUIVersion)
	{
		if(g_bSilent)
		{
			// log to message window

			vmsg(message, va);
			msg("\n");
		}
		else
		{
			// show dialog box with

			if(askbuttons_cv("~O~k", "~S~ilent", "~C~ancel", 1, message, va) == 0)
			{
				// confirm "SILENT" mode

				if(askyn_c(1, "This will turn off all error messages for the plugin. "
							  "The error messages will only appear in the message window. "
				              "Do you want to turn off the error messages ?") == 1 )
					g_bSilent = TRUE;
			}
		}
	}
	else
	{
		// just the default for text version

		vwarning(message, va);
	}
	va_end(va);
}

/*----------------------------------------------------------------------
Show error message, called from	ShowHresultError() in CoClassSyms
----------------------------------------------------------------------*/

void ShowError(const char *message)
{
	my_warning("%s", message);
}

/*----------------------------------------------------------------------
relocate linear (virtual) address (bases address is imagebase)
----------------------------------------------------------------------*/

BOOL RelocateVA(PVOID address, DWORD &relocateAddr)
{
	MEMORY_BASIC_INFORMATION mbi;

	// Tricky way to get the containing module from a linear address	
	VirtualQuery( address, &mbi, sizeof(mbi) );

	// "AllocationBase" is the same as an HMODULE
	LPVOID hModule = (LPVOID)mbi.AllocationBase;

	// Use IMAGEHLP API to get a pointer to the PE header.
	PIMAGE_NT_HEADERS pNtHeaders = ImageNtHeader(hModule);
	if ( !pNtHeaders )
		return FALSE;
		
	// Calculate relocated virtual address

	relocateAddr = (DWORD)address - (DWORD)hModule + pNtHeaders->OptionalHeader.ImageBase + g_delta;

	return TRUE;
}

/*----------------------------------------------------------------------
The plugin is a DLL, this is the standard DLL entry point.
Not really needed for this plugin, just for testing
----------------------------------------------------------------------*/

#if 1 // can be set to 0

__declspec (dllexport) BOOL APIENTRY DllMain(HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
	switch(ul_reason_for_call) 
	{
		case DLL_PROCESS_ATTACH:

		// determine TXT/GUI version

		g_bGUIVersion = (callui(ui_get_hwnd).vptr != NULL);
#if SHOW_DBG
		my_warning("DLL_PROCESS_ATTACH.");
#endif
		break;

		case DLL_THREAD_ATTACH:
		break;

		case DLL_THREAD_DETACH:
		break;

		case DLL_PROCESS_DETACH:
#if SHOW_DBG
		my_warning("DLL_PROCESS_DETACH.");
#endif
		break;
	}
	return TRUE;
}

#endif

/*----------------------------------------------------------------------
Thanks to Ilfak, here is the code that shows how apply_cdecl() works
internally. This can be used to customize apply_cdecl() if required
Attention: not sure if this is still true for IDA 5.1 !
----------------------------------------------------------------------*/

#if 1 // set to 1 to use the code instead of the built-in version of apply_cdecl()

bool ida_export apply_cdecl(ea_t ea, const char *decl)
{
  char *name;
  type_t *type;
  p_list *fields;
  bool ok = parse_type(decl, &name, &type, &fields, PT_SIL);
  if ( ok ) ok = apply_type(ea, type, fields);
  qfree(name);
  qfree(type);
  qfree(fields);
  return ok;
}

#endif

/*----------------------------------------------------------------------
Check if address is the start of a function
----------------------------------------------------------------------*/

BOOL IsFunctionStart(ea_t ea)
{
	func_t *fn = get_func(ea);
	return (fn && fn->startEA == ea);
}

/*----------------------------------------------------------------------
Check if address is inside an opcode, if not, try to clean up
Return TRUE if everything is OK.
----------------------------------------------------------------------*/

BOOL CleanupIfInsideOpcode(ea_t ea)
{
	// check if we are in the middle of an object

    if(!isTail(getFlags(ea)))
		return TRUE;

	// make unknown, expand

	do_unknown(get_item_head(ea), true);

	// wait till auatoanalysis is done, cause expanding
	// to "unknown" happends in the background if second
	// parameter of do_unknown() is true

	autoWait();

	// check if there already exists a function
	// if so, truncate it

	func_t nfn;
	if(find_func_bounds(ea, &nfn, FIND_FUNC_DEFINE) == FIND_FUNC_EXIST)
	{
		if(func_setend(ea, ea))
			return FALSE;

		// truncating a function may cause autoanalysis
		// to start again. Wait till autoanalysis is done to
		// get best results

		autoWait();
	}

	// create function, return result

	return add_func(ea, BADADDR) == 1;
}

/*----------------------------------------------------------------------
Called from EnumTypeInfoMembers() in CoClassSyms.cpp for every interface
found. Returns FALSE if interface should not be processed
----------------------------------------------------------------------*/

BOOL CheckInterface(LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr, const char *szInterfaceName)
{
	// IDispatch interface must be a dual interface, otherwise the method address is
	// the same for all methodes of that interface

	if(pTypeAttr->typekind == TKIND_DISPATCH && (pTypeAttr->wTypeFlags & TYPEFLAG_FDUAL) == 0)
	{
		my_warning("IDispatch interface %s is not dual.\n", szInterfaceName);
		return FALSE;
	}

	return TRUE;
}

/*----------------------------------------------------------------------
Called from EnumTypeInfoMembers() in CoClassSyms.cpp for every function 
found
----------------------------------------------------------------------*/

void ProcessFunction(FUNCDESC *pFuncDesc, LPTYPEINFO pITypeInfo, DWORD pFunction, const char *pszMungedName, const char *pszCommentAnsi)
{
	// Convert the virtual address to a relative virtual address (RVA)
	
	unsigned long rva;
	if(!RelocateVA((PVOID)pFunction, rva))
	{
		my_warning("RelocateVA() for address %X failed", rva);
		return;
	}

	ea_t ea = rva;

	// check if ea is on a valid opcode address
	// if not,try to make correct code

	if(!CleanupIfInsideOpcode(ea))
	{
		g_nNoFunctions++;

		msg("%08X: function address not valid yet: %s\n", ea, pszMungedName);

		return;
	}

	// Tell IDA kernel: create the function

	add_func(ea, BADADDR);

	// Tell IDA kernel: rename the function

	set_name(ea, pszMungedName, SN_NOWARN);

	// check if the function is properly set

	BOOL bFunctionOK = IsFunctionStart(ea);

	// most of the time a function is not properly set cause IDA things
	// that it is part of a very big function, so try to truncate that function

	if(!bFunctionOK)
	{
		// truncate function just above new function

		if(func_setend(ea, ea))
		{
			// truncating a function may cause autoanalysis
			// to start again. Wait till autoanalysis is done to
			// get best results

			autoWait();

			// Tell IDA kernel: create the function

			add_func(ea, BADADDR);

			// check again

			bFunctionOK = IsFunctionStart(ea);
		}
	}

	// get function declaration (IDL style)

	char szDecl[512];
	c_stringifyCOMMethod(pFuncDesc, pITypeInfo, szDecl, sizeof(szDecl));

	// set comment (two lines, IDL declaration and help text)

	if(pszCommentAnsi[0] || szDecl[0])
	{
		// build comment string

		char szBuf[1024];
		szBuf[0] = 0;
		const char *pszCmtLine1 = NULL;
		const char *pszCmtLine2 = NULL;

		// IDL declaration

		if(szDecl[0])
		{
			qstrncat(szBuf, szDecl, sizeof(szBuf)); 
			pszCmtLine1 = szDecl;
		}

		// set comment in second line

		if(pszCommentAnsi[0])
		{
			if(szBuf[0])
			{
				qstrncat(szBuf, "\n", sizeof(szBuf));
				pszCmtLine2 = pszCommentAnsi;
			}
			else
				pszCmtLine1 = pszCommentAnsi;

			qstrncat(szBuf, pszCommentAnsi, sizeof(szBuf)); 
		}

		// set function comment
		// if it not a function or a has a different start address, 
		// add standard comment

		if(bFunctionOK)
		{
			// tell IDA to set the function comment

			func_t *fn = get_func(ea);
			if(fn)
				set_func_cmt(fn, szBuf, 0);

			g_nFunctions++;
		}
		else
		{
			// check if comment is alread set
			// currently add_long_cmt() appends new comments.

			BOOL bCmdEqual = FALSE;
			char szCmt[512];
			int nLenCmt;

			// we have at most two comment lines, ExtraGet() 
			// returns a single line
			// Also ExtraGet() returns comments including leading "; "

			if(pszCmtLine1)
			{
				nLenCmt = ExtraGet(ea, E_PREV, szCmt, sizeof(szCmt)); // get first comment line before this address
				if(nLenCmt && strcmp(pszCmtLine1, szCmt + 2) == 0)
					bCmdEqual = TRUE;
			}
			if(pszCmtLine2 && bCmdEqual)
			{
				nLenCmt = ExtraGet(ea, E_PREV + 1, szCmt, sizeof(szCmt)); // get second comment line before this address
				if(!(nLenCmt && strcmp(pszCmtLine2, szCmt + 2) == 0))
					bCmdEqual = FALSE;
			}

			if(!bCmdEqual)
			{
				// tell IDA to add a comment (anterior)

				add_long_cmt(ea, 1, szBuf);
			}

			g_nNoFunctions++;
		}
	}

	// tell IDA to apply the function declaration
	// Note: function name does not care, but declaration has to
	// be correct C-style (e.g. has a trailing semicolon)
	// sample: int __cdecl sub(LPVOID lpMem1,DWORD dwBytes1);

	if(bFunctionOK)
	{
		// get function declaration (C style)

		c_stringifyCOMMethod(pFuncDesc, pITypeInfo, szDecl, sizeof(szDecl), FALSE);

		// tell IDA to use the declaration to ameliorate the disassembly
		
		if(!apply_cdecl(ea, szDecl))
		{
			// If using the "01234567: xxxx" style for warnings,
			// the GUI version permits to jump to an address by  
			// double clicking on the message line.

			msg("%08X: cannot set type for %s\n", ea, szDecl);
		}
	}
	else
		msg("%08X: no function: %s\n", ea, pszMungedName);
}

//--------------------------------------------------------------------------
//
//      The plugin method
//
//      This is the main function of plugin.
//
//      It will be called when the user selects the plugin.
//
//              arg - the input argument, it can be specified in
//                    plugins.cfg file. The default is zero.
//
//

void plugin_main(int)
{
#if SHOW_DBG
	my_warning("plugin_main() called.");
#endif
	
	// unload plugin if plugin debug flag set (-z00000020 command line parameter),
	// this is very helpful for debugging, cause IDA will always unload the
	// plugin after plugin_main() has been call. So the plugin can be replaced
	// without leaving IDA.

	if(debug & IDA_DEBUG_PLUGIN)
		PLUGIN.flags |= PLUGIN_UNL;

	// Warn the user that this plugin executes code in the COM component

	if(askyn_c(1, "Please be aware that this plugin will execute code in the "
				  "COM component. "
				  "Do you want to apply the plugin ?") <= 0 )
		return;

	// The results will be better if initial autoanalysis is finished,
	// warn the user if necessary

	if(!autoIsOk())
	{
		if(askyn_c(-1, "The analysis has not finished yet. "
					   "The results would be better if you wait till its finished. "
					   "Do you want to apply the plugin now?") <= 0 )
		return;
	}

	// initialize 

	if(!InitProcessTypeLib())
		return;

	// rest counters

	g_nFunctions = 0;
	g_nNoFunctions = 0;

    // Calculate the relocation amount (usually 0 since we load
    // the files at their ImageBase)
	//
	// Not equal to 0 if "Manual load" is selected when IDA load 
	// a file
    {
		g_delta = 0;
		netnode penode("$ PE header");
		ea_t loaded_base = penode.altval(-2);
		IMAGE_NT_HEADERS pe;
		int nPeSize;
		nPeSize = penode.valobj(&pe, sizeof(pe));

		if ( nPeSize != 0 && loaded_base != 0 )
			g_delta = loaded_base - pe.OptionalHeader.ImageBase;

#if SHOW_DBG
		my_warning("delta = 0x%X.", g_delta);
#endif
    }  

    // Get the input file name fom IDA and try to guess the file location
    // If failed, ask the user
	char *input;
    char szInputPath[_MAX_PATH];
	get_input_file_path(szInputPath, sizeof(szInputPath));
	input = szInputPath;
    if ( !qfileexist(input) && (input = askfile_c(false, input, "Please specify the input file")) == NULL )
        return;

	// tell IDA to display a message box

    show_wait_box("Getting symbols");

	// its a good idea to catch exception in case our plugin crashes ;-)

    try
    {
		// process typelibray, see CoClassSyms.cpp

		ProcessTypeLib(input);
    }
    catch(...)
    {
		my_warning("Exception while processing type library");
    }

	// tell IDA to remove the message box again

    hide_wait_box();

	// show result, turn off "SILENT" mode to display the message box

	int save_batch = batch;
	batch = 0; // turn off "SILENT" mode 

	if(g_nNoFunctions)
		info("%d function names set.\n%d names set for code not recognized as function yet.", g_nFunctions, g_nNoFunctions);
	else
		info("%d function names set.", g_nFunctions);

	batch = save_batch;

	// cleanup

	ExitProcessTypeLib();
}


//--------------------------------------------------------------------------
//
//      initialize plugin
//
//      IDA will call this function only once.
//      If this function returns PLGUIN_SKIP, IDA will never load it again.
//      If this function returns PLUGIN_OK, IDA will unload the plugin but
//      remember that the plugin agreed to work with the database.
//      The plugin will be loaded again if the user invokes it by
//      pressing the hotkey or selecting it from the menu.
//      After the second load the plugin will stay on memory.
//
// Some notes: to distinguish between TXT and GUI version of IDA,
// callui(ui_get_hwnd) or getvcl() can be used.	See kernwin.hpp for
// detail.

int init(void)
{
#if SHOW_DBG
	my_warning("init() called.");
#endif

	// determine TXT/GUI version

	g_bGUIVersion = (callui(ui_get_hwnd).vptr != NULL);

	if (inf.filetype != f_PE) 
		return PLUGIN_SKIP; // only for PE files

	return PLUGIN_OK;
}

//--------------------------------------------------------------------------
//      terminate
//      usually this callback is empty
//
//      IDA will call this function when the user asks to exit.
//      This function won't be called in the case of emergency exits.

void term(void)
{
#if SHOW_DBG
	my_warning("term() called.");
#endif
}

char comment[] = "Load symbol information from a COM typelibray";

char help[] =
"COM symbol loader\n"
"\n"
"This module allows you to load symbol information\n"
"from a COM typelibrary.\n\n"
"Please note that the component has to be registered\n"
"and may also be licensed first to access the typelibrary.\n";


//--------------------------------------------------------------------------
// This is the preferred name of the plugin module in the menu system
// The preferred name may be overriden in plugins.cfg file

char wanted_name[] = "Load COM symbols";


// This is the preferred hotkey for the plugin module
// The preferred hotkey may be overriden in plugins.cfg file
// Note: IDA won't tell you if the hotkey is not correct
//       It will just disable the hotkey.

char wanted_hotkey[] = "Ctrl-F11";


//--------------------------------------------------------------------------
//
//      PLUGIN DESCRIPTION BLOCK
//
//--------------------------------------------------------------------------

extern "C" plugin_t PLUGIN =
{
  IDP_INTERFACE_VERSION,
  0,                    // plugin flags
  init,                 // initialize

  term,                 // terminate. this pointer may be NULL.

  plugin_main,          // invoke plugin

  comment,              // long comment about the plugin
                        // it could appear in the status line
                        // or as a hint

  help,                 // multiline help about the plugin

  wanted_name,          // the preferred short name of the plugin
  wanted_hotkey         // the preferred hotkey to run the plugin
};
