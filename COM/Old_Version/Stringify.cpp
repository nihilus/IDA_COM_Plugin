// This code was taken from t_browse.cpp by Sean Baxter
//
// It generates an IDL-style and C-style declaration of a COM interface method
//
// See http://ript.net/~spec/typeinfo/
// for the full details
//
// I extended the code a little bit (parameter names, C-style,
// declaration) and deleted some code not used. Dieter Spaar

#include <windows.h>
#include <atlbase.h>
#include <sstream>

#include "stringify.h"

#ifndef UNICODE
  #error "This file must be compiled as unicode"
#endif

// internal declarations

std::string stringifyParameterAttributes(PARAMDESC* paramDesc);
std::string stringifyCustomType(HREFTYPE refType, ITypeInfo* pti);
std::string stringifyTypeDesc(TYPEDESC* typeDesc, ITypeInfo* pti);
std::string stringifyVarDesc(VARDESC* varDesc, ITypeInfo* pti);
std::string stringifyFunctionArgument(ELEMDESC* elemDesc, ITypeInfo* pti);
std::string stringifyCOMMethod(FUNCDESC* funcDesc, ITypeInfo* pti);

// usefull wrappers

struct EVeryBadThing { };

struct CComTypeAttr 
{
	TYPEATTR* _typeAttr;
	CComPtr<ITypeInfo> _pTypeInfo;
	operator TYPEATTR*() { return _typeAttr; }
	TYPEATTR* operator->() { return _typeAttr; }
	explicit CComTypeAttr(ITypeInfo* ti) throw(EVeryBadThing) : _pTypeInfo(ti) 
	{
		HRESULT hr(_pTypeInfo->GetTypeAttr(&_typeAttr));
		if(hr) throw EVeryBadThing();
	}
	~CComTypeAttr() { _pTypeInfo->ReleaseTypeAttr(_typeAttr); }
};

struct CComFuncDesc {
	FUNCDESC* _funcDesc;
	operator FUNCDESC*() { return _funcDesc; }
	FUNCDESC* operator->() { return _funcDesc; }
	CComPtr<ITypeInfo> _pTypeInfo;
	CComFuncDesc(ITypeInfo* ti, int index) throw(EVeryBadThing) : _pTypeInfo(ti) 
	{
		HRESULT hr(_pTypeInfo->GetFuncDesc(index, &_funcDesc));
		if(hr) throw EVeryBadThing();
	}
	~CComFuncDesc() { _pTypeInfo->ReleaseFuncDesc(_funcDesc); }
};

struct CComVarDesc {
	VARDESC* _varDesc;
	operator VARDESC*() { return _varDesc; }
	VARDESC* operator->() { return _varDesc; }
	CComPtr<ITypeInfo> _pTypeInfo;
	CComVarDesc(ITypeInfo* ti, int index) throw(EVeryBadThing) : _pTypeInfo(ti) 
	{
		HRESULT hr(_pTypeInfo->GetVarDesc(index, &_varDesc));
		if(hr) throw EVeryBadThing();
	}
	~CComVarDesc() { _pTypeInfo->ReleaseVarDesc(_varDesc); }
};

struct CComLibAttr {
	TLIBATTR* _libAttr;
	operator TLIBATTR*() { return _libAttr; }
	TLIBATTR* operator->() { return _libAttr; }
	CComPtr<ITypeLib> _pTypeLib;
	explicit CComLibAttr(ITypeLib* tlb) throw(EVeryBadThing) : _pTypeLib(tlb) 
	{
		HRESULT hr(_pTypeLib->GetLibAttr(&_libAttr));
		if(hr) throw EVeryBadThing();
	}
	~CComLibAttr() { _pTypeLib->ReleaseTLibAttr(_libAttr); }
};

// globals

BOOL g_bIDLStyle = TRUE;

std::string stringifyParameterAttributes(PARAMDESC* paramDesc) 
{
	USHORT paramFlags = paramDesc->wParamFlags;
	int numFlags(0);

	for(DWORD bit(1); bit <= PARAMFLAG_FHASDEFAULT; bit<<=1) 
		numFlags += (paramFlags & bit) ? 1 : 0;

	if(!numFlags) 
		return "";

	std::ostringstream oss;
	oss<< '[';

	if(paramFlags & PARAMFLAG_FIN) 
	{ 
		oss<< "in"; 
		if(--numFlags) 
			oss<< ", "; 
	}
	if(paramFlags & PARAMFLAG_FOUT) 
	{ 
		oss<< "out"; 
		if(--numFlags) 
			oss<< ", "; 
	}
	if(paramFlags & PARAMFLAG_FLCID) 
	{ 
		oss<< "lcid"; 
		if(--numFlags) 
			oss<< ", "; 
	}
	if(paramFlags & PARAMFLAG_FRETVAL) 
	{ 
		oss<< "retval"; 
		if(--numFlags) 
			oss<< ", "; 
	}
	if(paramFlags & PARAMFLAG_FOPT) 
	{ 
		oss<< "optional"; 
		if(--numFlags) 
			oss<< ", "; 
	}

	if(paramFlags & PARAMFLAG_FHASDEFAULT) 
	{
		oss<< "defaultvalue";
		if(paramDesc->pparamdescex) 
		{
			oss<< '(';

			PARAMDESCEX& paramDescEx = *(paramDesc->pparamdescex);
			VARIANT defVal();
			CComBSTR bstrDefValue;
			CComVariant variant;
			HRESULT hr(VariantChangeType(&variant, &paramDescEx.varDefaultValue, 0, VT_BSTR));

			if(hr) 
				oss<< "???)";
			else 
			{
				char ansiDefValue[MAX_PATH];
				WideCharToMultiByte(CP_ACP, 0, variant.bstrVal, SysStringLen(variant.bstrVal) + 1, ansiDefValue, MAX_PATH, 0, 0);

				if(paramDescEx.varDefaultValue.vt == VT_BSTR)
					oss<< '\"'<< ansiDefValue<< '\"'<< ')';
				else 
					oss<< ansiDefValue<< ')';
			}
		}
	}
	oss<< ']';

	return oss.str();
}


std::string stringifyCustomType(HREFTYPE refType, ITypeInfo* pti) 
{
	CComPtr<ITypeInfo> pTypeInfo(pti);
	CComPtr<ITypeInfo> pCustTypeInfo;
	HRESULT hr(pTypeInfo->GetRefTypeInfo(refType, &pCustTypeInfo));

	if(hr) 
		return "UnknownCustomType";

	CComBSTR bstrType;
	hr = pCustTypeInfo->GetDocumentation(-1, &bstrType, 0, 0, 0);
	if(hr) 
		return "UnknownCustomType";

	char ansiType[MAX_PATH];
	WideCharToMultiByte(CP_ACP, 0, bstrType, bstrType.Length() + 1, ansiType, MAX_PATH, 0, 0);

	return ansiType;
}

std::string stringifyTypeDesc(TYPEDESC* typeDesc, ITypeInfo* pti) 
{
	std::ostringstream oss;
	if(typeDesc->vt == VT_PTR) 
	{
		oss<< stringifyTypeDesc(typeDesc->lptdesc, pti)<< '*';
		return oss.str();
	}

	if(typeDesc->vt == VT_SAFEARRAY) 
	{
		oss<< "SAFEARRAY("
			<< stringifyTypeDesc(typeDesc->lptdesc, pti)<< ')';
		return oss.str();
	}

	if(typeDesc->vt == VT_CARRAY) 
	{
		oss<< stringifyTypeDesc(&typeDesc->lpadesc->tdescElem, pti);

		for(int dim(0); typeDesc->lpadesc->cDims; ++dim) 
			oss<< '['<< typeDesc->lpadesc->rgbounds[dim].cElements<< ']';

		return oss.str();
	}

	if(typeDesc->vt == VT_USERDEFINED) 
	{
		if(g_bIDLStyle)
			oss<< stringifyCustomType(typeDesc->hreftype, pti);
		else
			oss<< "int"; // default type
		return oss.str();
	}
	
	switch(typeDesc->vt) 
	{
		// VARIANT compatible types
		case VT_I2: 
			return "short";
		case VT_I4: 
			return "long";
		case VT_R4: 
			return "float";
		case VT_R8: 
			return "double";
		case VT_CY: 
			return "CY";
		case VT_DATE: 
			return "DATE";
		case VT_BSTR: 
			return "BSTR";
		case VT_DISPATCH: 
			return "IDispatch*";
		case VT_ERROR: 
			return "SCODE";
		case VT_BOOL: 
			return "VARIANT_BOOL";
		case VT_VARIANT: 
			if(g_bIDLStyle)
				return "VARIANT";
			else
				return "VARIANTARG";
		case VT_UNKNOWN: 
			return "IUnknown*";
		case VT_UI1: 
			return "BYTE";
		// VARIANTARG compatible types
		case VT_DECIMAL: 
			return "DECIMAL";
		case VT_I1: 
			return "char";
		case VT_UI2: 
			return "USHORT";
		case VT_UI4: 
			return "ULONG";
		case VT_I8: 
			return "__int64";
		case VT_UI8: 
			return "unsigned __int64";
		case VT_INT: 
			return "int";
		case VT_UINT: 
			return "UINT";
		case VT_HRESULT: 
			return "HRESULT";
		case VT_VOID: 
			return "void";
		case VT_LPSTR: 
			return "char*";
		case VT_LPWSTR: 
			return "wchar_t*";
	}
	return "BIG ERROR!";
}

// for enums and structure members, not used yet

std::string stringifyVarDesc(VARDESC* varDesc, ITypeInfo* pti) 
{
	CComPtr<ITypeInfo> pTypeInfo(pti);
	std::ostringstream oss;
	if(varDesc->varkind == VAR_CONST) 
		oss<< "const ";

	oss<< stringifyTypeDesc(&varDesc->elemdescVar.tdesc, pTypeInfo);
	CComBSTR bstrName;
	HRESULT hr(pTypeInfo->GetDocumentation(varDesc->memid, &bstrName, 0, 0, 0));
	if(hr) 
		return "UnknownName";

	char ansiName[MAX_PATH];
	WideCharToMultiByte(CP_ACP, 0, bstrName, bstrName.Length() + 1, ansiName, MAX_PATH, 0, 0);

	oss<< ' '<< ansiName;
	if(varDesc->varkind != VAR_CONST) 
		return oss.str();

	oss<< " = ";
	CComVariant variant;
	hr = VariantChangeType(&variant, varDesc->lpvarValue, 0, VT_BSTR);
	if(hr) 
		oss<< "???";
	else 
	{
		WideCharToMultiByte(CP_ACP, 0, variant.bstrVal, SysStringLen(variant.bstrVal) + 1, ansiName, MAX_PATH, 0, 0);
		oss<< ansiName;
	}

	return oss.str();
}


std::string stringifyFunctionArgument(ELEMDESC* elemDesc, ITypeInfo* pti) 
{
	CComPtr<ITypeInfo> pTypeInfo(pti);
	std::ostringstream oss;

	if(g_bIDLStyle)
		oss<< stringifyParameterAttributes(&elemDesc->paramdesc);

	if(oss.str().size()) 
		oss<< ' ';
	oss<< stringifyTypeDesc(&elemDesc->tdesc, pti);

	return oss.str();
}

std::string stringifyCOMMethod(FUNCDESC* funcDesc, ITypeInfo* pti) 
{
	CComPtr<ITypeInfo> pTypeInfo(pti);
	std::ostringstream oss;
	const int MAX_NAMES = 256;
	CComBSTR name[MAX_NAMES];
	UINT cNames;

	// get parameter names

	HRESULT hr(pTypeInfo->GetNames(funcDesc->memid, reinterpret_cast<BSTR*>(&name), MAX_NAMES - 1, &cNames));
	if(hr)
		cNames = 0;
	else
	{
  		// fix for 'rhs' problem (see tblodl.cpp of MS Visual C++ OLEVIEW sample for more details)

		if (cNames != MAX_NAMES - 1 && (int)cNames < funcDesc->cParams + 1)
		{
			name[cNames] = ::SysAllocString(OLESTR("rhs")) ;
			cNames++;
		}
	}

	// function attribute (only IDL-style)

	if(g_bIDLStyle)
	{
		if(funcDesc->funckind == FUNC_DISPATCH)	
			oss<< "[id("<< (int)funcDesc->memid<< ')';
		else 
			oss<< "[VOffset("<< funcDesc->oVft<< ')';

		switch(funcDesc->invkind) 
		{
			case INVOKE_PROPERTYGET: 
				oss<< ", propget] "; 
				break;
			case INVOKE_PROPERTYPUT: 
				oss<< ", propput] "; 
				break;
			case INVOKE_PROPERTYPUTREF: 
				oss<< ", propputref] "; 
				break;
			case INVOKE_FUNC: oss<< "] "; 
				break;
		}
	}

	// function type

	oss<< stringifyTypeDesc(&funcDesc->elemdescFunc.tdesc, pTypeInfo);

	// calling convention (only C-style)

	if(!g_bIDLStyle)
	{
		CComTypeAttr typeAttr(pTypeInfo);

		if (typeAttr->typekind != TKIND_DISPATCH)
		{   // Write calling convention
			switch(funcDesc->callconv)
			{
				case CC_CDECL:
					oss<< " _cdecl";           
					break ;
				/*
				case CC_MSCPASCAL:  
					oss<< " _mspascal";        
					break ;
				*/
				case CC_PASCAL:     
					oss<< " _pascal";          
					break ;
				case CC_MACPASCAL:  
					oss<< " _macpascal";       
					break ;
				case CC_STDCALL :   
					oss<< " _stdcall";          
					break ;
				/*
				case CC_RESERVED:   
					oss<< " _reserved";         
					break ;
				*/
				case CC_SYSCALL:    
					oss<< " _syscall";          
					break ;
				case CC_MPWCDECL:   
					oss<< " _mpwcdecl";         
					break ;
				case CC_MPWPASCAL:  
					oss<< " _mpwpascal";        
					break ;
			}
		}
	}

	// function name

	CComBSTR bstrName;
	pTypeInfo->GetDocumentation(funcDesc->memid, &bstrName, 0, 0, 0);

	char ansiName[MAX_PATH];
	WideCharToMultiByte(CP_ACP, 0, bstrName, bstrName.Length() + 1, ansiName, MAX_PATH, 0, 0);

	oss<< ' '<< ansiName<< '(';


	// first parameter is hidden "this" pointer (C-style only)

	if(!g_bIDLStyle)
	{
		oss<< "void *pThis";
		if(funcDesc->cParams) 
			oss<< ", ";
	}

	// parameters

	for(int curParam(0); curParam < funcDesc->cParams; ++curParam) 
	{
		oss<< stringifyFunctionArgument(&funcDesc->lprgelemdescParam[curParam], pTypeInfo);

		// function parameter name (index 0 is the function name)

		if(cNames)
		{
			if(curParam + 1 < cNames)
			{
				WideCharToMultiByte(CP_ACP, 0, name[curParam + 1], name[curParam + 1].Length() + 1, ansiName, MAX_PATH, 0, 0);
				oss<< ' '<< ansiName;
			}
			else
				oss<< " arg" << (curParam + 1);
		}

		if(curParam < funcDesc->cParams - 1) 
			oss<< ", ";
	}

	// trailer

	if(g_bIDLStyle)
		oss<< ')';
	else
		oss<< ");";

	return oss.str();
}

void c_stringifyCOMMethod(FUNCDESC* funcDesc, ITypeInfo* pti, char *szBuf, int nMaxLen, BOOL bIDLStyle /* = TRUE */)
{
	// The original code produced an IDL-style declaration.
	// IDA requires C-style. This probably will not work
	// for all combinations yet.

	g_bIDLStyle = bIDLStyle;

	strncpy(szBuf, stringifyCOMMethod(funcDesc, pti).c_str(), nMaxLen - 1);
	szBuf[nMaxLen - 1] = 0;
}
