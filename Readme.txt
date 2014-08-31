COM Plugin V 1.3, September 2007

Author: Dieter Spaar (spaar@mirider.augusta.de)

The plugin tries to extract the symbol information from
the typelibrary of the COM component. It will then set the 
function names of interface methods and their parameters, and 
finally add a comment with the MIDL-style declaration of the 
interface method.

To get the addresses of the interface methods, the plugin has 
to create the interface class first. This means that the COM 
component will be loaded, so be aware of malicious code. 
Loading of the COM component may fail if the component is 
not already registered or if the component checks if it is 
licensed to run on the computer. Also, it can happen for 
some interfaces that retrieving the function addresses
does not work. The plugin will display an error messages 
if something went wrong.

To compile the project (Microsoft Visual C++ 6 or Visual Studio 2003), 
the directory COM has to be put inside the PLUGIN directory of the 
IDA Pro SDK. I use the file "cpy1.bat" in the plugin build to copy 
the plugin to the IDA plugin directory, this file has to be adjusted
and also the "Custom Build Step" Project Setting and its "Command Line" 
and "Outputs" entry.

Please note that the code has been adjusted for the IDA Pro SDK 5.1.
There is also a subdirectory "Old_Version" which contains the old
version of the COM plugin for IDA Pro 4.3 in case someone really 
needs it. However I do not plan to support this version further,
any possible updates will be for the current version only.


Credits: Matt Pietrek, Sean Baxter, Ilfak Guilfanov
         Please also have a look at the source code
         for more details.

History:         

  9. Jul 2001:     added directories for different IDA Pro versions
                   which contain the compiled plugin. There is now
                   the plugin for IDA Pro 4.16 and 4.17.

 27. Oct 2002:     added compiled plugin for IDA Pro version 4.30.

 05. Sep 2007:     adjusted the code for IDA Pro 5.1          