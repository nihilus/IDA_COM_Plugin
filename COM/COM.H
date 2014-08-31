#ifndef _COM_H_
#define _COM_H_

extern void ShowError(const char *message);
extern BOOL CheckInterface(LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr, const char *szInterfaceName);
extern void ProcessFunction(FUNCDESC *pFuncDesc, LPTYPEINFO pITypeInfo, DWORD pFunction, const char *pszMungedName, const char *pszCommentAnsi);

#endif
