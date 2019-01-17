// OutlookImporter.h: interface for the COutlookImporter class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_OUTLOOKIMPORTER_H__C35841C7_32A6_4705_A061_24A85C3A0223__INCLUDED_)
#define AFX_OUTLOOKIMPORTER_H__C35841C7_32A6_4705_A061_24A85C3A0223__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "..\Interfaces\IImportExport.h"

class ITransText;

class COutlookImporter : public IImportTasklist  
{
public:
	COutlookImporter();
	virtual ~COutlookImporter();

	void Release() { delete this; }
	void SetLocalizer(ITransText* pTT);

	LPCTSTR GetMenuText() const { return _T("Microsoft Outlook"); }
	LPCTSTR GetFileFilter() const { return NULL; }
	LPCTSTR GetFileExtension() const { return NULL; }
	
	bool Import(LPCTSTR szSrcFilePath, ITaskList* pDestTaskFile, BOOL bSilent, IPreferences* pPrefs, LPCTSTR szKey);

};

#endif // !defined(AFX_OUTLOOKIMPORTER_H__C35841C7_32A6_4705_A061_24A85C3A0223__INCLUDED_)
