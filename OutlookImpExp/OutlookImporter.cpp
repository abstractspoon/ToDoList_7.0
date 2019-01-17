// OutlookImporter.cpp: implementation of the COutlookImporter class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "OutlookImpExp.h"
#include "OutlookImporter.h"
#include "OutlookImportDlg.h"

#include "..\shared\Localizer.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

COutlookImporter::COutlookImporter()
{

}

COutlookImporter::~COutlookImporter()
{

}

void COutlookImporter::SetLocalizer(ITransText* pTT)
{
	CLocalizer::Initialize(pTT);
}

bool COutlookImporter::Import(LPCTSTR /*szSrcFilePath*/, ITaskList* pDestTaskFile, BOOL bSilent, IPreferences* pPrefs, LPCTSTR szKey)
{
	COutlookImportDlg dlg;

	return (!bSilent && dlg.ImportTasks(pDestTaskFile, pPrefs, szKey) != FALSE);
}
