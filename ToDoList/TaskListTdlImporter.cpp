// TaskListCsvImporter.cpp: implementation of the CTaskListTdlImporter class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TaskListTdlImporter.h"
#include "TDLTasklistImportDlg.h"
#include "Taskfile.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CTaskListTdlImporter::CTaskListTdlImporter()
{

}

CTaskListTdlImporter::~CTaskListTdlImporter()
{

}

bool CTaskListTdlImporter::Import(LPCTSTR szSrcFilePath, ITaskList* pDestTaskFile, BOOL bSilent, IPreferences* /*pPrefs*/, LPCTSTR /*szKey*/)
{
	ITaskList8* pTasks8 = GetITLInterface<ITaskList8>(pDestTaskFile, IID_TASKLIST8);
	
	CTDLTasklistImportDlg dialog(szSrcFilePath);

	if (!bSilent)
	{
		if (dialog.DoModal() != IDOK)
			return false;

		// else
		return (dialog.GetSelectedTasks(pTasks8) > 0);
	}

	// else 
	CTaskFile tasks;

	if (!tasks.Load(szSrcFilePath))
		return false;

	tasks.CopyTasksTo(pTasks8);
	return (pTasks8->GetTaskCount() > 0);
}

