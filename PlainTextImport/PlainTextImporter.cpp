// PlainTextImporter.cpp: implementation of the CPlainTextImporter class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PlainTextImporter.h"
#include "optionsdlg.h"

#include "..\shared\misc.h"

#include "..\Interfaces\ITasklist.h"
#include "..\Interfaces\IPreferences.h"

#include "..\3rdParty\stdiofileex.h"

#include <time.h>
#include <unknwn.h>

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CPlainTextImporter::CPlainTextImporter()
{

}

CPlainTextImporter::~CPlainTextImporter()
{

}

void CPlainTextImporter::SetLocalizer(ITransText* /*pTT*/)
{
	//CLocalizer::Initialize(pTT);
}

bool CPlainTextImporter::InitConsts(BOOL bSilent, IPreferences* pPrefs, LPCTSTR szKey)
{
	CString sKey(szKey);
	sKey += _T("\\PlainText");

	WANTPROJECT = pPrefs->GetProfileInt(szKey, _T("IncludeProject"), FALSE);
	INDENT = pPrefs->GetProfileString(szKey, _T("Indent"), _T("  "));

	if (!bSilent)
	{
		COptionsDlg dlg(TRUE, WANTPROJECT, INDENT);

		if (dlg.DoModal() != IDOK)
			return false;
		
		INDENT = dlg.GetIndent();
		WANTPROJECT = dlg.GetWantProject();

		pPrefs->WriteProfileInt(szKey, _T("IncludeProject"), WANTPROJECT);
		pPrefs->WriteProfileString(szKey, _T("Indent"), INDENT);
	}

	return true;
}

bool CPlainTextImporter::Import(LPCTSTR szSrcFilePath, ITaskList* pDestTaskFile, BOOL bSilent, IPreferences* pPrefs, LPCTSTR szKey)
{
	if (!InitConsts(bSilent, pPrefs, szKey))
		return false;

	CStdioFileEx file;

	if (!file.Open(szSrcFilePath, CFile::modeRead))
		return false;

	// else
	ITaskList4* pITL4 = GetITLInterface<ITaskList4>(pDestTaskFile, IID_TASKLIST4);

	// the first line can be the project name
	if (WANTPROJECT)
	{
		CString sProjectName;
		file.ReadString(sProjectName);
		Misc::Trim(sProjectName);

		pITL4->SetProjectName(sProjectName);
	}

	// what follows are the tasks, indented to express subtasks
	int nLastDepth = 0;
	HTASKITEM hLastTask = NULL;

	ROOTDEPTH = -1; // gets set to the first task's depth

	CString sLine;
	
	while (file.ReadString(sLine)) 
	{
		CString sTitle, sComments;

		if (!GetTitleComments(sLine, sTitle, sComments))
			continue;

		// find the appropriate parent fro this task
		HTASKITEM hParent = NULL;
		int nDepth = GetDepth(sLine);

		if (nDepth == nLastDepth) // sibling
			hParent = hLastTask ? pITL4->GetTaskParent(hLastTask) : NULL;

		else if (nDepth > nLastDepth) // child
			hParent = hLastTask;

		else if (hLastTask) // we need to work up the tree
		{
			hParent = pITL4->GetTaskParent(hLastTask);

			while (hParent && nDepth < nLastDepth)
			{
				hParent = pITL4->GetTaskParent(hParent);
				nLastDepth--;
			}
		}
		
		HTASKITEM hTask = pITL4->NewTask(sTitle, hParent);

		if (!sComments.IsEmpty())
			pITL4->SetTaskComments(hTask, sComments);

		// update state
		hLastTask = hTask;
		nLastDepth = nDepth;
	}

	return true;
}

int CPlainTextImporter::GetDepth(const CString& sLine)
{
	if (INDENT.IsEmpty() || sLine.IsEmpty())
		return 0;

	// else
	int nDepth = 0;
	
	if (INDENT == "\t")
	{
		while (nDepth < sLine.GetLength())
		{
			if (sLine[nDepth] == '\t')
				nDepth++;
			else
				break;
		}
	}
	else // one or more spaces
	{
		int nPos = 0;

		while (nPos < sLine.GetLength())
		{
			if (sLine.Find(INDENT, nPos) == nPos)
				nDepth++;
			else
				break;

			// next
			nPos = nDepth * INDENT.GetLength();
		}
	}

	// set root depth if not set 
	if (ROOTDEPTH == -1)
		ROOTDEPTH = nDepth;

	// and take allowance for it
	nDepth -= ROOTDEPTH;

	return nDepth;
}

BOOL CPlainTextImporter::GetTitleComments(const CString& sLine, 
										  CString& sTitle, CString& sComments)
{
	int nDelim = sLine.Find(_T("|")); 
	
	if (nDelim != -1)
	{
		sTitle = sLine.Left(nDelim);
		sComments = sLine.Mid(nDelim + 1);
	}
	else
	{
		sTitle = sLine;
		sComments.Empty();
	}

	// cleanup
	Misc::Trim(sTitle);
	Misc::Trim(sComments);

	return sTitle.GetLength();
}
