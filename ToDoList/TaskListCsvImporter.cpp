// TaskListCsvImporter.cpp: implementation of the CTaskListCsvImporter class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TaskListCsvImporter.h"
#include "tdlschemadef.h"
#include "resource.h"
#include "recurringtaskedit.h"
#include "TDLCsvImportExportDlg.h"

#include <locale.h>

#include "..\shared\timehelper.h"
#include "..\shared\enstring.h"
#include "..\shared\misc.h"
#include "..\shared\filemisc.h"
#include "..\shared\Preferences.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

LPCTSTR SPACE = _T(" ");
LPCTSTR ENDL = _T("\n");
LPCTSTR ONEDBLQUOTE = _T("\"");
LPCTSTR TWODBLQUOTE = _T("\"\"");

////////////////////////////////////////////////////////

CTaskListCsvImporter::CTaskListCsvImporter()
{

}

CTaskListCsvImporter::~CTaskListCsvImporter()
{

}

bool CTaskListCsvImporter::InitConsts(LPCTSTR szSrcFilePath, BOOL bSilent, 
									  IPreferences* pPrefs, LPCTSTR szKey)
{
	// we go straight for the application preferences
	szKey = _T("Preferences");

	if (pPrefs->GetProfileInt(_T("Preferences"), _T("UseSpaceIndents"), TRUE))
		INDENT = CString(' ', pPrefs->GetProfileInt(_T("Preferences"), _T("TextIndent"), 2));
	else
		INDENT = "  "; // don't use tabs - excel strips them off

	CTDLCsvImportExportDlg dialog(szSrcFilePath, pPrefs, szKey);

	if (!bSilent)
	{
		while (true)
		{
			if (dialog.DoModal() != IDOK)
				return false;

			// valid mapping must include title
			if (!dialog.GetColumnMapping(m_aColumnMapping) ||
				(AttributeIndex(TDCA_TASKNAME) == -1))
			{
				AfxMessageBox(IDS_CSV_MUSTMAPTASKTITLE);
				// try again
			}
			else
			{
				break; // all good
			}
		} 
	}
	else
	{
		VERIFY (dialog.GetColumnMapping(m_aColumnMapping));
		ASSERT (AttributeIndex(TDCA_TASKNAME) != -1);
	}

	DELIM = dialog.GetDelimiter();

	return true;
}

bool CTaskListCsvImporter::Import(LPCTSTR szSrcFilePath, ITaskList* pDestTaskFile, BOOL bSilent, IPreferences* pPrefs, LPCTSTR szKey)
{
	if (!InitConsts(szSrcFilePath, bSilent, pPrefs, szKey))
		return false;

	// Load csv
	CStringArray aLines;
	FileMisc::LoadFileLines(szSrcFilePath, aLines);

	// remove header line
	aLines.RemoveAt(0);

	ITaskList14* pTasks = GetITLInterface<ITaskList14>(pDestTaskFile, IID_TASKLIST12);
	ASSERT(pTasks);

	AddCustomAttributes(pTasks);

	BOOL bSomeFailed = FALSE;

	for (int nLine = 0; nLine < aLines.GetSize(); nLine++)
	{
	    CString sLine = GetLine(aLines, nLine);

		if (!ImportTask(pTasks, sLine))
			bSomeFailed = TRUE;
	}

	if (!bSilent && bSomeFailed)
		AfxMessageBox(IDS_IMPORT_SOMEFAILED);

	return true;
}

CString CTaskListCsvImporter::GetLine(const CStringArray& aLines, int& nLine)
{
    CEnString sLine(aLines[nLine]);
	
    // if the line contains an odd number of double-quotes
    // then assume we're in a comment and keep appending lines
    // until we hit the closing double-quote
    if (sLine.GetCharacterCount('\"') % 2)
    {
		while (++nLine < aLines.GetSize())
		{
			CEnString sNextLine(aLines[nLine]);
			sLine += "\n" + sNextLine;
			
			if (sNextLine.GetCharacterCount('\"') % 2)
				break;
		}
    }
	
    return sLine;
}

void CTaskListCsvImporter::GetTaskAndParentIDs(const CStringArray& sValues, DWORD& dwTaskID, DWORD& dwParentID) const
{
	dwTaskID = dwParentID = 0;

	int nCol = AttributeIndex(TDCA_ID);

	if (nCol != -1)
		dwTaskID = _ttoi(sValues[nCol]);

	nCol = AttributeIndex(TDCA_PARENTID);

	if (nCol != -1)
		dwParentID = _ttoi(sValues[nCol]);
}

CString CTaskListCsvImporter::GetTaskTitle(const CStringArray& sValues) const
{
	int nCol = AttributeIndex(TDCA_TASKNAME);

	if (nCol != -1)
		return sValues[nCol];

	// else
	static CString sEmpty;
	return sEmpty;
}

BOOL CTaskListCsvImporter::ImportTask(ITaskList14* pTasks, const CString& sLine) const
{
	// must have at least one field
	CStringArray aValues;
	int nNumVals = Misc::Split(sLine, aValues, DELIM, TRUE);

	ASSERT(nNumVals >= 1);

	if (nNumVals < 1)
		return FALSE; // then we can report this when we've finished importing

	// must have a title
	CString sTitle = GetTaskTitle(aValues);
	ASSERT(!sTitle.IsEmpty());

	if (sTitle.IsEmpty())
		return FALSE;

	// get taskID and ParentID
	DWORD dwTaskID, dwParentID;
	GetTaskAndParentIDs(aValues, dwTaskID, dwParentID);

	// find the parent task
	HTASKITEM hParent = pTasks->FindTask(dwParentID);

	// create task
	HTASKITEM hTask = pTasks->NewTask(sTitle, hParent, dwTaskID);

	AddAttributeToTask(pTasks, hTask, TDCA_CREATEDBY, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_CATEGORY, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_STATUS, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_EXTERNALID, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_ALLOCBY, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_ALLOCTO, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_VERSION, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_FILEREF, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_DEPENDENCY, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_COMMENTS, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_PRIORITY, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_RISK, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_FLAG, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_TIMEEST, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_TIMESPENT, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_COST, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_STARTDATE, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_DUEDATE, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_DONEDATE, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_LASTMOD, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_CREATIONDATE, aValues);
	AddAttributeToTask(pTasks, hTask, TDCA_TAGS, aValues);

	// custom attributes have to be handled differently because
	// multiple unique attributes may exist
	AddCustomAttributesToTask(pTasks, hTask, aValues);

	return TRUE;
}

int CTaskListCsvImporter::AttributeIndex(TDC_ATTRIBUTE attrib) const
{
	int nAttrib = m_aColumnMapping.GetSize();

	while (nAttrib--)
	{
		const CSVCOLUMNMAPPING& col = m_aColumnMapping[nAttrib];

		if (col.nTDCAttrib == attrib && !col.sColumnName.IsEmpty())
			break;
	}

	return nAttrib;
}

int CTaskListCsvImporter::GetDepth(const CString& sLine)
{
	if (INDENT.IsEmpty() || sLine.IsEmpty())
		return 0;

	// else
	int nDepth = 0;
	
	if (INDENT == _T("\t"))
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

BOOL CTaskListCsvImporter::String2Date(const CString& sDate, time64_t& t64, BOOL bAndTime)
{
	return (CDateHelper::DecodeDate(sDate, t64, bAndTime));
}

BOOL CTaskListCsvImporter::GetCustomAttribIDAndLabel(const CSVCOLUMNMAPPING& col, CString& sCustID, CString& sCustLabel)
{
	if ((col.nTDCAttrib == TDCA_CUSTOMATTRIB) && !col.sColumnName.IsEmpty())
	{
		sCustLabel = col.sColumnName;
		
		// extract custom ID in brackets
		int nCustIDStart = col.sColumnName.Find('(');
		
		if (nCustIDStart > 0)
		{
			sCustLabel = col.sColumnName.Left(nCustIDStart++);
			
			int nCustIDEnd = col.sColumnName.Find(')', nCustIDStart);
			
			if (nCustIDEnd == -1)
				return FALSE;
			
			sCustID = col.sColumnName.Mid(nCustIDStart, nCustIDEnd - nCustIDStart);
		}
		else // create custom ID
		{
			sCustID.Format(_T("_%s_ID_"), col.sColumnName);
			sCustID.Replace(' ', '_');
		}

		return (!sCustID.IsEmpty() && !sCustLabel.IsEmpty());
	}

	// else
	return FALSE;
}

void CTaskListCsvImporter::AddCustomAttributes(ITaskList14* pTasks) const
{
	int nAttrib = m_aColumnMapping.GetSize();
	
	while (nAttrib--)
	{
		const CSVCOLUMNMAPPING& col = m_aColumnMapping[nAttrib];
		CString sCustID, sCustLabel;

		if (GetCustomAttribIDAndLabel(col, sCustID, sCustLabel))
			pTasks->AddCustomAttribute(sCustID, sCustLabel);
	}
}

void CTaskListCsvImporter::AddCustomAttributesToTask(ITaskList14* pTasks, HTASKITEM hTask, const CStringArray& aValues) const
{
	int nAttrib = min(aValues.GetSize(), m_aColumnMapping.GetSize());
	
	while (nAttrib--)
	{
		const CSVCOLUMNMAPPING& col = m_aColumnMapping[nAttrib];
		CString sCustID, sDummy;
		
		if (GetCustomAttribIDAndLabel(col, sCustID, sDummy))
			pTasks->SetTaskCustomAttributeData(hTask, sCustID, aValues[nAttrib]);
	}
}


void CTaskListCsvImporter::AddAttributeToTask(ITaskList14* pTasks, HTASKITEM hTask, TDC_ATTRIBUTE nAttrib, const CStringArray& aColValues) const
{
	ASSERT(nAttrib < TDCA_CUSTOMATTRIB_FIRST);

	int nCol = AttributeIndex(nAttrib);
	
	if ((nCol == -1) || (nCol >= aColValues.GetSize()))
		return;

	const CString& sValue = Misc::GetItem(aColValues, nCol);
	time64_t t64 = 0;
	
// ---------------------------------------------------------------
#define SETTASKATTRIBARRAY(fn) \
	{ CStringArray aItems; if (Misc::Split(sValue, aItems, '+')) \
	for (int nItem = 0; nItem < aItems.GetSize(); nItem++) pTasks->fn(hTask, aItems[nItem]); }
// ---------------------------------------------------------------

	switch(nAttrib)
	{
	case TDCA_CREATEDBY: 
		pTasks->SetTaskCreatedBy(hTask, sValue);
		break;

	case TDCA_STATUS: 
		pTasks->SetTaskStatus(hTask, sValue);
		break;

	case TDCA_EXTERNALID: 
		pTasks->SetTaskExternalID(hTask, sValue);
		break;

	case TDCA_ALLOCBY: 
		pTasks->SetTaskAllocatedBy(hTask, sValue);
		break;

	case TDCA_VERSION: 
		pTasks->SetTaskVersion(hTask, sValue);
		break;

	case TDCA_COMMENTS: 
		pTasks->SetTaskComments(hTask, sValue);
		break;

	case TDCA_STARTDATE: 
		if (String2Date(sValue, t64, TRUE))
			pTasks->SetTaskStartDate64(hTask, t64);
		break;

	case TDCA_DUEDATE: 
		if (String2Date(sValue, t64, TRUE))
			pTasks->SetTaskDueDate64(hTask, t64);
		break;

	case TDCA_DONEDATE: 
		if (String2Date(sValue, t64, TRUE))
			pTasks->SetTaskDoneDate64(hTask, t64);
		break;

	case TDCA_LASTMOD: 
		if (String2Date(sValue, t64, TRUE))
			pTasks->SetTaskLastModified64(hTask, t64);
		break;

	case TDCA_CREATIONDATE: 
		if (String2Date(sValue, t64, FALSE))
			pTasks->SetTaskCreationDate64(hTask, t64);
		break;

	case TDCA_PRIORITY: 
		pTasks->SetTaskPriority(hTask, (unsigned char)_ttoi(sValue));
		break;

	case TDCA_RISK: 
		pTasks->SetTaskRisk(hTask, (unsigned char)_ttoi(sValue));
		break;

	case TDCA_FLAG: 
		pTasks->SetTaskFlag(hTask, (_ttoi(sValue) != 0));
		break;

	case TDCA_COST: 
		pTasks->SetTaskCost(hTask, _ttof(sValue));
		break;

	case TDCA_TIMEEST: 
		{
			int nUnits = 0;
			double dAmount = 0.0;

			if (CTimeHelper::DecodeOffset(sValue, dAmount, nUnits, FALSE))
				pTasks->SetTaskTimeEstimate(hTask, dAmount, (TCHAR)nUnits);
		}
		break;

	case TDCA_TIMESPENT: 
		{
			int nUnits = 0;
			double dAmount = 0.0;

			if (CTimeHelper::DecodeOffset(sValue, dAmount, nUnits, FALSE))
				pTasks->SetTaskTimeSpent(hTask, dAmount, (TCHAR)nUnits);
		}
		break;
		
		// arrays
	case TDCA_CATEGORY: 
		SETTASKATTRIBARRAY(AddTaskCategory);
		break;
		
	case TDCA_TAGS: 
		SETTASKATTRIBARRAY(AddTaskTag);
		break;
		
	case TDCA_ALLOCTO: 
		SETTASKATTRIBARRAY(AddTaskAllocatedTo);
		break;
		
	case TDCA_FILEREF: 
		SETTASKATTRIBARRAY(AddTaskFileReference);
		break;
		
	case TDCA_DEPENDENCY: 
		SETTASKATTRIBARRAY(AddTaskDependency);
		break;
	}
}

TCHAR CTaskListCsvImporter::GetTimeUnits(const CStringArray& aValues, int nCol)
{
	TCHAR nUnits = 0;
	
	// get units from following column
	if (nCol < aValues.GetSize())
	{
		const CString& sUnits = Misc::GetItem(aValues, nCol);
		
		if (sUnits.GetLength() == 1)
		{
			switch (sUnits[0])
			{
			case THU_MINS:
			case THU_HOURS:
			case THU_DAYS:
			case THU_WEEKS:
			case THU_MONTHS:
			case THU_YEARS:
				nUnits = sUnits[0];
				break;
			}
		}
	}
	
	return nUnits;
}
