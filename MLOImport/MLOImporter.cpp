// MLOImporter.cpp: implementation of the CMLOImporter class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MLOImporter.h"
#include <time.h>
#include <unknwn.h>

#include "..\SHARED\xmlfile.h"
#include "..\SHARED\datehelper.h"

#include "..\Interfaces\iTasklist.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMLOImporter::CMLOImporter()
{

}

CMLOImporter::~CMLOImporter()
{

}

void CMLOImporter::SetLocalizer(ITransText* /*pTT*/)
{
	//CLocalizer::Initialize(pTT);
}

bool CMLOImporter::Import(LPCTSTR szSrcFilePath, ITaskList* pDestTaskFile, BOOL /*bSilent*/, IPreferences* /*pPrefs*/, LPCTSTR /*szKey*/)
{
	ITaskList5* pTL5 = GetITLInterface<ITaskList5>(pDestTaskFile, IID_TASKLIST5);

	if (!pTL5)
		return false;

	CXmlFile file;

	if (!file.Load(szSrcFilePath, _T("MyLifeOrganized-xml")))
		return false;

	const CXmlItem* pXIMLOTree = file.GetItem(_T("TaskTree"));

	if (!pXIMLOTree)
		return false;

	const CXmlItem* pXIMLOTask = pXIMLOTree->GetItem(_T("TaskNode"));

	if (!pXIMLOTask)
		return false;

	// this first node always seems to be untitled 
	// so we get the first subnode
	pXIMLOTask = pXIMLOTask->GetItem(_T("TaskNode"));

	if (!pXIMLOTask)
		return true; // just means it was an empty tasklist
	
	return ImportTask(pXIMLOTask, pTL5, NULL, TRUE); // NULL ==  root
}

bool CMLOImporter::ImportTask(const CXmlItem* pXIMLOTask, ITaskList5* pDestTaskFile, HTASKITEM hParent, BOOL bAndSiblings) const
{
	ASSERT (pXIMLOTask);

	HTASKITEM hTask = pDestTaskFile->NewTask(pXIMLOTask->GetItemValue(_T("Caption")), hParent);

	if (!hTask)
		return false;

	// priority (== Importance)
	int nPriority = (pXIMLOTask->GetItemValueI(_T("Importance")) * 10) / 100;

	pDestTaskFile->SetTaskPriority(hTask, (unsigned char)nPriority);

	// categories (== places)
	const CXmlItem* pXIPlace = pXIMLOTask->GetItem(_T("Places"), _T("Place"));

	while (pXIPlace)
	{
		pDestTaskFile->AddTaskCategory(hTask, pXIPlace->GetValue());
		pXIPlace = pXIPlace->GetSibling();
	}

	// estimated time (== EstimateMax)
	double dTimeEst = pXIMLOTask->GetItemValueF(_T("EstimateMax")); // in days

	pDestTaskFile->SetTaskTimeEstimate(hTask, dTimeEst, 'D');

	// due date (== DueDateTime)
	pDestTaskFile->SetTaskDueDate(hTask, GetDate(pXIMLOTask, _T("DueDateTime")));

	// completion (== CompletionDateTime)
	pDestTaskFile->SetTaskDoneDate(hTask, GetDate(pXIMLOTask, _T("CompletionDateTime")));

	// comments (== Note)
	CString sComments = pXIMLOTask->GetItemValue(_T("Note"));

	pDestTaskFile->SetTaskComments(hTask, sComments);

	// children
	const CXmlItem* pXIMLOSubTask = pXIMLOTask->GetItem(_T("TaskNode"));

	if (pXIMLOSubTask)
	{
		if (!ImportTask(pXIMLOSubTask, pDestTaskFile, hTask, TRUE))
			return false;
	}

	// handle sibling tasks WITHOUT RECURSION
	if (bAndSiblings)
	{
		pXIMLOTask = pXIMLOTask->GetSibling();
		
		while (pXIMLOTask)
		{
			// FALSE == don't recurse on siblings
			if (!ImportTask(pXIMLOTask, pDestTaskFile, hParent, FALSE))
				return false;
			
			pXIMLOTask = pXIMLOTask->GetSibling();
		}
	}

	return true;
}

time_t CMLOImporter::GetDate(const CXmlItem* pXIMLOTask, LPCTSTR szDateField) const
{
	CString sDate = pXIMLOTask->GetItemValue(szDateField);
	time_t date = 0;
	
	if (!CDateHelper::DecodeDate(sDate, date))
		date = 0;
	
	return date;
}
