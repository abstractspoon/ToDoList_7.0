// Fi M_BlISlteredToDoCtrl.cpp: implementation of the CFilteredToDoCtrl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "FilteredToDoCtrl.h"
#include "todoitem.h"
#include "resource.h"
#include "tdcstatic.h"
#include "tdcmsg.h"
#include "TDCCustomAttributeHelper.h"
#include "TDCSearchParamHelper.h"
#include "taskclipboard.h"

#include "..\shared\holdredraw.h"
#include "..\shared\datehelper.h"
#include "..\shared\enstring.h"
#include "..\shared\preferences.h"
#include "..\shared\deferwndmove.h"
#include "..\shared\autoflag.h"
#include "..\shared\holdredraw.h"
#include "..\shared\osversion.h"
#include "..\shared\graphicsmisc.h"
#include "..\shared\savefocus.h"
#include "..\shared\filemisc.h"

#include "..\Interfaces\IUIExtension.h"

#include <math.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

#ifndef LVS_EX_DOUBLEBUFFER
#define LVS_EX_DOUBLEBUFFER 0x00010000
#endif

#ifndef LVS_EX_LABELTIP
#define LVS_EX_LABELTIP     0x00004000
#endif

//////////////////////////////////////////////////////////////////////

const UINT SORTWIDTH = 10;

#ifdef _DEBUG
const UINT ONE_MINUTE = 10000;
#else
const UINT ONE_MINUTE = 60000;
#endif

const UINT TEN_MINUTES = (ONE_MINUTE * 10);

//////////////////////////////////////////////////////////////////////

const UINT WM_TDC_REFRESHFILTER	= (WM_APP + 11);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CFilteredToDoCtrl::CFilteredToDoCtrl(CUIExtensionMgr& mgrUIExt, CContentMgr& mgrContent, 
									 const CONTENTFORMAT& cfDefault, const TDCCOLEDITFILTERVISIBILITY& visDefault) 
	:
	CTabbedToDoCtrl(mgrUIExt, mgrContent, cfDefault, visDefault), 
	m_bLastFilterWasCustom(FALSE),
	m_bCustomFilter(FALSE)
{
}

CFilteredToDoCtrl::~CFilteredToDoCtrl()
{

}

BEGIN_MESSAGE_MAP(CFilteredToDoCtrl, CTabbedToDoCtrl)
//{{AFX_MSG_MAP(CFilteredToDoCtrl)
	ON_WM_DESTROY()
	ON_WM_TIMER()
	//}}AFX_MSG_MAP
	ON_REGISTERED_MESSAGE(WM_TDCN_VIEWPRECHANGE, OnPreTabViewChange)
	ON_NOTIFY(TVN_ITEMEXPANDED, IDC_TASKTREELIST, OnTreeExpandItem)
	ON_MESSAGE(WM_TDC_REFRESHFILTER, OnRefreshFilter)
	ON_CBN_EDITCHANGE(IDC_DUETIME, OnEditChangeDueTime)
END_MESSAGE_MAP()

///////////////////////////////////////////////////////////////////////////

BOOL CFilteredToDoCtrl::OnInitDialog()
{
	CTabbedToDoCtrl::OnInitDialog();

	return FALSE;
}

BOOL CFilteredToDoCtrl::LoadTasks(const CTaskFile& file)
{
	// handle reloading of tasklist with a filter present
	if (GetTaskCount() && m_filter.IsSet())
	{
		SaveSettings();
	}

	BOOL bViewWasUnset = (GetView() == FTCV_UNSET);

	if (!CTabbedToDoCtrl::LoadTasks(file))
		return FALSE;

	FTC_VIEW nView = GetView();

	// save visible state
	BOOL bHidden = !IsWindowVisible();

	// reload last view
	if (bViewWasUnset)
	{
		LoadSettings();

		// always refresh the tree filter because all other
		// views depend on it
		if (IsFilterSet(FTCV_TASKTREE))
			RefreshTreeFilter(); // always

		// handle other views
		switch (nView)
		{
		case FTCV_TASKLIST:
			if (IsFilterSet(nView))
			{
				RefreshListFilter();
			}
			else if (!GetPreferencesKey().IsEmpty()) // first time
			{
				GetViewData2(nView)->bNeedRefilter = TRUE;
			}
			break;

		case FTCV_UIEXTENSION1:
		case FTCV_UIEXTENSION2:
		case FTCV_UIEXTENSION3:
		case FTCV_UIEXTENSION4:
		case FTCV_UIEXTENSION5:
		case FTCV_UIEXTENSION6:
		case FTCV_UIEXTENSION7:
		case FTCV_UIEXTENSION8:
		case FTCV_UIEXTENSION9:
		case FTCV_UIEXTENSION10:
		case FTCV_UIEXTENSION11:
		case FTCV_UIEXTENSION12:
		case FTCV_UIEXTENSION13:
		case FTCV_UIEXTENSION14:
		case FTCV_UIEXTENSION15:
		case FTCV_UIEXTENSION16:
			// Note: By way of virtual functions CTabbedToDoCtrl::LoadTasks
			// will already have initialized the active view if it is an
			// extension so we only need to update if the tree actually
			// has a filter
			if (IsFilterSet(FTCV_TASKTREE))
				RefreshExtensionFilter(nView);
			break;
		}
	}
	else if (IsFilterSet(nView))
		RefreshFilter();

	// restore previously visibility
	if (bHidden)
		ShowWindow(SW_HIDE);

	return TRUE;
}

BOOL CFilteredToDoCtrl::DelayLoad(const CString& sFilePath, COleDateTime& dtEarliestDue)
{
	if (CTabbedToDoCtrl::DelayLoad(sFilePath, dtEarliestDue))
	{
		LoadSettings();
		return TRUE;
	}
	
	// else
	return FALSE;
}

void CFilteredToDoCtrl::SaveSettings() const
{
	CPreferences prefs;

	SaveFilter(m_filter, prefs);
	SaveFilter(m_lastFilter, prefs, _T("Last"));
	
	// record if we had a custom filter
	CString sKey = GetPreferencesKey(_T("Filter"));

	if (!sKey.IsEmpty())
	{
		prefs.WriteProfileInt(sKey, _T("CustomFilter"), m_bCustomFilter);
		prefs.WriteProfileInt(sKey, _T("LastFilterWasCustom"), m_bLastFilterWasCustom);

		SaveCustomFilter(m_sCustomFilter, m_customFilter, prefs);
		SaveCustomFilter(m_sCustomFilter, m_lastCustomFilter, prefs, _T("Last"));
	}
}

void CFilteredToDoCtrl::LoadSettings()
{
	if (HasStyle(TDCS_RESTOREFILTERS))
	{
		CPreferences prefs;

		// restore in reverse order
		LoadFilter(prefs, m_lastFilter, _T("Last"));
		LoadFilter(prefs, m_filter);
		
		// record if we had a custom filter
		CString sKey = GetPreferencesKey(_T("Filter"));
		
		if (!sKey.IsEmpty())
		{
			m_bCustomFilter = prefs.GetProfileInt(sKey, _T("CustomFilter"));
			m_bLastFilterWasCustom = prefs.GetProfileInt(sKey, _T("LastFilterWasCustom"));
			
			// restore in reverse order
			LoadCustomFilter(prefs, m_sCustomFilter, m_lastCustomFilter, _T("Last"));
			LoadCustomFilter(prefs, m_sCustomFilter, m_customFilter);
		}
	}
}

void CFilteredToDoCtrl::LoadCustomFilter(const CPreferences& prefs, CString& sFilter, SEARCHPARAMS& params, const CString& sSubKey)
{
	CString sKey = GetPreferencesKey(_T("Filter\\Custom"));

	if (!sKey.IsEmpty())
	{
		if (!sSubKey.IsEmpty())
			sKey += "\\" + sSubKey;

		sFilter = prefs.GetProfileString(sKey, _T("Name"));

		// reset
		params.Clear();

		// load
		params.bIgnoreDone = prefs.GetProfileInt(sKey, _T("IgnoreDone"), FALSE);
		params.bIgnoreOverDue = prefs.GetProfileInt(sKey, _T("IgnoreOverDue"), FALSE);
		params.bWantAllSubtasks = prefs.GetProfileInt(sKey, _T("WantAllSubtasks"), FALSE);

		int nNumRules = prefs.GetProfileInt(sKey, _T("NumRules"), 0), nNumDupes = 0;
		SEARCHPARAM rule, rulePrev;

		for (int nRule = 0; nRule < nNumRules; nRule++)
		{
			CString sRule = Misc::MakeKey(_T("Rule%d"), nRule, sKey);

			// stop loading if we meet an invalid rule
			if (!CTDCSearchParamHelper::LoadRule(prefs, sRule, m_aCustomAttribDefs, rule))
				break;

			// stop loading if we get more than 2 duplicates in a row.
			// this handles a bug resulting in an uncontrolled 
			// proliferation of duplicates that I cannot reproduce
			if (rule == rulePrev)
			{
				nNumDupes++;

				if (nNumDupes > 2)
					break;
			}
			else
			{
				nNumDupes = 0; // reset
				rulePrev = rule; // new 'previous rule
			}

			params.aRules.Add(rule);
		}
	}
}

void CFilteredToDoCtrl::SaveCustomFilter(const CString& sFilter, const SEARCHPARAMS& params, CPreferences& prefs, const CString& sSubKey) const
{
	// check anything to save
	if (params.aRules.GetSize() == 0)
		return;

	CString sKey = GetPreferencesKey(_T("Filter\\Custom"));

	if (!sKey.IsEmpty())
	{
		if (!sSubKey.IsEmpty())
			sKey += "\\" + sSubKey;

		// delete existing filter
		prefs.DeleteSection(sKey, TRUE);

		prefs.WriteProfileString(sKey, _T("Name"), sFilter);
		prefs.WriteProfileInt(sKey, _T("IgnoreDone"), params.bIgnoreDone);
		prefs.WriteProfileInt(sKey, _T("IgnoreOverDue"), params.bIgnoreOverDue);
		prefs.WriteProfileInt(sKey, _T("WantAllSubtasks"), params.bWantAllSubtasks);

		int nNumRules = 0, nNumDupes = 0;
		SEARCHPARAM rulePrev;

		for (int nRule = 0; nRule < params.aRules.GetSize(); nRule++)
		{
			CString sRule = Misc::MakeKey(_T("Rule%d"), nRule, sKey);
			const SEARCHPARAM& rule = params.aRules[nRule];

			// stop saving if we meet an invalid rule
			if (!CTDCSearchParamHelper::SaveRule(prefs, sRule, rule))
				break;

			// stop saving if we get more than 2 duplicates in a row.
			// this handles a bug resulting in an uncontrolled 
			// proliferation of duplicates that I cannot reproduce
			if (rule == rulePrev)
			{
				nNumDupes++;

				if (nNumDupes > 2)
					break;
			}
			else
			{
				nNumDupes = 0; // reset
				rulePrev = rule; // new 'previous rule
			}

			nNumRules++;
		}
		
		prefs.WriteProfileInt(sKey, _T("NumRules"), nNumRules);
	}
}

void CFilteredToDoCtrl::LoadFilter(const CPreferences& prefs, FTDCFILTER& filter, const CString& sSubKey)
{
	CString sKey = GetPreferencesKey(_T("Filter"));

	if (!sKey.IsEmpty())
	{
		if (!sSubKey.IsEmpty())
			sKey += "\\" + sSubKey;

		filter.nShow = (FILTER_SHOW)prefs.GetProfileInt(sKey, _T("Filter"), FS_ALL);
		filter.nStartBy = (FILTER_DATE)prefs.GetProfileInt(sKey, _T("Start"), FD_ANY);
		filter.nDueBy = (FILTER_DATE)prefs.GetProfileInt(sKey, _T("Due"), FD_ANY);

		filter.sTitle = prefs.GetProfileString(sKey, _T("Title"));
		filter.nPriority = prefs.GetProfileInt(sKey, _T("Priority"), FM_ANYPRIORITY);
		filter.nRisk = prefs.GetProfileInt(sKey, _T("Risk"), FM_ANYRISK);

		// dates
		if (filter.nStartBy == FD_USER)
			filter.dtUserStart = prefs.GetProfileDouble(sKey, _T("UserStart"), COleDateTime::GetCurrentTime());
		else
			filter.dtUserStart = COleDateTime::GetCurrentTime();

		if (filter.nDueBy == FD_USER)
			filter.dtUserDue = prefs.GetProfileDouble(sKey, _T("UserDue"), COleDateTime::GetCurrentTime());
		else
			filter.dtUserDue = COleDateTime::GetCurrentTime();

		// arrays
		prefs.GetProfileArray(sKey + _T("\\Category"), filter.aCategories, TRUE);
		prefs.GetProfileArray(sKey + _T("\\AllocTo"), filter.aAllocTo, TRUE);
		prefs.GetProfileArray(sKey + _T("\\AllocBy"), filter.aAllocBy, TRUE);
		prefs.GetProfileArray(sKey + _T("\\Status"), filter.aStatus, TRUE);
		prefs.GetProfileArray(sKey + _T("\\Version"), filter.aVersions, TRUE);
		prefs.GetProfileArray(sKey + _T("\\Tags"), filter.aTags, TRUE);

		// options
		filter.SetFlag(FO_ANYCATEGORY, prefs.GetProfileInt(sKey, _T("AnyCategory"), TRUE));
		filter.SetFlag(FO_ANYALLOCTO, prefs.GetProfileInt(sKey, _T("AnyAllocTo"), TRUE));
		filter.SetFlag(FO_ANYTAG, prefs.GetProfileInt(sKey, _T("AnyTag"), TRUE));
		filter.SetFlag(FO_HIDEPARENTS, prefs.GetProfileInt(sKey, _T("HideParents"), FALSE));
		filter.SetFlag(FO_HIDEOVERDUE, prefs.GetProfileInt(sKey, _T("HideOverDue"), FALSE));
		filter.SetFlag(FO_HIDEDONE, prefs.GetProfileInt(sKey, _T("HideDone"), FALSE));
		filter.SetFlag(FO_HIDECOLLAPSED, prefs.GetProfileInt(sKey, _T("HideCollapsed"), FALSE));
		filter.SetFlag(FO_SHOWALLSUB, prefs.GetProfileInt(sKey, _T("ShowAllSubtasks"), FALSE));
	}
}

void CFilteredToDoCtrl::SaveFilter(const FTDCFILTER& filter, CPreferences& prefs, const CString& sSubKey) const
{
	CString sKey = GetPreferencesKey(_T("Filter"));

	if (!sKey.IsEmpty())
	{
		if (!sSubKey.IsEmpty())
			sKey += _T("\\") + sSubKey;

		prefs.WriteProfileInt(sKey, _T("Filter"), filter.nShow);
		prefs.WriteProfileInt(sKey, _T("Start"), filter.nStartBy);
		prefs.WriteProfileInt(sKey, _T("Due"), filter.nDueBy);
		prefs.WriteProfileString(sKey, _T("Title"), filter.sTitle);
		prefs.WriteProfileInt(sKey, _T("Priority"), filter.nPriority);
		prefs.WriteProfileInt(sKey, _T("Risk"), filter.nRisk);

		// dates
		prefs.WriteProfileDouble(sKey, _T("UserStart"), filter.dtUserStart);
		prefs.WriteProfileDouble(sKey, _T("UserDue"), filter.dtUserDue);
		
		// arrays
		prefs.WriteProfileArray(sKey + _T("\\AllocBy"), filter.aAllocBy);
		prefs.WriteProfileArray(sKey + _T("\\Status"), filter.aStatus);
		prefs.WriteProfileArray(sKey + _T("\\Version"), filter.aVersions);
		prefs.WriteProfileArray(sKey + _T("\\AllocTo"), filter.aAllocTo);
		prefs.WriteProfileArray(sKey + _T("\\Category"), filter.aCategories);
		prefs.WriteProfileArray(sKey + _T("\\Tags"), filter.aTags);

		// options
		prefs.WriteProfileInt(sKey, _T("AnyAllocTo"), filter.HasFlag(FO_ANYALLOCTO));
		prefs.WriteProfileInt(sKey, _T("AnyCategory"), filter.HasFlag(FO_ANYCATEGORY));
		prefs.WriteProfileInt(sKey, _T("AnyTag"), filter.HasFlag(FO_ANYTAG));
		prefs.WriteProfileInt(sKey, _T("HideParents"), filter.HasFlag(FO_HIDEPARENTS));
		prefs.WriteProfileInt(sKey, _T("HideOverDue"), filter.HasFlag(FO_HIDEOVERDUE));
		prefs.WriteProfileInt(sKey, _T("HideDone"), filter.HasFlag(FO_HIDEDONE));
		prefs.WriteProfileInt(sKey, _T("HideCollapsed"), filter.HasFlag(FO_HIDECOLLAPSED));
		prefs.WriteProfileInt(sKey, _T("ShowAllSubtasks"), filter.HasFlag(FO_SHOWALLSUB));
	}
}

void CFilteredToDoCtrl::OnDestroy() 
{
	SaveSettings();

	CTabbedToDoCtrl::OnDestroy();
}

void CFilteredToDoCtrl::OnEditChangeDueTime()
{
	// need some special hackery to prevent a re-filter in the middle
	// of the user manually typing into the time field
	BOOL bNeedsRefilter = ModNeedsRefilter(TDCA_DUEDATE, FTCV_TASKTREE, GetSelectedTaskID());
	
	if (bNeedsRefilter)
		SetStyle(TDCS_REFILTERONMODIFY, FALSE, FALSE);
	
	CTabbedToDoCtrl::OnSelChangeDueTime();
	
	if (bNeedsRefilter)
		SetStyle(TDCS_REFILTERONMODIFY, TRUE, FALSE);
}

void CFilteredToDoCtrl::OnTreeExpandItem(NMHDR* /*pNMHDR*/, LRESULT* /*pResult*/)
{
	if (m_filter.HasFlag(FO_HIDECOLLAPSED))
	{
		if (InListView())
			RefreshListFilter();
		else
			GetViewData2(FTCV_TASKLIST)->bNeedRefilter = TRUE;
	}
}

LRESULT CFilteredToDoCtrl::OnPreTabViewChange(WPARAM nOldView, LPARAM nNewView) 
{
	if (nNewView != FTCV_TASKTREE)
	{
		const VIEWDATA2* pData = GetViewData2((FTC_VIEW)nNewView);
		BOOL bFiltered = FALSE;

		// take a note of what task is currently singly selected
		// so that we can prevent unnecessary calls to UpdateControls
		DWORD dwSelTaskID = GetSingleSelectedTaskID();

		switch (nNewView)
		{
		case FTCV_TASKLIST:
			// update filter as required
			if (pData->bNeedRefilter)
			{
				bFiltered = TRUE;
				RefreshListFilter();
			}
			break;

		case FTCV_UIEXTENSION1:
		case FTCV_UIEXTENSION2:
		case FTCV_UIEXTENSION3:
		case FTCV_UIEXTENSION4:
		case FTCV_UIEXTENSION5:
		case FTCV_UIEXTENSION6:
		case FTCV_UIEXTENSION7:
		case FTCV_UIEXTENSION8:
		case FTCV_UIEXTENSION9:
		case FTCV_UIEXTENSION10:
		case FTCV_UIEXTENSION11:
		case FTCV_UIEXTENSION12:
		case FTCV_UIEXTENSION13:
		case FTCV_UIEXTENSION14:
		case FTCV_UIEXTENSION15:
		case FTCV_UIEXTENSION16:
			// update filter as required
			if (pData && pData->bNeedRefilter)
			{
				// initialise progress depending on whether extension
				// window is already created
				UINT nProgressMsg = 0;

				if (GetExtensionWnd((FTC_VIEW)nNewView, FALSE) == NULL)
					nProgressMsg = IDS_INITIALISINGTABBEDVIEW;

				BeginExtensionProgress(pData, nProgressMsg);
				RefreshExtensionFilter((FTC_VIEW)nNewView);

				bFiltered = TRUE;
			}
			break;
		}
		
		// update controls only if the selection has changed and 
		// we didn't refilter (RefreshFilter will already have called UpdateControls)
		BOOL bSelChange = HasSingleSelectionChanged(dwSelTaskID);
		
		if (bSelChange && !bFiltered)
			UpdateControls();
	}

	return CTabbedToDoCtrl::OnPreTabViewChange(nOldView, nNewView);
}


BOOL CFilteredToDoCtrl::CopyCurrentSelection() const
{
	if (!GetSelectedCount())
		return FALSE;
	
	ClearCopiedItem();
	
	// NOTE: we are overriding this function else
	// filtered out tasks will not get copied

	// NOTE: We DON'T override GetSelectedTasks because
	// most often that only wants visible tasks

	// ISO date strings
	// must be done first before any tasks are added
	CTaskFile tasks;
	tasks.EnableISODates(HasStyle(TDCS_SHOWDATESINISO));
	
	// get selected tasks ordered, removing duplicate subtasks
	CHTIList selection;
	TSH().CopySelection(selection, TRUE, TRUE);
	
	// copy items
	POSITION pos = selection.GetHeadPosition();
	
	while (pos)
	{
		HTREEITEM hti = selection.GetNext(pos);
		DWORD dwTaskID = GetTrueTaskID(hti);

		const TODOSTRUCTURE* pTDS = m_data.LocateTask(dwTaskID);
		const TODOITEM* pTDI = GetTask(dwTaskID);

		if (!pTDS || !pTDI)
			return FALSE;

		// add task
		HTASKITEM hTask = tasks.NewTask(pTDI->sTitle, NULL, dwTaskID);
		ASSERT(hTask);
		
		if (!hTask)
			return FALSE;
		
		SetTaskAttributes(pTDI, pTDS, tasks, hTask, TDCGT_ALL, FALSE);

		// and subtasks
		AddSubTasksToTaskFile(pTDS, tasks, hTask);
	}
	
	// extra processing to identify the originally selected tasks
	// in case the user want's to paste as references
	pos = TSH().GetFirstItemPos();
	
	while (pos)
	{
		DWORD dwSelID = TSH().GetNextItemData(pos);
		ASSERT(dwSelID);
		
		HTASKITEM hSelTask = tasks.FindTask(dwSelID);
		ASSERT(hSelTask);
		
		tasks.SetTaskMetaData(hSelTask, _T("selected"), _T("1"));
	}
	
	// filename
	tasks.SetFileName(m_sLastSavePath);
	
	// meta data
	tasks.SetMetaData(m_mapMetaData);
	
	// custom data definitions
	if (m_aCustomAttribDefs.GetSize())
		tasks.SetCustomAttributeDefs(m_aCustomAttribDefs);
	
	// and their titles (including child dupes)
	CStringArray aTitles;
	
	VERIFY(TSH().CopySelection(selection, FALSE, TRUE));
	VERIFY(TSH().GetItemTitles(selection, aTitles));
	
	return CTaskClipboard::SetTasks(tasks, GetClipboardID(), Misc::FormatArray(aTitles, '\n'));
}

BOOL CFilteredToDoCtrl::ArchiveDoneTasks(TDC_ARCHIVE nFlags, BOOL bRemoveFlagged)
{
	if (CTabbedToDoCtrl::ArchiveDoneTasks(nFlags, bRemoveFlagged))
	{
		if (InListView())
		{
			if (IsFilterSet(FTCV_TASKLIST))
				RefreshListFilter();
		}
		else if (IsFilterSet(FTCV_TASKTREE))
		{
			RefreshTreeFilter();
		}

		return TRUE;
	}

	// else
	return FALSE;
}

BOOL CFilteredToDoCtrl::ArchiveSelectedTasks(BOOL bRemove)
{
	if (CTabbedToDoCtrl::ArchiveSelectedTasks(bRemove))
	{
		if (InListView())
		{
			if (IsFilterSet(FTCV_TASKLIST))
				RefreshListFilter();
		}
		else if (IsFilterSet(FTCV_TASKTREE))
		{
			RefreshTreeFilter();
		}

		return TRUE;
	}

	// else
	return FALSE;
}

int CFilteredToDoCtrl::GetArchivableTasks(CTaskFile& tasks, BOOL bSelectedOnly) const
{
	if (bSelectedOnly || !IsFilterSet(FTCV_TASKTREE))
		return CTabbedToDoCtrl::GetArchivableTasks(tasks, bSelectedOnly);

	// else process the entire data hierarchy
	GetCompletedTasks(m_data.GetStructure(), tasks, NULL, FALSE);

	return tasks.GetTaskCount();
}

BOOL CFilteredToDoCtrl::RemoveArchivedTask(DWORD dwTaskID)
{
	ASSERT(m_data.HasTask(dwTaskID,FALSE));
	
	// note: if the tasks does not exist in the tree then this is not a bug
	// if a filter is set
	HTREEITEM hti = m_taskTree.GetItem(dwTaskID);
	ASSERT(hti || IsFilterSet(FTCV_TASKTREE));
	
	if (!hti && !IsFilterSet(FTCV_TASKTREE))
		return FALSE;
	
	if (hti)
		m_taskTree.Tree().DeleteItem(hti);

	return m_data.DeleteTask(dwTaskID);
}

void CFilteredToDoCtrl::GetCompletedTasks(const TODOSTRUCTURE* pTDS, CTaskFile& tasks, HTASKITEM hTaskParent, BOOL bSelectedOnly) const
{
	const TODOITEM* pTDI = NULL;

	if (!pTDS->IsRoot())
	{
		DWORD dwTaskID = pTDS->GetTaskID();

		pTDI = m_data.GetTask(dwTaskID);
		ASSERT(pTDI);

		if (!pTDI)
			return;

		// we add the task if it is completed (and optionally selected) or it has children
		if (pTDI->IsDone() || pTDS->HasSubTasks())
		{
			HTASKITEM hTask = tasks.NewTask(_T(""), hTaskParent, dwTaskID);
			ASSERT(hTask);

			// copy attributes
			TDCGETTASKS allTasks;
			SetTaskAttributes(pTDI, pTDS, tasks, hTask, allTasks, FALSE);

			// this task is now the new parent
			hTaskParent = hTask;
		}
	}

	// children
	if (pTDS->HasSubTasks())
	{
		for (int nSubtask = 0; nSubtask < pTDS->GetSubTaskCount(); nSubtask++)
		{
			const TODOSTRUCTURE* pTDSChild = pTDS->GetSubTask(nSubtask);
			GetCompletedTasks(pTDSChild, tasks, hTaskParent, bSelectedOnly); // RECURSIVE call
		}

		// if no subtasks were added and the parent is not completed 
		// (and optionally selected) then we remove it
		if (hTaskParent && tasks.GetFirstTask(hTaskParent) == NULL)
		{
			ASSERT(pTDI);

			if (pTDI && !pTDI->IsDone())
				tasks.DeleteTask(hTaskParent);
		}
	}
}

int CFilteredToDoCtrl::GetFilteredTasks(CTaskFile& tasks, const TDCGETTASKS& filter) const
{
	// synonym for GetTasks which always returns the filtered tasks
	return GetTasks(tasks, filter);
}

FILTER_SHOW CFilteredToDoCtrl::GetFilter(FTDCFILTER& filter) const
{
	filter = m_filter;
	return m_filter.nShow;
}

void CFilteredToDoCtrl::SetFilter(const FTDCFILTER& filter)
{
	FTDCFILTER fPrev = m_filter;
	m_filter = filter;

	FTC_VIEW nView = GetView();

	if (m_bDelayLoaded)
	{
		// mark everything needing refilter
		GetViewData2(FTCV_TASKTREE)->bNeedRefilter = TRUE;
		SetListNeedRefilter(TRUE);
		SetExtensionsNeedRefilter(TRUE);
	}
	else
	{
		BOOL bTreeNeedsFilter = (m_bCustomFilter || !FiltersMatch(fPrev, filter, FTCV_TASKTREE));
		BOOL bListNeedRefilter = (m_bCustomFilter || !FiltersMatch(fPrev, filter, FTCV_TASKLIST)); 

		m_bCustomFilter = FALSE;

		if (bTreeNeedsFilter)
		{
			// this will mark all other views as needing refiltering
			// and refilter them if they are active
			RefreshFilter();
		}
		else if (bListNeedRefilter)
		{
			if (nView == FTCV_TASKLIST)
				RefreshListFilter();
			else
				SetListNeedRefilter(TRUE);
		}
	}

	ResetNowFilterTimer();
}
	
void CFilteredToDoCtrl::ClearFilter()
{
	if (HasFilter() || HasCustomFilter())
	{
		FTDCFILTER filter; // empty filter
		SetFilter(filter);
		m_lastFilter = filter; // clear

		m_customFilter = SEARCHPARAMS(); // empty ruleset
		m_bCustomFilter = FALSE;
		m_bLastFilterWasCustom = FALSE;
	}

	ResetNowFilterTimer();
}

void CFilteredToDoCtrl::ToggleFilter()
{
	if (HasFilter() || HasCustomFilter()) // turn off
	{
		// save last filter and clear
		// the order here is important because ClearFilter()
		// will also reset the last filter
		if (m_bCustomFilter)
		{
			SEARCHPARAMS temp(m_customFilter);
			ClearFilter();

			m_bLastFilterWasCustom = TRUE;
			m_lastCustomFilter = temp;
		}
		else
		{
			FTDCFILTER temp(m_filter);
			ClearFilter();

			m_lastFilter = temp; 
		}
	}
	else // restore
	{
		if (m_bLastFilterWasCustom)
			SetCustomFilter(m_lastCustomFilter, m_sCustomFilter);
		else
			SetFilter(m_lastFilter);
	}

	ResetNowFilterTimer();
}

BOOL CFilteredToDoCtrl::FiltersMatch(const FTDCFILTER& filter1, const FTDCFILTER& filter2, FTC_VIEW nView) const
{
	if (nView == FTCV_UNSET)
		return FALSE;

	DWORD dwIgnore = 0;

	if (nView == FTCV_TASKTREE)
		dwIgnore = (FO_HIDECOLLAPSED | FO_HIDEPARENTS);

	return FTDCFILTER::FiltersMatch(filter1, filter2, dwIgnore);
}

BOOL CFilteredToDoCtrl::IsFilterSet(FTC_VIEW nView) const
{
	return IsFilterSet(m_filter, nView) || HasCustomFilter();
}

BOOL CFilteredToDoCtrl::IsFilterSet(const FTDCFILTER& filter, FTC_VIEW nView) const
{
	if (nView == FTCV_UNSET)
		return FALSE;

	DWORD dwIgnore = 0;

	if (nView == FTCV_TASKTREE)
		dwIgnore = (FO_HIDECOLLAPSED | FO_HIDEPARENTS);

	return filter.IsSet(dwIgnore);
}

UINT CFilteredToDoCtrl::GetTaskCount(UINT* pVisible) const
{
	if (pVisible)
	{
		if (InListView())
			*pVisible = m_taskList.GetItemCount();
		else
			*pVisible = m_taskTree.GetItemCount();
	}

	return CTabbedToDoCtrl::GetTaskCount();
}

int CFilteredToDoCtrl::FindTasks(const SEARCHPARAMS& params, CResultArray& aResults) const
{
	if (params.bIgnoreFilteredOut)
		return CTabbedToDoCtrl::FindTasks(params, aResults);
	
	// else all tasks
	return m_data.FindTasks(params, aResults);
}

BOOL CFilteredToDoCtrl::HasCustomFilter() const 
{ 
	return m_bCustomFilter; 
}

CString CFilteredToDoCtrl::GetCustomFilterName() const 
{ 
	if (HasCustomFilter())
		return m_sCustomFilter; 

	// else
	return _T("");
}

BOOL CFilteredToDoCtrl::SetCustomFilter(const SEARCHPARAMS& params, LPCTSTR szName)
{
	if (!m_bCustomFilter || (m_customFilter != params))
	{
		m_customFilter = params;
		m_customFilter.aAttribDefs.Copy(m_aCustomAttribDefs);
		
		RestoreCustomFilter();
	}

	m_sCustomFilter = szName;
	return TRUE;
}

BOOL CFilteredToDoCtrl::RestoreCustomFilter()
{
	m_bCustomFilter = TRUE;

	if (m_bDelayLoaded)
	{
		// mark everything needing refilter
		GetViewData2(FTCV_TASKTREE)->bNeedRefilter = TRUE;
		SetListNeedRefilter(TRUE);
		SetExtensionsNeedRefilter(TRUE);
	}
	else
	{
		RefreshFilter();
	}

	return TRUE;
}

void CFilteredToDoCtrl::BuildFilterQuery(SEARCHPARAMS& params) const
{
	if (m_bCustomFilter)
	{
		params = m_customFilter;
		params.aAttribDefs.Copy(m_aCustomAttribDefs);

		return;
	}

	// reset the search
	params.aRules.RemoveAll();
	params.bIgnoreDone = m_filter.HasFlag(FO_HIDEDONE);
	params.bWantAllSubtasks = m_filter.HasFlag(FO_SHOWALLSUB);
	
	// handle principle filter
	switch (m_filter.nShow)
	{
	case FS_ALL:
		break;

	case FS_DONE:
		params.aRules.Add(SEARCHPARAM(TDCA_DONEDATE, FOP_SET));
		params.bIgnoreDone = FALSE;
		break;

	case FS_NOTDONE:
		params.aRules.Add(SEARCHPARAM(TDCA_DONEDATE, FOP_NOT_SET));
		params.bIgnoreDone = TRUE;
		break;

	case FS_FLAGGED:
		params.aRules.Add(SEARCHPARAM(TDCA_FLAG, FOP_SET));
		break;

	case FS_SELECTED:
		params.aRules.Add(SEARCHPARAM(TDCA_SELECTION, FOP_SET));
		break;

	case FS_RECENTMOD:
		params.aRules.Add(SEARCHPARAM(TDCA_RECENTMODIFIED, FOP_SET));
		break;

	default:
		ASSERT(0); // to catch unimplemented filters
		break;
	}

	// handle start date filters
	COleDateTime dateStart;

	if (InitFilterDate(m_filter.nStartBy, m_filter.dtUserStart, dateStart))
	{
		// special case: FD_NOW
		if (m_filter.nStartBy == FD_NOW)
			params.aRules.Add(SEARCHPARAM(TDCA_STARTTIME, FOP_ON_OR_BEFORE, dateStart));
		else
			params.aRules.Add(SEARCHPARAM(TDCA_STARTDATE, FOP_ON_OR_BEFORE, dateStart));

		params.aRules.Add(SEARCHPARAM(TDCA_DONEDATE, FOP_NOT_SET));
	}
	else if (m_filter.nStartBy == FD_NONE)
	{
		params.aRules.Add(SEARCHPARAM(TDCA_STARTDATE, FOP_NOT_SET));
		params.aRules.Add(SEARCHPARAM(TDCA_DONEDATE, FOP_NOT_SET));
	}

	// handle due date filters
	COleDateTime dateDue;

	if (InitFilterDate(m_filter.nDueBy, m_filter.dtUserDue, dateDue))
	{
		// special case: FD_NOW
		if (m_filter.nDueBy == FD_NOW)
			params.aRules.Add(SEARCHPARAM(TDCA_DUETIME, FOP_ON_OR_BEFORE, dateDue));
		else
			params.aRules.Add(SEARCHPARAM(TDCA_DUEDATE, FOP_ON_OR_BEFORE, dateDue));

		params.aRules.Add(SEARCHPARAM(TDCA_DONEDATE, FOP_NOT_SET));
			
		// this flag only applies to due filters
		params.bIgnoreOverDue = m_filter.HasFlag(FO_HIDEOVERDUE);
	}
	else if (m_filter.nDueBy == FD_NONE)
	{
		params.aRules.Add(SEARCHPARAM(TDCA_DUEDATE, FOP_NOT_SET));
		params.aRules.Add(SEARCHPARAM(TDCA_DONEDATE, FOP_NOT_SET));
	}

	// handle other attributes
	AddNonDateFilterQueryRules(params);
}

void CFilteredToDoCtrl::AddNonDateFilterQueryRules(SEARCHPARAMS& params) const
{
	// title text
	if (!m_filter.sTitle.IsEmpty())
	{
		switch (m_filter.nTitleOption)
		{
		case FT_FILTERONTITLECOMMENTS:
			params.aRules.Add(SEARCHPARAM(TDCA_TASKNAMEORCOMMENTS, FOP_INCLUDES, m_filter.sTitle));
			break;
			
		case FT_FILTERONANYTEXT:
			params.aRules.Add(SEARCHPARAM(TDCA_ANYTEXTATTRIBUTE, FOP_INCLUDES, m_filter.sTitle));
			break;
			
		case FT_FILTERONTITLEONLY:
		default:
			params.aRules.Add(SEARCHPARAM(TDCA_TASKNAME, FOP_INCLUDES, m_filter.sTitle));
			break;
		}
	}

	// note: these are all 'AND' 
	// category
	if (m_filter.aCategories.GetSize())
	{
		CString sMatchBy = Misc::FormatArray(m_filter.aCategories);

		if (m_filter.aCategories.GetSize() == 1 && sMatchBy.IsEmpty())
			params.aRules.Add(SEARCHPARAM(TDCA_CATEGORY, FOP_NOT_SET));

		else if (m_filter.dwFlags & FO_ANYCATEGORY)
			params.aRules.Add(SEARCHPARAM(TDCA_CATEGORY, FOP_INCLUDES, sMatchBy));
		else
			params.aRules.Add(SEARCHPARAM(TDCA_CATEGORY, FOP_EQUALS, sMatchBy));
	}

	// allocated to
	if (m_filter.aAllocTo.GetSize())
	{
		CString sMatchBy = Misc::FormatArray(m_filter.aAllocTo);

		if (m_filter.aAllocTo.GetSize() == 1 && sMatchBy.IsEmpty())
			params.aRules.Add(SEARCHPARAM(TDCA_ALLOCTO, FOP_NOT_SET));

		else if (m_filter.dwFlags & FO_ANYALLOCTO)
			params.aRules.Add(SEARCHPARAM(TDCA_ALLOCTO, FOP_INCLUDES, sMatchBy));
		else
			params.aRules.Add(SEARCHPARAM(TDCA_ALLOCTO, FOP_EQUALS, sMatchBy));
	}

	// allocated by
	if (m_filter.aAllocBy.GetSize())
	{
		CString sMatchBy = Misc::FormatArray(m_filter.aAllocBy);

		if (m_filter.aAllocBy.GetSize() == 1 && sMatchBy.IsEmpty())
			params.aRules.Add(SEARCHPARAM(TDCA_ALLOCBY, FOP_NOT_SET));
		else
			params.aRules.Add(SEARCHPARAM(TDCA_ALLOCBY, FOP_INCLUDES, sMatchBy));
	}

	// status
	if (m_filter.aStatus.GetSize())
	{
		CString sMatchBy = Misc::FormatArray(m_filter.aStatus);

		if (m_filter.aStatus.GetSize() == 1 && sMatchBy.IsEmpty())
			params.aRules.Add(SEARCHPARAM(TDCA_STATUS, FOP_NOT_SET));
		else
			params.aRules.Add(SEARCHPARAM(TDCA_STATUS, FOP_INCLUDES, sMatchBy));
	}

	// version
	if (m_filter.aVersions.GetSize())
	{
		CString sMatchBy = Misc::FormatArray(m_filter.aVersions);

		if (m_filter.aVersions.GetSize() == 1 && sMatchBy.IsEmpty())
			params.aRules.Add(SEARCHPARAM(TDCA_VERSION, FOP_NOT_SET));
		else
			params.aRules.Add(SEARCHPARAM(TDCA_VERSION, FOP_INCLUDES, sMatchBy));
	}

	// tags
	if (m_filter.aTags.GetSize())
	{
		CString sMatchBy = Misc::FormatArray(m_filter.aTags);

		if (m_filter.aTags.GetSize() == 1 && sMatchBy.IsEmpty())
			params.aRules.Add(SEARCHPARAM(TDCA_TAGS, FOP_NOT_SET));
		else
			params.aRules.Add(SEARCHPARAM(TDCA_TAGS, FOP_INCLUDES, sMatchBy));
	}

	// priority
	if (m_filter.nPriority != FM_ANYPRIORITY)
	{
		if (m_filter.nPriority == FM_NOPRIORITY)
			params.aRules.Add(SEARCHPARAM(TDCA_PRIORITY, FOP_NOT_SET));

		else if (m_filter.nPriority != FM_ANYPRIORITY)
			params.aRules.Add(SEARCHPARAM(TDCA_PRIORITY, FOP_GREATER_OR_EQUAL, m_filter.nPriority));
	}

	// risk
	if (m_filter.nRisk != FM_ANYRISK)
	{
		if (m_filter.nRisk == FM_NORISK)
			params.aRules.Add(SEARCHPARAM(TDCA_RISK, FOP_NOT_SET));
		
		else if (m_filter.nRisk != FM_ANYRISK)
			params.aRules.Add(SEARCHPARAM(TDCA_RISK, FOP_GREATER_OR_EQUAL, m_filter.nRisk));
	}

	// special case: no aRules + ignore completed
	if ((params.bIgnoreDone) && params.aRules.GetSize() == 0)
		params.aRules.Add(SEARCHPARAM(TDCA_DONEDATE, FOP_NOT_SET));
}

LRESULT CFilteredToDoCtrl::OnRefreshFilter(WPARAM wParam, LPARAM lParam)
{
	BOOL bUndo = lParam;
	FTC_VIEW nView = (FTC_VIEW)wParam;
	
	if (nView == FTCV_TASKTREE)
		RefreshFilter();
	
	// if undoing then we must also refresh the list filter because
	// otherwise ResyncListSelection will fail in the case where
	// we are undoing a delete because the undone item will not yet be in the list.
	if (nView == FTCV_TASKLIST || bUndo)
		RefreshListFilter();
	
	// resync selection?
	if (nView == FTCV_TASKLIST)
		ResyncListSelection();
	
	return 0L;
}

void CFilteredToDoCtrl::RefreshFilter() 
{
	CSaveFocus sf;

	RefreshTreeFilter(); // always

	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
		// mark all other views as needing refiltering
		SetListNeedRefilter(TRUE);
		SetExtensionsNeedRefilter(TRUE);
		break;

	case FTCV_TASKLIST:
		// mark extensions as needing refiltering
		RefreshListFilter();
		SetExtensionsNeedRefilter(TRUE);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		SetExtensionsNeedRefilter(TRUE);
		SetListNeedRefilter(TRUE);
		RefreshExtensionFilter(nView, TRUE);
		break;
	}
}

void CFilteredToDoCtrl::SetListNeedRefilter(BOOL bRefilter)
{
	GetViewData2(FTCV_TASKLIST)->bNeedRefilter = bRefilter;
}

void CFilteredToDoCtrl::SetExtensionsNeedRefilter(BOOL bRefilter, FTC_VIEW nIgnore)
{
	for (int nExt = 0; nExt < m_aExtViews.GetSize(); nExt++)
	{
		FTC_VIEW nView = (FTC_VIEW)(FTCV_UIEXTENSION1 + nExt);

		if (nView == nIgnore)
			continue;

		// else
		VIEWDATA2* pData = GetViewData2(nView);

		if (pData)
			pData->bNeedRefilter = bRefilter;
	}
}

void CFilteredToDoCtrl::RefreshTreeFilter() 
{
	if (m_data.GetTaskCount())
	{
		BOOL bTreeVis = !InListView();

		// save and reset current focus to work around a bug
		CSaveFocus sf;
		SetFocusToTasks();
		
		// cache current selection and task breadcrumbs before clearing selection
		TDCSELECTIONCACHE cache;
		CacheTreeSelection(cache);

		// grab the selected items for the filtering
		m_aSelectedTaskIDs.Copy(cache.aSelTaskIDs);
		
		TSH().RemoveAll();
		
		// and expanded state
		CPreferences prefs;
		SaveTasksState(prefs);
		
		CHoldRedraw hr(GetSafeHwnd());
		CWaitCursor cursor;
		
		// rebuild the tree
		RebuildTree();
		
		// redo last sort
		if (bTreeVis && IsSorting())
		{
			Resort();
			m_bTreeNeedResort = FALSE;
		}
		else
		{
			m_bTreeNeedResort = TRUE;
		}
		
		// restore expanded state
		LoadTasksState(prefs); 
		
		// restore selection
		if (!RestoreTreeSelection(cache))
		{
			HTREEITEM hti = m_taskTree.GetChildItem(NULL); // select first item

			if (hti)
				SelectItem(hti);
		}
	}
	
	// modify the tree prompt depending on whether there is a filter set
	UINT nPromptID = (IsFilterSet(FTCV_TASKTREE) ? IDS_TDC_FILTEREDTASKLISTPROMPT : IDS_TDC_TASKLISTPROMPT);
	m_taskTree.SetWindowPrompt(CEnString(nPromptID));
}

HTREEITEM CFilteredToDoCtrl::RebuildTree(const void* /*pContext*/)
{
	m_taskTree.DeleteAll();
	TCH().SelectItem(NULL);

	// build a find query that matches the filter
	if (HasFilter() || HasCustomFilter())
	{
		SEARCHPARAMS filter;
		BuildFilterQuery(filter);

		return CTabbedToDoCtrl::RebuildTree(&filter);
	}

	// else
	return CTabbedToDoCtrl::RebuildTree(NULL);
}

BOOL CFilteredToDoCtrl::WantAddTask(const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS, const void* pContext) const
{
	BOOL bWantTask = CTabbedToDoCtrl::WantAddTask(pTDI, pTDS, pContext);
	
	if (bWantTask && pContext != NULL) // it's a filter
	{
		const SEARCHPARAMS* pFilter = static_cast<const SEARCHPARAMS*>(pContext);
		SEARCHRESULT result;
		
		// special case: selected item
		if (pFilter->HasAttribute(TDCA_SELECTION))
		{
			bWantTask = (Misc::FindT(m_aSelectedTaskIDs, pTDS->GetTaskID()) != -1);

			// check parents
			if (!bWantTask && pFilter->bWantAllSubtasks)
			{
				TODOSTRUCTURE* pTDSParent = pTDS->GetParentTask();

				while (pTDSParent && !bWantTask)
				{
					bWantTask = (Misc::FindT(m_aSelectedTaskIDs, pTDSParent->GetTaskID()) != -1);
					pTDSParent = pTDSParent->GetParentTask();
				}
			}
		}
		else // rest of attributes
		{
			bWantTask = m_data.TaskMatches(pTDI, pTDS, *pFilter, result);
		}

		if (bWantTask && pTDS->HasSubTasks())
		{
			// NOTE: the only condition under which this method is called for
			// a parent is if none of its subtasks matched the filter.
			//
			// So if we're a parent and match the filter we need to do an extra check
			// to see if what actually matched was the absence of attributes
			//
			// eg. if the parent category is "" and the filter rule is 
			// TDCA_CATEGORY is (FOP_NOT_SET or FOP_NOT_INCLUDES or FOP_NOT_EQUAL) 
			// then we don't treat this as a match.
			//
			// The attributes to check are:
			//  Category
			//  Status
			//  Alloc To
			//  Alloc By
			//  Version
			//  Priority
			//  Risk
			//  Tags
			
			int nNumRules = pFilter->aRules.GetSize();
			
			for (int nRule = 0; nRule < nNumRules && bWantTask; nRule++)
			{
				const SEARCHPARAM& sp = pFilter->aRules[nRule];

				if (!sp.OperatorIs(FOP_NOT_EQUALS) && 
					!sp.OperatorIs(FOP_NOT_INCLUDES) && 
					!sp.OperatorIs(FOP_NOT_SET))
				{
					continue;
				}
				
				// else check for empty parent attributes
				switch (sp.GetAttribute())
				{
				case TDCA_ALLOCTO:
					bWantTask = (pTDI->aAllocTo.GetSize() > 0);
					break;
					
				case TDCA_ALLOCBY:
					bWantTask = !pTDI->sAllocBy.IsEmpty();
					break;
					
				case TDCA_VERSION:
					bWantTask = !pTDI->sVersion.IsEmpty();
					break;
					
				case TDCA_STATUS:
					bWantTask = !pTDI->sStatus.IsEmpty();
					break;
					
				case TDCA_CATEGORY:
					bWantTask = (pTDI->aCategories.GetSize() > 0);
					break;
					
				case TDCA_TAGS:
					bWantTask = (pTDI->aTags.GetSize() > 0);
					break;
					
				case TDCA_PRIORITY:
					bWantTask = (pTDI->nPriority != FM_NOPRIORITY);
					break;
					
				case TDCA_RISK:
					bWantTask = (pTDI->nRisk != FM_NORISK);
					break;
				}
			}
		}
	}
	
	return bWantTask; 
}

void CFilteredToDoCtrl::RefreshExtensionFilter(FTC_VIEW nView, BOOL bShowProgress)
{
	IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView, TRUE);
	ASSERT(pExtWnd);

	if (pExtWnd)
	{
		VIEWDATA2* pData = GetViewData2(nView);
		ASSERT(pData);

		if (bShowProgress)
			BeginExtensionProgress(pData, IDS_UPDATINGTABBEDVIEW);
		
		// update view with filtered tasks
		// NOTE: Need to be a bit clever here to always
		// refresh using the TREE's filtered tasks otherwise
		// we may only get a subset from the list if it's 
		// currently active
		CTaskFile tasks;
		CToDoCtrl::GetTasks(tasks);

		UpdateExtensionView(pExtWnd, tasks, IUI_ALL);

		// and clear all update flags
		pData->bNeedTaskUpdate = FALSE;
		pData->bNeedRefilter = FALSE;

		if (bShowProgress)
			EndExtensionProgress();
	}
}

// base class override
void CFilteredToDoCtrl::RebuildList(const void* pContext)
{
	if (pContext)
		CTabbedToDoCtrl::RebuildList(pContext);
	else
		RefreshListFilter();
}

void CFilteredToDoCtrl::RefreshListFilter() 
{
	GetViewData2(FTCV_TASKLIST)->bNeedRefilter = FALSE;

	// build a find query that matches the filter
	SEARCHPARAMS filter;
	BuildFilterQuery(filter);

	// rebuild the list
	RebuildList(filter);

	// modify the list prompt depending on whether there is a filter set
	UINT nPromptID = (IsFilterSet(FTCV_TASKLIST) ? IDS_TDC_FILTEREDTASKLISTPROMPT : IDS_TDC_TASKLISTPROMPT);
	m_taskList.SetWindowPrompt(CEnString(nPromptID));
}

void CFilteredToDoCtrl::RebuildList(const SEARCHPARAMS& filter)
{
	// since the tree will have already got the items we want 
	// we can optimize the rebuild under certain circumstances
	// which are: 
	// 1. the list is sorted
	// 2. or the tree is unsorted
	// otherwise we need the data in it's unsorted state and the 
	// tree doesn't have it
	if (m_taskList.IsSorting() || !m_taskTree.IsSorting())
	{
		// rebuild the list from the tree
		CTabbedToDoCtrl::RebuildList(&filter);
	}
	else // rebuild from scratch
	{
		// cache current selection
		TDCSELECTIONCACHE cache;
		CacheListSelection(cache);

		// grab the selected items for the filtering
		m_aSelectedTaskIDs.Copy(cache.aSelTaskIDs);

		// remove all existing items
		m_taskList.DeleteAll();

		// do the find
		CResultArray aResults;
		m_data.FindTasks(filter, aResults);

		// add tasks to list
		for (int nRes = 0; nRes < aResults.GetSize(); nRes++)
		{
			const SEARCHRESULT& res = aResults[nRes];

			// some more filtering required
			if (HasStyle(TDCS_ALWAYSHIDELISTPARENTS) || m_filter.HasFlag(FO_HIDEPARENTS))
			{
				const TODOSTRUCTURE* pTDS = m_data.LocateTask(res.dwTaskID);
				ASSERT(pTDS);

				if (pTDS->HasSubTasks())
					continue;
			}
			else if (m_filter.HasFlag(FO_HIDECOLLAPSED))
			{
				HTREEITEM hti = m_taskTree.GetItem(res.dwTaskID);
				ASSERT(hti);			

				if (m_taskTree.ItemHasChildren(hti) && !TCH().IsItemExpanded(hti))
					continue;
			}

			CTabbedToDoCtrl::AddItemToList(res.dwTaskID);
		}

		// restore selection
		RestoreListSelection(cache);
	}
}

void CFilteredToDoCtrl::AddTreeItemToList(HTREEITEM hti, const void* pContext)
{
	if (pContext == NULL)
	{
		CTabbedToDoCtrl::AddTreeItemToList(hti, NULL);
		return;
	}

	// else it's a filter
	const SEARCHPARAMS* pFilter = static_cast<const SEARCHPARAMS*>(pContext);

	if (hti)
	{
		BOOL bAdd = TRUE;
		DWORD dwTaskID = GetTaskID(hti);

		// check if parent items are to be ignored
		if (m_filter.HasFlag(FO_HIDEPARENTS) || HasStyle(TDCS_ALWAYSHIDELISTPARENTS))
		{
			// quick test first
			if (m_taskTree.ItemHasChildren(hti))
			{
				bAdd = FALSE;
			}
			else
			{
				const TODOSTRUCTURE* pTDS = m_data.LocateTask(dwTaskID);
				ASSERT(pTDS);

				bAdd = !pTDS->HasSubTasks();
			}
		}
		// else check if it's a parent item that's only present because
		// it has matching subtasks
		else if (m_taskTree.ItemHasChildren(hti))
		{
			const TODOSTRUCTURE* pTDS = m_data.LocateTask(dwTaskID);
			const TODOITEM* pTDI = GetTask(dwTaskID); 

			if (pTDI)
			{
				SEARCHRESULT result;

				// ie. check that parent actually matches
				bAdd = m_data.TaskMatches(pTDI, pTDS, *pFilter, result);
			}
		}

		if (bAdd)
			AddItemToList(dwTaskID);
	}

	// always check the children unless collapsed tasks ignored
	if (!m_filter.HasFlag(FO_HIDECOLLAPSED) || !hti || TCH().IsItemExpanded(hti))
	{
		HTREEITEM htiChild = m_taskTree.GetChildItem(hti);

		while (htiChild)
		{
			// check
			AddTreeItemToList(htiChild, pContext);
			htiChild = m_taskTree.GetNextItem(htiChild);
		}
	}
}

BOOL CFilteredToDoCtrl::SetStyles(const CTDCStylesMap& styles)
{
	if (CTabbedToDoCtrl::SetStyles(styles))
	{
		// do we need to re-filter?
		if (HasFilter() && GetViewData2(FTCV_TASKLIST)->bNeedRefilter)
		{
			RefreshTreeFilter(); // always

			if (InListView())
				RefreshListFilter();
		}

		return TRUE;
	}

	return FALSE;
}

BOOL CFilteredToDoCtrl::SetStyle(TDC_STYLE nStyle, BOOL bOn, BOOL bWantUpdate)
{
	// base class processing
	if (CTabbedToDoCtrl::SetStyle(nStyle, bOn, bWantUpdate))
	{
		// post-precessing
		switch (nStyle)
		{
		case TDCS_DUEHAVEHIGHESTPRIORITY:
		case TDCS_DONEHAVELOWESTPRIORITY:
		case TDCS_ALWAYSHIDELISTPARENTS:
			GetViewData2(FTCV_TASKLIST)->bNeedRefilter = TRUE;
			break;

		case TDCS_TREATSUBCOMPLETEDASDONE:
			if (m_filter.nShow != FS_ALL)
				GetViewData2(FTCV_TASKLIST)->bNeedRefilter = TRUE;
			break;
		}

		return TRUE;
	}

	return FALSE;
}

BOOL CFilteredToDoCtrl::InitFilterDate(FILTER_DATE nDate, const COleDateTime& dateUser, COleDateTime& date) const
{
	switch (nDate)
	{
	case FD_TODAY:
		date = CDateHelper::GetDate(DHD_TODAY);
		break;

	case FD_TOMORROW:
		date = CDateHelper::GetDate(DHD_TOMORROW);
		break;

	case FD_ENDTHISWEEK:
		date = CDateHelper::GetDate(DHD_ENDTHISWEEK);
		break;

	case FD_ENDNEXTWEEK: 
		date = CDateHelper::GetDate(DHD_ENDNEXTWEEK);
		break;

	case FD_ENDTHISMONTH:
		date = CDateHelper::GetDate(DHD_ENDTHISMONTH);
		break;

	case FD_ENDNEXTMONTH:
		date = CDateHelper::GetDate(DHD_ENDNEXTMONTH);
		break;

	case FD_ENDTHISYEAR:
		date = CDateHelper::GetDate(DHD_ENDTHISYEAR);
		break;

	case FD_ENDNEXTYEAR:
		date = CDateHelper::GetDate(DHD_ENDNEXTYEAR);
		break;

	case FD_NEXTSEVENDAYS:
		date = CDateHelper::GetDate(DHD_TODAY) + 6; // +6 because filter is FOP_ON_OR_BEFORE
		break;

	case FD_NOW:
		date = COleDateTime::GetCurrentTime();
		break;

	case FD_USER:
		ASSERT(CDateHelper::IsDateSet(dateUser));

		date = CDateHelper::GetDateOnly(dateUser);
		break;

	case FD_ANY:
	case FD_NONE:
		break;

	default:
		ASSERT(0);
		break;
	}

	return CDateHelper::IsDateSet(date);
}

BOOL CFilteredToDoCtrl::SplitSelectedTask(int nNumSubtasks)
{
   if (CTabbedToDoCtrl::SplitSelectedTask(nNumSubtasks))
   {
      if (InListView())
         RefreshListFilter();
 
      return TRUE;
   }
 
   return FALSE;
}

HTREEITEM CFilteredToDoCtrl::CreateNewTask(LPCTSTR szText, TDC_INSERTWHERE nWhere, BOOL bEditText)
{
	HTREEITEM htiNew = CTabbedToDoCtrl::CreateNewTask(szText, nWhere, bEditText);

	if (htiNew)
	{
		SetListNeedRefilter(!InListView());
		SetExtensionsNeedRefilter(TRUE, GetView());
	}

	return htiNew;
}

BOOL CFilteredToDoCtrl::CanCreateNewTask(TDC_INSERTWHERE nInsertWhere) const
{
	return CTabbedToDoCtrl::CanCreateNewTask(nInsertWhere);
}

void CFilteredToDoCtrl::SetModified(BOOL bMod, TDC_ATTRIBUTE nAttrib, DWORD dwModTaskID)
{
	CTabbedToDoCtrl::SetModified(bMod, nAttrib, dwModTaskID);

	if (bMod)
	{
		if (ModNeedsRefilter(nAttrib, FTCV_TASKTREE, dwModTaskID))
		{
			PostMessage(WM_TDC_REFRESHFILTER, FTCV_TASKTREE, (nAttrib == TDCA_UNDO));
		}
		else if (ModNeedsRefilter(nAttrib, FTCV_TASKLIST, dwModTaskID))
		{
			if (InListView())
				PostMessage(WM_TDC_REFRESHFILTER, FTCV_TASKLIST, (nAttrib == TDCA_UNDO));
			else
				GetViewData2(FTCV_TASKLIST)->bNeedRefilter = TRUE;
		}
	}
}

void CFilteredToDoCtrl::EndTimeTracking(BOOL bAllowConfirm)
{
	BOOL bWasTimeTracking = IsActivelyTimeTracking();
	
	CTabbedToDoCtrl::EndTimeTracking(bAllowConfirm);
	
	// do we need to refilter?
	if (bWasTimeTracking && m_bCustomFilter && m_customFilter.HasAttribute(TDCA_TIMESPENT))
	{
		RefreshFilter();
	}
}

BOOL CFilteredToDoCtrl::ModNeedsRefilter(TDC_ATTRIBUTE nModType, FTC_VIEW nView, DWORD dwModTaskID) const 
{
	// sanity checks
	if ((nModType == TDCA_NONE) || !HasStyle(TDCS_REFILTERONMODIFY))
		return FALSE;

	if (!HasFilter() && !HasCustomFilter())
		return FALSE;
	
	// we only need to refilter if the modified attribute
	// actually affects the filter
	BOOL bNeedRefilter = FALSE;

	if (m_bCustomFilter) // 'Find' filter
	{
		bNeedRefilter = m_customFilter.HasAttribute(nModType); 
		
		if (bNeedRefilter)
		{
			// don't refilter on Time Estimate/Spent, Cost or Comments, or 
			// similar custom attributes  because the user typically hasn't 
			// finished editing when this notification is first received
			switch (nModType)
			{
			case TDCA_TIMESPENT:
			case TDCA_TIMEEST:
			case TDCA_COST:
			case TDCA_COMMENTS:
				bNeedRefilter = FALSE;
				break;

			default:
				if (CTDCCustomAttributeHelper::IsCustomAttribute(nModType))
				{
					TDCCUSTOMATTRIBUTEDEFINITION attribDef;
					VERIFY(CTDCCustomAttributeHelper::GetAttributeDef(nModType, m_aCustomAttribDefs, attribDef));

					if (!attribDef.IsList())
					{
						switch (attribDef.GetDataType())
						{
						case TDCCA_DOUBLE:
						case TDCCA_INTEGER:
						case TDCCA_STRING:
							bNeedRefilter = FALSE;
							break;
						}
					}
				}
			}
		}
	}
	else if (HasFilter()) // 'Filter Bar' filter
	{
		switch (nModType)
		{
		case TDCA_TASKNAME:		
			bNeedRefilter = !m_filter.sTitle.IsEmpty(); 
			break;
			
		case TDCA_PRIORITY:		
			bNeedRefilter = (m_filter.nPriority != -1); 
			break;
			
		case TDCA_FLAG:		
			bNeedRefilter = (m_filter.nShow == FS_FLAGGED); 
			break;
			
		case TDCA_RISK:			
			bNeedRefilter = (m_filter.nRisk != -1);
			break;
			
		case TDCA_ALLOCBY:		
			bNeedRefilter = (m_filter.aAllocBy.GetSize() > 0);
			break;
			
		case TDCA_STATUS:		
			bNeedRefilter = (m_filter.aStatus.GetSize() > 0);
			break;
			
		case TDCA_VERSION:		
			bNeedRefilter = (m_filter.aVersions.GetSize() > 0);
			break;
			
		case TDCA_CATEGORY:		
			bNeedRefilter = (m_filter.aCategories.GetSize() > 0);
			break;
			
		case TDCA_TAGS:		
			bNeedRefilter = (m_filter.aTags.GetSize() > 0);
			break;
			
		case TDCA_ALLOCTO:		
			bNeedRefilter = (m_filter.aAllocTo.GetSize() > 0);
			break;
			
		case TDCA_PERCENT:
			bNeedRefilter = ((m_filter.nShow == FS_DONE) || (m_filter.nShow == FS_NOTDONE));
			break;
			
		case TDCA_DONEDATE:		
			// changing the DONE date requires refiltering if:
			bNeedRefilter = 
				// 1. The user wants to hide completed tasks
				(m_filter.HasFlag(FO_HIDEDONE) ||
				// 2. OR the user wants only completed tasks
				(m_filter.nShow == FS_DONE) || 
				// 3. OR the user wants only incomplete tasks
				(m_filter.nShow == FS_NOTDONE) ||
				// 4. OR a due date filter is active
				(m_filter.nDueBy != FD_ANY) ||
				// 5. OR a start date filter is active
				(m_filter.nStartBy != FD_ANY) ||
				// 6. OR the user is filtering on priority
				(m_filter.nPriority > 0));
			break;
			
		case TDCA_DUEDATE:		
			// changing the DUE date requires refiltering if:
			bNeedRefilter = 
				// 1. The user wants to hide overdue tasks
				((m_filter.HasFlag(FO_HIDEOVERDUE) ||
				// 2. OR the user is filtering on priority
				(m_filter.nPriority > 0) ||
				// 3. OR a due date filter is active
				(m_filter.nDueBy != FD_ANY)) &&
				// 3. AND the user doesn't want only completed tasks
				(m_filter.nShow != FS_DONE));
			break;
			
		case TDCA_STARTDATE:		
			// changing the START date requires refiltering if:
			bNeedRefilter = 
				// 1. A start date filter is active
				((m_filter.nStartBy != FD_ANY) &&
				// 2. AND the user doesn't want only completed tasks
				(m_filter.nShow != FS_DONE));
			break;
			
		default:
			// fall thru
			break;
		}
	}

	// handle attributes common to both filter types
	if (!bNeedRefilter)
	{
		switch (nModType)
		{
		case TDCA_NEWTASK:
			// if we refilter in the middle of adding a task it messes
			// up the tree items so we handle it in CreateNewTask
			break;
			
		case TDCA_DELETE:
			// this is handled in SetModified
			break;
			
		case TDCA_UNDO:
		case TDCA_PASTE:
			bNeedRefilter = TRUE;
			break;
			
		case TDCA_POSITION: // == move
			bNeedRefilter = (nView == FTCV_TASKLIST && !IsSorting());
			break;

		case TDCA_SELECTION:
			// never need to refilter
			break;
		}
	}

	// finally, if this was a task edit we can just test to 
	// see if the modified task still matches the filter.
	if (bNeedRefilter && dwModTaskID)
	{
		// VERY SPECIAL CASE
		// The task being time tracked has been filtered out
		// in which case we don't need to check if it matches
		if (dwModTaskID == m_dwTimeTrackTaskID)
		{
			if (m_taskTree.GetItem(dwModTaskID) == NULL)
			{
				ASSERT(HasTask(dwModTaskID));
				ASSERT(nModType == TDCA_TIMESPENT);

				return FALSE;
			}
			// else fall thru
		}

		SEARCHPARAMS params;
		SEARCHRESULT result;

		// This will handle custom and 'normal' filters
		BuildFilterQuery(params);
		bNeedRefilter = !m_data.TaskMatches(dwModTaskID, params, result);
		
		// extra handling for 'Find Tasks' filters 
		if (bNeedRefilter && m_bCustomFilter)
		{
			// don't refilter on Time Spent if time tracking
			bNeedRefilter = !(nModType == TDCA_TIMESPENT && IsActivelyTimeTracking());
		}
	}

	return bNeedRefilter;
}

void CFilteredToDoCtrl::Sort(TDC_COLUMN nBy, BOOL bAllowToggle)
{
	CTabbedToDoCtrl::Sort(nBy, bAllowToggle);
}

VIEWDATA2* CFilteredToDoCtrl::GetViewData2(FTC_VIEW nView) const
{
	return (VIEWDATA2*)CTabbedToDoCtrl::GetViewData(nView);
}

VIEWDATA2* CFilteredToDoCtrl::GetActiveViewData2() const
{
	return GetViewData2(GetView());
}

void CFilteredToDoCtrl::OnTimerMidnight()
{
	CTabbedToDoCtrl::OnTimerMidnight();

	// don't re-filter delay-loaded tasklists
	if (IsDelayLoaded())
		return;
	
	if ((m_filter.nStartBy != FD_NONE && m_filter.nStartBy != FD_ANY) ||
		(m_filter.nDueBy != FD_NONE && m_filter.nDueBy != FD_ANY))
	{
		RefreshFilter();
	}
}

void CFilteredToDoCtrl::ResetNowFilterTimer()
{
	if (!HasCustomFilter() && ((m_filter.nDueBy == FD_NOW) || (m_filter.nStartBy == FD_NOW)))
		SetTimer(TIMER_NOWFILTER, ONE_MINUTE, NULL);
	else
		KillTimer(TIMER_NOWFILTER);
}

void CFilteredToDoCtrl::OnTimer(UINT nIDEvent) 
{
	AF_NOREENTRANT;
	
	switch (nIDEvent)
	{
	case TIMER_NOWFILTER:
		OnTimerNow();
		return;
	}

	CTabbedToDoCtrl::OnTimer(nIDEvent);
}

void CFilteredToDoCtrl::OnTimerNow()
{
	// Since this timer gets called every minute we have to
	// find an efficient way of detecting tasks that are
	// currently hidden but need to be shown
	
	// So first thing we do is find reasons not to:
	
	// We are hidden
	if (!IsWindowVisible())
	{
		TRACE(_T("CFilteredToDoCtrl::OnTimerNow eaten (Window not visible)\n"));
		return;
	}
	
	// We're already displaying all tasks
	if (m_taskTree.GetItemCount() == m_data.GetTaskCount())
	{
		TRACE(_T("CFilteredToDoCtrl::OnTimerNow eaten (All tasks showing)\n"));
		return;
	}
	
	// App is minimized or hidden
	if (AfxGetMainWnd()->IsIconic() || !AfxGetMainWnd()->IsWindowVisible())
	{
		TRACE(_T("CFilteredToDoCtrl::OnTimerNow eaten (App not visible)\n"));
		return;
	}
	
	// App is not the foreground window
	if (GetForegroundWindow() != AfxGetMainWnd())
	{
		TRACE(_T("CFilteredToDoCtrl::OnTimerNow eaten (App not active)\n"));
		return;
	}
	
	// iterate the full data looking for items not in the
	// tree and test them for inclusion in the filter
	CHTIMap htiMap;
	VERIFY (TCH().BuildHTIMap(htiMap) < (int)m_data.GetTaskCount());
	
	SEARCHPARAMS params;
	BuildFilterQuery(params);
	
	const TODOSTRUCTURE* pTDS = m_data.GetStructure();
	ASSERT(pTDS);
	
	if (FindNewNowFilterTasks(pTDS, params, htiMap))
	{
		BOOL bRefilter = FALSE;
		
		if (m_filter.nDueBy == FD_NOW)
		{
			bRefilter = (AfxMessageBox(IDS_DUEBYNOW_CONFIRMREFILTER, MB_YESNO | MB_ICONQUESTION) == IDYES);
		}
		else if (m_filter.nStartBy == FD_NOW)
		{
			bRefilter = (AfxMessageBox(IDS_STARTBYNOW_CONFIRMREFILTER, MB_YESNO | MB_ICONQUESTION) == IDYES);
		}
		
		if (bRefilter)
		{
			RefreshFilter();
		}
		else // make the timer 10 minutes so we don't annoy them too soon
		{
			SetTimer(TIMER_NOWFILTER, TEN_MINUTES, NULL);
		}
	}
}

BOOL CFilteredToDoCtrl::FindNewNowFilterTasks(const TODOSTRUCTURE* pTDS, const SEARCHPARAMS& params, const CHTIMap& htiMap) const
{
	ASSERT(pTDS);

	// process task
	if (!pTDS->IsRoot())
	{
		// is the task invisible?
		HTREEITEM htiDummy;
		DWORD dwTaskID = pTDS->GetTaskID();

		if (!htiMap.Lookup(dwTaskID, htiDummy))
		{
			// does the task match the current filter
			SEARCHRESULT result;

			// This will handle custom and 'normal' filters
			if (m_data.TaskMatches(dwTaskID, params, result))
				return TRUE;
		}
	}

	// then children
	for (int nTask = 0; nTask < pTDS->GetSubTaskCount(); nTask++)
	{
		if (FindNewNowFilterTasks(pTDS->GetSubTask(nTask), params, htiMap))
			return TRUE;
	}

	// no new tasks
	return FALSE;
}
