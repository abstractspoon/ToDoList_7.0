// FilteredToDoCtrl.h: interface for the CFilteredToDoCtrl class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_FILTEREDTODOCTRL_H__356A6EB9_C7EC_4395_8716_623AFF4A269B__INCLUDED_)
#define AFX_FILTEREDTODOCTRL_H__356A6EB9_C7EC_4395_8716_623AFF4A269B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "TabbedToDoCtrl.h"
#include "..\shared\misc.h"

/////////////////////////////////////////////////////////////////////////////

struct VIEWDATA2 : public VIEWDATA
{
	VIEWDATA2() : bNeedRefilter(TRUE) {}
	virtual ~VIEWDATA2() {}

	BOOL bNeedRefilter;
};

/////////////////////////////////////////////////////////////////////////////

class CFilteredToDoCtrl : public CTabbedToDoCtrl  
{
public:
	CFilteredToDoCtrl(CUIExtensionMgr& mgrUIExt, 
						CContentMgr& mgrContent, 
						const CONTENTFORMAT& cfDefault,
						const TDCCOLEDITFILTERVISIBILITY& visDefault);
	
	virtual ~CFilteredToDoCtrl();

	BOOL DelayLoad(const CString& sFilePath, COleDateTime& dtEarliestDue);

	HTREEITEM CreateNewTask(LPCTSTR szText, TDC_INSERTWHERE nWhere = TDC_INSERTATTOPOFSELTASKPARENT, BOOL bEditText = TRUE); 
	BOOL CanCreateNewTask(TDC_INSERTWHERE nInsertWhere) const;

	BOOL ArchiveDoneTasks(TDC_ARCHIVE nFlags, BOOL bRemoveFlagged); // returns true if any tasks were removed
	BOOL ArchiveSelectedTasks(BOOL bRemove); // returns true if any tasks were removed

	void SetFilter(const FTDCFILTER& filter);
	FILTER_SHOW GetFilter(FTDCFILTER& filter) const;
	FILTER_SHOW GetFilter() const;
	void RefreshFilter();
	void ClearFilter();
	void ToggleFilter();
	BOOL HasFilter() const { return m_filter.IsSet(); }
	BOOL HasLastFilter() const { return m_lastFilter.IsSet() || m_bLastFilterWasCustom; }

	BOOL HasCustomFilter() const;
	BOOL SetCustomFilter(const SEARCHPARAMS& params, LPCTSTR szName);
	BOOL CustomFilterHasRules() const { return m_customFilter.GetRuleCount(); }
	BOOL RestoreCustomFilter();
	CString GetCustomFilterName() const;

	UINT GetTaskCount(UINT* pVisible = 0) const;
	int GetFilteredTasks(CTaskFile& tasks, const TDCGETTASKS& filter = TDCGT_ALL) const;
	int FindTasks(const SEARCHPARAMS& params, CResultArray& aResults) const;

	BOOL SplitSelectedTask(int nNumSubtasks);
	BOOL SetStyles(const CTDCStylesMap& styles);
	void Sort(TDC_COLUMN nBy, BOOL bAllowToggle = TRUE);
	void SetModified(BOOL bMod = TRUE) { CTabbedToDoCtrl::SetModified(bMod); }
	BOOL SetStyle(TDC_STYLE nStyle, BOOL bOn = TRUE) { return CTabbedToDoCtrl::SetStyle(nStyle, bOn); }

	virtual int GetArchivableTasks(CTaskFile& tasks, BOOL bSelectedOnly) const;
	virtual void EndTimeTracking(BOOL bClosing);

protected:
	enum // visible for derived classes
	{ 
		TIMER_NOWFILTER = TIMER_LAST, 
	};
	
	FTDCFILTER m_filter;
	FTDCFILTER m_lastFilter;
	SEARCHPARAMS m_customFilter;
	SEARCHPARAMS m_lastCustomFilter;
	BOOL m_bCustomFilter;
	BOOL m_bLastFilterWasCustom;
	CString m_sCustomFilter;
	CDWordArray m_aSelectedTaskIDs; // for FT_SELECTED filter

// Overrides
protected:
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFilteredToDoCtrl)
	//}}AFX_VIRTUAL
	virtual BOOL OnInitDialog();

protected:
	// Generated message map functions
	//{{AFX_MSG(CFilteredToDoCtrl)
	afx_msg void OnDestroy();
	afx_msg void OnTimer(UINT nIDEvent);
	//}}AFX_MSG
	afx_msg void OnTreeExpandItem(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg LRESULT OnRefreshFilter(WPARAM wParam, LPARAM lParam);
	afx_msg void OnEditChangeDueTime();
	afx_msg LRESULT OnPreTabViewChange(WPARAM nOldView, LPARAM nNewView);

	DECLARE_MESSAGE_MAP()

	// pseudo-message handlers
	void OnTimerNow();
	void OnTimerMidnight();

protected:
	BOOL InitFilterDate(FILTER_DATE nDate, const COleDateTime& dateUser, COleDateTime& dateDue) const;
	BOOL ModNeedsRefilter(TDC_ATTRIBUTE nModType, FTC_VIEW nView, DWORD dwModTaskID) const;

	virtual void SetModified(BOOL bMod, TDC_ATTRIBUTE nAttrib, DWORD dwModTaskID);
	virtual BOOL SetStyle(TDC_STYLE nStyle, BOOL bOn, BOOL bWantUpdate); // one style at a time only 

	virtual BOOL LoadTasks(const CTaskFile& file);
	void GetCompletedTasks(const TODOSTRUCTURE* pTDS, CTaskFile& tasks, HTASKITEM hTaskParent, BOOL bSelectedOnly = FALSE) const;
	virtual BOOL RemoveArchivedTask(DWORD dwTaskID);
	virtual BOOL CopyCurrentSelection() const;

	void BuildFilterQuery(SEARCHPARAMS& params) const;
	void AddNonDateFilterQueryRules(SEARCHPARAMS& params) const;
	
	void SaveSettings() const;
	void LoadSettings();

	void LoadFilter(const CPreferences& prefs, FTDCFILTER& filter, const CString& sSubKey = _T(""));
	void SaveFilter(const FTDCFILTER& filter, CPreferences& prefs, const CString& sSubKey = _T("")) const;
	void RefreshListFilter();
	void RefreshTreeFilter();
	void RefreshExtensionFilter(FTC_VIEW nView, BOOL bShowProgress = FALSE);
	void LoadCustomFilter(const CPreferences& prefs, CString& sFilter, SEARCHPARAMS& params, const CString& sSubKey = _T(""));
	void SaveCustomFilter(const CString& sFilter, const SEARCHPARAMS& params, CPreferences& prefs, const CString& sSubKey = _T("")) const;
	BOOL IsFilterSet(FTC_VIEW nView) const;
	BOOL IsFilterSet(const FTDCFILTER& filter, FTC_VIEW nView) const;
	BOOL FiltersMatch(const FTDCFILTER& filter1, const FTDCFILTER& filter2, FTC_VIEW nView) const;
	void SetExtensionsNeedRefilter(BOOL bRefilter, FTC_VIEW nIgnore = FTCV_UNSET);
	void SetListNeedRefilter(BOOL bRefilter);

	void ResetNowFilterTimer();
	BOOL FindNewNowFilterTasks(const TODOSTRUCTURE* pTDS, const SEARCHPARAMS& params, const CHTIMap& htiMap) const;

	VIEWDATA2* GetActiveViewData2() const;
	VIEWDATA2* GetViewData2(FTC_VIEW nView) const;

	HTREEITEM RebuildTree(const void* pContext = NULL);
	BOOL WantAddTask(const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS, const void* pContext) const; 

	void RebuildList(const void* pContext); 
	void RebuildList(const SEARCHPARAMS& filter);
	virtual void AddTreeItemToList(HTREEITEM hti, const void* pContext);

	virtual VIEWDATA* NewViewData() { return new VIEWDATA2(); }
};

#endif // !defined(AFX_FILTEREDTODOCTRL_H__356A6EB9_C7EC_4395_8716_623AFF4A269B__INCLUDED_)
