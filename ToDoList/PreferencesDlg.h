#if !defined(AFX_PREFERENCESDLG_H__C3FCC72A_6C69_49A6_930D_D5C94EC31298__INCLUDED_)
#define AFX_PREFERENCESDLG_H__C3FCC72A_6C69_49A6_930D_D5C94EC31298__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// PreferencesDlg.h : header file
//

#include "..\shared\preferencesbase.h"
#include "..\shared\uithemefile.h"
#include "..\shared\enstatic.h"

#include "preferencesgenpage.h"
#include "preferencestaskpage.h"
#include "preferencestaskdefpage.h"
#include "preferencestoolpage.h"
#include "preferencesuivisibilitypage.h"
#include "preferencesuipage.h"
#include "preferencesuitasklistpage.h"
#include "preferencesuitasklistcolorspage.h"
#include "preferencesshortcutspage.h"
#include "preferencesfilepage.h"
#include "preferencesexportpage.h"
#include "preferencesmultiuserpage.h"
#include "preferencesfile2page.h"
#include "preferencestaskcalcpage.h"

// matches order of pages
enum 
{ 
	PREFPAGE_GEN, 
	PREFPAGE_MULTIUSER, 
	PREFPAGE_FILE, 
	PREFPAGE_FILE2, 
	PREFPAGE_UI, 
	PREFPAGE_UITASK, 
	PREFPAGE_UIFONTCOLOR, 
	PREFPAGE_TASK, 
	PREFPAGE_TASKCALC, 
	PREFPAGE_TASKDEF, 
	PREFPAGE_IMPEXP, 
	PREFPAGE_TOOL, 
	PREFPAGE_SHORTCUT 
};

/////////////////////////////////////////////////////////////////////////////
// CPreferencesDlg dialog

class CPreferencesDlg : public CPreferencesDlgBase
{
// Construction
public:
	CPreferencesDlg(CShortcutManager* pShortcutMgr = NULL, 
					const CContentMgr* pContentMgr = NULL, 
					const CImportExportMgr* pExportMgr = NULL, 
					const CUIExtensionMgr* pMgrUIExt = NULL,
					CWnd* pParent = NULL);   // standard constructor
	virtual ~CPreferencesDlg();

	void SetUITheme(const CUIThemeFile& theme);

	// CPreferencesGenPage
	BOOL GetAlwaysOnTop() const { return m_pageGen.GetAlwaysOnTop(); }
	BOOL GetUseSysTray() const { return m_pageGen.GetUseSysTray(); }
	BOOL GetConfirmDelete() const { return m_pageGen.GetConfirmDelete(); }
	BOOL GetConfirmSaveOnExit() const { return m_pageGen.GetConfirmSaveOnExit(); }
	BOOL GetShowOnStartup() const { return m_pageGen.GetShowOnStartup(); }
	int GetSysTrayOption() const { return m_pageGen.GetSysTrayOption(); }
	BOOL GetToggleTrayVisibility() const { return m_pageGen.GetToggleTrayVisibility(); }
	BOOL GetMultiInstance() const { return m_pageGen.GetMultiInstance(); }
	BOOL GetGlobalHotkey() const { return m_pageGen.GetGlobalHotkey(); }
	BOOL GetAddFilesToMRU() const { return m_pageGen.GetAddFilesToMRU(); }
	BOOL GetEnableTDLExtension() const { return m_pageGen.GetEnableTDLExtension(); }
	BOOL GetEnableTDLProtocol() const { return m_pageGen.GetEnableTDLProtocol(); }
	BOOL GetAutoCheckForUpdates() const { return m_pageGen.GetAutoCheckForUpdates(); }
	BOOL GetEscapeMinimizes() const { return m_pageGen.GetEscapeMinimizes(); }
	BOOL GetEnableDelayedLoading() const { return m_pageGen.GetEnableDelayedLoading(); }
	CString GetLanguageFile() const { return m_pageGen.GetLanguageFile(); }
	BOOL GetSaveStoragePasswords() const { return m_pageGen.GetSaveStoragePasswords(); }
	int GetAutoMinimizeFrequency() const { return m_pageGen.GetAutoMinimizeFrequency(); }
	BOOL GetUseStickies(CString& sStickiesPath) const { return m_pageGen.GetUseStickies(sStickiesPath); }
	BOOL GetReloadTasklists() const { return m_pageGen.GetReloadTasklists(); }
	BOOL GetEnableRTLInput() const { return m_pageGen.GetEnableRTLInput(); }

	// CPreferencesMultiUserPage
	int GetReadonlyReloadOption() const { return m_pageMultiUser.GetReadonlyReloadOption(); }
	int GetTimestampReloadOption() const { return m_pageMultiUser.GetTimestampReloadOption(); }
	BOOL GetAutoCheckOut() const { return m_pageMultiUser.GetAutoCheckOut(); }
	BOOL GetKeepTryingToCheckout() const { return m_pageMultiUser.GetKeepTryingToCheckout(); }
	BOOL GetCheckinOnClose() const { return m_pageMultiUser.GetCheckinOnClose(); }
	int GetRemoteFileCheckFrequency() const { return m_pageMultiUser.GetRemoteFileCheckFrequency(); }
	int GetAutoCheckinFrequency() const { return m_pageMultiUser.GetAutoCheckinFrequency(); }
	BOOL GetUsing3rdPartySourceControl() const { return m_pageMultiUser.GetUsing3rdPartySourceControl(); }
	BOOL GetIncludeUserNameInCheckout() const { return m_pageMultiUser.GetIncludeUserNameInCheckout(); }

	// CPreferencesFilePage
	int GetNotifyDueByOnLoad() const { return m_pageFile.GetNotifyDueByOnLoad(); }
	int GetNotifyDueByOnSwitch() const { return m_pageFile.GetNotifyDueByOnSwitch(); }
	BOOL GetDisplayDueTasksInHtml() const { return m_pageFile.GetDisplayDueTasksInHtml(); }
	BOOL GetDisplayDueCommentsInHtml() const { return m_pageFile.GetDisplayDueCommentsInHtml(); }
	BOOL GetAutoArchive() const { return m_pageFile.GetAutoArchive(); }
	BOOL GetNotifyReadOnly() const { return m_pageFile.GetNotifyReadOnly(); }
	BOOL GetRemoveArchivedTasks() const { return m_pageFile.GetRemoveArchivedTasks(); }
	BOOL GetRemoveOnlyOnAbsoluteCompletion() const { return m_pageFile.GetRemoveOnlyOnAbsoluteCompletion(); }
	BOOL GetRefreshFindOnLoad() const { return m_pageFile.GetRefreshFindOnLoad(); }
	BOOL GetDueTaskTitlesOnly() const { return m_pageFile.GetDueTaskTitlesOnly(); }
	CString GetDueTaskStylesheet() const { return m_pageFile.GetDueTaskStylesheet(); }
	CString GetDueTaskPerson() const { return m_pageFile.GetDueTaskPerson(); }
	BOOL GetWarnAddDeleteArchive() const { return m_pageFile.GetWarnAddDeleteArchive(); }
	BOOL GetDontRemoveFlagged() const { return m_pageFile.GetDontRemoveFlagged(); }
	BOOL GetExpandTasksOnLoad() const { return m_pageFile.GetExpandTasksOnLoad(); }

	// CPreferencesFile2Page
	BOOL GetBackupOnSave() const { return m_pageFile2.GetBackupOnSave(); }
	CString GetBackupLocation(LPCTSTR szFilePath) const { return m_pageFile2.GetBackupLocation(szFilePath); }
	int GetKeepBackupCount() const { return m_pageFile2.GetKeepBackupCount(); }
	BOOL GetAutoExport() const { return m_pageFile2.GetAutoExport(); }
	BOOL GetExportToHTML() const { return m_pageFile2.GetExportToHTML(); }
	int GetOtherExporter() const { return m_pageFile2.GetOtherExporter(); }
	int GetAutoSaveFrequency() const { return m_pageFile2.GetAutoSaveFrequency(); }
	CString GetAutoExportFolderPath() const { return m_pageFile2.GetAutoExportFolderPath(); }
	CString GetSaveExportStylesheet() const { return m_pageFile2.GetSaveExportStylesheet(); }
	BOOL GetAutoSaveOnSwitchTasklist() const { return m_pageFile2.GetAutoSaveOnSwitchTasklist(); }
	BOOL GetAutoSaveOnSwitchApp() const { return m_pageFile2.GetAutoSaveOnSwitchApp(); }
	BOOL GetExportFilteredOnly() const { return m_pageFile2.GetExportFilteredOnly(); }
	BOOL GetAutoSaveOnRunTools() const { return m_pageFile2.GetAutoSaveOnRunTools(); }

	// CPreferencesExportPage
	CString GetHtmlFont() const { return m_pageExport.GetHtmlFont(); }
	int GetHtmlFontSize() const { return m_pageExport.GetHtmlFontSize(); }
	BOOL GetPreviewExport() const { return m_pageExport.GetPreviewExport(); }
	int GetTextIndent() const { return m_pageExport.GetTextIndent(); }
	BOOL GetExportParentTitleCommentsOnly() const { return m_pageExport.GetExportParentTitleCommentsOnly(); }
	BOOL GetExportSpaceForNotes() const { return m_pageExport.GetExportSpaceForNotes(); }
	BOOL GetExportVisibleColsOnly() const { return m_pageExport.GetExportVisibleColsOnly(); }
	BOOL GetExportAllAttributes() const { return !GetExportVisibleColsOnly(); }
	CString GetHtmlCharSet() const { return m_pageExport.GetHtmlCharSet(); }
	
	// CPreferencesTaskDefPage
	void GetDefaultTaskAttributes(TODOITEM& tdiDefault) const;
	int GetParentAttribsUsed(CTDCAttributeArray& aAttribs, BOOL& bUpdateAttrib) const { return m_pageTaskDef.GetParentAttribsUsed(aAttribs, bUpdateAttrib); }
	int GetDefaultListItems(TDC_ATTRIBUTE nList, CStringArray& aItems) const { return m_pageTaskDef.GetListItems(nList, aItems); }
	BOOL GetDefaultListIsReadonly(TDC_ATTRIBUTE nList) const { return m_pageTaskDef.GetListIsReadonly(nList); }

	// CPreferencesTaskPage
	BOOL GetTrackNonActiveTasklists() const { return m_pageTask.GetTrackNonActiveTasklists(); }
	BOOL GetTrackNonSelectedTasks() const { return m_pageTask.GetTrackNonSelectedTasks(); }
	BOOL GetTrackOnScreenSaver() const { return m_pageTask.GetTrackOnScreenSaver(); }
	BOOL GetTrackHibernated() const { return m_pageTask.GetTrackHibernated(); }
	double GetHoursInOneDay() const { return m_pageTask.GetHoursInOneDay(); }
	double GetDaysInOneWeek() const { return m_pageTask.GetDaysInOneWeek(); }
	BOOL GetLogTimeTracking() const { return m_pageTask.GetLogTimeTracking(); }
	BOOL GetLogTaskTimeSeparately() const { return m_pageTask.GetLogTaskTimeSeparately(); }
	BOOL GetExclusiveTimeTracking() const { return m_pageTask.GetExclusiveTimeTracking(); }
	BOOL GetAllowParentTimeTracking() const { return m_pageTask.GetAllowParentTimeTracking(); }
	DWORD GetWeekendDays() const { return m_pageTask.GetWeekendDays(); }
	BOOL GetAllowReferenceEditing() const { return m_pageTask.GetAllowReferenceEditing(); }
	BOOL GetDisplayLogConfirm() const { return m_pageTask.GetDisplayLogConfirm(); }

	// CPreferencesTaskCalcPage
	BOOL GetAveragePercentSubCompletion() const { return m_pageTaskCalc.GetAveragePercentSubCompletion(); }
	BOOL GetIncludeDoneInAverageCalc() const { return m_pageTaskCalc.GetIncludeDoneInAverageCalc(); }
	BOOL GetUsePercentDoneInTimeEst() const { return m_pageTaskCalc.GetUsePercentDoneInTimeEst(); }
	BOOL GetTreatSubCompletedAsDone() const { return m_pageTaskCalc.GetTreatSubCompletedAsDone(); }
	BOOL GetUseHighestPriority() const { return m_pageTaskCalc.GetUseHighestPriority(); }
	BOOL GetUseHighestRisk() const { return m_pageTaskCalc.GetUseHighestRisk(); }
	BOOL GetDueTasksHaveHighestPriority() const { return m_pageTaskCalc.GetDueTasksHaveHighestPriority(); }
	BOOL GetDoneTasksHaveLowestPriority() const { return m_pageTaskCalc.GetDoneTasksHaveLowestPriority(); }
	BOOL GetDoneTasksHaveLowestRisk() const { return m_pageTaskCalc.GetDoneTasksHaveLowestRisk(); } 
	BOOL GetAutoCalcTimeEstimates() const { return m_pageTaskCalc.GetAutoCalcTimeEstimates(); }
	BOOL GetIncludeDoneInPriorityRiskCalc() const { return m_pageTaskCalc.GetIncludeDoneInPriorityRiskCalc(); }
	BOOL GetWeightPercentCompletionByNumSubtasks() const { return m_pageTaskCalc.GetWeightPercentCompletionByNumSubtasks(); }
	BOOL GetAutoCalcPercentDone() const { return m_pageTaskCalc.GetAutoCalcPercentDone(); }
	BOOL GetAutoAdjustDependentsDates() const { return m_pageTaskCalc.GetAutoAdjustDependentsDates(); }
	BOOL GetNoDueDateIsDueToday() const { return m_pageTaskCalc.GetNoDueDateIsDueToday(); }
	BOOL GetCompletionStatus(CString& sStatus) const { return m_pageTaskCalc.GetCompletionStatus(sStatus); }
	COleDateTimeSpan GetRecentlyModifiedPeriod() const { return m_pageTaskCalc.GetRecentlyModifiedPeriod(); }

	PTCP_CALCTIMEREMAINING GetTimeRemainingCalculation() const { return m_pageTaskCalc.GetTimeRemainingCalculation(); }
	PTCP_CALCDUEDATE GetDueDateCalculation() const { return m_pageTaskCalc.GetDueDateCalculation(); }
	PTCP_CALCSTARTDATE GetStartDateCalculation() const { return m_pageTaskCalc.GetStartDateCalculation(); }

	// CPreferencesUIPage
	BOOL GetShowEditMenuAsColumns() const { return m_pageUI.GetShowEditMenuAsColumns(); }
	BOOL GetShowCommentsAlways() const { return m_pageUI.GetShowCommentsAlways(); }
	BOOL GetAutoReposCtrls() const { return m_pageUI.GetAutoReposCtrls(); }
	BOOL GetSharedCommentsHeight() const { return m_pageUI.GetSharedCommentsHeight(); }
	BOOL GetAutoHideTabbar() const { return m_pageUI.GetAutoHideTabbar(); }
	BOOL GetStackTabbarItems() const { return m_pageUI.GetStackTabbarItems(); }
	BOOL GetFocusTreeOnEnter() const { return m_pageUI.GetFocusTreeOnEnter(); }
	PUIP_NEWTASKPOS GetNewTaskPos() const { return m_pageUI.GetNewTaskPos(); }
	PUIP_NEWTASKPOS GetNewSubtaskPos() const { return m_pageUI.GetNewSubtaskPos(); }
	BOOL GetKeepTabsOrdered() const { return m_pageUI.GetKeepTabsOrdered(); }
	BOOL GetShowTasklistCloseButton() const { return m_pageUI.GetShowTasklistCloseButton(); }
	BOOL GetSortDoneTasksAtBottom() const { return m_pageUI.GetSortDoneTasksAtBottom(); }
	BOOL GetRTLComments() const { return m_pageUI.GetRTLComments(); }
	PUIP_LOCATION GetCommentsPos() const { return m_pageUI.GetCommentsPos(); }
	PUIP_LOCATION GetControlsPos() const { return m_pageUI.GetControlsPos(); }
	BOOL GetMultiSelFilters() const { return m_pageUI.GetMultiSelFilters(); }
	BOOL GetRestoreTasklistFilters() const { return m_pageUI.GetRestoreTasklistFilters(); }
	BOOL GetReFilterOnModify() const { return m_pageUI.GetReFilterOnModify(); }
	BOOL GetReSortOnModify() const { return m_pageUI.GetReSortOnModify(); }
	CString GetUITheme() const { return m_pageUI.GetUITheme(); }
	BOOL GetEnableLightboxMgr() const { return m_pageUI.GetEnableLightboxMgr(); }
	PUIP_MATCHTITLE GetTitleFilterOption() const { return m_pageUI.GetTitleFilterOption(); }
	BOOL GetShowDefaultTaskIcons() const { return m_pageUI.GetShowDefaultTaskIcons(); }
	BOOL GetShowDefaultFilters() const { return m_pageUI.GetShowDefaultFilters(); }
	BOOL GetEnableColumnHeaderSorting() const { return m_pageUI.GetEnableColumnHeaderSorting(); }
	int GetDefaultTaskViews(CStringArray& aViews) const { return m_pageUI.GetDefaultTaskViews(aViews); }
	BOOL GetStackEditFieldsAndComments() const { return m_pageUI.GetStackEditFieldsAndComments(); }

	// CPreferencesUIVisibilityPage
	void GetDefaultColumnEditFilterVisibility(TDCCOLEDITFILTERVISIBILITY& vis) const { m_pageUIVisibility.GetColumnAttributeVisibility(vis); }
	void SetDefaultColumnEditFilterVisibility(const TDCCOLEDITFILTERVISIBILITY& vis) { m_pageUIVisibility.SetColumnAttributeVisibility(vis); }

	// CPreferencesUITasklistPage
	BOOL GetShowInfoTips() const { return m_pageUITasklist.GetShowInfoTips(); }
	BOOL GetShowComments() const { return m_pageUITasklist.GetShowComments(); }
	BOOL GetShowPathInHeader() const { return m_pageUITasklist.GetShowPathInHeader(); }
	BOOL GetStrikethroughDone() const { return m_pageUITasklist.GetStrikethroughDone(); }
	BOOL GetDisplayDatesInISO() const { return m_pageUITasklist.GetDisplayDatesInISO(); }
	BOOL GetShowWeekdayInDates() const { return m_pageUITasklist.GetShowWeekdayInDates(); }
	BOOL GetShowParentsAsFolders() const { return m_pageUITasklist.GetShowParentsAsFolders(); }
	BOOL GetDisplayFirstCommentLine() const { return m_pageUITasklist.GetDisplayFirstCommentLine(); }
	int GetMaxInfoTipCommentsLength() const { return m_pageUITasklist.GetMaxInfoTipCommentsLength(); }
	int GetMaxColumnWidth() const { return m_pageUITasklist.GetMaxColumnWidth(); }
	BOOL GetHidePercentForDoneTasks() const { return m_pageUITasklist.GetHidePercentForDoneTasks(); }
	BOOL GetHideZeroTimeCost() const { return m_pageUITasklist.GetHideZeroTimeCost(); }
	BOOL GetHideStartDueForDoneTasks() const { return m_pageUITasklist.GetHideStartDueForDoneTasks(); }
	BOOL GetShowPercentAsProgressbar() const { return m_pageUITasklist.GetShowPercentAsProgressbar(); }
	BOOL GetRoundTimeFractions() const { return m_pageUITasklist.GetRoundTimeFractions(); }
	BOOL GetShowNonFilesAsText() const { return m_pageUITasklist.GetShowNonFilesAsText(); }
	BOOL GetUseHMSTimeFormat() const { return m_pageUITasklist.GetUseHMSTimeFormat(); }
	BOOL GetAutoFocusTasklist() const { return m_pageUITasklist.GetAutoFocusTasklist(); }
	BOOL GetShowColumnsOnRight() const { return m_pageUITasklist.GetShowColumnsOnRight(); }
	BOOL GetAlwaysHideListParents() const { return m_pageUITasklist.GetAlwaysHideListParents(); }
	int GetPercentDoneIncrement() const { return m_pageUITasklist.GetPercentDoneIncrement(); }
	BOOL GetHideZeroPercentDone() const { return m_pageUITasklist.GetHideZeroPercentDone(); }
	CString GetDateTimePasteTrailingText() const { return m_pageUITasklist.GetDateTimePasteTrailingText(); }
	BOOL GetAppendUserToDateTimePaste() const { return m_pageUITasklist.GetAppendUserToDateTimePaste(); }
	BOOL GetAllowCheckboxAgainstTreeItem() const { return m_pageUITasklist.GetAllowCheckboxAgainstTreeItem(); }
	BOOL GetHidePaneSplitBar() const { return m_pageUITasklist.GetHidePaneSplitBar(); }

	// CPreferencesUITasklistColorsPage
	int GetTextColorOption() const { return m_pageUITasklistColors.GetTextColorOption(); }
	BOOL GetColorTaskBackground() const { return m_pageUITasklistColors.GetColorTaskBackground(); }
	BOOL GetCommentsUseTreeFont() const { return m_pageUITasklistColors.GetCommentsUseTreeFont(); }
	BOOL GetColorPriority() const { return m_pageUITasklistColors.GetColorPriority(); }
	int GetPriorityColors(CDWordArray& aColors) const { return m_pageUITasklistColors.GetPriorityColors(aColors); }
	void GetDueTaskColors(COLORREF& crDue, COLORREF& crDueToday) const { m_pageUITasklistColors.GetDueTaskColors(crDue, crDueToday); }
	void GetStartedTaskColors(COLORREF& crStarted, COLORREF& crStartedToday) const { m_pageUITasklistColors.GetStartedTaskColors(crStarted, crStartedToday); }
	BOOL GetTreeFont(CString& sFaceName, int& nPointSize) const { return m_pageUITasklistColors.GetTreeFont(sFaceName, nPointSize); }
	BOOL GetCommentsFont(CString& sFaceName, int& nPointSize) const { return m_pageUITasklistColors.GetCommentsFont(sFaceName, nPointSize); }
	COLORREF GetGridlineColor() const { return m_pageUITasklistColors.GetGridlineColor(); }
	COLORREF GetDoneTaskColor() const { return m_pageUITasklistColors.GetDoneTaskColor(); }
	COLORREF GetFlaggedTaskColor() const { return m_pageUITasklistColors.GetFlaggedTaskColor(); }
	COLORREF GetReferenceTaskColor() const { return m_pageUITasklistColors.GetReferenceTaskColor(); }
	COLORREF GetHidePriorityNumber() const { return m_pageUITasklistColors.GetHidePriorityNumber(); }
	COLORREF GetAlternateLineColor() const { return m_pageUITasklistColors.GetAlternateLineColor(); }
	int GetAttributeColors(TDC_ATTRIBUTE& nAttrib, CAttribColorArray& aColors) const { return m_pageUITasklistColors.GetAttributeColors(nAttrib, aColors); }

	// CPreferencesToolPage
	int GetUserTools(CUserToolArray& aTools) const { return m_pageTool.GetUserTools(aTools); }
	BOOL GetUserTool(int nTool, USERTOOL& tool) const { return m_pageTool.GetUserTool(nTool, tool); } 

//	BOOL Get() const { return m_b; }

protected:
// Dialog Data
	//{{AFX_DATA(CPreferencesDlg)
	//}}AFX_DATA
	CPreferencesGenPage m_pageGen;
	CPreferencesFilePage m_pageFile;
	CPreferencesFile2Page m_pageFile2;
	CPreferencesExportPage m_pageExport;
	CPreferencesUIPage m_pageUI;
	CPreferencesUIVisibilityPage m_pageUIVisibility;
	CPreferencesUITasklistPage m_pageUITasklist;
	CPreferencesUITasklistColorsPage m_pageUITasklistColors;
	CPreferencesTaskPage m_pageTask;
	CPreferencesTaskCalcPage m_pageTaskCalc;
	CPreferencesTaskDefPage m_pageTaskDef;
	CPreferencesToolPage m_pageTool;
	CPreferencesShortcutsPage m_pageShortcuts;
	CPreferencesMultiUserPage m_pageMultiUser;

	CSize m_sizeCurrent;
	CTreeCtrl m_tcPages;
	CString m_sPageTitle;
	CEnStatic m_stCategoryTitle, m_stPageTitle; 
	CUIThemeFile m_theme;

	CMap<CPreferencesPageBase*, CPreferencesPageBase*, HTREEITEM, HTREEITEM&> m_mapPP2HTI;
	CMap<HTREEITEM, HTREEITEM, UINT, UINT&> m_mapHTIToSection;


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPreferencesDlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CPreferencesDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnHelp();
	afx_msg void OnSelchangedPages(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnApply();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnGetMinMaxInfo(MINMAXINFO FAR* lpMMI);
	//}}AFX_MSG
	afx_msg LRESULT OnColourPageAttributeChange(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnToolPageTestTool(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnGenPageClearMRU(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnGenPageCleanupDictionary(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnControlChange(WPARAM wp, LPARAM lp);
	afx_msg BOOL OnHelpInfo(HELPINFO* pHelpInfo);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	DECLARE_MESSAGE_MAP()

protected:
	void AddPage(CPreferencesPageBase* pPage, UINT nIDPath, UINT nIDSection = 0);
	BOOL SetActivePage(int nPage); // override
	CString GetItemPath(HTREEITEM hti) const;
	void SynchronizeTree();
	BOOL IsPrefTab(HWND hWnd) const;
	void Resize(int cx = 0, int cy = 0);

	static void SetTitleThemeColors(CEnStatic& stTitle, const CUIThemeFile& theme);

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PREFERENCESDLG_H__C3FCC72A_6C69_49A6_930D_D5C94EC31298__INCLUDED_)
