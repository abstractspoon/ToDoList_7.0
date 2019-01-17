#if !defined(AFX_TODOCTRLREMINDERS_H__E0CF538D_61A2_418B_8702_85A1BFFE05BF__INCLUDED_)
#define AFX_TODOCTRLREMINDERS_H__E0CF538D_61A2_418B_8702_85A1BFFE05BF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ToDoCtrlReminders.h : header file
//

#include <afxtempl.h>
#include "tdcenum.h"

#include "..\shared\StickiesWnd.h"

/////////////////////////////////////////////////////////////////////////////

const UINT WM_TD_REMINDER = ::RegisterWindowMessage(_T("WM_TD_REMINDER"));

/////////////////////////////////////////////////////////////////////////////

class CFilteredToDoCtrl;
class IPreferences;

/////////////////////////////////////////////////////////////////////////////

struct TDCREMINDER
{
	TDCREMINDER();

	CString FormatWhenString() const;
	CString GetTaskTitle() const;
	CString GetTaskComments() const;

	void Save(IPreferences* pPrefs, LPCTSTR szKey) const;
	void Load(const IPreferences* pPrefs, LPCTSTR szKey);

	DWORD dwTaskID;
	const CFilteredToDoCtrl* pTDC;
	COleDateTime dtAbsolute;
	BOOL bRelative;
	double dRelativeDaysLeadIn;
	double dDaysSnooze;
	TDC_REMINDER nRelativeFromWhen;
	BOOL bEnabled;
	CString sSoundFile;
	CString sStickiesID;
	int nLastSnoozeMins;
};

/////////////////////////////////////////////////////////////////////////////
// CToDoCtrlReminders window

class CToDoCtrlReminders : protected CStickiesWnd
{
// Construction
public:
	CToDoCtrlReminders();

	BOOL Initialize(CWnd* pNotify);
	BOOL UseStickies(BOOL bEnable, LPCTSTR szStickiesPath = NULL);

	void AddToDoCtrl(const CFilteredToDoCtrl& tdc);
	void CloseToDoCtrl(const CFilteredToDoCtrl& tdc);
	void SetReminder(const TDCREMINDER& rem);
	void RemoveReminder(const TDCREMINDER& rem);
	void RemoveReminder(DWORD dwTaskID, const CFilteredToDoCtrl* pTDC);
	BOOL GetReminder(int nRem, TDCREMINDER& rem) const;
	int FindReminder(const TDCREMINDER& rem, BOOL bIncludeDisabled = TRUE) const;
	int FindReminder(DWORD dwTaskID, const CFilteredToDoCtrl* pTDC, BOOL bIncludeDisabled = TRUE) const;
	BOOL ToDoCtrlHasReminders(const CFilteredToDoCtrl& tdc);
	BOOL ToDoCtrlHasReminders(const CString& sFilePath);
	void RemoveDeletedTaskReminders(const CFilteredToDoCtrl* pTDC = NULL);

	static CString FormatWhenString(const TDCREMINDER& rem);

// Attributes
protected:
	CWnd* m_pWndNotify;
	CArray<TDCREMINDER, TDCREMINDER&> m_aReminders;
	BOOL m_bUseStickies;
	CString m_sStickiesPath;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CToDoCtrlReminders)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CToDoCtrlReminders();

	// Generated message map functions
protected:
	//{{AFX_MSG(CToDoCtrlReminders)
	afx_msg void OnTimer(UINT nIDEvent);
	//}}AFX_MSG
	afx_msg void OnDestroy();
	DECLARE_MESSAGE_MAP()

protected:
	void SaveAndRemoveReminders(const CFilteredToDoCtrl& tdc);
	void LoadReminders(const CFilteredToDoCtrl& tdc);
	void StartTimer();
	LRESULT SendParentReminder(const TDCREMINDER& rem);

	static COleDateTime GetReminderDate(const TDCREMINDER& rem);

};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TODOCTRLREMINDERS_H__E0CF538D_61A2_418B_8702_85A1BFFE05BF__INCLUDED_)
