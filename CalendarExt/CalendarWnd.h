#if !defined(AFX_CALENDARWND_H__47616F96_0B5B_4F86_97A2_93B9DC796EB4__INCLUDED_)
#define AFX_CALENDARWND_H__47616F96_0B5B_4F86_97A2_93B9DC796EB4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// CalendarWnd.h : header file
//

#include "TaskCalendarCtrl.h"
#include "calendarpreferencesdlg.h"

#include "..\Shared\uitheme.h"
#include "..\Shared\menubutton.h"
#include "..\Shared\enstring.h"

#include "..\Interfaces\IUIExtension.h"

#include "..\3rdparty\fpsminicalendarctrl.h"

class CCalendarData;
struct UITHEME;

/////////////////////////////////////////////////////////////////////////////
// CCalendarWnd frame

class CCalendarWnd : public CDialog, public IUIExtensionWindow
{
	DECLARE_DYNAMIC(CCalendarWnd)

public:
	CCalendarWnd();
	virtual ~CCalendarWnd();

	BOOL Create(DWORD dwStyle, const RECT &rect, CWnd* pParentWnd, UINT nID);
	void Release();

	// IUIExtensionWindow
	LPCTSTR GetTypeID() const { return CAL_TYPEID; }
	HICON GetIcon() const { return m_hIcon; }
	LPCTSTR GetMenuText() const;

	void SetReadOnly(bool bReadOnly);
	HWND GetHwnd() const { return GetSafeHwnd(); }
	void SetUITheme(const UITHEME* pTheme);

	bool GetLabelEditRect(LPRECT pEdit);
	IUI_HITTEST HitTest(const POINT& ptScreen) const;

	void LoadPreferences(const IPreferences* pPrefs, LPCTSTR szKey, BOOL bAppOnly);
	void SavePreferences(IPreferences* pPrefs, LPCTSTR szKey) const;

	void UpdateTasks(const ITaskList* pTasks, IUI_UPDATETYPE nUpdate, int nEditAttribute);
	bool WantUpdate(int nAttribute) const;
	bool PrepareNewTask(ITaskList* pTask) const;

	bool SelectTask(DWORD dwTaskID);
	bool SelectTasks(DWORD* pdwTaskIDs, int nTaskCount);

	bool ProcessMessage(MSG* pMsg);
	void DoAppCommand(IUI_APPCOMMAND nCmd, DWORD dwExtra);
	bool CanDoAppCommand(IUI_APPCOMMAND nCmd, DWORD dwExtra) const;

//protected member variables
protected:
	CTaskCalendarCtrl m_BigCalendar;
	CFPSMiniCalendarCtrl m_MiniCalendar;
	CComboBox m_cbNumWeeks;
	CCalendarPreferencesDlg m_dlgPrefs;

	CBrush m_brBack;
	UITHEME m_theme;
	HICON m_hIcon;
	BOOL m_bReadOnly;

	CString m_sSelectedTaskDates;
	CComboBox m_cbSnapModes;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCalendarWnd)
	protected:
	//}}AFX_VIRTUAL
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	virtual void OnCancel();
	virtual void OnOK() {}

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CCalendarWnd)
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnGototoday();
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	//}}AFX_MSG
	afx_msg void OnPreferences();
	afx_msg void OnSelChangeNumWeeks();
	afx_msg void OnSelChangeSnapMode();
	afx_msg void OnBigCalendarNotifyClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnBigCalendarNotifyDblClk(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnMiniCalendarNotifyClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnMiniCalendarNotifyDblClk(NMHDR* pNMHDR, LRESULT* pResult);
	
	afx_msg LRESULT OnBigCalendarNotifyDateChange(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnBigCalendarNotifySelectionChange(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnBigCalendarNotifyDragChange(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnBigCalendarNotifyVisibleWeekChange(WPARAM wp, LPARAM lp);
	DECLARE_MESSAGE_MAP()

//protected member functions
protected:
	void ResizeControls(int cx, int cy);
	void UpdateSelectedTaskDates();
	void InitSnapCombo();
	void UpdateCalendarCtrlPreferences();

	static BOOL CALLBACK IsMiniCalSpecialDateCallback(COleDateTime &dt, DWORD dwUserData);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CALENDARFRAMEWND_H__47616F96_0B5B_4F86_97A2_93B9DC796EB4__INCLUDED_)
