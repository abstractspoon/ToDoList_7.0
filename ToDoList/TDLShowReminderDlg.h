#if !defined(AFX_TDLSHOWREMINDERDLG_H__6B20922F_A2AA_4C95_B9E6_45F5EBEF18BF__INCLUDED_)
#define AFX_TDLSHOWREMINDERDLG_H__6B20922F_A2AA_4C95_B9E6_45F5EBEF18BF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// TDLShowReminderDlg.h : header file
//

#include "..\shared\enstatic.h"

// predec
struct TDCREMINDER;

/////////////////////////////////////////////////////////////////////////////
// CTDLShowReminderDlg dialog

class CTDLShowReminderDlg : public CDialog
{
// Construction
public:
	CTDLShowReminderDlg(CWnd* pParent = NULL);   // standard constructor

	int DoModal(const TDCREMINDER& rem);
	UINT GetSnoozeMinutes() const;
	double GetSnoozeDays() const;

// Dialog Data
	//{{AFX_DATA(CTDLShowReminderDlg)
	enum { IDD = IDD_SHOWREMINDER_DIALOG };
	CString	m_sWhen;
	CString	m_sTaskTitle;
	int		m_nSnoozeIndex;
	//}}AFX_DATA
	CString m_sSoundFile;
	UINT m_nSnoozeMins;
	CEnStatic m_stTaskTitle;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTDLShowReminderDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL
	virtual int OnInitDialog();

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CTDLShowReminderDlg)
	afx_msg void OnSnooze();
	afx_msg void OnGototask();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

protected:
	void SetSnoozeMinutes(UINT nMins);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TDLSHOWREMINDERDLG_H__6B20922F_A2AA_4C95_B9E6_45F5EBEF18BF__INCLUDED_)
