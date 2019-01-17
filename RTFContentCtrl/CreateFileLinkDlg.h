#if !defined(AFX_CREATEFILELINKDLG_H__F70FEFCC_D6D8_40D7_8FB3_ACE6CD2DEEDB__INCLUDED_)
#define AFX_CREATEFILELINKDLG_H__F70FEFCC_D6D8_40D7_8FB3_ACE6CD2DEEDB__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// CreateFileLinkDlg.h : header file
//

#include "..\shared\RichEditHelper.h"

/////////////////////////////////////////////////////////////////////////////
// CCreateFileLinkDlg dialog

class CCreateFileLinkDlg : public CDialog
{
// Construction
public:
	CCreateFileLinkDlg(LPCTSTR szRefFile, RE_PASTE nLinkOption, BOOL bDefault, BOOL bPreferences = FALSE, CWnd* pParent = NULL);   // standard constructor

	RE_PASTE GetLinkOption() const { return m_nLinkOption; }
	BOOL GetMakeLinkOptionDefault() const { return m_bMakeDefault; }

protected:
// Dialog Data
	//{{AFX_DATA(CCreateFileLinkDlg)
	BOOL	m_bMakeDefault;
	//}}AFX_DATA
	RE_PASTE m_nLinkOption;
	CString m_sRefFile;
	BOOL m_bPreferences;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCreateFileLinkDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CCreateFileLinkDlg)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CREATEFILELINKDLG_H__F70FEFCC_D6D8_40D7_8FB3_ACE6CD2DEEDB__INCLUDED_)
