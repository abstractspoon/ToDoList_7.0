// TDLLanguageDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLLanguageDlg.h"

#include "..\shared\graphicsmisc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTDLLanguageDlg dialog

CTDLLanguageDlg::CTDLLanguageDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTDLLanguageDlg::IDD, pParent), m_cbLanguages(_T("*.csv"))
{
	//{{AFX_DATA_INIT(CTDLLanguageDlg)
	//}}AFX_DATA_INIT
}

void CTDLLanguageDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTDLLanguageDlg)
	DDX_Control(pDX, IDC_LANGUAGES, m_cbLanguages);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CTDLLanguageDlg, CDialog)
	//{{AFX_MSG_MAP(CTDLLanguageDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLLanguageDlg message handlers

BOOL CTDLLanguageDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	HICON hIcon = GraphicsMisc::LoadIcon(IDR_MAINFRAME);
	SetIcon(hIcon, FALSE);

	m_cbLanguages.SetFocus();
	
	return FALSE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

CString CTDLLanguageDlg::GetLanguageFile() const 
{ 
	return m_cbLanguages.GetLanguageFile(); 
}

CString CTDLLanguageDlg::GetDefaultLanguage() 
{ 
	return CTDLLanguageComboBox::GetDefaultLanguage(); 
}
