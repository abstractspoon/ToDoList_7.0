// TDLTransformDialog.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLTransformDialog.h"

#include "..\shared\enstring.h"
#include "..\shared\preferences.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTDLTransformDialog dialog


CTDLTransformDialog::CTDLTransformDialog(LPCTSTR szTitle, FTC_VIEW nView, CWnd* pParent /*=NULL*/)
	: CDialog(IDD_TRANSFORM_DIALOG, pParent), m_taskSel(_T("Transform"), nView),
		m_sTitle(szTitle), m_eStylesheet(FES_COMBOSTYLEBTN | FES_RELATIVEPATHS, CEnString(IDS_XSLFILEFILTER))

{
	//{{AFX_DATA_INIT(CTDLTransformDialog)
	//}}AFX_DATA_INIT
	// see what we had last time
	CPreferences prefs;

	// init the stylesheet folder to point to the resource folder
	CString sFolder = FileMisc::GetAppResourceFolder() + _T("\\Stylesheets");
	
	m_eStylesheet.SetCurrentFolder(sFolder);
	m_eStylesheet.SetBrowseTitle(CEnString(IDS_SELECTSTYLESHEET_TITLE));	
	
	m_sStylesheet = prefs.GetProfileString(_T("Transform"), _T("Stylesheet"));
	m_sStylesheet = FileMisc::GetRelativePath(m_sStylesheet, sFolder, FALSE);

	m_bDate = prefs.GetProfileInt(_T("Transform"), _T("WantDate"), TRUE);

	// share same title history as print dialog
	m_cbTitle.Load(&prefs, _T("Print"));
}


void CTDLTransformDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTDLTransformDialog)
	DDX_Control(pDX, IDC_STYLESHEET, m_eStylesheet);
	DDX_Text(pDX, IDC_STYLESHEET, m_sStylesheet);
	//}}AFX_DATA_MAP
	DDX_Text(pDX, IDC_TRANSFORMTITLE, m_sTitle);
	DDX_Check(pDX, IDC_TRANSFORMDATE, m_bDate);
	DDX_Control(pDX, IDC_TRANSFORMTITLE, m_cbTitle);
}


BEGIN_MESSAGE_MAP(CTDLTransformDialog, CDialog)
	//{{AFX_MSG_MAP(CTDLTransformDialog)
	ON_EN_CHANGE(IDC_STYLESHEET, OnChangeStylesheet)
	//}}AFX_MSG_MAP
//	ON_BN_CLICKED(IDC_USESTYLESHEET, OnUsestylesheet)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLTransformDialog message handlers

void CTDLTransformDialog::OnOK() 
{
	CDialog::OnOK();

	CPreferences prefs;

	prefs.WriteProfileString(_T("Transform"), _T("Stylesheet"), m_sStylesheet);
	prefs.WriteProfileInt(_T("Transform"), _T("WantDate"), m_bDate);

	// share same title history as print dialog
	m_cbTitle.Save(&prefs, _T("Print"));
}

BOOL CTDLTransformDialog::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
    VERIFY(m_taskSel.Create(IDC_FRAME, this));

	BOOL bEnable = FileMisc::FileExists(GetStylesheet());
	GetDlgItem(IDOK)->EnableWindow(bEnable);
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CTDLTransformDialog::OnChangeStylesheet() 
{
	UpdateData();

	BOOL bEnable = FileMisc::FileExists(GetStylesheet());
	GetDlgItem(IDOK)->EnableWindow(bEnable);
}

CString CTDLTransformDialog::GetStylesheet() const 
{ 
	return FileMisc::GetFullPath(m_sStylesheet, FileMisc::GetAppResourceFolder() + _T("\\Stylesheets")); 
}
