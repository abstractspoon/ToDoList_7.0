// TDLPrefMigrationDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLPrefMigrationDlg.h"

#include "..\shared\filemisc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTDLPrefMigrationDlg dialog


CTDLPrefMigrationDlg::CTDLPrefMigrationDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTDLPrefMigrationDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTDLPrefMigrationDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CTDLPrefMigrationDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTDLPrefMigrationDlg)
	DDX_Control(pDX, IDC_ANIMATION, m_animation);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CTDLPrefMigrationDlg, CDialog)
	//{{AFX_MSG_MAP(CTDLPrefMigrationDlg)
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLPrefMigrationDlg message handlers

BOOL CTDLPrefMigrationDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	m_sAnimationFilePath = FileMisc::GetTempFilePath(_T("tdl_migrate"), _T("avi"));

	if (FileMisc::ExtractResource(_T("Shell32.dll"), 161, _T("AVI"), m_sAnimationFilePath))
		m_animation.Open(m_sAnimationFilePath);

	CenterWindow();
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

BOOL CTDLPrefMigrationDlg::Create(CWnd* pParentWnd) 
{
	return CDialog::Create(IDD, pParentWnd);
}

void CTDLPrefMigrationDlg::OnDestroy() 
{
	CDialog::OnDestroy();
	
	FileMisc::DeleteFile(m_sAnimationFilePath, TRUE);	
}
