// TDLAddLoggedTimeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLAddLoggedTimeDlg.h"

#include "..\shared\DateHelper.h"
#include "..\Shared\localizer.h"
#include "..\Shared\preferences.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTDLAddLoggedTimeDlg dialog

CTDLAddLoggedTimeDlg::CTDLAddLoggedTimeDlg(DWORD dwTaskID, LPCTSTR szTaskTitle, double dHours, CWnd* pParent /*=NULL*/)
	: 
	CDialog(CTDLAddLoggedTimeDlg::IDD, pParent), 
	m_cbTimeWhen(TCB_HALFHOURS),
	m_dLoggedTime(0.0),
	m_dwTaskID(dwTaskID),
	m_sTaskTitle(szTaskTitle),
	m_bShowAddTimeToTimeSpent(FALSE),
	m_bTracked(dHours != 0.0)
{
	//{{AFX_DATA_INIT(CTDLAddLoggedTimeDlg)
	//}}AFX_DATA_INIT
	m_dtWhen = COleDateTime::GetCurrentTime();
	m_nUnits = (TCHAR)AfxGetApp()->GetProfileInt(_T("Preferences"), _T("AddLoggedTimeUnits"), THU_MINS);

	// convert log time from hours to current units
	m_dLoggedTime = CTimeHelper().GetTime(dHours, THU_HOURS, m_nUnits);
	m_eLoggedTime.EnableNegativeTimes(TRUE);
}

void CTDLAddLoggedTimeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTDLAddLoggedTimeDlg)
	DDX_Control(pDX, IDC_WHENDATE, m_dateWhen);
	DDX_Control(pDX, IDC_WHENTIME, m_cbTimeWhen);
	DDX_Text(pDX, IDC_LOGGEDTIME, m_dLoggedTime);
	DDX_Text(pDX, IDC_TASKID, m_dwTaskID);
	DDX_Text(pDX, IDC_TASKTITLE, m_sTaskTitle);
	DDX_Control(pDX, IDC_LOGGEDTIME, m_eLoggedTime);
	DDX_Check(pDX, IDC_ADDTIMETOTIMESPENT, m_bAddTimeToTimeSpent);
	DDX_Text(pDX, IDC_COMMENT, m_sComment);
	//}}AFX_DATA_MAP
	DDX_Control(pDX, IDC_TASKTITLE, m_stTaskTitle);

	if (pDX->m_bSaveAndValidate)
	{
		m_nUnits = m_eLoggedTime.GetUnits();

		m_dateWhen.GetTime(m_dtWhen);
		COleDateTime time = m_cbTimeWhen.GetOleTime();

		m_dtWhen = CDateHelper::GetDateOnly(m_dtWhen) + time;
	}
	else
	{
		m_eLoggedTime.SetUnits(m_nUnits);

		m_dateWhen.SetTime(m_dtWhen);
		m_cbTimeWhen.SetOleTime(CDateHelper::GetTimeOnly(m_dtWhen));
	}
}


BEGIN_MESSAGE_MAP(CTDLAddLoggedTimeDlg, CDialog)
	//{{AFX_MSG_MAP(CTDLAddLoggedTimeDlg)
	ON_BN_CLICKED(IDAPPLY, OnApply)
	ON_EN_CHANGE(IDC_LOGGEDTIME, OnChange)
	ON_EN_CHANGE(IDC_COMMENT, OnChange)
	ON_BN_CLICKED(IDC_ADDTIMETOTIMESPENT, OnChange)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_WHENDATE, OnNotifyChange)
	ON_CBN_SELCHANGE(IDC_WHENTIME, OnChange)
	ON_CBN_EDITCHANGE(IDC_WHENTIME, OnChange)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLAddLoggedTimeDlg message handlers

int CTDLAddLoggedTimeDlg::DoModal(BOOL bShowAddTimeToTimeSpent)
{
	m_bShowAddTimeToTimeSpent = bShowAddTimeToTimeSpent;

	if (bShowAddTimeToTimeSpent)
		m_bAddTimeToTimeSpent = AfxGetApp()->GetProfileInt(_T("Preferences"), _T("AddLoggedTimeToTimeSpent"), TRUE);

	return CDialog::DoModal();
}

double CTDLAddLoggedTimeDlg::GetLoggedHours() const
{ 
	return CTimeHelper().GetTime(m_dLoggedTime, m_nUnits, THU_HOURS);
}

COleDateTime CTDLAddLoggedTimeDlg::GetWhen() const
{
	return m_dtWhen;
}

void CTDLAddLoggedTimeDlg::OnOK()
{
	CDialog::OnOK();

	CPreferences prefs;

	if (m_bShowAddTimeToTimeSpent)
		prefs.WriteProfileInt(_T("Preferences"), _T("AddLoggedTimeToTimeSpent"), m_bAddTimeToTimeSpent);

	prefs.WriteProfileInt(_T("Preferences"), _T("AddLoggedTimeUnits"), m_nUnits);
	
	// Treat as 'cancel' if 'Apply' button is still disabled
	if (!GetDlgItem(IDAPPLY)->IsWindowEnabled())
		EndDialog(IDCANCEL);
}

BOOL CTDLAddLoggedTimeDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();

	CLocalizer::EnableTranslation(::GetDlgItem(*this, IDC_TASKTITLE), FALSE);
	
	if (!m_bShowAddTimeToTimeSpent)
	{
		GetDlgItem(IDC_ADDTIMETOTIMESPENT)->EnableWindow(FALSE);
		GetDlgItem(IDC_ADDTIMETOTIMESPENT)->ShowWindow(SW_HIDE);
	}
	
	m_stTaskTitle.SetFontStyle(TRUE);

	// set focus to time spent if no time specified
	if (m_dLoggedTime == 0.0)
	{
		m_eLoggedTime.SetFocus();
		m_eLoggedTime.SetSel(0, -1);

		return FALSE;
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CTDLAddLoggedTimeDlg::OnApply() 
{
	// make sure we are up to date
	UpdateData();

	// notify parent
	CWnd* pParent = (m_pParentWnd ? m_pParentWnd : GetParent());
	ASSERT(pParent);

	if (pParent->SendMessage(WM_TDL_APPLYADDLOGGEDTIME, m_dwTaskID, (LPARAM)this))
	{
		// disable apply button until another change is made
		GetDlgItem(IDAPPLY)->EnableWindow(FALSE);
	}
}

void CTDLAddLoggedTimeDlg::OnChange() 
{
	// reenable apply button
	GetDlgItem(IDAPPLY)->EnableWindow(TRUE);
}

void CTDLAddLoggedTimeDlg::OnNotifyChange(NMHDR* /*pNMHDR*/, LRESULT* pResult) 
{
	OnChange();
	
	*pResult = 0;
}
