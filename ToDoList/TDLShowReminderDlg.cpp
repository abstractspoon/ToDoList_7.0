// TDLShowReminderDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLShowReminderDlg.h"
#include "filteredtodoctrl.h"
#include "todoctrlreminders.h"

#include "..\Shared\enstring.h"
#include "..\Shared\graphicsmisc.h"
#include "..\Shared\preferences.h"
#include "..\Shared\localizer.h"

#pragma warning(push)
#pragma warning(disable: 4201)
#include <Mmsystem.h>
#pragma warning(pop)

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////

// for PlaySound
#pragma comment(lib, "winmm.lib")

/////////////////////////////////////////////////////////////////////////////

const UINT ONE_HOUR = 60;
const UINT ONE_DAY = ONE_HOUR * 24;
const UINT ONE_WEEK = ONE_DAY * 7;
const UINT ONE_MONTH = ONE_DAY * 30;
const UINT ONE_YEAR = ONE_DAY * 365;

const UINT SNOOZE[] = { 5, 10, 15, 20, 30, 45, 
						ONE_HOUR, 2 * ONE_HOUR, 3 * ONE_HOUR, 4 * ONE_HOUR, 5 * ONE_HOUR, 6 * ONE_HOUR, 12 * ONE_HOUR, 
						ONE_DAY, 2 * ONE_DAY, 3 * ONE_DAY, 4 * ONE_DAY, 5 * ONE_DAY, 6 * ONE_DAY, 
						ONE_WEEK, 2 * ONE_WEEK, 3 * ONE_WEEK, 4 * ONE_WEEK, 
						ONE_MONTH, 2 * ONE_MONTH, 3 * ONE_MONTH, 4 * ONE_MONTH, 5 * ONE_MONTH, 6 * ONE_MONTH, 9 * ONE_MONTH, 
						ONE_YEAR };

const UINT NUM_SNOOZE = (sizeof(SNOOZE) / sizeof(UINT));

/////////////////////////////////////////////////////////////////////////////
// CTDLShowReminderDlg dialog

CTDLShowReminderDlg::CTDLShowReminderDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTDLShowReminderDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTDLShowReminderDlg)
	//}}AFX_DATA_INIT
	m_nSnoozeIndex = -1;
}


void CTDLShowReminderDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTDLShowReminderDlg)
	DDX_Text(pDX, IDC_WHENTEXT, m_sWhen);
	DDX_Text(pDX, IDC_TASKTITLE, m_sTaskTitle);
	DDX_CBIndex(pDX, IDC_SNOOZE, m_nSnoozeIndex);
	//}}AFX_DATA_MAP
	DDX_Control(pDX, IDC_TASKTITLE, m_stTaskTitle);
}


BEGIN_MESSAGE_MAP(CTDLShowReminderDlg, CDialog)
	//{{AFX_MSG_MAP(CTDLShowReminderDlg)
	ON_BN_CLICKED(IDSNOOZE, OnSnooze)
	ON_BN_CLICKED(IDGOTOTASK, OnGototask)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLShowReminderDlg message handlers

int CTDLShowReminderDlg::DoModal(const TDCREMINDER& rem)
{
	m_sWhen = rem.FormatWhenString();
	m_sSoundFile = rem.sSoundFile;

	m_sTaskTitle = rem.GetTaskTitle();
	m_sTaskTitle.Replace(_T("&"), _T("&&"));

	// init snooze value
	UINT nSnooze = CPreferences().GetProfileInt(_T("Reminders"), _T("Snooze"), 5);

	// this reminder's previous snooze overrides
	if (rem.nLastSnoozeMins > 0)
		nSnooze = rem.nLastSnoozeMins;

	SetSnoozeMinutes(nSnooze);

	return CDialog::DoModal();
}

BOOL CTDLShowReminderDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	
	CLocalizer::EnableTranslation(::GetDlgItem(*this, IDC_TASKTITLE), FALSE);

	// do we need to play a sound?
	if (!m_sSoundFile.IsEmpty())
		PlaySound(m_sSoundFile, NULL, SND_FILENAME);
	
	m_stTaskTitle.SetFontStyle(TRUE);
		
	return TRUE;
}

void CTDLShowReminderDlg::SetSnoozeMinutes(UINT nMins)
{
	int nSnooze = NUM_SNOOZE;
	
	while (nSnooze--)
	{
		if (SNOOZE[nSnooze] == nMins)
			break;
	}

	if (nSnooze == -1)
		nSnooze = 0;

	if (m_nSnoozeIndex != nSnooze)
	{
		m_nSnoozeIndex = nSnooze;

		if (GetSafeHwnd())
			UpdateData(FALSE);
	}
}

void CTDLShowReminderDlg::OnSnooze() 
{
	UpdateData();
	EndDialog(IDSNOOZE);	
	
	// save snooze value for next time
	CPreferences().WriteProfileInt(_T("Reminders"), _T("Snooze"), GetSnoozeMinutes());
}

UINT CTDLShowReminderDlg::GetSnoozeMinutes() const
{
	ASSERT (m_nSnoozeIndex < NUM_SNOOZE);

	if (m_nSnoozeIndex >= NUM_SNOOZE)
		return ONE_HOUR;
	else
		return SNOOZE[m_nSnoozeIndex];
}

double CTDLShowReminderDlg::GetSnoozeDays() const
{
	return (GetSnoozeMinutes() / (double)ONE_DAY);
}

void CTDLShowReminderDlg::OnGototask() 
{
	EndDialog(IDGOTOTASK);	
}
