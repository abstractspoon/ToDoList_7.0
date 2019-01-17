// ToDoCtrlReminders.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "ToDoCtrlReminders.h"
#include "filteredToDoCtrl.h"

#include "..\shared\preferences.h"
#include "..\shared\filemisc.h"
#include "..\shared\autoflag.h"

#include <math.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////

TDCREMINDER::TDCREMINDER() 
	: 
	dwTaskID(0), 
	pTDC(NULL), 
	dRelativeDaysLeadIn(0.0), 
	dDaysSnooze(0.0), 
	nRelativeFromWhen(TDCR_DUEDATE), 
	bEnabled(TRUE), 
	bRelative(FALSE),
	nLastSnoozeMins(0)
{
}

CString TDCREMINDER::GetTaskTitle() const
{
	ASSERT(pTDC);

	if (pTDC)
		return pTDC->GetTaskTitle(dwTaskID);

	// else
	return _T("");
}

CString TDCREMINDER::GetTaskComments() const
{
	ASSERT(pTDC);

	if (pTDC)
	{
		CString sComments = pTDC->GetTaskComments(dwTaskID);

		if (sComments.GetLength() > 100)
			sComments = sComments.Left(100) + _T("...");

		return sComments;
	}

	// else
	return _T("");
}

CString TDCREMINDER::FormatWhenString() const
{
	ASSERT(pTDC);

	CString sWhen;

	if (pTDC && bRelative)
	{
		CEnString sFormat;
		COleDateTime date;
		double dWhen = 0;

		if (nRelativeFromWhen == TDCR_DUEDATE)
		{
			date = pTDC->GetTaskDate(dwTaskID, TDCD_DUE);
			dWhen = date - COleDateTime::GetCurrentTime();

			if (dWhen < 1.0)
			{
				sFormat.LoadString(IDS_DUEWHENREMINDERNOW);
			}
			else if (fabs(dWhen) < 1.0)
			{
				dWhen *= 24 * 60; // convert to minutes
				sFormat.LoadString(IDS_DUEWHENREMINDERMINS);
			}
			else
			{
				sFormat.LoadString(IDS_DUEWHENREMINDERHOURS);
			}
		}
		else
		{
			date = pTDC->GetTaskDate(dwTaskID, TDCD_START);
			dWhen = date - COleDateTime::GetCurrentTime();

			if (dWhen < 1.0)
			{
				sFormat.LoadString(IDS_BEGINWHENREMINDERNOW);
			}
			else if (fabs(dWhen) < 1.0)
			{
				dWhen *= 24 * 60; // convert to minutes
				sFormat.LoadString(IDS_BEGINWHENREMINDERMINS);
			}
			else
			{
				sFormat.LoadString(IDS_BEGINWHENREMINDERHOURS);
			}
		}

		CString sDateTime = CDateHelper::FormatDate(date, DHFD_DOW | DHFD_NOSEC | DHFD_TIME);

		if (dWhen < 1.0)
			sWhen.Format(sFormat, sDateTime);
		else
			sWhen.Format(sFormat, dWhen, sDateTime);
	}

	return sWhen;
}

void TDCREMINDER::Save(IPreferences* pPrefs, LPCTSTR szKey) const
{
	pPrefs->WriteProfileInt(szKey, _T("TaskID"), dwTaskID);
	pPrefs->WriteProfileInt(szKey, _T("Relative"), bRelative);
	pPrefs->WriteProfileDouble(szKey, _T("Snooze"), dDaysSnooze);
	pPrefs->WriteProfileInt(szKey, _T("LastSnoozeMins"), nLastSnoozeMins);
	
	if (bRelative)
	{
		pPrefs->WriteProfileDouble(szKey, _T("LeadIn"), dRelativeDaysLeadIn * 24 * 60); // save as minutes
		pPrefs->WriteProfileInt(szKey, _T("FromWhen"), nRelativeFromWhen);
	}
	else
		pPrefs->WriteProfileDouble(szKey, _T("AbsoluteDate"), dtAbsolute);
	
	pPrefs->WriteProfileInt(szKey, _T("Enabled"), bEnabled);
	pPrefs->WriteProfileString(szKey, _T("SoundFile"), sSoundFile);
}

void TDCREMINDER::Load(const IPreferences* pPrefs, LPCTSTR szKey)
{
	dwTaskID = pPrefs->GetProfileInt(szKey, _T("TaskID"));
	bRelative = pPrefs->GetProfileInt(szKey, _T("Relative"));
	dDaysSnooze = pPrefs->GetProfileDouble(szKey, _T("Snooze"));
	nLastSnoozeMins = pPrefs->GetProfileInt(szKey, _T("LastSnoozeMins"));
	
	if (bRelative)
	{
		dRelativeDaysLeadIn = pPrefs->GetProfileDouble(szKey, _T("LeadIn")) / (24 * 60);
		nRelativeFromWhen = (TDC_REMINDER)pPrefs->GetProfileInt(szKey, _T("FromWhen"));
	}
	else
		dtAbsolute = pPrefs->GetProfileDouble(szKey, _T("AbsoluteDate"));
	
	bEnabled = pPrefs->GetProfileInt(szKey, _T("Enabled"));
	sSoundFile = pPrefs->GetProfileString(szKey, _T("SoundFile"));
	sStickiesID = pPrefs->GetProfileString(szKey, _T("StickiesID"));
}

/////////////////////////////////////////////////////////////////////////////
// CToDoCtrlReminders

CToDoCtrlReminders::CToDoCtrlReminders() : m_pWndNotify(NULL), m_bUseStickies(FALSE)
{
}

CToDoCtrlReminders::~CToDoCtrlReminders()
{
}


BEGIN_MESSAGE_MAP(CToDoCtrlReminders, CStickiesWnd)
	//{{AFX_MSG_MAP(CToDoCtrlReminders)
	ON_WM_TIMER()
	//}}AFX_MSG_MAP
	ON_WM_DESTROY()
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CToDoCtrlReminders message handlers

BOOL CToDoCtrlReminders::Initialize(CWnd* pNotify)
{
	ASSERT_VALID(pNotify);

	if (!pNotify || !pNotify->GetSafeHwnd())
		return FALSE;

	m_pWndNotify = pNotify;

	// create ourselves so that we can receive timer messages
	if (!Create(_T("STATIC"), _T(""), WS_CHILD, CRect (0, 0, 0, 0), pNotify, 0xffff))
	{
		m_pWndNotify = NULL;
		return FALSE;
	}

	// else
	return TRUE;
}

BOOL CToDoCtrlReminders::UseStickies(BOOL bEnable, LPCTSTR szStickiesPath)
{
	if (!bEnable || FileMisc::FileExists(szStickiesPath))
	{
		m_bUseStickies = bEnable;
		m_sStickiesPath = szStickiesPath;

		return TRUE;
	}

	// else
	return FALSE;
}

void CToDoCtrlReminders::AddToDoCtrl(const CFilteredToDoCtrl& tdc)
{
	LoadReminders(tdc);
}

void CToDoCtrlReminders::CloseToDoCtrl(const CFilteredToDoCtrl& tdc)
{
	SaveAndRemoveReminders(tdc);
}

void CToDoCtrlReminders::SetReminder(const TDCREMINDER& rem)
{
	ASSERT (m_pWndNotify);

	int nExist = FindReminder(rem);
	TDCREMINDER temp = rem; // to get around constness

	if (nExist != -1) // already exists
	{
		// replace
		m_aReminders[nExist] = temp;
	}
	else
	{
		// add
		m_aReminders.Add(temp);
	}

	StartTimer();
}

void CToDoCtrlReminders::StartTimer()
{
	if (!GetSafeHwnd())
		return;

	if (m_aReminders.GetSize())
	{
#ifdef _DEBUG
		SetTimer(1, 10000, NULL);
#else
		SetTimer(1, 60000, NULL); // every minute
#endif
	}
	else
		KillTimer(1);
}

void CToDoCtrlReminders::RemoveReminder(const TDCREMINDER& rem)
{
	RemoveReminder(rem.dwTaskID, rem.pTDC);
}

void CToDoCtrlReminders::RemoveReminder(DWORD dwTaskID, const CFilteredToDoCtrl* pTDC)
{
	ASSERT (m_pWndNotify);

	int nRem = FindReminder(dwTaskID, pTDC);

	if (nRem != -1) // already exists
	{
		TDCREMINDER rem = m_aReminders[nRem];
		m_aReminders.RemoveAt(nRem);

		// kill times if we have no reminders
		if (m_aReminders.GetSize() == 0)
			KillTimer(1);
	}
}

BOOL CToDoCtrlReminders::GetReminder(int nRem, TDCREMINDER& rem) const
{
	if (nRem < 0 || nRem >= m_aReminders.GetSize())
		return FALSE;

	rem = m_aReminders[nRem];
	return TRUE;
}

void CToDoCtrlReminders::RemoveDeletedTaskReminders(const CFilteredToDoCtrl* pTDC)
{
	int nRem = m_aReminders.GetSize();

	while (nRem--)
	{
		TDCREMINDER rem = m_aReminders[nRem];

		if (pTDC == NULL || pTDC == rem.pTDC)
		{
			if (!rem.pTDC->HasTask(rem.dwTaskID))
				m_aReminders.RemoveAt(nRem);
		}
	}
}

int CToDoCtrlReminders::FindReminder(const TDCREMINDER& rem, BOOL bIncludeDisabled) const
{
	return FindReminder(rem.dwTaskID, rem.pTDC, bIncludeDisabled);
}

int CToDoCtrlReminders::FindReminder(DWORD dwTaskID, const CFilteredToDoCtrl* pTDC, BOOL bIncludeDisabled) const
{
	int nRem = m_aReminders.GetSize();

	while (nRem--)
	{
		TDCREMINDER rem = m_aReminders[nRem];

		if (dwTaskID == rem.dwTaskID && pTDC == rem.pTDC)
		{
			if (bIncludeDisabled || rem.bEnabled)
				return nRem;
		}
	}

	return -1;
}

void CToDoCtrlReminders::SaveAndRemoveReminders(const CFilteredToDoCtrl& tdc)
{
	CPreferences prefs;

	// nRem is the total number of reminders for all tasklists
	// nRemCount is the number of reminders for 'tdc'
	int nRemCount = 0, nRem = m_aReminders.GetSize();
	CString sFileKey = tdc.GetPreferencesKey(_T("Reminders"));

	while (nRem--)
	{
		TDCREMINDER rem = m_aReminders[nRem];

		if (rem.pTDC == &tdc)
		{
			// verify that the task for this reminder still exists
			if (!tdc.HasTask(rem.dwTaskID))
				continue;

			CString sKey = sFileKey + Misc::MakeKey(_T("\\Reminder%d"), nRemCount);
			rem.Save(&prefs, sKey);

			nRemCount++;
			m_aReminders.RemoveAt(nRem);
		}
	}

	prefs.WriteProfileInt(sFileKey, _T("NumReminders"), nRemCount);

	// kill timer if no reminders
	if (GetSafeHwnd() && m_aReminders.GetSize() == 0)
		KillTimer(1);
}

BOOL CToDoCtrlReminders::ToDoCtrlHasReminders(const CFilteredToDoCtrl& tdc)
{
	return ToDoCtrlHasReminders(tdc.GetFilePath());
}

BOOL CToDoCtrlReminders::ToDoCtrlHasReminders(const CString& sFilePath)
{
	if (sFilePath.IsEmpty())
		return FALSE;

	CPreferences prefs;
	CString sFileKey = CFilteredToDoCtrl::GetPreferencesKey(sFilePath, _T("Reminders"));
	
	return (!sFileKey.IsEmpty() && prefs.GetProfileInt(sFileKey, _T("NumReminders")) > 0);
}

void CToDoCtrlReminders::LoadReminders(const CFilteredToDoCtrl& tdc)
{
	CPreferences prefs;
	CString sFileKey = tdc.GetPreferencesKey(_T("Reminders"));
	int nRemCount = prefs.GetProfileInt(sFileKey, _T("NumReminders"));

	for (int nRem = 0; nRem < nRemCount; nRem++)
	{
		CString sKey = Misc::MakeKey(_T("\\Reminder%d"), nRem, sFileKey);
		TDCREMINDER rem;

		rem.pTDC = &tdc;
		rem.Load(&prefs, sKey);

		m_aReminders.Add(rem);
	}

	// start timer if some reminders and not using Stickies
	if (!m_stickies.IsValid())
		StartTimer();
}

void CToDoCtrlReminders::OnDestroy()
{
	CStickiesWnd::OnDestroy();
}

void CToDoCtrlReminders::OnTimer(UINT nIDEvent) 
{
	// prevent re-entrancy
	AF_NOREENTRANT

	int nRem = m_aReminders.GetSize();

	if (nRem == 0)
	{
		KillTimer(1);
		return;
	}

	// iterate all the reminders looking for matches
	while (nRem--)
	{
		TDCREMINDER rem = m_aReminders[nRem];

		// check for disabled reminders
		if (!rem.bEnabled)
			continue;

		// check for completed tasks
		if (rem.pTDC->IsTaskDone(rem.dwTaskID, TDCCHECKALL))
		{
			m_aReminders.RemoveAt(nRem);
			continue;
		}

		// else
		COleDateTime dateNow = COleDateTime::GetCurrentTime();
		COleDateTime dateRem = GetReminderDate(rem);

		if (dateNow > dateRem)
		{
			if (SendParentReminder(rem) == 0L)
			{
				m_aReminders[nRem] = rem;
			}
			else // delete
			{
				m_aReminders.RemoveAt(nRem);
				rem.pTDC->RedrawReminders();
			}
		}
	}
	
	CStickiesWnd::OnTimer(nIDEvent);
}

LRESULT CToDoCtrlReminders::SendParentReminder(const TDCREMINDER& rem)
{
	if (m_bUseStickies && InitStickiesAPI(m_sStickiesPath))
	{
		CString sContent(rem.GetTaskTitle()), 
				sWhen(rem.FormatWhenString()), 
				sComments(rem.GetTaskComments()), 
				sDummy;

		if (!sComments.IsEmpty())
		{
			sContent += _T("\n\n");
			sContent += sComments;
		}

		if (!sWhen.IsEmpty())
		{
			sContent += _T("\n\n");
			sContent += sWhen;
		}
		
		if (CreateSticky(CEnString(IDS_STICKIES_TITLE), sDummy, sContent))
		{
			return 1L; // delete reminder as Stickies takes over
		}
	}
		
	// all else (fallback)
	ASSERT (m_pWndNotify);
	return m_pWndNotify->SendMessage(WM_TD_REMINDER, 0, (LPARAM)&rem);
}

COleDateTime CToDoCtrlReminders::GetReminderDate(const TDCREMINDER& rem)
{
	COleDateTime dateRem(rem.dtAbsolute);
	
	if (rem.bRelative)
	{
		ASSERT(rem.pTDC);
		ASSERT(rem.dwTaskID);
		
		if (rem.nRelativeFromWhen == TDCR_DUEDATE)
		{
			dateRem = rem.pTDC->GetTaskDate(rem.dwTaskID, TDCD_DUE);
		}
		else // start date
		{
			dateRem = rem.pTDC->GetTaskDate(rem.dwTaskID, TDCD_START);
		}
		
		dateRem -= rem.dRelativeDaysLeadIn;
	}

	dateRem += rem.dDaysSnooze;
	
	return dateRem;
}
