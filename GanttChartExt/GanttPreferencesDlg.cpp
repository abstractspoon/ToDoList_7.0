// GanttPreferencesDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "GanttPreferencesDlg.h"

#include "..\shared\dialoghelper.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

const COLORREF DEF_TODAYCOLOR		= RGB(255, 0, 0);
const COLORREF DEF_WEEKENDCOLOR		= RGB(224, 224, 224);
const COLORREF DEF_PARENTCOLOR		= RGB(0, 0, 0);

/////////////////////////////////////////////////////////////////////////////
// CGanttPreferencesPage dialog

CGanttPreferencesPage::CGanttPreferencesPage(CWnd* /*pParent*/ /*=NULL*/)
	: CPropertyPage(IDD_PREFERENCES_PAGE), m_brush(NULL)
{
	//{{AFX_DATA_INIT(CGanttPreferencesPage)
	m_bDisplayAllocTo = TRUE;
	m_bAutoScrollSelection = TRUE;
	m_bSpecifyWeekendColor = TRUE;
	m_bSpecifyTodayColor = TRUE;
	m_bAutoCalcParentDates = TRUE;
	m_bCalculateMissingStartDates = TRUE;
	m_bCalculateMissingDueDates = TRUE;
	m_nParentColoring = 0;
	m_bUseTagForMilestone = FALSE;
	m_sMilestoneTag = _T("");
	m_bDisplayProgressInBar = FALSE;
	//}}AFX_DATA_INIT
	m_crParent = DEF_PARENTCOLOR;
	m_crToday = DEF_TODAYCOLOR;
	m_crWeekend = DEF_WEEKENDCOLOR;
}

void CGanttPreferencesPage::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);

	//{{AFX_DATA_MAP(CGanttPreferencesPage)
	DDX_Check(pDX, IDC_DISPLAYALLOCTO, m_bDisplayAllocTo);
	DDX_Check(pDX, IDC_AUTOSCROLLSELECTION, m_bAutoScrollSelection);
	DDX_Check(pDX, IDC_WEEKENDCOLOR, m_bSpecifyWeekendColor);
	DDX_Check(pDX, IDC_TODAYCOLOR, m_bSpecifyTodayColor);
	DDX_Check(pDX, IDC_CALCULATEPARENTDATES, m_bAutoCalcParentDates);
	DDX_Check(pDX, IDC_CALCULATEMISSINGSTARTDATE, m_bCalculateMissingStartDates);
	DDX_Check(pDX, IDC_CALCULATEMISSINGDUEDATE, m_bCalculateMissingDueDates);
	DDX_Radio(pDX, IDC_DEFAULTPARENTCOLORS, m_nParentColoring);
	DDX_Check(pDX, IDC_USETAGFORMILESTONE, m_bUseTagForMilestone);
	DDX_Text(pDX, IDC_MILESTONETAG, m_sMilestoneTag);
	DDX_Check(pDX, IDC_DISPLAYPROGRESS, m_bDisplayProgressInBar);
	//}}AFX_DATA_MAP
	DDX_Control(pDX, IDC_SETTODAYCOLOR, m_btTodayColor);
	DDX_Control(pDX, IDC_SETWEEKENDCOLOR, m_btWeekendColor);
	DDX_Control(pDX, IDC_SETPARENTCOLOR, m_btParentColor);
}


BEGIN_MESSAGE_MAP(CGanttPreferencesPage, CPropertyPage)
	//{{AFX_MSG_MAP(CGanttPreferencesPage)
	ON_BN_CLICKED(IDC_SETWEEKENDCOLOR, OnSetWeekendcolor)
	ON_BN_CLICKED(IDC_SETTODAYCOLOR, OnSetTodaycolor)
	ON_BN_CLICKED(IDC_WEEKENDCOLOR, OnWeekendcolor)
	ON_BN_CLICKED(IDC_TODAYCOLOR, OnTodaycolor)
	ON_BN_CLICKED(IDC_DEFAULTPARENTCOLORS, OnChangeParentColoring)
	ON_BN_CLICKED(IDC_USETAGFORMILESTONE, OnUseTagForMilestone)
	ON_BN_CLICKED(IDC_NOPARENTCOLOR, OnChangeParentColoring)
	ON_BN_CLICKED(IDC_SETPARENTCOLOR, OnSetParentColor)
	ON_BN_CLICKED(IDC_SPECIFYPARENTCOLOR, OnChangeParentColoring)
	//}}AFX_MSG_MAP
	ON_WM_CTLCOLOR()
	ON_WM_ERASEBKGND()
	ON_WM_SHOWWINDOW()
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CGanttPreferencesPage message handlers

COLORREF CGanttPreferencesPage::GetTodayColor() const 
{ 
	return m_bSpecifyTodayColor ? m_crToday : CLR_NONE; 
}

COLORREF CGanttPreferencesPage::GetWeekendColor() const 
{ 
	return m_bSpecifyWeekendColor ? m_crWeekend : CLR_NONE; 
}

void CGanttPreferencesPage::OnSetWeekendcolor() 
{
	m_crWeekend = m_btWeekendColor.GetColor();
	Invalidate();
}

void CGanttPreferencesPage::OnSetTodaycolor() 
{
	m_crToday = m_btTodayColor.GetColor();
	Invalidate();
}

void CGanttPreferencesPage::OnWeekendcolor() 
{
	UpdateData();
	m_btWeekendColor.EnableWindow(m_bSpecifyWeekendColor);
}

void CGanttPreferencesPage::OnTodaycolor() 
{
	UpdateData();
	m_btTodayColor.EnableWindow(m_bSpecifyTodayColor);
}

BOOL CGanttPreferencesPage::OnInitDialog() 
{
	CPropertyPage::OnInitDialog();
	
	m_mgrGroupLines.AddGroupLine(IDC_COLORSGROUP, *this);
	m_mgrGroupLines.AddGroupLine(IDC_DATESGROUP, *this);

	m_btWeekendColor.EnableWindow(m_bSpecifyWeekendColor);
	m_btTodayColor.EnableWindow(m_bSpecifyTodayColor);
	m_btParentColor.EnableWindow(m_nParentColoring == 2);
	m_btWeekendColor.SetColor(m_crWeekend);
	m_btTodayColor.SetColor(m_crToday);
	m_btParentColor.SetColor(m_crParent);

	GetDlgItem(IDC_MILESTONETAG)->EnableWindow(m_bUseTagForMilestone);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

BOOL CGanttPreferencesPage::OnEraseBkgnd(CDC* pDC)
{
	if (m_brush == NULL)
		m_brush = ::CreateSolidBrush(GetSysColor(COLOR_WINDOW));

	CRect rClient;
	GetClientRect(rClient);
	pDC->FillSolidRect(rClient, GetSysColor(COLOR_WINDOW));

	return TRUE;
}

HBRUSH CGanttPreferencesPage::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CPropertyPage::OnCtlColor(pDC, pWnd, nCtlColor);

	if (nCtlColor == CTLCOLOR_STATIC && m_brush != NULL)
	{
		pDC->SetBkMode(TRANSPARENT);
		hbr = m_brush;
	}
	
	return hbr;
}

void CGanttPreferencesPage::OnShowWindow(BOOL bShow, UINT nStatus)
{
	CPropertyPage::OnShowWindow(bShow, nStatus);

	if (bShow)
		CDialogHelper::ResizeButtonStaticTextFieldsToFit(this);
}

void CGanttPreferencesPage::OnChangeParentColoring() 
{
	UpdateData();
	m_btParentColor.EnableWindow(m_nParentColoring == 2);
}

void CGanttPreferencesPage::OnSetParentColor() 
{
	m_crParent = m_btParentColor.GetColor();
}

int CGanttPreferencesPage::GetParentColoring(COLORREF& crParent) const
{
	crParent = m_crParent;
	return m_nParentColoring;
}

void CGanttPreferencesPage::SavePreferences(IPreferences* pPrefs, LPCTSTR szKey) const
{
	pPrefs->WriteProfileString(szKey, _T("MilestoneTag"), m_sMilestoneTag);

	pPrefs->WriteProfileInt(szKey, _T("UseTagForMilestone"), m_bUseTagForMilestone);
	pPrefs->WriteProfileInt(szKey, _T("DisplayAllocTo"), m_bDisplayAllocTo);
	pPrefs->WriteProfileInt(szKey, _T("AutoScrollSelection"), m_bAutoScrollSelection);
	pPrefs->WriteProfileInt(szKey, _T("AutoCalcParentDates"), m_bAutoCalcParentDates);
	pPrefs->WriteProfileInt(szKey, _T("CalculateMissingStartDates"), m_bCalculateMissingStartDates);
	pPrefs->WriteProfileInt(szKey, _T("CalculateMissingDueDates"), m_bCalculateMissingDueDates);
	pPrefs->WriteProfileInt(szKey, _T("TodayColor"), (int)m_crToday);
	pPrefs->WriteProfileInt(szKey, _T("WeekendColor"), (int)m_crWeekend);
	pPrefs->WriteProfileInt(szKey, _T("ParentColoring"), m_nParentColoring);
	pPrefs->WriteProfileInt(szKey, _T("ParentColor"), (int)m_crParent);
	pPrefs->WriteProfileInt(szKey, _T("DisplayProgressInBar"), m_bDisplayProgressInBar);
}

void CGanttPreferencesPage::LoadPreferences(const IPreferences* pPrefs, LPCTSTR szKey) 
{
	m_sMilestoneTag = pPrefs->GetProfileString(szKey, _T("MilestoneTag"), _T("Milestone"));
	m_bUseTagForMilestone = pPrefs->GetProfileInt(szKey, _T("UseTagForMilestone"), TRUE);
	m_bDisplayAllocTo = pPrefs->GetProfileInt(szKey, _T("DisplayAllocTo"), TRUE);
	m_bAutoScrollSelection = pPrefs->GetProfileInt(szKey, _T("AutoScrollSelection"), TRUE);
	m_bAutoCalcParentDates = pPrefs->GetProfileInt(szKey, _T("AutoCalcParentDates"), TRUE);
	m_bCalculateMissingStartDates = pPrefs->GetProfileInt(szKey, _T("CalculateMissingStartDates"), TRUE);
	m_bCalculateMissingDueDates = pPrefs->GetProfileInt(szKey, _T("CalculateMissingDueDates"), TRUE);
	m_bSpecifyTodayColor = pPrefs->GetProfileInt(szKey, _T("SpecifyTodayColor"), TRUE);
	m_crToday = (COLORREF)pPrefs->GetProfileInt(szKey, _T("TodayColor"), DEF_TODAYCOLOR);
	m_bSpecifyWeekendColor = pPrefs->GetProfileInt(szKey, _T("SpecifyWeekendColor"), TRUE);
	m_crWeekend = (COLORREF)pPrefs->GetProfileInt(szKey, _T("WeekendColor"), DEF_WEEKENDCOLOR);
	m_crParent = (COLORREF)pPrefs->GetProfileInt(szKey, _T("ParentColor"), DEF_PARENTCOLOR);
	m_nParentColoring = pPrefs->GetProfileInt(szKey, _T("ParentColoring"), 0);
	m_bDisplayProgressInBar = pPrefs->GetProfileInt(szKey, _T("DisplayProgressInBar"), FALSE);
}

void CGanttPreferencesPage::OnUseTagForMilestone() 
{
	UpdateData();
	GetDlgItem(IDC_MILESTONETAG)->EnableWindow(m_bUseTagForMilestone);
}

CString CGanttPreferencesPage::GetMilestoneTag() const
{
	return m_bUseTagForMilestone ? m_sMilestoneTag : _T("");
}

/////////////////////////////////////////////////////////////////////////////
// CGanttPreferencesDlg dialog

CGanttPreferencesDlg::CGanttPreferencesDlg(CWnd* pParent /*=NULL*/)
	: CDialog(IDD_PREFERENCES_DIALOG, pParent)
{
	//{{AFX_DATA_INIT(CGanttPreferencesDlg)
	//}}AFX_DATA_INIT

	m_ppHost.AddPage(&m_page);
}

BEGIN_MESSAGE_MAP(CGanttPreferencesDlg, CDialog)
	//{{AFX_MSG_MAP(CGanttPreferencesPage)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BOOL CGanttPreferencesDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();

	m_ppHost.Create(IDC_PPHOST, this);

	// shrink the ppHost by 1 pixel to allow the border to show
	CRect rHost;
	m_ppHost.GetWindowRect(rHost);
	ScreenToClient(rHost);

	rHost.DeflateRect(1, 1);
	m_ppHost.MoveWindow(rHost, FALSE);
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CGanttPreferencesDlg::OnOK()
{
	CDialog::OnOK();

	m_ppHost.OnOK();
}
