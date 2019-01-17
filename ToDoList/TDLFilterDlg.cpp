// FilterDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLFilterDlg.h"
#include "tdcstatic.h"

#include "..\shared\enstring.h"
#include "..\shared\preferences.h"
#include "..\shared\localizer.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTDLFilterDlg dialog


CTDLFilterDlg::CTDLFilterDlg(BOOL bMultiSelFilters, CWnd* pParent /*=NULL*/)
	: CDialog(CTDLFilterDlg::IDD, pParent), 
	  m_cbCategoryFilter(bMultiSelFilters, IDS_TDC_NONE, IDS_TDC_ANY),
	  m_cbAllocToFilter(bMultiSelFilters, IDS_TDC_NOBODY, IDS_TDC_ANYONE),
	  m_cbAllocByFilter(bMultiSelFilters, IDS_TDC_NOBODY, IDS_TDC_ANYONE),
	  m_cbStatusFilter(bMultiSelFilters, IDS_TDC_NONE, IDS_TDC_ANY),
	  m_cbVersionFilter(bMultiSelFilters, IDS_TDC_NONE, IDS_TDC_ANY),
	  m_cbTagFilter(bMultiSelFilters, IDS_TDC_NONE, IDS_TDC_ANY),
	  m_nView(FTCV_UNSET)
{
	//{{AFX_DATA_INIT(CTDLFilterDlg)
	//}}AFX_DATA_INIT
}


void CTDLFilterDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTDLFilterDlg)
	DDX_Control(pDX, IDC_FILTEROPTIONS, m_cbOptions);
	DDX_Control(pDX, IDC_CATEGORYFILTERCOMBO, m_cbCategoryFilter);
	DDX_Control(pDX, IDC_STATUSFILTERCOMBO, m_cbStatusFilter);
	DDX_Control(pDX, IDC_PRIORITYFILTERCOMBO, m_cbPriorityFilter);
	DDX_Control(pDX, IDC_RISKFILTERCOMBO, m_cbRiskFilter);
	DDX_Control(pDX, IDC_ALLOCTOFILTERCOMBO, m_cbAllocToFilter);
	DDX_Control(pDX, IDC_ALLOCBYFILTERCOMBO, m_cbAllocByFilter);
	DDX_Control(pDX, IDC_VERSIONFILTERCOMBO, m_cbVersionFilter);
	DDX_Control(pDX, IDC_TAGFILTERCOMBO, m_cbTagFilter);
	DDX_Control(pDX, IDC_FILTERCOMBO, m_cbTaskFilter);
	//}}AFX_DATA_MAP
	DDX_Control(pDX, IDC_STARTFILTERCOMBO, m_cbStartFilter);
	DDX_Control(pDX, IDC_DUEFILTERCOMBO, m_cbDueFilter);
	DDX_Text(pDX, IDC_TITLEFILTERTEXT, m_filter.sTitle);
	DDX_DateTimeCtrl(pDX, IDC_USERSTARTDATE, m_filter.dtUserStart);
	DDX_DateTimeCtrl(pDX, IDC_USERDUEDATE, m_filter.dtUserDue);

	// special handling
	if (pDX->m_bSaveAndValidate)
	{
		// filter
		m_filter.nShow = m_cbTaskFilter.GetSelectedFilter(m_sCustomFilter);
		m_filter.nStartBy = m_cbStartFilter.GetSelectedFilter();
		m_filter.nDueBy = m_cbDueFilter.GetSelectedFilter();

		// priority
		int nIndex;

		DDX_CBIndex(pDX, IDC_PRIORITYFILTERCOMBO, nIndex);

		if (nIndex == 0) // any
			m_filter.nPriority = FM_ANYPRIORITY;

		else if (nIndex == 1) // none
			m_filter.nPriority = FM_NOPRIORITY;
		else
			m_filter.nPriority = nIndex - 2;

		// risk
		DDX_CBIndex(pDX, IDC_RISKFILTERCOMBO, nIndex);

		if (nIndex == 0) // any
			m_filter.nRisk = FM_ANYRISK;

		else if (nIndex == 1) // none
			m_filter.nRisk = FM_NORISK;
		else
			m_filter.nRisk = nIndex - 2;

		// auto droplists
		m_cbCategoryFilter.GetChecked(m_filter.aCategories);
		m_cbAllocToFilter.GetChecked(m_filter.aAllocTo);
		m_cbStatusFilter.GetChecked(m_filter.aStatus);
		m_cbAllocByFilter.GetChecked(m_filter.aAllocBy);
		m_cbVersionFilter.GetChecked(m_filter.aVersions);
		m_cbTagFilter.GetChecked(m_filter.aTags);

		// build the filter options from the listbox
		m_filter.dwFlags = m_cbOptions.GetSelectedOptions();
	}
	else
	{
		// filter
		if (m_filter.nShow == FS_CUSTOM)
			m_cbTaskFilter.SelectFilter(m_sCustomFilter);
		else
			m_cbTaskFilter.SelectFilter(m_filter.nShow);

		m_cbStartFilter.SelectFilter(m_filter.nStartBy);
		m_cbDueFilter.SelectFilter(m_filter.nDueBy);

		// priority
		int nIndex;
		
		if (m_filter.nPriority == FM_ANYPRIORITY)
			nIndex = 0;

		else if (m_filter.nPriority == FM_NOPRIORITY)
			nIndex = 1;
		else
			nIndex = m_filter.nPriority + 2;

		DDX_CBIndex(pDX, IDC_PRIORITYFILTERCOMBO, nIndex);

		// risk
		if (m_filter.nRisk == FM_ANYRISK)
			nIndex = 0;

		else if (m_filter.nRisk == FM_NORISK)
			nIndex = 1;
		else
			nIndex = m_filter.nRisk + 2;

		DDX_CBIndex(pDX, IDC_RISKFILTERCOMBO, nIndex);

		// auto droplists
		m_cbCategoryFilter.SetChecked(m_filter.aCategories);
		m_cbAllocToFilter.SetChecked(m_filter.aAllocTo);
		m_cbStatusFilter.SetChecked(m_filter.aStatus);
		m_cbAllocByFilter.SetChecked(m_filter.aAllocBy);
		m_cbVersionFilter.SetChecked(m_filter.aVersions);
		m_cbTagFilter.SetChecked(m_filter.aTags);

		// options states set in OnInitDialog
		m_cbOptions.Initialize(m_filter, m_nView, m_bWantHideParents);
	}
}


BEGIN_MESSAGE_MAP(CTDLFilterDlg, CDialog)
	//{{AFX_MSG_MAP(CTDLFilterDlg)
	ON_BN_CLICKED(IDC_CLEARFILTER, OnClearfilter)
	ON_CBN_SELCHANGE(IDC_FILTERCOMBO, OnSelchangeFiltercombo)
	//}}AFX_MSG_MAP
	ON_CBN_SELCHANGE(IDC_DUEFILTERCOMBO, OnSelchangeDatecombo)
	ON_CBN_SELCHANGE(IDC_STARTFILTERCOMBO, OnSelchangeDatecombo)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnToolTipNotify)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLFilterDlg message handlers

int CTDLFilterDlg::DoModal(const CStringArray& aCustomFilters,
						   const CFilteredToDoCtrl& tdc, 
						   const CDWordArray& aPriorityColors)
{
	// main filter
	tdc.GetFilter(m_filter);

	// get custom filter
	m_aCustomFilters.Copy(aCustomFilters);
	
	if (tdc.HasCustomFilter())
	{
		m_sCustomFilter = tdc.GetCustomFilterName();
		m_filter.nShow = FS_CUSTOM;
	}
	else
		m_sCustomFilter.Empty();

	// auto-droplists
	tdc.GetAutoListData(m_tldListData);
	
	m_aPriorityColors.Copy(aPriorityColors);
	m_nView = tdc.GetView();

	m_bWantHideParents = tdc.HasStyle(TDCS_ALWAYSHIDELISTPARENTS);

	return CDialog::DoModal();
}

void CTDLFilterDlg::GetFilter(FTDCFILTER& filter, CString& sCustom)
{
	filter = m_filter;
	sCustom = m_sCustomFilter;
}

BOOL CTDLFilterDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// disable translation of user-data combos
	CLocalizer::EnableTranslation(m_cbAllocByFilter, FALSE);
	CLocalizer::EnableTranslation(m_cbAllocToFilter, FALSE);
	CLocalizer::EnableTranslation(m_cbCategoryFilter, FALSE);
	CLocalizer::EnableTranslation(m_cbStatusFilter, FALSE);
	CLocalizer::EnableTranslation(m_cbVersionFilter, FALSE);
	CLocalizer::EnableTranslation(m_cbTagFilter, FALSE);

	// auto droplist filters
	m_cbAllocToFilter.SetStrings(m_tldListData.aAllocTo);
	m_cbAllocToFilter.AddEmptyString(); // add blank item to top

	// alloc by filter
	m_cbAllocByFilter.SetStrings(m_tldListData.aAllocBy);
	m_cbAllocByFilter.AddEmptyString(); // add blank item to top

	// category filter
	m_cbCategoryFilter.SetStrings(m_tldListData.aCategory);
	m_cbCategoryFilter.AddEmptyString(); // add blank item to top

	// status filter
	m_cbStatusFilter.SetStrings(m_tldListData.aStatus);
	m_cbStatusFilter.AddEmptyString(); // add blank item to top

	// version filter
	m_cbVersionFilter.SetStrings(m_tldListData.aVersion);
	m_cbVersionFilter.AddEmptyString(); // add blank item to top

	// tag filter
	m_cbTagFilter.SetStrings(m_tldListData.aTags);
	m_cbTagFilter.AddEmptyString(); // add blank item to top

	// priority
	CEnString sAny(IDS_TDC_ANY);

	m_cbPriorityFilter.SetColors(m_aPriorityColors);
	m_cbPriorityFilter.InsertColor(0, CLR_NONE, sAny); // add a blank item

	// risk
	m_cbRiskFilter.InsertString(0, sAny); // add a blank item

	// title
	m_mgrPrompts.SetEditPrompt(IDC_TITLEFILTERTEXT, *this, sAny);

	// custom filters
	m_cbTaskFilter.AddCustomFilters(m_aCustomFilters, m_sCustomFilter);

	// update UI
	UpdateData(FALSE);
	EnableDisableControls();

	EnableToolTips();
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CTDLFilterDlg::EnableDisableControls()
{
	BOOL bSelFilter = (m_filter.nShow == FS_SELECTED);
	BOOL bCustomFilter = (m_filter.nShow == FS_CUSTOM);
	BOOL bEnable = !(bSelFilter || bCustomFilter);

	GetDlgItem(IDC_TITLEFILTERTEXT)->EnableWindow(bEnable);

	m_cbStartFilter.EnableWindow(bEnable);
	m_cbDueFilter.EnableWindow(bEnable);
	m_cbCategoryFilter.EnableWindow(bEnable);
	m_cbStatusFilter.EnableWindow(bEnable);
	m_cbPriorityFilter.EnableWindow(bEnable);
	m_cbRiskFilter.EnableWindow(bEnable);
	m_cbAllocToFilter.EnableWindow(bEnable);
	m_cbAllocByFilter.EnableWindow(bEnable);
	m_cbVersionFilter.EnableWindow(bEnable);
	m_cbTagFilter.EnableWindow(bEnable);

	// special cases
	m_cbOptions.EnableWindow(!bCustomFilter);

	// user dates
	GetDlgItem(IDC_USERSTARTDATE)->EnableWindow(bEnable && (m_filter.nStartBy == FD_USER));
	GetDlgItem(IDC_USERDUEDATE)->EnableWindow(bEnable && (m_filter.nDueBy == FD_USER));
}


void CTDLFilterDlg::OnClearfilter() 
{
	m_filter.Reset();
	
	UpdateData(FALSE);
	
	// the combos don't get set properly if they
	// are empty but were not previously so we do it manually
// 	m_cbAllocToFilter.SetCurSel(0);
// 	m_cbAllocByFilter.SetCurSel(0);
// 	m_cbCategoryFilter.SetCurSel(0);
// 	m_cbVersionFilter.SetCurSel(0);
// 	m_cbStatusFilter.SetCurSel(0);
// 	m_cbTagFilter.SetCurSel(0);
}

void CTDLFilterDlg::OnSelchangeFiltercombo() 
{
	// enable/disable controls if necessary
	UpdateData();
	EnableDisableControls();

	m_cbOptions.Initialize(m_filter, m_nView);
}

void CTDLFilterDlg::OnSelchangeDatecombo()
{
	// enable/disable controls if necessary
	UpdateData();
	EnableDisableControls();
}

BOOL CTDLFilterDlg::OnToolTipNotify(UINT /*id*/, NMHDR* pNMHDR, LRESULT* /*pResult*/)
{
    // Get the tooltip structure.
    TOOLTIPTEXT *pTTT = (TOOLTIPTEXT *)pNMHDR;

    // Actually the idFrom holds Control's handle.
    UINT CtrlHandle = pNMHDR->idFrom;

    // Check once again that the idFrom holds handle itself.
    if (pTTT->uFlags & TTF_IDISHWND)
    {
		static CString sTooltip;
		sTooltip.Empty();

        // Get the control's ID.
        UINT nID = ::GetDlgCtrlID( HWND( CtrlHandle ));

        switch( nID )
        {
        case IDC_CATEGORYFILTERCOMBO:
			sTooltip = m_cbCategoryFilter.GetTooltip();
            break;

        case IDC_ALLOCTOFILTERCOMBO:
			sTooltip = m_cbAllocToFilter.GetTooltip();
            break;

        case IDC_STATUSFILTERCOMBO:
			sTooltip = m_cbStatusFilter.GetTooltip();
            break;

        case IDC_ALLOCBYFILTERCOMBO:
			sTooltip = m_cbAllocByFilter.GetTooltip();
            break;

        case IDC_VERSIONFILTERCOMBO:
			sTooltip = m_cbVersionFilter.GetTooltip();
            break;

        case IDC_TAGFILTERCOMBO:
			sTooltip = m_cbTagFilter.GetTooltip();
            break;

        case IDC_OPTIONFILTERCOMBO:
			sTooltip = m_cbOptions.GetTooltip();
            break;
        }

		if (!sTooltip.IsEmpty())
		{
			// Set the tooltip text.
			::SendMessage(pNMHDR->hwndFrom, TTM_SETMAXTIPWIDTH, 0, 300);
			pTTT->lpszText = (LPTSTR)(LPCTSTR)sTooltip;
	        return TRUE;
		}
    }

    // Not handled.
    return FALSE;
}