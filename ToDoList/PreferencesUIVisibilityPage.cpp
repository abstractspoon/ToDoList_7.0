// PreferencesUIVisibilityPage.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "PreferencesUIVisibilityPage.h"

#include "..\Shared\DialogHelper.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPreferencesUIVisibilityPage property page

IMPLEMENT_DYNCREATE(CPreferencesUIVisibilityPage, CPreferencesPageBase)

CPreferencesUIVisibilityPage::CPreferencesUIVisibilityPage() 
	: 
	CPreferencesPageBase(IDD_PREFUIVISIBILITY_PAGE)
{
	//{{AFX_DATA_INIT(CPreferencesUIVisibilityPage)
	//}}AFX_DATA_INIT
	m_nAttribShow = TDLSA_ASCOLUMN;
}

CPreferencesUIVisibilityPage::~CPreferencesUIVisibilityPage()
{
}

void CPreferencesUIVisibilityPage::DoDataExchange(CDataExchange* pDX)
{
	CPreferencesPageBase::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPreferencesUIVisibilityPage)
	DDX_Control(pDX, IDC_VISIBILITY, m_lcVisibility);
	//}}AFX_DATA_MAP
	DDX_Radio(pDX, IDC_SHOWALLATTRIB, (int&)m_nAttribShow);
}


BEGIN_MESSAGE_MAP(CPreferencesUIVisibilityPage, CPreferencesPageBase)
	//{{AFX_MSG_MAP(CPreferencesUIVisibilityPage)
	ON_WM_SIZE()
	ON_BN_CLICKED(IDC_SHOWALLATTRIB, OnChangeAttribShow)
	ON_BN_CLICKED(IDC_SHOWANYATTRIB, OnChangeAttribShow)
	ON_BN_CLICKED(IDC_SHOWATTRIBASCOLUMN, OnChangeAttribShow)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPreferencesUIVisibilityPage message handlers

BOOL CPreferencesUIVisibilityPage::OnInitDialog()
{
	CPreferencesPageBase::OnInitDialog();

	m_lcVisibility.SetCurSel(0, 0);

	return TRUE;
}

void CPreferencesUIVisibilityPage::GetColumnAttributeVisibility(TDCCOLEDITFILTERVISIBILITY& vis) const
{
	m_lcVisibility.GetVisibility(vis);
}

void CPreferencesUIVisibilityPage::SetColumnAttributeVisibility(const TDCCOLEDITFILTERVISIBILITY& vis)
{
	m_lcVisibility.SetVisibility(vis);
}

void CPreferencesUIVisibilityPage::LoadPreferences(const CPreferences& prefs)
{
	TDCCOLEDITFILTERVISIBILITY vis;

	// columns
	CTDCColumnIDArray aColumns;
	int nCol = prefs.GetProfileInt(_T("Preferences\\ColumnVisibility"), _T("Count"), -1);
	
	if (nCol >= 0)
	{
		while (nCol--)
		{
			CString sKey = Misc::MakeKey(_T("Col%d"), nCol);
			TDC_COLUMN col = (TDC_COLUMN)prefs.GetProfileInt(_T("Preferences\\ColumnVisibility"), sKey, TDCC_NONE);
			
			if (col != TDCC_NONE)
				aColumns.Add(col);
		}
	}
	else // first time => backwards compatibility
	{
		// same order as enum
		aColumns.Add(TDCC_PRIORITY); 
		aColumns.Add(TDCC_PERCENT); 
		aColumns.Add(TDCC_TIMEEST); 
		aColumns.Add(TDCC_TIMESPENT); 
		aColumns.Add(TDCC_TRACKTIME); 
		aColumns.Add(TDCC_DUEDATE); 
		aColumns.Add(TDCC_ALLOCTO); 
		aColumns.Add(TDCC_STATUS); 
		aColumns.Add(TDCC_CATEGORY); 
		aColumns.Add(TDCC_FILEREF); 
		aColumns.Add(TDCC_RECURRENCE); 
		aColumns.Add(TDCC_REMINDER); 
		aColumns.Add(TDCC_TAGS); 
	}
	
	vis.SetVisibleColumns(aColumns);

	// backwards compatibility
	if (prefs.GetProfileInt(_T("Preferences"), _T("ShowSubtaskCompletion"), FALSE))
		vis.SetColumnVisible(TDCC_SUBTASKDONE);

	// attrib
	CTDCAttributeArray aAttrib;
	int nEdit = prefs.GetProfileInt(_T("Preferences\\EditVisibility"), _T("Count"), -1);
	
	if (nEdit >= 0) 
	{
		// show attributes mode
		TDL_SHOWATTRIB nShow = (TDL_SHOWATTRIB)prefs.GetProfileInt(_T("Preferences"), _T("ShowAttributes"), TDLSA_ASCOLUMN);
		vis.SetShowEditsAndFilters(nShow);

		if (nShow == TDLSA_ANY)
		{
			// edit fields
			while (nEdit--)
			{
				CString sKey = Misc::MakeKey(_T("Edit%d"), nEdit);
				TDC_ATTRIBUTE att = (TDC_ATTRIBUTE)prefs.GetProfileInt(_T("Preferences\\EditVisibility"), sKey, TDCA_NONE);
				
				if (att != TDCA_NONE)
					aAttrib.Add(att);
			}
			
			vis.SetVisibleEditFields(aAttrib);

			// filter fields
			int nFilter = prefs.GetProfileInt(_T("Preferences\\FilterVisibility"), _T("Count"), -1);

			while (nFilter--)
			{
				CString sKey = Misc::MakeKey(_T("Filter%d"), nFilter);
				TDC_ATTRIBUTE att = (TDC_ATTRIBUTE)prefs.GetProfileInt(_T("Preferences\\FilterVisibility"), sKey, TDCA_NONE);
				
				if (att != TDCA_NONE)
					aAttrib.Add(att);
			}
			
			vis.SetVisibleFilterFields(aAttrib);
		}
	}
	else // first time => backwards compatibility
	{
		BOOL bShowCtrlsAsColumns = prefs.GetProfileInt(_T("Preferences"), _T("ShowCtrlsAsColumns"), FALSE);

		vis.SetShowEditsAndFilters(bShowCtrlsAsColumns ? TDLSA_ASCOLUMN : TDLSA_ALL);

		// if any time field is hidden we must enable 'any' attribute
		// and remove those fields
		BOOL bHideDueTimeField = prefs.GetProfileInt(_T("Preferences"), _T("HideDueTimeField"), FALSE);
		BOOL bHideStartTimeField = prefs.GetProfileInt(_T("Preferences"), _T("HideStartTimeField"), FALSE);
		BOOL bHideDoneTimeField = prefs.GetProfileInt(_T("Preferences"), _T("HideDoneTimeField"), FALSE);

		if (bHideDoneTimeField || bHideDueTimeField || bHideStartTimeField)
		{
			vis.SetShowEditsAndFilters(TDLSA_ANY);
			
			if (bHideStartTimeField)
				vis.SetEditFieldVisible(TDCA_STARTTIME, FALSE);
			
			if (bHideDueTimeField)
				vis.SetEditFieldVisible(TDCA_DUETIME, FALSE);
			
			if (bHideDoneTimeField)
				vis.SetEditFieldVisible(TDCA_DONETIME, FALSE);
			
			vis.SetVisibleEditFields(aAttrib);
		}
	}

	m_lcVisibility.SetVisibility(vis);

	m_nAttribShow = vis.GetShowEditsAndFilters();
}

void CPreferencesUIVisibilityPage::SavePreferences(CPreferences& prefs)
{
	TDCCOLEDITFILTERVISIBILITY vis;
	m_lcVisibility.GetVisibility(vis);

	// columns
	int nCol = vis.GetVisibleColumns().GetSize();

	prefs.WriteProfileInt(_T("Preferences\\ColumnVisibility"), _T("Count"), nCol);
	
	while (nCol--)
	{
		CString sKey = Misc::MakeKey(_T("Col%d"), nCol);
		prefs.WriteProfileInt(_T("Preferences\\ColumnVisibility"), sKey, vis.GetVisibleColumns()[nCol]);
	}

	// show attributes mode
	prefs.WriteProfileInt(_T("Preferences"), _T("ShowAttributes"), vis.GetShowEditsAndFilters());

	if (vis.GetShowEditsAndFilters() == TDLSA_ANY)
	{
		// edit fields
		int nEdit = vis.GetVisibleEditFields().GetSize();
		
		prefs.WriteProfileInt(_T("Preferences\\EditVisibility"), _T("Count"), nEdit);
		
		while (nEdit--)
		{
			CString sKey = Misc::MakeKey(_T("Edit%d"), nEdit);
			prefs.WriteProfileInt(_T("Preferences\\EditVisibility"), sKey, vis.GetVisibleEditFields()[nEdit]);
		}

		// filters
		int nFilter = vis.GetVisibleFilterFields().GetSize();
		
		prefs.WriteProfileInt(_T("Preferences\\FilterVisibility"), _T("Count"), nFilter);
		
		while (nFilter--)
		{
			CString sKey = Misc::MakeKey(_T("Filter%d"), nFilter);
			prefs.WriteProfileInt(_T("Preferences\\FilterVisibility"), sKey, vis.GetVisibleFilterFields()[nFilter]);
		}
	}
}

void CPreferencesUIVisibilityPage::OnSize(UINT nType, int cx, int cy) 
{
	CPreferencesPageBase::OnSize(nType, cx, cy);
	
	if (m_lcVisibility.GetSafeHwnd())
	{
		// calculate border
		CPoint ptBorders = CDialogHelper::GetCtrlRect(this, IDC_LISTLABEL).TopLeft();

		// calc offsets
		CRect rVis = CDialogHelper::GetCtrlRect(this, IDC_ATTRIBGROUP);

		int nXOffset = (cx - rVis.right - ptBorders.x);
		int nYOffset = (cy - rVis.bottom - ptBorders.y);

		CDialogHelper::OffsetCtrl(this, IDC_ATTRIBGROUP, 0, nYOffset);
		CDialogHelper::OffsetCtrl(this, IDC_SHOWALLATTRIB, 0, nYOffset);
		CDialogHelper::OffsetCtrl(this, IDC_SHOWATTRIBASCOLUMN, 0, nYOffset);
		CDialogHelper::OffsetCtrl(this, IDC_SHOWANYATTRIB, 0, nYOffset);

		CDialogHelper::ResizeCtrl(this, IDC_ATTRIBGROUP, nXOffset, 0);
		CDialogHelper::ResizeChild(&m_lcVisibility, nXOffset, nYOffset);
	}
}

void CPreferencesUIVisibilityPage::OnChangeAttribShow() 
{
	UpdateData();

	m_lcVisibility.SetShowEditsAndFilters(m_nAttribShow);
}
