#if !defined(AFX_FILTERBAR_H__FAABD9A1_72C4_4731_B7A4_48251860672C__INCLUDED_)
#define AFX_FILTERBAR_H__FAABD9A1_72C4_4731_B7A4_48251860672C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FilterBar.h : header file
//

#include "..\shared\dialoghelper.h"
#include "..\shared\encheckcombobox.h"
#include "..\shared\enedit.h"
#include "..\shared\wndprompt.h"
#include "..\shared\uithemefile.h"

#include "tdcstruct.h"
#include "tdlprioritycombobox.h"
#include "tdlriskcombobox.h"
#include "tdlfilteroptioncombobox.h"
#include "tdlfiltercombobox.h"
#include "tdlfilterdatecombobox.h"

/////////////////////////////////////////////////////////////////////////////
// CFilterBar dialog

class CFilteredToDoCtrl;
class CDeferWndMove;

class CTDLFilterBar : public CDialog, public CDialogHelper
{
// Construction
public:
	CTDLFilterBar(CWnd* pParent = NULL);   // standard constructor

	BOOL Create(CWnd* pParentWnd, UINT nID = 0, BOOL bVisible = TRUE);
	void GetFilter(FTDCFILTER& filter) const { filter = m_filter; }
	void SetFilter(const FTDCFILTER& filter, FTC_VIEW nView);
	
	FILTER_SHOW GetActiveFilter() const;
	FILTER_SHOW GetActiveFilter(CString& sCustom) const;
	DWORD GetFilterOptions() const;

	void SetCustomFilter(BOOL bCustom = TRUE, LPCTSTR szFilter = NULL);
	void AddCustomFilters(const CStringArray& aFilters);
	int GetCustomFilters(CStringArray& aFilters) const;
	void RemoveCustomFilters();
	void ShowDefaultFilters(BOOL bShow);

	void RefreshFilterControls(const CFilteredToDoCtrl& tdc);
	void SetFilterLabelAlignment(BOOL bLeft);
	void SetPriorityColors(const CDWordArray& aColors);
	int CalcHeight(int nWidth);
	void SetVisibleFilters(const CTDCAttributeArray& aFilters);
	BOOL FilterMatches(const FTDCFILTER& filter) { return (filter == m_filter); }

	void EnableMultiSelection(BOOL bEnable);
	void SetUITheme(const CUIThemeFile& theme);

protected:
// Dialog Data
	//{{AFX_DATA(CFilterBar)
	//}}AFX_DATA
	CTDLFilterComboBox	m_cbTaskFilter;
	CTDLFilterDateComboBox	m_cbStartFilter;
	CTDLFilterDateComboBox	m_cbDueFilter;
	CEnEdit m_eTitleFilter;
	CEnCheckComboBox	m_cbAllocToFilter;
	CEnCheckComboBox	m_cbAllocByFilter;
	CEnCheckComboBox	m_cbCategoryFilter;
	CEnCheckComboBox	m_cbStatusFilter;
	CEnCheckComboBox	m_cbVersionFilter;
	CEnCheckComboBox	m_cbTagFilter;
	CTDLPriorityComboBox	m_cbPriorityFilter;
	CTDLRiskComboBox	m_cbRiskFilter;
	CTDLFilterOptionComboBox	m_cbOptions;
	CDateTimeCtrl m_dtcUserStart, m_dtcUserDue;

	FTDCFILTER m_filter;
	CDWordArray m_aPriorityColors;
	CTDCAttributeArray m_aVisibility;
	FTC_VIEW m_nView;
	CBrush m_brUIBack;
	CUIThemeFile m_theme;
	BOOL m_bCustomFilter;
	CString m_sCustomFilter;
	BOOL m_bWantHideParents;
	CWndPromptManager m_mgrPrompts;
	BOOL m_bRefreshBkgndColor;

protected:
	int DoModal() { return -1; }

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFilterBar)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL
	virtual void OnOK() { /* do nothing */ }
	virtual void OnCancel() { /* do nothing */ }
	virtual BOOL PreTranslateMessage(MSG* pMsg);

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CFilterBar)
	afx_msg void OnSize(UINT nType, int cx, int cy);
	virtual BOOL OnInitDialog();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	//}}AFX_MSG
	afx_msg void OnSelchangeFilter();
	afx_msg void OnSelchangeFilter(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSelchangeDateFilter();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg BOOL OnToolTipNotify( UINT id, NMHDR* pNMHDR, LRESULT* pResult );	
	afx_msg LRESULT OnEEBtnClick(WPARAM wp, LPARAM lp);
	DECLARE_MESSAGE_MAP()

protected:
	int ReposControls(int nWidth = -1, BOOL bCalcOnly = FALSE);
	BOOL WantShowFilter(TDC_ATTRIBUTE nType);
	void SetVisibleFilters(const CTDCAttributeArray& aFilters, BOOL bRepos);
	void RefreshUIBkgndBrush();
	COLORREF CalcUIBkgndColor() const;

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FILTERBAR_H__FAABD9A1_72C4_4731_B7A4_48251860672C__INCLUDED_)
