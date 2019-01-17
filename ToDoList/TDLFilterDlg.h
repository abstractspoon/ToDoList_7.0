#if !defined(AFX_FILTERDLG_H__9BFA276C_9BD6_4B23_A714_BB9D059CC2EB__INCLUDED_)
#define AFX_FILTERDLG_H__9BFA276C_9BD6_4B23_A714_BB9D059CC2EB__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FilterDlg.h : header file
//

#include "filteredtodoctrl.h"
#include "tdlfiltercombobox.h"
#include "tdlfilteroptioncombobox.h"
#include "tdlprioritycombobox.h"
#include "tdlriskcombobox.h"
#include "tdlfilterdatecombobox.h"

#include "..\shared\dialoghelper.h"
#include "..\shared\encheckcombobox.h"
#include "..\shared\wndprompt.h"

#include <afxtempl.h>

/////////////////////////////////////////////////////////////////////////////
// CTDLFilterDlg dialog

class CTDLFilterDlg : public CDialog, public CDialogHelper
{
// Construction
public:
	CTDLFilterDlg(BOOL bMultiSelFilters, CWnd* pParent = NULL);   // standard constructor

	int DoModal(const CStringArray& aCustomFilters, 
				const CFilteredToDoCtrl& tdc, 
				const CDWordArray& aPriorityColors);

	void GetFilter(FTDCFILTER& filter, CString& sCustom);

protected:
// Dialog Data
	//{{AFX_DATA(CTDLFilterDlg)
	enum { IDD = IDD_FILTER_DIALOG };
	//}}AFX_DATA
	CEnCheckComboBox	m_cbCategoryFilter;
	CEnCheckComboBox	m_cbStatusFilter;
	CEnCheckComboBox	m_cbAllocToFilter;
	CEnCheckComboBox	m_cbAllocByFilter;
	CEnCheckComboBox	m_cbVersionFilter;
	CEnCheckComboBox	m_cbTagFilter;
	CTDLFilterComboBox	m_cbTaskFilter;
	CTDLFilterDateComboBox	m_cbStartFilter;
	CTDLFilterDateComboBox	m_cbDueFilter;
	CTDLPriorityComboBox	m_cbPriorityFilter;
	CTDLRiskComboBox	m_cbRiskFilter;
	FTDCFILTER m_filter;
	TDCAUTOLISTDATA m_tldListData;
	CDWordArray m_aPriorityColors;
	CTDLFilterOptionComboBox m_cbOptions;
	FTC_VIEW m_nView;
	CString m_sCustomFilter;
	BOOL m_bWantHideParents;
	CStringArray m_aCustomFilters;
	CWndPromptManager m_mgrPrompts;
	CDateTimeCtrl m_dtcUserStart, m_dtcUserDue;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTDLFilterDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CTDLFilterDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnClearfilter();
	afx_msg void OnSelchangeFiltercombo();
	//}}AFX_MSG
	afx_msg void OnSelchangeDatecombo();
	afx_msg BOOL OnToolTipNotify( UINT id, NMHDR* pNMHDR, LRESULT* pResult );	
	DECLARE_MESSAGE_MAP()
		
	void EnableDisableControls();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FILTERDLG_H__9BFA276C_9BD6_4B23_A714_BB9D059CC2EB__INCLUDED_)
