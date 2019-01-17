#if !defined(AFX_FINDTASKEXPRESSIONLISTCTRL_H__42272309_6C54_4901_A56D_D6FDA87F1E48__INCLUDED_)
#define AFX_FINDTASKEXPRESSIONLISTCTRL_H__42272309_6C54_4901_A56D_D6FDA87F1E48__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FindTaskExpressionListCtrl.h : header file
//

#include "tdcenum.h"
#include "tdcstruct.h"
#include "tdlfindtaskattributecombobox.h"
#include "tdlprioritycombobox.h"
#include "tdlriskcombobox.h"

#include "..\shared\InputListCtrl.h"
#include "..\shared\timeedit.h"
#include "..\shared\datetimectrlex.h"
#include "..\shared\checkcombobox.h"

#include <afxtempl.h>

/////////////////////////////////////////////////////////////////////////////
// CTDLFindTaskExpressionListCtrl window

class CTDLFindTaskExpressionListCtrl : public CInputListCtrl
{
// Construction
public:
	CTDLFindTaskExpressionListCtrl();

	void SetSearchParams(const SEARCHPARAM& param);
	void SetSearchParams(const CSearchParamArray& params);
	int GetSearchParams(CSearchParamArray& params) const;

	void ClearSearch();

	void SetCustomAttributes(const CTDCCustomAttribDefinitionArray& aAttribDefs);
	void SetAttributeListData(const TDCAUTOLISTDATA& tld);

	BOOL AddRule();
	BOOL DeleteSelectedRule();
	BOOL CanDeleteSelectedRule() const { return CanDeleteSelectedCell(); }
	BOOL HasRules() const { return m_aSearchParams.GetSize(); }

	void MoveSelectedRuleUp();
	BOOL CanMoveSelectedRuleUp() const;
	void MoveSelectedRuleDown();
	BOOL CanMoveSelectedRuleDown() const;

// Attributes
protected:
	CComboBox						m_cbOperators;
	CTDLFindTaskAttributeComboBox	m_cbAttributes;
	CComboBox						m_cbAndOr;
	CDateTimeCtrlEx					m_dtcDate;
	CTimeEdit						m_eTime;
	CCheckComboBox					m_cbListValues;
	CTDLPriorityComboBox			m_cbPriority;
	CTDLRiskComboBox				m_cbRisk;
	CSearchParamArray				m_aSearchParams;
	CTDCCustomAttribDefinitionArray m_aAttribDefs;
	TDCAUTOLISTDATA					m_tldListContents;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTDLFindTaskExpressionListCtrl)
	protected:
	virtual void PreSubclassWindow();
	//}}AFX_VIRTUAL
	virtual BOOL PreTranslateMessage(MSG* pMsg);

// Implementation
public:
	virtual ~CTDLFindTaskExpressionListCtrl();

	// Generated message map functions
protected:
	//{{AFX_MSG(CTDLFindTaskExpressionListCtrl)
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	//}}AFX_MSG
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnAttribEditCancel();
	afx_msg void OnAttribEditOK();
	afx_msg void OnOperatorEditCancel();
	afx_msg void OnOperatorEditOK();
	afx_msg void OnAndOrEditCancel();
	afx_msg void OnAndOrEditOK();
	afx_msg void OnListValuesEditCancel();
	afx_msg void OnListValuesEditOK();
	afx_msg void OnPriorityEditCancel();
	afx_msg void OnPriorityEditOK();
	afx_msg void OnRiskEditCancel();
	afx_msg void OnRiskEditOK();
	afx_msg void MeasureItem(LPMEASUREITEMSTRUCT lpMeasureItemStruct);
	afx_msg void OnValueEditOK(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg BOOL OnSelItemChanged(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDateChange(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDateCloseUp(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnTimeChange();
	afx_msg LRESULT OnTimeUnitsChange(WPARAM wp, LPARAM lp);
	DECLARE_MESSAGE_MAP()

	virtual void EditCell(int nItem, int nCol);
	virtual BOOL IsEditing() const;
	virtual BOOL DeleteSelectedCell();
	virtual BOOL CanEditSelectedCell() const;
	virtual IL_COLUMNTYPE GetCellType(int nRow, int nCol) const;
	virtual void OnCancelEdit();

	void PrepareEdit(int nRow, int nCol);
	void PrepareControl(CWnd& ctrl, int nRow, int nCol);
	int GetValueType(int nRow) const;
	void BuildListCtrl();
	int InsertRule(int nRow, const SEARCHPARAM& sp);
	int GetRuleCount() const { return m_aSearchParams.GetSize(); }
	CWnd* GetEditControl(int nRow, int nCol);
	void ValidateListData() const;
	void UpdateValueColumnText(int nRow);
	void AddOperatorToCombo(FIND_OPERATOR op);
	void RefreshAndOrColumnText();
	void HideAllControls(const CWnd* pWndIgnore = NULL);

	static CString GetOpName(FIND_OPERATOR op);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FINDTASKEXPRESSIONLISTCTRL_H__42272309_6C54_4901_A56D_D6FDA87F1E48__INCLUDED_)
