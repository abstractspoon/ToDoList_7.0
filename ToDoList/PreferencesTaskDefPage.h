#if !defined(AFX_PREFERENCESTASKDEFPAGE_H__852964E3_4ABD_4B66_88BA_F553177616F2__INCLUDED_)
#define AFX_PREFERENCESTASKDEFPAGE_H__852964E3_4ABD_4B66_88BA_F553177616F2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// PreferencesTaskDefPage.h : header file
//

#include "tdcenum.h"
#include "tdlprioritycombobox.h"
#include "tdlriskcombobox.h"

#include <afxtempl.h>

#include "..\shared\preferencesbase.h"
#include "..\shared\colorbutton.h"
#include "..\shared\timeedit.h"
#include "..\shared\wndPrompt.h"
#include "..\shared\autocombobox.h"
#include "..\shared\groupline.h"
#include "..\shared\maskedit.h"
#include "..\Shared\checklistboxex.h"

/////////////////////////////////////////////////////////////////////////////
// CPreferencesTaskDefPage dialog

class TODOITEM;

/////////////////////////////////////////////////////////////////////////////

class CPreferencesTaskDefPage : public CPreferencesPageBase
{
	DECLARE_DYNCREATE(CPreferencesTaskDefPage)

// Construction
public:
	CPreferencesTaskDefPage();
	~CPreferencesTaskDefPage();

	void SetPriorityColors(const CDWordArray& aColors);

	void GetTaskAttributes(TODOITEM& tdiDefault) const;
	int GetParentAttribsUsed(CTDCAttributeArray& aAttribs, BOOL& bUpdateAttrib) const;

	int GetListItems(TDC_ATTRIBUTE nList, CStringArray& aItems) const;
	BOOL AddListItem(TDC_ATTRIBUTE nList, LPCTSTR szItem);
	BOOL DeleteListItem(TDC_ATTRIBUTE nList, LPCTSTR szItem);
	BOOL GetListIsReadonly(TDC_ATTRIBUTE nList) const;

protected:
// Dialog Data
	//{{AFX_DATA(CPreferencesTaskDefPage)
	enum { IDD = IDD_PREFTASKDEF_PAGE };
	CTDLRiskComboBox	m_cbDefRisk;
	CTDLPriorityComboBox	m_cbDefPriority;
	CString	m_sDefCreatedBy;
	double	m_dDefCost;
	BOOL	m_bUpdateInheritAttributes;
	CString	m_sDefCategoryList;
	CString	m_sDefStatusList;
	CString	m_sDefAllocToList;
	CString	m_sDefAllocByList;
	BOOL	m_bCatListReadonly;
	BOOL	m_bStatusListReadonly;
	BOOL	m_bAllocToListReadonly;
	BOOL	m_bAllocByListReadonly;
	//}}AFX_DATA
	CTimeEdit	m_eTimeEst;
	CTimeEdit	m_eTimeSpent;
	CMaskEdit m_eCost;
	CCheckListBoxEx	m_lbAttribUse;
	int		m_nDefPriority;
	int		m_nDefRisk;
	double	m_dDefTimeEst, m_dDefTimeSpent;
	CString	m_sDefAllocTo;
	CString	m_sDefAllocBy;
	CString	m_sDefStatus;
	CString	m_sDefTags;
	CString	m_sDefCategory;
	CColorButton	m_btDefColor;
	COLORREF m_crDef;
	BOOL	m_bInheritParentAttributes;
	int		m_nSelAttribUse;
	BOOL	m_bUseCreationForDefStartDate;
	BOOL	m_bUseCreationForDefDueDate;
	CWndPromptManager m_mgrPrompts;
	CGroupLineManager m_mgrGroupLines;

	struct ATTRIBPREF
	{
		ATTRIBPREF() : nAttrib(TDCA_NONE), bUse(FALSE) {}
		ATTRIBPREF(UINT nIDName, TDC_ATTRIBUTE attrib, BOOL use) : nAttrib(attrib), bUse(use) { sName.LoadString(nIDName); }

		CString sName;
		TDC_ATTRIBUTE nAttrib;
		BOOL bUse;
	};
	CArray<ATTRIBPREF, ATTRIBPREF&> m_aAttribPrefs;


// Overrides
	// ClassWizard generate virtual function overrides
	//{{AFX_VIRTUAL(CPreferencesTaskDefPage)
	public:
	virtual void OnOK();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CPreferencesTaskDefPage)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	afx_msg void OnSetdefaultcolor();
	afx_msg void OnUseparentattrib();
	afx_msg void OnAttribUseChange();
	DECLARE_MESSAGE_MAP()
		
	virtual void LoadPreferences(const CPreferences& prefs);
	virtual void SavePreferences(CPreferences& prefs);
	
	BOOL HasCheckedAttributes() const;
	
	static TDC_ATTRIBUTE MapCtrlIDToList(UINT nListID);

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PREFERENCESTASKDEFPAGE_H__852964E3_4ABD_4B66_88BA_F553177616F2__INCLUDED_)
