// PreferencesTaskDefPage.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "PreferencesTaskDefPage.h"
#include "tdcenum.h"
#include "todoitem.h"

#include "..\shared\enstring.h"
#include "..\shared\misc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

const LPCTSTR ENDL = _T("\r\n");

/////////////////////////////////////////////////////////////////////////////
// CPreferencesTaskDefPage property page

IMPLEMENT_DYNCREATE(CPreferencesTaskDefPage, CPreferencesPageBase)

CPreferencesTaskDefPage::CPreferencesTaskDefPage() : 
	CPreferencesPageBase(CPreferencesTaskDefPage::IDD)
{
	//{{AFX_DATA_INIT(CPreferencesTaskDefPage)
	//}}AFX_DATA_INIT
	m_nSelAttribUse = -1;

	// attrib use
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_PRIORITY, TDCA_PRIORITY, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_RISK, TDCA_RISK, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_TIMEEST, TDCA_TIMEEST, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_ALLOCTO, TDCA_ALLOCTO, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_ALLOCBY, TDCA_ALLOCBY, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_STATUS, TDCA_STATUS, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_CATEGORY, TDCA_CATEGORY, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_PTDP_COLOR, TDCA_COLOR, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_PTDP_DUEDATE, TDCA_DUEDATE, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_PTDP_VERSION, TDCA_VERSION, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_STARTDATE, TDCA_STARTDATE, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_FLAG, TDCA_FLAG, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_EXTERNALID, TDCA_EXTERNALID, -1)); 
	m_aAttribPrefs.Add(ATTRIBPREF(IDS_TDLBC_TAGS, TDCA_TAGS, -1)); 

	m_eCost.SetMask(_T(".0123456789"), ME_LOCALIZEDECIMAL);
}

CPreferencesTaskDefPage::~CPreferencesTaskDefPage()
{
}

void CPreferencesTaskDefPage::DoDataExchange(CDataExchange* pDX)
{
	CPreferencesPageBase::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPreferencesTaskDefPage)
	DDX_Control(pDX, IDC_DEFAULTRISK, m_cbDefRisk);
	DDX_Control(pDX, IDC_DEFAULTPRIORITY, m_cbDefPriority);
	DDX_Text(pDX, IDC_DEFAULTCREATEDBY, m_sDefCreatedBy);
	DDX_Text(pDX, IDC_DEFAULTCOST, m_dDefCost);
	DDX_Check(pDX, IDC_UPDATEINHERITATTRIB, m_bUpdateInheritAttributes);
	DDX_Text(pDX, IDC_CATEGORYLIST, m_sDefCategoryList);
	DDX_Text(pDX, IDC_STATUSLIST, m_sDefStatusList);
	DDX_Text(pDX, IDC_ALLOCTOLIST, m_sDefAllocToList);
	DDX_Text(pDX, IDC_ALLOCBYLIST, m_sDefAllocByList);
	DDX_Check(pDX, IDC_CATLIST_READONLY, m_bCatListReadonly);
	DDX_Check(pDX, IDC_STATUSLIST_READONLY, m_bStatusListReadonly);
	DDX_Check(pDX, IDC_ALLOCTOLIST_READONLY, m_bAllocToListReadonly);
	DDX_Check(pDX, IDC_ALLOCBYLIST_READONLY, m_bAllocByListReadonly);
	//}}AFX_DATA_MAP
	DDX_Control(pDX, IDC_DEFAULTTIMESPENT, m_eTimeSpent);
	DDX_Control(pDX, IDC_DEFAULTTIMEEST, m_eTimeEst);
	DDX_Control(pDX, IDC_DEFAULTCOST, m_eCost);
	DDX_Control(pDX, IDC_SETDEFAULTCOLOR, m_btDefColor);
	DDX_Control(pDX, IDC_INHERITATTRIBUTES, m_lbAttribUse);
	DDX_CBPriority(pDX, IDC_DEFAULTPRIORITY, m_nDefPriority);
	DDX_CBRisk(pDX, IDC_DEFAULTRISK, m_nDefRisk);
	DDX_Text(pDX, IDC_DEFAULTTIMEEST, m_dDefTimeEst);
	DDX_Text(pDX, IDC_DEFAULTTIMESPENT, m_dDefTimeSpent);
	DDX_Text(pDX, IDC_DEFAULTALLOCTO, m_sDefAllocTo);
	DDX_LBIndex(pDX, IDC_INHERITATTRIBUTES, m_nSelAttribUse);
	DDX_Text(pDX, IDC_DEFAULTALLOCBY, m_sDefAllocBy);
	DDX_Text(pDX, IDC_DEFAULTSTATUS, m_sDefStatus);
	DDX_Text(pDX, IDC_DEFAULTTAGS, m_sDefTags);
	DDX_Text(pDX, IDC_DEFAULTCATEGORY, m_sDefCategory);
	DDX_Check(pDX, IDC_USEPARENTATTRIB, m_bInheritParentAttributes);
	DDX_Check(pDX, IDC_USECREATIONFORDEFSTARTDATE, m_bUseCreationForDefStartDate);
	DDX_Check(pDX, IDC_USECREATIONFORDEFDUEDATE, m_bUseCreationForDefDueDate);
}


BEGIN_MESSAGE_MAP(CPreferencesTaskDefPage, CPreferencesPageBase)
	//{{AFX_MSG_MAP(CPreferencesTaskDefPage)
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_SETDEFAULTCOLOR, OnSetdefaultcolor)
	ON_BN_CLICKED(IDC_USEPARENTATTRIB, OnUseparentattrib)
	ON_CLBN_CHKCHANGE(IDC_INHERITATTRIBUTES, OnAttribUseChange)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPreferencesTaskDefPage message handlers

BOOL CPreferencesTaskDefPage::OnInitDialog() 
{
	CPreferencesPageBase::OnInitDialog();
	
	m_mgrGroupLines.AddGroupLine(IDC_DEFGROUP, *this);
	m_mgrGroupLines.AddGroupLine(IDC_INHERITGROUP, *this);
	m_mgrGroupLines.AddGroupLine(IDC_DROPLISTGROUP, *this);

	GetDlgItem(IDC_INHERITATTRIBUTES)->EnableWindow(m_bInheritParentAttributes);
	GetDlgItem(IDC_UPDATEINHERITATTRIB)->EnableWindow(m_bInheritParentAttributes);
	
	m_btDefColor.SetColor(m_crDef);

	int nAttrib = m_aAttribPrefs.GetSize();
	
	while (nAttrib--)
	{
		int nIndex = m_lbAttribUse.InsertString(0, m_aAttribPrefs[nAttrib].sName);

		m_lbAttribUse.SetItemData(nIndex, m_aAttribPrefs[nAttrib].nAttrib);
		m_lbAttribUse.SetCheck(nIndex, m_aAttribPrefs[nAttrib].bUse ? 1 : 0);
	}

	// init edit prompts()
	m_mgrPrompts.SetEditPrompt(IDC_DEFAULTALLOCTO, *this, CEnString(IDS_PTDP_NAMEPROMPT));
	m_mgrPrompts.SetEditPrompt(IDC_DEFAULTALLOCBY, *this, CEnString(IDS_PTDP_NAMEPROMPT));
	m_mgrPrompts.SetEditPrompt(IDC_DEFAULTSTATUS, *this, CEnString(IDS_PTDP_STATUSPROMPT));
	m_mgrPrompts.SetEditPrompt(IDC_DEFAULTTAGS, *this, CEnString(IDS_PTDP_TAGSPROMPT));
	m_mgrPrompts.SetEditPrompt(IDC_DEFAULTCATEGORY, *this, CEnString(IDS_PTDP_CATEGORYPROMPT));
	m_mgrPrompts.SetEditPrompt(IDC_DEFAULTCREATEDBY, *this, CEnString(IDS_PTDP_NAMEPROMPT));
	m_mgrPrompts.SetComboEditPrompt(IDC_ALLOCTOLIST, *this, CEnString(IDS_PTDP_NAMEPROMPT));
	m_mgrPrompts.SetComboEditPrompt(IDC_ALLOCBYLIST, *this, CEnString(IDS_PTDP_NAMEPROMPT));
	m_mgrPrompts.SetComboEditPrompt(IDC_STATUSLIST, *this, CEnString(IDS_PTDP_STATUSPROMPT));
	m_mgrPrompts.SetComboEditPrompt(IDC_CATEGORYLIST, *this, CEnString(IDS_PTDP_CATEGORYPROMPT));

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CPreferencesTaskDefPage::SetPriorityColors(const CDWordArray& aColors)
{
	m_cbDefPriority.SetColors(aColors);
}

void CPreferencesTaskDefPage::OnOK() 
{
	CPreferencesPageBase::OnOK();
}

void CPreferencesTaskDefPage::OnSetdefaultcolor() 
{
	m_crDef = m_btDefColor.GetColor();

	CPreferencesPageBase::OnControlChange();
}

void CPreferencesTaskDefPage::OnUseparentattrib() 
{
	UpdateData();

	GetDlgItem(IDC_INHERITATTRIBUTES)->EnableWindow(m_bInheritParentAttributes);
	GetDlgItem(IDC_UPDATEINHERITATTRIB)->EnableWindow(m_bInheritParentAttributes);

	CPreferencesPageBase::OnControlChange();
}

int CPreferencesTaskDefPage::GetListItems(TDC_ATTRIBUTE nList, CStringArray& aItems) const
{
	aItems.RemoveAll();

	// include default task attributes
	CString sDef;

	switch (nList)
	{
	case TDCA_ALLOCBY:	sDef = (m_sDefAllocByList + ENDL + m_sDefAllocBy); break;
	case TDCA_ALLOCTO:	sDef = (m_sDefAllocToList + ENDL + m_sDefAllocTo); break;
	case TDCA_STATUS:	sDef = (m_sDefStatusList + ENDL + m_sDefStatus); break; 
	case TDCA_CATEGORY:	sDef = (m_sDefCategoryList + ENDL + m_sDefCategory); break;

	default: ASSERT(0);	return 0;
	}

	CStringArray aDef;
	Misc::Split(sDef, aDef, ENDL);
	Misc::AddUniqueItems(aDef, aItems);

	return aItems.GetSize();
}

BOOL CPreferencesTaskDefPage::GetListIsReadonly(TDC_ATTRIBUTE nList) const
{
	switch (nList)
	{
	case TDCA_ALLOCBY:	return m_bAllocByListReadonly;
	case TDCA_ALLOCTO:	return m_bAllocToListReadonly;
	case TDCA_STATUS:	return m_bStatusListReadonly; 
	case TDCA_CATEGORY:	return m_bCatListReadonly;
	}

	// else
	ASSERT(0);	
	return TRUE;
}

BOOL CPreferencesTaskDefPage::AddListItem(TDC_ATTRIBUTE nList, LPCTSTR szItem)
{
	UpdateData();

	CString* pList = NULL;

	switch (nList)
	{
	case TDCA_ALLOCBY:	pList = &m_sDefAllocByList; break;
	case TDCA_ALLOCTO:	pList = &m_sDefAllocToList; break;
	case TDCA_STATUS:	pList = &m_sDefStatusList; break;
	case TDCA_CATEGORY:	pList = &m_sDefCategoryList; break;

	default: ASSERT(0);	return FALSE;
	}

	// parse string into array
	CStringArray aItems;

	if (Misc::Split(*pList, aItems, ENDL))
	{
		// add to array
		if (Misc::AddUniqueItem(szItem, aItems))
		{
			// update edit control
			*pList = Misc::FormatArray(aItems, ENDL);
			UpdateData(FALSE);

			return TRUE;
		}
	}

	return FALSE; // not unique
}

BOOL CPreferencesTaskDefPage::DeleteListItem(TDC_ATTRIBUTE nList, LPCTSTR szItem)
{
	UpdateData();

	CString* pList = NULL;

	switch (nList)
	{
	case TDCA_ALLOCBY:	pList = &m_sDefAllocByList; break;
	case TDCA_ALLOCTO:	pList = &m_sDefAllocToList; break;
	case TDCA_STATUS:	pList = &m_sDefStatusList; break;
	case TDCA_CATEGORY:	pList = &m_sDefCategoryList; break;

	default: ASSERT(0);	return FALSE;
	}

	// parse string into array
	CStringArray aItems;

	if (Misc::Split(*pList, aItems, ENDL))
	{
		// delete from array
		if (Misc::RemoveItem(szItem, aItems))
		{
			// update edit control
			*pList = Misc::FormatArray(aItems, ENDL);
			UpdateData(FALSE);

			return TRUE;
		}
	}

	return FALSE; // not found
}

TDC_ATTRIBUTE CPreferencesTaskDefPage::MapCtrlIDToList(UINT nListID)
{
	switch (nListID)
	{
	case IDC_ALLOCBYLIST:	return TDCA_ALLOCBY;
	case IDC_ALLOCTOLIST:	return TDCA_ALLOCTO;
	case IDC_STATUSLIST:	return TDCA_STATUS;
	case IDC_CATEGORYLIST:	return TDCA_CATEGORY;
	}

	ASSERT(0);
	return (TDC_ATTRIBUTE)-1;
}

void CPreferencesTaskDefPage::OnAttribUseChange()
{
	UpdateData();

	ASSERT (m_nSelAttribUse >= 0);
	
	if (m_nSelAttribUse >= 0)
	{
		TDC_ATTRIBUTE nSelAttrib = (TDC_ATTRIBUTE)m_lbAttribUse.GetItemData(m_nSelAttribUse);

		// search for this item
		int nAttrib = m_aAttribPrefs.GetSize();
		
		while (nAttrib--)
		{
			if (m_aAttribPrefs[nAttrib].nAttrib == nSelAttrib)
			{
				m_aAttribPrefs[nAttrib].bUse = m_lbAttribUse.GetCheck(m_nSelAttribUse);
				break;
			}
		}
	}

	GetDlgItem(IDC_UPDATEINHERITATTRIB)->EnableWindow(m_bInheritParentAttributes);

	CPreferencesPageBase::OnControlChange();
}

BOOL CPreferencesTaskDefPage::HasCheckedAttributes() const
{
	int nAttrib = m_aAttribPrefs.GetSize();

	while (nAttrib--)
	{
		if (m_aAttribPrefs[nAttrib].bUse)
			return TRUE;
	}

	// else
	return FALSE;
}

int CPreferencesTaskDefPage::GetParentAttribsUsed(CTDCAttributeArray& aAttribs, BOOL& bUpdateAttrib) const
{
	aAttribs.RemoveAll();

	if (m_bInheritParentAttributes)
	{
		bUpdateAttrib = m_bUpdateInheritAttributes;
		int nIndex = (int)m_aAttribPrefs.GetSize();
		
		while (nIndex--)
		{
			if (m_aAttribPrefs[nIndex].bUse)
				aAttribs.Add(m_aAttribPrefs[nIndex].nAttrib);
		}
	}
	else
		bUpdateAttrib = FALSE;

	return aAttribs.GetSize();
}


void CPreferencesTaskDefPage::LoadPreferences(const CPreferences& prefs)
{
	// load settings
	m_nDefPriority = prefs.GetProfileInt(_T("Preferences"), _T("DefaultPriority"), 5); 
	m_nDefRisk = prefs.GetProfileInt(_T("Preferences"), _T("DefaultRisk"), 0); 
	m_sDefAllocTo = prefs.GetProfileString(_T("Preferences"), _T("DefaultAllocTo"));
	m_sDefAllocBy = prefs.GetProfileString(_T("Preferences"), _T("DefaultAllocBy"));
	m_sDefStatus = prefs.GetProfileString(_T("Preferences"), _T("DefaultStatus"));
	m_sDefTags = prefs.GetProfileString(_T("Preferences"), _T("DefaultTags"));
	m_sDefCategory = prefs.GetProfileString(_T("Preferences"), _T("DefaultCategory"));
	m_sDefCreatedBy = prefs.GetProfileString(_T("Preferences"), _T("DefaultCreatedBy"), Misc::GetUserName());
	m_crDef = prefs.GetProfileInt(_T("Preferences"), _T("DefaultColor"), 0);
	m_bInheritParentAttributes = prefs.GetProfileInt(_T("Preferences"), _T("InheritParentAttributes"), prefs.GetProfileInt(_T("Preferences"), _T("UseParentAttributes")));
	m_bUpdateInheritAttributes = prefs.GetProfileInt(_T("Preferences"), _T("UpdateInheritAttributes"), FALSE);
	m_bUseCreationForDefStartDate = prefs.GetProfileInt(_T("Preferences"), _T("UseCreationForDefStartDate"), TRUE);
	m_bUseCreationForDefDueDate = prefs.GetProfileInt(_T("Preferences"), _T("UseCreationForDefDueDate"), FALSE);
	m_dDefCost = Misc::Atof(prefs.GetProfileString(_T("Preferences"), _T("DefaultCost"), _T("0")));
	m_dDefTimeEst = prefs.GetProfileDouble(_T("Preferences"), _T("DefaultTimeEstimate"), 0);
	m_eTimeEst.SetUnits((TCHAR)prefs.GetProfileInt(_T("Preferences"), _T("DefaultTimeEstUnits"), THU_HOURS));
	m_dDefTimeSpent = prefs.GetProfileDouble(_T("Preferences"), _T("DefaultTimeSpent"), 0);
	m_eTimeSpent.SetUnits((TCHAR)prefs.GetProfileInt(_T("Preferences"), _T("DefaultTimeSpentUnits"), THU_HOURS));
	m_bCatListReadonly = prefs.GetProfileInt(_T("Preferences"), _T("CatListReadonly"), FALSE);
	m_bStatusListReadonly = prefs.GetProfileInt(_T("Preferences"), _T("StatusListReadonly"), FALSE);
	m_bAllocToListReadonly = prefs.GetProfileInt(_T("Preferences"), _T("AllocToListReadonly"), FALSE);
	m_bAllocByListReadonly = prefs.GetProfileInt(_T("Preferences"), _T("AllocByListReadonly"), FALSE);

   // attribute use
	int nIndex = m_aAttribPrefs.GetSize();
	
	while (nIndex--)
	{
		CString sKey = Misc::MakeKey(_T("Attrib%d"), m_aAttribPrefs[nIndex].nAttrib);
		m_aAttribPrefs[nIndex].bUse = prefs.GetProfileInt(_T("Preferences\\AttribUse"), sKey, FALSE);
	}

	// combo lists
	CStringArray aItems;

	prefs.GetProfileArray(_T("Preferences\\CategoryList"), aItems);
	Misc::SortArray(aItems);
	m_sDefCategoryList = Misc::FormatArray(aItems, ENDL);

	prefs.GetProfileArray(_T("Preferences\\StatusList"), aItems);
	Misc::SortArray(aItems);
	m_sDefStatusList = Misc::FormatArray(aItems, ENDL);

	prefs.GetProfileArray(_T("Preferences\\AllocToList"), aItems);
	Misc::SortArray(aItems);
	m_sDefAllocToList = Misc::FormatArray(aItems, ENDL);

	prefs.GetProfileArray(_T("Preferences\\AllocByList"), aItems);
	Misc::SortArray(aItems);
	m_sDefAllocByList = Misc::FormatArray(aItems, ENDL);
}

void CPreferencesTaskDefPage::SavePreferences(CPreferences& prefs)
{
	// save settings
	prefs.WriteProfileInt(_T("Preferences"), _T("DefaultPriority"), m_nDefPriority);
	prefs.WriteProfileInt(_T("Preferences"), _T("DefaultRisk"), m_nDefRisk);
	prefs.WriteProfileString(_T("Preferences"), _T("DefaultAllocTo"), m_sDefAllocTo);
	prefs.WriteProfileString(_T("Preferences"), _T("DefaultAllocBy"), m_sDefAllocBy);
	prefs.WriteProfileString(_T("Preferences"), _T("DefaultStatus"), m_sDefStatus);
	prefs.WriteProfileString(_T("Preferences"), _T("DefaultTags"), m_sDefTags);
	prefs.WriteProfileString(_T("Preferences"), _T("DefaultCategory"), m_sDefCategory);
	prefs.WriteProfileString(_T("Preferences"), _T("DefaultCreatedBy"), m_sDefCreatedBy);
	prefs.WriteProfileInt(_T("Preferences"), _T("DefaultColor"), m_crDef);
	prefs.WriteProfileInt(_T("Preferences"), _T("InheritParentAttributes"), m_bInheritParentAttributes);
	prefs.WriteProfileInt(_T("Preferences"), _T("UpdateInheritAttributes"), m_bUpdateInheritAttributes);
	prefs.WriteProfileInt(_T("Preferences"), _T("UseCreationForDefStartDate"), m_bUseCreationForDefStartDate);
	prefs.WriteProfileInt(_T("Preferences"), _T("UseCreationForDefDueDate"), m_bUseCreationForDefDueDate);
	prefs.WriteProfileDouble(_T("Preferences"), _T("DefaultCost"), m_dDefCost);
	prefs.WriteProfileDouble(_T("Preferences"), _T("DefaultTimeEstimate"), m_dDefTimeEst);
	prefs.WriteProfileInt(_T("Preferences"), _T("DefaultTimeEstUnits"), m_eTimeEst.GetUnits());
	prefs.WriteProfileDouble(_T("Preferences"), _T("DefaultTimeSpent"), m_dDefTimeSpent);
	prefs.WriteProfileInt(_T("Preferences"), _T("DefaultTimeSpentUnits"), m_eTimeSpent.GetUnits());
	prefs.WriteProfileInt(_T("Preferences"), _T("CatListReadonly"), m_bCatListReadonly);
	prefs.WriteProfileInt(_T("Preferences"), _T("StatusListReadonly"), m_bStatusListReadonly);
	prefs.WriteProfileInt(_T("Preferences"), _T("AllocToListReadonly"), m_bAllocToListReadonly);
	prefs.WriteProfileInt(_T("Preferences"), _T("AllocByListReadonly"), m_bAllocByListReadonly);
	
	// attribute usage
	int nIndex = m_aAttribPrefs.GetSize();

	while (nIndex--)
	{
		CString sKey = Misc::MakeKey(_T("Attrib%d"), m_aAttribPrefs[nIndex].nAttrib);
		prefs.WriteProfileInt(_T("Preferences\\AttribUse"), sKey, m_aAttribPrefs[nIndex].bUse);
	}

	// combo lists
	CStringArray aItems;

	Misc::Split(m_sDefCategoryList, aItems, ENDL);
	prefs.WriteProfileArray(_T("Preferences\\CategoryList"), aItems);

	Misc::Split(m_sDefStatusList, aItems, ENDL);
	prefs.WriteProfileArray(_T("Preferences\\StatusList"), aItems);

	Misc::Split(m_sDefAllocToList, aItems, ENDL);
	prefs.WriteProfileArray(_T("Preferences\\AllocToList"), aItems);

	Misc::Split(m_sDefAllocByList, aItems, ENDL);
	prefs.WriteProfileArray(_T("Preferences\\AllocByList"), aItems);
}

void CPreferencesTaskDefPage::GetTaskAttributes(TODOITEM& tdiDefault) const
{
	tdiDefault.sTitle = CEnString(IDS_TASK);
	tdiDefault.color = m_crDef;
	tdiDefault.sAllocBy = m_sDefAllocBy;
	tdiDefault.sStatus = m_sDefStatus;
	tdiDefault.dTimeEstimate = m_dDefTimeEst;
	tdiDefault.dTimeSpent = m_dDefTimeSpent;
	tdiDefault.nTimeEstUnits = m_eTimeEst.GetUnits();
	tdiDefault.nTimeSpentUnits = m_eTimeSpent.GetUnits();
	tdiDefault.dCost = m_dDefCost;
	tdiDefault.nPriority = m_nDefPriority;
	tdiDefault.nRisk = m_nDefRisk;
	tdiDefault.sCreatedBy = m_sDefCreatedBy;
	
	tdiDefault.dateStart.m_dt = (m_bUseCreationForDefStartDate ? -1 : 0);
	tdiDefault.dateStart.m_status = COleDateTime::null;
	tdiDefault.dateDue.m_dt = (m_bUseCreationForDefDueDate ? -1 : 0);
	tdiDefault.dateDue.m_status = COleDateTime::null;
	
	Misc::Split(m_sDefCategory, tdiDefault.aCategories);
	Misc::Split(m_sDefAllocTo, tdiDefault.aAllocTo);
	Misc::Split(m_sDefTags, tdiDefault.aTags);
}

