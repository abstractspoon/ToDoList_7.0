// TDLTaskViewListBox.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "TDLTaskViewListBox.h"

#include "..\shared\DialogHelper.h"
#include "..\shared\UIExtensionMgr.h"
#include "..\shared\EnString.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////

#define WM_INITLISTBOX (WM_APP+1)
#define LISTVIEW_INDEX 0xFFFF

/////////////////////////////////////////////////////////////////////////////
// CTDLTaskViewListBox

CTDLTaskViewListBox::CTDLTaskViewListBox(const CUIExtensionMgr* pMgr) : m_pMgrUIExt(pMgr)
{
}

CTDLTaskViewListBox::~CTDLTaskViewListBox()
{
}


BEGIN_MESSAGE_MAP(CTDLTaskViewListBox, CCheckListBoxEx) 
	//{{AFX_MSG_MAP(CTDLTaskViewListBox)
	ON_WM_CREATE()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_INITLISTBOX, OnInitListBox)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLTaskViewListBox message handlers

void CTDLTaskViewListBox::PreSubclassWindow() 
{
	if (m_pMgrUIExt)
		PostMessage(WM_INITLISTBOX);
	
	CCheckListBoxEx::PreSubclassWindow();
}

int CTDLTaskViewListBox::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CCheckListBoxEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	if (!m_pMgrUIExt)
		return -1;
	
	PostMessage(WM_INITLISTBOX);
	
	return 0;
}

LRESULT CTDLTaskViewListBox::OnInitListBox(WPARAM /*wp*/, LPARAM /*lp*/)
{
	BuildList();
	return 0L;
}

void CTDLTaskViewListBox::BuildList()
{
	// once only
	if (CCheckListBoxEx::GetCount())
		return;

	ASSERT(m_pMgrUIExt);

	if (m_pMgrUIExt)
	{
		// 'list view' is special
		int nItem = AddString(CString(MAKEINTRESOURCE(IDS_LISTVIEW)));

		SetItemData(nItem, LISTVIEW_INDEX);
		SetCheck(nItem, TRUE); // default

		// add rest of extensions
		int nNumExt = m_pMgrUIExt->GetNumUIExtensions();

		for (int nExt = 0; nExt < nNumExt; nExt++)
		{
			nItem = AddString(m_pMgrUIExt->GetUIExtensionMenuText(nExt));

			SetItemData(nItem, nExt);
			SetCheck(nItem, TRUE); // default
		}

		if (m_aVisibleViews.GetSize())
		{
			SetVisibleViews(m_aVisibleViews);
			m_aVisibleViews.RemoveAll();
		}
	}
}

void CTDLTaskViewListBox::SetVisibleViews(const CStringArray& aTypeIDs)
{
	if (GetSafeHwnd())
	{
		ASSERT(m_pMgrUIExt);

		if (!m_pMgrUIExt)
			return;

		ASSERT(GetCount());

		SetAllChecked(FALSE);

		int nExt = aTypeIDs.GetSize();

		while (nExt--)
		{
			CString sTypeID = aTypeIDs[nExt];

			int nFind = m_pMgrUIExt->FindUIExtension(sTypeID);

			if ((nFind == -1) && (sTypeID == LISTVIEW_TYPE))
				nFind = LISTVIEW_INDEX;

			ASSERT(nFind != -1);

			if (nFind != -1)
			{
				int nItem = CDialogHelper::FindItemByData(*this, nFind);
				ASSERT(nItem != -1);
				
				SetCheck(nItem, TRUE);
			}
		}
	}
	else
		m_aVisibleViews.Copy(aTypeIDs);

}

int CTDLTaskViewListBox::GetVisibleViews(CStringArray& aTypeIDs) const
{
	if (GetSafeHwnd())
	{
		ASSERT(m_pMgrUIExt);

		if (!m_pMgrUIExt)
			return 0;

		aTypeIDs.RemoveAll();

		int nItem = GetCount();
		CTDLTaskViewListBox* pThis = const_cast<CTDLTaskViewListBox*>(this);

		while (nItem--)
		{
			if (pThis->GetCheck(nItem))
			{
				DWORD dwItemData = GetItemData(nItem);

				if (dwItemData == LISTVIEW_INDEX)
					aTypeIDs.Add(LISTVIEW_TYPE);
				else
					aTypeIDs.Add(m_pMgrUIExt->GetUIExtensionTypeID((int)dwItemData));
			}
		}
	}
	else
		aTypeIDs.Copy(m_aVisibleViews);

	return aTypeIDs.GetSize();
}

void CTDLTaskViewListBox::OnDestroy() 
{
	GetVisibleViews(m_aVisibleViews);

	CCheckListBoxEx::OnDestroy();
}
