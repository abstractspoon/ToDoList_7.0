// colorcombobox.cpp : implementation file
//

#include "stdafx.h"
#include "colorcombobox.h"
#include "dlgunits.h"
#include "graphicsmisc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CColorComboBox

CColorComboBox::CColorComboBox(BOOL bRoundRect) : m_bRoundRect(bRoundRect)
{
}

CColorComboBox::~CColorComboBox()
{
}


BEGIN_MESSAGE_MAP(CColorComboBox, COwnerdrawComboBoxBase)
	//{{AFX_MSG_MAP(CColorComboBox)
	ON_WM_CREATE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CColorComboBox message handlers

void CColorComboBox::DrawItemText(CDC& dc, const CRect& rect, int nItem, UINT nItemState, 
								  DWORD dwItemData, const CString& sItem, BOOL bList)
{
	CDlgUnits dlu(GetParent());
	
 	BOOL bHasColor = (dwItemData != (DWORD)-1);
	CRect rColor(rect), rText(rect);
	
	// draw text
	if (!sItem.IsEmpty())
	{
		rColor.right = rColor.left + (bHasColor ? rColor.Height() : 0);
		rText.left = rColor.right + 2;

		COwnerdrawComboBoxBase::DrawItemText(dc, rText, nItem, nItemState, dwItemData, sItem, bList);
	}

	// draw color
	if (bHasColor)
	{
		int nCornerRadius = m_bRoundRect ? (rColor.Width() / 4) : 0;
		COLORREF crFill = (COLORREF)dwItemData, crBorder = CLR_NONE;
		
		if (!IsWindowEnabled() || (nItemState & ODS_GRAYED))
		{
			crFill = GetSysColor(COLOR_3DFACE);
			crBorder = GetSysColor(COLOR_3DDKSHADOW);
		}
		else
		{
			crBorder = GraphicsMisc::Darker(crFill, 0.5);
		}
		
		GraphicsMisc::DrawRect(&dc, rColor, crFill, crBorder, nCornerRadius);
	}
}	

int CColorComboBox::AddColor(COLORREF color, LPCTSTR szDescription)
{
	ASSERT (GetStyle() & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE));

	if (szDescription)
		ASSERT (GetStyle() & CBS_HASSTRINGS);
	else
		szDescription = _T("");

	int nIndex = AddString(szDescription);

	if (nIndex != CB_ERR)
		SetItemData(nIndex, (DWORD)color);

	return nIndex;
}

int CColorComboBox::InsertColor(int nIndex, COLORREF color, LPCTSTR szDescription)
{
	ASSERT (GetStyle() & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE));

	if (szDescription)
		ASSERT (GetStyle() & CBS_HASSTRINGS);
	else
		szDescription = _T("");

	nIndex = InsertString(nIndex, szDescription);

	if (nIndex != CB_ERR)
		SetItemData(nIndex, (DWORD)color);

	return nIndex;
}

COLORREF CColorComboBox::SetColor(int nIndex, COLORREF color)
{
	ASSERT (GetStyle() & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE));
	ASSERT (nIndex >= 0 && nIndex < GetCount());

	COLORREF crOld = (COLORREF)GetItemData(nIndex);
	SetItemData(nIndex, (DWORD)color);

	return crOld;
}
