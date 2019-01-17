// TDLFilterComboBox.cpp : implementation file
//

#include "stdafx.h"
#include "TDLFilterDateComboBox.h"
#include "tdcstatic.h"

#include "..\shared\enstring.h"
#include "..\shared\localizer.h"
#include "..\shared\dialoghelper.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTDLFilterDateComboBox

CTDLFilterDateComboBox::CTDLFilterDateComboBox()
{
}

CTDLFilterDateComboBox::~CTDLFilterDateComboBox()
{
}


BEGIN_MESSAGE_MAP(CTDLFilterDateComboBox, CTabbedComboBox)
	//{{AFX_MSG_MAP(CTDLFilterDateComboBox)
	ON_WM_CREATE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTDLFilterDateComboBox message handlers

void CTDLFilterDateComboBox::PreSubclassWindow() 
{
	CTabbedComboBox::PreSubclassWindow();

	FillCombo();
}

int CTDLFilterDateComboBox::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CTabbedComboBox::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	FillCombo();
	
	return 0;
}

void CTDLFilterDateComboBox::FillCombo()
{
	ASSERT(GetSafeHwnd());

	if (GetCount())
		return; // already called

	CLocalizer::EnableTranslation(*this, FALSE);

	TCHAR nLetter = 'A';

	for (int nItem = 0; nItem < NUM_DATEFILTER; nItem++)
	{
		CEnString sFilter(DATE_FILTERS[nItem][0]);
		UINT nFilter = DATE_FILTERS[nItem][1];

		CString sText;

		if ((nFilter == FD_ANY) || 
			(nFilter == FD_NONE) ||
			(nFilter == FD_USER))
		{
			sText = sFilter;
		}
		else
		{
			sText.Format(_T("%c)\t%s"), nLetter, sFilter); 
			nLetter++;
		}

		VERIFY(CDialogHelper::AddString(*this, sText, nFilter) != CB_ERR);
// 		int nIndex = AddString(sText);
// 		ASSERT (nIndex != CB_ERR);
// 
// 		if (nIndex != CB_ERR)
// 			SetItemData(nIndex, nFilter);
	}

	CDialogHelper::RefreshMaxDropWidth(*this);
}

FILTER_DATE CTDLFilterDateComboBox::GetSelectedFilter() const
{
	int nSel = GetCurSel();

	if (nSel == CB_ERR)
		return FD_ANY;

	// else
	return (FILTER_DATE)GetItemData(nSel);
}

BOOL CTDLFilterDateComboBox::SelectFilter(FILTER_DATE nFilter)
{
	return CDialogHelper::SelectItemByData(*this, (DWORD)nFilter);
}
