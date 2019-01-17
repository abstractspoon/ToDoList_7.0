/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "EnHeaderCtrl.h"
#include "themed.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////

enum // item flags
{
	EHCF_NOTRACK	= 0x0001,
	EHCF_HIDDEN		= 0x0002,
	EHCF_TRACKED	= 0x0004,
	EHCF_SORTUP		= 0x0008,
	EHCF_SORTDOWN	= 0x0010,
	EHCF_SORTMASK	= (EHCF_SORTUP | EHCS_DOWN),
	EHCF_NODRAG		= 0x0020,
};

/////////////////////////////////////////////////////////////////////////////

const UINT WM_EHC_RECALCITEMTOOLTIPRECTS = (WM_APP + 1);

/////////////////////////////////////////////////////////////////////////////
// CEnHeaderCtrl

CEnHeaderCtrl::CEnHeaderCtrl() : m_nRowCount(1), m_bEnableTracking(TRUE)
{
}

CEnHeaderCtrl::~CEnHeaderCtrl()
{
}


BEGIN_MESSAGE_MAP(CEnHeaderCtrl, CHeaderCtrl)
	//{{AFX_MSG_MAP(CEnHeaderCtrl)
	ON_WM_SETCURSOR()
	ON_WM_CONTEXTMENU()
	//}}AFX_MSG_MAP
	ON_WM_LBUTTONDBLCLK()
	ON_NOTIFY_REFLECT(HDN_BEGINTRACK, OnBeginTrackHeader)
	ON_NOTIFY_REFLECT(HDN_ENDTRACK, OnEndTrackHeader)
	ON_NOTIFY_REFLECT(HDN_BEGINDRAG, OnBeginDragHeader)
	ON_NOTIFY_REFLECT_EX(HDN_ENDDRAG, OnEndDragHeader)
	ON_MESSAGE(HDM_LAYOUT, OnLayout)
	ON_MESSAGE(HDM_INSERTITEM, OnInsertItem)
	ON_MESSAGE(HDM_DELETEITEM, OnDeleteItem)
	ON_MESSAGE(WM_EHC_RECALCITEMTOOLTIPRECTS, OnRecalcItemTooltipRects)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CEnHeaderCtrl message handlers

void CEnHeaderCtrl::DeleteAllItems()
{
	while (DeleteItem(0));
}

void CEnHeaderCtrl::SetRowCount(int nRows)
{
	if (nRows <= 0 || nRows == m_nRowCount)
		return;

	m_nRowCount = nRows;

	// force resize
	if (GetSafeHwnd())
		SetWindowPos(NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER); 
}

LRESULT CEnHeaderCtrl::OnLayout(WPARAM wp, LPARAM lp)
{
	// default processing
	LRESULT lResult = CHeaderCtrl::DefWindowProc(HDM_LAYOUT, wp, lp);

	// our handling
	if (m_nRowCount > 1)
	{
		HD_LAYOUT* pLayout = (HD_LAYOUT*)lp;
		int nHeight = (int)(((0.6 * m_nRowCount) + 0.4) * pLayout->pwpos->cy);

		pLayout->pwpos->cy = nHeight;
		pLayout->prc->top = nHeight;

		Invalidate(TRUE);
	}

	return lResult;
}

LRESULT CEnHeaderCtrl::OnInsertItem(WPARAM wp, LPARAM lp)
{
	// default processing
	LRESULT lResult = Default();

	// our handling
	if ((int)lResult != -1)
		m_aItemFlags.SetAtGrow(lResult, 0);

	return lResult;
}

LRESULT CEnHeaderCtrl::OnDeleteItem(WPARAM wp, LPARAM lp)
{
	// default processing
	LRESULT lResult = Default();

	// our handling
	if (lResult)
	{
		if ((int)wp < m_aItemFlags.GetSize())
			m_aItemFlags.RemoveAt(wp);
	}

	return lResult;
}

void CEnHeaderCtrl::OnEndTrackHeader(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LPNMHEADER pNMH = (LPNMHEADER)pNMHDR;

	if (IsItemTrackable(pNMH->iItem))
	{
		// mark item as having been tracked
		ModifyItemFlags(pNMH->iItem, EHCF_TRACKED, TRUE);

		RecalcItemTooltipRects();
	}
}

void CEnHeaderCtrl::OnBeginTrackHeader(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LPNMHEADER pNMH = (LPNMHEADER)pNMHDR;

	// return TRUE to disable tracking
	*pResult = ((pNMH->iItem == -1) || !IsItemTrackable(pNMH->iItem));
}

void CEnHeaderCtrl::OnBeginDragHeader(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LPNMHEADER pNMH = (LPNMHEADER)pNMHDR;

	// Note: Dragging can only be prevented by rejecting from OnEndDragHeader
	*pResult = (pNMH->iItem == -1);
}

BOOL CEnHeaderCtrl::OnEndDragHeader(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LPNMHEADER pNMH = (LPNMHEADER)pNMHDR;
	*pResult = 0;

	int nItem = pNMH->iItem;

	// return TRUE to revert drag
	if ((nItem == -1) || !IsItemDraggable(nItem))
	{
		*pResult = 1;
	}
	else // don't allow non-draggable items to move
	{
		CPoint ptCursor(GetMessagePos());
		ScreenToClient(&ptCursor);

		// where are we inserting this item
		int nInsert = HitTest(ptCursor);
		ASSERT(nInsert != -1);

		// inserting before or after this item
		CRect rItem;
		GetItemRect(nInsert, rItem);

		if (ptCursor.x >= rItem.CenterPoint().x)
			nInsert++;

		// don't allow item to cross a non-draggable item
		if (nItem < nInsert)
		{
			for (nItem++; nItem < nInsert; nItem++)
			{
				if (!IsItemDraggable(nItem))
				{
					*pResult = 1; // prevent
					break;
				}
			}
		}
		else if (nItem > nInsert)
		{
			while (--nItem >= nInsert)
			{
				if (!IsItemDraggable(nItem))
				{
					*pResult = 1; // prevent
					break;
				}
			}
		}
	}

	if (*pResult == 0) // success
		PostMessage(WM_EHC_RECALCITEMTOOLTIPRECTS);

	return FALSE; // continue routing
}

int CEnHeaderCtrl::HitTest(CPoint ptClient, UINT* pFlags) const
{
	HDHITTESTINFO hdhit = { 0 };
	hdhit.pt = ptClient;

	int nHit = ::SendMessage(GetSafeHwnd(), HDM_HITTEST, 0, (LPARAM)&hdhit);

	if (pFlags)
		*pFlags = hdhit.flags;

	return nHit;
}

BOOL CEnHeaderCtrl::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message)
{
	CPoint ptCursor = GetCurrentMessage()->pt;
	ScreenToClient(&ptCursor);

	UINT nFlags = 0;
	int nItem = HitTest(ptCursor, &nFlags);

	if (nItem != -1)
	{
		if ((nFlags & (HHT_ONDIVIDER | HHT_ONDIVOPEN)) && !IsItemTrackable(nItem))
		{
			AfxGetApp()->LoadStandardCursor(IDC_ARROW);
			return TRUE;
		}
	}

	// else 
	return CHeaderCtrl::OnSetCursor(pWnd, nHitTest, message);
}

void CEnHeaderCtrl::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	int nItem = HitTest(point, &nFlags); // reuse nFlags
	
	if (nItem != -1)
	{
		if ((nFlags & (HHT_ONDIVIDER | HHT_ONDIVOPEN)) && !IsItemTrackable(nItem))
		{
			return; // eat it
		}
	}
	
	// else
	CHeaderCtrl::OnLButtonDblClk(nFlags, point);
}

void CEnHeaderCtrl::ShowItem(int nItem, BOOL bShow)
{
	if (ModifyItemFlags(nItem, EHCF_HIDDEN, !bShow))
	{
		// resize as necessary
		if (bShow)
		{
			ASSERT(nItem < m_aItemWidths.GetSize());

			int nWidth = m_aItemWidths[nItem];
			SetItemWidth(nItem, nWidth);
		}
		else // hide
		{
			m_aItemWidths.SetAtGrow(nItem, (WORD)GetItemWidth(nItem));
			SetItemWidth(nItem, 0);
		}
	}
}

BOOL CEnHeaderCtrl::IsItemVisible(int nItem) const
{
	if (ItemHasFlag(nItem, EHCF_HIDDEN))
		return FALSE;

	return TRUE;
}

int CEnHeaderCtrl::GetFirstVisibleItem() const
{
	CIntArray aItems;
	int nNumItem = GetItemOrder(aItems);
	
	for (int nOrder = 0; nOrder < nNumItem; nOrder++)
	{
		int nItem = aItems[nOrder];

		if (IsItemVisible(nItem) && GetItemWidth(nItem))
			return nItem;
	}
	
	// all hidden
	return -1;
}

int CEnHeaderCtrl::GetLastVisibleItem() const
{
	CIntArray aOrder;
	int nOrder = GetItemOrder(aOrder);

	while (nOrder--)
	{
		int nItem = aOrder[nOrder];

		if (IsItemVisible(nItem) && GetItemWidth(nItem))
			return nItem;
	}

	// all hidden
	return -1;
}

int CEnHeaderCtrl::GetNextVisibleItem(int nItem, BOOL bNext) const
{
	CIntArray aOrder;
	int nNumItems = GetItemOrder(aOrder);
	
	int nPos = GetItemPosition(nItem, aOrder);
	
	if (nPos != -1)
	{
		if (bNext)
		{
			while (++nPos < nNumItems)
			{
				nItem = OrderToIndex(nPos);

				if (IsItemVisible(nItem) && GetItemWidth(nItem))
					return nItem;
			}
		}
		else
		{
			while (nPos--)
			{
				nItem = OrderToIndex(nPos);
				
				if (IsItemVisible(nItem) && GetItemWidth(nItem))
					return nItem;
			}
		}
	}

	// outside range
	return -1;
}

int CEnHeaderCtrl::GetItemPosition(int nItem) const
{
	CIntArray aOrder;
	GetItemOrder(aOrder);

	return GetItemPosition(nItem, aOrder);
}

int CEnHeaderCtrl::GetItemPosition(int nItem, const CIntArray& aOrder)
{
	int i = aOrder.GetSize();

	while (i--)
	{
		if (aOrder[i] == nItem)
			return i;
	}

	// not found ?
	ASSERT(0);
	return -1;
}

BOOL CEnHeaderCtrl::SetItemSorted(int nItem, EHC_SORT nDir)
{
	if (!IsItemVisible(nItem))
		return FALSE;

	// set sorting only on this item
	ClearAllSorted();
	
	switch(nDir)
	{
	case EHCS_UP:
		ModifyItemFlags(nItem, EHCF_SORTUP, TRUE);
		break;

	case EHCS_NONE:
		ModifyItemFlags(nItem, EHCF_SORTDOWN, TRUE);
		break;

	case EHCS_DOWN:
		break;
	}

	return TRUE;
}

void CEnHeaderCtrl::ClearAllSorted()
{
	ModifyAllItemFlags(EHCF_SORTMASK, FALSE);
}

EHC_SORT CEnHeaderCtrl::GetItemSorted(int nItem) const
{
	if (ItemHasFlag(nItem, EHCF_SORTDOWN))
	{
		return EHCS_UP;
	}
	else if (ItemHasFlag(nItem, EHCF_SORTUP))
	{
		return EHCS_UP;
	}

	// else
	return EHCS_NONE;
}

int CEnHeaderCtrl::GetSortedItem() const
{
	int nItem = GetItemCount();
	
	while (nItem--)
	{
		if (GetItemFlags(nItem) & EHCF_SORTMASK)
			return nItem;
	}

	// else
	return -1;
}

BOOL CEnHeaderCtrl::DrawItemSortArrow(CDC* pDC, int nItem) const
{
	EHC_SORT nDir = GetItemSorted(nItem);
	ASSERT (nDir != EHCS_NONE);

	if (nDir != EHCS_NONE)
		return DrawItemSortArrow(pDC, nItem, (nDir == EHCS_UP));

	// else
	return FALSE;
}

BOOL CEnHeaderCtrl::DrawItemSortArrow(CDC* pDC, int nItem, BOOL bUp) const
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());

	CRect rArrow;
	
	if (!GetItemRect(nItem, rArrow))
		return FALSE;

	static CThemed th(this, _T("Header"));
	
	if (th.AreControlsThemed())
	{
		rArrow.left = rArrow.CenterPoint().x - 6;
		rArrow.right = rArrow.left + 12;
		rArrow.bottom = rArrow.top + 8;
		
		return th.DrawBackground(pDC, HP_HEADERSORTARROW, (bUp ? HSAS_SORTEDUP : HSAS_SORTEDDOWN), rArrow);
	}
	
	// else
	int nOffY = (bUp ? 5 : 3), nOffX = (rArrow.CenterPoint().x - 4);
	int nDir = (bUp ? -1 : 1);
	
	POINT ptArrow[3] = { { 0, 0 }, { 3, (int)nDir * 3 }, { 7, -(int)nDir } };
	
	// translate the arrow to the appropriate location
	int nPoint = 3;
	
	while (nPoint--)
	{
		ptArrow[nPoint].x += nOffX;
		ptArrow[nPoint].y += nOffY;
	}
	pDC->Polyline(ptArrow, 3);

	return TRUE;
}

BOOL CEnHeaderCtrl::ItemHasFlag(int nItem, DWORD dwFlag) const
{
	return ((GetItemFlags(nItem) & dwFlag) == dwFlag);
}

DWORD CEnHeaderCtrl::GetItemFlags(int nItem) const
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());

	if (nItem < m_aItemFlags.GetSize())
		return m_aItemFlags[nItem];

	// else 
	return 0;
}

BOOL CEnHeaderCtrl::ModifyItemFlags(int nItem, DWORD dwFlag, BOOL bAdd)
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());

	DWORD dwFlags = GetItemFlags(nItem), dwNewFlags = dwFlags;

	if (bAdd)
		dwNewFlags |= dwFlag;
	else
		dwNewFlags &= ~dwFlag;

	m_aItemFlags.SetAtGrow(nItem, dwNewFlags);

	return (dwNewFlags != dwFlags);
}

BOOL CEnHeaderCtrl::ModifyAllItemFlags(DWORD dwFlag, BOOL bAdd)
{
	BOOL bMod = FALSE;
	int nItem = GetItemCount();

	while (nItem--)
		bMod |= ModifyItemFlags(nItem, dwFlag, bAdd);

	return bMod;
}

void CEnHeaderCtrl::EnableTracking(BOOL bAllow) 
{ 
	// globally allows/disallows tracking
	m_bEnableTracking = bAllow; 
}

void CEnHeaderCtrl::EnableItemTracking(int nItem, BOOL bAllow)
{
	ModifyItemFlags(nItem, EHCF_NOTRACK, !bAllow);
}

BOOL CEnHeaderCtrl::IsItemTrackable(int nItem) const
{
	return (m_bEnableTracking &&
			!ItemHasFlag(nItem, EHCF_NOTRACK) && 
			!ItemHasFlag(nItem, EHCF_HIDDEN));
}

BOOL CEnHeaderCtrl::IsItemTracked(int nItem) const
{
	return (IsItemTrackable(nItem) && ItemHasFlag(nItem, EHCF_TRACKED));
}

void CEnHeaderCtrl::EnableItemDragging(int nItem, BOOL bAllow)
{
	ASSERT(GetStyle() & HDS_DRAGDROP);

	ModifyItemFlags(nItem, EHCF_NODRAG, !bAllow);
}

BOOL CEnHeaderCtrl::IsItemDraggable(int nItem) const
{
	ASSERT(GetStyle() & HDS_DRAGDROP);
	
	return ((GetStyle() & HDS_DRAGDROP) && !ItemHasFlag(nItem, EHCF_NODRAG));
}

int CEnHeaderCtrl::AppendItem(int nWidth, LPCTSTR szText, int nFormat, UINT uIDBitmap)
{
	return InsertItem(GetItemCount(), nWidth, szText, nFormat, uIDBitmap);
}

int CEnHeaderCtrl::InsertItem(int nItem, int nWidth, LPCTSTR szText, int nFormat, UINT uIDBitmap)
{
	HD_ITEM hdi = { 0 };

	hdi.mask = HDI_WIDTH | HDI_FORMAT | HDI_TEXT | (uIDBitmap ? HDI_BITMAP : 0);
	hdi.fmt = nFormat;
	hdi.cxy = nWidth;
	hdi.pszText = (LPTSTR)szText;

	// bitmap
	if (uIDBitmap)
	{
		CBitmap bm;
		bm.LoadBitmap(uIDBitmap);
		hdi.hbm = (HBITMAP)bm.Detach();
	}

	return CHeaderCtrl::InsertItem(nItem, &hdi);
}

BOOL CEnHeaderCtrl::SetItem(int nItem, HDITEM* pHeaderItem, BOOL bRecalcTooltipRect)
{
	if (CHeaderCtrl::SetItem(nItem, pHeaderItem))
	{
		if (bRecalcTooltipRect && (pHeaderItem->mask & HDI_WIDTH))
			RecalcItemTooltipRects();
		
		return TRUE;
	}
	
	return FALSE;
}

BOOL CEnHeaderCtrl::SetItem(int nItem, HDITEM* pHeaderItem)
{
	return SetItem(nItem, pHeaderItem, TRUE);
}

BOOL CEnHeaderCtrl::SetItem(int nItem, int nWidth, LPCTSTR szText, DWORD dwData)
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());
	HD_ITEM hdi = { 0 };

	hdi.mask = HDI_WIDTH | HDI_LPARAM | HDI_TEXT;
	hdi.cxy = nWidth;
	hdi.pszText = (LPTSTR)szText;
	hdi.lParam = dwData;

	return SetItem(nItem, &hdi);
}

BOOL CEnHeaderCtrl::SetItemWidth(int nItem, int nWidth)
{
	return SetItemWidth(nItem, nWidth, TRUE);
}

BOOL CEnHeaderCtrl::SetItemWidth(int nItem, int nWidth, BOOL bRecalcTooltipRect)
{
	if (nWidth == GetItemWidth(nItem))
		return TRUE;

	HD_ITEM hdi = { 0 };

	hdi.mask = HDI_WIDTH;
	hdi.cxy = nWidth;

	if (SetItem(nItem, &hdi, bRecalcTooltipRect))
	{
		SetTrackedItem(nItem, FALSE);
		return TRUE;
	}

	// else
	return FALSE;
}

BOOL CEnHeaderCtrl::SetTrackedItem(int nItem, BOOL bTracked)
{
	return ModifyItemFlags(nItem, EHCF_TRACKED, bTracked);
}

#ifdef _DEBUG
void CEnHeaderCtrl::TraceVisibleItemWidths(LPCTSTR szKey) const
{
	TRACE(_T("CEnHeaderCtrl::TraceVisibleItemWidths(%s "), szKey);

	CIntArray aItems;
	int nNumItem = GetItemOrder(aItems);

	if (nNumItem)
	{
		for (int nItem = 0; nItem < nNumItem; nItem++)
		{
			if (IsItemVisible(aItems[nItem]))
				TRACE(_T("%d "), GetItemWidth(aItems[nItem]));
		}
	}
	else
	{
		TRACE(_T("empty"));
	}

	TRACE(_T(")\n"));
}

void CEnHeaderCtrl::TraceVisibleToolTipRects(LPCTSTR szKey)
{
	TRACE(_T("CEnHeaderCtrl::TraceVisibleItemWidths(%s "), szKey);
	
	CIntArray aItems;
	int nNumItem = GetItemOrder(aItems);
	
	if (nNumItem)
	{
		for (int nItem = 0; nItem < nNumItem; nItem++)
		{
			if (IsItemVisible(aItems[nItem]))
			{
				CToolInfo ti;
				CRect rItem;

				if (m_tooltips.GetToolInfo(ti, this, (nItem + 1)) && GetItemRect(nItem, rItem))
				{
					TRACE(_T("\n   %s: item (%d, %d, %d, %d), tip (%d, %d, %d, %d)"), ti.szText,
							rItem.left, rItem.top, rItem.right, rItem.bottom,
							ti.rect.left, ti.rect.top, ti.rect.right, ti.rect.bottom);
				}
			}
		}
	}
	else
	{
		TRACE(_T("empty"));
	}
	
	TRACE(_T(")\n"));
}
#endif

int CEnHeaderCtrl::GetItemWidths(CIntArray& aWidths) const
{
	int nItem = GetItemCount();
	aWidths.SetSize(nItem);

	while (nItem--)
		aWidths[nItem] = GetItemWidth(nItem);

	return aWidths.GetSize();
}

BOOL CEnHeaderCtrl::SetItemWidths(const CIntArray& aWidths)
{
	int nItem = GetItemCount();
	ASSERT(nItem == aWidths.GetSize());

	if (nItem != aWidths.GetSize())
		return FALSE;

	while (nItem--)
	{
		if (!SetItemWidth(nItem, aWidths[nItem], FALSE))
			return FALSE;
	}

	RecalcItemTooltipRects();
	return TRUE;
}

int CEnHeaderCtrl::GetTrackedItems(CIntArray& aTracked) const
{
	int nItem = GetItemCount();
	aTracked.SetSize(nItem);
	
	while (nItem--)
		aTracked[nItem] = IsItemTracked(nItem);
	
	return aTracked.GetSize();
}

BOOL CEnHeaderCtrl::SetTrackedItems(const CIntArray& aTracked)
{
	int nItem = GetItemCount();
	ASSERT(nItem == aTracked.GetSize());
	
	if (nItem != aTracked.GetSize())
		return FALSE;
	
	while (nItem--)
		SetTrackedItem(nItem, aTracked[nItem]);
	
	return TRUE;
}

void CEnHeaderCtrl::ClearAllTracked()
{
	ModifyAllItemFlags(EHCF_TRACKED, FALSE);
}

int CEnHeaderCtrl::GetItemOrder(CIntArray& aOrder) const
{
	int nNumItems = GetItemCount();
	aOrder.SetSize(nNumItems);
	
	CEnHeaderCtrl* pThis = const_cast<CEnHeaderCtrl*>(this);
	
	if (pThis->GetOrderArray(aOrder.GetData(), nNumItems))
		return nNumItems;

	// else
	aOrder.RemoveAll();
	return 0;
}

BOOL CEnHeaderCtrl::SetItemOrder(const CIntArray& aOrder)
{
	return SetOrderArray(aOrder.GetSize(), (LPINT)aOrder.GetData());
}

int CEnHeaderCtrl::GetItemWidth(int nItem) const
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());
	HD_ITEM hdi = { 0 };

	hdi.mask = HDI_WIDTH;
	GetItem(nItem, &hdi);

	return hdi.cxy;
}

int CEnHeaderCtrl::FindItem(DWORD dwData) const
{
	int nItem = GetItemCount();

	while (nItem--)
	{
		if (GetItemData(nItem) == dwData)
			return nItem;
	}

	// not found
	return -1;
}

BOOL CEnHeaderCtrl::SetItemData(int nItem, DWORD dwData)
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());
	HD_ITEM hdi = { 0 };

	hdi.mask = HDI_LPARAM;
	hdi.lParam = dwData;

	return SetItem(nItem, &hdi);
}

DWORD CEnHeaderCtrl::GetItemData(int nItem) const
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());
	HD_ITEM hdi = { 0 };

	hdi.mask = HDI_LPARAM;
	GetItem(nItem, &hdi);

	return hdi.lParam;
}

BOOL CEnHeaderCtrl::SetItemText(int nItem, LPCTSTR szText)
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());
	HD_ITEM hdi = { 0 };

	hdi.mask = HDI_TEXT;
	hdi.pszText = (LPTSTR)szText;

	return SetItem(nItem, &hdi);
}

CString CEnHeaderCtrl::GetItemText(int nItem) const
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());
	HD_ITEM hdi = { 0 };

	TCHAR szText[128];

	hdi.mask = HDI_TEXT;
	hdi.pszText = szText;
	hdi.cchTextMax = 127;

	GetItem(nItem, &hdi);

	return hdi.pszText;
}

int CEnHeaderCtrl::GetItemTextWidth(int nItem, CDC* pDC) const
{
	return pDC->GetTextExtent(GetItemText(nItem)).cx;
}

void CEnHeaderCtrl::OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/) 
{
	// do nothing	
}

BOOL CEnHeaderCtrl::PreTranslateMessage(MSG* pMsg)
{
	if (m_tooltips.GetSafeHwnd())
		m_tooltips.RelayEvent(pMsg);

	return CHeaderCtrl::PreTranslateMessage(pMsg);
}

BOOL CEnHeaderCtrl::EnableToolTips(BOOL bEnable)
{
	if (bEnable)
		return InitializeTooltips();
	
	// else 
	if (m_tooltips.GetSafeHwnd())
		m_tooltips.DestroyWindow();

	return TRUE;
}

BOOL CEnHeaderCtrl::InitializeTooltips()
{
	return (m_tooltips.GetSafeHwnd() || m_tooltips.Create(this, (WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP)));
}

void CEnHeaderCtrl::RecalcItemTooltipRects()
{
	if (m_tooltips.GetSafeHwnd())
	{
		int nItem = GetItemCount();

		while (nItem--)
		{
// 			CToolInfo ti;
// 			
// 			if (m_tooltips.GetToolInfo(ti, this, (nItem + 1)))
			{
				CRect rItem;
				GetItemRect(nItem, rItem);
				
				m_tooltips.SetToolRect(this, (nItem + 1), rItem);
			}
		}
	}
}

LRESULT CEnHeaderCtrl::OnRecalcItemTooltipRects(WPARAM wp, LPARAM lp)
{
	RecalcItemTooltipRects();

	return 0L;
}

BOOL CEnHeaderCtrl::SetItemToolTip(int nItem, LPCTSTR szTip)
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());
	ASSERT (szTip);

	if (InitializeTooltips())
	{
		CToolInfo ti;

		if (!m_tooltips.GetToolInfo(ti, this, (nItem + 1)))
		{
			CRect rItem;
			GetItemRect(nItem, rItem);

			m_tooltips.AddTool(this, szTip, rItem, (nItem + 1));
		}
		else
		{
			lstrcpyn(ti.szText, szTip, ((sizeof(ti.szText) / sizeof(TCHAR)) - 1)); 
			m_tooltips.SetToolInfo(&ti);
		}

		return TRUE;
	}

	return FALSE;
}

CString CEnHeaderCtrl::GetItemToolTip(int nItem)
{
	ASSERT (nItem >= 0 && nItem < GetItemCount());

	if (m_tooltips.GetSafeHwnd())
	{
		CToolInfo ti;
		
		if (m_tooltips.GetToolInfo(ti, this, nItem))
			return ti.szText;
	}

	// else
	return _T("");
}
