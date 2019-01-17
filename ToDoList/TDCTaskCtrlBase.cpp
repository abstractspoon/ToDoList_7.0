// TDCTreeListCtrl.cpp: implementation of the CTDCTaskCtrlBase class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TDCTaskCtrlBase.h"
#include "todoctrldata.h"
#include "tdcstatic.h"
#include "tdcmsg.h"
#include "tdccustomattributehelper.h"
#include "tdcimagelist.h"
#include "resource.h"

/////////////////////////////////////////////////////////////////////////////

#include "..\shared\graphicsmisc.h"
#include "..\shared\autoflag.h"
#include "..\shared\holdredraw.h"
#include "..\shared\timehelper.h"
#include "..\shared\misc.h"
#include "..\shared\filemisc.h"
#include "..\shared\colordef.h"
#include "..\shared\preferences.h"
#include "..\shared\themed.h"
#include "..\shared\wndprompt.h"
#include "..\shared\osversion.h"
#include "..\shared\webmisc.h"
#include "..\shared\enbitmap.h"

/////////////////////////////////////////////////////////////////////////////

#include <MATH.H>

/////////////////////////////////////////////////////////////////////////////

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

const int LV_COLPADDING			= 3;
const int HD_COLPADDING			= 6;
const int MIN_RESIZE_WIDTH		= 17; 
const int MIN_COL_WIDTH			= 6;
const int MIN_TASKS_WIDTH		= 200;

const COLORREF COMMENTSCOLOR	= RGB(98, 98, 98);
const COLORREF ALTCOMMENTSCOLOR = RGB(164, 164, 164);

const UINT TIMER_BOUNDINGSEL	= 100;

//////////////////////////////////////////////////////////////////////

const UINT WM_NOTIFYPARENTSELECTIONUPDATE = (WM_APP + 1);

//////////////////////////////////////////////////////////////////////

enum
{
	IDC_TASKTREE = 100,		
	IDC_TASKTREECOLUMNS,		
	IDC_TASKTREEHEADER,		
};

//////////////////////////////////////////////////////////////////////

CMap<TDC_COLUMN, TDC_COLUMN, TDCCOLUMN*, TDCCOLUMN*&>
			CTDCTaskCtrlBase::s_mapColumns;
short		CTDCTaskCtrlBase::s_nExtendedSelection = HOTKEYF_CONTROL | HOTKEYF_SHIFT;
double		CTDCTaskCtrlBase::s_dRecentModPeriod = 0.0;												

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

IMPLEMENT_DYNAMIC(CTDCTaskCtrlBase, CWnd)

//////////////////////////////////////////////////////////////////////

CTDCTaskCtrlBase::CTDCTaskCtrlBase(BOOL bSyncSelection,
								   const CTDCImageList& ilIcons,
								   const CToDoCtrlData& data, 
								   const CToDoCtrlFind& find,
								   const CWordArray& aStyles,
								   const CTDCColumnIDArray& aVisibleCols,
								   const CTDCCustomAttribDefinitionArray& aCustAttribDefs) 
	: 
	CTreeListSyncer(TLSF_SYNCFOCUS | TLSF_BORDER | TLSF_SYNCDATA | TLSF_SPLITTER | (bSyncSelection ? TLSF_SYNCSELECTION : 0)),
	m_data(data),
	m_find(find),
	m_aStyles(aStyles),
	m_ilTaskIcons(ilIcons),
	m_aVisibleCols(aVisibleCols),
	m_aCustomAttribDefs(aCustAttribDefs),
	m_crDone(CLR_NONE),
	m_crDue(CLR_NONE), 
	m_crDueToday(CLR_NONE),
	m_crFlagged(CLR_NONE),
	m_crStarted(CLR_NONE), 
	m_crStartedToday(CLR_NONE),
	m_crReference(CLR_NONE),
	m_crAltLine(CLR_NONE),
	m_crGridLine(CLR_NONE),
	m_nSortColID(TDCC_NONE),
	m_nSortDir(TDC_SORTNONE),
	m_dwTimeTrackTaskID(0), 
	m_dwEditTitleTaskID(0),
	m_dwNextUniqueTaskID(100),
	m_nMaxInfotipCommentsLength(-1),
	m_bSortingColumns(FALSE),
	m_nColorByAttrib(TDCA_NONE),
	m_bBoundSelecting(FALSE),
	m_nDefTimeEstUnits(THU_HOURS), 
	m_nDefTimeSpentUnits(THU_HOURS),
	m_imageIcons(16, 16)
{
	// build one time column map
	if (s_mapColumns.IsEmpty())
	{
		// add all columns
		int nCol = NUM_COLUMNS;

		while (nCol--)
		{
			TDCCOLUMN& tdcc = COLUMNS[nCol];
			s_mapColumns[tdcc.nColID] = &tdcc;
		}
	}
}

CTDCTaskCtrlBase::~CTDCTaskCtrlBase()
{
	Release();
}

///////////////////////////////////////////////////////////////////////////

BEGIN_MESSAGE_MAP(CTDCTaskCtrlBase, CWnd)
//{{AFX_MSG_MAP(CTDCTaskCtrlBase)
//}}AFX_MSG_MAP
ON_WM_DESTROY()
ON_WM_SIZE()
ON_WM_ERASEBKGND()
ON_WM_CREATE()
ON_MESSAGE(WM_SETREDRAW, OnSetRedraw)
ON_WM_SETCURSOR()
ON_WM_TIMER()
END_MESSAGE_MAP()

///////////////////////////////////////////////////////////////////////////

BOOL CTDCTaskCtrlBase::Create(CWnd* pParentWnd, const CRect& rect, UINT nID, BOOL bVisible)
{
	DWORD dwStyle = (WS_CHILD | (bVisible ? WS_VISIBLE : 0) | WS_TABSTOP);
	
	// create ourselves
	return CWnd::CreateEx(WS_EX_CONTROLPARENT, NULL, NULL, dwStyle, rect, pParentWnd, nID);
}

int CTDCTaskCtrlBase::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;

	BOOL bVisible = (lpCreateStruct->style & WS_VISIBLE);
	CRect rect(0, 0, lpCreateStruct->cx, lpCreateStruct->cy);
	
	if (!CreateTasksWnd(this, rect, bVisible))
		return -1;

	DWORD dwStyle = (WS_CHILD | (bVisible ? WS_VISIBLE : 0));

	// Tasks Header ---------------------------------------------------------------------
	if (!m_hdrTasks.Create((dwStyle | HDS_BUTTONS), rect, this, IDC_TASKTREEHEADER))
	{
		return FALSE;
	}

	// Column List ---------------------------------------------------------------------
	if (!m_lcColumns.Create((dwStyle | WS_TABSTOP),	rect, this, IDC_TASKTREECOLUMNS))
	{
		return FALSE;
	}
	
	// extended styles
	ListView_SetExtendedListViewStyleEx(m_lcColumns, LVS_EX_HEADERDRAGDROP, LVS_EX_HEADERDRAGDROP);
	
	// subclass the tree and list
	if (HasStyle(TDCS_RIGHTSIDECOLUMNS))
	{
		if (!Sync(Tasks(), m_lcColumns, TLSL_RIGHTDATA_IS_LEFTITEM, m_hdrTasks))
		{
			return FALSE;
		}
	}
	else // left side
	{
		if (!Sync(m_lcColumns, Tasks(), TLSL_LEFTDATA_IS_RIGHTITEM, m_hdrTasks))
		{
			return FALSE;
		}
	}
		
	// Column Header ---------------------------------------------------------------------
	if (!m_hdrColumns.SubclassWindow(ListView_GetHeader(m_lcColumns)))
	{
		return FALSE;
	}
	m_hdrColumns.EnableToolTips();
	
	// set header font and calc char width
	CFont* pFont = m_hdrColumns.GetFont();
	m_hdrTasks.SetFont(pFont);

	CClientDC dc(&m_hdrTasks);
	m_fAveHeaderCharWidth = GraphicsMisc::GetAverageCharWidth(&dc, pFont);

	VERIFY(InitCheckboxImageList());

	BuildColumns();
	RecalcColumnWidths();
	OnColumnVisibilityChange();
	PostResize();
	
	return 0;
}

void CTDCTaskCtrlBase::UpdateSelectedTaskPath()
{
	CEnString sHeader(IDS_TDC_COLUMN_TASK);
	
	// add the item path to the header
	if (HasStyle(TDCS_SHOWPATHINHEADER) && m_hdrTasks.GetItemCount())
	{
		if (GetSelectedCount() == 1)
		{
			CRect rHeader;
			::GetClientRect(m_hdrTasks, rHeader);
			
			int nColWidthInChars = (int)(rHeader.Width() / m_fAveHeaderCharWidth);
			CString sPath = m_data.GetTaskPath(GetSelectedTaskID(), nColWidthInChars);
			
			if (!sPath.IsEmpty())
				sHeader.Format(_T("%s [%s]"), CEnString(IDS_TDC_COLUMN_TASK), sPath);
		}
	}
	
	m_hdrTasks.SetItemText(0, sHeader);
	m_hdrTasks.Invalidate(FALSE);
}

void CTDCTaskCtrlBase::SetTimeTrackTaskID(DWORD dwTaskID)
{
	if (m_dwTimeTrackTaskID != dwTaskID)
	{
		RedrawColumn(TDCC_TRACKTIME);
		RedrawColumn(TDCC_TIMEEST);

		m_dwTimeTrackTaskID = dwTaskID;
		
		// resort if appropriate
		if (m_sort.IsSortingBy(TDCC_TRACKTIME, FALSE))
			Sort(TDCC_TRACKTIME, FALSE);
	}
}

void CTDCTaskCtrlBase::SetEditTitleTaskID(DWORD dwTaskID)
{
	if (m_dwEditTitleTaskID != dwTaskID)
	{
		m_dwEditTitleTaskID = dwTaskID;
	}
}

void CTDCTaskCtrlBase::SetNextUniqueTaskID(DWORD dwTaskID)
{
	if (m_dwNextUniqueTaskID != dwTaskID)
	{
		m_dwNextUniqueTaskID = dwTaskID;
	}
}

void CTDCTaskCtrlBase::OnDestroy() 
{
	Release();
	
	CWnd::OnDestroy();
}

void CTDCTaskCtrlBase::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);
	
	if (cx && cy)
	{
		CRect rect(0, 0, cx, cy);
		CTreeListSyncer::Resize(rect);

		switch (GetHidden())
		{
		case TLSH_LEFT:
			OnNotifySplitterChange(0);
			break;

		case TLSH_RIGHT:
			OnNotifySplitterChange(cx);
			break;
		}
	}
}

LRESULT CTDCTaskCtrlBase::OnSetRedraw(WPARAM wp, LPARAM /*lp*/)
{
	::SendMessage(Tasks(), WM_SETREDRAW, wp, 0);
	m_lcColumns.SetRedraw(wp);

	return 0L;
}

BOOL CTDCTaskCtrlBase::OnEraseBkgnd(CDC* pDC)
{
	return CTreeListSyncer::HandleEraseBkgnd(pDC);
}

BOOL CTDCTaskCtrlBase::IsListItemSelected(HWND hwnd, int nItem) const
{
	return (ListView_GetItemState(hwnd, nItem, LVIS_SELECTED) & LVIS_SELECTED);
}

void CTDCTaskCtrlBase::OnStylesUpdated()
{
	SetTasksImageList(m_ilCheckboxes, TRUE, (!IsColumnShowing(TDCC_DONE) && HasStyle(TDCS_ALLOWTREEITEMCHECKBOX)));

	RecalcColumnWidths();
	UpdateHeaderSorting();
	PostResize();

	if (IsVisible())
		InvalidateAll();
}

void CTDCTaskCtrlBase::OnStyleUpdated(TDC_STYLE nStyle, BOOL bOn, BOOL bDoUpdate)
{
	switch (nStyle)
	{
	case TDCS_NODUEDATEISDUETODAY:
	case TDCS_SHOWDATESINISO:
	case TDCS_USEEARLIESTDUEDATE:
	case TDCS_USELATESTDUEDATE:
	case TDCS_USEEARLIESTSTARTDATE:
	case TDCS_USELATESTSTARTDATE:
	case TDCS_SHOWCOMMENTSINLIST:
	case TDCS_SHOWFIRSTCOMMENTLINEINLIST:
	case TDCS_STRIKETHOUGHDONETASKS:
	case TDCS_TASKCOLORISBACKGROUND:
	case TDCS_CALCREMAININGTIMEBYDUEDATE:
	case TDCS_CALCREMAININGTIMEBYSPENT:
	case TDCS_CALCREMAININGTIMEBYPERCENT:
	case TDCS_COLORTEXTBYATTRIBUTE:
		if (bDoUpdate)
			InvalidateAll();
		break;

	case TDCS_SORTDONETASKSATBOTTOM:
 		Resort();
		break;
		
	case TDCS_DUEHAVEHIGHESTPRIORITY:
	case TDCS_DONEHAVELOWESTPRIORITY:
		if (IsSortingBy(TDCC_PRIORITY))
			Resort();
		break;
		
	case TDCS_DONEHAVELOWESTRISK:
		if (IsSortingBy(TDCC_RISK))
			Resort();
		break;
		
	case TDCS_RIGHTSIDECOLUMNS:
		if (bOn != IsShowingColumnsOnRight())
			SwapSides();
		break;
		
	case TDCS_DISPLAYHMSTIMEFORMAT:
	case TDCS_TREATSUBCOMPLETEDASDONE:
		if (bDoUpdate)
			RecalcColumnWidths();
		break;
		
	case TDCS_USEHIGHESTPRIORITY:
	case TDCS_INCLUDEDONEINPRIORITYCALC:
	case TDCS_HIDEPRIORITYNUMBER:
		if (bDoUpdate && IsColumnShowing(TDCC_PRIORITY))
			InvalidateAll();
		break;
		
	case TDCS_USEHIGHESTRISK:
	case TDCS_INCLUDEDONEINRISKCALC:
		if (bDoUpdate && IsColumnShowing(TDCC_RISK))
			InvalidateAll();
		break;
		
	case TDCS_SHOWNONFILEREFSASTEXT:
		if (bDoUpdate && IsColumnShowing(TDCC_FILEREF))
			RecalcColumnWidths();
		break;
		
	case TDCS_USEPERCENTDONEINTIMEEST:
		if (bDoUpdate && IsColumnShowing(TDCC_TIMEEST))
			RecalcColumnWidths();
		break;
		
	case TDCS_HIDEZEROTIMECOST:
		if (bDoUpdate && 
			(IsColumnShowing(TDCC_TIMEEST) || 
			IsColumnShowing(TDCC_TIMESPENT) || 
			IsColumnShowing(TDCC_COST)))
		{
			RecalcColumnWidths();
		}
		break;
		
	case TDCS_ROUNDTIMEFRACTIONS:
		if (bDoUpdate && 
			(IsColumnShowing(TDCC_TIMEEST) || 
			IsColumnShowing(TDCC_TIMESPENT)))
		{
			RecalcColumnWidths();
		}
		break;
		
	case TDCS_HIDEPERCENTFORDONETASKS:
	case TDCS_INCLUDEDONEINAVERAGECALC:
	case TDCS_WEIGHTPERCENTCALCBYNUMSUB:
	case TDCS_SHOWPERCENTASPROGRESSBAR:
	case TDCS_HIDEZEROPERCENTDONE:
		if (bDoUpdate && IsColumnShowing(TDCC_PERCENT))
			InvalidateAll();
		break;
		
	case TDCS_AVERAGEPERCENTSUBCOMPLETION:
	case TDCS_AUTOCALCPERCENTDONE:
		if (bDoUpdate)
		{
			if (IsColumnShowing(TDCC_PERCENT))
				RecalcColumnWidths();
			else
				InvalidateAll();
		}
		break;
		
	case TDCS_HIDESTARTDUEFORDONETASKS:
		if (bDoUpdate && 
			(IsColumnShowing(TDCC_STARTDATE) || 
			IsColumnShowing(TDCC_DUEDATE)))
		{
			RecalcColumnWidths();
		}
		break;
		
	case TDCS_SHOWWEEKDAYINDATES:
		if (bDoUpdate && 
			(IsColumnShowing(TDCC_STARTDATE) || 
			IsColumnShowing(TDCC_LASTMOD) ||
			IsColumnShowing(TDCC_DUEDATE) || 
			IsColumnShowing(TDCC_DONEDATE)))
		{
			RecalcColumnWidths();
		}
		break;
		
	case TDCS_SHOWPATHINHEADER:
		UpdateSelectedTaskPath();
		break;
		
	case TDCS_HIDEPANESPLITBAR:
		CTreeListSyncer::SetSplitBarWidth(bOn ? 0 : 10);
		break;

	// all else not handled
	}
}

BOOL CTDCTaskCtrlBase::InvalidateColumnItem(int nItem, BOOL bUpdate)
{
	return InvalidateItem(m_lcColumns, nItem, bUpdate);
}

BOOL CTDCTaskCtrlBase::InvalidateColumnSelection(BOOL bUpdate)
{
	return InvalidateSelection(m_lcColumns, bUpdate);
}

BOOL CTDCTaskCtrlBase::InvalidateSelection(CListCtrl& lc, BOOL bUpdate)
{
	BOOL bInvalidated = FALSE;
	POSITION pos = lc.GetFirstSelectedItemPosition();

	while (pos)
	{
		int nItem = lc.GetNextSelectedItem(pos);
		bInvalidated |= InvalidateItem(lc, nItem, FALSE); // don't update until the end
	}
	
	if (bUpdate && bInvalidated)
		lc.UpdateWindow();

	return bInvalidated;
}

BOOL CTDCTaskCtrlBase::InvalidateItem(CListCtrl& lc, int nItem, BOOL bUpdate)
{
	ASSERT(nItem != -1);
	
	if (nItem == -1)
		return FALSE;

	BOOL bInvalidated = FALSE;
	CRect rItem;

	if (lc.GetItemRect(nItem, rItem, LVIR_BOUNDS))
	{
		lc.InvalidateRect(rItem);
		bInvalidated = TRUE;
	}
	
	if (bUpdate && bInvalidated)
		lc.UpdateWindow();

	return bInvalidated;
}

void CTDCTaskCtrlBase::InvalidateAll(BOOL bErase, BOOL bUpdate) 
{ 
	CTreeListSyncer::InvalidateAll(bErase, bUpdate); 
}

BOOL CTDCTaskCtrlBase::IsShowingColumnsOnRight() const
{
	return (Right() == m_lcColumns);
}

void CTDCTaskCtrlBase::OnUndoRedo(BOOL /*bUndo*/)
{
	// resync scroll pos
	PostResync(m_lcColumns, FALSE);
}

void CTDCTaskCtrlBase::OnColumnVisibilityChange()
{
	// everything depends on the visibility map being up to date
	RebuildColumnVisibilityMap();

	CHoldRedraw hr(m_lcColumns);
	
	int nNumCols = m_hdrColumns.GetItemCount();
	BOOL nNumShowing = 0;
	
	for (int nItem = 1; nItem < nNumCols; nItem++)
	{		
		TDC_COLUMN nColID = (TDC_COLUMN)m_hdrColumns.GetItemData(nItem);

		if (IsColumnShowing(nColID))
		{
			m_hdrColumns.ShowItem(nItem, TRUE);
			nNumShowing++;
		}
		else
		{
			m_hdrColumns.ShowItem(nItem, FALSE);
		}
	}

	if (nNumShowing == 0)
		SetHidden(IsLeft(m_lcColumns) ? TLSH_LEFT : TLSH_RIGHT);
	else
		SetHidden(TLSH_NONE);

	OnImageListChange();
	RecalcColumnWidths(FALSE); // FALSE -> standard columns
}

void CTDCTaskCtrlBase::OnImageListChange()
{
	SetTasksImageList(m_ilTaskIcons, FALSE, !IsColumnShowing(TDCC_ICON));
	SetTasksImageList(m_ilCheckboxes, TRUE, (!IsColumnShowing(TDCC_DONE) && HasStyle(TDCS_ALLOWTREEITEMCHECKBOX)));

	if (IsVisible())
		::InvalidateRect(Tasks(), NULL, FALSE);
}

BOOL CTDCTaskCtrlBase::IsVisible() const
{
	HWND hwnd = GetSafeHwnd();

	return (hwnd && ::IsWindowVisible(::GetParent(hwnd)) && ::IsWindowVisible(hwnd));
}

void CTDCTaskCtrlBase::OnCustomAttributeChange()
{
	// everything depends on the visibility map up to date
	RebuildColumnVisibilityMap();

	for (int nAttrib = 0; nAttrib < m_aCustomAttribDefs.GetSize(); nAttrib++)
	{
		const TDCCUSTOMATTRIBUTEDEFINITION& attribDef = m_aCustomAttribDefs.GetData()[nAttrib];
		
		int nItem = m_hdrColumns.FindItem(attribDef.GetColumnID());
		ASSERT(nItem != -1);

		m_hdrColumns.EnableItemTracking(nItem, attribDef.bEnabled);
		m_hdrColumns.ShowItem(nItem, attribDef.bEnabled);

		if (attribDef.bEnabled)
			m_hdrColumns.SetItemText(nItem, attribDef.GetColumnTitle());
		else
			m_hdrColumns.SetItemText(nItem, _T(""));
	}

	RecalcColumnWidths(TRUE); // TRUE -> Custom columns
}

void CTDCTaskCtrlBase::RebuildColumnVisibilityMap()
{
	m_mapColVisibility.RemoveAll();

	// title always visible
	m_mapColVisibility[TDCC_CLIENT] = TRUE;

	// default columns
	for (int nCol = 0; nCol < m_aVisibleCols.GetSize(); nCol++)
	{
		m_mapColVisibility[m_aVisibleCols[nCol]] = TRUE;
	}
	
	// custom columns
	for (int nAttrib = 0; nAttrib < m_aCustomAttribDefs.GetSize(); nAttrib++)
	{
		const TDCCUSTOMATTRIBUTEDEFINITION& attribDef = m_aCustomAttribDefs.GetData()[nAttrib];
		
		if (attribDef.bEnabled)
			m_mapColVisibility[attribDef.GetColumnID()] = TRUE;
	}
}

BOOL CTDCTaskCtrlBase::IsColumnShowing(TDC_COLUMN nColID) const
{
	BOOL bShowing = FALSE;

	return (m_mapColVisibility.Lookup(nColID, bShowing) && bShowing);
}

BOOL CTDCTaskCtrlBase::SetColumnOrder(const CDWordArray& aColumns)
{
	CIntArray aOrder;
	aOrder.SetSize(aColumns.GetSize() + 1);

	// hidden column is always first
	aOrder[0] = 0;
	
	// convert columns IDs to indices
	int nNumCols = aColumns.GetSize();
	
	for (int nCol = 0; nCol < nNumCols; nCol++)
	{		
		int nItem = m_hdrColumns.FindItem((DWORD)aColumns[nCol]);
		ASSERT(nItem != -1);

		aOrder[nCol + 1] = nItem;
	}
	
	return m_lcColumns.SetColumnOrderArray(aOrder.GetSize(), aOrder.GetData());
}

BOOL CTDCTaskCtrlBase::GetColumnOrder(CDWordArray& aColumnIDs) const
{
	CIntArray aOrder;
	int nNumCols = m_hdrColumns.GetItemOrder(aOrder);

	if (nNumCols)
	{
		// ignore first column because that's our dummy column
		aColumnIDs.SetSize(nNumCols - 1); 

		for (int nItem = 1; nItem < nNumCols; nItem++)
		{		
			TDC_COLUMN nColID = (TDC_COLUMN)m_hdrColumns.GetItemData(aOrder[nItem]);
			ASSERT(nColID != TDCC_NONE);
			
			aColumnIDs[nItem - 1] = nColID;
		}
	
		return TRUE;
	}
	
	return FALSE;
}

void CTDCTaskCtrlBase::SetColumnWidths(const CDWordArray& aWidths)
{
	int nNumCols = aWidths.GetSize();
	
	// omit first column because that's our dummy column
	for (int nCol = 0; nCol < nNumCols; nCol++)
		m_hdrColumns.SetItemWidth(nCol + 1, aWidths[nCol]);
}

void CTDCTaskCtrlBase::GetColumnWidths(CDWordArray& aWidths) const
{
	CIntArray aIntWidths;
	int nNumCols = m_hdrColumns.GetItemWidths(aIntWidths);

	// omit first column because that's our dummy column
	aWidths.SetSize(nNumCols - 1);
	
	for (int nCol = 1; nCol < nNumCols; nCol++)
		aWidths[nCol - 1] = aIntWidths[nCol];
}

void CTDCTaskCtrlBase::SetTrackedColumns(const CDWordArray& aTracked)
{
	int nNumCols = aTracked.GetSize();
	
	// omit first column because that's our dummy column
	for (int nCol = 0; nCol < nNumCols; nCol++)
		m_hdrColumns.SetTrackedItem(nCol + 1, (int)aTracked[nCol]);
}

void CTDCTaskCtrlBase::GetTrackedColumns(CDWordArray& aTracked) const
{
	CIntArray aIntTracked;
	int nNumCols = m_hdrColumns.GetTrackedItems(aIntTracked);

	// omit first column because that's our dummy column
	aTracked.SetSize(nNumCols - 1);
	
	for (int nCol = 1; nCol < nNumCols; nCol++)
		aTracked[nCol - 1] = aIntTracked[nCol];
}

BOOL CTDCTaskCtrlBase::BuildColumns()
{
	if (m_hdrColumns.GetItemCount())
		return FALSE;

	// primary header
	const TDCCOLUMN* pClient = GetColumn(TDCC_CLIENT);
	ASSERT(pClient);
		
	// add empty column as placeholder so we can easily replace the 
	// other columns without losing all our items too
	m_lcColumns.InsertColumn(0, _T(""));
	m_hdrColumns.ShowItem(0, FALSE);
	
	// add all columns in two stages because m_lcColumns 
	// doesn't immediately update the header control
	int nCol = 0;
	
	for (nCol = 0; nCol < NUM_COLUMNS; nCol++)
	{
		TDCCOLUMN& tdcc = COLUMNS[nCol];
		
		if (tdcc.nColID != TDCC_CLIENT)
		{
			m_lcColumns.InsertColumn((nCol + 1), CEnString(tdcc.nIDName), tdcc.GetColumnHeaderAlignment(), 100);
		}
		else
		{
			ASSERT(!m_hdrTasks.GetItemCount());

			m_hdrTasks.AppendItem(150, CEnString(pClient->nIDName), HDF_LEFT);
			m_hdrTasks.SetItemData(0, TDCC_CLIENT);
			m_hdrTasks.EnableItemTracking(0, FALSE); // always
		}
	}
	
	// Add column IDs to header
	for (nCol = 0; nCol < NUM_COLUMNS; nCol++)
	{
		TDCCOLUMN& tdcc = COLUMNS[nCol];
		
		if (tdcc.nColID != TDCC_CLIENT)
		{
			int nItem = (nCol + 1); // zero'th column ignored

			m_hdrColumns.SetItemData(nItem, tdcc.nColID);
			m_hdrColumns.SetItemToolTip(nItem, CEnString(tdcc.nIDLongName));
		}
	}
	
	// add custom columns
	int nNumCols = (TDCC_CUSTOMCOLUMN_LAST - TDCC_CUSTOMCOLUMN_FIRST + 1);

	for (nCol = 0; nCol < nNumCols; nCol++)
	{
		m_lcColumns.InsertColumn((NUM_COLUMNS + nCol), _T(""), LVCFMT_LEFT, 0);
	}

	// and their IDs
	for (nCol = 0; nCol < nNumCols; nCol++)
	{
		int nItem = (NUM_COLUMNS + nCol);

		m_hdrColumns.SetItemData(nItem, (TDCC_CUSTOMCOLUMN_FIRST + nCol));
		m_hdrColumns.EnableItemTracking(nItem, FALSE);
		m_hdrColumns.ShowItem(nItem, TRUE);
	}

	RecalcColumnWidths();

	return TRUE;
}

void CTDCTaskCtrlBase::RecalcAllColumnWidths()
{
	m_hdrColumns.ClearAllTracked();
	RecalcColumnWidths();
}

void CTDCTaskCtrlBase::RecalcColumnWidths()
{
	RecalcColumnWidths(FALSE); // standard
	RecalcColumnWidths(TRUE); // custom
}

void CTDCTaskCtrlBase::RecalcColumnWidths(BOOL bCustom)
{
	CHoldRedraw hr(m_lcColumns);

	// recalc the widths of 'non-tracked' items
	CDC* pDC = m_lcColumns.GetDC();
	CFont* pOldFont = GraphicsMisc::PrepareDCFont(pDC, m_lcColumns);

	int nNumCols = m_hdrColumns.GetItemCount();
	
	for (int nItem = 1; nItem < nNumCols; nItem++)
	{		
		TDC_COLUMN nColID = (TDC_COLUMN)m_hdrColumns.GetItemData(nItem);
		BOOL bCustomCol = CTDCCustomAttributeHelper::IsCustomColumn(nColID);

		if ((bCustom && bCustomCol) || (!bCustom && !bCustomCol))
		{
			if (m_hdrColumns.IsItemVisible(nItem))
			{
				if (!m_hdrColumns.IsItemTracked(nItem))
				{
					int nColWidth = RecalcColumnWidth(nItem, pDC);
					m_hdrColumns.SetItemWidth(nItem, nColWidth);
				}
			}
			else
			{
				m_hdrColumns.SetItemWidth(nItem, 0);
			}
		}
	}

	// cleanup
	pDC->SelectObject(pOldFont);
	m_lcColumns.ReleaseDC(pDC);
}

void CTDCTaskCtrlBase::RecalcColumnWidth(TDC_COLUMN nColID)
{
	switch (nColID)
	{
	case TDCC_NONE:
		break;

	case TDCC_ALL:
		RecalcColumnWidths();
		break;

	default:
		{
			int nItem = m_hdrColumns.FindItem(nColID);
			ASSERT(nItem != -1);
			
			if (m_hdrColumns.IsItemVisible(nItem) && !m_hdrColumns.IsItemTracked(nItem))
			{
				CDC* pDC = m_lcColumns.GetDC();
				CFont* pOldFont = GraphicsMisc::PrepareDCFont(pDC, m_lcColumns);
				
				int nColWidth = RecalcColumnWidth(nItem, pDC);
				m_hdrColumns.SetItemWidth(nItem, nColWidth);
				
				pDC->SelectObject(pOldFont);
				m_lcColumns.ReleaseDC(pDC);
			}
		}
		break;
	}
}

void CTDCTaskCtrlBase::SaveState(CPreferences& prefs, const CString& sKey) const
{
	ASSERT (GetSafeHwnd());
	ASSERT (!sKey.IsEmpty());
	
	CDWordArray aOrder, aWidths, aTracked;

	GetColumnOrder(aOrder);
	GetColumnWidths(aWidths);
	GetTrackedColumns(aTracked);

	prefs.WriteProfileArray((sKey + _T("\\ColumnOrder")), aOrder);
	prefs.WriteProfileArray((sKey + _T("\\ColumnWidth")), aWidths);
	prefs.WriteProfileArray((sKey + _T("\\ColumnTracked")), aTracked);

	prefs.WriteProfileInt(sKey, _T("SplitPos"), GetSplitPos());
	
	SaveSortState(prefs, sKey);
}

void CTDCTaskCtrlBase::LoadState(const CPreferences& prefs, const CString& sKey)
{
	ASSERT (GetSafeHwnd());
	ASSERT (!sKey.IsEmpty());

	// make sure columns are configured right
	OnColumnVisibilityChange();
	
	// load column customisations
	CDWordArray aOrder, aWidths, aTracked;
	
	if (prefs.GetProfileArray((sKey + _T("\\ColumnOrder")), aOrder))
		SetColumnOrder(aOrder);

	if (prefs.GetProfileArray((sKey + _T("\\ColumnWidth")), aWidths))
		SetColumnWidths(aWidths);
	
	if (prefs.GetProfileArray((sKey + _T("\\ColumnTracked")), aTracked))
		SetTrackedColumns(aTracked);
	
	if (aOrder.GetSize() || aWidths.GetSize() || aTracked.GetSize())
		SetSplitPos(prefs.GetProfileInt(sKey, _T("SplitPos"), 300));
	else
		SetSplitPos(CalcSplitterPosToFitColumns());

	RefreshSize();

	LoadSortState(prefs, sKey);
}

void CTDCTaskCtrlBase::SaveSortState(CPreferences& prefs, const CString& sKey) const
{
	// ignore this if we have no tasks
	if (GetTaskCount() == 0)
		return;
	
	CString sSortKey = sKey + (_T("\\SortColState"));
	
	if (!sSortKey.IsEmpty())
	{
		prefs.WriteProfileInt(sSortKey, _T("Multi"), m_sort.bMulti);
		
		prefs.WriteProfileInt(sSortKey, _T("Column"), m_sort.single.nBy);
		prefs.WriteProfileInt(sSortKey, _T("Ascending"), m_sort.single.bAscending);
		
		prefs.WriteProfileInt(sSortKey, _T("Column1"), m_sort.multi.col1.nBy);
		prefs.WriteProfileInt(sSortKey, _T("Column2"), m_sort.multi.col2.nBy);
		prefs.WriteProfileInt(sSortKey, _T("Column3"), m_sort.multi.col3.nBy);
		prefs.WriteProfileInt(sSortKey, _T("Ascending1"), m_sort.multi.col1.bAscending);
		prefs.WriteProfileInt(sSortKey, _T("Ascending2"), m_sort.multi.col2.bAscending);
		prefs.WriteProfileInt(sSortKey, _T("Ascending3"), m_sort.multi.col3.bAscending);
	}
}

void CTDCTaskCtrlBase::LoadSortState(const CPreferences& prefs, const CString& sKey)
{
	CString sSortKey = sKey + (_T("\\SortColState"));
	
	// single sort
	m_sort.single.nBy = (TDC_COLUMN)prefs.GetProfileInt(sSortKey, _T("Column"), TDCC_NONE);
	m_sort.single.bAscending = prefs.GetProfileInt(sSortKey, _T("Ascending"), TRUE);
	
	// multi sort
	m_sort.bMulti = prefs.GetProfileInt(sSortKey, _T("Multi"), FALSE);
	
	m_sort.multi.col1.nBy = (TDC_COLUMN)prefs.GetProfileInt(sSortKey, _T("Column1"), TDCC_NONE);
	m_sort.multi.col1.bAscending = prefs.GetProfileInt(sSortKey, _T("Ascending1"), TRUE);
	m_sort.multi.col2.nBy = (TDC_COLUMN)prefs.GetProfileInt(sSortKey, _T("Column2"), TDCC_NONE);
	m_sort.multi.col3.nBy = (TDC_COLUMN)prefs.GetProfileInt(sSortKey, _T("Column3"), TDCC_NONE);
	m_sort.multi.col2.bAscending = prefs.GetProfileInt(sSortKey, _T("Ascending2"), TRUE);
	m_sort.multi.col3.bAscending = prefs.GetProfileInt(sSortKey, _T("Ascending3"), TRUE);
	
	m_sort.Validate();
}

int CALLBACK CTDCTaskCtrlBase::SortFuncMulti(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	TDSORTPARAMS* pSS = (TDSORTPARAMS*)lParamSort;
	
	ASSERT (pSS->sort.bMulti && pSS->sort.multi.IsSorting());
	
	int nCompare = SortTasks(lParam1, lParam2, 
							pSS->base, 
							pSS->sort.multi.col1, 
							pSS->bSortDueTodayHigh, 
							pSS->dwTimeTrackID);
	
	if ((nCompare == 0) && pSS->sort.multi.col2.IsSorting())
	{
		nCompare = SortTasks(lParam1, lParam2, 
							pSS->base, 
							pSS->sort.multi.col2, 
							pSS->bSortDueTodayHigh, 
							pSS->dwTimeTrackID);
		
		if ((nCompare == 0) && pSS->sort.multi.col3.IsSorting())
		{
			nCompare = SortTasks(lParam1, lParam2, 
								pSS->base, 
								pSS->sort.multi.col3, 
								pSS->bSortDueTodayHigh, 
								pSS->dwTimeTrackID);
		}
	}
	
	// finally, if the items are equal we sort by raw
	// position so that the sort is stable
	if (nCompare == 0)
	{
		nCompare = SortTasks(lParam1, lParam2, 
							pSS->base, 
							TDSORTCOLUMN(TDCC_NONE, FALSE),
							FALSE, 
							0);
	}
	
	return nCompare;
}

int CALLBACK CTDCTaskCtrlBase::SortFunc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	TDSORTPARAMS* pSS = (TDSORTPARAMS*)lParamSort;
	
	ASSERT (!pSS->sort.bMulti);
	
	int nCompare = SortTasks(lParam1, lParam2, 
							pSS->base, 
							pSS->sort.single, 
							pSS->bSortDueTodayHigh, 
							pSS->dwTimeTrackID);
	
	// finally, if the items are equal we sort by raw
	// position so that the sort is stable
	if (nCompare == 0)
	{
		nCompare = SortTasks(lParam1, lParam2, 
							pSS->base, 
							TDSORTCOLUMN(TDCC_NONE, FALSE),
							FALSE, 
							0);
	}
	
	return nCompare;
}

int CTDCTaskCtrlBase::SortTasks(LPARAM lParam1, LPARAM lParam2, 
								 const CTDCTaskCtrlBase& base, const TDSORTCOLUMN& sort, 
								 BOOL bSortDueTodayHigh, DWORD dwTimeTrackedID)
{
	DWORD dwTaskID1 = lParam1, dwTaskID2 = lParam2;
	
	// special cases first
	if (sort.IsSortingBy(TDCC_TRACKTIME))
	{
		BOOL bTracked1 = ((dwTaskID1 == dwTimeTrackedID) ? 1 : 0);
		BOOL bTracked2 = ((dwTaskID2 == dwTimeTrackedID) ? 1 : 0);
		ASSERT(!(bTracked1 && bTracked2));
		
		return (sort.bAscending ? (bTracked1 - bTracked2) : (bTracked2 - bTracked1));
	}
	else if (sort.IsSortingBy(TDCC_REMINDER))
	{
		BOOL bHasReminder1 = base.TaskHasReminder(dwTaskID1);
		BOOL bHasReminder2 = base.TaskHasReminder(dwTaskID2);
		
		return (sort.bAscending ? (bHasReminder1 - bHasReminder2) : (bHasReminder2 - bHasReminder1));
	}
	// handle custom attribute
	else if (sort.IsSortingByCustom())
	{
		TDCCUSTOMATTRIBUTEDEFINITION attribDef;
		
		// this can still fail
		if (!CTDCCustomAttributeHelper::GetAttributeDef(sort.nBy, base.m_aCustomAttribDefs, attribDef))
			return 0;
		
		return base.m_data.CompareTasks(dwTaskID1, dwTaskID2, attribDef, sort.bAscending);
	}
	
	// else default attribute
	return base.m_data.CompareTasks(dwTaskID1, dwTaskID2, sort.nBy, sort.bAscending, bSortDueTodayHigh);
}

DWORD CTDCTaskCtrlBase::HitTestTask(const CPoint& ptScreen) const
{
	DWORD dwTaskID = HitTestColumnsTask(ptScreen);

	if (dwTaskID == 0)
		dwTaskID = HitTestTasksTask(ptScreen);

	return dwTaskID;
}

int CTDCTaskCtrlBase::HitTestColumnsItem(const CPoint& pt, BOOL bClient, TDC_COLUMN& nColID, DWORD* pTaskID) const
{
	LVHITTESTINFO lvHit = { 0 };
	lvHit.pt = pt;

	if (!bClient)
		m_lcColumns.ScreenToClient(&lvHit.pt);
	
	ListView_SubItemHitTest(m_lcColumns, &lvHit);

	if ((lvHit.iItem < 0) || (lvHit.iSubItem < 0))
		return -1;

	nColID = GetColumnID(lvHit.iSubItem);
	ASSERT(nColID != TDCC_NONE);

	if (pTaskID)
	{
		*pTaskID = GetColumnItemTaskID(lvHit.iItem);
		ASSERT(*pTaskID);
	}
		
	return lvHit.iItem;
}

TDC_HITTEST CTDCTaskCtrlBase::HitTest(const CPoint& ptScreen) const
{
	if (PointInRect(ptScreen, m_hdrTasks, TRUE)) // task header
	{
		return TDCHT_COLUMNHEADER;
	}
	else if (PointInRect(ptScreen, m_hdrColumns, TRUE))	// column header
	{
		return TDCHT_COLUMNHEADER;
	}
	else if (PointInRect(ptScreen, Tasks(), TRUE)) // tasks
	{
		// see if we hit a task
		if (HitTestTasksTask(ptScreen))
		{
			return TDCHT_TASK;
		}
		else
		{
			CPoint ptClient(ptScreen);
			::ScreenToClient(Tasks(), &ptClient);

			if (PointInRect(ptClient, Tasks(), FALSE))
			{
				return TDCHT_TASKLIST;
			}
		}
	}
	else if (PointInRect(ptScreen, m_lcColumns, TRUE)) // columns
	{
		// see if we hit a task
		if (HitTestColumnsTask(ptScreen))
		{
			return TDCHT_TASK;
		}
		else
		{
			CPoint ptClient(ptScreen);
			m_lcColumns.ScreenToClient(&ptClient);

			if (PointInRect(ptClient, m_lcColumns, FALSE))
			{
				return TDCHT_TASKLIST;
			}
		}
	}
	
	// all else
	return TDCHT_NOWHERE;
}

DWORD CTDCTaskCtrlBase::HitTestColumnsTask(const CPoint& ptScreen) const
{
	// see if we hit a task in the list
	CPoint ptClient(ptScreen);
	m_lcColumns.ScreenToClient(&ptClient);
	
	int nItem = m_lcColumns.HitTest(ptClient);

	if (nItem != -1)
		return GetColumnItemTaskID(nItem);

	// all else
	return 0;
}

TDC_COLUMN CTDCTaskCtrlBase::HitTestColumn(const CPoint& ptScreen) const
{
	if (PointInRect(ptScreen, m_hdrColumns, TRUE) || // tree header
		PointInRect(ptScreen, Tasks(), TRUE)) // tree
	{
		return TDCC_CLIENT;
	}
	else if (PointInRect(ptScreen, m_hdrColumns, TRUE))	// column header
	{
		CPoint ptHeader(ptScreen);
		m_hdrColumns.ScreenToClient(&ptHeader);
		
		int nCol = m_hdrColumns.HitTest(ptHeader);
		
		if (nCol >= 0)
			return (TDC_COLUMN)m_hdrColumns.GetItemData(nCol);
	}
	else if (PointInRect(ptScreen, m_lcColumns, TRUE)) // columns
	{
		TDC_COLUMN nColID = TDCC_NONE;
		
		if (HitTestColumnsItem(ptScreen, FALSE, nColID) != -1)
			return nColID;
	}

	// else
	return TDCC_NONE;
}

void CTDCTaskCtrlBase::Release() 
{ 
	m_imageIcons.Clear();
	m_fonts.Clear();
	m_ilCheckboxes.DeleteImageList();
	
	GraphicsMisc::VerifyDeleteObject(m_brDue);
	GraphicsMisc::VerifyDeleteObject(m_brDueToday);

	if (::IsWindow(m_hdrColumns))
		m_hdrColumns.UnsubclassWindow();
	
	Unsync(); 
}

BOOL CTDCTaskCtrlBase::SetFont(HFONT hFont)
{	
	ASSERT(Tasks() != NULL);
	ASSERT(hFont);

	HFONT hTaskFont = (HFONT)::SendMessage(m_lcColumns, WM_GETFONT, 0, 0);
	BOOL bChange = !GraphicsMisc::SameFontNameSize(hFont, hTaskFont);
	
	if (bChange)
	{
		m_fonts.Clear();
		::SendMessage(Tasks(), WM_SETFONT, (WPARAM)hFont, TRUE);

		RecalcColumnWidths();
	}
	
	return bChange;
}

BOOL CTDCTaskCtrlBase::IsColumnLineOdd(int nItem) const
{
	return ((nItem % 2) == 1);
}

BOOL CTDCTaskCtrlBase::IsSorting() const
{
	return m_sort.IsSorting();
}

BOOL CTDCTaskCtrlBase::IsSortingBy(TDC_COLUMN nSortBy) const
{
	return (m_sort.IsSortingBy(nSortBy, TRUE));
}

void CTDCTaskCtrlBase::UpdateHeaderSorting()
{
	BOOL bEnable = HasStyle(TDCS_COLUMNHEADERSORTING);
	DWORD dwAdd(bEnable ? HDS_BUTTONS : 0), dwRemove(bEnable ? 0 : HDS_BUTTONS);

	if (m_hdrTasks.GetSafeHwnd())
		m_hdrTasks.ModifyStyle(dwRemove, dwAdd);

	if (m_hdrColumns)
		m_hdrColumns.ModifyStyle(dwRemove, dwAdd);
}

void CTDCTaskCtrlBase::Resort(BOOL bAllowToggle) 
{ 
	if (IsMultiSorting())
	{
		TDSORTCOLUMNS sort;

		GetSortBy(sort);
		MultiSort(sort);
	}
	else
	{
		Sort(GetSortBy(), bAllowToggle); 
	}
}

BOOL CTDCTaskCtrlBase::IsMultiSorting() const
{
	return (m_sort.bMulti && m_sort.multi.IsSorting());
}

void CTDCTaskCtrlBase::Sort(TDC_COLUMN nBy, BOOL bAllowToggle)
{
	// special case
	if (nBy == TDCC_NONE && !m_sort.IsSorting())
		return; // nothing to do

	BOOL bAscending = m_sort.single.bAscending;

	if (nBy != TDCC_NONE)
	{
		// first time?
		if ((bAscending == -1) || !m_sort.single.IsSortingBy(nBy))
		{
			const TDCCOLUMN* pTDCC = GetColumn(nBy);

			if (pTDCC)
			{
				bAscending = pTDCC->bSortAscending;
			}
			else if (CTDCCustomAttributeHelper::IsCustomColumn(nBy))
			{
				// TODO
				bAscending = FALSE;//(m_ctrlTreeList.Tree().GetGutterColumnSort(nBy) != NCGSORT_DOWN);
			}
		}
		// if there's been a mod since last sorting then its reasonable to assume
		// that the user is not toggling direction but wants to simply resort
		// in the same direction, otherwise we toggle the direction.
		else if (!m_sort.bModSinceLastSort && bAllowToggle)
		{
			bAscending = !bAscending; // toggle 
		}
		
		// update the column header
		SetSortColumn(nBy, (bAscending ? TDC_SORTUP : TDC_SORTDOWN));
	}
	else
	{
		ClearSortColumn();
		bAscending = TRUE;
	}

	m_sort.SetSortBy(nBy, bAscending);
	m_sort.bModSinceLastSort = FALSE;
	
	if (m_data.GetTaskCount() > 1)
		DoSort();
}

PFNTLSCOMPARE CTDCTaskCtrlBase::PrepareSort(TDSORTPARAMS& ss) const
{
	ss.sort = m_sort;
	ss.bSortChildren = TRUE;
	ss.bSortDueTodayHigh = HasColor(m_crDueToday);
	ss.dwTimeTrackID = m_dwTimeTrackTaskID;
	
	return (ss.sort.bMulti ? SortFuncMulti : SortFunc);
}

void CTDCTaskCtrlBase::DoSort()
{
	CHoldRedraw hr(*this, NCR_PAINT | NCR_NCPAINT | NCR_UPDATE);
	TDSORTPARAMS ss(*this);

	// Save the scroll pos of the columns
	int nScrollPos = m_lcColumns.GetScrollPos(SB_VERT);

	CTreeListSyncer::Sort(PrepareSort(ss), (LPARAM)&ss);

	ResyncSelection(m_lcColumns, Tasks(), FALSE);

	// restore the column scroll pos and sync the tree to it
	if (nScrollPos)
	{
		m_lcColumns.SendMessage(LVM_SCROLL, 0, (nScrollPos * GetItemHeight(m_lcColumns)));
		ResyncScrollPos(Tasks(), m_lcColumns);
	}

	EnsureSelectionVisible();
}

void CTDCTaskCtrlBase::GetSortBy(TDSORTCOLUMNS& sort) const
{
	sort = m_sort.multi;
}

void CTDCTaskCtrlBase::MultiSort(const TDSORTCOLUMNS& sort)
{
	ASSERT (sort.col1.IsSorting());

	if (!sort.col1.IsSorting())
		return;

	m_sort.SetSortBy(sort);
	m_sort.bModSinceLastSort = FALSE;

	if (m_data.GetTaskCount() > 1)
		DoSort();

	// clear sort direction on header
	ClearSortColumn();
}

void CTDCTaskCtrlBase::Resize(const CRect& rect)
{
	if (GetSafeHwnd())
		MoveWindow(rect);
	else
		CTreeListSyncer::Resize(rect);
}

void CTDCTaskCtrlBase::Resize()
{
	CRect rect;
	GetBoundingRect(rect);
	
	Resize(rect);
}

POSITION CTDCTaskCtrlBase::GetFirstSelectedTaskPos() const
{
	return m_lcColumns.GetFirstSelectedItemPosition();
}

DWORD CTDCTaskCtrlBase::GetNextSelectedTaskID(POSITION& pos) const
{
	if (pos == NULL)
		return 0;

	int nSel = m_lcColumns.GetNextSelectedItem(pos);

	if (nSel == -1)
		return 0;

	return GetColumnItemTaskID(nSel);
}

void CTDCTaskCtrlBase::OnNotifySplitterChange(int /*nSplitPos*/)
{
	InvalidateAll(TRUE);
}

void CTDCTaskCtrlBase::DrawSplitBar(CDC* pDC, const CRect& rSplitter, COLORREF crSplitBar)
{
	GraphicsMisc::DrawSplitBar(pDC, rSplitter, crSplitBar);
}

BOOL CTDCTaskCtrlBase::GetTaskTextColors(DWORD dwTaskID, COLORREF& crText, COLORREF& crBack, BOOL bRef) const
{
	const TODOITEM* pTDI = NULL;
	const TODOSTRUCTURE* pTDS = NULL;
	
	if (m_data.GetTask(dwTaskID, pTDI, pTDS))
		return GetTaskTextColors(pTDI, pTDS, crText, crBack, bRef, FALSE);
	
	// else
	return FALSE;
}

BOOL CTDCTaskCtrlBase::GetTaskTextColors(const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS,
										COLORREF& crText, COLORREF& crBack, BOOL bRef) const
{
	return GetTaskTextColors(pTDI, pTDS, crText, crBack, bRef, FALSE);
}

BOOL CTDCTaskCtrlBase::GetTaskTextColors(const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS, COLORREF& crText, 
										COLORREF& crBack, BOOL bRef, BOOL bSelected) const
{
	ASSERT(pTDI && pTDS);

	if (!(pTDI && pTDS))
		return FALSE;

	if (bRef == -1)
		bRef = pTDI->IsReference();

	// all else
	crBack = CLR_NONE;
	crText = GetSysColor(COLOR_WINDOWTEXT);

	BOOL bDone = m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL);

	if (bDone)
	{
		if (HasColor(m_crDone))
		{
			crText = m_crDone; // parent and/or item is done
		}
		else if (bRef && HasColor(m_crReference))
		{
			crText = m_crReference;
		}
		else
		{
			crText = pTDI->color; 
		}
	}
	else // all incomplete tasks
	{
		// if it's a ref task just return the ref colour
		if (bRef && HasColor(m_crReference))
		{
			crText = m_crReference;
		}
		else
		{
			BOOL bDue = m_data.IsTaskDue(pTDI, pTDS); // due today or overdue
			BOOL bDueToday = bDue ? m_data.IsTaskDue(pTDI, pTDS, TRUE) : FALSE;
			BOOL bOverDue = (bDue && !bDueToday);

			BOOL bStarted = m_data.IsTaskStarted(pTDI, pTDS); // started by now
			BOOL bStartedToday = bStarted ? m_data.IsTaskStarted(pTDI, pTDS, TRUE) : FALSE;
			BOOL bStartedBefore = (bStarted && !bStartedToday);

			// do we have a custom 'due today' color
			if (bDueToday && HasColor(m_crDueToday))
			{
				crText = m_crDueToday;
			}
			// likewise for overdue
			else if (bOverDue && HasColor(m_crDue))
			{
				crText = m_crDue;
			}
			// do we have a custom 'started today' color
			else if (bStartedToday && HasColor(m_crStartedToday))
			{
				crText = m_crStartedToday;
			}
			// likewise for started
			else if (bStartedBefore && HasColor(m_crStarted))
			{
				crText = m_crStarted;
			}
			else if (HasColor(m_crFlagged) && pTDI->bFlagged)
			{
				crText = m_crFlagged;
			}
			else if (HasStyle(TDCS_COLORTEXTBYPRIORITY))
			{
				int nPriority = FM_NOPRIORITY;
				
				if (bDueToday)
				{
					nPriority = m_data.GetHighestPriority(pTDI, pTDS, FALSE); // ignore due tasks
				}
				else if (bOverDue)
				{
					nPriority = 10;
				}
				else
				{
					nPriority = m_data.GetHighestPriority(pTDI, pTDS);
				}
				
				if (nPriority != FM_NOPRIORITY)
					crText = GetPriorityColor(nPriority); 
			}
			else if (HasStyle(TDCS_COLORTEXTBYATTRIBUTE))
			{
				switch (m_nColorByAttrib)
				{
				case TDCA_CATEGORY:
					GetAttributeColor(pTDI->GetCategory(0), crText);
					break;

				case TDCA_ALLOCBY:
					GetAttributeColor(pTDI->sAllocBy, crText);
					break;

				case TDCA_ALLOCTO:
					GetAttributeColor(pTDI->GetAllocTo(0), crText);
					break;

				case TDCA_STATUS:
					GetAttributeColor(pTDI->sStatus, crText);
					break;

				case TDCA_VERSION:
					GetAttributeColor(pTDI->sVersion, crText);
					break;

				case TDCA_EXTERNALID:
					GetAttributeColor(pTDI->sExternalID, crText);
					break;

				case TDCA_TAGS:
					GetAttributeColor(pTDI->GetTag(0), crText);
					break;

				default:
					ASSERT(0);
					break;
				}
			}
			else if (!HasStyle(TDCS_COLORTEXTBYNONE) && pTDI->color)
			{
				crText = pTDI->color; 
			}
		}
	}
	ASSERT(HasColor(crText));
	
	if (bSelected)
	{
		crText = GraphicsMisc::GetExplorerItemTextColor(crText, GMIS_SELECTED, GMIB_THEMECLASSIC);
	}
	else
	{
		if (HasStyle(TDCS_TASKCOLORISBACKGROUND) && 
			(crText != GetSysColor(COLOR_WINDOWTEXT)) &&
			!m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL))
		{
			crBack = crText;
			crText = GraphicsMisc::GetBestTextColor(crBack);
		}
	}

	return TRUE;
}

BOOL CTDCTaskCtrlBase::SetPriorityColors(const CDWordArray& aColors)
{
	ASSERT (aColors.GetSize() == 11);
	
	if ((aColors.GetSize() == 11) && !Misc::MatchAllT(aColors, m_aPriorityColors))
	{
		m_aPriorityColors.Copy(aColors);
			
		if (GetSafeHwnd())
			InvalidateAll();
			
		return TRUE;
	}
	
	// else
	return FALSE; // invalid combination or no change
}

BOOL CTDCTaskCtrlBase::SetStartedTaskColors(COLORREF crStarted, COLORREF crStartedToday)
{
	if ((m_crStarted != crStarted) || (m_crStartedToday != crStartedToday))
	{
		m_crStarted = crStarted;
		m_crStartedToday = crStartedToday;
		
		if (GetSafeHwnd())
			InvalidateAll();

		return TRUE;
	}

	// else no change
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SetDueTaskColors(COLORREF crDue, COLORREF crDueToday)
{
	if ((m_crDue != crDue) || (m_crDueToday != crDueToday))
	{
		if (m_crDue != crDue)
		{
			GraphicsMisc::VerifyDeleteObject(m_brDue);
			
			if (HasColor(crDue))
				m_brDue.CreateSolidBrush(crDue);

			m_crDue = crDue;
		}

		if (m_crDueToday != crDueToday)
		{
			GraphicsMisc::VerifyDeleteObject(m_brDueToday);
			
			if (HasColor(crDueToday))
				m_brDueToday.CreateSolidBrush(crDueToday);
			
			m_crDueToday = crDueToday;
		}

		if (GetSafeHwnd())
			InvalidateAll();

		return TRUE;
	}

	// else no change
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SetCompletedTaskColor(COLORREF color)
{
	if (m_crDone != color)
	{
		m_crDone = color;
		
		if (GetSafeHwnd())
			InvalidateAll();

		return TRUE;
	}

	// else no change
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SetFlaggedTaskColor(COLORREF color)
{
	if (m_crFlagged != color)
	{
		m_crFlagged = color;
		
		if (GetSafeHwnd())
			InvalidateAll();

		return TRUE;
	}

	// else no change
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SetReferenceTaskColor(COLORREF color)
{
	if (m_crReference != color)
	{
		m_crReference = color;
			
		if (GetSafeHwnd())
			InvalidateAll();

		return TRUE;
	}

	// else no change
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SetAttributeColors(TDC_ATTRIBUTE nAttrib, const CTDCColorMap& colors)
{
	// see if there is any change
	if ((m_nColorByAttrib == nAttrib) && m_mapAttribColors.GetCount())
	{
		POSITION pos = m_mapAttribColors.GetStartPosition();
		BOOL bMatch = TRUE;
		
		while (pos && bMatch)
		{
			CString sAttrib;
			COLORREF crNew, color;
			
			m_mapAttribColors.GetNextAssoc(pos, sAttrib, color);
			
			bMatch = (colors.Lookup(sAttrib, crNew) && (color != crNew));
		}
		
		if (bMatch)
			return FALSE; // no change

	}

	m_nColorByAttrib = nAttrib;
	m_mapAttribColors.RemoveAll();
	
	CString sAttrib;
	COLORREF color;
	POSITION pos = colors.GetStartPosition();
	
	while (pos)
	{
		colors.GetNextAssoc(pos, sAttrib, color);
		sAttrib.MakeUpper();
		
		m_mapAttribColors[sAttrib] = color;
	}
	
	if (GetSafeHwnd() && HasStyle(TDCS_COLORTEXTBYATTRIBUTE))
		InvalidateAll();

	return TRUE;
}

BOOL CTDCTaskCtrlBase::GetAttributeColor(const CString& sAttrib, COLORREF& color) const
{
	if (!sAttrib.IsEmpty())
	{
		CString sUpper(sAttrib);
		sUpper.MakeUpper();
		
		if (m_mapAttribColors.Lookup(sUpper, color))
			return TRUE; 
	}
	
	// else
	return FALSE;
}

COLORREF CTDCTaskCtrlBase::GetPriorityColor(int nPriority) const
{
	if (nPriority < 0 || nPriority >= m_aPriorityColors.GetSize())
		return 0;
	
	return (COLORREF)m_aPriorityColors[nPriority];
}

void CTDCTaskCtrlBase::DrawCommentsText(CDC* pDC, const CRect& rRow, const CRect& rLabel, const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS)
{
	CRect rClip;
	pDC->GetClipBox(rClip);
	
 	CRect rComments((rLabel.right + 10), (rRow.top + 1), min(rRow.right, m_hdrTasks.GetItemWidth(0)), rRow.bottom);

	if ((rClip.right <= rComments.left) || (rComments.Width() <= 0))
	{
		return; // nothing to do
	}

	if ((m_dwEditTitleTaskID != pTDS->GetTaskID()) && !pTDI->sComments.IsEmpty())
	{
		CString sComments(_T("[...]"));
		
		// if we're showing the comments but not the first line
		// and the first line is empty then scoot
		if (HasStyle(TDCS_SHOWCOMMENTSINLIST))
		{
			CString sTemp(pTDI->sComments);
			int nFind = sTemp.FindOneOf(_T("\n\r")); 
			
			if (HasStyle(TDCS_SHOWFIRSTCOMMENTLINEINLIST))
			{
				if (nFind == 0) 
				{
					return; 
				}
				else if (nFind > 0) 
				{
					sTemp = sTemp.Left(nFind); 
				}
			}
			else if (nFind != -1)
			{
				sTemp.Replace(_T("\r\n"), _T("|"));
				sTemp.Replace('\n', '|');

				// trim trailing delimiters
				sTemp.TrimRight(_T(" |"));
			}
			
			sComments = sTemp;
		}
		
		COLORREF crPrev = pDC->SetTextColor(COMMENTSCOLOR);
		
		if (m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL))
		{
			if (HasColor(m_crDone))
				crPrev = pDC->SetTextColor(m_crDone);
		}

		int nPrevBkMode = pDC->SetBkMode(TRANSPARENT);
		UINT nFlags = (DT_LEFT | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS | GraphicsMisc::GetRTLDrawTextFlags(Tasks()));

		pDC->DrawText(sComments, (LPRECT)&rComments, nFlags);
		pDC->SetTextColor(crPrev);
		pDC->SetBkMode(nPrevBkMode);
	}
}

CFont* CTDCTaskCtrlBase::GetTaskFont(const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS, BOOL bColumns)
{
	if (!m_fonts.GetHwnd() && !m_fonts.Initialise(Tasks()))
		return NULL;

	BOOL bStrikeThru = (HasStyle(TDCS_STRIKETHOUGHDONETASKS) && pTDI->IsDone());
	BOOL bBold = (pTDS->ParentIsRoot() && !bColumns);

	if (bBold || bStrikeThru)
		return m_fonts.GetFont(bBold, FALSE, FALSE, bStrikeThru);

	// else
	return NULL;
}

BOOL CTDCTaskCtrlBase::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	if ((pWnd == &m_lcColumns) && Misc::ModKeysArePressed(0))
	{
		TDC_COLUMN nColID = TDCC_NONE;
		int nHit = HitTestColumnsItem(::GetMessagePos(), FALSE, nColID);
		
		if (ItemColumnSupportsClickHandling(nHit, nColID))
		{
			::SetCursor(GraphicsMisc::HandCursor());
			return TRUE;
		}
	}
	
	// else
	return CWnd::OnSetCursor(pWnd, nHitTest, message);
}

BOOL CTDCTaskCtrlBase::TaskHasReminder(DWORD dwTaskID) const
{
	return CWnd::GetParent()->SendMessage(WM_TDCM_TASKHASREMINDER, dwTaskID, (LPARAM)CWnd::GetSafeHwnd());
}

void CTDCTaskCtrlBase::DrawGridlines(CDC* pDC, const CRect& rect, BOOL bSelected, BOOL bHorz, BOOL bVert)
{
	if (HasColor(m_crGridLine))
	{
		if (bHorz)
			GraphicsMisc::DrawHorzLine(pDC, rect.left, rect.right, (rect.bottom - 1), m_crGridLine);

		if (bVert)
		{
			CRect rGridline(rect);

			// don't overdraw selection
			if (bSelected)
				rGridline.DeflateRect(0, 1);

			GraphicsMisc::DrawVertLine(pDC, rGridline.top, rGridline.bottom, rect.right - 1, m_crGridLine);
		}
	}
}

LRESULT CTDCTaskCtrlBase::OnListCustomDraw(NMLVCUSTOMDRAW* pLVCD)
{
	ASSERT(pLVCD->nmcd.hdr.hwndFrom == m_lcColumns);

	if (pLVCD->nmcd.hdr.hwndFrom == m_lcColumns)
	{
		switch (pLVCD->nmcd.dwDrawStage)
		{
		case CDDS_PREPAINT:
			return CDRF_NOTIFYITEMDRAW;
			
		case CDDS_ITEMPREPAINT:
			{
				// task data
				int nItem = (int)pLVCD->nmcd.dwItemSpec;
				DWORD dwTaskID = GetColumnItemTaskID(nItem), dwTrueID(dwTaskID);
				
				const TODOITEM* pTDI = NULL;
				const TODOSTRUCTURE* pTDS = NULL;
				
				if (m_data.GetTask(dwTrueID, pTDI, pTDS))
				{
					CDC* pDC = CDC::FromHandle(pLVCD->nmcd.hdc);

					// draw gridlines and row colour full width of list
					BOOL bAlternate = (HasColor(m_crAltLine) && IsColumnLineOdd(nItem));
					COLORREF crRowBack = (bAlternate ? m_crAltLine : GetSysColor(COLOR_WINDOW));
					
					CRect rItem;
					m_lcColumns.GetItemRect(nItem, rItem, LVIR_BOUNDS);

					// this call will update rFullWidth to full client width
					CRect rFullWidth(rItem);
					GraphicsMisc::FillItemRect(pDC, rFullWidth, crRowBack, m_lcColumns);

					// if the columns are to the right then
					// we don't want to draw the rounded end 
					// on the left so it looks continuous with the tasks
					DWORD dwFlags = GMIB_THEMECLASSIC;

					if (HasStyle(TDCS_RIGHTSIDECOLUMNS))
						dwFlags |= GMIB_CLIPLEFT;

					// selection state
					GM_ITEMSTATE nState = GetColumnItemState(nItem);
					
					if ((nState == GMIS_SELECTEDNOTFOCUSED) && m_dwEditTitleTaskID)
						nState = GMIS_SELECTED;
					
					BOOL bSelected = (nState != GMIS_NONE);
					BOOL bRef = (dwTaskID != dwTrueID);
					
					// colors
					COLORREF crText, crBack;
					GetTaskTextColors(pTDI, pTDS, crText, crBack, bRef, bSelected);

					// draw task background
					// Note: using the non-full width item rect
					if (!bSelected && HasColor(crBack))
						pDC->FillSolidRect(rItem, crBack);

					// draw horz gridline before selection
					DrawGridlines(pDC, rFullWidth, FALSE, TRUE, FALSE);

					// draw task background
					// Note: using the non-full width item rect
					if (bSelected)
					{
						crText = GraphicsMisc::GetExplorerItemTextColor(crText, nState, GMIB_THEMECLASSIC);
						GraphicsMisc::DrawExplorerItemBkgnd(pDC, m_lcColumns, nState, rItem, dwFlags);
					}

					// draw row text
					CFont* pOldFont = pDC->SelectObject(GetTaskFont(pTDI, pTDS));
					
					DrawColumnsRowText(pDC, nItem, dwTaskID, pTDI, pTDS, crText, bSelected);

					pDC->SelectObject(pOldFont);
				}
			}
			return CDRF_SKIPDEFAULT;
		}
	}
	
	return CDRF_DODEFAULT;
}

BOOL CTDCTaskCtrlBase::HasThemedState(GM_ITEMSTATE nState) const
{ 
	return ((nState != GMIS_NONE) && CThemed::AreControlsThemed()); 
}

void CTDCTaskCtrlBase::DrawTasksRowBackground(CDC* pDC, const CRect& rRow, const CRect& rLabel, GM_ITEMSTATE nState, BOOL bAlternate, COLORREF crBack)
{
	if (nState == GMIS_NONE)
	{
		// if we have gridlines we don't fill the bottom line so 
		// as to avoid overwriting gridlines previously drawn
		CRect rBack(rLabel);

		if (HasColor(m_crGridLine))
			rBack.bottom--;
		
		if (!HasColor(crBack))
			crBack = ((bAlternate && HasColor(m_crAltLine)) ? m_crAltLine : GetSysColor(COLOR_WINDOW));
		
		pDC->FillSolidRect(rBack, crBack);
	}
	
	// draw horz gridline before selection
	DrawGridlines(pDC, rRow, FALSE, TRUE, FALSE);
	
	if (nState != GMIS_NONE) // selection of some sort
	{
		DWORD dwFlags = (GMIB_THEMECLASSIC | GMIB_EXTENDRIGHT);
		
		// if the columns are on the right then
		// we don't want to draw the rounded end 
		// on the right so it looks continuous with the columns
		if (HasStyle(TDCS_RIGHTSIDECOLUMNS))
			dwFlags |= GMIB_CLIPRIGHT;
		
		GraphicsMisc::DrawExplorerItemBkgnd(pDC, Tasks(), nState, rLabel, dwFlags);
	}
}

void CTDCTaskCtrlBase::DrawColumnsRowText(CDC* pDC, int nItem, DWORD dwTaskID, const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS, 
										  COLORREF crText, BOOL bSelected)
{
	DWORD dwTrueID = pTDS->GetTaskID();

	// draw each column separately
	int nNumCol = m_hdrColumns.GetItemCount();
	int nLastCol = m_hdrColumns.GetLastVisibleItem();
	CRect rSubItem, rClient;

	m_lcColumns.GetClientRect(rClient);

	for (int nCol = 1; nCol < nNumCol; nCol++)
	{
		if (!m_lcColumns.GetSubItemRect(nItem, nCol, LVIR_BOUNDS, rSubItem))
			continue;

		// don't draw columns outside of client rect
		if (!CRect().IntersectRect(rClient, rSubItem))
			continue;

		// draw vertical gridlines for all but the
		// last item if the row is selected
		if (!(bSelected && (nCol == nLastCol)))
			DrawGridlines(pDC, rSubItem, bSelected, FALSE, TRUE);

		// don't draw min sized columns
		if (rSubItem.Width() <= MIN_COL_WIDTH)
			continue;
		
		TDC_COLUMN nColID = (TDC_COLUMN)m_hdrColumns.GetItemData(nCol);

		// Note: we pass dwTaskID NOT dwTrueID here so that references 
		// can be handled correctly
		CString sTaskColText = GetTaskColumnText(dwTaskID, pTDI, pTDS, nColID);
		
		const TDCCOLUMN* pCol = GetColumn(nColID);

		switch (nColID)
		{
		case TDCC_POSITION:
			DrawColumnText(pDC, sTaskColText, rSubItem, pCol->nTextAlignment, crText);
			break;
			
		case TDCC_RISK:
		case TDCC_RECURRENCE:
		case TDCC_ID:
		case TDCC_PARENTID:
		case TDCC_RECENTEDIT:
		case TDCC_COST:
		case TDCC_EXTERNALID:
		case TDCC_VERSION:
		case TDCC_ALLOCTO:
		case TDCC_ALLOCBY:
		case TDCC_STATUS:
		case TDCC_CATEGORY:
		case TDCC_TAGS:
		case TDCC_CREATEDBY:
		case TDCC_PATH:
		case TDCC_REMAINING:
		case TDCC_SUBTASKDONE:
		case TDCC_TIMEEST:
			DrawColumnText(pDC, sTaskColText, rSubItem, pCol->nTextAlignment, crText);
			break;
			
		case TDCC_TIMESPENT:
			{
				// show text in red if we're currently tracking
				COLORREF crTemp = ((m_dwTimeTrackTaskID == dwTrueID) ? 255 : crText);

				DrawColumnText(pDC, sTaskColText, rSubItem, pCol->nTextAlignment, crTemp);
			}
			break;
			
		case TDCC_PRIORITY:
			// priority color
			if (!HasStyle(TDCS_DONEHAVELOWESTPRIORITY) || !m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL))
			{
				rSubItem.DeflateRect(2, 1, 3, 2);
				
				// first draw the priority colour
				int nPriority = m_data.GetHighestPriority(pTDI, pTDS, FALSE);
				BOOL bHasPriority = (nPriority != FM_NOPRIORITY);
				
				if (bHasPriority)
				{
					COLORREF crFill = GetPriorityColor(nPriority);
					COLORREF crBorder = GraphicsMisc::Darker(crFill, 0.5);
					
					GraphicsMisc::DrawRect(pDC, rSubItem, crFill, crBorder);//, nCornerRadius);
				}
				
				// then, if the task is also due, draw a small tag
				if (m_data.IsTaskDue(pTDI, pTDS))
				{
					BOOL bDueToday = m_data.IsTaskDue(pTDI, pTDS, TRUE);
					
					// unless it's due today and the user doesn't want today's tasks marked as due
					// or there's no due color 
					if ((!bDueToday && HasColor(m_crDue)) || 
						(bDueToday && HasColor(m_crDueToday)))
					{
						POINT pt[3] = 
						{ 
							{ rSubItem.left, rSubItem.top + 7 }, 
						 	{ rSubItem.left, rSubItem.top }, 
						 	{ rSubItem.left + 7, rSubItem.top } 
						};
						
						CBrush* pOldBr = pDC->SelectObject(bDueToday ? &m_brDueToday : &m_brDue);
						pDC->SelectStockObject(NULL_PEN);
						 
						pDC->Polygon(pt, 3);
						pDC->SelectObject(pOldBr);
						
						// a black line between the two
						pDC->SelectStockObject(BLACK_PEN);
						pDC->MoveTo(rSubItem.left, rSubItem.top + 6);
						pDC->LineTo(rSubItem.left + 7, rSubItem.top - 1);
					}
				}
				
				// draw priority number over the top
				if (bHasPriority && !HasStyle(TDCS_HIDEPRIORITYNUMBER))
				{
					COLORREF color = GetPriorityColor(nPriority);
					
					// pick best colour for text
					BYTE nLum = RGBX(color).Luminance();
					COLORREF crTemp = ((nLum < 128) ? RGB(255, 255, 255) : 0);
					
					DrawColumnText(pDC, sTaskColText, rSubItem, pCol->nTextAlignment, crTemp);
				}
			}
			break;
			
		case TDCC_PERCENT:
			if (!sTaskColText.IsEmpty())
			{
				if (HasStyle(TDCS_SHOWPERCENTASPROGRESSBAR))
				{
					// if the task is done then draw in 'done' colour else priority colour
					BOOL bDone = m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL);
					
					COLORREF crBar(m_crDone);
					
					if (!bDone || !HasStyle(TDCS_DONEHAVELOWESTPRIORITY)) // determine appropriate priority
					{
						int nPriority = m_data.GetHighestPriority(pTDI, pTDS, FALSE);
						crBar = GetPriorityColor(nPriority);
						
						// check for due
						if (m_data.IsTaskDue(pTDI, pTDS))
						{
							BOOL bDueToday = m_data.IsTaskDue(pTDI, pTDS, TRUE);
							
							if (bDueToday && HasColor(m_crDueToday))
							{
								crBar = m_crDueToday;
							}
							else if (HasColor(m_crDue))
							{
								crBar = m_crDue;
							}
						}
					}
					
					if (HasColor(crBar))
					{
						rSubItem.DeflateRect(2, 1, 3, 2);
						CRect rProgress(rSubItem);
						
						// draw border
						COLORREF crBorder = GraphicsMisc::Darker(crBar, 0.5);
						GraphicsMisc::DrawRect(pDC, rProgress, CLR_NONE, crBorder);
						
						// draw fill
						int nPercent = m_data.CalcPercentDone(pTDI, pTDS);

						rProgress.DeflateRect(1, 1);
						rProgress.right = rProgress.left + MulDiv(rProgress.Width(), nPercent, 100);
						
						if (rProgress.Width() > 0)
						{
							pDC->FillSolidRect(rProgress, crBar); 
							
							// draw text in colour to suit pbar
							// modify text colour to suit
							BYTE nLum = RGBX(crBar).Luminance();
							COLORREF crTemp = (nLum < 128) ? RGB(255, 255, 255) : 0;
							
							DrawColumnText(pDC, sTaskColText, rSubItem, pCol->nTextAlignment, crTemp);
							
							// then clip out the pbar rect
							pDC->ExcludeClipRect(rProgress);
						}
					}
				}
				
				// draw default text
				DrawColumnText(pDC, sTaskColText, rSubItem, pCol->nTextAlignment, crText);
			}
			break;
			
		case TDCC_TRACKTIME:
			// show icon in red if we're currently tracking
			if (m_dwTimeTrackTaskID == dwTrueID)
			{
				DrawColumnText(pDC, CLOCKCHAR, rSubItem, pCol->nTextAlignment, 255, FALSE, &GraphicsMisc::WingDings(), TRUE);
			}
			break;
			
		case TDCC_FLAG:
			if (pTDI->bFlagged)
			{
				DrawColumnText(pDC, FLAGCHAR, rSubItem, pCol->nTextAlignment, crText, FALSE, &GraphicsMisc::WingDings(), TRUE);
			}
			break;
			
		case TDCC_REMINDER:
			if (TaskHasReminder(dwTrueID))
			{
				DrawColumnText(pDC, BELLCHAR, rSubItem, pCol->nTextAlignment, crText, FALSE, &GraphicsMisc::WingDings(), TRUE);
			}
			break;
			
		case TDCC_STARTDATE:
			{
				COleDateTime date;
				BOOL bDone = m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL);
				
				if (bDone && !HasStyle(TDCS_HIDESTARTDUEFORDONETASKS))
					date = pTDI->dateStart;
				
				else if (!bDone)
					date = m_data.CalcStartDate(pTDI, pTDS);
				
				DrawColumnDate(pDC, date, TDCD_START, rSubItem, crText);
			}
			break;
			
		case TDCC_DUEDATE:
			{
				COleDateTime date;
				BOOL bDone = m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL);
				
				if (bDone && !HasStyle(TDCS_HIDESTARTDUEFORDONETASKS))
					date = pTDI->dateDue;
				
				else if (!bDone)
					date = m_data.CalcDueDate(pTDI, pTDS);
				
				DrawColumnDate(pDC, date, TDCD_DUE, rSubItem, crText);
			}
			break;
			
			
		case TDCC_DONEDATE:
			DrawColumnDate(pDC, pTDI->dateDone, TDCD_DONE, rSubItem, crText);
			break;
			
		case TDCC_CREATIONDATE:
			DrawColumnDate(pDC, pTDI->dateCreated, TDCD_CREATE, rSubItem, crText);
			break;
			
		case TDCC_LASTMOD:
			DrawColumnDate(pDC, pTDI->dateLastMod, TDCD_LASTMOD, rSubItem, crText);
			break;
			
		case TDCC_ICON:
			if (rSubItem.Width() >= 16)
			{
				int nIcon = m_ilTaskIcons.GetImageIndex(pTDI->sIcon);
									
				if (nIcon == -1 && pTDS->HasSubTasks() && HasStyle(TDCS_SHOWPARENTSASFOLDERS))
					nIcon = 0;
												
				if (nIcon >= 0)
				{
					CPoint pt(CalcColumnIconRenderPos(rSubItem));
					m_ilTaskIcons.Draw(pDC, nIcon, pt, ILD_TRANSPARENT);
				}
			}
			break;
			
		case TDCC_DEPENDENCY:
			if (pTDI->aDependencies.GetSize())
			{
				// draw our app icon 
				HICON hIcon = (HICON)AfxGetMainWnd()->SendMessage(WM_GETICON, ICON_SMALL);
				
				if (hIcon)
				{
					CPoint pt(CalcColumnIconRenderPos(rSubItem));
					::DrawIconEx(pDC->GetSafeHdc(), pt.x, pt.y, hIcon, 16, 16, 0, NULL, DI_NORMAL);
				}
			}
			break;
			
		case TDCC_FILEREF:
			if (!sTaskColText.IsEmpty())
			{
				CPoint pt(CalcColumnIconRenderPos(rSubItem));

				// first check for a tdl://
				if (sTaskColText.Find(TDL_PROTOCOL) != -1)
				{
					// draw our app icon 
					HICON hIcon = (HICON)AfxGetMainWnd()->SendMessage(WM_GETICON, ICON_SMALL);
					
					if (hIcon)
						::DrawIconEx(pDC->GetSafeHdc(), pt.x, pt.y, hIcon, 16, 16, 0, NULL, DI_NORMAL);
				}
				else
				{
					// get the associated image, failing if necessary
					sTaskColText.Remove('\"'); // remove quotes

					if (CEnBitmap::IsSupportedImageFile(sTaskColText) && 
						(m_imageIcons.HasIcon(sTaskColText) || m_imageIcons.Add(sTaskColText, sTaskColText)))
					{
						m_imageIcons.Draw(pDC, sTaskColText, pt);
					}
					else
					{
						int nImage = m_ilFileRef.GetFileImageIndex(sTaskColText, TRUE);
						
						if (nImage == -1 && HasStyle(TDCS_SHOWNONFILEREFSASTEXT))
						{
							DrawColumnText(pDC, sTaskColText, rSubItem, DT_LEFT, crText);
						}
						else
						{
							if (nImage == -1)
								nImage = m_ilFileRef.GetFileImageIndex(sTaskColText, FALSE);
							
							if (nImage >= 0)
							{
								m_ilFileRef.GetImageList()->Draw(pDC, nImage, pt, ILD_TRANSPARENT);
							}
						}
					}
				}
			}
			break;
			
		case TDCC_DONE:
			DrawColumnCheckBox(pDC, rSubItem, pTDI->IsDone());
			break;
			
		default:
			// custom attribute columns
			VERIFY (DrawItemCustomColumn(pTDI, pTDS, nColID, pDC, rSubItem, crText));
			break;
		}
	}
}

void CTDCTaskCtrlBase::DrawColumnCheckBox(CDC* pDC, const CRect& rSubItem, BOOL bChecked)
{
	CPoint pt(CalcColumnIconRenderPos(rSubItem, TRUE));
				
	int nImage = (bChecked ? 2 : 1); // first image is blank
	m_ilCheckboxes.Draw(pDC, nImage, pt, ILD_TRANSPARENT);
}

CPoint CTDCTaskCtrlBase::CalcColumnIconRenderPos(const CRect& rSubItem, BOOL bCheckbox)
{
	CPoint pt(rSubItem.CenterPoint());

	// if the line height is odd, move one pixel down
	// to avoid collision with selection rect
	if (bCheckbox)
		pt.y += (rSubItem.Height() % 2);

	// assume 16 px icons
	pt.x -= 8;
	pt.y -= 8;

	return pt;
}

BOOL CTDCTaskCtrlBase::DrawItemCustomColumn(const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS, TDC_COLUMN nColID, 
											CDC* pDC, const CRect& rSubItem, COLORREF crText)
{
	if (!CTDCCustomAttributeHelper::IsCustomColumn(nColID))
		return FALSE;

	TDCCUSTOMATTRIBUTEDEFINITION attribDef;
	
	if (!CTDCCustomAttributeHelper::GetAttributeDef(nColID, m_aCustomAttribDefs, attribDef))
		return TRUE;

	CString sTaskColText = GetTaskCustomColumnText(pTDI, pTDS, nColID);
	CRect rCol(rSubItem);
	
	switch (attribDef.GetDataType())
	{
	case TDCCA_DATE:
		{
			double dDate = TDCCADATA(sTaskColText).AsDate();
			m_data.CalcTaskCustomAttributeData(pTDI, pTDS, attribDef, dDate);

			DrawColumnDate(pDC, dDate, TDCD_CUSTOM, rCol, crText);
		}
		break;
		
	case TDCCA_ICON:
		if (!sTaskColText.IsEmpty() && (rCol.Width() > 16)) // min width for one icon
		{
			TDCCADATA data(sTaskColText);

			CStringArray aImages;
			int nNumImage = data.AsArray(aImages);
			int nTotalWidth = (nNumImage * 18) - 2;

			// centre icon vertically
			CPoint pt(rCol.left, (rCol.CenterPoint().y - 8));
			int nTextAlign = attribDef.nTextAlignment;
			
			CString sName;
			
			if (nNumImage == 1)
				sName = attribDef.GetImageName(data.AsString());
			
			switch (nTextAlign)
			{
			case DT_RIGHT:
				rCol.right -= (LV_COLPADDING + 16);
				pt.x = rCol.right;
				break;
				
			case DT_CENTER:
				// if there is associated text then we align left
				if (sName.IsEmpty())
				{
					pt.x = (rCol.left + ((rCol.Width() - nTotalWidth) / 2));
					break;
				}
				else 
				{
					nTextAlign = DT_LEFT;
					// fall thru
				}
				
			case DT_LEFT:
			default:
				rCol.left += 16;
				pt.x += LV_COLPADDING;
				break;
			}

			// calculate icon and text position
			rCol.left += LV_COLPADDING;
			BOOL bOverrun = FALSE;

			for (int nImg = 0; ((nImg < nNumImage) && !bOverrun); nImg++)
			{
				CString sImage, sDummy;

				if (TDCCUSTOMATTRIBUTEDEFINITION::DecodeImageTag(aImages[nImg], sImage, sDummy))
				{
					m_ilTaskIcons.Draw(pDC, sImage, pt, ILD_TRANSPARENT);
					pt.x += 18;

					bOverrun = ((pt.x + 16) > rCol.right);
				}
			}
			
			// optional text for single list images
			if (!bOverrun && (nNumImage == 1) && attribDef.IsList() && !sName.IsEmpty())
			{
				DrawColumnText(pDC, sName, rCol, nTextAlign, crText);
			}
		}
		break;
		
	case TDCCA_BOOL:
		DrawColumnCheckBox(pDC, rSubItem, !sTaskColText.IsEmpty());
		break;

	case TDCCA_DOUBLE:
	case TDCCA_INTEGER:
		DrawColumnText(pDC, sTaskColText, rCol, attribDef.nTextAlignment, crText);
		break;

	default:
		if (!sTaskColText.IsEmpty())
		{
			if (attribDef.GetListType() == TDCCA_FIXEDMULTILIST || attribDef.GetListType() == TDCCA_AUTOMULTILIST)
			{
				CStringArray aItems;
				
				if (TDCCADATA(sTaskColText).AsArray(aItems))
					DrawColumnText(pDC, Misc::FormatArray(aItems), rCol, attribDef.nTextAlignment, crText);
			}
			else
			{
				DrawColumnText(pDC, sTaskColText, rCol, attribDef.nTextAlignment, crText);
			}
		}
	}

	return TRUE; // we handled it
}

BOOL CTDCTaskCtrlBase::FormatDate(const COleDateTime& date, TDC_DATE nDate, CString& sDate, CString& sTime, CString& sDow) const
{	
	sDate.Empty();
	sTime.Empty();
	sDow.Empty();
	
	if (CDateHelper::IsDateSet(date))
	{
		DWORD dwFmt = 0;
		
		if (HasStyle(TDCS_SHOWWEEKDAYINDATES))
			dwFmt |= DHFD_DOW;
		
		if (HasStyle(TDCS_SHOWDATESINISO))
			dwFmt |= DHFD_ISO;
		
		if (WantDrawColumnTime(nDate) && CDateHelper::DateHasTime(date))
		{
			dwFmt |= DHFD_TIME | DHFD_NOSEC;
		}
		
		return CDateHelper::FormatDate(date, dwFmt, sDate, sTime, sDow);
	}
	
	// else
	return FALSE;
}

void CTDCTaskCtrlBase::DrawColumnDate(CDC* pDC, const COleDateTime& date, TDC_DATE nDate, const CRect& rect, COLORREF crText)
{
	if (!CDateHelper::IsDateSet(date)) 
		return; // nothing to do
	
	CString sDate, sTime, sDow;
	CRect rDraw(rect);
	
	if (!FormatDate(date, nDate, sDate, sTime, sDow))
		return; // nothing to do
	
	// calculate max date widths for aligning the various parts
	COleDateTime dateMax(2000, 12, 31, 23, 59, 0);
	
	CString sDateMax, sTimeMax, sDummy;
	FormatDate(dateMax, nDate, sDateMax, sTimeMax, sDummy);
	
	// also compensate for the padding that happens inside DrawGutterItemText
	int nSpace = pDC->GetTextExtent(_T(" ")).cx;
	
	int nMaxDate = (pDC->GetTextExtent(sDateMax).cx + nSpace);
	int nMaxTime = (pDC->GetTextExtent(sTimeMax).cx + nSpace);
	
	// draw time
	if (WantDrawColumnTime(nDate))
	{
		COLORREF crTime = crText;
		
		// if NO time component, render 'default' start and due time 
		// in a lighter colour to indicate it wasn't user-set
		if (sTime.IsEmpty())
		{
			int nHour = ((nDate == TDCD_DUE) ? 23 : 0);
			int nMin = ((nDate == TDCD_DUE) ? 59 : 0);
			
			sTime = CTimeHelper::FormatClockTime(nHour, nMin, 0, FALSE, HasStyle(TDCS_SHOWDATESINISO));
			
			COLORREF crActualText = (HasColor(crText) ? crText : pDC->GetTextColor());
			crTime = GraphicsMisc::Lighter(crActualText, 0.5);
		}
		
		// draw and adjust rect
		if (!sTime.IsEmpty())
		{
			DrawColumnText(pDC, sTime, rDraw, TA_RIGHT, crTime);
			rDraw.right -= nMaxTime;
		}
	}
	
	// then date
	DrawColumnText(pDC, sDate, rDraw, TA_RIGHT, crText);
	rDraw.right -= nMaxDate;
	
	// then dow
	if (!sDow.IsEmpty())
		DrawColumnText(pDC, sDow, rDraw, TA_RIGHT, crText);
}

void CTDCTaskCtrlBase::DrawColumnText(CDC* pDC, const CString& sText, const CRect& rect, int nAlign, COLORREF crText, BOOL bTaskTitle, CFont* pFont, BOOL bSymbol)
{
	ASSERT(crText != CLR_NONE);

	if (sText.IsEmpty())
		return;
	
	CRect rText(rect);
	CPoint ptText(0, rText.top);
	
	switch (nAlign)
	{
	case DT_LEFT:
		rText.left += LV_COLPADDING;
		break;
		
	case DT_RIGHT:
		rText.right -= LV_COLPADDING;
		break;
		
	case DT_CENTER:
		break;
	}
	
	UINT nFlags = (nAlign | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | (bTaskTitle ? DT_END_ELLIPSIS : 0) | GraphicsMisc::GetRTLDrawTextFlags(Tasks()));
	COLORREF crOld = pDC->SetTextColor(crText);
	
	pDC->SetBkMode(TRANSPARENT);
	
	if (bSymbol)
		GraphicsMisc::DrawAnsiSymbol(pDC, (char)sText[0], rect, nFlags, pFont);
	else
		pDC->DrawText(sText, rText, nFlags);
	
	// cleanup
	if (HasColor(crOld))
		pDC->SetTextColor(crOld);
}

LRESULT CTDCTaskCtrlBase::OnHeaderCustomDraw(NMCUSTOMDRAW* pNMCD)
{
	switch (pNMCD->dwDrawStage)
	{
	case CDDS_PREPAINT:
		return CDRF_NOTIFYITEMDRAW;
		
	case CDDS_ITEMPREPAINT:
		{
			// don't draw columns having min width
			CRect rItem(pNMCD->rc);

			if (rItem.Width() > MIN_COL_WIDTH)
				return CDRF_NOTIFYPOSTPAINT;
		}
		break;

	case CDDS_ITEMPOSTPAINT:
		{
			// don't draw columns having min width
			CRect rItem(pNMCD->rc);

			if (rItem.Width() > MIN_COL_WIDTH)
			{
				CDC* pDC = CDC::FromHandle(pNMCD->hdc);
				CFont* pFont = NULL;

				// draw sort direction
				int nCol = (int)pNMCD->dwItemSpec;
				TDC_COLUMN nColID = (TDC_COLUMN)pNMCD->lItemlParam;

				if (nColID == m_nSortColID)
				{
					BOOL bUp = (m_nSortDir == TDC_SORTUP);
					GetColumnHeaderCtrl(nColID).DrawItemSortArrow(pDC, nCol, bUp);
				}

				const TDCCOLUMN* pTDCC = GetColumn(nColID);
				int nAlignment = DT_LEFT;
				
				if (pTDCC)
				{
					nAlignment = pTDCC->GetColumnHeaderAlignment();

					// handle symbol fonts
					if (pTDCC->bSymbolFont)
					{
						CString sFont(Misc::ToUpper(pTDCC->szFont));
						
						if (sFont == _T("WINGDINGS"))
						{
							pFont = &GraphicsMisc::WingDings();
						}
						else if (sFont == _T("MARLETT"))
						{
							pFont = &GraphicsMisc::Marlett();
						}
						else
						{
							ASSERT(0);
							break;
						}

						TCHAR szCol[2] = { (TCHAR)pTDCC->nIDName, 0 };
						DrawColumnText(pDC, szCol, pNMCD->rc, DT_CENTER, GetSysColor(COLOR_WINDOWTEXT), FALSE, pFont, TRUE);
						
						return CDRF_SKIPDEFAULT;
					}
				}

				// Handle RTL text column headers
				if (GraphicsMisc::GetRTLDrawTextFlags(pNMCD->hdr.hwndFrom) == DT_RTLREADING)
				{
					CString sColumn(GetColumnHeaderCtrl(nColID).GetItemText(nCol));
					DrawColumnText(pDC, sColumn, pNMCD->rc, nAlignment, GetSysColor(COLOR_WINDOWTEXT));
					
					return CDRF_SKIPDEFAULT;
				}
			}
		}
		break;
	}
	
	return CDRF_DODEFAULT;
}

const CEnHeaderCtrl& CTDCTaskCtrlBase::GetColumnHeaderCtrl(TDC_COLUMN nColID) const
{
	return ((nColID == TDCC_CLIENT) ? m_hdrTasks : m_hdrColumns);
}

TDCCOLUMN* CTDCTaskCtrlBase::GetColumn(TDC_COLUMN nColID)
{
	if (CTDCCustomAttributeHelper::IsCustomColumn(nColID))
		return NULL;

	ASSERT(!s_mapColumns.IsEmpty());

	TDCCOLUMN* pCol = NULL;
	VERIFY(s_mapColumns.Lookup(nColID, pCol));

	ASSERT(pCol);
	return pCol;
}

TDC_COLUMN CTDCTaskCtrlBase::GetColumnID(int nCol) const
{
	switch (nCol)
	{
	case 0:
		// zero is always 'tasks'
		return TDCC_CLIENT;

	default:
		if (nCol > 0)
		{
			return (TDC_COLUMN)m_hdrColumns.GetItemData(nCol);
		}
		break;
	}

	ASSERT(0);
	return TDCC_NONE;
}


void CTDCTaskCtrlBase::SetCompletionStatus(const CString& sStatus) 
{ 
	if (sStatus != m_sCompletionStatus)
	{
		m_sCompletionStatus = sStatus; 

		if (IsColumnShowing(TDCC_STATUS))
			RedrawColumn(TDCC_STATUS);
	}
}

CString CTDCTaskCtrlBase::GetTaskColumnText(DWORD dwTaskID, 
	const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS, TDC_COLUMN nColID) const
{
	ASSERT(pTDS && pTDI && dwTaskID && (nColID != TDCC_NONE));

	if (!(pTDS && pTDI && dwTaskID && (nColID != TDCC_NONE)))
		return _T("");

	CString sTaskColText;
		
	switch (nColID)
	{
	case TDCC_CLIENT:
		sTaskColText = pTDI->sTitle;
		break;

	case TDCC_POSITION:
		sTaskColText = m_data.GetTaskPosition(pTDI, pTDS);
		break;

	case TDCC_PRIORITY:
		// priority color
		if (!HasStyle(TDCS_DONEHAVELOWESTPRIORITY) || !m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL))
		{
			int nPriority = m_data.GetHighestPriority(pTDI, pTDS, FALSE);
			BOOL bHasPriority = (nPriority != FM_NOPRIORITY);

			// draw priority number over the top
			if (bHasPriority && !HasStyle(TDCS_HIDEPRIORITYNUMBER))
				sTaskColText = Misc::Format(nPriority);
		}
		break;

	case TDCC_RISK:
		if (HasStyle(TDCS_INCLUDEDONEINRISKCALC) || !m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL))
		{
			int nRisk = m_data.GetHighestRisk(pTDI, pTDS);

			if (nRisk != FM_NORISK)
				sTaskColText = Misc::Format(nRisk);
		}
		break;

	case TDCC_RECURRENCE:
		sTaskColText = pTDI->trRecurrence.GetRegularity(FALSE);
		break;

	case TDCC_ID:
		// Figure out is this is really a reference
		if (pTDS->GetTaskID() != dwTaskID)
			sTaskColText.Format(_T("(%lu) %lu"), pTDS->GetTaskID(), dwTaskID);
		else
			sTaskColText.Format(_T("%lu"), dwTaskID);
		break;

	case TDCC_PARENTID:
		sTaskColText = Misc::Format(pTDS->GetParentTaskID());
		break;

	case TDCC_RECENTEDIT:
		if (pTDI->IsRecentlyEdited())
			sTaskColText = _T("*");
		break;

	case TDCC_COST:
		{
			double dCost = m_data.CalcCost(pTDI, pTDS);

			if (dCost != 0.0 || !HasStyle(TDCS_HIDEZEROTIMECOST))
				sTaskColText = Misc::Format(dCost, 2);
		}
		break;

	case TDCC_EXTERNALID:
		sTaskColText = pTDI->sExternalID;
		break;

	case TDCC_VERSION:
		sTaskColText = pTDI->sVersion;
		break;

	case TDCC_ALLOCTO:
		sTaskColText = m_data.FormatTaskAllocTo(pTDI);
		break;

	case TDCC_ALLOCBY:
		sTaskColText = pTDI->sAllocBy;
		break;

	case TDCC_STATUS:
		{
			sTaskColText = pTDI->sStatus;

			// if a task is completed and has no status and the completion status
			// has been specified then draw the completion status
			if (sTaskColText.IsEmpty() && !m_sCompletionStatus.IsEmpty() && m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL))
				sTaskColText = m_sCompletionStatus;
		}
		break;

	case TDCC_CATEGORY:
		sTaskColText = m_data.FormatTaskCategories(pTDI);
		break;

	case TDCC_TAGS:
		sTaskColText = m_data.FormatTaskTags(pTDI);
		break;

	case TDCC_CREATEDBY:
		sTaskColText = pTDI->sCreatedBy;
		break;

	case TDCC_PERCENT:
		{
			BOOL bDone = -1; // so we know when it's set
			BOOL bHideDone = HasStyle(TDCS_HIDEPERCENTFORDONETASKS);

			if (bHideDone)
				bDone = m_data.IsTaskDone(pTDI, pTDS, TDCCHECKALL);

			if (bHideDone && bDone)
				break; // nothing to do

			int nPercent = m_data.CalcPercentDone(pTDI, pTDS);

			if (!nPercent && HasStyle(TDCS_HIDEZEROPERCENTDONE))
				break; // nothing to do

			sTaskColText = Misc::Format(nPercent, _T("%"));
		}
		break;

	case TDCC_TIMEEST:
	case TDCC_TIMESPENT:
		{
			BOOL bTimeEst = (nColID == TDCC_TIMEEST);
			int nUnits = (bTimeEst ? m_nDefTimeEstUnits : m_nDefTimeSpentUnits); // good default value

			// get actual task time units
			if (!pTDS->HasSubTasks() || HasStyle(TDCS_ALLOWPARENTTIMETRACKING))
				nUnits = pTDI->GetTimeUnits(bTimeEst);

			// draw time
			double dTime = (bTimeEst ? m_data.CalcTimeEstimate(pTDI, pTDS, nUnits) :
				m_data.CalcTimeSpent(pTDI, pTDS, nUnits));

			// first handle zero times
			if ((dTime == 0.0) && HasStyle(TDCS_HIDEZEROTIMECOST))
			{
				break;
			}
			// then check for negative times
			else if (bTimeEst && (dTime < 0.0))
			{
				break;
			}

			// else
			int nDecPlaces = HasStyle(TDCS_ROUNDTIMEFRACTIONS) ? 0 : 2;

			if (HasStyle(TDCS_DISPLAYHMSTIMEFORMAT))
				sTaskColText = CTimeHelper().FormatTimeHMS(dTime, nUnits, (BOOL)nDecPlaces);
			else
				sTaskColText = CTimeHelper().FormatTime(dTime, nUnits, nDecPlaces);
		}
		break;

	case TDCC_PATH:
		if (pTDS->GetParentTaskID())
			sTaskColText = m_data.GetTaskPath(pTDI, pTDS);
		break;

	case TDCC_REMINDER:
		if (TaskHasReminder(dwTaskID))
			sTaskColText = BELLCHAR;
		break;

		// items having no text or rendered differently
	case TDCC_STARTDATE:
	case TDCC_DUEDATE:
	case TDCC_DONEDATE:
	case TDCC_CREATIONDATE:
	case TDCC_LASTMOD:
	case TDCC_ICON:
	case TDCC_DEPENDENCY:
	case TDCC_DONE:
	case TDCC_TRACKTIME:
	case TDCC_FLAG:
		break;

	case TDCC_REMAINING:
		{
			int nUnits = THU_HOURS;
			double dRemaining = m_data.CalcRemainingTime(pTDI, pTDS, nUnits);

			// format appropriately
			if (HasStyle(TDCS_CALCREMAININGTIMEBYPERCENT) || HasStyle(TDCS_CALCREMAININGTIMEBYSPENT))
			{
				sTaskColText = CTimeHelper().FormatTime(dRemaining, nUnits, 2);
			}
			else // TDCS_CALCREMAININGTIMEBYDUEDATE
			{
				COleDateTime date = m_data.CalcDueDate(pTDI, pTDS);

				if (CDateHelper::IsDateSet(date)) 
				{
					if (HasStyle(TDCS_DISPLAYHMSTIMEFORMAT))
					{
						sTaskColText = CTimeHelper().FormatTimeHMS(dRemaining, THU_DAYS, TRUE);
					}
					else
					{
						// find best units for display
						if (fabs(dRemaining) >= 1.0)
						{
							sTaskColText = CTimeHelper().FormatTime(dRemaining, THU_DAYS, 1);
						}
						else
						{
							dRemaining *= 24; // to hours

							if (fabs(dRemaining) >= 1.0)
							{
								sTaskColText = CTimeHelper().FormatTime(dRemaining, THU_HOURS, 1);
							}
							else
							{
								dRemaining *= 60; // to mins
								sTaskColText = CTimeHelper().FormatTime(dRemaining, THU_MINS, 0);
							}
						}
					}
				}
			}
		}
		break;

	case TDCC_FILEREF:
		sTaskColText = pTDI->GetFileRef(0);

		if (!sTaskColText.IsEmpty() && !WebMisc::IsURL(sTaskColText))
		{
			ASSERT(!m_sTasklistFolder.IsEmpty());
			FileMisc::MakeFullPath(sTaskColText, m_sTasklistFolder);
		}
		break;

	case TDCC_SUBTASKDONE:
		{
			int nSubtasksCount, nSubtasksDone;

			if (m_data.GetSubtaskTotals(pTDI, pTDS, nSubtasksCount, nSubtasksDone))
				sTaskColText.Format(_T("%d/%d"), nSubtasksDone, nSubtasksCount);
		}
		break;

	default:
		// handled during drawing
		ASSERT(CTDCCustomAttributeHelper::IsCustomColumn(nColID));
		break;
	}
	
	return sTaskColText;
}

CString CTDCTaskCtrlBase::GetTaskCustomColumnText(const TODOITEM* pTDI, const TODOSTRUCTURE* pTDS, TDC_COLUMN nColID) const
{
	if (!CTDCCustomAttributeHelper::IsCustomColumn(nColID))
		return _T("");

	TDCCUSTOMATTRIBUTEDEFINITION attribDef;
	VERIFY (CTDCCustomAttributeHelper::GetAttributeDef(nColID, m_aCustomAttribDefs, attribDef));
	
	TDCCADATA data(pTDI->GetCustomAttribValue(attribDef.sUniqueID));
	
	switch (attribDef.GetDataType())
	{
	case TDCCA_DATE:
	case TDCCA_ICON:
	case TDCCA_BOOL:
		break;
		
	case TDCCA_DOUBLE:
		{
			double dValue = data.AsDouble();
			m_data.CalcTaskCustomAttributeData(pTDI, pTDS, attribDef, dValue);

			return Misc::Format(dValue, 2);
		}
		break;

	case TDCCA_INTEGER:
		{
			// calculate as double to prevent overflow
			double dValue = data.AsDouble();
			m_data.CalcTaskCustomAttributeData(pTDI, pTDS, attribDef, dValue);

			return Misc::Format(dValue, 0);
		}
		break;
		
	default:
		if (!data.IsEmpty())
		{
			if (attribDef.GetListType() == TDCCA_FIXEDMULTILIST || attribDef.GetListType() == TDCCA_AUTOMULTILIST)
			{
				CStringArray aItems;
				
				if (data.AsArray(aItems))
					return Misc::FormatArray(aItems);
			}
		}
	}

	// else
	return data.AsString();
}

// message and notifications for 'us'
LRESULT CTDCTaskCtrlBase::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	if (!IsResyncEnabled())
		return CTreeListSyncer::WindowProc(hRealWnd, msg, wp, lp);
	
	ASSERT(hRealWnd == GetHwnd()); // 'us'

	switch (msg)
	{
	case WM_CONTEXTMENU:
		if ((HWND)wp == m_scLeft.GetHwnd() || 
			(HWND)wp == m_scRight.GetHwnd() ||
			(HWND)wp == m_hdrTasks.GetSafeHwnd() ||
			(HWND)wp == m_hdrColumns.GetSafeHwnd())
		{
			// pass on to parent
			::SendMessage(GetHwnd(), WM_CONTEXTMENU, (WPARAM)hRealWnd, lp);

			return 0L; // eat
		}
		break;
		
	case WM_SETFOCUS:
		::SetFocus(Tasks());
		break;

	case WM_LBUTTONDBLCLK:
		{
			CPoint ptCursor(lp);
			CRect rSplitter;
			
			if (GetSplitterRect(rSplitter) && rSplitter.PtInRect(ptCursor))
			{
				ResizeSplitterToFitColumns();
				return 0L; // eat
			}
		}
		break;
		
	case WM_NOTIFY:
		// only interested in notifications from the tree/list pair to their parent
		if (wp == m_scLeft.GetDlgCtrlID() || 
			wp == m_scRight.GetDlgCtrlID() ||
			wp == (UINT)m_hdrTasks.GetDlgCtrlID() ||
			wp == (UINT)m_hdrColumns.GetDlgCtrlID())
		{
			// let base class have its turn first
			LRESULT lr = CTreeListSyncer::WindowProc(hRealWnd, msg, wp, lp);
			
			// our extra handling
			LPNMHDR pNMHDR = (LPNMHDR)lp;
			HWND hwnd = pNMHDR->hwndFrom;
			
			switch (pNMHDR->code)
			{
			case HDN_ITEMCLICK:
				// column header clicks handled in ScWindowProc
				if (hwnd == m_hdrTasks)
				{
					NMHEADER* pHDN = (NMHEADER*)pNMHDR;
					
					if (pHDN->iButton == 0)
						OnHeaderClick(TDCC_CLIENT);

					return 0L; // eat
				}
				break;
				
			case NM_RCLICK:
				// headers don't generate their own WM_CONTEXTMENU
				if (hwnd == m_hdrTasks)
				{
					// pass on to parent
					::SendMessage(GetHwnd(), WM_CONTEXTMENU, (WPARAM)hwnd, (LPARAM)::GetMessagePos());
				}
				break;

			case NM_CLICK:
				if (hwnd == m_lcColumns)
				{
					const NMITEMACTIVATE* pNMIA = (NMITEMACTIVATE*)lp;

					// only handle if item is valid
					if (pNMIA->iItem != -1)
					{
						TDC_COLUMN nColID = GetColumnID(pNMIA->iSubItem);
						DWORD dwTaskID = GetColumnItemTaskID(pNMIA->iItem); // task ID
						
						if (ItemColumnSupportsClickHandling(pNMIA->iItem, nColID))
						{
							NotifyParentOfColumnEditClick(nColID, dwTaskID);
						}
					}
				}
				break;
			}
				
			return lr;
		}
		break;
	}
	
	return CTreeListSyncer::WindowProc(hRealWnd, msg, wp, lp);
}

int CTDCTaskCtrlBase::CalcSplitterPosToFitColumns() const
{
	int nFirst = m_hdrColumns.GetFirstVisibleItem();
	ASSERT(nFirst != -1);
	
	int nLast = m_hdrColumns.GetLastVisibleItem();
	ASSERT(nLast != -1);

	CRect rFirst, rLast;
	VERIFY(m_hdrColumns.GetItemRect(nFirst, rFirst) && m_hdrColumns.GetItemRect(nLast, rLast));
	
	CRect rClient;
	CWnd::GetClientRect(rClient);

	int nNewSplitPos = 0;
	int nColsWidth = ((rLast.right - rFirst.left) + 10/*GetSplitterWidth()*/);

	if (IsRight(m_lcColumns))
	{
		if (HasVScrollBar(m_lcColumns))
			nColsWidth += GetSystemMetrics(SM_CXVSCROLL);

		nNewSplitPos = max(MIN_TASKS_WIDTH, (rClient.right - nColsWidth));
	}
	else // cols on left
	{
		nNewSplitPos = min(nColsWidth, (rClient.right - MIN_TASKS_WIDTH));
	}

	return nNewSplitPos;
}

void CTDCTaskCtrlBase::ResizeSplitterToFitColumns()
{
	int nNewSplitPos = CalcSplitterPosToFitColumns();

	if (nNewSplitPos != GetSplitPos())
	{
		CRect rClient;
		CWnd::GetClientRect(rClient);
		
		CTreeListSyncer::Resize(rClient, nNewSplitPos);
	}
}

void CTDCTaskCtrlBase::RepackageAndSendToParent(UINT msg, WPARAM /*wp*/, LPARAM lp)
{
	// sanity check
	ASSERT(msg == WM_NOTIFY);

	if (msg == WM_NOTIFY)
	{
		NMHDR* pNMHDR = (NMHDR*)lp;
		NMHDR hdrOrg = *pNMHDR; // so we can restore after

		pNMHDR->hwndFrom = GetSafeHwnd();
		pNMHDR->idFrom = CWnd::GetDlgCtrlID();
		
		CWnd::GetParent()->SendMessage(msg, pNMHDR->idFrom, lp);

		// restore
		pNMHDR->hwndFrom = hdrOrg.hwndFrom;
		pNMHDR->idFrom = hdrOrg.idFrom;
	}
}

// messages and notifications sent to m_lcColumns 
LRESULT CTDCTaskCtrlBase::ScWindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	if (!IsResyncEnabled())
		return CTreeListSyncer::ScWindowProc(hRealWnd, msg, wp, lp);
	
	if (hRealWnd == Tasks())
	{
		switch (msg)
		{
		case WM_SETFOCUS:
			InvalidateAll(FALSE, TRUE);
			break;

		case WM_PRINT:
			if (!m_lcColumns.GetItemCount() && !m_sTasksWndPrompt.IsEmpty())
			{
				LRESULT lr = ScDefault(hRealWnd);
				CWndPrompt::DrawPrompt(hRealWnd, m_sTasksWndPrompt, (HDC)wp, TRUE);

				return lr;
			}
			break;
			
		case WM_PAINT:
			if (!m_lcColumns.GetItemCount() && !m_sTasksWndPrompt.IsEmpty())
			{
				LRESULT lr = ScDefault(hRealWnd);
				CWndPrompt::DrawPrompt(hRealWnd, m_sTasksWndPrompt, (HDC)wp, TRUE);

				return lr;
			}
			break;
			
		case WM_HSCROLL:
			// Windows only invalidates the item labels but
			// we need the whole row because we render the 
			// comments after the task text
			::InvalidateRect(Tasks(), NULL, FALSE);
			break;
		}
	}
	else if (hRealWnd == m_lcColumns)
	{
		switch (msg)
		{
		case WM_TIMER:
			// very weird horizontal scroll issue when LVS_EX_FULLROWSELECT
			// is enabled and the mouse is very slowly moved off a just
			// selected task. Research suggests that this relates to the
			// listview's handling of it's internal timer messages, and
			// it scrolls to the start because it thinks it's about to
			// start editing the label, even though LVS_EDITLABELS is not set
			if ((wp == 0x2B) && (lp == 0))
			{
				ASSERT(!CTreeListSyncer::HasStyle(hRealWnd, LVS_EDITLABELS, FALSE));
				return TRUE; // eat it
			}
			break;
			
		case WM_NOTIFY:
			{
				LPNMHDR pNMHDR = (LPNMHDR)lp;
				HWND hwnd = pNMHDR->hwndFrom;
				
				switch (pNMHDR->code)
				{
				case NM_RCLICK:
					// headers don't generate their own WM_CONTEXTMENU
					if (hwnd == m_hdrColumns)
					{
						// pass on to parent
						::SendMessage(GetHwnd(), WM_CONTEXTMENU, (WPARAM)hwnd, ::GetMessagePos());
					}
					break;

				case HDN_DIVIDERDBLCLICK:
					if (hwnd == m_hdrColumns)
					{
						// resize just that column
						int nItem = ((NMHEADER*)pNMHDR)->iItem;
						
						if (m_hdrColumns.IsItemVisible(nItem))
						{
							CDC* pDC = m_lcColumns.GetDC();
							CFont* pOldFont = GraphicsMisc::PrepareDCFont(pDC, m_lcColumns);
							
							int nColWidth = RecalcColumnWidth(nItem, pDC);
							m_hdrColumns.SetItemWidth(nItem, nColWidth);
							
							pDC->SelectObject(pOldFont);
							m_lcColumns.ReleaseDC(pDC);
						}
						return 0L;
					}

				case HDN_ITEMCLICK:
					if (hwnd == m_hdrColumns)
					{
						NMHEADER* pHDN = (NMHEADER*)pNMHDR;

						// forward on to our parent
						if ((pHDN->iButton == 0) && m_hdrColumns.IsItemVisible(pHDN->iItem))
						{
							OnHeaderClick((TDC_COLUMN)m_hdrColumns.GetItemData(pHDN->iItem));
						}
						return 0L;
					}
					break;

				case HDN_ITEMCHANGING:
					if (hwnd == m_hdrColumns)
					{
						NMHEADER* pHDN = (NMHEADER*)pNMHDR;
						
						// don't let user drag column too narrow
						if ((pHDN->iButton == 0) && (pHDN->pitem->mask & HDI_WIDTH))
						{
							TDC_COLUMN nColID = GetColumnID(pHDN->iItem);

							if (IsColumnShowing(nColID))
							{
								pHDN->pitem->cxy = max(MIN_COL_WIDTH, pHDN->pitem->cxy);
							}
						}
					}
					break;
				}
			}
			break;

		case WM_SETFOCUS:
			InvalidateAll();
			break;
			
 		case WM_ERASEBKGND:
			if (COSVersion() == OSV_LINUX)
			{
				CRect rClient;
				m_lcColumns.GetClientRect(rClient);

				CDC::FromHandle((HDC)wp)->FillSolidRect(rClient, GetSysColor(COLOR_WINDOW));
			}
			return TRUE;

		case WM_KEYUP:
			switch (wp)
			{
			case VK_NEXT:  
			case VK_DOWN:
			case VK_UP:
			case VK_PRIOR: 
			case VK_HOME:
			case VK_END: 
			case VK_SHIFT: 
			case VK_CONTROL: 
				// force a parent notification
				NotifyParentSelChange(SC_BYKEYBOARD);
				break;
			}
			break;
			
		case WM_LBUTTONDBLCLK:
		case WM_RBUTTONDBLCLK:
		case WM_RBUTTONDOWN:
			{
				// let parent handle any focus changes first
				m_lcColumns.SetFocus();

				// don't let the selection to be set to -1
				// when clicking below the last item
				if (m_lcColumns.HitTest(lp) == -1)
				{
					CPoint pt(lp);
					::ClientToScreen(hRealWnd, &pt);

					if (!::DragDetect(m_lcColumns, pt)) // we don't want to disable drag selecting
					{
						TRACE(_T("Ate Listview ButtonDown\n"));
						return 0; // eat it
					}
				}
			}
			break;
				
		case WM_LBUTTONDOWN:
			if (HandleListLBtnDown(m_lcColumns, lp))
				return 0L; // eat it
			break;
		}
	}
	
	return CTreeListSyncer::ScWindowProc(hRealWnd, msg, wp, lp);
}

BOOL CTDCTaskCtrlBase::HandleListLBtnDown(CListCtrl& lc, CPoint pt)
{
	m_bBoundSelecting = FALSE;

	// let parent handle any focus changes first
	lc.SetFocus();

	int nHit = -1;
	TDC_COLUMN nColID = TDCC_NONE;

	// if the user clicked on a column that allows direct input
	// AND multi items are selected and the item clicked is 
	// already selected then we generate a NM_CLICK and eat the 
	// message to prevent a selection change
	if (lc == m_lcColumns)
	{
		DWORD dwTaskID = 0;
		nHit = HitTestColumnsItem(pt, TRUE, nColID, &dwTaskID);

		BOOL bMultiSelection = (m_lcColumns.GetSelectedCount() > 1);
		BOOL bTaskSelected = IsListItemSelected(m_lcColumns, nHit);

		if (bMultiSelection && bTaskSelected && ItemColumnSupportsClickHandling(nHit, nColID))
		{
			NotifyParentOfColumnEditClick(nColID, dwTaskID);

			TRACE(_T("Ate Listview LButtonDown\n"));
			return TRUE; // eat it
		}
	}
	else
	{
		nHit = lc.HitTest(pt);
	}

	// De-selecting a lot of items can be slow because 
	// OnListSelectionChange is called once for each.
	// So we try to detect big changes here and handle 
	// them ourselves.
	lc.ClientToScreen(&pt);

	if (nHit != -1)
	{
		BOOL bHitSelected = IsListItemSelected(m_lcColumns, nHit);

		if (Misc::ModKeysArePressed(0) && 
			(!bHitSelected || !ItemColumnSupportsClickHandling(nHit, nColID)))
		{
			DeselectAll();
		}
	}
	else if (::DragDetect(lc, pt))
	{
		m_bBoundSelecting = -1;

		if (!Misc::IsKeyPressed(VK_CONTROL))
		{
			DeselectAll();
		}

		// there's no reliable to way to detect the end of a
		// bounding-box selection especially if the mouse 
		// cursor ends up outside the window so we use a timer
		SetTimer(TIMER_BOUNDINGSEL, 50, NULL);
	}
	else // prevent deselection
	{
		TRACE(_T("Ate Listview ButtonDown\n"));
		return TRUE; // eat it
	}

	return FALSE;
}

void CTDCTaskCtrlBase::OnTimer(UINT nIDEvent)
{
	switch (nIDEvent)
	{
	case TIMER_BOUNDINGSEL:
		if (m_bBoundSelecting && !Misc::IsKeyPressed(VK_LBUTTON))
		{
			m_bBoundSelecting = FALSE;

			KillTimer(TIMER_BOUNDINGSEL);
			NotifyParentSelChange(SC_BYMOUSE);
		}
		break;
	}

	CWnd::OnTimer(nIDEvent);
}

void CTDCTaskCtrlBase::NotifyParentOfColumnEditClick(TDC_COLUMN nColID, DWORD dwTaskID)
{
	CWnd::GetParent()->SendMessage(WM_TDCN_COLUMNEDITCLICK, nColID, dwTaskID);
}

void CTDCTaskCtrlBase::OnHeaderClick(TDC_COLUMN nColID)
{
	TDC_COLUMN nPrev = m_sort.single.nBy;
	TDC_COLUMN nSortBy = TDCC_NONE;
	
	// check for default attribute
	TDCCOLUMN* pTDCC = GetColumn(nColID);
	
	// could also be a custom column
	if (pTDCC)
	{
		nSortBy = nColID;
	}	
	else if (CTDCCustomAttributeHelper::IsColumnSortable(nColID, m_aCustomAttribDefs))
	{
		nSortBy = nColID;
	}
	
	// do the sort
	if (nSortBy != TDCC_NONE)
	{
		Sort(nSortBy);
		
		// notify parent
		CWnd::GetParent()->SendMessage(WM_TDCN_SORT, CWnd::GetDlgCtrlID(), MAKELPARAM((WORD)nPrev, (WORD)nSortBy));
	}
}

BOOL CTDCTaskCtrlBase::ItemColumnSupportsClickHandling(int nItem, TDC_COLUMN nColID) const
{
	if ((nItem == -1) || !Misc::ModKeysArePressed(0))
		return FALSE;

	DWORD dwTaskID = GetColumnItemTaskID(nItem);
	ASSERT(dwTaskID);

	BOOL bSupported = FALSE;
	BOOL bNoModifiers = Misc::ModKeysArePressed(0);
	BOOL bSingleSelection = (GetSelectedCount() == 1);
	BOOL bTaskSelected = IsListItemSelected(m_lcColumns, nItem);

	switch (nColID)
	{
	case TDCC_DONE:
	case TDCC_FLAG:
	case TDCC_ICON:
	case TDCC_RECURRENCE:
		bSupported = !IsReadOnly();
		break;
		
	case TDCC_REMINDER:
		bSupported = !m_data.IsTaskDone(dwTaskID);
		break;

	case TDCC_FILEREF:
		bSupported = m_data.TaskHasFileRef(dwTaskID);
		break;
			
	case TDCC_DEPENDENCY:
		bSupported = m_data.IsTaskDependent(dwTaskID);
		break;
			
	case TDCC_TRACKTIME:
		// check tasklist is editable, task is trackable and 
		// neither the ctrl not shift keys are pressed (ie => multiple selection)
		// and either the task is not selected or it's only singly selected
		bSupported = (bNoModifiers && !IsReadOnly() && 
					m_data.IsTaskTimeTrackable(dwTaskID) &&
					(!bTaskSelected || bSingleSelection));
		break;

	default: // try custom columns
		if (!IsReadOnly() && CTDCCustomAttributeHelper::IsCustomColumn(nColID))
		{
			TDCCUSTOMATTRIBUTEDEFINITION attribDef;
			
			if (CTDCCustomAttributeHelper::GetAttributeDef(nColID, m_aCustomAttribDefs, attribDef))
			{
				switch (attribDef.GetDataType())
				{
				case TDCCA_BOOL:
					bSupported = TRUE;
					break;
					
				case TDCCA_ICON:
					switch (attribDef.GetListType())
					{
					case TDCCA_FIXEDLIST:
					case TDCCA_NOTALIST:
						bSupported = TRUE;
						break;
					}
					break;
					
					default:
						// do item cycling for fixed lists unless they support calculation
						bSupported = ((attribDef.GetListType() == TDCCA_FIXEDLIST) && !attribDef.SupportsCalculation());
						break;
				}
			}
		}
		break;
	}

	return bSupported;
}

void CTDCTaskCtrlBase::SetModified(TDC_ATTRIBUTE nAttrib)
{
	// Recalc or redraw columns as required
	BOOL bRedrawCols = FALSE, bRedrawTasks = ModCausesTaskTextColorChange(nAttrib);
	
	TDC_COLUMN nColID = TDC::MapAttributeToColumn(nAttrib);
	TDC_COLUMN nRecalcColID = TDCC_NONE;
	
	switch (nAttrib)
	{
	case TDCA_CREATIONDATE:
		// this can only be modified for new tasks via the commandline
		// so nothing needs to be done
		break;
		
	case TDCA_DONEDATE:
		AccumulateRecalcColumn(TDCC_DONEDATE, nRecalcColID);
		AccumulateRecalcColumn(TDCC_DUEDATE, nRecalcColID);
		AccumulateRecalcColumn(TDCC_DONE, nRecalcColID);
		
		if (HasStyle(TDCS_USEPERCENTDONEINTIMEEST))
			AccumulateRecalcColumn(TDCC_TIMEEST, nRecalcColID);
		
		if (!m_sCompletionStatus.IsEmpty())
			AccumulateRecalcColumn(TDCC_STATUS, nRecalcColID);
		
		break;
		
	case TDCA_DUEDATE:
		if (!AccumulateRecalcColumn(TDCC_DUEDATE, nRecalcColID))
			bRedrawCols = IsColumnShowing(TDCC_PRIORITY);
		
		// don't redraw while month cal control is visible
		// 			if (m_dateDue.GetMonthCalCtrl() == NULL)
		// 				m_taskTree.InvalidateAll(); // text color can be affected 
		break;
		
	case TDCA_PRIORITY:
		if (!bRedrawTasks)
		{
			bRedrawCols = IsColumnShowing(TDCC_PRIORITY);
			
			if (!bRedrawCols && HasStyle(TDCS_SHOWPERCENTASPROGRESSBAR))
				bRedrawCols = IsColumnShowing(TDCC_PERCENT);
		}
		break;
		
	case TDCA_ALLOCTO:
	case TDCA_ALLOCBY:
	case TDCA_STATUS:
	case TDCA_VERSION:
	case TDCA_CATEGORY:
	case TDCA_TAGS:
	case TDCA_COST:
	case TDCA_STARTDATE:
	case TDCA_EXTERNALID:
	case TDCA_RECURRENCE:
		AccumulateRecalcColumn(nColID, nRecalcColID);
		break;
		
	case TDCA_TIMEEST:
		if (!AccumulateRecalcColumn(TDCC_TIMEEST, nRecalcColID))
			bRedrawCols = HasStyle(TDCS_AUTOCALCPERCENTDONE);
		break;
		
	case TDCA_TIMESPENT:
		if (!AccumulateRecalcColumn(TDCC_TIMESPENT, nRecalcColID))
			bRedrawCols = HasStyle(TDCS_AUTOCALCPERCENTDONE);
		break;
		
	case TDCA_DEPENDENCY:
	case TDCA_RISK:
	case TDCA_FILEREF:
	case TDCA_FLAG:
	case TDCA_PERCENT:
		if (!bRedrawTasks)
			bRedrawCols = IsColumnShowing(nColID);
		break;
		
	case TDCA_ICON:
		if (IsColumnShowing(TDCC_ICON))
			bRedrawCols = TRUE;
		else
			bRedrawTasks = TRUE;
		break;
		
	case TDCA_TASKNAME:
		bRedrawCols = IsColumnShowing(TDCC_PATH);
		break;
		
	case TDCA_PROJNAME:
	case TDCA_COMMENTS:
	case TDCA_ENCRYPT:
	case TDCA_COLOR:
		break;
		
	case TDCA_NONE:
	case TDCA_PASTE:
	case TDCA_POSITION: // == move
	case TDCA_DELETE:
	case TDCA_ARCHIVE:
	case TDCA_UNDO:
	case TDCA_NEWTASK:
	case TDCA_CUSTOMATTRIB:
	case TDCA_CUSTOMATTRIBDEFS:
		nRecalcColID = TDCC_ALL;
		break;
		
	case TDCA_TASKNAMEORCOMMENTS:
	case TDCA_ANYTEXTATTRIBUTE:
	default:
		ASSERT(0);
		break;
	}
		
	RecalcColumnWidth(nRecalcColID);
	
	if (bRedrawTasks)
	{
		InvalidateAll();
	}
	else if (bRedrawCols || (nRecalcColID != TDCC_NONE))
	{
		m_lcColumns.Invalidate();
	}
}

BOOL CTDCTaskCtrlBase::ModCausesTaskTextColorChange(TDC_ATTRIBUTE nModType) const
{
	switch (nModType)
	{
	case TDCA_COLOR:
		return !HasStyle(TDCS_COLORTEXTBYPRIORITY) &&
				!HasStyle(TDCS_COLORTEXTBYATTRIBUTE) &&
				!HasStyle(TDCS_COLORTEXTBYNONE);

	case TDCA_CATEGORY:
	case TDCA_ALLOCBY:
	case TDCA_ALLOCTO:
	case TDCA_STATUS:
	case TDCA_VERSION:
	case TDCA_EXTERNALID:
	case TDCA_TAGS:
		return (HasStyle(TDCS_COLORTEXTBYATTRIBUTE) && (GetColorByAttribute() == nModType));

	case TDCA_DONEDATE:
		return (GetCompletedTaskColor() != CLR_NONE);

	case TDCA_DUEDATE:
		{
			COLORREF crDue, crDueToday;
			GetDueTaskColors(crDue, crDueToday);

			return ((crDue != CLR_NONE) || (crDueToday != CLR_NONE));
		}

	case TDCA_PRIORITY:
		return HasStyle(TDCS_COLORTEXTBYPRIORITY);
	}

	// all else
	return FALSE;
}

BOOL CTDCTaskCtrlBase::AccumulateRecalcColumn(TDC_COLUMN nColID, TDC_COLUMN& nRecalcColID) const
{
	// optimisation
	if (nRecalcColID == TDCC_ALL)
	{
		return TRUE;
	}
	else if (IsColumnShowing(nColID))
	{
		if (nRecalcColID == TDCC_NONE)
			nRecalcColID = nColID;
		else
			nRecalcColID = TDCC_ALL;
		
		return TRUE;
	}
	
	// else
	return FALSE;
}

BOOL CTDCTaskCtrlBase::ModNeedsResort(TDC_ATTRIBUTE nModType) const
{
	if (!m_sort.IsSorting())
		return FALSE;

	return ModNeedsResort(nModType, m_sort);
}

BOOL CTDCTaskCtrlBase::ModNeedsResort(TDC_ATTRIBUTE nModType, const TDSORT& sort) const
{
	if (!HasStyle(TDCS_RESORTONMODIFY))
		return FALSE;

	BOOL bNeedSort = FALSE;
	
	if (sort.bMulti)
	{
		bNeedSort = (ModNeedsResort(nModType, sort.multi.col1.nBy) ||
					ModNeedsResort(nModType, sort.multi.col2.nBy) ||
					ModNeedsResort(nModType, sort.multi.col3.nBy));
	}
	else
	{
		bNeedSort = ModNeedsResort(nModType, sort.single.nBy);
	}

	return bNeedSort;
}

BOOL CTDCTaskCtrlBase::ModNeedsResort(TDC_ATTRIBUTE nModType, TDC_COLUMN nSortBy) const
{
	if (nSortBy == TDCC_NONE)
		return FALSE;

	TDC_COLUMN nModCol = TDC::MapAttributeToColumn(nModType);

	switch (nModType)
	{
	case TDCA_TASKNAME:		
	case TDCA_STARTDATE:	
	case TDCA_PRIORITY:		
	case TDCA_ALLOCTO:		
	case TDCA_ALLOCBY:		
	case TDCA_STATUS:		
	case TDCA_CATEGORY:		
	case TDCA_TAGS:			
	case TDCA_TIMEEST:		
	case TDCA_RISK:			
	case TDCA_EXTERNALID:	
	case TDCA_VERSION:		
	case TDCA_RECURRENCE:	
	case TDCA_ICON:			
	case TDCA_COLOR:		
	case TDCA_FLAG:			
	case TDCA_DUEDATE:
	case TDCA_PERCENT:
	case TDCA_TIMESPENT:
	case TDCA_DEPENDENCY:
	case TDCA_COST:
		{
			ASSERT(nModCol != TDCC_NONE);

			if (nModCol == nSortBy)
				return TRUE;
		}
		break;

	case TDCA_DONEDATE:
		{
			ASSERT(nModCol != TDCC_NONE);

			if (nModCol == nSortBy)
				return TRUE;

			if (HasStyle(TDCS_SORTDONETASKSATBOTTOM))
			{
				// some sort columns are unaffected by completed tasks
				switch (nSortBy)
				{
				case TDCC_ID:
					return FALSE;
				}

				// all else
				return TRUE;
			}
		}
		break;

	case TDCA_NEWTASK:
	case TDCA_POSITION:
	case TDCA_UNDO:
	case TDCA_PASTE:
		ASSERT(nModCol == TDCC_NONE);
		return TRUE;

	case TDCA_DELETE:
	case TDCA_ARCHIVE:
	case TDCA_ENCRYPT:
	case TDCA_NONE:
	case TDCA_PROJNAME:
	case TDCA_COMMENTS:
		ASSERT(nModCol == TDCC_NONE);
		return FALSE;

	case TDCA_CUSTOMATTRIBDEFS:	// Resort all custom columns
		ASSERT(nModCol == TDCC_NONE);
		return CTDCCustomAttributeHelper::IsCustomColumn(nSortBy);

	default:
		if (CTDCCustomAttributeHelper::IsCustomAttribute(nModType))
		{
			ASSERT(nModCol != TDCC_NONE);
			ASSERT (CTDCCustomAttributeHelper::IsColumnSortable(nSortBy, m_aCustomAttribDefs));

			return (nModCol == nSortBy);
		}
		// else unhandled attribute
		ASSERT(0);
		return FALSE;
	}

	ASSERT(nModCol != TDCC_NONE);

	// Attribute interdependencies
	switch (nSortBy)
	{
	case TDCC_DONE:
	case TDCC_STARTDATE:
	case TDCC_DUEDATE:
	case TDCC_PERCENT:
		return (nModType == TDCA_DONEDATE);
		
	case TDCC_RISK:
		return ((nModType == TDCA_DONEDATE) && HasStyle(TDCS_DONEHAVELOWESTRISK));
		
	case TDCC_PRIORITY:
		return (((nModType == TDCA_DONEDATE) && HasStyle(TDCS_DONEHAVELOWESTPRIORITY)) ||
				((nModType == TDCA_DUEDATE) && HasStyle(TDCS_DUEHAVEHIGHESTPRIORITY)));
		
	case TDCC_REMAINING: 
		return (((nModType == TDCA_DUEDATE) && HasStyle(TDCS_CALCREMAININGTIMEBYDUEDATE)) ||
				((nModType == TDCA_TIMESPENT) && HasStyle(TDCS_CALCREMAININGTIMEBYSPENT)) ||
				((nModType == TDCA_PERCENT) && HasStyle(TDCS_CALCREMAININGTIMEBYPERCENT)));
	}

	// all else
	return FALSE;
}

void CTDCTaskCtrlBase::SetAlternateLineColor(COLORREF crAltLine)
{
	SetColor(m_crAltLine, crAltLine);
}

void CTDCTaskCtrlBase::SetGridlineColor(COLORREF crGridLine)
{
	SetColor(m_crGridLine, crGridLine);
}

void CTDCTaskCtrlBase::SetColor(COLORREF& color, COLORREF crNew)
{
	if (IsHooked() && (crNew != color))
		InvalidateAll(FALSE);
	
	color = crNew;
}

void CTDCTaskCtrlBase::RedrawColumns(BOOL bErase) const
{
	::InvalidateRect(m_lcColumns, NULL, bErase);
	::UpdateWindow(m_lcColumns);
}

void CTDCTaskCtrlBase::RedrawColumn(TDC_COLUMN nColID) const
{
	CRect rCol, rClient;
	m_lcColumns.GetClientRect(rClient);

	int nCol = m_hdrColumns.FindItem(nColID);
	m_hdrColumns.GetItemRect(nCol, rCol);

	rCol.top = rClient.top;
	rCol.bottom = rClient.bottom;

	::InvalidateRect(m_lcColumns, rCol, FALSE);
}

void CTDCTaskCtrlBase::RedrawTasks(BOOL bErase) const
{
	::InvalidateRect(Tasks(), NULL, bErase);
	::UpdateWindow(Tasks());
}

int CTDCTaskCtrlBase::CalcMaxDateColWidth(TDC_DATE nDate, CDC* pDC) const
{
	// basic date string
	COleDateTime date(2000, 12, 31, 23, 59, 0);
	DWORD dwFlags = 0;
	
	if (HasStyle(TDCS_SHOWDATESINISO))
		dwFlags |= DHFD_ISO;
	
	// time component
	if (WantDrawColumnTime(nDate))
		dwFlags |= DHFD_TIME | DHFD_NOSEC;
	
	CString sDate = CDateHelper::FormatDate(date, dwFlags);
	int nWidth = pDC->GetTextExtent(sDate).cx;
	
	// add longest day of week if required
	if (HasStyle(TDCS_SHOWWEEKDAYINDATES))
	{
		nWidth += pDC->GetTextExtent(" ").cx; // spacer
		nWidth += CDateHelper::CalcLongestWeekdayName(pDC, TRUE); // short day
	}
	
	return nWidth;
}

BOOL CTDCTaskCtrlBase::WantDrawColumnTime(TDC_DATE nDate) const
{
	switch (nDate)
	{
	case TDCD_START:		
	case TDCD_STARTDATE:	
	case TDCD_STARTTIME:
		return IsColumnShowing(TDCC_STARTTIME);
		
	case TDCD_DUE:		
	case TDCD_DUEDATE:	
	case TDCD_DUETIME:	
		return IsColumnShowing(TDCC_DUETIME);
		
	case TDCD_DONE:		
	case TDCD_DONEDATE:	
	case TDCD_DONETIME:	
		return IsColumnShowing(TDCC_DONETIME);
		
	case TDCD_LASTMOD:
		return TRUE; // always
		
	case TDCD_CUSTOM:
	case TDCD_CREATE:
		return FALSE; // never
	}
	
	// all else
	ASSERT(0);
	return FALSE;
}

void CTDCTaskCtrlBase::ClearSortColumn()
{
	if ((m_nSortColID != TDCC_NONE) && GetSafeHwnd())
	{
		::InvalidateRect(m_hdrTasks, NULL, FALSE);
		m_hdrColumns.Invalidate(FALSE);
	}

	m_nSortColID = TDCC_NONE;
	m_nSortDir = TDC_SORTNONE;
}

void CTDCTaskCtrlBase::SetSortColumn(TDC_COLUMN nColID, TDC_SORTDIR nSortDir)
{
	m_nSortDir = nSortDir;

	if (m_nSortColID != nColID)
	{
		if ((m_nSortColID == TDCC_CLIENT) || (nColID == TDCC_CLIENT))
			m_hdrTasks.Invalidate(FALSE);

		m_nSortColID = nColID;
	}
}

TDC_COLUMN CTDCTaskCtrlBase::GetSortColumn(TDC_SORTDIR& nSortDir) const
{
	nSortDir = m_nSortDir;
	return m_nSortColID;
}

int CTDCTaskCtrlBase::RecalcColumnWidth(int nCol, CDC* pDC, BOOL bVisibleOnly) const
{
	TDC_COLUMN nColID = (TDC_COLUMN)m_hdrColumns.GetItemData(nCol);

	// handle hidden columns
	if (!IsColumnShowing(nColID))
 		return 0;
	
	int nColWidth = 0;
	
	switch (nColID)
	{
	case TDCC_TRACKTIME:
	case TDCC_FLAG:
	case TDCC_RECENTEDIT:
	case TDCC_DEPENDENCY:
	case TDCC_ICON:
	case TDCC_DONE:
	case TDCC_REMINDER:
		// use MINCOLWIDTH
		break; 
		
	case TDCC_ID:
		{
			DWORD dwRefID = m_find.GetLargestReferenceID();

			if (dwRefID)
				nColWidth = GraphicsMisc::GetTextWidth(pDC, _T("%u (%u)"), m_dwNextUniqueTaskID - 1, dwRefID);
			else
				nColWidth = GraphicsMisc::GetTextWidth(pDC, _T("%u"), m_dwNextUniqueTaskID - 1);
		}
		break; 

	case TDCC_PARENTID:
		nColWidth = GraphicsMisc::GetTextWidth(pDC, _T("%u"), m_dwNextUniqueTaskID - 1);
		break; 

	case TDCC_POSITION:
	case TDCC_RECURRENCE:
	case TDCC_EXTERNALID:
	case TDCC_VERSION:
	case TDCC_ALLOCBY:
	case TDCC_CREATEDBY:
	case TDCC_COST:
	case TDCC_PATH:
	case TDCC_SUBTASKDONE:
	case TDCC_ALLOCTO:
	case TDCC_CATEGORY:
	case TDCC_TAGS:
		{
			TDC_ATTRIBUTE nAttrib = TDC::MapColumnToAttribute(nColID);
			ASSERT(nAttrib != TDCA_NONE);
			
			// determine the longest visible string
			CString sLongest = m_find.GetLongestValue(nAttrib, bVisibleOnly);
			nColWidth = GraphicsMisc::GetAverageMaxStringWidth(sLongest, pDC);
		}
		break;
		
	case TDCC_STATUS:
		{
			// determine the longest visible string
			CString sLongest = m_find.GetLongestValue(TDCA_STATUS, m_sCompletionStatus, bVisibleOnly);
			nColWidth = GraphicsMisc::GetAverageMaxStringWidth(sLongest, pDC);
		}
		break;
		
	case TDCC_PRIORITY:
		if (!HasStyle(TDCS_HIDEPRIORITYNUMBER))
			nColWidth = pDC->GetTextExtent("10").cx;
		break; 
		
	case TDCC_RISK:
		nColWidth = pDC->GetTextExtent("10").cx;
		break; 
		
	case TDCC_REMAINING:
		{
			int nWidth1 = pDC->GetTextExtent("12m4w").cx;
			int nWidth2 = pDC->GetTextExtent("100.0W").cx;

			nColWidth = max(nWidth1, nWidth2);
		}
		break; 
		
	case TDCC_FILEREF:
		if (HasStyle(TDCS_SHOWNONFILEREFSASTEXT))
			nColWidth = 60; 
		break; 
		
	case TDCC_TIMEEST:
	case TDCC_TIMESPENT:
		if (HasStyle(TDCS_DISPLAYHMSTIMEFORMAT))
		{
			nColWidth = pDC->GetTextExtent("12m4w").cx;
		}
		else
		{
			BOOL bTimeEst = (nColID == TDCC_TIMEEST);
			CString sLongest = m_find.GetLongestTime(m_nDefTimeEstUnits, bTimeEst);

			nColWidth = GraphicsMisc::GetAverageMaxStringWidth(sLongest, pDC) + 4; // add a bit to handle different time unit widths
		}
		break;
		
	case TDCC_PERCENT:
		nColWidth = pDC->GetTextExtent("100%").cx;
		break;
		
	case TDCC_LASTMOD:
	case TDCC_DUEDATE:
	case TDCC_CREATIONDATE:
	case TDCC_STARTDATE:
	case TDCC_DONEDATE:
		nColWidth = CalcMaxDateColWidth(TDC::MapColumnToDate(nColID), pDC);
		break;

	default:
		// custom columns
		if (CTDCCustomAttributeHelper::IsCustomColumn(nColID))
		{
			// determine the longest visible string depending on type
			TDCCUSTOMATTRIBUTEDEFINITION attribDef;

			if (CTDCCustomAttributeHelper::GetAttributeDef(nColID, m_aCustomAttribDefs, attribDef))
			{
				if (!attribDef.bEnabled)
				{
					return 0; // hidden
				}
				else 
				{
					switch (attribDef.GetDataType())
					{
					case TDCCA_DATE:
						nColWidth = CalcMaxDateColWidth(TDCD_CUSTOM, pDC);
						break;

					case TDCCA_ICON:
						if (attribDef.IsList())
						{
							switch (attribDef.GetListType())
							{
							case TDCCA_FIXEDLIST:
								nColWidth = CTDCCustomAttributeHelper::CalcLongestListItem(attribDef, pDC);
								break;

							case TDCCA_FIXEDMULTILIST:
								nColWidth = (attribDef.aDefaultListData.GetSize() * 18);
								break;
							}
						}
						// else single icon, no text: use MINCOLWIDTH
						break;

					case TDCCA_BOOL:
						if (attribDef.sColumnTitle.GetLength() == 1)
						{
							nColWidth = GraphicsMisc::GetTextWidth(pDC, attribDef.sColumnTitle);
						}
						else 
						{
							nColWidth = GraphicsMisc::GetTextWidth(pDC, _T("+"));
						}
						break;

					case TDCCA_DOUBLE:
						{
							// numerals are always the same width so we don't need average width
							CString sLongest = m_find.GetLongestCustomDoubleAttribute(attribDef.sUniqueID, bVisibleOnly);
							nColWidth = GraphicsMisc::GetTextWidth(pDC, sLongest);
						}
						break;

					case TDCCA_INTEGER:
						{
							// numerals are always the same width so we don't need average width
							CString sLongest = m_find.GetLongestCustomAttribute(attribDef.sUniqueID, bVisibleOnly);
							nColWidth = GraphicsMisc::GetTextWidth(pDC, sLongest); // 
						}
						break;

					default:
						{
							CString sLongest = m_find.GetLongestCustomAttribute(attribDef.sUniqueID, bVisibleOnly);
							nColWidth = GraphicsMisc::GetAverageMaxStringWidth(sLongest, pDC);
						}
						break;
					}
				}
			}
			else
			{
				return 0; // hidden
			}
		}
		else
		{
			ASSERT (0);
		}
		break;
	}

	if (nColWidth == 0)
		nColWidth = MIN_RESIZE_WIDTH;
	else
		nColWidth = max((nColWidth + (2 * LV_COLPADDING)), MIN_RESIZE_WIDTH);
	
	// take max of this and column title
	int nTitleWidth = (m_hdrColumns.GetItemTextWidth(nCol, pDC) + (2 * HD_COLPADDING));

	return max(nTitleWidth, nColWidth);
}

BOOL CTDCTaskCtrlBase::SelectionHasIncompleteDependencies(CString& sIncomplete) const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		if (TaskHasIncompleteDependencies(dwTaskID, sIncomplete))
			return TRUE;
	}

	return FALSE;
}

BOOL CTDCTaskCtrlBase::TaskHasIncompleteDependencies(DWORD dwTaskID, CString& sIncomplete) const
{
	CStringArray aDepends;
	int nNumDepends = m_data.GetTaskDependencies(dwTaskID, aDepends);
	
	for (int nDepends = 0; nDepends < nNumDepends; nDepends++)
	{
		CString sFile;
		DWORD dwDependID;
		
		ParseTaskLink(aDepends[nDepends], dwDependID, sFile);
		
		// see if dependent is one of 'our' tasks
		if (dwDependID && sFile.IsEmpty())
		{
			if (!m_data.IsTaskDone(dwDependID))
			{
				sIncomplete = aDepends[nDepends];
				return TRUE;
			}
		}
		else if (!sFile.IsEmpty()) // pass to parent if we can't handle
		{
			BOOL bDependentIsDone = CWnd::GetParent()->SendMessage(WM_TDCM_TASKISDONE, dwDependID, (LPARAM)(LPCTSTR)sFile);
			
			if (!bDependentIsDone)
			{
				sIncomplete = aDepends[nDepends];
				return TRUE;
			}
		}
	}
	
	// check this tasks subtasks
	const TODOSTRUCTURE* pTDS = m_data.LocateTask(dwTaskID);
	ASSERT(pTDS);
	
	if (pTDS && pTDS->HasSubTasks())
	{
		int nPos = pTDS->GetSubTaskCount();
		
		while (nPos--)
		{
			if (TaskHasIncompleteDependencies(pTDS->GetSubTaskID(nPos), sIncomplete))
				return TRUE;
		}
	}
	
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SelectionHasDates(TDC_DATE nDate, BOOL bAll) const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	// traverse the selected items looking for the first one
	// who has a due date or the first that doesn't
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		double dDate = m_data.GetTaskDate(dwTaskID, nDate);

		if (!bAll && dDate > 0)
		{
			return TRUE;
		}
		else if (bAll && dDate == 0)
		{
			return FALSE;
		}
	}

	return bAll; // there were no dissenting tasks
}

BOOL CTDCTaskCtrlBase::SelectionHasIcons() const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		if (!m_data.GetTaskIcon(dwTaskID).IsEmpty())
			return TRUE;
	}

	return FALSE; // there were no tasks with icons
}

BOOL CTDCTaskCtrlBase::SelectionAreAllDone() const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		if (!m_data.IsTaskDone(dwTaskID))
			return FALSE;
	}

	return TRUE; // all tasks were completed
}

int CTDCTaskCtrlBase::SelectionHasCircularDependencies(CDWordArray& aTaskIDs) const
{
	aTaskIDs.RemoveAll();

	POSITION pos = GetFirstSelectedTaskPos();

	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		if (m_data.TaskHasLocalCircularDependencies(dwTaskID))
		{
			aTaskIDs.Add(dwTaskID);
			return TRUE;
		}
	}

	return aTaskIDs.GetSize();
}

BOOL CTDCTaskCtrlBase::SelectionHasDependents() const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		if (m_data.TaskHasDependents(dwTaskID))
			return TRUE;
	}
	
	// no refs
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SelectionHasReferences() const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		if (m_data.IsTaskReference(dwTaskID))
			return TRUE;
	}
	
	// no refs
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SelectionHasNonReferences() const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		if (!m_data.IsTaskReference(dwTaskID))
			return TRUE;
	}
	
	// no non-refs
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SelectionHasIncompleteSubtasks(BOOL bExcludeRecurring) const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		if (m_data.TaskHasIncompleteSubtasks(dwTaskID, bExcludeRecurring))
			return TRUE;
	}
	
	return FALSE;
}

BOOL CTDCTaskCtrlBase::SelectionHasSubtasks() const
{
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		const TODOSTRUCTURE* pTDS = m_data.LocateTask(dwTaskID);
		ASSERT(pTDS);

		if (pTDS && pTDS->GetSubTaskCount())
			return TRUE;
	}
	
	return FALSE;
}

BOOL CTDCTaskCtrlBase::ParseTaskLink(const CString& sLink, DWORD& dwTaskID, CString& sFile)
{
	CString sCleaned(sLink);
	
	// strip off protocol (if not done)
	int nProtocol = sCleaned.Find(TDL_PROTOCOL);
	
	if (nProtocol != -1)
		sCleaned = sCleaned.Mid(nProtocol + lstrlen(TDL_PROTOCOL));
	
	// cleanup
	sCleaned.Replace(_T("%20"), _T(" "));
	sCleaned.Replace(_T("/"), _T("\\"));
	
	// parse the url
	return TODOITEM::ParseTaskLink(sCleaned, dwTaskID, sFile);
}

void CTDCTaskCtrlBase::EnableExtendedSelection(BOOL bCtrl, BOOL bShift)
{
	if (bCtrl)
		s_nExtendedSelection |= HOTKEYF_CONTROL;
	else
		s_nExtendedSelection &= ~HOTKEYF_CONTROL;
	
	if (bShift)
		s_nExtendedSelection |= HOTKEYF_SHIFT;
	else
		s_nExtendedSelection &= ~HOTKEYF_SHIFT;
}

BOOL CTDCTaskCtrlBase::IsReservedShortcut(DWORD dwShortcut)
{
	// check this is not a reserved shortcut used by the tree or a.n.other ctrl
	switch (dwShortcut)
	{
	case MAKELONG(VK_UP, HOTKEYF_EXT):
	case MAKELONG(VK_PRIOR, HOTKEYF_EXT):
	case MAKELONG(VK_DOWN, HOTKEYF_EXT):
	case MAKELONG(VK_NEXT, HOTKEYF_EXT):
		
	case MAKELONG(VK_SPACE, HOTKEYF_CONTROL):
	case MAKELONG(VK_DELETE, HOTKEYF_CONTROL | HOTKEYF_EXT):
		return TRUE;
		
	case MAKELONG(VK_UP, HOTKEYF_CONTROL | HOTKEYF_EXT):
	case MAKELONG(VK_PRIOR, HOTKEYF_CONTROL | HOTKEYF_EXT):
	case MAKELONG(VK_DOWN, HOTKEYF_CONTROL | HOTKEYF_EXT):
	case MAKELONG(VK_NEXT, HOTKEYF_CONTROL | HOTKEYF_EXT):
		return (s_nExtendedSelection & HOTKEYF_CONTROL);
		
	case MAKELONG(VK_UP, HOTKEYF_SHIFT | HOTKEYF_EXT):
	case MAKELONG(VK_PRIOR, HOTKEYF_SHIFT | HOTKEYF_EXT):
	case MAKELONG(VK_DOWN, HOTKEYF_SHIFT | HOTKEYF_EXT):
	case MAKELONG(VK_NEXT, HOTKEYF_SHIFT | HOTKEYF_EXT):
		return (s_nExtendedSelection & HOTKEYF_SHIFT);
		
	case MAKELONG(VK_UP, HOTKEYF_SHIFT | HOTKEYF_CONTROL | HOTKEYF_EXT):
	case MAKELONG(VK_PRIOR, HOTKEYF_SHIFT | HOTKEYF_CONTROL | HOTKEYF_EXT):
	case MAKELONG(VK_DOWN, HOTKEYF_SHIFT | HOTKEYF_CONTROL | HOTKEYF_EXT):
	case MAKELONG(VK_NEXT, HOTKEYF_SHIFT | HOTKEYF_CONTROL | HOTKEYF_EXT):
		return (s_nExtendedSelection & (HOTKEYF_CONTROL | HOTKEYF_SHIFT)); // both
	}
	
	// all else
	return FALSE;
}

CString CTDCTaskCtrlBase::GetSelectedTaskComments() const
{
	if (GetSelectedCount() == 1)
		return m_data.GetTaskComments(GetSelectedTaskID());
	
	// else
	return _T("");
}

const CBinaryData& CTDCTaskCtrlBase::GetSelectedTaskCustomComments(CString& sCommentsTypeID) const
{
	if (GetSelectedCount() == 1)
		return m_data.GetTaskCustomComments(GetSelectedTaskID(), sCommentsTypeID);
	
	// else
	sCommentsTypeID.Empty();
    POSITION pos = GetFirstSelectedTaskPos();
	
    while (pos)
    {
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		const TODOITEM* pTDI = m_data.GetTask(dwTaskID);
		ASSERT (pTDI);
		
		if (pTDI)
		{
			if (sCommentsTypeID.IsEmpty())
			{
				sCommentsTypeID = pTDI->sCommentsTypeID;
			}
			else if (sCommentsTypeID != pTDI->sCommentsTypeID)
			{
				sCommentsTypeID.Empty();
				break;
			}
		}
	}
	
	static CBinaryData content;
	return content;
}

CString CTDCTaskCtrlBase::GetSelectedTaskTitle() const
{
	if (GetSelectedCount() == 1)
		return m_data.GetTaskTitle(GetSelectedTaskID());
	
	// else
	return _T("");
}

int CTDCTaskCtrlBase::GetSelectedTaskPriority() const
{
   int nPriority = -1;
   POSITION pos = GetFirstSelectedTaskPos();

   while (pos)
   {
      DWORD dwTaskID = GetNextSelectedTaskID(pos);
      int nTaskPriority = m_data.GetTaskPriority(dwTaskID);

      if (nPriority == -1)
         nPriority = nTaskPriority;

      else if (nPriority != nTaskPriority)
         return -1;
   }
	
	return nPriority;
}

DWORD CTDCTaskCtrlBase::GetSelectedTaskParentID() const
{
	DWORD dwParentID = 0;
	POSITION pos = GetFirstSelectedTaskPos();
	
	while (pos)
	{
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		DWORD dwTaskParentID = m_data.GetTaskParentID(dwTaskID);
		
		if (dwParentID == -1)
			dwParentID = dwTaskParentID;
		
		else if (dwParentID != dwTaskParentID)
			return 0;
	}
	
	return dwParentID;
}

int CTDCTaskCtrlBase::GetSelectedTaskRisk() const
{
   int nRisk = -1;
   POSITION pos = GetFirstSelectedTaskPos();

   while (pos)
   {
      DWORD dwTaskID = GetNextSelectedTaskID(pos);
      int nTaskRisk = m_data.GetTaskRisk(dwTaskID);

      if (nRisk == -1)
         nRisk = nTaskRisk;

      else if (nRisk != nTaskRisk)
         return -1; // == various
   }
	
	return nRisk;
}

CString CTDCTaskCtrlBase::GetSelectedTaskIcon() const
{
	CString sIcon;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		sIcon = m_data.GetTaskIcon(dwTaskID);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			CString sTaskIcon = m_data.GetTaskIcon(dwTaskID);
			
			if (sIcon != sTaskIcon)
				return _T("");
		}
	}
	
	return sIcon;
}

BOOL CTDCTaskCtrlBase::SelectedTaskHasDate(TDC_DATE nDate) const
{
	return CDateHelper::IsDateSet(GetSelectedTaskDate(nDate));
}

COleDateTime CTDCTaskCtrlBase::GetSelectedTaskDate(TDC_DATE nDate) const
{
	COleDateTime date; // == 0
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		date = m_data.GetTaskDate(dwTaskID, nDate);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			COleDateTime dateTask = m_data.GetTaskDate(dwTaskID, nDate);
			
			// first check if both dates are not set
			if (!CDateHelper::IsDateSet(date) && !CDateHelper::IsDateSet(dateTask))
				continue;

			if (!CDateHelper::IsDateSet(date)) // means dateTask.m_dt != 0
				return dateTask;

			// else 
			return date;
		}
		// if we get here then no dates were set
	}
	
	return date;
}

int CTDCTaskCtrlBase::IsSelectedTaskFlagged() const
{
	return m_data.IsTaskFlagged(GetSelectedTaskID());
}

BOOL CTDCTaskCtrlBase::IsSelectedTaskReference() const
{
	return m_data.IsTaskReference(GetSelectedTaskID());
}

double CTDCTaskCtrlBase::GetSelectedTaskTimeEstimate(TCHAR& nUnits) const
{
	double dTime = 0.0;
	nUnits = m_nDefTimeEstUnits;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		dTime = m_data.GetTaskTimeEstimate(dwTaskID, nUnits);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			
			TCHAR nTaskUnits;
			double dTaskTime = m_data.GetTaskTimeEstimate(dwTaskID, nTaskUnits);
			
			if (dTime != dTaskTime || nUnits != nTaskUnits)
				return 0.0;
		}
	}
	
	return dTime;
}

double CTDCTaskCtrlBase::GetSelectedTaskTimeSpent(TCHAR& nUnits) const
{
	double dTime = 0.0;
	nUnits = m_nDefTimeSpentUnits;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		dTime = m_data.GetTaskTimeSpent(dwTaskID, nUnits);
		
		while (pos)
		{
			DWORD dwTaskID = GetNextSelectedTaskID(pos);
			
			TCHAR nTaskUnits;
			double dTaskTime = m_data.GetTaskTimeSpent(dwTaskID, nTaskUnits);
			
			if (dTime != dTaskTime || nUnits != nTaskUnits)
				return 0.0;
		}
	}
	
	return dTime;
}

COLORREF CTDCTaskCtrlBase::GetSelectedTaskColor() const
{
	return m_data.GetTaskColor(GetSelectedTaskID());
}

BOOL CTDCTaskCtrlBase::GetSelectedTaskRecurrence(TDIRECURRENCE& tr) const
{
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		m_data.GetTaskRecurrence(dwTaskID, tr);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			
			TDIRECURRENCE trTask;
			m_data.GetTaskRecurrence(dwTaskID, trTask);
			
			if (tr != trTask)
				tr = TDIRECURRENCE();
		}
	}
	
	return (tr.nRegularity != TDIR_ONCE);
}

int CTDCTaskCtrlBase::GetSelectedTaskPercent() const
{
	int nPercent = 0;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		nPercent = m_data.GetTaskPercent(dwTaskID, FALSE);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			int nTaskPercent = m_data.GetTaskPercent(dwTaskID, FALSE);
			
			if (nPercent != nTaskPercent)
				return -1;
		}
	}
	
	return nPercent;
}

CString CTDCTaskCtrlBase::GetSelectedTaskPath(BOOL bIncludeTaskName, int nMaxLen) const
{
	CString sPath;

	if (GetSelectedCount() == 1)
	{
		DWORD dwTaskID = GetSelectedTaskID();
		CString sTaskTitle = bIncludeTaskName ? m_data.GetTaskTitle(dwTaskID) : _T("");

		if (bIncludeTaskName && nMaxLen != -1)
			nMaxLen -= sTaskTitle.GetLength();

		sPath = m_data.GetTaskPath(dwTaskID, nMaxLen);
	
		if (bIncludeTaskName)
			sPath += sTaskTitle;
	}

	return sPath;
}

double CTDCTaskCtrlBase::GetSelectedTaskCost() const
{
	double dCost = 0.0;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		dCost = m_data.GetTaskCost(dwTaskID);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			double dTaskCost = m_data.GetTaskCost(dwTaskID);

			if (dCost != dTaskCost)
				return 0.0;
		}
	}
	
	return dCost;
}

CString CTDCTaskCtrlBase::GetSelectedTaskCustomAttributeData(const CString& sAttribID, BOOL bFormatted) const
{
	CString sData;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		sData = m_data.GetTaskCustomAttributeData(dwTaskID, sAttribID);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			CString sTaskData = m_data.GetTaskCustomAttributeData(dwTaskID, sAttribID);

			if (sData != sTaskData)
				return _T("");
		}

		if (bFormatted)
			sData = CTDCCustomAttributeHelper::FormatData(sData, sAttribID, m_aCustomAttribDefs);
	}
	
	return sData;
}

int CTDCTaskCtrlBase::GetSelectedTaskAllocTo(CStringArray& aAllocTo) const
{
	aAllocTo.RemoveAll();

	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		m_data.GetTaskAllocTo(dwTaskID, aAllocTo);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);

			CStringArray aTaskAllocTo;
			m_data.GetTaskAllocTo(dwTaskID, aTaskAllocTo);
			
			if (!Misc::MatchAll(aAllocTo, aTaskAllocTo))
			{
				aAllocTo.RemoveAll();
				break;
			}
		}
	}
	
	return aAllocTo.GetSize();
}

CString CTDCTaskCtrlBase::GetSelectedTaskAllocBy() const
{
	CString sAllocBy;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		sAllocBy = m_data.GetTaskAllocBy(dwTaskID);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			CString sTaskAllocBy = m_data.GetTaskAllocBy(dwTaskID);
			
			if (sAllocBy != sTaskAllocBy)
				return _T("");
		}
	}
	
	return sAllocBy;
}

CString CTDCTaskCtrlBase::GetSelectedTaskVersion() const
{
	CString sVersion;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		sVersion = m_data.GetTaskVersion(dwTaskID);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			CString sTaskVersion = m_data.GetTaskVersion(dwTaskID);
			
			if (sVersion != sTaskVersion)
				return _T("");
		}
	}
	
	return sVersion;
}

CString CTDCTaskCtrlBase::GetSelectedTaskStatus() const
{
	CString sStatus;
	
	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);

		sStatus = m_data.GetTaskStatus(dwTaskID);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);
			CString sTaskStatus = m_data.GetTaskStatus(dwTaskID);
			
			if (sStatus != sTaskStatus)
				return _T("");
		}
	}
	
	return sStatus;
}

int CTDCTaskCtrlBase::GetSelectedTaskCategories(CStringArray& aCats) const
{
	aCats.RemoveAll();

	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		m_data.GetTaskCategories(dwTaskID, aCats);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);

			CStringArray aTaskCats;
			m_data.GetTaskCategories(dwTaskID, aTaskCats);
			
			if (!Misc::MatchAll(aCats, aTaskCats))
			{
				aCats.RemoveAll();
				break;
			}
		}
	}
	
	return aCats.GetSize();
}

int CTDCTaskCtrlBase::GetSelectedTaskTags(CStringArray& aTags) const
{
	aTags.RemoveAll();

	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		m_data.GetTaskTags(dwTaskID, aTags);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);

			CStringArray aTaskTags;
			m_data.GetTaskTags(dwTaskID, aTaskTags);
			
			if (!Misc::MatchAll(aTags, aTaskTags))
			{
				aTags.RemoveAll();
				break;
			}
		}
	}
	
	return aTags.GetSize();
}

int CTDCTaskCtrlBase::GetSelectedTaskDependencies(CStringArray& aDepends) const
{
	aDepends.RemoveAll();

	if (GetSelectedCount())
	{
		// get first item's value as initial
		POSITION pos = GetFirstSelectedTaskPos();
		DWORD dwTaskID = GetNextSelectedTaskID(pos);
		
		m_data.GetTaskDependencies(dwTaskID, aDepends);
		
		while (pos)
		{
			dwTaskID = GetNextSelectedTaskID(pos);

			CStringArray aTaskaDepends;
			m_data.GetTaskDependencies(dwTaskID, aTaskaDepends);
			
			if (!Misc::MatchAll(aDepends, aTaskaDepends))
			{
				aDepends.RemoveAll();
				break;
			}
		}
	}
	
	return aDepends.GetSize();
}

CString CTDCTaskCtrlBase::GetSelectedTaskFileRef(int nFile) const
{
	if (GetSelectedCount() == 1)
		return m_data.GetTaskFileRef(GetSelectedTaskID(), nFile);
	
	// else
	return _T("");
}

int CTDCTaskCtrlBase::GetSelectedTaskFileRefs(CStringArray& aFiles) const
{
	if (GetSelectedCount() == 1)
		return m_data.GetTaskFileRefs(GetSelectedTaskID(), aFiles);
	
	// else
	return 0;
}

int CTDCTaskCtrlBase::GetSelectedTaskFileRefCount() const
{
	if (GetSelectedCount() == 1)
		return m_data.GetTaskFileRefCount(GetSelectedTaskID());
	
	// else
	return 0;
}

CString CTDCTaskCtrlBase::GetSelectedTaskExtID() const
{
	if (GetSelectedCount() == 1)
		return m_data.GetTaskExtID(GetSelectedTaskID());
	
	// else
	return _T("");
}

BOOL CTDCTaskCtrlBase::CanSplitSelectedTask() const
{
	if (IsReadOnly())
		return FALSE;
	
	if (SelectionHasReferences())
		return FALSE;
	
	int nSelCount = GetSelectedCount();
	
	if (nSelCount == 1)
	{
		if (IsSelectedTaskDone() || SelectionHasSubtasks())
			return FALSE;
	}
	
	return (nSelCount > 0);
}

BOOL CTDCTaskCtrlBase::IsSelectedTaskDone() const
{
	return m_data.IsTaskDone(GetSelectedTaskID());
}

BOOL CTDCTaskCtrlBase::IsSelectedTaskDue() const
{
	return m_data.IsTaskDue(GetSelectedTaskID());
}

BOOL CTDCTaskCtrlBase::InitCheckboxImageList()
{
	if (m_ilCheckboxes.GetSafeHandle())
		return TRUE;
	
	const int nStates[] = { -1, CBS_UNCHECKEDNORMAL, CBS_CHECKEDNORMAL };//, CBS_MIXEDNORMAL };
	const int nNumStates = sizeof(nStates) / sizeof(int);
	
	CThemed th;
	
	if (th.Open(this, _T("BUTTON")) && th.AreControlsThemed())
	{
		th.BuildImageList(m_ilCheckboxes, BP_CHECKBOX, nStates, nNumStates);
	}
	
	// unthemed + fallback
	if (!m_ilCheckboxes.GetSafeHandle())
	{
		CBitmap bitmap;
		bitmap.LoadBitmap(IDB_CHECKBOXES);
		
		BITMAP BM;
		bitmap.GetBitmap(&BM);
		
		if (m_ilCheckboxes.Create(BM.bmWidth / nNumStates, BM.bmHeight, ILC_COLOR32 | ILC_MASK, 0, 1))
			m_ilCheckboxes.Add(&bitmap, 255);
	}
	
	return (NULL != m_ilCheckboxes.GetSafeHandle());
}

void CTDCTaskCtrlBase::PrepareInfoTip(HWND /*hwndTooltip*/)
{	
//	ASSERT(hwndTooltip);

// 	::SendMessage(hwndTooltip, TTM_SETDELAYTIME, 30000, MAKELPARAM(TTDT_AUTOPOP, 0));
// 	::SendMessage(hwndTooltip, TTM_SETDELAYTIME, 1000,  MAKELPARAM(TTDT_INITIAL, 0));
// 	::SendMessage(hwndTooltip, TTM_SETDELAYTIME, 100,   MAKELPARAM(TTDT_RESHOW, 0));
}

CString CTDCTaskCtrlBase::FormatInfoTip(DWORD dwTaskID) const
{
	const TODOITEM* pTDI = m_data.GetTask(dwTaskID);
	ASSERT(pTDI);
	
	CString sTip, sItem;
	
	if (pTDI)
	{
		// format text multilined
		sItem.Format(_T("%s %s"), CEnString(IDS_TDCTIP_TASK), pTDI->sTitle);
		sTip += sItem;
		
		CString sComments = pTDI->sComments;
		int nLen = sComments.GetLength();
		
		if (nLen && m_nMaxInfotipCommentsLength != 0)
		{
			if ((m_nMaxInfotipCommentsLength > 0) && (nLen > m_nMaxInfotipCommentsLength))
				sComments = (sComments.Left(m_nMaxInfotipCommentsLength) + _T("..."));
			
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_COMMENTS), sComments);
			sTip += sItem;
		}
		
		if (IsColumnShowing(TDCC_STATUS) && !pTDI->sStatus.IsEmpty())
		{
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_STATUS), pTDI->sStatus);
			sTip += sItem;
		}
		
		if (IsColumnShowing(TDCC_CATEGORY) && pTDI->aCategories.GetSize())
		{
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_CATEGORY), m_data.FormatTaskCategories(pTDI));
			sTip += sItem;
		}
		
		if (IsColumnShowing(TDCC_TAGS) && pTDI->aTags.GetSize())
		{
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_TAGS), m_data.FormatTaskTags(pTDI));
			sTip += sItem;
		}
		
		if (IsColumnShowing(TDCC_ALLOCTO) && pTDI->aAllocTo.GetSize())
		{
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_ALLOCTO), m_data.FormatTaskAllocTo(pTDI));
			sTip += sItem;
		}
		
		if (IsColumnShowing(TDCC_ALLOCBY) && !pTDI->sAllocBy.IsEmpty())
		{
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_ALLOCBY), pTDI->sAllocBy); 
			sTip += sItem;
		}
		
		if (IsColumnShowing(TDCC_VERSION) && !pTDI->sVersion.IsEmpty())
		{
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_VERSION), pTDI->sVersion); 
			sTip += sItem;
		}
		
		DWORD dwDateFmt = HasStyle(TDCS_SHOWDATESINISO) ? DHFD_ISO : 0;
		
		if (pTDI->IsDone())
		{
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_DONEDATE), CDateHelper::FormatDate(pTDI->dateDone, dwDateFmt));
			sTip += sItem;
		}
		else
		{
			if (IsColumnShowing(TDCC_PRIORITY) && pTDI->nPriority != FM_NOPRIORITY)
			{
				sItem.Format(_T("\n%s %d"), CEnString(IDS_TDCTIP_PRIORITY), pTDI->nPriority);
				sTip += sItem;
			}
			
			if (IsColumnShowing(TDCC_RISK) && pTDI->nRisk != FM_NORISK)
			{
				sItem.Format(_T("\n%s %d"), CEnString(IDS_TDCTIP_RISK), pTDI->nRisk);
				sTip += sItem;
			}
			
			if (IsColumnShowing(TDCC_STARTDATE) && pTDI->HasStart())
			{
				sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_STARTDATE), CDateHelper::FormatDate(pTDI->dateStart, dwDateFmt));
				sTip += sItem;
			}
			
			if (IsColumnShowing(TDCC_DUEDATE) && pTDI->HasDue())
			{
				if (pTDI->HasDueTime())
					dwDateFmt |= DHFD_TIME | DHFD_NOSEC;
				
				sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_DUEDATE), CDateHelper::FormatDate(pTDI->dateDue, dwDateFmt));
				sTip += sItem;
			}
			
			if (IsColumnShowing(TDCC_PERCENT))
			{
				sItem.Format(_T("\n%s %d"), CEnString(IDS_TDCTIP_PERCENT), m_data.CalcPercentDone(dwTaskID));
				sTip += sItem;
			}
			
			if (IsColumnShowing(TDCC_TIMEEST))
			{
				sItem.Format(_T("\n%s %.2f %c"), CEnString(IDS_TDCTIP_TIMEEST), m_data.CalcTimeEstimate(dwTaskID, TDCU_HOURS), CTimeHelper::GetUnits(THU_HOURS));
				sTip += sItem;
			}
			
			if (IsColumnShowing(TDCC_TIMESPENT))
			{
				sItem.Format(_T("\n%s %.2f %c"), CEnString(IDS_TDCTIP_TIMESPENT), m_data.CalcTimeSpent(dwTaskID, TDCU_HOURS), CTimeHelper::GetUnits(THU_HOURS));
				sTip += sItem;
			}
		}
		
		if (IsColumnShowing(TDCC_COST))
		{
			sItem.Format(_T("\n%s %.2f"), CEnString(IDS_TDCTIP_COST), m_data.CalcCost(dwTaskID));
			sTip += sItem;
		}
		
		if (IsColumnShowing(TDCC_DEPENDENCY))
		{
			if (pTDI->aDependencies.GetSize())
			{
				sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_DEPENDS), Misc::FormatArray(pTDI->aDependencies));
				sTip += sItem;
			}
			
			CDWordArray aDependents;
			
			if (m_data.GetTaskLocalDependents(dwTaskID, aDependents))
			{
				sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_DEPENDENTS), Misc::FormatArray(aDependents));
				sTip += sItem;
			}
		}
		
		if (IsColumnShowing(TDCC_FILEREF) && pTDI->aFileRefs.GetSize())
		{
			sItem.Format(_T("\n%s %s"), CEnString(IDS_TDCTIP_FILEREF), pTDI->aFileRefs[0]);
			sTip += sItem;
		}
		
		if (pTDI->dateLastMod > 0)
		{
			sItem.Format(_T("\n%s %s (%s)"), CEnString(IDS_TDCTIP_LASTMOD), pTDI->dateLastMod.Format(VAR_DATEVALUEONLY), pTDI->dateLastMod.Format(VAR_TIMEVALUEONLY));
			sTip += sItem;
		}
	}
	
	return sTip;
}
