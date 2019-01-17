// GanttStruct.cpp: implementation of the CGanttStruct class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "GanttStruct.h"

#include "..\shared\DateHelper.h"
#include "..\shared\graphicsMisc.h"
#include "..\shared\misc.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

int GANTTDEPENDENCY::STUB = 0;

GANTTDEPENDENCY::GANTTDEPENDENCY() : ptFrom(0), ptTo(0), dwFromID(0), dwToID(0)
{
}

void GANTTDEPENDENCY::SetFrom(const CPoint& pt, DWORD dwTaskID)
{
	dwFromID = dwTaskID;
	ptFrom = pt;
}

void GANTTDEPENDENCY::SetTo(const CPoint& pt, DWORD dwTaskID)
{
	dwToID = dwTaskID;
	ptTo = pt;
}

DWORD GANTTDEPENDENCY::GetFromID() const
{
	ASSERT(dwFromID);
	return dwFromID;
}

DWORD GANTTDEPENDENCY::GetToID() const
{
	ASSERT(dwToID);
	return dwToID;
}

BOOL GANTTDEPENDENCY::Matches(DWORD dwFrom, DWORD dwTo) const
{
	ASSERT(dwFromID && dwToID);
	return ((dwFromID == dwFrom) && (dwToID == dwTo));
}

BOOL GANTTDEPENDENCY::Draw(CDC* pDC, const CRect& rClient, BOOL bDragging)
{
	if (!HitTest(rClient))
		return FALSE;
	
	// draw 3x3 box at ptTo
	if (!bDragging)
	{
		CRect rBox(ptTo.x - 1, ptTo.y - 1, ptTo.x + 2, ptTo.y + 2);
		pDC->FillSolidRect(rBox, 0);
	}
		
	CPoint pts[3];
	CalcDependencyPath(pts);
	
	int nOldROP2 = pDC->SetROP2(bDragging ? R2_NOT : R2_BLACK);
	pDC->Polyline(pts, 3);

	DrawDependencyArrow(pDC, pts[0]);
	pDC->SetROP2(nOldROP2);
	
	return TRUE;
}

void GANTTDEPENDENCY::DrawDependencyArrow(CDC* pDC, const CPoint& pt) const
{
	CPoint pts[3], ptArrow(pt);

	CalcDependencyArrow(ptArrow, pts);
	pDC->Polyline(pts, 3);
	
	// offset and draw again
	if (IsFromAboveTo())
		ptArrow.y++;
	else
		ptArrow.y--;
	
	CalcDependencyArrow(ptArrow, pts);
	pDC->Polyline(pts, 3);
}

BOOL GANTTDEPENDENCY::HitTest(const CRect& rect) const
{
	CRect rThis;
	return (CalcBoundingRect(rThis) && CRect().IntersectRect(rect, rThis));
}

BOOL GANTTDEPENDENCY::HitTest(const CPoint& point, int nTol) const
{
	CRect rThis;
	
	if (!CalcBoundingRect(rThis))
		return FALSE;

	// add tolerance
	rThis.InflateRect(nTol, nTol);

	if (!rThis.PtInRect(point))
		return FALSE;

	// check each line segment
	CPoint pts[3];
	CalcDependencyPath(pts);

	nTol = max(nTol, 1);
	
	for (int i = 0; i < 2; i++)
	{
		CRect rSeg;

		rSeg.left	= min(pts[i].x, pts[i+1].x) - nTol;
		rSeg.right	= max(pts[i].x, pts[i+1].x) + nTol;
		rSeg.top	= min(pts[i].y, pts[i+1].y) - nTol;
		rSeg.bottom = max(pts[i].y, pts[i+1].y) + nTol;

		if (rSeg.PtInRect(point))
			return TRUE;
	}

	// no hit
	return FALSE;
}

BOOL GANTTDEPENDENCY::IsFromAboveTo() const
{
	return (ptFrom.y < ptTo.y);
}

void GANTTDEPENDENCY::CalcDependencyPath(CPoint pts[3]) const
{
	CPoint ptTemp(ptFrom);

	// first point
	if (IsFromAboveTo())
		ptTemp.y -= (2 - STUB);
	else
		ptTemp.y += (1 - STUB);

	pts[0] = ptTemp;

	// mid point
	ptTemp.y = ptTo.y;
	pts[1] = ptTemp;

	// last point
	pts[2] = ptTo;
}

void GANTTDEPENDENCY::CalcDependencyArrow(const CPoint& pt, CPoint pts[3]) const
{
	pts[0] = pts[1] = pts[2] = pt;

	const int ARROW = (STUB / 2);
	
	if (IsFromAboveTo())
	{
		pts[0].Offset(-ARROW, ARROW);
		pts[2].Offset(ARROW+1, ARROW+1);
	}
	else
	{
		pts[0].Offset(-ARROW, -ARROW);
		pts[2].Offset(ARROW+1, -(ARROW+1));
	}
}

BOOL GANTTDEPENDENCY::CalcBoundingRect(CRect& rect) const
{
	if (ptFrom == ptTo)
		return FALSE;
	
	// allow for stub overhang
	rect.left	= min(ptFrom.x, ptTo.x) - STUB;
	rect.right	= max(ptFrom.x, ptTo.x) + STUB;
	rect.top	= min(ptFrom.y, ptTo.y);
	rect.bottom = max(ptFrom.y, ptTo.y);
	
	return TRUE;
}

//////////////////////////////////////////////////////////////////////

CString GANTTITEM::MILESTONE_TAG;

GANTTITEM::GANTTITEM() 
	: 
	crText(CLR_NONE), 
	bParent(FALSE), 
	dwRefID(0), 
	dwOrgRefID(0), 
	bGoodAsDone(FALSE)
{

}

GANTTITEM::GANTTITEM(const GANTTITEM& gi)
{
	*this = gi;
}

GANTTITEM& GANTTITEM::operator=(const GANTTITEM& gi)
{
	sTitle = gi.sTitle;
	dtStart = gi.dtStart;
	dtMinStart = gi.dtMinStart;
	dtDue = gi.dtDue;
	dtMaxDue = gi.dtMaxDue;
	dtDone = gi.dtDone;
	crText = gi.crText;
	sAllocTo = gi.sAllocTo;
	bParent = gi.bParent;
	dwRefID = gi.dwRefID;
	nPercent = gi.nPercent;
	bGoodAsDone = gi.bGoodAsDone;
	
	aTags.Copy(gi.aTags);
	aDepends.Copy(gi.aDepends);
	
	return (*this);
}

BOOL GANTTITEM::operator==(const GANTTITEM& gi)
{
	return ((sTitle == gi.sTitle) &&
			(dtStart == gi.dtStart) &&
			(dtMinStart == gi.dtMinStart) &&
			(dtDue == gi.dtDue) &&
			(dtMaxDue == gi.dtMaxDue) &&
			(dtDone == gi.dtDone) &&
			(crText == gi.crText) &&
			(sAllocTo == gi.sAllocTo) &&
			(bParent == gi.bParent) &&
			(dwRefID == gi.dwRefID) &&
			(nPercent == gi.nPercent) &&	
			(bGoodAsDone == gi.bGoodAsDone) &&
			Misc::MatchAll(aTags, gi.aTags) &&
			Misc::MatchAll(aDepends, gi.aDepends));
}

GANTTITEM::~GANTTITEM()
{
	
}

void GANTTITEM::MinMaxDates(const GANTTITEM& giOther)
{
	if (giOther.bParent)
	{
		CDateHelper::Max(dtMaxDue, giOther.dtMaxDue);
		CDateHelper::Min(dtMinStart, giOther.dtMinStart);
	}
	else // leaf task
	{
		CDateHelper::Max(dtMaxDue, giOther.dtDue);
		CDateHelper::Min(dtMinStart, giOther.dtStart);
	}
}

BOOL GANTTITEM::IsDone() const
{
	return (bGoodAsDone || CDateHelper::IsDateSet(dtDone));
}

BOOL GANTTITEM::HasStart() const
{
	return CDateHelper::IsDateSet(dtStart);
}

BOOL GANTTITEM::HasDue() const
{
	return CDateHelper::IsDateSet(dtDue);
}

BOOL GANTTITEM::IsMilestone() const
{
	if (MILESTONE_TAG.IsEmpty() || (aTags.GetSize() == 0))
		return FALSE;

	if (!bParent && !CDateHelper::IsDateSet(dtDue))
		return FALSE;
	
	if (bParent && !CDateHelper::IsDateSet(dtMaxDue))
		return FALSE;

	// else
	return (Misc::Find(aTags, MILESTONE_TAG, FALSE, FALSE) != -1);
}

COLORREF GANTTITEM::GetDefaultFillColor() const
{
	if (IsDone())
		return GetSysColor(COLOR_3DLIGHT);

	if (HasTextColor() && (crText != 0))
		return crText;

	// else
	return GetSysColor(COLOR_WINDOW);
}

BOOL GANTTITEM::HasTextColor() const
{
	return ((crText != CLR_NONE) && (crText != GetSysColor(COLOR_WINDOWTEXT)));
}

COLORREF GANTTITEM::GetDefaultBorderColor() const
{
	COLORREF crDefFill = GetDefaultFillColor();

	if (crDefFill == GetSysColor(COLOR_WINDOW))
		return 0;

	// else
	return GraphicsMisc::Darker(crDefFill, 0.5);
}

//////////////////////////////////////////////////////////////////////

CGanttItemMap::~CGanttItemMap()
{
	RemoveAll();
}

void CGanttItemMap::RemoveAll()
{
	DWORD dwTaskID = 0;
	GANTTITEM* pGI = NULL;
	
	POSITION pos = GetStartPosition();
	
	while (pos)
	{
		GetNextAssoc(pos, dwTaskID, pGI);
		ASSERT(pGI);
		
		delete pGI;
	}
	
	CMap<DWORD, DWORD, GANTTITEM*, GANTTITEM*&>::RemoveAll();
}

BOOL CGanttItemMap::RemoveKey(DWORD dwKey)
{
	GANTTITEM* pGI = NULL;
	
	if (Lookup(dwKey, pGI))
	{
		delete pGI;
		return CMap<DWORD, DWORD, GANTTITEM*, GANTTITEM*&>::RemoveKey(dwKey);
	}
	
	// else
	return FALSE;
}

BOOL CGanttItemMap::HasTask(DWORD dwKey) const
{
	GANTTITEM* pGI = NULL;
	
	if (Lookup(dwKey, pGI))
	{
		ASSERT(pGI);
		return TRUE;
	}
	
	// else
	return FALSE;
}

//////////////////////////////////////////////////////////////////////

GANTTDISPLAY::GANTTDISPLAY() 
	: 
	nStartPos(GCDR_NOTDRAWN), 
	nEndPos(GCDR_NOTDRAWN), 
	nDonePos(GCDR_NOTDRAWN) 
{
}

int GANTTDISPLAY::GetBestTextPos() const
{
	return max(nEndPos, nDonePos);
}

BOOL GANTTDISPLAY::IsPosSet() const
{
	return (!HasNoDates() && (GetBestTextPos() > GCDR_NOTDRAWN));
}

BOOL GANTTDISPLAY::IsStartSet() const
{
	return (nStartPos > GCDR_NOTDRAWN);
}

BOOL GANTTDISPLAY::IsEndSet() const
{
	return (nEndPos > GCDR_NOTDRAWN);
}

BOOL GANTTDISPLAY::IsDoneSet() const
{
	return (nDonePos > GCDR_NOTDRAWN);
}

void GANTTDISPLAY::SetHasNoDates()
{
	nStartPos = nEndPos = nDonePos = GCDR_NODATES;
}

BOOL GANTTDISPLAY::HasNoDates() const
{
	// must all be set or none be set
	BOOL bAllSet = ((nStartPos == GCDR_NODATES) && (nEndPos == GCDR_NODATES) && (nDonePos == GCDR_NODATES));

	if (!bAllSet)
		ASSERT((nStartPos != GCDR_NODATES) && (nEndPos != GCDR_NODATES) && (nDonePos != GCDR_NODATES));
	
	return bAllSet;
}

//////////////////////////////////////////////////////////////////////

CGanttDisplayMap::~CGanttDisplayMap()
{
	RemoveAll();
}

void CGanttDisplayMap::RemoveAll()
{
	DWORD dwTaskID = 0;
	GANTTDISPLAY* pGD = NULL;
	
	POSITION pos = GetStartPosition();
	
	while (pos)
	{
		GetNextAssoc(pos, dwTaskID, pGD);
		ASSERT(pGD);
		
		delete pGD;
	}
	
	CMap<DWORD, DWORD, GANTTDISPLAY*, GANTTDISPLAY*&>::RemoveAll();
}

BOOL CGanttDisplayMap::RemoveKey(DWORD dwKey)
{
	GANTTDISPLAY* pGD = NULL;

	if (Lookup(dwKey, pGD))
	{
		delete pGD;
		return CMap<DWORD, DWORD, GANTTDISPLAY*, GANTTDISPLAY*&>::RemoveKey(dwKey);
	}

	// else
	return FALSE;
}

BOOL CGanttDisplayMap::HasTask(DWORD dwKey) const
{
	GANTTDISPLAY* pGD = NULL;
	
	if (Lookup(dwKey, pGD))
	{
		ASSERT(pGD);
		return TRUE;
	}

	// else
	return FALSE;
}

//////////////////////////////////////////////////////////////////////

