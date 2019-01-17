#if !defined(AFX_TASKCALENDARCTRL_H__09FB7C3D_BBA8_43B3_A7B3_1D95C946892B__INCLUDED_)
#define AFX_TASKCALENDARCTRL_H__09FB7C3D_BBA8_43B3_A7B3_1D95C946892B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// TaskCalendarCtrl.h : header file
//

/////////////////////////////////////////////////////////////////////////////

#include <afxtempl.h>

#include "..\Interfaces\IUIExtension.h"
#include "..\Interfaces\ITaskList.h"

#include "..\3rdParty\CalendarCtrl.h"

/////////////////////////////////////////////////////////////////////////////

enum // options
{
	TCCO_CALCMISSINGSTARTASCREATION				= 0x0001,
	TCCO_CALCMISSINGSTARTASDUE					= 0x0002,
	TCCO_CALCMISSINGSTARTASEARLIESTDUEANDTODAY	= 0x0004,
	TCCO_CALCMISSINGDUEASSTART					= 0x0008,
	TCCO_CALCMISSINGDUEASLATESTSTARTANDTODAY	= 0x0010,
	TCCO_DISPLAYCONTINUOUS						= 0x0020,
	TCCO_DISPLAYSTART							= 0x0040,
	TCCO_DISPLAYDUE								= 0x0080,
	TCCO_DISPLAYCALCSTART						= 0x0100,
	TCCO_DISPLAYCALCDUE							= 0x0200,
	TCCO_ADJUSTTASKHEIGHTS						= 0x0400,
	TCCO_DISPLAYDONE							= 0x0800,
	TCCO_TASKTEXTCOLORISBKGND					= 0x1000,
};

/////////////////////////////////////////////////////////////////////////////

enum TCC_HITTEST
{
	TCCHT_NOWHERE = -1,
	TCCHT_BEGIN,
	TCCHT_MIDDLE,
	TCCHT_END,
};

/////////////////////////////////////////////////////////////////////////////

enum TCC_SNAPMODE
{
	TCCSM_NONE = -1,
	TCCSM_NEARESTDAY,
	TCCSM_NEARESTHALFDAY,
	TCCSM_NEARESTHOUR,
	TCCSM_FREE,
};

/////////////////////////////////////////////////////////////////////////////

// WPARAM = Hit test, LPARAM = Task ID
const UINT WM_CALENDAR_DATECHANGE = ::RegisterWindowMessage(_T("WM_CALENDAR_DATECHANGE"));

// WPARAM = 0, LPARAM = Task ID
const UINT WM_CALENDAR_SELCHANGE = ::RegisterWindowMessage(_T("WM_CALENDAR_SELCHANGE"));

// WPARAM = Drag Mode, LPARAM = Task ID
const UINT WM_CALENDAR_DRAGCHANGE = ::RegisterWindowMessage(_T("WM_CALENDAR_DRAGCHANGE"));

/////////////////////////////////////////////////////////////////////////////

const int DEF_TASK_HEIGHT = 18;

/////////////////////////////////////////////////////////////////////////////
// CTaskCalendarCtrl window

struct TASKCALITEM
{
public:
	TASKCALITEM();
	TASKCALITEM(const TASKCALITEM& tci);
	TASKCALITEM(const ITaskList14* pTasks, HTASKITEM hTask, DWORD dwCalcDates);

	TASKCALITEM& TASKCALITEM::operator=(const TASKCALITEM& tci);
	BOOL TASKCALITEM::operator==(const TASKCALITEM& tci);

	BOOL UpdateTask(const ITaskList14* pTasks, HTASKITEM hTask, int nEditAttrib, DWORD dwCalcDates);
	void RecalcDates(DWORD dwCalcDates);

	BOOL IsValid() const;
	BOOL IsDone() const;
	
	BOOL IsStartDateSet() const;
	void SetStartDate(const COleDateTime& date);

	COleDateTime GetAnyStartDate() const;
	BOOL HasAnyStartDate() const;

	BOOL IsEndDateSet() const;
	void SetEndDate(const COleDateTime& date);

	COleDateTime GetAnyEndDate() const;
	BOOL HasAnyEndDate() const;

	COLORREF GetFillColor(BOOL bSelected, BOOL bTextIsBack = FALSE) const;
	COLORREF GetBorderColor(BOOL bSelected, BOOL bTextIsBack = FALSE) const;
	COLORREF GetTextColor(BOOL bSelected, BOOL bTextIsBack = FALSE) const;
	BOOL HasTextColor() const;

	CString GetName() const;

public:
	COLORREF crText;
	DWORD dwTaskID;
	BOOL bDone, bGoodAsDone;

	static COLORREF crDone;
	
protected:
	COleDateTime dtCreation, dtStart, dtEnd;
	COleDateTime dtStartCalc, dtEndCalc;
	CString sName, sFormattedName;

protected:
	void UpdateTaskDates(const ITaskList14* pTasks, HTASKITEM hTask, int nEditAttrib, DWORD dwCalcDates);
	void ReformatName();

	static COleDateTime GetDate(time64_t tDate);
	static BOOL WantAttributeUpdate(int nEditAttrib, int nAttribMask);
};

/////////////////////////////////////////////////////////////////////////////

typedef CArray<TASKCALITEM*, TASKCALITEM*&> CTaskCalItemArray;
typedef CMap<DWORD, DWORD, TASKCALITEM*, TASKCALITEM*&> CTaskCalItemMap;
typedef CMap<double, double, BOOL, BOOL&> CSpecialDateMap;

/////////////////////////////////////////////////////////////////////////////

class CTaskCalendarCtrl : public CCalendarCtrl
{
// Construction
public:
	CTaskCalendarCtrl();
	virtual ~CTaskCalendarCtrl();

	void UpdateTasks(const ITaskList* pTasks, IUI_UPDATETYPE nUpdate, int nEditAttribute);
	bool PrepareNewTask(ITaskList* pTask) const;

	BOOL IsSpecialDate(const COleDateTime& date) const;
	BOOL GetSelectedTaskDates(COleDateTime& dtStart, COleDateTime& dtDue) const;
	DWORD GetSelectedTaskID() const { return m_dwSelectedTaskID; }
	BOOL CancelDrag();
	BOOL SelectTask(DWORD dwTaskID);
	BOOL HasTask(DWORD dwTaskID) const;
	DWORD HitTest(const CPoint& ptCursor) const;
	void EnsureVisible(DWORD dwTaskID, BOOL bShowStart);
	BOOL GetTaskLabelRect(DWORD dwTaskID, CRect& rLabel) const;
	void SetReadOnly(bool bReadOnly) { m_bReadOnly = bReadOnly; }
	BOOL SetVisibleWeeks(int nWeeks);
	void SetDoneTaskAttributes(COLORREF color, BOOL bStrikeThru);

	TCC_SNAPMODE GetSnapMode() const;
	void SetSnapMode(TCC_SNAPMODE nSnap) { m_nSnapMode = nSnap; }

	void SetOption(DWORD dwOption, BOOL bSet = TRUE);
	BOOL HasOption(DWORD dwOption) const { return ((m_dwOptions & dwOption) == dwOption); }

	static BOOL WantAttributeUpdate(int nEditAttribute);

protected:
	CTaskCalItemMap m_mapData;
	CSpecialDateMap m_mapSpecial;

	DWORD m_dwSelectedTaskID;
	DWORD m_dwOptions;
	BOOL m_bDraggingStart, m_bDraggingEnd, m_bDragging;
	TASKCALITEM m_tciPreDrag;
	CPoint m_ptDragOrigin;
	CToolTipCtrl m_tooltip;
	BOOL m_bReadOnly;
	int m_nCellVScrollPos;
	CScrollBar m_sbCellVScroll;
	CFont m_fontAltText, m_fontDone;

	mutable CMap<DWORD, DWORD, int, int> m_mapVertPos, m_mapTextOffset;
	mutable int m_nMaxDayTaskCount;
	mutable TCC_SNAPMODE m_nSnapMode;

	// Generated message map functions
protected:
	//{{AFX_MSG(CTaskCalendarCtrl)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnCaptureChanged(CWnd *pWnd);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	//}}AFX_MSG
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnSetFocus(CWnd* pFocus);
	afx_msg void OnKillFocus(CWnd* pFocus);
#if _MSC_VER >= 1400
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
#else
	afx_msg void OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
#endif
	DECLARE_MESSAGE_MAP()

protected:
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTaskCalendarCtrl)
	//}}AFX_VIRTUAL
	virtual int OnToolHitTest( CPoint point, TOOLINFO* pTI ) const;

	virtual void DrawHeader(CDC* pDC);
	virtual void DrawGrid(CDC* pDC);
	virtual void DrawCells(CDC* pDC);
	virtual void DrawCell(CDC* pDC, const CCalendarCell* pCell, const CRect& rCell, 
							BOOL bSelected, BOOL bToday, BOOL bShowMonth);
	virtual void DrawCellContent(CDC* pDC, const CCalendarCell* pCell, const CRect& rCell, 
									BOOL bSelected, BOOL bToday);
	virtual void DrawCellFocus(CDC* pDC, const CCalendarCell* pCell, const CRect& rCell);

	void BuildData(const ITaskList14* pTasks, HTASKITEM hTask, BOOL bAndSiblings);
	void DeleteData();

	void ResetPositions();
	int GetCellTasks(const COleDateTime& dtCell, CTaskCalItemArray& aTasks, BOOL bOrdered = TRUE) const;
	BOOL CalcTaskCellRect(DWORD dwTaskID, const CCalendarCell* pCell, const CRect& rCell, CRect& rTask) const;
	int GetTaskVertPos(DWORD dwTaskID, BOOL bVScrolled/* = FALSE*/) const;
	int GetTaskTextOffset(DWORD dwTaskID) const;
	TASKCALITEM* GetTaskCalItem(DWORD dwTaskID) const;
	bool GetGridCellFromTask(DWORD dwTaskID, int &nRow, int &nCol) const;

	BOOL UpdateCellScrollBarVisibility();
	BOOL IsCellScrollBarActive() const;
	int GetTaskHeight() const;
	int CalcRequiredTaskFontPointSize() const;

	DWORD HitTest(const CPoint& ptCursor, TCC_HITTEST& nHit) const;
	BOOL GetDateFromPoint(const CPoint& ptCursor, COleDateTime& date) const;
	BOOL StartDragging(const CPoint& ptCursor);
	BOOL EndDragging(const CPoint& ptCursor);
	BOOL UpdateDragging(const CPoint& ptCursor);
	BOOL IsValidDrag(const CPoint& ptDrag) const;
	BOOL ValidateDragPoint(CPoint& ptDrag) const;
	void CancelDrag(BOOL bReleaseCapture);
	BOOL IsDragging() const;
	BOOL GetValidDragDate(const CPoint& ptCursor, COleDateTime& dtDrag) const;
	double CalcDateDragTolerance() const;
	BOOL SelectTask(DWORD dwTaskID, BOOL bNotify);
	void RecalcTaskDates();

	void NotifyParentDateChange(TCC_HITTEST nHit);
	void NotifyParentDragChange();

	BOOL UpdateTask(const ITaskList14* pTasks, HTASKITEM hTask, IUI_UPDATETYPE nUpdate, int nEditAttribute, BOOL bAndSiblings);
	BOOL RemoveDeletedTasks(const ITaskList14* pTasks);

	// helpers
	static int CompareTCItems(const void* pV1, const void* pV2);

};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TASKCALENDARCTRL_H__09FB7C3D_BBA8_43B3_A7B3_1D95C946892B__INCLUDED_)
