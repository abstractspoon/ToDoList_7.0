// TDCTreeListCtrl.h: interface for the CTDCListListCtrl class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_TDCLISTLISTCTRL_H__155791A3_22AC_4083_B933_F39E9EBDADEF__INCLUDED_)
#define AFX_TDCLISTLISTCTRL_H__155791A3_22AC_4083_B933_F39E9EBDADEF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////

#include "tdcmsg.h"
#include "tdcstruct.h"
#include "todoctrldata.h"
#include "tdctaskctrlbase.h"

/////////////////////////////////////////////////////////////////////////////

#include "..\Shared\EnHeaderCtrl.h"
#include "..\Shared\TreeListSyncer.h"
#include "..\Shared\runtimedlg.h"
#include "..\shared\sysimagelist.h"

/////////////////////////////////////////////////////////////////////////////

#include <afxtempl.h>
typedef CArray<int, int> CIntArray;

/////////////////////////////////////////////////////////////////////////////

class CToDoCtrlData;

/////////////////////////////////////////////////////////////////////////////

class CTDCTaskListCtrl : public CTDCTaskCtrlBase  
{
	DECLARE_DYNAMIC(CTDCTaskListCtrl);
	
public:
	CTDCTaskListCtrl(const CTDCImageList& ilIcons,
					const CToDoCtrlData& data, 
					const CToDoCtrlFind& find,
					const CWordArray& aStyles,
					const CTDCColumnIDArray& aVisibleCols,
					const CTDCCustomAttribDefinitionArray& aCustAttribDefs);

	virtual ~CTDCTaskListCtrl();
	operator HWND() const { return GetSafeHwnd(); }
	
	inline CListCtrl& List() { return m_lcTasks; }
	inline const CListCtrl& List() const { return m_lcTasks; }
	
	inline const TODOITEM* GetTask(DWORD dwTaskID) const { return m_data.GetTask(dwTaskID); }
	inline const TODOITEM* GetTask(int nItem) const { return GetTask(GetTaskID(nItem)); }
	inline DWORD GetTaskID(int nItem) const { return m_lcTasks.GetItemData(nItem); }
	inline DWORD GetTrueTaskID(int nItem) const { return m_data.GetTrueTaskID(GetTaskID(nItem)); }
	inline UINT GetTaskCount() const { return m_data.GetTaskCount(); }
	inline DWORD GetSelectedTaskID() const { return GetTaskID(GetSelectedItem()); }
	inline int GetSelectedCount() const { return m_lcTasks.GetSelectedCount(); }
	inline int GetItemCount() const { return m_lcTasks.GetItemCount(); }
	inline int HitTestItem(POINT point, UINT* pFlags = NULL) const { return m_lcTasks.HitTest(point, pFlags); }

	const TODOITEM* GetSelectedTask() const;
	int GetSelectedTaskIDs(CDWordArray& aTaskIDs, BOOL bTrue = FALSE) const;
	int GetSelectedTaskIDs(CDWordArray& aTaskIDs, DWORD& dwFocusedTaskID) const;
	BOOL SelectTasks(const CDWordArray& aTasks, BOOL bTrue = FALSE);
	void CacheSelection(TDCSELECTIONCACHE& cache) const;
	void RestoreSelection(const TDCSELECTIONCACHE& cache);
	BOOL IsTaskSelected(DWORD dwTaskID) const;
	BOOL EnsureSelectionVisible();
	BOOL GetSelectionBoundingRect(CRect& rSelection) const;
	BOOL GetLabelEditRect(CRect& rLabel) const;

	// list related
	int GetSelectedItem() const;
	BOOL SelectItem(int nItem);
	BOOL SelectAll();
	BOOL InvalidateSelection(BOOL bUpdate = FALSE);
	BOOL InvalidateItem(int nItem, BOOL bUpdate = FALSE);
	int GetFirstSelectedItem() const;
	DWORD GetFocusedListTaskID() const;
	int GetFocusedListItem() const;
	int FindTaskItem(DWORD dwTaskID) const;

	void GetWindowRect(CRect& rWindow) const { CWnd::GetWindowRect(rWindow); }
	void DeleteAll();
	void RemoveDeletedItems();
	
	BOOL CancelOperation();

	void OnStyleUpdated(TDC_STYLE nStyle, BOOL bOn, BOOL bDoUpdate);

protected:
	CListCtrl m_lcTasks;

	// Overrides
protected:
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTDCListListCtrl)
	//}}AFX_VIRTUAL
	
protected:
	// Generated message map functions
	//{{AFX_MSG(CTDCListListCtrl)
	//}}AFX_MSG
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	
	DECLARE_MESSAGE_MAP()

protected:
	// base-class overrides
	LRESULT WindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
	LRESULT ScWindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
	
 	LRESULT OnListCustomDraw(NMLVCUSTOMDRAW* pLVCD);
	LRESULT OnListGetDispInfo(NMLVDISPINFO* pLVDI);

	void OnListSelectionChange(NMLISTVIEW* pNMLV);

	void OnNotifySplitterChange(int nSplitPos);
	BOOL IsListItemSelected(HWND hwnd, int nItem) const;
	int RecalcColumnWidth(int nCol, CDC* pDC) const;
	DWORD GetColumnItemTaskID(int nItem) const { return GetTaskID(nItem); }
	BOOL InvalidateTask(DWORD dwTaskID, BOOL bUpdate);
	void DeselectAll();
	void NotifyParentSelChange(SELCHANGE_ACTION nAction = SC_UNKNOWN);

protected:
	void GetWindowRect(CRect& rWindow, BOOL bWithHeader) const;
	BOOL BuildColumns();
	void Release();
	BOOL CreateTasksWnd(CWnd* pParentWnd, const CRect& rect, BOOL bVisible);
	DWORD HitTestTasksTask(const CPoint& ptScreen) const;
	CString GetLongestValue(TDC_ATTRIBUTE nAttrib, const CString& sExtra) const;
	void SetTasksImageList(HIMAGELIST hil, BOOL bState, BOOL bOn = TRUE);
	inline HWND TasksHeader() const { return m_hdrTasks; }
	inline HWND Tasks() const { return m_lcTasks; }
	int GetTopIndex() const;
	BOOL SetTopIndex(int nIndex);
	BOOL NeedDrawTaskSelection() { return (HasFocus() && (GetFocus() != &m_lcTasks)); }
	void SetSelectedTasks(const CDWordArray& aTaskIDs, DWORD dwFocusedTaskID);
	BOOL HandleClientColumnClick(const CPoint& pt, BOOL bDblClk);
	BOOL GetItemTitleRect(int nItem, TDC_TITLERECT nArea, CRect& rect, CDC* pDC = NULL, LPCTSTR szTitle = NULL) const;
	
	static BOOL HasHitTestFlag(UINT nFlags, UINT nFlag);

	GM_ITEMSTATE GetListItemState(int nItem) const;
	GM_ITEMSTATE GetColumnItemState(int nItem) const;

};

#endif // !defined(AFX_TDCLISTLISTCTRL_H__155791A3_22AC_4083_B933_F39E9EBDADEF__INCLUDED_)
