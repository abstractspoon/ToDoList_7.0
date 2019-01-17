#if !defined(AFX_StatisticsWND_H__F2F5ABDC_CDB2_4197_A8E1_6FE134F95A20__INCLUDED_)
#define AFX_StatisticsWND_H__F2F5ABDC_CDB2_4197_A8E1_6FE134F95A20__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// StatisticsWnd.h : header file
//

#include <afxtempl.h>

#include "..\Shared\uitheme.h"

#include "..\Interfaces\Itasklist.h"
#include "..\Interfaces\IUIExtension.h"

#include "..\3rdparty\hmxchart.h"

/////////////////////////////////////////////////////////////////////////////

struct STATSITEM 
{ 
	STATSITEM();
	virtual ~STATSITEM();

	BOOL HasStart() const;
	BOOL IsDone() const;
	
	void MinMax(COleDateTime& dtMin, COleDateTime& dtMax) const;
	
	COleDateTime dtStart, dtDone; 
	DWORD dwTaskID;

protected:
	static void MinMax(const COleDateTime& date, COleDateTime& dtMin, COleDateTime& dtMax);

};
typedef CMap<DWORD, DWORD, STATSITEM, STATSITEM&> CMapStatsItems;

/////////////////////////////////////////////////////////////////////////////
// CStatisticsWnd dialog

class CStatisticsWnd : public CDialog, public IUIExtensionWindow
{
// Construction
public:
	CStatisticsWnd(CWnd* pParent = NULL);   // standard constructor
	virtual ~CStatisticsWnd();

	// IUIExtensionWindow
	BOOL Create(DWORD dwStyle, const RECT &rect, CWnd* pParentWnd, UINT nID);
	void Release();

	// IUIExtensionWindow
	LPCTSTR GetMenuText() const { return _T("Statistics"); }
	HICON GetIcon() const { return m_hIcon; }
	LPCTSTR GetTypeID() const { return STATS_TYPEID; }

	void SetReadOnly(bool /*bReadOnly*/) {}
	HWND GetHwnd() const { return GetSafeHwnd(); }
	void SetUITheme(const UITHEME* pTheme);

	bool GetLabelEditRect(LPRECT pEdit);
	IUI_HITTEST HitTest(const POINT& ptScreen) const;

	void LoadPreferences(const IPreferences* pPrefs, LPCTSTR szKey, BOOL bAppOnly);
	void SavePreferences(IPreferences* pPrefs, LPCTSTR szKey) const;

	void UpdateTasks(const ITaskList* pTasks, IUI_UPDATETYPE nUpdate, int nEditAttribute);
	bool WantUpdate(int nAttribute) const;
	bool PrepareNewTask(ITaskList* /*pTask*/) const { return true; }

	bool SelectTask(DWORD dwTaskID);
	bool SelectTasks(DWORD* pdwTaskIDs, int nTaskCount);

	bool ProcessMessage(MSG* pMsg);

	void DoAppCommand(IUI_APPCOMMAND nCmd, DWORD dwExtra);
	bool CanDoAppCommand(IUI_APPCOMMAND nCmd, DWORD dwExtra) const;

protected:
// Dialog Data
	//{{AFX_DATA(CStatisticsWnd)
	enum { IDD = IDD_STATISTICS_DLG };
	CStatic	m_stFrame;
	int		m_nDisplay;
	//}}AFX_DATA
	HICON m_hIcon;
	CBrush m_brBack;
	UITHEME m_theme;
	CHMXChart m_graph;
	int m_nScale;

	CMapStatsItems m_data;
	CDWordArray m_aDateOrdered;
	COleDateTime m_dtEarliestDone, m_dtLatestDone;

protected:
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CStatisticsWnd)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	//}}AFX_VIRTUAL
	virtual void OnCancel() {}
	virtual void OnOK() {}

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CStatisticsWnd)
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

protected:
	void UpdateTask(const ITaskList12* pTasks, HTASKITEM hTask, IUI_UPDATETYPE nUpdate, int nEditAttribute, BOOL bAndSiblings);
	void UpdateDataExtents();
	BOOL RemoveDeletedTasks(const ITaskList12* pTasks);
	void BuildData(const ITaskList12* pTasks);
	void BuildData(const ITaskList12* pTasks, HTASKITEM hTask, BOOL bAndSiblings);
	void SortData();
	BOOL IsDataSorted() const;
	void BuildGraph();
	int CalculateIncompleteTaskCount(const COleDateTime& date);
	BOOL GetStatsItem(DWORD dwTaskID, STATSITEM& si) const;
	void RebuildXScale(); // returns the previous scale
	int GetDataDuration() const;
	COleDateTime GetGraphStartDate() const;
	COleDateTime GetGraphEndDate() const;
	int CalculateRequiredScale() const;
	COleDateTime GetTaskStartDate(const ITaskList12* pTasks, HTASKITEM hTask);
	COleDateTime GetTaskDoneDate(const ITaskList12* pTasks, HTASKITEM hTask);

	static int CompareStatItems(const void* pV1, const void* pV2);
	static COleDateTime GetTaskDate(time64_t tDate);

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_StatisticsWND_H__F2F5ABDC_CDB2_4197_A8E1_6FE134F95A20__INCLUDED_)
