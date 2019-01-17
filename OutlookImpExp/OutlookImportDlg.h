#if !defined(AFX_OUTLOOKIMPORTDLG_H__E9D6C9F1_A424_4CB8_9AEF_0CE8641F1C53__INCLUDED_)
#define AFX_OUTLOOKIMPORTDLG_H__E9D6C9F1_A424_4CB8_9AEF_0CE8641F1C53__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// OutlookImportDlg.h : header file
//

#include "..\Shared\wndprompt.h"

/////////////////////////////////////////////////////////////////////////////
// COutlookImportDlg dialog

// predecs
namespace OutlookAPI
{
	class _Application;
	class MAPIFolder;
	class _TaskItem;
}

class ITaskList10;
class ITaskList;
class IPreferences;
typedef void* HTASKITEM;

class COutlookImportDlg : public CDialog
{
// Construction
public:
	COutlookImportDlg(CWnd* pParent = NULL);   // standard constructor

	BOOL ImportTasks(ITaskList* pDestTaskFile, IPreferences* pPrefs, LPCTSTR szKey);

// Dialog Data
	//{{AFX_DATA(COutlookImportDlg)
	enum { IDD = IDD_IMPORT_DIALOG };
	BOOL	m_bRemoveOutlookTasks;
	CString	m_sCurFolder;
	//}}AFX_DATA
	CTreeCtrl	m_tcTasks;
	ITaskList10* m_pDestTaskFile;
	OutlookAPI::_Application* m_pOutlook;
	OutlookAPI::MAPIFolder* m_pFolder;
	CWndPromptManager m_wndPrompt;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(COutlookImportDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL
	virtual void OnOK();

// Implementation
protected:
	int DoModal() { return CDialog::DoModal(); }

	// Generated message map functions
	//{{AFX_MSG(COutlookImportDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnDestroy();
	afx_msg void OnChoosefolder();
	afx_msg void OnClickTasklist(NMHDR* pNMHDR, LRESULT* pResult);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

protected:
	static time_t ConvertDate(DATE date);

	void AddFolderItemsToTree(OutlookAPI::MAPIFolder* pFolder, HTREEITEM htiParent = NULL);
	void SetTaskAttributes(HTASKITEM hTask, OutlookAPI::_TaskItem* pTask);
	BOOL DeleteTaskFromFolder(OutlookAPI::_TaskItem* pTask, OutlookAPI::MAPIFolder* pFolder);
	void AddTreeItemsToTasks(HTREEITEM htiParent, HTASKITEM hTaskParent, OutlookAPI::MAPIFolder* pFolder);
	void SetChildItemsChecked(HTREEITEM hti, BOOL bChecked);

	static BOOL TaskPathsMatch(OutlookAPI::_TaskItem* pTask1, OutlookAPI::_TaskItem* pTask2);
	static CString GetFullPath(OutlookAPI::_TaskItem* pTask);

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_OUTLOOKIMPORTDLG_H__E9D6C9F1_A424_4CB8_9AEF_0CE8641F1C53__INCLUDED_)
