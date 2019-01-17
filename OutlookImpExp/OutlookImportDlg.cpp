// OutlookImportDlg.cpp : implementation file
//

#include "stdafx.h"
#include "OutlookImpExp.h"
#include "OutlookImportDlg.h"

#include "..\shared\misc.h"
#include "..\shared\localizer.h"

#include "..\Interfaces\IPreferences.h"
#include "..\Interfaces\ITaskList.h"

#include "..\3rdparty\msoutl.h"

#include <afxtempl.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////

const int olFolderDeletedItems	= 3;
const int olFolderOutbox		= 4;
const int olFolderSentMail		= 5;
const int olFolderInbox			= 6;
const int olFolderCalendar		= 9;
const int olFolderContacts		= 10;
const int olFolderJournal		= 11;
const int olFolderNotes			= 12;
const int olFolderTasks			= 13;

/////////////////////////////////////////////////////////////////////////////

using namespace OutlookAPI;

/////////////////////////////////////////////////////////////////////////////

const LPCTSTR PATHDELIM = _T(" \\ ");

/////////////////////////////////////////////////////////////////////////////
// COutlookImportDlg dialog


COutlookImportDlg::COutlookImportDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COutlookImportDlg::IDD, pParent), m_pDestTaskFile(NULL), m_pOutlook(NULL)
{
	//{{AFX_DATA_INIT(COutlookImportDlg)
	m_sCurFolder = _T("");
	//}}AFX_DATA_INIT
}

void COutlookImportDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(COutlookImportDlg)
	DDX_Check(pDX, IDC_REMOVEOUTLOOKTASKS, m_bRemoveOutlookTasks);
	DDX_Control(pDX, IDC_TASKLIST, m_tcTasks);
	DDX_Text(pDX, IDC_CURFOLDER, m_sCurFolder);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(COutlookImportDlg, CDialog)
	//{{AFX_MSG_MAP(COutlookImportDlg)
	ON_WM_DESTROY()
	ON_BN_CLICKED(IDC_CHOOSEFOLDER, OnChoosefolder)
	ON_NOTIFY(NM_CLICK, IDC_TASKLIST, OnClickTasklist)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// COutlookImportDlg message handlers

BOOL COutlookImportDlg::ImportTasks(ITaskList* pDestTaskFile, IPreferences* pPrefs, LPCTSTR szKey)
{
	m_pDestTaskFile = GetITLInterface<ITaskList10>(pDestTaskFile, IID_TASKLIST10);
	ASSERT(m_pDestTaskFile);

	if (!m_pDestTaskFile)
		return FALSE;

	CString sKey(szKey);
	sKey += _T("\\Outlook");

	m_bRemoveOutlookTasks = pPrefs->GetProfileInt(sKey, _T("RemoveOutlookTasks"), FALSE);

	if (DoModal() == IDOK)
	{
		pPrefs->WriteProfileInt(sKey, _T("RemoveOutlookTasks"), m_bRemoveOutlookTasks);
		return TRUE;
	}

	// else
	return FALSE;
}

BOOL COutlookImportDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	CLocalizer::EnableTranslation(m_tcTasks, FALSE);
	
	ASSERT(m_pOutlook == NULL);
	m_pOutlook = new _Application;

	if (m_pOutlook->CreateDispatch(_T("Outlook.Application")))
	{
		_NameSpace nmspc(m_pOutlook->GetNamespace(_T("MAPI")));
		nmspc.m_lpDispatch->AddRef(); // to keep it alive

		m_pFolder = new MAPIFolder(nmspc.GetDefaultFolder(olFolderTasks));
		m_pFolder->m_lpDispatch->AddRef(); // to keep it alive

		AddFolderItemsToTree(m_pFolder);
		
		m_sCurFolder = m_pFolder->GetName();
		UpdateData(FALSE);

		m_wndPrompt.SetPrompt(m_tcTasks, IDS_FOLDERNOITEMS, TVM_GETCOUNT);
	}
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void COutlookImportDlg::AddFolderItemsToTree(MAPIFolder* pFolder, HTREEITEM htiParent)
{
	_Items items(pFolder->GetItems());
	items.m_lpDispatch->AddRef(); // to keep it alive

	int nItem, nCount = items.GetCount();

	for (nItem = 1; nItem <= nCount; nItem++)
	{
		LPDISPATCH lpd = items.Item(COleVariant((short)nItem));
		lpd->AddRef(); // to keep it alive
		_TaskItem task(lpd);

		HTREEITEM hti = m_tcTasks.InsertItem(task.GetSubject(), htiParent);

		TRACE(_T("%s / %s (item)\n"), m_tcTasks.GetItemText(htiParent), m_tcTasks.GetItemText(hti));

		m_tcTasks.SetItemData(hti, (DWORD)lpd);
		m_tcTasks.SetCheck(hti, TRUE);
	}

	// likewise for subfolders
	_Items folders(pFolder->GetFolders());
	nCount = folders.GetCount();

	for (nItem = 1; nItem <= nCount; nItem++)
	{
		MAPIFolder folder(folders.Item(COleVariant((short)nItem)));
		folder.m_lpDispatch->AddRef();

		HTREEITEM hti = m_tcTasks.InsertItem(folder.GetName(), htiParent);

		TRACE(_T("%s / %s (folder)\n"), m_tcTasks.GetItemText(htiParent), m_tcTasks.GetItemText(hti));
		
		m_tcTasks.SetItemData(hti, (DWORD)folder.m_lpDispatch);
		m_tcTasks.SetCheck(hti, TRUE);

		AddFolderItemsToTree(&folder, hti); // RECURSIVE call

		// and expand
		m_tcTasks.Expand(hti, TVE_EXPAND);
	}
}

void COutlookImportDlg::OnOK()
{
	CDialog::OnOK();

	ASSERT(m_pOutlook && m_pDestTaskFile);

	// make sure nothing has changed
	_NameSpace nmspc(m_pOutlook->GetNamespace(_T("MAPI")));
	nmspc.m_lpDispatch->AddRef(); // to keep it alive

	MAPIFolder mapi(nmspc.GetDefaultFolder(olFolderTasks));
	mapi.m_lpDispatch->AddRef(); // to keep it alive

	_Items items(mapi.GetItems());
	items.m_lpDispatch->AddRef(); // to keep it alive

	AddTreeItemsToTasks(NULL, NULL, &mapi);
}

void COutlookImportDlg::AddTreeItemsToTasks(HTREEITEM htiParent, HTASKITEM hTaskParent, MAPIFolder* pFolder)
{
	ASSERT ((htiParent && hTaskParent) || (!htiParent && !hTaskParent));

	// iterate the tree items adding checked items
	HTREEITEM hti = m_tcTasks.GetChildItem(htiParent);

	while (hti)
	{
		if (m_tcTasks.GetCheck(hti))
		{
			// create a new task
			CString sTitle = m_tcTasks.GetItemText(hti);

			HTASKITEM hTask = m_pDestTaskFile->NewTask(sTitle, hTaskParent, 0);
			ASSERT(hTask);

			// add attributes for leaf tasks
			if (m_tcTasks.GetChildItem(hti) == NULL)
			{
				// get the task index that this item points to
				LPDISPATCH lpd = (LPDISPATCH)m_tcTasks.GetItemData(hti);
				_TaskItem task(lpd);

					SetTaskAttributes(hTask, &task);

					// delete the item as we go if required
					if (m_bRemoveOutlookTasks)
					DeleteTaskFromFolder(&task, pFolder);
				}
			else // process RECURSIVELY
			{
				AddTreeItemsToTasks(hti, hTask, pFolder);
			}
		}

		hti = m_tcTasks.GetNextItem(hti, TVGN_NEXT);
	}
}

BOOL COutlookImportDlg::DeleteTaskFromFolder(_TaskItem* pTask, MAPIFolder* pFolder)
{
	// look through this folders tasks first
	_Items items(pFolder->GetItems());
	items.m_lpDispatch->AddRef(); // to keep it alive
	int nItem, nTaskCount = items.GetCount();

	for (nItem = 1; nItem <= nTaskCount; nItem++)
	{
		LPDISPATCH lpd = items.Item(COleVariant((short)nItem));
		_TaskItem taskTest(lpd);

		if (TaskPathsMatch(pTask, &taskTest))
		{
			items.Remove(nItem);
			return TRUE;
		}
	}

	// then for subfolders
	_Items folders(pFolder->GetFolders());
	int nCount = folders.GetCount();

	for (nItem = 1; nItem <= nCount; nItem++)
	{
		MAPIFolder folder(folders.Item(COleVariant((short)nItem)));

		if (DeleteTaskFromFolder(pTask, &folder))
			return TRUE;
	}

	return FALSE;
}

BOOL COutlookImportDlg::TaskPathsMatch(_TaskItem* pTask1, _TaskItem* pTask2)
{
	CString sPath1 = GetFullPath(pTask1);
	CString sPath2 = GetFullPath(pTask2);

	return (sPath1 == sPath2);
}

CString COutlookImportDlg::GetFullPath(_TaskItem* pTask)
{
	CString sPath(pTask->GetSubject()), sFolder;
	LPDISPATCH lpd = pTask->GetParent();

	do
	{
		try
		{
			MAPIFolder folder(lpd);
			sFolder = folder.GetName(); // will throw when we hit the highest level
			sPath = sFolder + PATHDELIM + sPath;

			lpd = folder.GetParent(); 
		}
		catch (...)
		{
			break;
		}
	}
	while (true);

	return sPath;
}

void COutlookImportDlg::SetTaskAttributes(HTASKITEM hTask, _TaskItem* pTask)
{
	// set it's attributes
	m_pDestTaskFile->SetTaskComments(hTask, pTask->GetBody());

	// can have multiple categories
	CStringArray aCats;
	Misc::Split(pTask->GetCategories(), aCats);

	for (int nCat = 0; nCat < aCats.GetSize(); nCat++)
		m_pDestTaskFile->AddTaskCategory(hTask, aCats[nCat]);
	
	if (pTask->GetComplete())
		m_pDestTaskFile->SetTaskDoneDate(hTask, ConvertDate(pTask->GetDateCompleted()));
	
	m_pDestTaskFile->SetTaskDueDate(hTask, ConvertDate(pTask->GetDueDate()));
	m_pDestTaskFile->SetTaskStartDate(hTask, ConvertDate(pTask->GetStartDate()));
	m_pDestTaskFile->SetTaskCreationDate(hTask, ConvertDate(pTask->GetCreationTime()));
	m_pDestTaskFile->SetTaskLastModified(hTask, ConvertDate(pTask->GetLastModificationTime()));
	m_pDestTaskFile->SetTaskAllocatedBy(hTask, pTask->GetDelegator());
	m_pDestTaskFile->SetTaskAllocatedTo(hTask, pTask->GetOwner());
	m_pDestTaskFile->SetTaskPriority(hTask, (unsigned char)(pTask->GetImportance() * 5));
	m_pDestTaskFile->SetTaskComments(hTask, pTask->GetBody());

	// save outlook ID unless removing from Outlook
	if (!m_bRemoveOutlookTasks)
	{
		CString sFileLink;
		sFileLink.Format(_T("outlook:%s"), pTask->GetEntryID());
		m_pDestTaskFile->SetTaskFileReferencePath(hTask, sFileLink);

		// and save entry ID as meta-data so we
		// maintain the association for synchronization
		m_pDestTaskFile->SetTaskMetaData(hTask, OUTLOOK_METADATAKEY, pTask->GetEntryID());
	}

/*
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
	m_pDestTaskFile->SetTask(hTask, pTask->Get());
*/
}

void COutlookImportDlg::OnDestroy() 
{
	CDialog::OnDestroy();
	
	delete m_pOutlook;
	delete m_pFolder;

	m_pOutlook = NULL;
	m_pDestTaskFile = NULL;
	m_pFolder = NULL;
}

time_t COutlookImportDlg::ConvertDate(DATE date)
{
	if (date <= 0.0)
		return 0;

	SYSTEMTIME st;
	COleDateTime dt(date);

	dt.GetAsSystemTime(st);

	tm t = { st.wSecond, st.wMinute, st.wHour, st.wDay, st.wMonth - 1, st.wYear - 1900, 0 };
	return mktime(&t);
}

void COutlookImportDlg::OnChoosefolder() 
{
	_NameSpace nmspc(m_pOutlook->GetNamespace(_T("MAPI")));
	nmspc.m_lpDispatch->AddRef(); // to keep it alive

	LPDISPATCH pDisp = nmspc.PickFolder();

	if (pDisp)
	{
		delete m_pFolder;
		m_pFolder = new MAPIFolder(pDisp);

		m_tcTasks.DeleteAllItems();
		AddFolderItemsToTree(m_pFolder);
		
		m_sCurFolder = m_pFolder->GetName();
		UpdateData(FALSE);
	}
}

void COutlookImportDlg::OnClickTasklist(NMHDR* /*pNMHDR*/, LRESULT* pResult) 
{
	CPoint ptClick(GetMessagePos());
	m_tcTasks.ScreenToClient(&ptClick);

	UINT nFlags = 0;
	HTREEITEM hti = m_tcTasks.HitTest(ptClick, &nFlags);

	if (hti && (TVHT_ONITEMSTATEICON & nFlags))
	{
		BOOL bChecked = !m_tcTasks.GetCheck(hti); // check hasn't happened yet

		// clear or set all children
		SetChildItemsChecked(hti, bChecked);

		// check parents if necessary
		if (bChecked)
		{
			HTREEITEM htiParent = m_tcTasks.GetParentItem(hti);

			while (htiParent)
			{
				// item
				m_tcTasks.SetCheck(htiParent, bChecked);
					
				// then parent
				htiParent = m_tcTasks.GetParentItem(htiParent);
			}
		}
    }

	*pResult = 0;
}

void COutlookImportDlg::SetChildItemsChecked(HTREEITEM hti, BOOL bChecked)
{
	HTREEITEM htiChild = m_tcTasks.GetChildItem(hti);

	while (htiChild)
	{
		// item
		m_tcTasks.SetCheck(htiChild, bChecked);
		
		// then children
		SetChildItemsChecked(htiChild, bChecked); // RECURSIVE

		htiChild = m_tcTasks.GetNextItem(htiChild, TVGN_NEXT);
	}
}
