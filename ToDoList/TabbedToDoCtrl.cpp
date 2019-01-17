// Fi M_BlISlteredToDoCtrl.cpp: implementation of the CTabbedToDoCtrl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TabbedToDoCtrl.h"
#include "todoitem.h"
#include "resource.h"
#include "tdcstatic.h"
#include "tdcmsg.h"
#include "tdccustomattributehelper.h"
#include "tdltaskicondlg.h"
#include "tdcuiextensionhelper.h"

#include "..\shared\holdredraw.h"
#include "..\shared\datehelper.h"
#include "..\shared\enstring.h"
#include "..\shared\preferences.h"
#include "..\shared\deferwndmove.h"
#include "..\shared\autoflag.h"
#include "..\shared\holdredraw.h"
#include "..\shared\osversion.h"
#include "..\shared\graphicsmisc.h"
#include "..\shared\uiextensionmgr.h"
#include "..\shared\filemisc.h"

#include "..\Interfaces\iuiextension.h"

#include <math.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

#ifndef LVS_EX_DOUBLEBUFFER
#define LVS_EX_DOUBLEBUFFER 0x00010000
#endif

#ifndef LVS_EX_LABELTIP
#define LVS_EX_LABELTIP     0x00004000
#endif

//////////////////////////////////////////////////////////////////////

const UINT SORTWIDTH = 10;
const UINT DEFTEXTFLAGS = (DT_END_ELLIPSIS | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CTabbedToDoCtrl::CTabbedToDoCtrl(CUIExtensionMgr& mgrUIExt, CContentMgr& mgrContent, 
								 const CONTENTFORMAT& cfDefault, const TDCCOLEDITFILTERVISIBILITY& visDefault) 
	:
	CToDoCtrl(mgrContent, cfDefault, visDefault), 
	m_bTreeNeedResort(FALSE),
	m_bTaskColorChange(FALSE),
	m_bUpdatingExtensions(FALSE),
	m_bExtModifyingApp(FALSE),
	m_mgrUIExt(mgrUIExt),
	m_taskList(m_ilTaskIcons, m_data, m_taskTree.Find(), m_aStyles, m_visColAttrib.GetVisibleColumns(), m_aCustomAttribDefs)
{
	// add extra controls to implement list-view
	for (int nCtrl = 0; nCtrl < NUM_FTDCCTRLS; nCtrl++)
	{
		const TDCCONTROL& ctrl = FTDCCONTROLS[nCtrl];

		AddRCControl(_T("CONTROL"), ctrl.szClass, CString((LPCTSTR)ctrl.nIDCaption), 
					ctrl.dwStyle, ctrl.dwExStyle,
					ctrl.nX, ctrl.nY, ctrl.nCx, ctrl.nCy, ctrl.nID);
	}

	// tab is on by default
	m_aStyles.SetAt(TDCS_SHOWTREELISTBAR, 1);
}

CTabbedToDoCtrl::~CTabbedToDoCtrl()
{
	// cleanup extension views
	int nView = m_aExtViews.GetSize();

	while (nView--)
	{
		IUIExtensionWindow* pExtWnd = m_aExtViews[nView];

		if (pExtWnd)
			pExtWnd->Release();
	}

	m_aExtViews.RemoveAll();
}

BEGIN_MESSAGE_MAP(CTabbedToDoCtrl, CToDoCtrl)
//{{AFX_MSG_MAP(CTabbedToDoCtrl)
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_FTC_TASKLISTLIST, OnListSelChanged)
	ON_NOTIFY(NM_CLICK, IDC_FTC_TASKLISTLIST, OnListClick)
	ON_NOTIFY(NM_RCLICK, IDC_FTC_TABCTRL, OnTabCtrlRClick)

	ON_REGISTERED_MESSAGE(WM_TDCN_COLUMNEDITCLICK, OnColumnEditClick)
	ON_REGISTERED_MESSAGE(WM_TDCN_VIEWPOSTCHANGE, OnPostTabViewChange)
	ON_REGISTERED_MESSAGE(WM_TDCN_VIEWPRECHANGE, OnPreTabViewChange)

	ON_REGISTERED_MESSAGE(WM_IUI_EDITSELECTEDTASKTITLE, OnUIExtEditSelectedTaskTitle)
	ON_REGISTERED_MESSAGE(WM_IUI_MODIFYSELECTEDTASK, OnUIExtModifySelectedTask)
	ON_REGISTERED_MESSAGE(WM_IUI_SELECTTASK, OnUIExtSelectTask)
	ON_REGISTERED_MESSAGE(WM_IUI_SORTCOLUMNCHANGE, OnUIExtSortColumnChange)

	ON_REGISTERED_MESSAGE(WM_PCANCELEDIT, OnEditCancel)
	ON_REGISTERED_MESSAGE(WM_TLDT_DROP, OnDropObject)
	ON_REGISTERED_MESSAGE(WM_TDCM_TASKHASREMINDER, OnTaskHasReminder)
	ON_WM_DRAWITEM()
	ON_WM_ERASEBKGND()
	ON_WM_MEASUREITEM()
END_MESSAGE_MAP()

///////////////////////////////////////////////////////////////////////////

void CTabbedToDoCtrl::DoDataExchange(CDataExchange* pDX)
{
	CToDoCtrl::DoDataExchange(pDX);
	
	DDX_Control(pDX, IDC_FTC_TABCTRL, m_tabViews);
}

BOOL CTabbedToDoCtrl::OnInitDialog()
{
	CToDoCtrl::OnInitDialog();

	// create the list-list before anything else
	CRect rect(0, 0, 0, 0);
	VERIFY(m_taskList.Create(this, rect, IDC_FTC_TASKLISTLIST));
	
	// list initialisation
	m_dtList.Register(&m_taskList.List(), this);
	m_taskList.SetWindowPos(&m_taskTree, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOREDRAW);
	m_taskList.SetWindowPrompt(CEnString(IDS_TDC_TASKLISTPROMPT));

	// add tree and list as tabbed views
	m_tabViews.AttachView(m_taskTree.GetSafeHwnd(), FTCV_TASKTREE, CEnString(IDS_TASKTREE), GraphicsMisc::LoadIcon(IDI_TASKTREE_STD), NULL);
	m_tabViews.AttachView(m_taskList, FTCV_TASKLIST, CEnString(IDS_LISTVIEW), GraphicsMisc::LoadIcon(IDI_LISTVIEW_STD), NewViewData());

	for (int nExt = 0; nExt < m_mgrUIExt.GetNumUIExtensions(); nExt++)
		AddView(m_mgrUIExt.GetUIExtension(nExt));

	Resize();

	return FALSE;
}

void CTabbedToDoCtrl::OnStylesUpdated()
{
	CToDoCtrl::OnStylesUpdated();

	m_taskList.OnStylesUpdated();
}

void CTabbedToDoCtrl::OnTaskIconsChanged()
{
	CToDoCtrl::OnTaskIconsChanged();
	
	m_taskList.OnImageListChange();
}

LRESULT CTabbedToDoCtrl::OnTaskHasReminder(WPARAM wp, LPARAM lp)
{
	UNREFERENCED_PARAMETER(lp);
	ASSERT(wp && (((HWND)lp == m_taskTree.GetSafeHwnd()) || 
					((HWND)lp == m_taskList.GetSafeHwnd())));
	
	return TaskHasReminder(wp);
}

BOOL CTabbedToDoCtrl::PreTranslateMessage(MSG* pMsg) 
{
	// see if an UI extension wants this
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKLIST:
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		break;
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView, FALSE);
			ASSERT(pExtWnd);

			if (pExtWnd->ProcessMessage(pMsg))
				return true;
		}
		break;
		
	default:
		ASSERT(0);
	}

	return CToDoCtrl::PreTranslateMessage(pMsg);
}

void CTabbedToDoCtrl::SetUITheme(const CUIThemeFile& theme)
{
	CToDoCtrl::SetUITheme(theme);

	m_tabViews.SetBackgroundColor(theme.crAppBackLight);
	m_taskList.SetSplitBarColor(theme.crAppBackDark);

	// update extensions
	int nExt = m_aExtViews.GetSize();
			
	while (nExt--)
	{
		IUIExtensionWindow* pExtWnd = m_aExtViews[nExt];
		
		if (pExtWnd)
		{
			// prepare theme file
			CUIThemeFile themeExt(theme);

			themeExt.SetToolbarImageFile(pExtWnd->GetTypeID());
			pExtWnd->SetUITheme(&themeExt);
		}
	}
}

BOOL CTabbedToDoCtrl::LoadTasks(const CTaskFile& file)
{
	if (!CToDoCtrl::LoadTasks(file))
		return FALSE;

	m_taskList.SetTasklistFolder(FileMisc::GetFolderFromFilePath(m_sLastSavePath));

	// reload last view
	CPreferences prefs;
	
	if (GetView() == FTCV_UNSET)
	{
		CString sKey = GetPreferencesKey(); // no subkey
		
		if (!sKey.IsEmpty()) // not first time
		{
			CPreferences prefs;

			// restore view visibility
			ShowListViewTab(prefs.GetProfileInt(sKey, _T("ListViewVisible"), TRUE));

			CStringArray aTypeIDs;
			int nExt = prefs.GetProfileInt(sKey, _T("VisibleExtensionCount"), -1);

			if (nExt >= 0)
			{
				while (nExt--)
				{
					CString sSubKey = Misc::MakeKey(_T("VisibleExt%d"), nExt);
					aTypeIDs.Add(prefs.GetProfileString(sKey, sSubKey));
				}

				SetVisibleExtensionViews(aTypeIDs);
			}
			
			FTC_VIEW nView = (FTC_VIEW)prefs.GetProfileInt(sKey, _T("View"), FTCV_UNSET);

			if ((nView != FTCV_UNSET) && (nView != GetView()))
				SetView(nView);

			// clear the view so we don't keep restoring it
			prefs.WriteProfileInt(sKey, _T("View"), FTCV_UNSET);
		}
	}
	else
	{
		FTC_VIEW nView = GetView();

		// handle extension views
		switch (nView)
		{
		case FTCV_TASKTREE:
		case FTCV_TASKLIST:
			SetExtensionsNeedUpdate(TRUE);
			break;

		case FTCV_UIEXTENSION1:
		case FTCV_UIEXTENSION2:
		case FTCV_UIEXTENSION3:
		case FTCV_UIEXTENSION4:
		case FTCV_UIEXTENSION5:
		case FTCV_UIEXTENSION6:
		case FTCV_UIEXTENSION7:
		case FTCV_UIEXTENSION8:
		case FTCV_UIEXTENSION9:
		case FTCV_UIEXTENSION10:
		case FTCV_UIEXTENSION11:
		case FTCV_UIEXTENSION12:
		case FTCV_UIEXTENSION13:
		case FTCV_UIEXTENSION14:
		case FTCV_UIEXTENSION15:
		case FTCV_UIEXTENSION16:
			{
				IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView);
				ASSERT(pExtWnd);

				if (pExtWnd == NULL)
					return FALSE;

				UpdateExtensionView(pExtWnd, file, IUI_ALL, TDCA_ALL);
				UpdateExtensionViewSelection();

				// mark rest of extensions needing update
				SetExtensionsNeedUpdate(TRUE, nView);
			}
			break;
		}
	}

	return TRUE;
}

void CTabbedToDoCtrl::OnDestroy() 
{
	if (GetView() != FTCV_UNSET)
	{
		CPreferences prefs;
		CString sKey = GetPreferencesKey(); // no subkey
		
		// save view
		if (!sKey.IsEmpty())
		{
			prefs.WriteProfileInt(sKey, _T("View"), GetView());
			prefs.WriteProfileInt(sKey, _T("View"), GetView());

			// save view visibility
			prefs.WriteProfileInt(sKey, _T("ListViewVisible"), IsListViewTabShowing());

			CStringArray aTypeIDs;
			int nExt = GetVisibleExtensionViews(aTypeIDs);

			prefs.WriteProfileInt(sKey, _T("VisibleExtensionCount"), nExt);

			while (nExt--)
			{
				CString sSubKey = Misc::MakeKey(_T("VisibleExt%d"), nExt);
				prefs.WriteProfileString(sKey, sSubKey, aTypeIDs[nExt]);
			}

			// extensions
			int nView = m_aExtViews.GetSize();

			if (nView)
			{
				CString sKey = GetPreferencesKey(_T("UIExtensions"));

				while (nView--)
				{
					IUIExtensionWindow* pExtWnd = m_aExtViews[nView];

					if (pExtWnd)
						pExtWnd->SavePreferences(&prefs, sKey);
				}
			}
		}
	}
		
	CToDoCtrl::OnDestroy();
}

void CTabbedToDoCtrl::UpdateVisibleColumns()
{
	CToDoCtrl::UpdateVisibleColumns();

	m_taskList.OnColumnVisibilityChange();
}

IUIExtensionWindow* CTabbedToDoCtrl::GetExtensionWnd(FTC_VIEW nView) const
{
	ASSERT(nView >= FTCV_FIRSTUIEXTENSION && nView <= FTCV_LASTUIEXTENSION);

	if (nView < FTCV_FIRSTUIEXTENSION || nView > FTCV_LASTUIEXTENSION)
		return NULL;

	int nExtension = (nView - FTCV_FIRSTUIEXTENSION);
	ASSERT(nExtension < m_aExtViews.GetSize());

	IUIExtensionWindow* pExtWnd = m_aExtViews[nExtension];
	ASSERT(pExtWnd || (m_tabViews.GetViewHwnd(nView) == NULL));

	return pExtWnd;
}

IUIExtensionWindow* CTabbedToDoCtrl::GetExtensionWnd(FTC_VIEW nView, BOOL bAutoCreate)
{
	ASSERT(nView >= FTCV_FIRSTUIEXTENSION && nView <= FTCV_LASTUIEXTENSION);

	if (nView < FTCV_FIRSTUIEXTENSION || nView > FTCV_LASTUIEXTENSION)
		return NULL;

	// try for existing first
	IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView);

	if (pExtWnd || !bAutoCreate)
		return pExtWnd;

	// sanity checks
	ASSERT(m_tabViews.GetViewHwnd(nView) == NULL);

	VIEWDATA* pData = GetViewData(nView);

	if (!pData)
		return NULL;
	
	// this may take a while
	CWaitCursor cursor;

	// Create the extension window
	int nExtension = (nView - FTCV_FIRSTUIEXTENSION);
	UINT nCtrlID = (IDC_FTC_EXTENSIONWINDOW1 + nExtension);

	pExtWnd = pData->pExtension->CreateExtWindow(nCtrlID, WS_CHILD, 0, 0, 0, 0, GetSafeHwnd());
	
	if (pExtWnd == NULL)
		return NULL;
	
	HWND hWnd = pExtWnd->GetHwnd();
	ASSERT (hWnd);
	
	if (!hWnd)
		return NULL;
	
	pExtWnd->SetUITheme(&m_theme);
	pExtWnd->SetReadOnly(HasStyle(TDCS_READONLY) != FALSE);
	
	// update focus first because initializing views can take time
	::SetFocus(hWnd);
	
	m_aExtViews[nExtension] = pExtWnd;
	
	// restore state
	CPreferences prefs;
	CString sKey = GetPreferencesKey(_T("UIExtensions"));
	
	pExtWnd->LoadPreferences(&prefs, sKey);
	
	// and update tab control with our new HWND
	m_tabViews.SetViewHwnd((FTC_VIEW)nView, hWnd);
	
	// initialize update state
	pData->bNeedTaskUpdate = TRUE;

	// and capabilities
	if (pData->bCanPrepareNewTask == -1)
	{
		CTaskFile task;
		task.NewTask(_T("Test Task"));

		pData->bCanPrepareNewTask = pExtWnd->PrepareNewTask(&task);
	}

	// insert the view after the list in z-order
	::SetWindowPos(pExtWnd->GetHwnd(), m_taskList, 0, 0, 0, 0, (SWP_NOMOVE | SWP_NOSIZE));
	
	Invalidate();

	return pExtWnd;
}

LRESULT CTabbedToDoCtrl::OnPreTabViewChange(WPARAM nOldView, LPARAM nNewView) 
{
	EndLabelEdit(FALSE);
	HandleUnsavedComments(); 		

	// notify parent
	GetParent()->SendMessage(WM_TDCN_VIEWPRECHANGE, nOldView, nNewView);

	// make sure something is selected
	if (GetSelectedCount() == 0)
	{
		HTREEITEM hti = m_taskTree.GetSelectedItem();
		
		if (!hti)
			hti = m_taskTree.GetChildItem(NULL);
		
		CToDoCtrl::SelectTask(GetTaskID(hti));
	}

	// take a note of what task is currently singly selected
	// so that we can prevent unnecessary calls to UpdateControls
	DWORD dwSelTaskID = GetSingleSelectedTaskID();
	
	switch (nNewView)
	{
	case FTCV_TASKTREE:

		// update sort
		if (m_bTreeNeedResort)
		{
			m_bTreeNeedResort = FALSE;
			CToDoCtrl::Resort();
		}

		m_taskTree.EnsureSelectionVisible();
		break;

	case FTCV_TASKLIST:
		{
			// update sort
			VIEWDATA* pLVData = GetViewData(FTCV_TASKLIST);
			
			if (pLVData->bNeedResort)
			{
				pLVData->bNeedResort = FALSE;
				m_taskList.Resort();
			}
			
			ResyncListSelection();
			m_taskList.EnsureSelectionVisible();
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			VIEWDATA* pData = GetViewData((FTC_VIEW)nNewView);

			if (!pData)
				return 1L; // prevent tab change

			// start progress if initializing from another view, 
			// will be cleaned up in OnPostTabViewChange
			UINT nProgressMsg = 0;

			if (nOldView != -1)
			{
				if (GetExtensionWnd((FTC_VIEW)nNewView, FALSE) == NULL)
					nProgressMsg = IDS_INITIALISINGTABBEDVIEW;

				else if (pData->bNeedTaskUpdate)
					nProgressMsg = IDS_UPDATINGTABBEDVIEW;

				if (nProgressMsg)
					BeginExtensionProgress(pData, nProgressMsg);
			}

			IUIExtensionWindow* pExtWnd = GetExtensionWnd((FTC_VIEW)nNewView, TRUE);
			ASSERT(pExtWnd && pExtWnd->GetHwnd());
			
			if (pData->bNeedTaskUpdate)
			{
				// start progress if not already
				// will be cleaned up in OnPostTabViewChange
				if (nProgressMsg == 0)
					BeginExtensionProgress(pData);

				CTaskFile tasks;
				GetAllTasks(tasks);
				
				UpdateExtensionView(pExtWnd, tasks, IUI_ALL);
				pData->bNeedTaskUpdate = FALSE;
			}
				
			// set the selection
			pExtWnd->SelectTask(dwSelTaskID);
		}
		break;
	}

	// update controls only if the selection has changed
	if (HasSingleSelectionChanged(dwSelTaskID))
		UpdateControls();

	return 0L; // allow tab change
}

LRESULT CTabbedToDoCtrl::OnPostTabViewChange(WPARAM nOldView, LPARAM nNewView)
{
	switch (nNewView)
	{
	case FTCV_TASKTREE:
	case FTCV_TASKLIST:
		break;
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		// stop any progress
		GetParent()->SendMessage(WM_TDCM_LENGTHYOPERATION, FALSE);
		
		// resync selection
		UpdateExtensionViewSelection();
		break;
	}
	
	// notify parent
	GetParent()->SendMessage(WM_TDCN_VIEWPOSTCHANGE, nOldView, nNewView);
	
	return 0L;
}

void CTabbedToDoCtrl::UpdateExtensionView(IUIExtensionWindow* pExtWnd, const CTaskFile& tasks, 
										  IUI_UPDATETYPE nType, TDC_ATTRIBUTE nAttrib)
{
	CWaitCursor cursor;
	CAutoFlag af(m_bUpdatingExtensions, TRUE);

	pExtWnd->UpdateTasks(&tasks, nType, nAttrib);
}

void CTabbedToDoCtrl::UpdateExtensionViewSelection()
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
	case FTCV_TASKLIST:
		ASSERT(0);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView);
			ASSERT(pExtWnd && pExtWnd->GetHwnd());

			CDWordArray aTaskIDs;
			int nNumSel = CToDoCtrl::GetSelectedTaskIDs(aTaskIDs);

			if (nNumSel && !pExtWnd->SelectTasks(aTaskIDs.GetData(), nNumSel))
			{
				if (!pExtWnd->SelectTask(aTaskIDs[0]))
				{
					// clear tasklist selection
					TSH().RemoveAll();
					UpdateControls();
				}
			}
		}
		break;

	default:
		ASSERT(0);
	}
}

DWORD CTabbedToDoCtrl::GetSingleSelectedTaskID() const
{
	if (GetSelectedCount() == 1) 
		return GetTaskID(GetSelectedItem());

	// else
	return 0;
}

BOOL CTabbedToDoCtrl::HasSingleSelectionChanged(DWORD dwSelID) const
{
	// multi-selection
	if (GetSelectedCount() != 1)
		return TRUE;

	// different selection
	if (GetTaskID(GetSelectedItem()) != dwSelID)
		return TRUE;

	// dwSelID is still the only selection
	return FALSE;
}

VIEWDATA* CTabbedToDoCtrl::GetViewData(FTC_VIEW nView) const
{
	VIEWDATA* pData = (VIEWDATA*)m_tabViews.GetViewData(nView);

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		ASSERT(pData == NULL);
		break;
		
	case FTCV_TASKLIST:
		ASSERT(pData && !pData->pExtension);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		ASSERT(pData && pData->pExtension);
		break;

	// all else
	default:
		ASSERT(0);
	}

	return pData;
}

VIEWDATA* CTabbedToDoCtrl::GetActiveViewData() const
{
	return GetViewData(GetView());
}

void CTabbedToDoCtrl::SetView(FTC_VIEW nView) 
{
	// take a note of what task is currently singly selected
	// so that we can prevent unnecessary calls to UpdateControls
	DWORD dwSelTaskID = GetSingleSelectedTaskID();
	
	if (!m_tabViews.SetActiveView(nView, TRUE))
		return;

	// update controls only if the selection has changed and 
	if (HasSingleSelectionChanged(dwSelTaskID))
		UpdateControls();
}

void CTabbedToDoCtrl::SetNextView() 
{
	// take a note of what task is currently singly selected
	// so that we can prevent unnecessary calls to UpdateControls
	DWORD dwSelTaskID = GetSingleSelectedTaskID();
	
	m_tabViews.ActivateNextView();

	// update controls only if the selection has changed and 
	if (HasSingleSelectionChanged(dwSelTaskID))
		UpdateControls();
}

LRESULT CTabbedToDoCtrl::OnUIExtSelectTask(WPARAM /*wParam*/, LPARAM lParam)
{
	if (!m_bUpdatingExtensions)
	{
		// check there's an actual change
		DWORD dwTaskID = (DWORD)lParam;

		if (HasSingleSelectionChanged(dwTaskID))
			return SelectTask(dwTaskID);
	}

	// else
	return 0L;
}

LRESULT CTabbedToDoCtrl::OnUIExtSortColumnChange(WPARAM /*wParam*/, LPARAM lParam)
{
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
	case FTCV_TASKLIST:
	default:
		ASSERT(0);
		break;
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			VIEWDATA* pData = GetViewData(nView);
			ASSERT(pData);

			if (pData)
				pData->sort.single.nBy = TDC_COLUMN(lParam);
		}
		break;
	}

	return 0L;
}

LRESULT CTabbedToDoCtrl::OnUIExtEditSelectedTaskTitle(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	BOOL bEdit = EditSelectedTask();
	ASSERT(bEdit);

	return bEdit;
}

LRESULT CTabbedToDoCtrl::OnUIExtModifySelectedTask(WPARAM wParam, LPARAM lParam)
{
	ASSERT(!IsReadOnly());

	if (IsReadOnly())
		return FALSE;

	HandleUnsavedComments();

	CStringArray aValues;
	CBinaryData bdEmpty;

	// prevent change being propagated back to active view
	CAutoFlag af(m_bExtModifyingApp, TRUE);

	BOOL bResult = TRUE;
	BOOL bDependChange = FALSE;
	
	CUndoAction ua(m_data, TDCUAT_EDIT);

	try
	{
		const IUITASKMOD* pMod = (const IUITASKMOD*)lParam;
		int nNumMod = (int)wParam;

		ASSERT(nNumMod > 0);

		for (int nMod = 0; ((nMod < nNumMod) && bResult); nMod++, pMod++)
		{
			switch (pMod->nAttrib)
			{
			case TDCA_TASKNAME:
				bResult &= SetSelectedTaskTitle(pMod->szValue);
				break;

			case TDCA_DONEDATE:
				bResult &= SetSelectedTaskDate(TDCD_DONE, CDateHelper::GetDate(pMod->tValue));
				break;

			case TDCA_DUEDATE:
				if (SetSelectedTaskDate(TDCD_DUE, CDateHelper::GetDate(pMod->tValue)))
				{
					if (HasStyle(TDCS_AUTOADJUSTDEPENDENCYDATES))
						bDependChange = TRUE;
				}
				else
					bResult = FALSE;
				break;

			case TDCA_STARTDATE:
				bResult &= SetSelectedTaskDate(TDCD_START, CDateHelper::GetDate(pMod->tValue));
				break;

			case TDCA_PRIORITY:
				bResult &= SetSelectedTaskPriority(pMod->nValue);
				break;

			case TDCA_COLOR:
				bResult &= SetSelectedTaskColor(pMod->crValue);
				break;

			case TDCA_ALLOCTO:
				Misc::Split(pMod->szValue, aValues);
				bResult &= SetSelectedTaskAllocTo(aValues);
				break;

			case TDCA_ALLOCBY:
				bResult &= SetSelectedTaskAllocBy(pMod->szValue);
				break;

			case TDCA_STATUS:
				bResult &= SetSelectedTaskStatus(pMod->szValue);
				break;

			case TDCA_CATEGORY:
				Misc::Split(pMod->szValue, aValues);
				bResult &= SetSelectedTaskCategories(aValues);
				break;

			case TDCA_TAGS:
				Misc::Split(pMod->szValue, aValues);
				bResult &= SetSelectedTaskTags(aValues);
				break;
				
			case TDCA_FILEREF:
				Misc::Split(pMod->szValue, aValues);
				bResult &= SetSelectedTaskFileRefs(aValues);
				break;

			case TDCA_PERCENT:
				bResult &= SetSelectedTaskPercentDone(pMod->nValue);
				break;

			case TDCA_TIMEEST:
				bResult &= SetSelectedTaskTimeEstimate(pMod->dValue);
				break;

			case TDCA_TIMESPENT:
				bResult &= SetSelectedTaskTimeSpent(pMod->dValue);
				break;

			case TDCA_COMMENTS:
				bResult &= SetSelectedTaskComments(pMod->szValue, bdEmpty);
				break;

			case TDCA_FLAG:
				bResult &= SetSelectedTaskFlag(pMod->bValue);
				break;

			case TDCA_RISK: 
				bResult &= SetSelectedTaskRisk(pMod->nValue);
				break;

			case TDCA_EXTERNALID: 
				bResult &= SetSelectedTaskExtID(pMod->szValue);
				break;

			case TDCA_COST: 
				bResult &= SetSelectedTaskCost(pMod->dValue);
				break;

			case TDCA_DEPENDENCY: 
				Misc::Split(pMod->szValue, aValues);

				if (SetSelectedTaskDependencies(aValues))
					bDependChange = TRUE;
				else
					bResult = FALSE;
				break;

			case TDCA_VERSION:
				bResult &= SetSelectedTaskVersion(pMod->szValue);
				break;

			case TDCA_RECURRENCE: 
			case TDCA_CREATIONDATE:
			case TDCA_CREATEDBY:
			default:
				ASSERT(0);
				return FALSE;
			}
		}
	}
	catch (...)
	{
		ASSERT(0);
		return FALSE;
	}

	// since we prevented changes being propagated back to the active view
	// we may need to update more than the selected task as a consequence
	// of dependency changes
	if (bDependChange)
	{
		// update all tasks
		CTaskFile tasks;
		GetTasks(tasks);
		
		IUIExtensionWindow* pExtWnd = GetExtensionWnd(GetView());
		UpdateExtensionView(pExtWnd, tasks, IUI_EDIT, TDCA_ALL);
	}
	
	return bResult;
}

void CTabbedToDoCtrl::RebuildCustomAttributeUI()
{
	CToDoCtrl::RebuildCustomAttributeUI();

	m_taskList.OnCustomAttributeChange();
}

void CTabbedToDoCtrl::ReposTaskTree(CDeferWndMove* pDWM, const CRect& rPos)
{
	CRect rView;
	m_tabViews.Resize(rPos, pDWM, rView);

	CToDoCtrl::ReposTaskTree(pDWM, rView);
}

void CTabbedToDoCtrl::UpdateTasklistVisibility()
{
	BOOL bTasksVis = (m_nMaxState != TDCMS_MAXCOMMENTS);
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::UpdateTasklistVisibility();
		break;

	case FTCV_TASKLIST:
		m_taskList.ShowWindow(bTasksVis ? SW_SHOW : SW_HIDE);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}

	// handle tab control
	m_tabViews.ShowWindow(bTasksVis && HasStyle(TDCS_SHOWTREELISTBAR) ? SW_SHOW : SW_HIDE);
}

BOOL CTabbedToDoCtrl::OnEraseBkgnd(CDC* pDC)
{
	// clip out tab ctrl
	if (m_tabViews.GetSafeHwnd())
		ExcludeChild(&m_tabViews, pDC);

	return CToDoCtrl::OnEraseBkgnd(pDC);
}

void CTabbedToDoCtrl::Resize(int cx, int cy, BOOL bSplitting)
{
	CToDoCtrl::Resize(cx, cy, bSplitting);
}

BOOL CTabbedToDoCtrl::WantTaskContextMenu() const
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
	case FTCV_TASKLIST:
		return TRUE;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		return TRUE;

	default:
		ASSERT(0);
	}

	return FALSE;
}

BOOL CTabbedToDoCtrl::GetSelectionBoundingRect(CRect& rSelection) const
{
	rSelection.SetRectEmpty();
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::GetSelectionBoundingRect(rSelection);
		break;

	case FTCV_TASKLIST:
		if (m_taskList.GetSelectionBoundingRect(rSelection))
		{
			m_taskList.ClientToScreen(rSelection);
			ScreenToClient(rSelection);
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView);
			ASSERT(pExtWnd);
			
			if (pExtWnd)
			{
				pExtWnd->GetLabelEditRect(rSelection);
				ScreenToClient(rSelection);
			}
		}
		break;

	default:
		ASSERT(0);
	}

	return (!rSelection.IsRectEmpty());
}

void CTabbedToDoCtrl::SelectAll()
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::SelectAll();
		break;

	case FTCV_TASKLIST:
		{
			int nNumItems = m_taskList.GetTaskCount();
			BOOL bAllTasks = (CToDoCtrl::GetTaskCount() == (UINT)nNumItems);

			// select items in tree
			if (bAllTasks)
			{
				CToDoCtrl::SelectAll();
				m_taskList.SelectAll();
			}
			else
			{
				// save IDs only not showing all tasks
				CDWordArray aTaskIDs;

				for (int nItem = 0; nItem < nNumItems; nItem++)
					aTaskIDs.Add(m_taskList.GetTaskID(nItem));

				m_taskList.SelectTasks(aTaskIDs);
			}
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}
}

int CTabbedToDoCtrl::GetTasks(CTaskFile& tasks, const TDCGETTASKS& filter) const
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::GetTasks(tasks, filter);

	case FTCV_TASKLIST:
		{
			PrepareTaskfileForTasks(tasks, filter);

			// we return exactly what's selected in the list and in the same order
			// so we make sure the filter includes TDCGT_NOTSUBTASKS
			for (int nItem = 0; nItem < m_taskList.GetItemCount(); nItem++)
			{
				HTREEITEM hti = GetTreeItem(nItem);
				DWORD dwParentID = m_data.GetTaskParentID(GetTaskID(hti));
				
				AddTreeItemToTaskFile(hti, tasks, NULL, filter, FALSE, dwParentID);
			}

			return tasks.GetTaskCount();
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		return CToDoCtrl::GetTasks(tasks, filter); // for now

	default:
		ASSERT(0);
	}

	return 0;
}

int CTabbedToDoCtrl::GetSelectedTasks(CTaskFile& tasks, const TDCGETTASKS& filter, DWORD dwFlags) const
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::GetSelectedTasks(tasks, filter, dwFlags);

	case FTCV_TASKLIST:
		{
			PrepareTaskfileForTasks(tasks, filter);

			// we return exactly what's selected in the list and in the same order
			POSITION pos = m_taskList.List().GetFirstSelectedItemPosition();

			while (pos)
			{
				int nItem = m_taskList.List().GetNextSelectedItem(pos);

				HTREEITEM hti = GetTreeItem(nItem);
				DWORD dwParentID = m_data.GetTaskParentID(GetTaskID(nItem));
				HTASKITEM htParent = NULL;
				
				// Add immediate parent as required.
				// Note: we can assume that the selected task is always added successfully
				if (dwParentID && (dwFlags & TDCGSTF_IMMEDIATEPARENT))
				{
					HTREEITEM htiParent = m_taskTree.GetParentItem(hti);
					DWORD dwParentsParentID = m_data.GetTaskParentID(dwParentID);
					
					if (AddTreeItemToTaskFile(htiParent, tasks, NULL, filter, FALSE, dwParentsParentID))  // FALSE == no subtasks
						htParent = tasks.FindTask(dwParentID);
				}

				VERIFY(AddTreeItemToTaskFile(hti, tasks, NULL, filter, FALSE, dwParentID)); // FALSE == no subtasks
			}

			return tasks.GetTaskCount();
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		return CToDoCtrl::GetSelectedTasks(tasks, filter, dwFlags); // for now

	default:
		ASSERT(0);
	}

	return 0;
}

BOOL CTabbedToDoCtrl::RedrawReminders() const
{ 
 	if (CToDoCtrl::RedrawReminders())
	{
		if (InListView())
 			m_taskList.RedrawColumn(TDCC_REMINDER);
	}

	return FALSE;
}

void CTabbedToDoCtrl::SetMaxInfotipCommentsLength(int nLength)
{
	CToDoCtrl::SetMaxInfotipCommentsLength(nLength);

	m_taskList.SetMaxInfotipCommentsLength(nLength);
}

void CTabbedToDoCtrl::SetEditTitleTaskID(DWORD dwTaskID)
{
	CToDoCtrl::SetEditTitleTaskID(dwTaskID);
	
	m_taskList.SetEditTitleTaskID(dwTaskID);
}

BOOL CTabbedToDoCtrl::SetStyle(TDC_STYLE nStyle, BOOL bOn, BOOL bWantUpdate)
{
	// base class processing
	if (CToDoCtrl::SetStyle(nStyle, bOn, bWantUpdate))
	{
		// post-precessing
		switch (nStyle)
		{
		case TDCS_SHOWTREELISTBAR:
			m_tabViews.ShowTabControl(bOn);

			if (bWantUpdate)
				Resize();
			break;

		// detect preferences that can affect task text colors
		// and then handle this in NotifyEndPreferencesUpdate
		case TDCS_COLORTEXTBYPRIORITY:
		case TDCS_COLORTEXTBYATTRIBUTE:
		case TDCS_COLORTEXTBYNONE:
		case TDCS_TREATSUBCOMPLETEDASDONE:
		case TDCS_USEEARLIESTDUEDATE:
		case TDCS_USELATESTDUEDATE:
		case TDCS_USEEARLIESTSTARTDATE:
		case TDCS_USELATESTSTARTDATE:
		case TDCS_USEHIGHESTPRIORITY:
		case TDCS_INCLUDEDONEINPRIORITYCALC:
		case TDCS_DUEHAVEHIGHESTPRIORITY:
		case TDCS_DONEHAVELOWESTPRIORITY:
		case TDCS_TASKCOLORISBACKGROUND:
			m_bTaskColorChange = TRUE;
			break;
		}

		// notify list-list to update itself
		m_taskList.OnStyleUpdated(nStyle, bOn, bWantUpdate);

		return TRUE;
	}

	return FALSE;
}

void CTabbedToDoCtrl::SetPriorityColors(const CDWordArray& aColors)
{
	if (m_taskList.SetPriorityColors(aColors))
		m_bTaskColorChange = TRUE;

	CToDoCtrl::SetPriorityColors(aColors);
}

void CTabbedToDoCtrl::SetCompletedTaskColor(COLORREF color)
{
	if (m_taskList.SetCompletedTaskColor(color))
		m_bTaskColorChange = TRUE;

	CToDoCtrl::SetCompletedTaskColor(color);
}

void CTabbedToDoCtrl::SetFlaggedTaskColor(COLORREF color)
{
	if (m_taskList.SetFlaggedTaskColor(color))
		m_bTaskColorChange = TRUE;

	CToDoCtrl::SetFlaggedTaskColor(color);
}

void CTabbedToDoCtrl::SetReferenceTaskColor(COLORREF color)
{
	if (m_taskList.SetReferenceTaskColor(color))
		m_bTaskColorChange = TRUE;

	CToDoCtrl::SetReferenceTaskColor(color);
}

void CTabbedToDoCtrl::SetAttributeColors(TDC_ATTRIBUTE nAttrib, const CTDCColorMap& colors)
{
	if (m_taskList.SetAttributeColors(nAttrib, colors))
		m_bTaskColorChange = TRUE;

	CToDoCtrl::SetAttributeColors(nAttrib, colors);
}

void CTabbedToDoCtrl::SetStartedTaskColors(COLORREF crStarted, COLORREF crStartedToday)
{
	if (m_taskList.SetStartedTaskColors(crStarted, crStartedToday))
		m_bTaskColorChange = TRUE;

	CToDoCtrl::SetStartedTaskColors(crStarted, crStartedToday);
}

void CTabbedToDoCtrl::SetDueTaskColors(COLORREF crDue, COLORREF crDueToday)
{
	if (m_taskList.SetDueTaskColors(crDue, crDueToday))
		m_bTaskColorChange = TRUE;

	CToDoCtrl::SetDueTaskColors(crDue, crDueToday);
}

void CTabbedToDoCtrl::SetGridlineColor(COLORREF color)
{
	m_taskList.SetGridlineColor(color);
	
	CToDoCtrl::SetGridlineColor(color);
}

void CTabbedToDoCtrl::SetAlternateLineColor(COLORREF color)
{	
	m_taskList.SetAlternateLineColor(color);

	CToDoCtrl::SetAlternateLineColor(color);
}

void CTabbedToDoCtrl::NotifyBeginPreferencesUpdate()
{
	// base class
	CToDoCtrl::NotifyBeginPreferencesUpdate();

	// nothing else for us to do
}

void CTabbedToDoCtrl::NotifyEndPreferencesUpdate()
{
	// base class
	CToDoCtrl::NotifyEndPreferencesUpdate();

	// notify extension windows
	if (HasAnyExtensionViews())
	{
		// we need to update in 2 ways:
		// 1. Tell the extensions that application settings have changed
		// 2. Refresh tasks if their text colour may have changed
		CPreferences prefs;
		CString sKey = GetPreferencesKey(_T("UIExtensions"));
		
		int nExt = m_aExtViews.GetSize();
		FTC_VIEW nCurView = GetView();

		while (nExt--)
		{
			FTC_VIEW nExtView = (FTC_VIEW)(FTCV_FIRSTUIEXTENSION + nExt);
			IUIExtensionWindow* pExtWnd = m_aExtViews[nExt];
			
			if (pExtWnd)
			{
				VIEWDATA* pData = GetViewData(nExtView);

				if (!pData)
					continue;

				// if this extension is active and wants a 
				// color update we want to start progress
				BOOL bWantColorUpdate = (m_bTaskColorChange && pExtWnd->WantUpdate(TDCA_COLOR));

				if (bWantColorUpdate && nExtView == nCurView)
					BeginExtensionProgress(pData);

				// notify all extensions of prefs change
				pExtWnd->LoadPreferences(&prefs, sKey, TRUE);

				// Update task colours if necessary
				if (bWantColorUpdate)
				{
					if (nExtView == nCurView)
					{
						TDCGETTASKS filter;
						filter.aAttribs.Add(TDCA_COLOR);
						
						CTaskFile tasks;
						GetTasks(tasks, filter);
						
						UpdateExtensionView(pExtWnd, tasks, IUI_EDIT, TDCA_COLOR);
						pData->bNeedTaskUpdate = FALSE;
					}
					else // mark for update
					{
						pData->bNeedTaskUpdate = TRUE;
					}
				}

				// cleanup progress
				if (bWantColorUpdate && nExtView == nCurView)
					EndExtensionProgress();
			}
		}
	}

	m_bTaskColorChange = FALSE; // always
}

void CTabbedToDoCtrl::BeginExtensionProgress(const VIEWDATA* pData, UINT nMsg)
{
	ASSERT(pData);

	if (nMsg == 0)
		nMsg = IDS_UPDATINGTABBEDVIEW;

	CEnString sMsg(nMsg, pData->pExtension->GetMenuText());
	GetParent()->SendMessage(WM_TDCM_LENGTHYOPERATION, TRUE, (LPARAM)(LPCTSTR)sMsg);
}

void CTabbedToDoCtrl::EndExtensionProgress()
{
	GetParent()->SendMessage(WM_TDCM_LENGTHYOPERATION, FALSE);
}

void CTabbedToDoCtrl::BeginTimeTracking(DWORD dwTaskID)
{
	CToDoCtrl::BeginTimeTracking(dwTaskID);

	m_taskList.SetTimeTrackTaskID(m_dwTimeTrackTaskID);
}

void CTabbedToDoCtrl::EndTimeTracking(BOOL bAllowConfirm)
{
	CToDoCtrl::EndTimeTracking(bAllowConfirm);
	
	m_taskList.SetTimeTrackTaskID(m_dwTimeTrackTaskID);
}

CString CTabbedToDoCtrl::GetControlDescription(const CWnd* pCtrl) const
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		break; // handled below
		
	case FTCV_TASKLIST:
		if (CDialogHelper::IsChildOrSame(m_taskList, pCtrl->GetSafeHwnd()))
			return CEnString(IDS_LISTVIEW);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		if (pCtrl)
		{
			HWND hwnd = m_tabViews.GetViewHwnd(nView);

			if (CDialogHelper::IsChildOrSame(hwnd, pCtrl->GetSafeHwnd()))
			{
				return m_tabViews.GetViewName(nView);
			}
		}
		break;

	default:
		ASSERT(0);
	}

	return CToDoCtrl::GetControlDescription(pCtrl);
}

BOOL CTabbedToDoCtrl::DeleteSelectedTask(BOOL bWarnUser, BOOL bResetSel)
{
	BOOL bDel = CToDoCtrl::DeleteSelectedTask(bWarnUser, bResetSel);

	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		// handled above
		break;

	case FTCV_TASKLIST:
		if (bDel)
		{
			// work out what to select
			DWORD dwSelID = GetSelectedTaskID();
			int nSel = m_taskList.FindTaskItem(dwSelID);

			if (nSel == -1 && m_taskList.GetItemCount())
				nSel = 0;

			if (nSel != -1)
				SelectTask(m_taskList.GetTaskID(nSel));
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}

	return bDel;
}

HTREEITEM CTabbedToDoCtrl::CreateNewTask(const CString& sText, TDC_INSERTWHERE nWhere, BOOL bEditText)
{
	HTREEITEM htiNew = NULL;
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		htiNew = CToDoCtrl::CreateNewTask(sText, nWhere, bEditText);
		break;

	case FTCV_TASKLIST:
		htiNew = CToDoCtrl::CreateNewTask(sText, nWhere, FALSE); // note FALSE

		if (htiNew)
		{
			DWORD dwTaskID = GetTaskID(htiNew);
			ASSERT(dwTaskID == m_dwNextUniqueID - 1);
			
			// make the new task appear
			RebuildList(NULL); 
			
			// select that task
			SelectTask(dwTaskID);
			
			if (bEditText)
			{
				m_dwLastAddedID = dwTaskID;
				EditSelectedTask(TRUE);
			}
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		htiNew = CToDoCtrl::CreateNewTask(sText, nWhere, FALSE); // note FALSE

		if (htiNew)
		{
			DWORD dwTaskID = GetTaskID(htiNew);
			ASSERT(dwTaskID == m_dwNextUniqueID - 1);
			
			// select that task
			SelectTask(dwTaskID);
			
			if (bEditText)
			{
				m_dwLastAddedID = dwTaskID;
				EditSelectedTask(TRUE);
			}
		}
		break;

	default:
		ASSERT(0);
	}

	return htiNew;
}

TODOITEM* CTabbedToDoCtrl::CreateNewTask(HTREEITEM htiParent)
{
	TODOITEM* pTDI = CToDoCtrl::CreateNewTask(htiParent);
	ASSERT(pTDI);

	// give active extension view a chance to initialise
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView);
			ASSERT(pExtWnd);

			if (pExtWnd)
			{
				CTaskFile task;
				HTASKITEM hTask = task.NewTask(pTDI->sTitle);

				task.SetTaskAttributes(hTask, pTDI);

				if (pExtWnd->PrepareNewTask(&task))
					task.GetTaskAttributes(hTask, pTDI);

				// fall thru
			}
		}
		break;
	}

	return pTDI;
}

BOOL CTabbedToDoCtrl::CanCreateNewTask(TDC_INSERTWHERE nInsertWhere) const
{
	FTC_VIEW nView = GetView();

	BOOL bCanCreate = CToDoCtrl::CanCreateNewTask(nInsertWhere);

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
	case FTCV_TASKLIST:
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			const VIEWDATA* pData = GetViewData(nView);

			if (pData)
				bCanCreate &= pData->bCanPrepareNewTask;
		}
		break;

	default:
		bCanCreate = FALSE;
		ASSERT(0);
		break;
	}

	return bCanCreate;
}

void CTabbedToDoCtrl::RebuildList(const void* pContext)
{
	if (!m_data.GetTaskCount())
	{
		m_taskList.DeleteAll(); 
		return;
	}

	// cache current selection
	TDCSELECTIONCACHE cache;
	CacheListSelection(cache);

	// note: the call to RestoreListSelection at the bottom fails if the 
	// list has redraw disabled so it must happen outside the scope of hr2
	{
		CHoldRedraw hr(GetSafeHwnd());
		CHoldRedraw hr2(m_taskList);
		CWaitCursor cursor;

		// remove all existing items
		m_taskList.DeleteAll();
		
		// rebuild the list from the tree
		AddTreeItemToList(NULL, pContext);

		m_taskList.SetNextUniqueTaskID(m_dwNextUniqueID);

		// redo last sort
		FTC_VIEW nView = GetView();

		switch (nView)
		{
		case FTCV_TASKTREE:
		case FTCV_UNSET:
			break;

		case FTCV_TASKLIST:
			if (IsSorting())
			{
				GetViewData(FTCV_TASKLIST)->bNeedResort = FALSE;
				Resort();
			}
			break;

		case FTCV_UIEXTENSION1:
		case FTCV_UIEXTENSION2:
		case FTCV_UIEXTENSION3:
		case FTCV_UIEXTENSION4:
		case FTCV_UIEXTENSION5:
		case FTCV_UIEXTENSION6:
		case FTCV_UIEXTENSION7:
		case FTCV_UIEXTENSION8:
		case FTCV_UIEXTENSION9:
		case FTCV_UIEXTENSION10:
		case FTCV_UIEXTENSION11:
		case FTCV_UIEXTENSION12:
		case FTCV_UIEXTENSION13:
		case FTCV_UIEXTENSION14:
		case FTCV_UIEXTENSION15:
		case FTCV_UIEXTENSION16:
			break;

		default:
			ASSERT(0);
		}
	}
	
	// restore selection
	ResyncListSelection();
	
	// don't update controls if only one item is selected and it did not
	// change as a result of the filter
	if (!(GetSelectedCount() == 1 && 
		cache.aSelTaskIDs.GetSize() == 1 &&
		GetTaskID(GetSelectedItem()) == cache.aSelTaskIDs[0]))
	{
		UpdateControls();
	}
}

int CTabbedToDoCtrl::AddItemToList(DWORD dwTaskID)
{
	// omit task references from list
	if (CToDoCtrl::IsTaskReference(dwTaskID))
		return -1;

	// else
	return m_taskList.List().InsertItem(LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM | LVIF_STATE, 
							m_taskList.GetItemCount(), 
							LPSTR_TEXTCALLBACK, 
							0,
							LVIS_STATEIMAGEMASK,
							I_IMAGECALLBACK, 
							dwTaskID);
}

void CTabbedToDoCtrl::AddTreeItemToList(HTREEITEM hti, const void* pContext)
{
	// add task
	if (hti)
	{
		// if the add fails then it's a task reference
		if (CTabbedToDoCtrl::AddItemToList(GetTaskID(hti)) == -1)
		{
			return; 
		}
	}

	// children
	HTREEITEM htiChild = m_taskTree.GetChildItem(hti);

	while (htiChild)
	{
		AddTreeItemToList(htiChild, pContext);
		htiChild = m_taskTree.GetNextItem(htiChild);
	}
}

void CTabbedToDoCtrl::SetExtensionsNeedUpdate(BOOL bUpdate, FTC_VIEW nIgnore)
{
	for (int nExt = 0; nExt < m_aExtViews.GetSize(); nExt++)
	{
		FTC_VIEW nView = (FTC_VIEW)(FTCV_UIEXTENSION1 + nExt);
		
		if (nView == nIgnore)
			continue;
		
		// else
		VIEWDATA* pData = GetViewData(nView);
		
		if (pData)
			pData->bNeedTaskUpdate = bUpdate;
	}
}

void CTabbedToDoCtrl::SetModified(BOOL bMod, TDC_ATTRIBUTE nAttrib, DWORD dwModTaskID)
{
	CToDoCtrl::SetModified(bMod, nAttrib, dwModTaskID);

	if (bMod)
	{
		m_taskList.SetModified(nAttrib); // always

		FTC_VIEW nView = GetView();
		
		switch (nView)
		{
		case FTCV_TASKLIST:
			{
				switch (nAttrib)
				{
				case TDCA_DELETE:
					if (m_taskTree.GetItemCount())
						m_taskList.RemoveDeletedItems();
					else
						m_taskList.DeleteAll();
					break;

				case TDCA_ARCHIVE:
					m_taskList.RemoveDeletedItems();
					break;
					
				case TDCA_NEWTASK:
				case TDCA_UNDO:
				case TDCA_PASTE:
					RebuildList(NULL);
					break;
					
				default: // all other attributes
					m_taskList.InvalidateSelection();
				}
			}
			break;
			
		case FTCV_TASKTREE:
		case FTCV_UNSET:
			GetViewData(FTCV_TASKLIST)->bNeedTaskUpdate = TRUE;
			break;
			
		case FTCV_UIEXTENSION1:
		case FTCV_UIEXTENSION2:
		case FTCV_UIEXTENSION3:
		case FTCV_UIEXTENSION4:
		case FTCV_UIEXTENSION5:
		case FTCV_UIEXTENSION6:
		case FTCV_UIEXTENSION7:
		case FTCV_UIEXTENSION8:
		case FTCV_UIEXTENSION9:
		case FTCV_UIEXTENSION10:
		case FTCV_UIEXTENSION11:
		case FTCV_UIEXTENSION12:
		case FTCV_UIEXTENSION13:
		case FTCV_UIEXTENSION14:
		case FTCV_UIEXTENSION15:
		case FTCV_UIEXTENSION16:
			// handled below
			break;
			
		default:
			ASSERT(0);
		}
		
		UpdateExtensionViews(nAttrib, dwModTaskID);
	}
}

void CTabbedToDoCtrl::UpdateExtensionViews(TDC_ATTRIBUTE nAttrib, DWORD dwTaskID)
{
	if (!HasAnyExtensionViews())
		return;

	FTC_VIEW nCurView = GetView();

	switch (nAttrib)
	{
	// for a simple attribute change (or addition) update all extensions
	// at the same time so that they won't need updating when the user switches view
	case TDCA_TASKNAME:
		// Initial edit of new task is special case
		if (m_dwLastAddedID == GetSelectedTaskID())
		{
			UpdateExtensionViews(TDCA_NEWTASK, m_dwLastAddedID); // RECURSIVE CALL
			return;
		}
		// else fall thru

	case TDCA_DONEDATE:
	case TDCA_DUEDATE:
	case TDCA_STARTDATE:
	case TDCA_PRIORITY:
	case TDCA_COLOR:
	case TDCA_ALLOCTO:
	case TDCA_ALLOCBY:
	case TDCA_STATUS:
	case TDCA_CATEGORY:
	case TDCA_TAGS:
	case TDCA_PERCENT:
	case TDCA_TIMEEST:
	case TDCA_TIMESPENT:
	case TDCA_FILEREF:
	case TDCA_COMMENTS:
	case TDCA_FLAG:
	case TDCA_CREATIONDATE:
	case TDCA_CREATEDBY:
	case TDCA_RISK: 
	case TDCA_EXTERNALID: 
	case TDCA_COST: 
	case TDCA_DEPENDENCY: 
	case TDCA_RECURRENCE: 
	case TDCA_VERSION:
		{	
			// Check extensions for this OR a color change
			BOOL bColorChange = m_taskList.ModCausesTaskTextColorChange(nAttrib);

			if (!AnyExtensionViewWantsChange(nAttrib))
			{
				// if noone wants either we can stop
				if (!bColorChange || !AnyExtensionViewWantsChange(TDCA_COLOR))
					return;
			}

			// note: we need to get 'True' tasks and all their parents
			// because of the possibility of colour changes
			CTaskFile tasks;
			DWORD dwFlags = TDCGSTF_RESOLVEREFERENCES;

			// don't include subtasks unless the completion date changed
			if (nAttrib != TDCA_DONEDATE)
				dwFlags |= TDCGSTF_NOTSUBTASKS;

			if (bColorChange)
				dwFlags |= TDCGSTF_ALLPARENTS;

			// update all tasks if the selected tasks have 
			// dependents and the due dates was modified
			if ((nAttrib == TDCA_DUEDATE) && m_taskTree.SelectionHasDependents())
			{
				GetTasks(tasks);
				nAttrib = TDCA_ALL; // so that start date changes get picked up
			}
			else
			{
				CToDoCtrl::GetSelectedTasks(tasks, TDCGT_ALL, dwFlags);
			}
			
			// refresh all extensions 
			int nExt = m_aExtViews.GetSize();
			
			while (nExt--)
			{
				IUIExtensionWindow* pExtWnd = m_aExtViews[nExt];
				
				if (pExtWnd)
				{
					if (ExtensionViewWantsChange(nExt, nAttrib))
					{
						UpdateExtensionView(pExtWnd, tasks, IUI_EDIT, nAttrib);
					}
					else if (bColorChange && ExtensionViewWantsChange(nExt, TDCA_COLOR))
					{
						UpdateExtensionView(pExtWnd, tasks, IUI_EDIT, TDCA_COLOR);
					}
				}

				// clear the update flag for all created extensions
				FTC_VIEW nExtView = (FTC_VIEW)(FTCV_FIRSTUIEXTENSION + nExt);
				VIEWDATA* pData = GetViewData(nExtView);

				if (pData)
					pData->bNeedTaskUpdate = (pExtWnd == NULL);
			}
		}
		break;
		
	// refresh the current view (if it's an extension) 
	// and mark the others as needing updates.
	case TDCA_NEWTASK: 
	case TDCA_DELETE:
	case TDCA_UNDO:
	case TDCA_POSITION: // == move
	case TDCA_PASTE:
	case TDCA_ARCHIVE:
		{
			int nExt = m_aExtViews.GetSize();
			
			while (nExt--)
			{
				FTC_VIEW nView = (FTC_VIEW)(FTCV_FIRSTUIEXTENSION + nExt);
				VIEWDATA* pData = GetViewData(nView);

				if (pData)
				{
					IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView);

					if (pExtWnd && (nView == nCurView))
					{
						BeginExtensionProgress(pData);

						// update all tasks
						CTaskFile tasks;
						GetTasks(tasks);

						IUI_UPDATETYPE nUpdate = IUI_ALL;

						if ((nAttrib == TDCA_DELETE) || (nAttrib == TDCA_ARCHIVE))
							nUpdate = IUI_DELETE;
						
						UpdateExtensionView(pExtWnd, tasks, nUpdate);
						pData->bNeedTaskUpdate = FALSE;

						if ((nAttrib == TDCA_NEWTASK) && dwTaskID)
							pExtWnd->SelectTask(dwTaskID);
						else
							UpdateExtensionViewSelection();

						EndExtensionProgress();
					}
					else
					{
						pData->bNeedTaskUpdate = TRUE;
					}
				}
			}
		}
		break;	
		
	case TDCA_PROJNAME:
	case TDCA_ENCRYPT:
	default:
		// do nothing
		break;
	}
}

BOOL CTabbedToDoCtrl::ExtensionViewWantsChange(int nExt, TDC_ATTRIBUTE nAttrib) const
{
	FTC_VIEW nCurView = GetView();
	FTC_VIEW nExtView = (FTC_VIEW)(FTCV_FIRSTUIEXTENSION + nExt);

	// if the window is not active and is already marked
	// for a full update then we don't need to do
	// anything more because it will get this update when
	// it is next activated
	if (nExtView != nCurView)
	{
		const VIEWDATA* pData = GetViewData(nExtView);

		if (!pData || pData->bNeedTaskUpdate)
			return FALSE;
	}
	else // active view
	{
		// if this update has come about as a consequence
		// of this extension window modifying the selected 
		// task, then we assume that it won't want the update
		if (m_bExtModifyingApp)
			return FALSE;
	}
	
	IUIExtensionWindow* pExtWnd = m_aExtViews[nExt];
	ASSERT(pExtWnd);
	
	return (pExtWnd && pExtWnd->WantUpdate(nAttrib));
}

BOOL CTabbedToDoCtrl::AnyExtensionViewWantsChange(TDC_ATTRIBUTE nAttrib) const
{
	// find the first extension wanting this change
	FTC_VIEW nCurView = GetView();
	int nExt = m_aExtViews.GetSize();
	
	while (nExt--)
	{
		if (ExtensionViewWantsChange(nExt, nAttrib))
			return TRUE;
	}

	// not found
	return FALSE;
}

void CTabbedToDoCtrl::ResortSelectedTaskParents()
{
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::ResortSelectedTaskParents();
		break;
		
	case FTCV_TASKLIST:
		m_taskList.Resort(); // resort all
		break;
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		// resorting handled by the extension itself
		break;
		
	default:
		ASSERT(0);
	}
}

BOOL CTabbedToDoCtrl::ModNeedsResort(TDC_ATTRIBUTE nModType) const
{
	if (!HasStyle(TDCS_RESORTONMODIFY))
		return FALSE;

	VIEWDATA* pLVData = GetViewData(FTCV_TASKLIST);

	BOOL bTreeNeedsResort = CToDoCtrl::ModNeedsResort(nModType);
	BOOL bListNeedsResort = m_taskList.ModNeedsResort(nModType);
	
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		pLVData->bNeedResort |= bListNeedsResort;
		return bTreeNeedsResort;

	case FTCV_TASKLIST:
		m_bTreeNeedResort |= bTreeNeedsResort;
		return bListNeedsResort;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		// resorting handled by the extension itself
		m_bTreeNeedResort |= bTreeNeedsResort;
		pLVData->bNeedResort |= bListNeedsResort;
		break;

	default:
		ASSERT(0);
	}

	return FALSE;
}

const TDSORT& CTabbedToDoCtrl::GetSort() const 
{ 
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return m_taskTree.GetSort();
		
	case FTCV_TASKLIST:
		return m_taskList.GetSort();
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			VIEWDATA* pVData = GetActiveViewData();
			ASSERT(pVData);
			
			if (pVData)
				return pVData->sort; 
		}
		break;
		
	default:
		ASSERT(0);
	}

	static TDSORT sort;
	return sort;
}

TDC_COLUMN CTabbedToDoCtrl::GetSortBy() const
{
	return GetSort().single.nBy;
}

void CTabbedToDoCtrl::GetSortBy(TDSORTCOLUMNS& sort) const
{
	sort = GetSort().multi;
}

BOOL CTabbedToDoCtrl::SelectTask(DWORD dwTaskID, BOOL bTrue)
{
	BOOL bRes = CToDoCtrl::SelectTask(dwTaskID, bTrue);

	// check task has not been filtered out
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		// handled above
		break;

	case FTCV_TASKLIST:
		{
			int nItem = m_taskList.FindTaskItem(dwTaskID);

			if (nItem == -1)
			{
				ASSERT(0);
				return FALSE;
			}
			
			// remove focused state from existing task
			int nFocus = m_taskList.List().GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);

			if (nFocus != -1)
				m_taskList.List().SetItemState(nFocus, 0, LVIS_SELECTED | LVIS_FOCUSED);

			m_taskList.List().SetItemState(nItem, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
			m_taskList.EnsureSelectionVisible();
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			IUIExtensionWindow* pExtWnd = GetExtensionWnd(nView);
			ASSERT(pExtWnd);

			if (pExtWnd)
				pExtWnd->SelectTask(dwTaskID);
		}
		break;

	default:
		ASSERT(0);
	}

	return bRes;
}

LRESULT CTabbedToDoCtrl::OnEditCancel(WPARAM wParam, LPARAM lParam)
{
	// check if we need to delete the just added item
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		// handled below
		break;

	case FTCV_TASKLIST:
		// delete the just added task
		if (GetSelectedTaskID() == m_dwLastAddedID)
		{
			int nDelItem = m_taskList.GetFirstSelectedItem();
			m_taskList.List().DeleteItem(nDelItem);
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		// delete the just added task from the active view
		if (GetSelectedTaskID() == m_dwLastAddedID)
		{
			LRESULT lr = CToDoCtrl::OnEditCancel(wParam, lParam);
			UpdateExtensionViews(TDCA_DELETE);
		}
		break;

	default:
		ASSERT(0);
	}

	return CToDoCtrl::OnEditCancel(wParam, lParam);
}

void CTabbedToDoCtrl::CacheListSelection(TDCSELECTIONCACHE& cache) const
{
	m_taskList.CacheSelection(cache);
}

void CTabbedToDoCtrl::RestoreListSelection(const TDCSELECTIONCACHE& cache)
{
	m_taskList.RestoreSelection(cache);
}

BOOL CTabbedToDoCtrl::SetTreeFont(HFONT hFont)
{
	if (CToDoCtrl::SetTreeFont(hFont))
	{
		if (!hFont) // set to our font
		{
			// for some reason i can not yet explain, our font
			// is not correctly set so we use our parent's font instead
			// hFont = (HFONT)SendMessage(WM_GETFONT);
			hFont = (HFONT)GetParent()->SendMessage(WM_GETFONT);
		}

		VERIFY(m_taskList.SetFont(hFont));

		// other views
		// TODO

		return TRUE;
	}

	// else
	return FALSE;
}

BOOL CTabbedToDoCtrl::AddView(IUIExtension* pExtension)
{
	if (!pExtension)
		return FALSE;

	// remove any existing views of this type
	RemoveView(pExtension);

	// add to tab control
	HICON hIcon = pExtension->GetIcon();
	CEnString sName(pExtension->GetMenuText());

	int nIndex = m_aExtViews.GetSize();
	FTC_VIEW nView = (FTC_VIEW)(FTCV_UIEXTENSION1 + nIndex);

	VIEWDATA* pData = NewViewData();
	ASSERT(pData);

	pData->pExtension = pExtension;

	// we pass NULL for the hWnd because we are going to load
	// only on demand
	if (m_tabViews.AttachView(NULL, nView, sName, hIcon, pData))
	{
		m_aExtViews.Add(NULL); // placeholder
		return TRUE;
	}

	return FALSE;
}

BOOL CTabbedToDoCtrl::RemoveView(IUIExtension* pExtension)
{
	// search for any views having this type
	int nView = m_aExtViews.GetSize();

	while (nView--)
	{
		IUIExtensionWindow* pExtWnd = m_aExtViews[nView];

		if (pExtWnd) // can be NULL
		{
			CString sExtType = pExtension->GetTypeID();
			CString sExtWndType = pExtWnd->GetTypeID();

			if (sExtType == sExtWndType)
			{
				VERIFY (m_tabViews.DetachView(pExtWnd->GetHwnd()));
				pExtWnd->Release();

				m_aExtViews.RemoveAt(nView);

				return TRUE;
			}
		}
	}

	return FALSE;
}

void CTabbedToDoCtrl::OnTabCtrlRClick(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
	CMenu menu;

	if (menu.LoadMenu(IDR_TASKVIEWVISIBILITY))
	{
		CMenu* pPopup = menu.GetSubMenu(0);
		CPoint ptCursor(GetMessagePos());

		// prepare list view
		// NOTE: task tree is already prepared
		pPopup->CheckMenuItem(ID_SHOWVIEW_LISTVIEW, IsListViewTabShowing() ? MF_CHECKED : 0);

		// extension views
		CStringArray aTypeIDs;
		GetVisibleExtensionViews(aTypeIDs);

		CTDCUIExtensionHelper::PrepareViewVisibilityMenu(pPopup, m_mgrUIExt, aTypeIDs);

		UINT nCmdID = ::TrackPopupMenu(*pPopup, TPM_RETURNCMD | TPM_LEFTALIGN | TPM_LEFTBUTTON, 
										ptCursor.x, ptCursor.y, 0, GetSafeHwnd(), NULL);

		m_tabViews.Invalidate(FALSE);
		m_tabViews.UpdateWindow();

		switch (nCmdID)
		{
		case ID_SHOWVIEW_TASKTREE:
			ASSERT(0); // this is not accessible
			break;
			
		case ID_SHOWVIEW_LISTVIEW:
			ShowListViewTab(!IsListViewTabShowing());
			break;
			
		default:
			if (CTDCUIExtensionHelper::ProcessViewVisibilityMenuCmd(nCmdID, m_mgrUIExt, aTypeIDs))
				SetVisibleExtensionViews(aTypeIDs);
			break;
		}
	}
	
	*pResult = 0;
}


LRESULT CTabbedToDoCtrl::OnColumnEditClick(WPARAM wParam, LPARAM lParam)
{
	if (InListView())
		UpdateTreeSelection();

	return CToDoCtrl::OnColumnEditClick(wParam, lParam);
}

void CTabbedToDoCtrl::OnListSelChanged(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
	*pResult = 0;

	UpdateTreeSelection();
}

void CTabbedToDoCtrl::OnListClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	// special case - ALT key
	if (Misc::IsKeyPressed(VK_MENU))
	{
		LPNMITEMACTIVATE pNMIA = (LPNMITEMACTIVATE)pNMHDR;

		TDC_COLUMN nColID = (TDC_COLUMN)pNMIA->iSubItem;
		UINT nCtrlID = MapColumnToCtrlID(nColID);
		
		if (nCtrlID)
		{
			// make sure the edit controls are visible
			if (m_nMaxState != TDCMS_NORMAL)
				SetMaximizeState(TDCMS_NORMAL);
			
			GetDlgItem(nCtrlID)->SetFocus();
		}
		
		return;
	}

	*pResult = 0;
}

TDC_HITTEST CTabbedToDoCtrl::HitTest(const CPoint& ptScreen) const
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::HitTest(ptScreen);

	case FTCV_TASKLIST:
		return m_taskList.HitTest(ptScreen);

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			IUIExtensionWindow* pExt = GetExtensionWnd(nView);
			ASSERT(pExt);

			if (pExt)
			{
				IUI_HITTEST nHit = pExt->HitTest(ptScreen);

				switch (nHit)
				{
				case IUI_TASKLIST:		return TDCHT_TASKLIST;
				case IUI_COLUMNHEADER:	return TDCHT_COLUMNHEADER;
				case IUI_TASK:			return TDCHT_TASK;

				case IUI_NOWHERE:
				default: // fall thru
					break;
				}
			}
		}
		break;

	default:
		ASSERT(0);
	}

	// else
	return TDCHT_NOWHERE;
}

TDC_COLUMN CTabbedToDoCtrl::ColumnHitTest(const CPoint& ptScreen) const
{
	return CToDoCtrl::ColumnHitTest(ptScreen);
}

void CTabbedToDoCtrl::Resort(BOOL bAllowToggle)
{
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::Resort(bAllowToggle);
		break;
		
	case FTCV_TASKLIST:
		m_taskList.Resort(bAllowToggle);
		break;
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			VIEWDATA* pData = GetViewData(nView);
			ASSERT(pData);

			if (pData)
				Sort(pData->sort.single.nBy, bAllowToggle);
		}
		break;
		
	default:
		ASSERT(0);
	}
}

BOOL CTabbedToDoCtrl::IsSortingBy(TDC_COLUMN nBy) const
{
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::IsSortingBy(nBy);
		
	case FTCV_TASKLIST:
		return m_taskList.IsSortingBy(nBy);
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		return FALSE;
		
	default:
		ASSERT(0);
	}

	return FALSE;
}

BOOL CTabbedToDoCtrl::IsSorting() const
{
	return GetSort().IsSorting();
}

BOOL CTabbedToDoCtrl::IsMultiSorting() const
{
	return GetSort().bMulti;
}

void CTabbedToDoCtrl::MultiSort(const TDSORTCOLUMNS& sort)
{
	ASSERT (sort.IsSorting());

	if (!sort.IsSorting())
		return;

	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::MultiSort(sort);
		break;

	case FTCV_TASKLIST:
		m_taskList.MultiSort(sort);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}
}

BOOL CTabbedToDoCtrl::CanMultiSort() const
{
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
	case FTCV_TASKLIST:
		return TRUE;
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		return FALSE;
	}
	
	// all else
	ASSERT(0);
	return FALSE;
}

void CTabbedToDoCtrl::Sort(TDC_COLUMN nBy, BOOL bAllowToggle)
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::Sort(nBy, bAllowToggle);
		break;

	case FTCV_TASKLIST:
		m_taskList.Sort(nBy, bAllowToggle);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			ExtensionDoAppCommand(nView, (bAllowToggle ? IUI_TOGGLABLESORT : IUI_SORT), nBy);

			VIEWDATA* pData = GetViewData(nView);
			ASSERT(pData);
			
			if (pData)
				pData->sort.single.nBy = nBy;
		}
		break;

	default:
		ASSERT(0);
	}
}

BOOL CTabbedToDoCtrl::CanSortBy(TDC_COLUMN nBy) const
{
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::CanSortBy(nBy);
		
	case FTCV_TASKLIST:
		return m_taskList.CanSortBy(nBy);
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		return ((nBy == TDCC_NONE) || ExtensionCanDoAppCommand(nView, IUI_SORT, nBy));
	}
	
	// else
	ASSERT(0);
	return FALSE;
}

BOOL CTabbedToDoCtrl::MoveSelectedTask(TDC_MOVETASK nDirection) 
{ 
	return !InTreeView() ? FALSE : CToDoCtrl::MoveSelectedTask(nDirection); 
}

BOOL CTabbedToDoCtrl::CanMoveSelectedTask(TDC_MOVETASK nDirection) const 
{ 
	return !InTreeView() ? FALSE : CToDoCtrl::CanMoveSelectedTask(nDirection); 
}

BOOL CTabbedToDoCtrl::GotoNextTask(TDC_GOTO nDirection)
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::GotoNextTask(nDirection);

	case FTCV_TASKLIST:
		if (CanGotoNextTask(nDirection))
		{
			int nSel = m_taskList.GetFirstSelectedItem();

			if (nDirection == TDCG_NEXT)
				nSel++;
			else
				nSel--;

			return SelectTask(m_taskList.GetTaskID(nSel));
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}
	
	// else
	return FALSE;
}

BOOL CTabbedToDoCtrl::CanGotoNextTask(TDC_GOTO nDirection) const
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::CanGotoNextTask(nDirection);

	case FTCV_TASKLIST:
		{
			int nSel = m_taskList.GetFirstSelectedItem();

			if (nDirection == TDCG_NEXT)
				return (nSel >= 0 && nSel < m_taskList.GetItemCount() - 1);
			
			// else prev
			return (nSel > 0 && nSel <= m_taskList.GetItemCount() - 1);
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}
	
	// else
	return FALSE;
}

BOOL CTabbedToDoCtrl::GotoNextTopLevelTask(TDC_GOTO nDirection)
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::GotoNextTopLevelTask(nDirection);

	case FTCV_TASKLIST:
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}

	// else
	return FALSE; // not supported
}

BOOL CTabbedToDoCtrl::CanGotoNextTopLevelTask(TDC_GOTO nDirection) const
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::CanGotoNextTopLevelTask(nDirection);

	case FTCV_TASKLIST:
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}

	// else
	return FALSE; // not supported
}

void CTabbedToDoCtrl::ExpandTasks(TDC_EXPANDCOLLAPSE nWhat, BOOL bExpand)
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::ExpandTasks(nWhat, bExpand);

	case FTCV_TASKLIST:
		// no can do!
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		if (bExpand)
		{
			switch (nWhat)
			{
			case TDCEC_ALL:
				ExtensionDoAppCommand(nView, IUI_EXPANDALL);
				break;

			case TDCEC_SELECTED:
				ExtensionDoAppCommand(nView, IUI_EXPANDSELECTED);
				break;
			}
		}
		else // collapse
		{
			switch (nWhat)
			{
			case TDCEC_ALL:
				ExtensionDoAppCommand(nView, IUI_COLLAPSEALL);
				break;

			case TDCEC_SELECTED:
				ExtensionDoAppCommand(nView, IUI_COLLAPSESELECTED);
				break;
			}
		}
		break;

	default:
		ASSERT(0);
	}
}

BOOL CTabbedToDoCtrl::CanExpandTasks(TDC_EXPANDCOLLAPSE nWhat, BOOL bExpand) const 
{ 
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::CanExpandTasks(nWhat, bExpand);

	case FTCV_TASKLIST:
		// no can do!
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		if (bExpand)
		{
			switch (nWhat)
			{
			case TDCEC_ALL:
				return ExtensionCanDoAppCommand(nView, IUI_EXPANDALL);

			case TDCEC_SELECTED:
				return ExtensionCanDoAppCommand(nView, IUI_EXPANDSELECTED);
			}
		}
		else // collapse
		{
			switch (nWhat)
			{
			case TDCEC_ALL:
				return ExtensionCanDoAppCommand(nView, IUI_COLLAPSEALL);

			case TDCEC_SELECTED:
				return ExtensionCanDoAppCommand(nView, IUI_COLLAPSESELECTED);
			}
		}
		break;

	default:
		ASSERT(0);
	}

	// else
	return FALSE; // not supported
}

void CTabbedToDoCtrl::ExtensionDoAppCommand(FTC_VIEW nView, IUI_APPCOMMAND nCmd, DWORD dwExtra)
{
	IUIExtensionWindow* pExt = GetExtensionWnd(nView, FALSE);
	ASSERT(pExt);

	if (pExt)
		pExt->DoAppCommand(nCmd, dwExtra);
}

BOOL CTabbedToDoCtrl::ExtensionCanDoAppCommand(FTC_VIEW nView, IUI_APPCOMMAND nCmd, DWORD dwExtra) const
{
	const IUIExtensionWindow* pExt = GetExtensionWnd(nView);
	ASSERT(pExt);

	if (pExt)
		return pExt->CanDoAppCommand(nCmd, dwExtra);

	return FALSE;
}

void CTabbedToDoCtrl::SetFocusToTasks()
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::SetFocusToTasks();
		break;

	case FTCV_TASKLIST:
		if (GetFocus() != &m_taskList)
		{
			// See CToDoCtrl::SetFocusToTasks() for why we need this
			SetFocusToComments();
			
			m_taskList.SetFocus();
		}
			
		// ensure the selected list item is visible
		if (!m_taskList.EnsureSelectionVisible())
			m_taskList.SelectItem(0);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		ExtensionDoAppCommand(nView, IUI_SETFOCUS);
		break;

	default:
		ASSERT(0);
	}
}

BOOL CTabbedToDoCtrl::TasksHaveFocus() const
{ 
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::TasksHaveFocus(); 

	case FTCV_TASKLIST:
		return m_taskList.HasFocus();

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		return ExtensionCanDoAppCommand(nView, IUI_SETFOCUS);

	default:
		ASSERT(0);
	}
	
	return FALSE;
}

int CTabbedToDoCtrl::FindTasks(const SEARCHPARAMS& params, CResultArray& aResults) const
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::FindTasks(params, aResults);

	case FTCV_TASKLIST:
		{
			for (int nItem = 0; nItem < m_taskList.GetItemCount(); nItem++)
			{
				DWORD dwTaskID = GetTaskID(nItem);
				SEARCHRESULT result;

				if (m_data.TaskMatches(dwTaskID, params, result))
					aResults.Add(result);
			}
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}

	return aResults.GetSize();
}


BOOL CTabbedToDoCtrl::SelectTask(CString sPart, TDC_SELECTTASK nSelect)
{
	int nFind = -1;
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::SelectTask(sPart, nSelect);

	case FTCV_TASKLIST:
		switch (nSelect)
		{
		case TDC_SELECTFIRST:
			nFind = FindListTask(sPart);
			break;
			
		case TDC_SELECTNEXT:
			nFind = FindListTask(sPart, m_taskList.GetFirstSelectedItem() + 1);
			break;
			
		case TDC_SELECTNEXTINCLCURRENT:
			nFind = FindListTask(sPart, m_taskList.GetFirstSelectedItem());
			break;
			
		case TDC_SELECTPREV:
			nFind = FindListTask(sPart, m_taskList.GetFirstSelectedItem() - 1, FALSE);
			break;
			
		case TDC_SELECTLAST:
			nFind = FindListTask(sPart, m_taskList.List().GetItemCount() - 1, FALSE);
			break;
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}

	// else
	if (nFind != -1)
		return SelectTask(GetTaskID(nFind));

	return FALSE;
}

int CTabbedToDoCtrl::FindListTask(const CString& sPart, int nStart, BOOL bNext)
{
	// build a search query
	SEARCHPARAMS params;
	params.aRules.Add(SEARCHPARAM(TDCA_TASKNAMEORCOMMENTS, FOP_INCLUDES, sPart));

	// we need to do this manually because CListCtrl::FindItem 
	// only looks at the start of the string
	SEARCHRESULT result;

	int nFrom = nStart;
	int nTo = bNext ? m_taskList.List().GetItemCount() : -1;
	int nInc = bNext ? 1 : -1;

	for (int nItem = nFrom; nItem != nTo; nItem += nInc)
	{
		DWORD dwTaskID = GetTaskID(nItem);

		if (m_data.TaskMatches(dwTaskID, params, result))
			return nItem;
	}

	return -1; // no match
}

void CTabbedToDoCtrl::SelectNextTasksInHistory()
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::SelectNextTasksInHistory();
		break;

	case FTCV_TASKLIST:
		if (CanSelectNextTasksInHistory())
		{
			// let CToDoCtrl do it's thing
			CToDoCtrl::SelectNextTasksInHistory();

			// then update our own selection
			ResyncListSelection();
		}
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}
}

BOOL CTabbedToDoCtrl::SelectTasks(const CDWordArray& aTasks, BOOL bRedraw)
{
	BOOL bRes = CToDoCtrl::SelectTasks(aTasks, bRedraw);

	// extra processing
	if (bRes)
	{
		FTC_VIEW nView = GetView();

		switch (nView)
		{
		case FTCV_TASKTREE:
		case FTCV_UNSET:
			break;

		case FTCV_TASKLIST:
			ResyncListSelection();
			break;

		case FTCV_UIEXTENSION1:
		case FTCV_UIEXTENSION2:
		case FTCV_UIEXTENSION3:
		case FTCV_UIEXTENSION4:
		case FTCV_UIEXTENSION5:
		case FTCV_UIEXTENSION6:
		case FTCV_UIEXTENSION7:
		case FTCV_UIEXTENSION8:
		case FTCV_UIEXTENSION9:
		case FTCV_UIEXTENSION10:
		case FTCV_UIEXTENSION11:
		case FTCV_UIEXTENSION12:
		case FTCV_UIEXTENSION13:
		case FTCV_UIEXTENSION14:
		case FTCV_UIEXTENSION15:
		case FTCV_UIEXTENSION16:
			break;

		default:
			ASSERT(0);
		}
	}

	return bRes;
}

void CTabbedToDoCtrl::ResyncListSelection()
{
	// optimisation when all items selected
	if (TSH().GetCount() == m_taskTree.GetItemCount())
	{
		m_taskList.SelectAll();
	}
	else
	{
		// save current states
		TDCSELECTIONCACHE cacheList, cacheTree;

		CacheListSelection(cacheList);
		CacheTreeSelection(cacheTree);

		if (!cacheList.SelectionMatches(cacheTree))
		{
			// now update the list selection using the tree's selection
			// but the list's breadcrumbs (if it has any), and save list 
			// scroll pos before restoring
			cacheTree.dwFirstVisibleTaskID = GetTaskID(m_taskList.List().GetTopIndex());

			if (cacheList.aBreadcrumbs.GetSize())
				cacheTree.aBreadcrumbs.Copy(cacheList.aBreadcrumbs);

			RestoreListSelection(cacheTree);
			
			// now check that the tree is correctly synced with us!
			CacheListSelection(cacheList);

			if (!cacheList.SelectionMatches(cacheTree))
				RestoreTreeSelection(cacheList);
		}
	}

	m_taskList.UpdateSelectedTaskPath();
}

void CTabbedToDoCtrl::SelectPrevTasksInHistory()
{
	if (CanSelectPrevTasksInHistory())
	{
		// let CToDoCtrl do it's thing
		CToDoCtrl::SelectPrevTasksInHistory();

		// extra processing
		FTC_VIEW nView = GetView();

		switch (nView)
		{
		case FTCV_TASKTREE:
		case FTCV_UNSET:
			// handled above
			break;

		case FTCV_TASKLIST:
			// then update our own selection
			ResyncListSelection();
			break;

		case FTCV_UIEXTENSION1:
		case FTCV_UIEXTENSION2:
		case FTCV_UIEXTENSION3:
		case FTCV_UIEXTENSION4:
		case FTCV_UIEXTENSION5:
		case FTCV_UIEXTENSION6:
		case FTCV_UIEXTENSION7:
		case FTCV_UIEXTENSION8:
		case FTCV_UIEXTENSION9:
		case FTCV_UIEXTENSION10:
		case FTCV_UIEXTENSION11:
		case FTCV_UIEXTENSION12:
		case FTCV_UIEXTENSION13:
		case FTCV_UIEXTENSION14:
		case FTCV_UIEXTENSION15:
		case FTCV_UIEXTENSION16:
			break;

		default:
			ASSERT(0);
		}
	}
}

void CTabbedToDoCtrl::InvalidateItem(HTREEITEM hti, BOOL bUpdate)
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::InvalidateItem(hti, bUpdate);
		break;

	case FTCV_TASKLIST:
		m_taskList.InvalidateItem(GetListItem(hti), bUpdate);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}
}

void CTabbedToDoCtrl::UpdateTreeSelection()
{
	// update the tree selection as required
	TDCSELECTIONCACHE cacheTree, cacheList;

	m_taskTree.CacheSelection(cacheTree);
	m_taskList.CacheSelection(cacheList);
	
	if (!cacheTree.SelectionMatches(cacheList, TRUE))
	{
		if (cacheList.aSelTaskIDs.GetSize() == 0)
		{
			m_taskTree.DeselectAll();
		}
		else if (cacheList.aSelTaskIDs.GetSize() == m_taskTree.GetItemCount())
		{
			m_taskTree.SelectAll();
		}
		else
		{
			m_taskTree.RestoreSelection(cacheList);
		}

		UpdateControls();
	}
}

BOOL CTabbedToDoCtrl::IsItemSelected(int nItem) const
{
	HTREEITEM hti = GetTreeItem(nItem);
	return hti ? TSH().HasItem(hti) : FALSE;
}

HTREEITEM CTabbedToDoCtrl::GetTreeItem(int nItem) const
{
	if (nItem < 0 || nItem >= m_taskList.GetItemCount())
		return NULL;

	DWORD dwID = m_taskList.GetTaskID(nItem);
	return m_taskTree.GetItem(dwID);
}

int CTabbedToDoCtrl::GetListItem(HTREEITEM hti) const
{
	DWORD dwID = GetTaskID(hti);
	return (dwID ? m_taskList.FindTaskItem(dwID) : -1);
}

LRESULT CTabbedToDoCtrl::OnDropObject(WPARAM wParam, LPARAM lParam)
{
	if (IsReadOnly())
		return 0L;

	TLDT_DATA* pData = (TLDT_DATA*)wParam;
	CWnd* pTarget = (CWnd*)lParam;
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::OnDropObject(wParam, lParam); // default handling
		break;


	case FTCV_TASKLIST:
		// simply convert the list item into the corresponding tree
		// item and pass to base class
		if (pTarget == &m_taskList.List())
		{
			ASSERT (InListView());

 			if (pData->nItem != -1)
 				m_taskList.SelectItem(pData->nItem);

			pData->hti = GetTreeItem(pData->nItem);
			pData->nItem = -1;
			lParam = (LPARAM)&m_taskTree.Tree();
		}
		// fall thru for default handling

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		CToDoCtrl::OnDropObject(wParam, lParam); // default handling
		break;

	default:
		ASSERT(0);
	}

	// else
	return 0L;
}

BOOL CTabbedToDoCtrl::GetLabelEditRect(CRect& rScreen)
{
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		return CToDoCtrl::GetLabelEditRect(rScreen);

	case FTCV_TASKLIST:
		if (m_taskList.EnsureSelectionVisible())
			return m_taskList.GetLabelEditRect(rScreen);
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		{
			IUIExtensionWindow* pExt = GetExtensionWnd(nView);
			ASSERT(pExt);

			return (pExt && pExt->GetLabelEditRect(rScreen));
		}
		break;

	default:
		ASSERT(0);
	}

	return FALSE;
}

void CTabbedToDoCtrl::UpdateSelectedTaskPath()
{
	CToDoCtrl::UpdateSelectedTaskPath();

	// extra processing
	FTC_VIEW nView = GetView();

	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		// handled above
		break;

	case FTCV_TASKLIST:
		m_taskList.UpdateSelectedTaskPath();
		break;

	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		break;

	default:
		ASSERT(0);
	}
}

void CTabbedToDoCtrl::SaveTasksState(CPreferences& prefs) const
{
	m_taskList.SaveState(prefs, GetPreferencesKey(_T("TaskList")));
	
	// base class
	CToDoCtrl::SaveTasksState(prefs);
}

HTREEITEM CTabbedToDoCtrl::LoadTasksState(const CPreferences& prefs)
{
	m_taskList.LoadState(prefs, GetPreferencesKey(_T("TaskList")));
	m_taskList.RecalcColumnWidths();

	if (m_taskList.IsSorting())
	{
		VIEWDATA* pData = GetViewData(FTCV_TASKLIST);
		ASSERT(pData);

		pData->bNeedResort = TRUE;
	}

	// base class
	return CToDoCtrl::LoadTasksState(prefs);
}

BOOL CTabbedToDoCtrl::IsViewSet() const 
{ 
	return (GetView() != FTCV_UNSET); 
}

BOOL CTabbedToDoCtrl::InListView() const 
{ 
	return (GetView() == FTCV_TASKLIST); 
}

BOOL CTabbedToDoCtrl::InTreeView() const 
{ 
	return (GetView() == FTCV_TASKTREE || !IsViewSet()); 
}

BOOL CTabbedToDoCtrl::InExtensionView() const
{
	return IsExtensionView(GetView());
}

BOOL CTabbedToDoCtrl::IsExtensionView(FTC_VIEW nView)
{
	return (nView >= FTCV_UIEXTENSION1 && nView <= FTCV_UIEXTENSION16);
}

BOOL CTabbedToDoCtrl::HasAnyExtensionViews() const
{
	int nView = m_aExtViews.GetSize();

	while (nView--)
	{
		if (m_aExtViews[nView] != NULL)
			return TRUE;
	}

	// else
	return FALSE;
}

void CTabbedToDoCtrl::ShowListViewTab(BOOL bVisible)
{
	m_tabViews.ShowViewTab(FTCV_TASKLIST, bVisible);
}

BOOL CTabbedToDoCtrl::IsListViewTabShowing() const
{
	return m_tabViews.IsViewTabShowing(FTCV_TASKLIST);
}

void CTabbedToDoCtrl::SetVisibleExtensionViews(const CStringArray& aTypeIDs)
{
	// update extension visibility
	int nExt = m_mgrUIExt.GetNumUIExtensions();

	while (nExt--)
	{
		FTC_VIEW nView = (FTC_VIEW)(FTCV_UIEXTENSION1 + nExt);

		VIEWDATA* pData = GetViewData(nView);
		ASSERT(pData);

		// update tab control
		CString sTypeID = m_mgrUIExt.GetUIExtensionTypeID(nExt);
		BOOL bVisible = (Misc::Find(aTypeIDs, sTypeID, FALSE, FALSE) != -1);

		m_tabViews.ShowViewTab(nView, bVisible);
	}
}

int CTabbedToDoCtrl::GetVisibleExtensionViews(CStringArray& aTypeIDs) const
{
	ASSERT(GetSafeHwnd());

	aTypeIDs.RemoveAll();

	int nExt = m_mgrUIExt.GetNumUIExtensions();

	while (nExt--)
	{
		FTC_VIEW nView = (FTC_VIEW)(FTCV_UIEXTENSION1 + nExt);

		if (m_tabViews.IsViewTabShowing(nView))
			aTypeIDs.Add(m_mgrUIExt.GetUIExtensionTypeID(nExt));
	}

	return aTypeIDs.GetSize();
}

BOOL CTabbedToDoCtrl::SetStyle(TDC_STYLE nStyle, BOOL bOn) 
{ 
	if (CToDoCtrl::SetStyle(nStyle, bOn))
	{
		// extra handling for extensions
		switch (nStyle)
		{
		case TDCS_READONLY:
			{
				int nView = m_aExtViews.GetSize();
				
				while (nView--)
				{
					if (m_aExtViews[nView] != NULL)
						m_aExtViews[nView]->SetReadOnly(bOn != FALSE);
				}
			}
			break;
		}

		return TRUE;
	}

	return FALSE;
}

BOOL CTabbedToDoCtrl::CanResizeAttributeColumnsToFit() const
{
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
	case FTCV_TASKLIST:
		return GetVisibleColumns().GetSize();
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		return ExtensionCanDoAppCommand(nView, IUI_RESIZEATTRIBCOLUMNS);
		
	default:
		ASSERT(0);
	}

	return FALSE;
}

void CTabbedToDoCtrl::ResizeAttributeColumnsToFit()
{
	FTC_VIEW nView = GetView();
	
	switch (nView)
	{
	case FTCV_TASKTREE:
	case FTCV_UNSET:
		CToDoCtrl::ResizeAttributeColumnsToFit();
		break;
		
	case FTCV_TASKLIST:
		m_taskList.RecalcAllColumnWidths();
		break;
		
	case FTCV_UIEXTENSION1:
	case FTCV_UIEXTENSION2:
	case FTCV_UIEXTENSION3:
	case FTCV_UIEXTENSION4:
	case FTCV_UIEXTENSION5:
	case FTCV_UIEXTENSION6:
	case FTCV_UIEXTENSION7:
	case FTCV_UIEXTENSION8:
	case FTCV_UIEXTENSION9:
	case FTCV_UIEXTENSION10:
	case FTCV_UIEXTENSION11:
	case FTCV_UIEXTENSION12:
	case FTCV_UIEXTENSION13:
	case FTCV_UIEXTENSION14:
	case FTCV_UIEXTENSION15:
	case FTCV_UIEXTENSION16:
		ExtensionDoAppCommand(nView, IUI_RESIZEATTRIBCOLUMNS);
		break;
		
	default:
		ASSERT(0);
	}
}
