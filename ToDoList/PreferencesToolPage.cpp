// PreferencesToolPage.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "PreferencesToolPage.h"
#include "tdcenum.h"

#include "..\shared\enstring.h"
#include "..\shared\misc.h"
#include "..\shared\filemisc.h"
#include "..\shared\enfiledialog.h"
#include "..\shared\enmenu.h"
#include "..\shared\dialoghelper.h"
#include "..\shared\themed.h"

#include "..\3rdparty\ini.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPreferencesToolPage property page

LPCTSTR REALQUOTE = _T("\"");
LPCTSTR SAFEQUOTE = _T("{QUOTES}");

IMPLEMENT_DYNCREATE(CPreferencesToolPage, CPreferencesPageBase)

CPreferencesToolPage::CPreferencesToolPage(int nMaxNumTools) 
	: 
	CPreferencesPageBase(CPreferencesToolPage::IDD),
	m_eToolPath(FES_COMBOSTYLEBTN | FES_ALLOWURL),
	m_eIconPath(FES_COMBOSTYLEBTN | FES_DISPLAYSIMAGES),
	m_nMaxNumTools(nMaxNumTools)
{
	//{{AFX_DATA_INIT(CPreferencesToolPage)
	m_sToolPath = _T("");
	m_sCommandLine = _T("");
	m_bRunMinimized = FALSE;
	m_sIconPath = _T("");
	//}}AFX_DATA_INIT

	m_eCmdLine.AddButton(1, _T(""), CEnString(IDS_PTP_PLACEHOLDERS));
	m_eCmdLine.SetDropMenuButton(1);
}

CPreferencesToolPage::~CPreferencesToolPage()
{
}

void CPreferencesToolPage::DoDataExchange(CDataExchange* pDX)
{
	CPreferencesPageBase::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPreferencesToolPage)
	DDX_Control(pDX, IDC_CMDLINE, m_eCmdLine);
	DDX_Control(pDX, IDC_TOOLPATH, m_eToolPath);
	DDX_Control(pDX, IDC_TOOLLIST, m_lcTools);
	DDX_Text(pDX, IDC_TOOLPATH, m_sToolPath);
	DDX_Text(pDX, IDC_CMDLINE, m_sCommandLine);
	DDX_Check(pDX, IDC_RUNMINIMIZED, m_bRunMinimized);
	DDX_Text(pDX, IDC_ICONPATH, m_sIconPath);
	DDX_Control(pDX, IDC_ICONPATH, m_eIconPath);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CPreferencesToolPage, CPreferencesPageBase)
	//{{AFX_MSG_MAP(CPreferencesToolPage)
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_NEWTOOL, OnNewtool)
	ON_BN_CLICKED(IDC_DELETETOOL, OnDeletetool)
	ON_NOTIFY(LVN_ENDLABELEDIT, IDC_TOOLLIST, OnEndlabeleditToollist)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_TOOLLIST, OnItemchangedToollist)
	ON_EN_CHANGE(IDC_TOOLPATH, OnChangeToolpath)
	ON_NOTIFY(LVN_KEYDOWN, IDC_TOOLLIST, OnKeydownToollist)
	ON_EN_CHANGE(IDC_CMDLINE, OnChangeCmdline)
	ON_COMMAND_RANGE(ID_TOOLARG_PATHNAME, ID_TOOLARG_SELTASKFILELINK, OnInsertPlaceholder)
	ON_BN_CLICKED(IDC_RUNMINIMIZED, OnRunminimized)
	ON_BN_CLICKED(IDC_TESTTOOL, OnTesttool)
	ON_EN_CHANGE(IDC_ICONPATH, OnChangeIconPath)
	ON_BN_CLICKED(IDC_IMPORT, OnImport)
	ON_REGISTERED_MESSAGE(WM_EE_BTNCLICK, OnEEBtnClick)
	ON_REGISTERED_MESSAGE(WM_FE_GETFILEICON, OnGetFileIcon)

END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPreferencesToolPage message handlers

BOOL CPreferencesToolPage::OnInitDialog() 
{
	CPreferencesPageBase::OnInitDialog();

	m_ilSys.Initialize();
	m_lcTools.SetImageList(m_ilSys.GetImageList(), LVSIL_SMALL);

	m_eToolPath.SetCurrentFolder(FileMisc::GetAppFolder());
	m_eIconPath.SetCurrentFolder(FileMisc::GetAppFolder());

	CRect rList;
	m_lcTools.GetClientRect(rList);

	m_lcTools.InsertColumn(0, CEnString(IDS_PTP_TOOLNAME), LVCFMT_LEFT, 150);
	m_lcTools.InsertColumn(1, CEnString(IDS_PTP_TOOLPATH), LVCFMT_LEFT, 250);
	m_lcTools.InsertColumn(2, CEnString(IDS_PTP_ARGUMENTS), LVCFMT_LEFT, rList.Width() - 400);
	m_lcTools.InsertColumn(3, CEnString(IDS_PTP_ICONPATH), LVCFMT_LEFT, 0);

	m_lcTools.SetExtendedStyle(m_lcTools.GetExtendedStyle() | LVS_EX_FULLROWSELECT);
	
	// add tools we loaded from the registry
	for (int nTool = 0; nTool < m_aTools.GetSize(); nTool++)
	{
		const USERTOOL& tool = m_aTools[nTool];

		int nIndex = m_lcTools.InsertItem(nTool, tool.sToolName, -1);

		m_lcTools.SetItemText(nIndex, 1, tool.sToolPath);
		m_lcTools.SetItemText(nIndex, 2, tool.sCmdline);
		m_lcTools.SetItemText(nIndex, 3, tool.sIconPath);
		m_lcTools.SetItemData(nIndex, tool.bRunMinimized);
	}
	RebuildListCtrlImages();

	m_lcTools.SetItemState(0, LVIS_SELECTED, LVIS_SELECTED);
	OnItemchangedToollist(NULL, NULL);

	CThemed::SetWindowTheme(&m_lcTools, _T("Explorer"));

	EnableControls();
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CPreferencesToolPage::OnNewtool() 
{
	int nIndex = m_lcTools.InsertItem(m_lcTools.GetItemCount(), CEnString(IDS_PTP_NEWTOOL), -1);
	m_lcTools.SetItemText(nIndex, 2, CmdIDToPlaceholder(ID_TOOLARG_PATHNAME));

	m_lcTools.SetItemState(nIndex, LVIS_SELECTED, LVIS_SELECTED);
	m_lcTools.SetFocus();
	m_lcTools.EditLabel(nIndex);

	CPreferencesPageBase::OnControlChange();
}

void CPreferencesToolPage::OnDeletetool() 
{
	int nSel = GetCurSel();

	if (nSel >= 0)
	{
		m_lcTools.DeleteItem(nSel);
		m_sToolPath.Empty();
		m_sIconPath.Empty();

		EnableControls();
		UpdateData(FALSE);
	}

	CPreferencesPageBase::OnControlChange();
}

void CPreferencesToolPage::OnEndlabeleditToollist(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LV_DISPINFO* pDispInfo = (LV_DISPINFO*)pNMHDR;
	
	if (pDispInfo->item.pszText)
	{
		int nSel = GetCurSel();

		if (nSel >= 0)
		{
			m_lcTools.SetItemText(nSel, 0, pDispInfo->item.pszText);

			GetDlgItem(IDC_TOOLPATH)->SetFocus();
		}
	}
	
	*pResult = 0;

	CPreferencesPageBase::OnControlChange();
}

void CPreferencesToolPage::OnItemchangedToollist(NMHDR* /*pNMHDR*/, LRESULT* /*pResult*/) 
{
	EnableControls();
	int nSel = GetCurSel();

	if (nSel >= 0)
	{
		m_sToolPath = m_lcTools.GetItemText(nSel, 1);
		m_sCommandLine = m_lcTools.GetItemText(nSel, 2);
		m_sIconPath = m_lcTools.GetItemText(nSel, 3);
		m_bRunMinimized = m_lcTools.GetItemData(nSel);
	}
	else
	{
		m_sToolPath.Empty();
		m_sCommandLine.Empty();
		m_bRunMinimized = FALSE;
		m_sIconPath.Empty();
	}

	UpdateData(FALSE);
}

void CPreferencesToolPage::EnableControls()
{
	int nSel = GetCurSel();

	GetDlgItem(IDC_NEWTOOL)->EnableWindow(m_lcTools.GetItemCount() < m_nMaxNumTools);
	GetDlgItem(IDC_DELETETOOL)->EnableWindow(nSel >= 0);
	GetDlgItem(IDC_TOOLPATH)->EnableWindow(nSel >= 0);
	GetDlgItem(IDC_ICONPATH)->EnableWindow(nSel >= 0);
	GetDlgItem(IDC_CMDLINE)->EnableWindow(nSel >= 0);
	GetDlgItem(IDC_RUNMINIMIZED)->EnableWindow(nSel >= 0);

	m_eCmdLine.EnableButton(1, nSel >= 0);
}

int CPreferencesToolPage::GetCurSel()
{
	int nSel = -1;
	POSITION pos = m_lcTools.GetFirstSelectedItemPosition();

	if (pos)
		nSel = m_lcTools.GetNextSelectedItem(pos);

	return nSel;
}

void CPreferencesToolPage::OnChangeToolpath() 
{
	int nSel = GetCurSel();

	if (nSel == CB_ERR)
	{
		ASSERT (0);
		return;
	}

	UpdateData();

	m_lcTools.SetItemText(nSel, 1, m_sToolPath);

	RebuildListCtrlImages();

	CPreferencesPageBase::OnControlChange();
}

void CPreferencesToolPage::RebuildListCtrlImages()
{
	// remove any existing images
	m_ilSys.GetImageList()->DeleteImageList();
	m_ilSys.Initialize();
	m_lcTools.SetImageList(m_ilSys.GetImageList(), LVSIL_SMALL);

	// Re-add
	int nTool = m_lcTools.GetItemCount();

	while (nTool--)
	{
		LVITEM lvi = { 0 };
		
		lvi.mask = LVIF_IMAGE;
		lvi.iItem = nTool;
		
		CString sIconPath = m_lcTools.GetItemText(nTool, 3);

		if (sIconPath.IsEmpty())
		{
			CString sToolPath = m_lcTools.GetItemText(nTool, 1);

			if (!CToolsCmdlineParser::PrepareToolPath(sToolPath))
				FileMisc::MakeFullPath(sToolPath, FileMisc::GetAppFolder());

			lvi.iImage = m_ilSys.GetFileImageIndex(sToolPath);	
		}
		else
		{
			FileMisc::MakeFullPath(sIconPath, FileMisc::GetAppFolder());

			HICON hIcon = CEnBitmap::LoadImageFileAsIcon(sIconPath, CLR_NONE, 16, 16);

			if (hIcon)
			{
				lvi.iImage = m_ilSys.GetImageList()->Add(hIcon);
				::DestroyIcon(hIcon);
			}
			else
			{
				lvi.iImage = m_ilSys.GetFileImageIndex(sIconPath);
			}
		}
		
		m_lcTools.SetItem(&lvi);
	}
}

void CPreferencesToolPage::OnChangeIconPath() 
{
	int nSel = GetCurSel();

	if (nSel == CB_ERR)
	{
		ASSERT (0);
		return;
	}

	UpdateData();

	m_lcTools.SetItemText(nSel, 3, m_sIconPath);

	RebuildListCtrlImages();

	CPreferencesPageBase::OnControlChange();
}

int CPreferencesToolPage::GetUserTools(CUserToolArray& aTools) const
{
	aTools.Copy(m_aTools);

	return aTools.GetSize();
}

BOOL CPreferencesToolPage::GetUserTool(int nTool, USERTOOL& tool) const
{
	if (nTool >= 0 && nTool < m_aTools.GetSize())
	{
		tool = m_aTools[nTool];
		return TRUE;
	}

	return FALSE;
}

void CPreferencesToolPage::OnOK() 
{
	CPreferencesPageBase::OnOK();
	
	// save tools to m_aTools
	m_aTools.RemoveAll();
	int nToolCount = m_lcTools.GetItemCount();

	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		USERTOOL ut;

		ut.sToolName = m_lcTools.GetItemText(nTool, 0);
		ut.sToolPath = m_lcTools.GetItemText(nTool, 1);
		ut.sCmdline = m_lcTools.GetItemText(nTool, 2);
		ut.sIconPath = m_lcTools.GetItemText(nTool, 3);
		ut.bRunMinimized = m_lcTools.GetItemData(nTool);

		// GetPrivateProfileString strips a leading/trailing quote pairs if 
		// it finds them so we replace quotes with safe quotes
		ut.sCmdline.Replace(REALQUOTE, SAFEQUOTE);
		
		m_aTools.Add(ut);
	}
}

void CPreferencesToolPage::OnKeydownToollist(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LV_KEYDOWN* pLVKeyDown = (LV_KEYDOWN*)pNMHDR;
	
	switch (pLVKeyDown->wVKey)
	{
	case VK_DELETE:
		OnDeletetool();
		break;

	case VK_F2:
		{
			int nSel = GetCurSel();

			if (nSel >= 0)
				m_lcTools.EditLabel(nSel);
		}
		break;
	}
	
	*pResult = 0;
}

void CPreferencesToolPage::OnChangeCmdline() 
{
	int nSel = GetCurSel();

	if (nSel >= 0)
	{
		UpdateData();

		m_lcTools.SetItemText(nSel, 2, m_sCommandLine);

		CPreferencesPageBase::OnControlChange();
	}
}

void CPreferencesToolPage::OnInsertPlaceholder(UINT nCmdID) 
{
	m_eCmdLine.ReplaceSel(CmdIDToPlaceholder(nCmdID), TRUE);

	CPreferencesPageBase::OnControlChange();
}

void CPreferencesToolPage::OnRunminimized() 
{
	int nSel = GetCurSel();

	if (nSel >= 0)
	{
		UpdateData();

		m_lcTools.SetItemData(nSel, m_bRunMinimized);

		CPreferencesPageBase::OnControlChange();
	}
}

LRESULT CPreferencesToolPage::OnEEBtnClick(WPARAM wParam, LPARAM lParam)
{
	switch (wParam)
	{
	case IDC_CMDLINE:
		if (lParam == 1) // button ID
		{
			CEnMenu menu;
			
			if (menu.LoadMenu(IDR_MISC, NULL, TRUE))
				m_eCmdLine.TrackPopupMenu(1U, menu.GetSubMenu(MM_TOOLARGS), EETPM_RIGHT);
		}
		break;
	}

	return 0L;
}

LRESULT CPreferencesToolPage::OnGetFileIcon(WPARAM wParam, LPARAM lParam)
{
	switch (wParam)
	{
	case IDC_TOOLPATH:
		{
			CString sToolPath((LPCTSTR)lParam);
			
			if (CToolsCmdlineParser::PrepareToolPath(sToolPath))
			{
				static HICON hIcon = m_ilSys.ExtractFileIcon(sToolPath);
				return (LRESULT)hIcon;
			}
		}
		break;
	}

	return 0L;
}

void CPreferencesToolPage::OnTesttool() 
{
	int nTool = GetCurSel();
	
	if (nTool != -1)
	{
		USERTOOL ut;
		
		ut.sToolName = m_lcTools.GetItemText(nTool, 0);
		ut.sToolPath = m_lcTools.GetItemText(nTool, 1);
		ut.sCmdline = m_lcTools.GetItemText(nTool, 2);
		ut.bRunMinimized = m_lcTools.GetItemData(nTool);
		
		GetParent()->SendMessage(WM_PTP_TESTTOOL, 0, (LPARAM)&ut);
	}
}

CString CPreferencesToolPage::CmdIDToPlaceholder(UINT nCmdID) 
{
	switch (nCmdID)
	{
	case ID_TOOLARG_PATHNAME:		return _T("\"$(pathName)\"");       
	case ID_TOOLARG_FILETITLE:		return _T("$(fileTitle)");      
	case ID_TOOLARG_FOLDER:			return _T("$(folder)");         
	case ID_TOOLARG_FILENAME:		return _T("$(fileName)");       
	case ID_TOOLARG_SELTASKID:		return _T("$(selTID)");       
	case ID_TOOLARG_SELTASKTITLE:	return _T("\"$(selTTitle)\"");       
	case ID_TOOLARG_USERDATE:		return _T("$(userDate, var_date1, \"Date Prompt\", default_date)");       
	case ID_TOOLARG_USERFILEPATH:	return _T("\"$(userFile, var_file1, \"File Prompt\", default_path)\"");   
	case ID_TOOLARG_USERFOLDER:		return _T("$(userFolder, var_folder1, \"Folder Prompt\", default_folder)");     
	case ID_TOOLARG_USERTEXT:		return _T("$(userText, var_text1, \"Text Prompt\", default_text)");       
	case ID_TOOLARG_TODAYSDATE:		return _T("$(todaysDate)");     
	case ID_TOOLARG_TODOLIST:		return _T("\"$(todolist)\""); 
	case ID_TOOLARG_SELTASKEXTID:	return _T("\"$(selTExtID)\""); 
	case ID_TOOLARG_SELTASKCOMMENTS:return _T("\"$(selTComments)\""); 
	case ID_TOOLARG_SELTASKFILELINK:return _T("\"$(selTFile)\""); 
	case ID_TOOLARG_SELTASKALLOCBY:	return _T("\"$(selTAllocBy)\""); 
	case ID_TOOLARG_SELTASKALLOCTO:	return _T("\"$(selTAllocTo)\""); 
	}

	return _T("");
}

void CPreferencesToolPage::OnImport() 
{
	BOOL bContinue = TRUE;
	CPreferences prefs;

	while (bContinue)
	{
		CFileOpenDialog dialog(IDS_IMPORTPREFS_TITLE, 
								_T("ini"), 
								NULL, 
								OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY | OFN_FILEMUSTEXIST, 
								CEnString(IDS_INIFILEFILTER));
		
		if (dialog.DoModal(&prefs) == IDOK)
		{
			CIni ini(dialog.GetPathName());

			int nTools = ini.GetInt(_T("Tools"), _T("ToolCount"), 0);

			if (!nTools)
				bContinue = (AfxMessageBox(IDS_INIHASNOTOOLS, MB_YESNO) == IDYES);
			else
			{
				int nCurCount = m_lcTools.GetItemCount();

				for (int nTool = 0; nTool < nTools; nTool++)
				{
					CString sKey = Misc::MakeKey(_T("Tools\\Tool%d"), nTool + 1);
					
					CString sName = ini.GetString(sKey, _T("Name"));
					CString sPath = ini.GetString(sKey, _T("Path"));
					CString sIconPath = ini.GetString(sKey, _T("IconPath"));
					BOOL bRunMinimized = ini.GetBool(sKey, _T("RunMinimized"), FALSE);
					CString sCmdline = ini.GetString(sKey, _T("Cmdline"));

					// replace safe quotes with real quotes
					sCmdline.Replace(SAFEQUOTE, REALQUOTE);

					// add tool to list
 					int nImage = -1;
					
					nImage = m_ilSys.GetFileImageIndex(sPath);
					
					if (!sIconPath.IsEmpty()) 
						nImage = m_ilSys.GetFileImageIndex(sIconPath);
						
					int nIndex = m_lcTools.InsertItem(nCurCount + nTool, sName, nImage);

					m_lcTools.SetItemText(nIndex, 1, sPath);
					m_lcTools.SetItemText(nIndex, 2, sCmdline);
					m_lcTools.SetItemText(nIndex, 3, sIconPath);
					m_lcTools.SetItemData(nIndex, bRunMinimized);
				}

				bContinue = FALSE;

				CPreferencesPageBase::OnControlChange();
			}
		}
		else
			bContinue = FALSE; // cancelled
	}
}

void CPreferencesToolPage::LoadPreferences(const CPreferences& prefs)
{
	// load tools
	int nToolCount = prefs.GetProfileInt(_T("Tools"), _T("ToolCount"), 0);

	for (int nTool = 1; nTool <= nToolCount; nTool++)
	{
		CString sKey = Misc::MakeKey(_T("Tools\\Tool%d"), nTool);

		USERTOOL ut;
		ut.sToolName = prefs.GetProfileString(sKey, _T("Name"), _T(""));
		ut.sToolPath = prefs.GetProfileString(sKey, _T("Path"), _T(""));
		ut.sCmdline = prefs.GetProfileString(sKey, _T("CmdLine"), _T("")); 
		ut.bRunMinimized = prefs.GetProfileInt(sKey, _T("RunMinimized"), FALSE);
		ut.sIconPath = prefs.GetProfileString(sKey, _T("IconPath"), _T(""));
		
		// replace safe quotes with real quotes
		ut.sCmdline.Replace(SAFEQUOTE, REALQUOTE);

		m_aTools.Add(ut);
	}
}

void CPreferencesToolPage::SavePreferences(CPreferences& prefs)
{
	// save tools to registry and m_aTools
	int nToolCount = m_aTools.GetSize();

	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		USERTOOL ut = m_aTools[nTool];

        CString sKey = Misc::MakeKey(_T("Tools\\Tool%d"), nTool + 1);
		
		prefs.WriteProfileString(sKey, _T("Name"), ut.sToolName);
		prefs.WriteProfileString(sKey, _T("Path"), ut.sToolPath);
		prefs.WriteProfileString(sKey, _T("IconPath"), ut.sIconPath);
		prefs.WriteProfileInt(sKey, _T("RunMinimized"), ut.bRunMinimized);
		
		// GetPrivateProfileString strips a leading/trailing quote pairs if 
		// it finds them so we replace quotes with safe quotes
		ut.sCmdline.Replace(REALQUOTE, SAFEQUOTE);
		prefs.WriteProfileString(sKey, _T("Cmdline"), ut.sCmdline);
	}

	prefs.WriteProfileInt(_T("Tools"), _T("ToolCount"), nToolCount);
}

void CPreferencesToolPage::OnSize(UINT nType, int cx, int cy) 
{
	CPreferencesPageBase::OnSize(nType, cx, cy);
	
	if (m_lcTools.GetSafeHwnd())
	{
		// calculate border
		CPoint ptBorders = CDialogHelper::GetCtrlRect(this, IDC_TOOLLIST).TopLeft();
		
		// calc offsets
		CPoint ptImport = CDialogHelper::GetCtrlRect(this, IDC_IMPORT).BottomRight();
		int nXOffset = cx - (ptImport.x + ptBorders.x);
		int nYOffset = cy - (ptImport.y + ptBorders.y);

		// move controls
		CUIntArray aCtrls;

		CDialogHelper::GetChildCtrlIDs(this, aCtrls);

		CDialogHelper::RemoveCtrlID(IDC_TOOLLIST, aCtrls);
		CDialogHelper::RemoveCtrlID(IDC_DELETETOOL, aCtrls);
		CDialogHelper::RemoveCtrlID(IDC_IMPORT, aCtrls);

		CDialogHelper::OffsetCtrls(this, aCtrls, 0, nYOffset);
		CDialogHelper::OffsetCtrl(this, IDC_DELETETOOL, nXOffset, nYOffset);
		CDialogHelper::OffsetCtrl(this, IDC_IMPORT, nXOffset, nYOffset);

		CDialogHelper::ResizeChild(&m_eToolPath, nXOffset, 0);
		CDialogHelper::ResizeChild(&m_eIconPath, nXOffset, 0);
		CDialogHelper::ResizeChild(&m_eCmdLine, nXOffset, 0);
		CDialogHelper::ResizeChild(&m_lcTools, nXOffset, nYOffset);
	}
}
