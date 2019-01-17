// ToolsHelper.cpp: implementation of the CToolsHelper class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "ToolsHelper.h"
#include "ToolsUserInputDlg.h"

#include "..\shared\holdredraw.h"
#include "..\shared\SysImageList.h"
#include "..\shared\enfiledialog.h"
#include "..\shared\enstring.h"
#include "..\shared\enbitmap.h"
#include "..\shared\filemisc.h"
#include "..\shared\preferences.h"

#include <shlwapi.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CToolsHelper::CToolsHelper(BOOL bTDLEnabled, UINT nStart, int nSize) : m_nStartID(nStart), m_nSize(nSize), m_bTDLEnabled(bTDLEnabled)
{
	
}

CToolsHelper::~CToolsHelper()
{
	
}

BOOL CToolsHelper::RunTool(const USERTOOL& tool, const USERTOOLARGS& args, CWnd* pWnd)
{
	CString sCmdline, sToolPath;
	
	if (PrepareCmdline(tool, args, sToolPath, sCmdline))
	{
		HWND hWnd = pWnd ? pWnd->GetSafeHwnd() : NULL;
		
		DWORD dwRes = FileMisc::Run(hWnd, sToolPath, sCmdline, (tool.bRunMinimized ? SW_MINIMIZE : SW_SHOWNORMAL));
		return (dwRes > 32);
	}
	
	return FALSE;
}

void CToolsHelper::TestTool(const USERTOOL& tool, const USERTOOLARGS& args, CWnd* pWnd)
{
	CString sCmdline, sToolPath;

	if (PrepareCmdline(tool, args, sToolPath, sCmdline))
	{
		CEnString sMessage;
		sMessage.Format(IDS_TH_TESTTOOL, sToolPath, sCmdline);
		
		HWND hWnd = pWnd ? pWnd->GetSafeHwnd() : NULL;
		
		if (::MessageBox(hWnd, sMessage, CEnString(IDS_TH_TESTTOOL_TITLE), MB_YESNO | MB_ICONQUESTION) == IDYES)
		{
			DWORD dwRes = FileMisc::Run(hWnd, sToolPath, sCmdline, (tool.bRunMinimized ? SW_MINIMIZE : SW_SHOWNORMAL));
			
			if (dwRes <= 32)
			{
				// TODO
			}
		}
	}
}

CString CToolsHelper::GetToolPath(const USERTOOL& tool)
{
	CString sToolPath = tool.sToolPath;
	
	if (!sToolPath.IsEmpty() && !::PathIsURL(sToolPath))
	{
		if (!CToolsCmdlineParser::PrepareToolPath(sToolPath))
			FileMisc::MakeFullPath(sToolPath, FileMisc::GetAppFolder());
	}

	return sToolPath;
}

BOOL CToolsHelper::GetToolPaths(const USERTOOL& tool, CString& sToolPath, CString& sIconPath)
{
	sToolPath = GetToolPath(tool);
	sIconPath = tool.sIconPath;

	if (!tool.sIconPath.IsEmpty())
		sIconPath = FileMisc::GetFullPath(tool.sIconPath, FileMisc::GetAppFolder());
	else
		sIconPath.IsEmpty();

	return !sToolPath.IsEmpty();
}

HICON CToolsHelper::GetToolIcon(CSysImageList& sil, const USERTOOL& tool)
{
	HICON hIcon = NULL;

	// Check valid tool path. Note: could also be url
	CString sToolPath, sIconPath;
	VERIFY(GetToolPaths(tool, sToolPath, sIconPath));

	if (FileMisc::FileExists(sToolPath) || FileMisc::FileExists(sIconPath) || ::PathIsURL(sToolPath))
	{
		if (tool.sIconPath.IsEmpty())
		{
			hIcon = sil.ExtractFileIcon(sToolPath);	
		}
		else // try for supported image
		{

			hIcon = CEnBitmap::LoadImageFileAsIcon(sIconPath, CLR_NONE, 16, 16);

			// All else/Fallback
			if (hIcon == NULL)
				hIcon = sil.ExtractFileIcon(sIconPath);
		}
	}

	return hIcon;
}

void CToolsHelper::UpdateMenu(CCmdUI* pCmdUI, const CUserToolArray& tools, CMenuIconMgr& iconMgr)
{
	if (pCmdUI->m_pMenu)
	{
		CUserToolArray aTools;
		aTools.Copy(tools);
		
		// delete existing tool entries and their icons first
		for (int nTool = 0; nTool < m_nSize; nTool++)
		{
			pCmdUI->m_pMenu->DeleteMenu(pCmdUI->m_nID + nTool, MF_BYCOMMAND);
			iconMgr.DeleteImage(m_nStartID + nTool);
		}
		
		// if we have any tools to add we do it here
		if (aTools.GetSize())
		{
			// add valid tools only by first removing invalid items
			int nTool = aTools.GetSize();
			
			while (nTool--)
			{
				if (aTools[nTool].sToolName.IsEmpty() || aTools[nTool].sToolPath.IsEmpty())
					aTools.RemoveAt(nTool);
			}
			
			if (aTools.GetSize())
			{
				CSysImageList sil; // for menu icons
				VERIFY(sil.Initialize());
		
				int nPos = 0;
				
				for (nTool = 0; nTool < aTools.GetSize(); nTool++)
				{
					const USERTOOL& tool = aTools[nTool];
					CString sMenuItem;
					
					if (nPos < 9)
						sMenuItem.Format(_T("&%d %s"), nPos + 1, tool.sToolName);
					else
						sMenuItem = tool.sToolName;
					
					pCmdUI->m_pMenu->InsertMenu(pCmdUI->m_nIndex++, MF_BYPOSITION | MF_STRING, 
						m_nStartID + nTool, sMenuItem);

					iconMgr.SetImage(m_nStartID + nTool, GetToolIcon(sil, tool));
					
					nPos++;
				}
				
				
				// update end menu count
				pCmdUI->m_nIndex--; // point to last menu added
				pCmdUI->m_nIndexMax = pCmdUI->m_pMenu->GetMenuItemCount();
				
				pCmdUI->m_bEnableChanged = TRUE;    // all the added items are enabled
			}
		}
		
		// if nothing to add just re-add placeholder
		if (!aTools.GetSize())
		{
			pCmdUI->m_pMenu->InsertMenu(pCmdUI->m_nIndex, MF_BYPOSITION | MF_STRING | MF_GRAYED, 
				m_nStartID, CEnString(IDS_USERDEFINEDTOOLS));
		}
	}
}

BOOL CToolsHelper::PrepareCmdline(const USERTOOL& tool, const USERTOOLARGS& args, CString& sToolPath, CString& sCmdline)
{
	// do necessary substitutions
	CToolsCmdlineParser tcp(tool.sCmdline);
	CString sTasklist(args.sTasklist);
	
	if (tcp.HasTasklistArgument())
	{
		// allow for no loaded file
		if (!FileMisc::FileExists(sTasklist))
		{
			CPreferences prefs;
			CFileOpenDialog dialog(IDS_SELECTTASKLIST_TITLE, 
									GetDefaultFileExt(), 
									NULL, 
									EOFN_DEFAULTOPEN, 
									GetFileFilter());
			
			if (dialog.DoModal(&prefs) != IDOK)
				return FALSE;
			
			sTasklist = dialog.GetPathName();
		}
		
		CString sDrive, sPath, sFName, sExt;
		
		FileMisc::SplitPath(sTasklist, &sDrive, &sPath, &sFName, &sExt);
		
		tcp.ReplaceArgument(CLAT_PATHNAME, sTasklist);
		tcp.ReplaceArgument(CLAT_FOLDER, sDrive + sPath);
		tcp.ReplaceArgument(CLAT_FILENAME, sFName + sExt);
		tcp.ReplaceArgument(CLAT_FILETITLE, sFName);
	}
	
	tcp.ReplaceArgument(CLAT_TODOLIST, FileMisc::GetAppFilePath());
	tcp.ReplaceArgument(CLAT_SELTASKID, args.sTaskIDs);
	tcp.ReplaceArgument(CLAT_SELTASKTITLE, args.sTaskTitle);
	tcp.ReplaceArgument(CLAT_SELTASKEXTID, args.sTaskExtID);
	tcp.ReplaceArgument(CLAT_SELTASKCOMMENTS, args.sTaskComments);
	tcp.ReplaceArgument(CLAT_SELTASKFILELINK, args.sTaskFileLink);
	tcp.ReplaceArgument(CLAT_SELTASKALLOCBY, args.sTaskAllocBy);
	tcp.ReplaceArgument(CLAT_SELTASKALLOCTO, args.sTaskAllocTo);

	CCLArgArray aCustomArgs;
	int nArg = tcp.GetCustomAttributeArguments(aCustomArgs);

	while (nArg--)
	{
		CString sAttribID(aCustomArgs[nArg].sName), sValue;
		args.mapTaskCustData.Lookup(sAttribID, sValue);

		tcp.ReplaceArgument(aCustomArgs[nArg].sPlaceHolder, sValue);
	}
	
	if (tcp.IsUserInputRequired())
	{
		CToolsUserInputDlg dialog(tcp);
		
		if (dialog.DoModal(tool.sToolName) != IDOK)
			return FALSE;
		
		// process the results
		CCLArgArray aArgs;
		int nArg = tcp.GetUserArguments(aArgs);
		
		while (nArg--)
		{
			CString sResult(dialog.GetResult(aArgs[nArg].sName));
			tcp.ReplaceArgument(aArgs[nArg].sName, sResult);
		}
	}
	
	sCmdline = tcp.GetCmdLine();
	sToolPath = GetToolPath(tool);

	return TRUE;
}

void CToolsHelper::RemoveToolsFromToolbar(CToolBar& toolbar, UINT nCmdAfter)
{
	int nRemoved = 0;
	
	TBBUTTON tbb;
	CImageList* pIL = toolbar.GetToolBarCtrl().GetImageList();
	
	for (UINT nToolID = m_nStartID; nToolID <= m_nStartID + m_nSize; nToolID++)
	{
		int nBtn = toolbar.CommandToIndex(nToolID);
		
		if (nBtn != -1)
		{
			VERIFY(toolbar.GetToolBarCtrl().GetButton(nBtn, &tbb));
			
			if (toolbar.GetToolBarCtrl().DeleteButton(nBtn))
			{
				// delete the image too
				pIL->Remove(tbb.iBitmap);
				
				nRemoved++;
			}
		}
	}
	
	// remove separator
	if (nRemoved)
	{
		int nSep = toolbar.CommandToIndex(nCmdAfter) + 1;
		toolbar.GetToolBarCtrl().DeleteButton(nSep);
	}
}

void CToolsHelper::AppendToolsToToolbar(const CUserToolArray& aTools, CToolBar& toolbar, UINT nCmdAfter)
{
	// remove tools first
	RemoveToolsFromToolbar(toolbar, nCmdAfter);
	
	// then re-add
	if (aTools.GetSize())
	{
		// figure out if we want the large or small images
		CSize sizeBtn(toolbar.GetToolBarCtrl().GetButtonSize());
		sizeBtn -= CSize(7, 7); // btn borders from BarTool.cpp

		CSysImageList sil((sizeBtn.cx > 16));
		VERIFY(sil.Initialize());
		
		// start adding after the pref button
		int nStartPos = toolbar.CommandToIndex(nCmdAfter) + 1;
		int nAdded = 0;
		
		for (int nTool = 0; nTool < aTools.GetSize(); nTool++)
		{
			const USERTOOL& tool = aTools[nTool];
			HICON hIcon = GetToolIcon(sil, tool);
				
			if (hIcon)
			{
				CImageList* pIL = toolbar.GetToolBarCtrl().GetImageList();
				int nImage = pIL->Add(hIcon);
				
				TBBUTTON tbb = { nImage, nTool + m_nStartID, 0, TBSTYLE_BUTTON, 0, 0, (UINT)-1 };
				
				if (toolbar.GetToolBarCtrl().InsertButton(nStartPos + nAdded, &tbb))
					nAdded++;
				else // remove image
					pIL->Remove(nImage);
				
				// cleanup
				::DestroyIcon(hIcon);
			}
		}
		
		// add a separator if any buttons added
		if (nAdded)
		{
			TBBUTTON tbb = { -1, 0, 0, TBSTYLE_SEP, 0, 0, (UINT)-1 };
			toolbar.GetToolBarCtrl().InsertButton(nStartPos, &tbb);
		}
	}
}

LPCTSTR CToolsHelper::GetFileFilter()
{
	static CEnString TDLFILEFILTER(IDS_TDLFILEOPENFILTER);
	static CEnString XMLFILEFILTER(IDS_XMLFILEOPENFILTER);

	return m_bTDLEnabled ? TDLFILEFILTER : XMLFILEFILTER;
}

LPCTSTR CToolsHelper::GetDefaultFileExt()
{
	static LPCTSTR TDLEXT = _T("tdl");
	static LPCTSTR XMLEXT = _T("xml");

	return m_bTDLEnabled ? TDLEXT : XMLEXT;
}

