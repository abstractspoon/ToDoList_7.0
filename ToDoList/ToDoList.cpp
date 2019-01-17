// ToDoList.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "resource.h"
#include "todolist.h"
#include "ToDoListWnd.h"
#include "Preferencesdlg.h"
#include "welcomewizard.h"
#include "tdcenum.h"
#include "tdcmsg.h"
#include "tdlprefmigrationdlg.h"
#include "tdllanguagedlg.h"
#include "TDLCmdlineOptionsDlg.h"
#include "tdlolemessagefilter.h"
#include "tdlwebupdatepromptdlg.h"
#include "tdcstartupoptions.h"

#include "..\shared\encommandlineinfo.h"
#include "..\shared\driveinfo.h"
#include "..\shared\dialoghelper.h"
#include "..\shared\enfiledialog.h"
#include "..\shared\regkey.h"
#include "..\shared\enstring.h"
#include "..\shared\filemisc.h"
#include "..\shared\autoflag.h"
#include "..\shared\preferences.h"
#include "..\shared\localizer.h"
#include "..\shared\fileregister.h"
#include "..\shared\osversion.h"
#include "..\shared\rtlstylemgr.h"

#include "..\3rdparty\xmlnodewrapper.h"
#include "..\3rdparty\ini.h"
#include "..\3rdparty\base64coder.h"

#include <afxpriv.h>

#include <shlwapi.h>
#pragma comment(lib, "shlwapi.lib")

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

LPCTSTR REGKEY				= _T("AbstractSpoon");
LPCTSTR APPREGKEY			= _T("Software\\AbstractSpoon\\ToDoList");
LPCTSTR UNINSTALLREGKEY		= _T("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\AbstractSpoon_ToDoList");
LPCTSTR APPDATA				= _T("Abstractspoon\\ToDoList");

LPCTSTR WIKI_URL			= _T("http://bit.ly/AbstrToDoListWiki"); 
LPCTSTR GOOGLEGROUP_URL		= _T("http://bit.ly/AbstrToDoListGG"); 
LPCTSTR LINKEDIN_URL		= _T("http://bit.ly/AbstrToDoListLinkedIn"); 
LPCTSTR LICENSE_URL			= _T("https://github.com/abstractspoon/ToDoList_Wiki/wiki/Free-Open-Source-Software"); 
LPCTSTR DONATE_URL			= _T("https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=abstractspoon2%40optusnet%2ecom%2eau&item_name=Software"); 
LPCTSTR FACEBOOK_URL		= _T("http://bit.ly/AbstrToDoListFB"); 
LPCTSTR TWITTER_URL			= _T("http://bit.ly/AbstrToDoListTwitter"); 
LPCTSTR GOOGLEPLUS_URL		= _T("http://bit.ly/AbstrToDoListGPlus"); 

LPCTSTR FILESTATEKEY		= _T("FileStates");
LPCTSTR REMINDERKEY			= _T("Reminders");
LPCTSTR DEFAULTKEY			= _T("Default");

/////////////////////////////////////////////////////////////////////////////
// CToDoListApp

BEGIN_MESSAGE_MAP(CToDoListApp, CWinApp)
	//{{AFX_MSG_MAP(CToDoListApp)
	//}}AFX_MSG_MAP
	ON_COMMAND(ID_HELP_GOOGLEGROUP, OnHelpGoogleGroup)
	ON_COMMAND(ID_HELP_LINKEDIN, OnHelpLinkedIn)
	ON_COMMAND(ID_HELP_LICENSE, OnHelpLicense)
	ON_COMMAND(ID_HELP_COMMANDLINE, OnHelpCommandline)
	ON_COMMAND(ID_HELP_DONATE, OnHelpDonate)
	ON_COMMAND(ID_HELP_UNINSTALL, OnHelpUninstall)
	ON_COMMAND(ID_HELP_RECORDBUGREPORT, OnHelpRecordBugReport)
	ON_COMMAND(ID_HELP_WIKI, OnHelpWiki)
	ON_COMMAND(ID_HELP_FACEBOOK, OnHelpFacebook)
	ON_COMMAND(ID_HELP_TWITTER, OnHelpTwitter)
	ON_COMMAND(ID_HELP_GOOGLEPLUS, OnHelpGooglePlus)

#ifdef _DEBUG
	ON_COMMAND(ID_DEBUGTASKDIALOG_INFO, OnDebugTaskDialogInfo)
	ON_COMMAND(ID_DEBUGSHOWUPDATEDLG, OnDebugShowUpdateDlg)
	ON_COMMAND(ID_DEBUGSHOWSCRIPTDLG, OnDebugShowScriptDlg)
	ON_COMMAND(ID_DEBUGTASKDIALOG_WARNING, OnDebugTaskDialogWarning)
	ON_COMMAND(ID_DEBUGTASKDIALOG_ERROR, OnDebugTaskDialogError)
#endif

	ON_COMMAND(ID_TOOLS_CHECKFORUPDATES, OnHelpCheckForUpdates)
	ON_COMMAND(ID_TOOLS_IMPORTPREFS, OnImportPrefs)
	ON_COMMAND(ID_TOOLS_EXPORTPREFS, OnExportPrefs)
	ON_UPDATE_COMMAND_UI(ID_TOOLS_IMPORTPREFS, OnUpdateImportPrefs)
	ON_UPDATE_COMMAND_UI(ID_TOOLS_EXPORTPREFS, OnUpdateExportPrefs)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CToDoListApp construction

CToDoListApp::CToDoListApp() : CWinApp(), m_bUseStaging(FALSE)
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CToDoListApp object

CToDoListApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CToDoListApp initialization

BOOL CToDoListApp::InitInstance()
{
	CEnCommandLineInfo cmdInfo(_T(".tdl;.xml"));
	ParseCommandLine(cmdInfo);

	// see if the user wants to uninstall
	if (cmdInfo.HasOption(SWITCH_UNINSTALL))
	{
		// we don't have the uninstaller run directly by the OS
		// to prevent the APPID appearing in plain text
		RunUninstaller();
		return FALSE;// quit app
	}

	// see if the user just wants to see the commandline options
	if (cmdInfo.HasOption(SWITCH_HELP1) || 
		cmdInfo.HasOption(SWITCH_HELP2) || 
		cmdInfo.HasOption(SWITCH_HELP3))
	{
		OnHelpCommandline();
		return FALSE; // quit app
	}

	// If this is a restart, wait until the previous instance has closed
	if (cmdInfo.HasOption(SWITCH_RESTART))
	{
		DWORD dwProcID = _ttoi(cmdInfo.GetOption(SWITCH_RESTART));
		ASSERT(dwProcID);

		if (!WaitForInstanceToClose(dwProcID))
		{
			// message box the user
			// TODO

			return FALSE; // quit app
		}

		// else strip off the restart switch and continue
		cmdInfo.DeleteOption(SWITCH_RESTART);
	}

	AfxOleInit(); // for initializing COM and handling drag and drop via explorer
	AfxEnableControlContainer(); // embedding IE

	// before anything else make sure we've got MSXML3 installed
	if (!CXmlDocumentWrapper::IsVersion3orGreater())
	{
		AfxMessageBox(IDS_BADMSXML);
		return FALSE; // quit app
	}

	// init prefs 
	if (!InitPreferences(cmdInfo))
		return FALSE; // quit app

	// commandline options
	CTDCStartupOptions startup(cmdInfo);

	// see if another instance can better handle this than us
	HWND hwndOtherInst = NULL;

	TDCFINDWND find;
	int nNumWnds = FindToDoListWnds(find);

	if (nNumWnds)
	{
		// if there IS a tasklist on the commandline and
		// we are NOT importing it, see if any other instance 
		// already has it loaded 
		if (startup.HasFilePath() && !startup.HasFlag(TLD_IMPORTFILE))
		{
			for (int nWnd = 0; nWnd < nNumWnds; nWnd++)
			{
				HWND hWnd = find.aResults[nWnd];

				if (SendStartupParams(hWnd, startup, TDLCD_FINDFILE))
				{
					hwndOtherInst = hWnd;
					break;
				}
			}

			if (hwndOtherInst) // someone else already has the tasklist open
			{
				if (!SendStartupParams(hwndOtherInst, startup, TDLCD_STARTUP))
				{
					hwndOtherInst = NULL;
				}
			}
		}

		// if we didn't find it already open then we try again to see if 
		// anyone else wants to handle it
		BOOL bTasklistOpened = FALSE;

		if (hwndOtherInst == NULL)
		{
			// But ONLY If we are single-instance OR this is an inter-tasklist task link
			BOOL bSingleInstance = !CPreferencesDlg().GetMultiInstance();
			BOOL bTaskLink = startup.HasFlag(TLD_TASKLINK);

			if (bSingleInstance || bTaskLink)
			{
				for (int nWnd = 0; nWnd < nNumWnds; nWnd++)
				{
					HWND hWnd = find.aResults[nWnd];

					if (SendStartupParams(hWnd, startup, TDLCD_STARTUP))
					{
						hwndOtherInst = hWnd;
						bTasklistOpened = TRUE;

						break;
					}
				}
			}

			// if this still was not handled and we are single instance
			// then notify the user appropriately and quit
			if ((hwndOtherInst == NULL) && bSingleInstance)
			{
				if (startup.GetTaskID())
				{
					// TODO
				}

				return FALSE;
			}
		}

		// if we found a window, we may need to bring it to the foreground
		if (hwndOtherInst)
		{
			// we only show and bring to the foreground 
			// the other app window if:
			// 0. There are no commandline options
			// 1. This is a task link
			// 2. A tasklist was opened or imported
			// 3. A new task name requires editing
			// 4. A command requiring UI executed
			BOOL bShow = (startup.IsEmpty(TRUE) || 
				startup.HasFlag(TLD_TASKLINK) ||
				(startup.HasFilePath() && bTasklistOpened) ||
				(startup.HasFlag(TLD_NEWTASK) && !startup.HasNewTaskName()) ||
				(startup.HasCommandID() && CommandRequiresUI(startup.GetCommandID())));

			if (bShow)
			{
				::SendMessage(hwndOtherInst, WM_TDL_SHOWWINDOW, 0, 0);
				::SetForegroundWindow(hwndOtherInst);
			}

			return FALSE; // to quit this instance
		}
	}

	// if no one handled it simply create a new instance
	CToDoListWnd* pTDL = new CToDoListWnd;
	
	if (pTDL && pTDL->Create(startup))
	{
		m_pMainWnd = pTDL;
		return TRUE;
	}

	// else
	return FALSE; // quit app
}

BOOL CToDoListApp::SendStartupParams(HWND hWnd, const CTDCStartupOptions& startup, TDL_COPYDATA nMsg)
{
	COPYDATASTRUCT cds = { nMsg, sizeof(startup), (void*)&startup };

	if (::IsWindow(hWnd))
		return (::SendMessage(hWnd, WM_COPYDATA, NULL, (LPARAM)&cds) != 0);

	// else
	return FALSE;
}

BOOL CToDoListApp::ValidateTasklistPath(CString& sPath)
{
	if (FileMisc::HasExtension(sPath, _T("tdl")) ||
		FileMisc::HasExtension(sPath, _T("xml")))
	{
		return ValidateFilePath(sPath);
	}

	// else log it
	FileMisc::LogText(_T("Taskfile '%s' had an invalid extension\n"), sPath);

	return FALSE;
}

BOOL CToDoListApp::GetDefaultIniPath(CString& sIniPath, BOOL bCheckExists)
{
	CString sExePath = FileMisc::GetAppFilePath();
	CString sExeFolder = FileMisc::GetAppFolder();
	
	// preferred location is app folder for portability
	CString sTestIniPath = FileMisc::ReplaceExtension(sExePath, _T("ini"));
	
	// but it may not be writable
	if (!FileMisc::IsFolderWritable(sExeFolder))
	{
		// try app data
		CString sAppDataFolder;
		VERIFY (FileMisc::GetSpecialFilePath(CSIDL_APPDATA, APPDATA, sAppDataFolder));
		
		// if the exe Ini exists then copy it to this location
		// if there is nothing there already
		CString sAppDataIniPath = (FileMisc::TerminatePath(sAppDataFolder) +
									FileMisc::GetFileNameFromPath(sTestIniPath));
		
		if (FileMisc::FileExists(sTestIniPath) && !FileMisc::FileExists(sAppDataIniPath))
		{
			FileMisc::CreateFolderFromFilePath(sAppDataIniPath);
			
			if (FileMisc::CopyFile(sTestIniPath, sAppDataIniPath))
			{
				FileMisc::DeleteFile(sTestIniPath, TRUE);
			}
		}
		
		sTestIniPath = sAppDataIniPath;
	}
	
	// first try file having the same name/location as the executable
	if (ValidateIniPath(sTestIniPath, bCheckExists))
	{
		sIniPath = sTestIniPath;
		return TRUE;
	}
	
	// else
	return FALSE;
}

BOOL CToDoListApp::ValidateIniPath(CString& sFilePath, BOOL bCheckExists)
{
	ASSERT(!::PathIsRelative(sFilePath));

	if (::PathIsRelative(sFilePath))
		return FALSE;

	CString sIniPath(sFilePath);
	FileMisc::ReplaceExtension(sIniPath, _T("ini"));

	// check existence as required
	BOOL bFileExists = FileMisc::FileExists(sIniPath);

	if (bCheckExists && !bFileExists)
		return FALSE;

	// make sure it's writable
	if (bFileExists)
		::SetFileAttributes(sIniPath, FILE_ATTRIBUTE_NORMAL);

	// check containing folder is writable
	CString sIniFolder = FileMisc::GetFolderFromFilePath(sIniPath);

	FileMisc::CreateFolder(sIniFolder);
	::SetFileAttributes(sIniFolder, FILE_ATTRIBUTE_NORMAL);
		
	if (FileMisc::IsFolderWritable(sIniFolder))
	{
		sFilePath = sIniPath;
		return TRUE;
	}

	return FALSE;
}

BOOL CToDoListApp::ValidateFilePath(CString& sPath, const CString& sExt)
{
	if (sPath.IsEmpty())
		return FALSE;

	CString sTemp(sPath);

	if (!sExt.IsEmpty())
		FileMisc::ReplaceExtension(sTemp, sExt);

	// don't change sTemp
	CString sFullPath(sTemp);

	// if relative check app folder first
	if (::PathIsRelative(sTemp))
	{
		sFullPath = FileMisc::GetFullPath(sTemp, FileMisc::GetAppFolder());

		// then try CWD
		if (!FileMisc::FileExists(sFullPath))
			sFullPath = FileMisc::GetFullPath(sTemp);
	}

	// test file existence
	if (FileMisc::FileExists(sFullPath))
	{
		sPath = sFullPath;
		return TRUE;
	}

	// else log it
	if (FileMisc::IsLoggingEnabled())
	{
		if (::PathIsRelative(sTemp))
		{
			FileMisc::LogText(_T("File '%s' not found in '%s' or '%s'\n"),
								sTemp, 
								FileMisc::GetAppFolder(), 
								FileMisc::GetCwd());
		}
		else 
		{
			FileMisc::LogText(_T("File '%s' not found\n"), sFullPath);
		}
	}

	return FALSE;
}

// our own local version
CString CToDoListApp::AfxGetAppName()
{
	return ((CToDoListWnd*)m_pMainWnd)->GetTitle();
}

void CToDoListApp::ParseCommandLine(CEnCommandLineInfo& cmdInfo)
{
	CWinApp::ParseCommandLine(cmdInfo); // default

	m_bUseStaging = cmdInfo.HasOption(SWITCH_STAGING);

	// turn on logging if requested
    if (cmdInfo.HasOption(SWITCH_LOGGING))
		FileMisc::EnableLogging(TRUE, _T("Abstractspoon"));

	// validate ini path if present
    if (cmdInfo.HasOption(SWITCH_INIFILE))
	{
		CString sIniPath = cmdInfo.GetOption(SWITCH_INIFILE);

		// always delete existing item
		cmdInfo.DeleteOption(SWITCH_INIFILE);

		// must have ini extension
		if (FileMisc::HasExtension(sIniPath, _T("ini")) &&
			ValidateFilePath(sIniPath))
		{
			// save full path
			cmdInfo.SetOption(SWITCH_INIFILE, sIniPath);
		}
	}

	// validate import path
	if (cmdInfo.HasOption(SWITCH_IMPORT))
	{
		CString sImportPath = cmdInfo.GetOption(SWITCH_IMPORT);

		// always delete existing item
		cmdInfo.DeleteOption(SWITCH_IMPORT);
		
		if (ValidateFilePath(sImportPath))
		{
			// save full path
			cmdInfo.SetOption(SWITCH_IMPORT, sImportPath);
		}
	}

	// validate main file path
	if (!cmdInfo.m_strFileName.IsEmpty())
	{
		if (!ValidateTasklistPath(cmdInfo.m_strFileName))
			cmdInfo.m_strFileName.Empty();
	}

	// validate multiple filepaths
	if (cmdInfo.HasOption(SWITCH_TASKFILE))
	{
		CString sTaskFiles = cmdInfo.GetOption(SWITCH_TASKFILE);
		CStringArray aTaskFiles; 

		int nFile = Misc::Split(sTaskFiles, aTaskFiles, '|');

		while (nFile--)
		{
			CString& sTaskfile = aTaskFiles[nFile];

			if (!ValidateTasklistPath(sTaskfile))
			{
				aTaskFiles.RemoveAt(nFile);
			}
			// also remove if it's a dupe of m_strFileName
			else if (cmdInfo.m_strFileName == sTaskfile)
			{
				aTaskFiles.RemoveAt(nFile);
			}
		}

		// save results
		cmdInfo.DeleteOption(SWITCH_TASKFILE);

		if (aTaskFiles.GetSize())
			cmdInfo.SetOption(SWITCH_TASKFILE, Misc::FormatArray(aTaskFiles, '|'));
	}
}

BOOL CToDoListApp::PreTranslateMessage(MSG* pMsg) 
{
	// give first chance to main window for handling accelerators
	if (m_pMainWnd && m_pMainWnd->PreTranslateMessage(pMsg))
		return TRUE;

	// -------------------------------------------------------------------
	// Implement CWinApp::PreTranslateMessage(pMsg)	ourselves
	// so as to not call CMainFrame::PreTranslateMessage(pMsg) twice

	// if this is a thread-message, short-circuit this function
	if (pMsg->hwnd == NULL && DispatchThreadMessageEx(pMsg))
		return TRUE;

	// walk from target to main window but excluding main window
	for (HWND hWnd = pMsg->hwnd; 
		hWnd != NULL && hWnd != m_pMainWnd->GetSafeHwnd(); 
		hWnd = ::GetParent(hWnd))
	{
		CWnd* pWnd = CWnd::FromHandlePermanent(hWnd);

		if (pWnd != NULL)
		{
			// target window is a C++ window
			if (pWnd->PreTranslateMessage(pMsg))
				return TRUE; // trapped by target window (eg: accelerators)
		}
	}
	// -------------------------------------------------------------------

	return FALSE;       // no special processing
	//return CWinApp::PreTranslateMessage(pMsg);
}

void CToDoListApp::OnHelpWiki() 
{ 
	DoHelp();
}

void CToDoListApp::WinHelp(DWORD dwData, UINT nCmd) 
{
	if (nCmd == HELP_CONTEXT)
		DoHelp((LPCTSTR)dwData);
}

void CToDoListApp::DoHelp(const CString& /*sHelpRef*/)
{
	CString sHelpUrl(WIKI_URL);
	FileMisc::Run(*m_pMainWnd, WIKI_URL, NULL, SW_SHOWNORMAL);
}

BOOL CToDoListApp::InitPreferences(CEnCommandLineInfo& cmdInfo)
{
	BOOL bUseIni = FALSE;
	BOOL bSetMultiInstance = FALSE;
	BOOL bRegKeyExists = CRegKey2::KeyExists(HKEY_CURRENT_USER, APPREGKEY);

#ifdef _DEBUG
	BOOL bQuiet = cmdInfo.HasOption(SWITCH_QUIET);
#else
	BOOL bQuiet = FALSE;
#endif

    CString sIniPath;

    // try command line override first
    if (cmdInfo.GetOption(SWITCH_INIFILE, sIniPath))
    {
		ASSERT(!sIniPath.IsEmpty());
		ASSERT(!::PathIsRelative(sIniPath));
		ASSERT(FileMisc::PathExists(sIniPath));

		bUseIni = ValidateIniPath(sIniPath, FALSE);

		if (!bUseIni)
		{
			FileMisc::LogText(_T("Specified ini file not found: %s"), sIniPath);
			sIniPath.Empty();
		}
	}
	else if (!cmdInfo.m_strFileName.IsEmpty())
	{
		ASSERT(!::PathIsRelative(cmdInfo.m_strFileName));
		ASSERT(FileMisc::PathExists(cmdInfo.m_strFileName));
		
		// else if there is a tasklist on the commandline 
		// then try for an ini file of the same name
		sIniPath = cmdInfo.m_strFileName;

		if (ValidateIniPath(sIniPath, TRUE))
		{
			bUseIni = TRUE;
			FileMisc::LogText(_T("Ini file matching specified tasklist found: %s"), sIniPath);

			// enable multi-instance because this tasklist
			// has its own ini file
			bSetMultiInstance = TRUE;
		}
	}

	// if all else fails then try for the default ini file
	if (!bUseIni)
	{
		if (GetDefaultIniPath(sIniPath, TRUE))
		{
			bUseIni = TRUE;
		}
		else // handle quiet start
		{
			bUseIni = bQuiet;
		}
	}

	// Has the user already chosen a language?
	BOOL bFirstTime = (!bUseIni && !bRegKeyExists);

	// check existing prefs
	if (!bFirstTime)
	{
		if (bUseIni)
		{
			FileMisc::LogText(_T("Using existing ini file for preferences: %s"), sIniPath);
			
			free((void*)m_pszProfileName);
			m_pszProfileName = _tcsdup(sIniPath);
		}
		else
		{
			FileMisc::LogText(_T("Using existing registry settings for preferences"));
			
			SetRegistryKey(REGKEY);
		}

		// language is stored as relative path
		CPreferences prefs;
		m_sLanguageFile = prefs.GetProfileString(_T("Preferences"), _T("LanguageFile"));

		if (!m_sLanguageFile.IsEmpty() && (m_sLanguageFile != CTDLLanguageComboBox::GetDefaultLanguage()))
		{
			FileMisc::MakeFullPath(m_sLanguageFile, FileMisc::GetAppFolder());
		}
	}

	// show language dialog if no language set
	if (m_sLanguageFile.IsEmpty())
	{
		if (bQuiet || !CTDLLanguageComboBox::HasLanguages())
		{
			m_sLanguageFile = CTDLLanguageComboBox::GetDefaultLanguage();
		}
		else
		{
			CTDLLanguageDlg dialog;
			
			if (dialog.DoModal() == IDCANCEL)
				return FALSE; // quit app

			// else
			m_sLanguageFile = dialog.GetLanguageFile();
		}
	}

	// init language translation. 
	// 'u' indicates uppercase mode
	if (cmdInfo.HasOption(SWITCH_TRANSUPPER))
	{
		CLocalizer::Release();
		CLocalizer::Initialize(m_sLanguageFile, ITTTO_UPPERCASE);
	}
	// 't' indicates 'translation' mode (aka 'Add2Dictionary')
	else if (FileMisc::FileExists(m_sLanguageFile))
	{
		CLocalizer::Release();

		if (cmdInfo.HasOption(SWITCH_ADDTODICT))
			CLocalizer::Initialize(m_sLanguageFile, ITTTO_ADD2DICTIONARY);
		else
			CLocalizer::Initialize(m_sLanguageFile, ITTTO_TRANSLATEONLY);
	}
	
	// save language choice 
	if (!bFirstTime)
	{
		CPreferences prefs;
		
		FileMisc::MakeRelativePath(m_sLanguageFile, FileMisc::GetAppFolder(), FALSE);
		prefs.WriteProfileString(_T("Preferences"), _T("LanguageFile"), m_sLanguageFile);

		if (bSetMultiInstance)
			prefs.WriteProfileInt(_T("Preferences"), _T("MultiInstance"), TRUE);
		
		UpgradePreferences(prefs);
		prefs.Save();

		// check for web updates
		if (prefs.GetProfileInt(_T("Preferences"), _T("AutoCheckForUpdates"), FALSE))
		{
			TDL_WEBUPDATE_CHECK nCheck = CheckForUpdates(FALSE);

			if ((nCheck == TDLWUC_WANTUPDATE) || (nCheck == TDLWUC_WANTPRERELEASEUPDATE))
			{
				RunUpdater(nCheck == TDLWUC_WANTPRERELEASEUPDATE);
				return FALSE; // quit app
			}
		}
	}
	else  // first time so no ini file exists. show wizard
	{
		ASSERT(!bQuiet);
		FileMisc::LogText(_T("Neither ini file nor registry settings found -> Showing setup wizard"));

		CTDLWelcomeWizard wizard;

		// before we show the wizard we need to enable '.tdl' 
		// as a recognized extension else the file icon will 
		// not display correctly on the last page of the wizard
		CFileRegister filereg(_T("tdl"), _T("tdl_Tasklist"));
		filereg.RegisterFileType(_T("Tasklist"), 0);
		
		int nRet = wizard.DoModal();
			
		// remove file extension enabling before continuing
		filereg.UnRegisterFileType();
			
		if (nRet != ID_WIZFINISH)
			return FALSE; // quit app
		
		// use whichever ini location has a writable folder
		bUseIni = (wizard.GetUseIniFile() && GetDefaultIniPath(sIniPath, FALSE));

		if (bUseIni)
		{
			FileMisc::LogText(_T("Using ini for preferences: %s"), sIniPath);

			free((void*)m_pszProfileName);
			m_pszProfileName = _tcsdup(sIniPath);
		}
		else
		{
			FileMisc::LogText(_T("Using registry for preferences"));
			
			SetRegistryKey(REGKEY);
		}

		// initialize prefs to defaults
		CPreferences prefs;
		CPreferencesDlg().Initialize(prefs);

		// set up some default preferences
		if (wizard.GetShareTasklists()) 
		{
			// set up source control for remote tasklists
			prefs.WriteProfileInt(_T("Preferences"), _T("EnableSourceControl"), TRUE);
			prefs.WriteProfileInt(_T("Preferences"), _T("SourceControlLanOnly"), TRUE);
			prefs.WriteProfileInt(_T("Preferences"), _T("AutoCheckOut"), TRUE);
			prefs.WriteProfileInt(_T("Preferences"), _T("CheckoutOnCheckin"), TRUE);
			prefs.WriteProfileInt(_T("Preferences"), _T("CheckinOnClose"), TRUE);
			prefs.WriteProfileInt(_T("Preferences"), _T("CheckinNoEditTime"), 1);
			prefs.WriteProfileInt(_T("Preferences"), _T("CheckinNoEdit"), TRUE);
			prefs.WriteProfileInt(_T("Preferences"), _T("Use3rdPartySourceControl"), FALSE);
		}
		
		// setup default columns
		CTDCColumnIDArray aColumns;
		wizard.GetVisibleColumns(aColumns);
		
		int nCol = aColumns.GetSize();
		prefs.WriteProfileInt(_T("Preferences\\ColumnVisibility"), _T("Count"), nCol);
		
		while (nCol--)
		{
			CString sKey = Misc::MakeKey(_T("Col%d"), nCol);
			prefs.WriteProfileInt(_T("Preferences\\ColumnVisibility"), sKey, aColumns[nCol]);
		}
		
		if (wizard.GetHideAttributes())
		{
			// hide clutter
			prefs.WriteProfileInt(_T("Preferences"), _T("ShowCtrlsAsColumns"), TRUE);
			prefs.WriteProfileInt(_T("Preferences"), _T("ShowEditMenuAsColumns"), TRUE);
		}
		
		// set up initial file
		CString sSample = wizard.GetSampleFilePath();
		
		if (!sSample.IsEmpty())
			cmdInfo.m_strFileName = sSample;

		// save language choice
		FileMisc::MakeRelativePath(m_sLanguageFile, FileMisc::GetAppFolder(), FALSE);
		prefs.WriteProfileString(_T("Preferences"), _T("LanguageFile"), m_sLanguageFile);

		// setup uninstaller to point to us
		if (!bUseIni)
		{
			CRegKey2 reg;

			if (reg.Open(HKEY_LOCAL_MACHINE, UNINSTALLREGKEY, FALSE) == ERROR_SUCCESS)
			{
				CString sUninstall;
				sUninstall.Format(_T("\"%s\" -%s"), FileMisc::GetAppFilePath(), SWITCH_UNINSTALL);

				reg.Write(_T("DisplayName"), CToDoListWnd::GetTitle());			
				reg.Write(_T("NoModify"), 1);
				reg.Write(_T("NoRepair"), 1);
				reg.Write(_T("UninstallString"), sUninstall);
			}
		}
	}

	return TRUE;
}

void CToDoListApp::UpgradePreferences(CPreferences& prefs)
{
	UNREFERENCED_PARAMETER(prefs);

	// we don't handle the registry because it's too hard (for now)
	if (m_pszRegistryKey)
		return;
	
	// remove preferences for all files _not_ in the MRU list
	// provided there's at least one file in the MRU list
	//
	// NEEDS WORK
/*
	BOOL bUseMRU = prefs.GetProfileInt(_T("Preferences"), _T("AddFilesToMRU"), FALSE);

	if (!bUseMRU)
		return;

	CStringArray aMRU;
	
	for (int nFile = 0; nFile < 16; nFile++)
	{
		CString sItem, sFile;
		
		sItem.Format(_T("TaskList%d"), nFile + 1);
		sFile = prefs.GetProfileString(_T("MRU"), sItem);
		
		if (sFile.IsEmpty())
			break;
		
		// else
		sFile = FileMisc::GetFileNameFromPath(sFile);
		Misc::AddUniqueItem(sFile, aMRU);
	}
	
	if (aMRU.GetSize())
	{
		CStringArray aSections;
		int nSection = prefs.GetSectionNames(aSections);
		
		while (nSection--)
		{
			const CString& sSection = aSections[nSection];
			
			// does it start with "FileStates\\" 
			if (sSection.Find(FILESTATEKEY) == 0 || sSection.Find(REMINDERKEY) == 0)
			{
				// split the section name into its parts
				CStringArray aSubSections;
				int nSubSections = Misc::Split(sSection, aSubSections, '\\');
				
				if (nSubSections > 1)
				{
					// the file name is the second item
					const CString& sFilename = aSubSections[1];

					// make sure it's an actual filepath
					if (FileMisc::IsPath(sFilename))
					{
						// ignore 'Default'
						if (sFilename.CompareNoCase(DEFAULTKEY) != 0)
						{
							if (Misc::Find(aMRU, sFilename) == -1)
								prefs.DeleteSection(sSection);
						}
					}
				}
			}
		}
	}
*/
}

int CToDoListApp::DoMessageBox(LPCTSTR lpszPrompt, UINT nType, UINT /*nIDPrompt*/) 
{
	HWND hwndMain = NULL;

	// make sure app window is visible
	if (m_pMainWnd)
	{
		hwndMain = *m_pMainWnd;
		m_pMainWnd->SendMessage(WM_TDL_SHOWWINDOW, 0, 0);
	}
	else
	{
		hwndMain = ::GetDesktopWindow();
	}
	
	CString sTitle(AfxGetAppName()), sInstruction, sText(lpszPrompt);
	CStringArray aPrompt;
	
	int nNumInputs = Misc::Split(lpszPrompt, aPrompt, '|');
	
	switch (nNumInputs)
	{
	case 0:
		ASSERT(0);
		break;
		
	case 1:
		// do nothing
		break;
		
	case 2:
		sInstruction = aPrompt[0];
		sText = aPrompt[1];
		break;
		
	case 3:
		sTitle, aPrompt[0];
		sInstruction = aPrompt[1];
		sText = aPrompt[2];
	}
	
	return CDialogHelper::ShowMessageBox(hwndMain, sTitle, sInstruction, sText, nType);
}

void CToDoListApp::OnImportPrefs() 
{
	// default location is always app folder
	CString sIniPath = FileMisc::GetAppFilePath();
	sIniPath.MakeLower();
	sIniPath.Replace(_T("exe"), _T("ini"));
	
	CPreferences prefs;
	CFileOpenDialog dialog(IDS_IMPORTPREFS_TITLE, 
							_T("ini"), 
							sIniPath, 
							EOFN_DEFAULTOPEN, 
							CEnString(IDS_INIFILEFILTER));
	
	if (dialog.DoModal(&prefs) == IDOK)
	{
		CRegKey2 reg;
		
		if (reg.Open(HKEY_CURRENT_USER, APPREGKEY, FALSE) == ERROR_SUCCESS)
		{
			sIniPath = dialog.GetPathName();
			
			if (reg.ImportFromIni(sIniPath)) // => import ini to registry
			{
				// use them now?
				// only ask if we're not already using the registry
				BOOL bUsingIni = (m_pszRegistryKey == NULL);

				if (bUsingIni)
				{
					if (AfxMessageBox(CEnString(IDS_POSTIMPORTPREFS), MB_YESNO | MB_ICONQUESTION) == IDYES)
					{
						// renames existing prefs file
						CString sNewName = (sIniPath + _T(".bak"));
						
						if (FileMisc::MoveFile(sIniPath, sNewName, TRUE, TRUE))
						{
							// and initialize the registry 
							SetRegistryKey(REGKEY);
							
							// reset prefs
							m_pMainWnd->SendMessage(WM_TDL_REFRESHPREFS);
						}
					}
				}
				else // reset prefs
				{
					m_pMainWnd->SendMessage(WM_TDL_REFRESHPREFS);
				}
			}
			else // notify user
			{
				CEnString sMessage(CEnString(IDS_INVALIDPREFFILE), dialog.GetFileName());
				AfxMessageBox(sMessage, MB_OK | MB_ICONEXCLAMATION);
			}
		}
	}
}

void CToDoListApp::OnUpdateImportPrefs(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
}

void CToDoListApp::OnExportPrefs() 
{
	ASSERT (m_pszRegistryKey != NULL);

	CRegKey2 reg;

	if (reg.Open(HKEY_CURRENT_USER, APPREGKEY, TRUE) == ERROR_SUCCESS)
	{
		// default location is always app folder
		CString sAppPath = FileMisc::GetAppFilePath();

		CString sIniPath(sAppPath);
		sIniPath.MakeLower();
		sIniPath.Replace(_T("exe"), _T("ini"));
		
		CPreferences prefs;
		CFileSaveDialog dialog(IDS_IMPORTPREFS_TITLE, 
								_T("ini"), 
								sIniPath, 
								EOFN_DEFAULTSAVE, 
								CEnString(IDS_INIFILEFILTER));
		
		if (dialog.DoModal(&prefs) == IDOK)
		{
			BOOL bUsingReg = (m_pszRegistryKey != NULL);
			sIniPath = dialog.GetPathName();

			if (bUsingReg && reg.ExportToIni(sIniPath))
			{
				// use them now? 
				CString sAppFolder, sIniFolder;
				
				FileMisc::SplitPath(sAppPath, NULL, &sAppFolder);
				FileMisc::SplitPath(sIniPath, NULL, &sIniFolder);
				
				// only if they're in the same folder as the exe
				if (sIniFolder.CompareNoCase(sAppFolder) == 0)
				{
					if (AfxMessageBox(CEnString(IDS_POSTEXPORTPREFS), MB_YESNO | MB_ICONQUESTION) == IDYES)
					{
						free((void*)m_pszRegistryKey);
						m_pszRegistryKey = NULL;
						
						free((void*)m_pszProfileName);
						m_pszProfileName = _tcsdup(sIniPath);
						
						// reset prefs
						m_pMainWnd->SendMessage(WM_TDL_REFRESHPREFS);
					}
				}
			}
		}
	}
}

void CToDoListApp::OnUpdateExportPrefs(CCmdUI* pCmdUI) 
{
	BOOL bUsingReg = (m_pszRegistryKey != NULL);
	pCmdUI->Enable(bUsingReg);
}

void CToDoListApp::OnHelpGoogleGroup() 
{
	FileMisc::Run(*m_pMainWnd, GOOGLEGROUP_URL);
}

void CToDoListApp::OnHelpLinkedIn() 
{
	FileMisc::Run(*m_pMainWnd, LINKEDIN_URL);
}

void CToDoListApp::OnHelpFacebook() 
{
	FileMisc::Run(*m_pMainWnd, FACEBOOK_URL);
}

void CToDoListApp::OnHelpTwitter() 
{
	FileMisc::Run(*m_pMainWnd, TWITTER_URL);
}

void CToDoListApp::OnHelpGooglePlus() 
{
	FileMisc::Run(*m_pMainWnd, GOOGLEPLUS_URL);
}

void CToDoListApp::OnHelpLicense() 
{
	FileMisc::Run(*m_pMainWnd, LICENSE_URL);
}

void CToDoListApp::OnHelpCommandline() 
{
	CTDLCmdlineOptionsDlg dialog;
	dialog.DoModal();
}

void CToDoListApp::OnHelpDonate() 
{
	FileMisc::Run(*m_pMainWnd, DONATE_URL);
}

int CToDoListApp::ExitInstance() 
{
	// TODO: Add your specialized code here and/or call the base class
	CLocalizer::Release();

	return CWinApp::ExitInstance();
}

void CToDoListApp::OnHelpUninstall() 
{
	// confirm uninstall first
	if (AfxMessageBox(CEnString(IDS_CONFIRM_UNINSTALL), MB_YESNO | MB_ICONWARNING) == IDYES)
	{
		RunUninstaller();
	}
}

DWORD CToDoListApp::RunHelperApp(const CString& sAppName, UINT nIDGenErrorMsg, UINT nIDSmartScreenErrorMsg, BOOL bPreRelease)
{
	CEnCommandLineInfo params;

	// Pass the centroid of the main wnd so that the
	// updater appears in that same location
	if (m_pMainWnd && (GetSystemMetrics(SM_CMONITORS) > 1))
	{
		CRect rWindow;
		m_pMainWnd->GetWindowRect(rWindow);
		CPoint ptPos = rWindow.CenterPoint();

		params.SetOption(SWITCH_POSITION, MAKELPARAM(ptPos.x, ptPos.y));
	}
		
	// the helper app path
	CString sAppFolder = FileMisc::GetAppFolder();

#ifdef _DEBUG // -----------------------------------------------------------------------

	// Copy ourselves to a temp location
	CString sTempFolder = (FileMisc::TerminatePath(sAppFolder) + _T("Debug") + sAppName);
	
	// sanity check
	VERIFY(FileMisc::DeleteFolder(sTempFolder, FMDF_SUBFOLDERS | FMDF_HIDDENREADONLY));
	
	// create folder and copy
	VERIFY(FileMisc::CreateFolder(sTempFolder));
	VERIFY(FileMisc::CopyFolder(sAppFolder, sTempFolder, _T("*.exe;*.dll;*.ini"), FMDF_HIDDENREADONLY));

	// set the temp folder as the app folder
	sAppFolder = sTempFolder;

#else // Release -----------------------------------------------------------------------

	// try to close all open instances of TDL
	if (!CloseAllToDoListWnds())
	{
		return 0; // user cancelled
	}

#endif // ------------------------------------------------------------------------------
	
	CString sAppPath;
	FileMisc::MakePath(sAppPath, NULL, sAppFolder, sAppName, _T("exe"));
	
	// to handle UAC on Vista and above we use the "RunAs" verb
	LPCTSTR szVerb = ((COSVersion() >= OSV_VISTA) ? _T("runas") : NULL);

	// pass our app id to app 
	params.SetOption(SWITCH_APPID, TDLAPPID);

	// and the commandline we were started with
	// use base64 encoding to mangle it so that the update
	// doesn't try to interpret the commandline itself
	params.SetOption(SWITCH_CMDLINE, Base64Coder::Encode(m_lpCmdLine));
	
	if (CRTLStyleMgr::IsRTL())
		params.SetOption(SWITCH_RTL);
	
	// and the current language
	if (m_sLanguageFile != CTDLLanguageComboBox::GetDefaultLanguage())
	{
		CString sLangFile = FileMisc::GetFullPath(m_sLanguageFile, FileMisc::GetAppFolder());
		params.SetOption(SWITCH_LANG, sLangFile);
		
		if (CLocalizer::GetTranslationOption() == ITTTO_ADD2DICTIONARY)
		{
			params.SetOption(SWITCH_ADDTODICT);

			// make sure all dictionary changes have been written
			CLocalizer::Release();
		}
	}

	// and whether this is a pre-prelease
	if (bPreRelease)
		params.SetOption(SWITCH_PRERELEASE);
	DWORD dwRes = FileMisc::Run(NULL, 
								sAppPath, 
								params.GetCommandLine(), 
								SW_SHOWNORMAL,
								NULL, 
								szVerb);
	
	// error handling
	if (dwRes <= 32)
	{
		switch (dwRes)
		{
		case SE_ERR_ACCESSDENIED:
			// if this is windows 8 or above, assume 
			// this was blocked by SmartScreen
			if (COSVersion() >= OSV_WIN8)
				AfxMessageBox(nIDSmartScreenErrorMsg);

			// else fall thru
			break;
		}

		// all else
		AfxMessageBox(nIDGenErrorMsg);
	}

	return dwRes;
}

void CToDoListApp::RunUninstaller()
{
	RunHelperApp(_T("TDLUninstall"), IDS_UNINSTALLER_RUNFAILURE, IDS_UNINSTALLER_SMARTSCREENBLOCK);
}

void CToDoListApp::RunUpdater(BOOL bPreRelease)
{
	RunHelperApp(_T("TDLUpdate"), IDS_UPDATER_RUNFAILURE, IDS_UPDATER_SMARTSCREENBLOCK, bPreRelease);
}

BOOL CToDoListApp::CloseAllToDoListWnds()
{
	TDCFINDWND find(NULL, TRUE);
	int nNumWnds = FindToDoListWnds(find);
	
	for (int nWnd = 0; nWnd < nNumWnds; nWnd++)
	{
		HWND hwndMain = find.aResults[nWnd];
		
		// final check for validity
		if (::IsWindow(hwndMain))
		{
			// close the application
			::SendMessage(hwndMain, WM_CLOSE, 0, 1);
		}
	}

	// check for user cancellation by seeing 
	// if any windows are still open
	return (FindToDoListWnds(find) == 0);
}

int CToDoListApp::FindToDoListWnds(TDCFINDWND& find)
{
	ASSERT(find.hWndIgnore == NULL || ::IsWindow(find.hWndIgnore));

	find.aResults.RemoveAll();
	EnumWindows(FindOtherInstance, (LPARAM)&find);

	return find.aResults.GetSize();
}

BOOL CALLBACK CToDoListApp::FindOtherInstance(HWND hWnd, LPARAM lParam)
{
	static CString COPYRIGHT(MAKEINTRESOURCE(IDS_COPYRIGHT));

	CString sCaption;
	CWnd::FromHandle(hWnd)->GetWindowText(sCaption);

	if (sCaption.Find(COPYRIGHT) != -1)
	{
		TDCFINDWND* pFind = (TDCFINDWND*)lParam;
		ASSERT(pFind);

		// check window to ignore
		if ((pFind->hWndIgnore == NULL) || (pFind->hWndIgnore == hWnd))
		{
			// check if it's closing
			DWORD bClosing = FALSE;
			BOOL bSendSucceeded = ::SendMessageTimeout(hWnd, 
														WM_TDL_ISCLOSING, 
														0, 
														0, 
														SMTO_ABORTIFHUNG | SMTO_BLOCK, 
														1000, 
														&bClosing);

			if (bSendSucceeded && (pFind->bIncClosing || !bClosing))
			{
				if (pFind->dwProcessID)
				{
					DWORD dwOtherProcID = 0;
					GetWindowThreadProcessId(hWnd, &dwOtherProcID);
					ASSERT(dwOtherProcID);

					if (dwOtherProcID == pFind->dwProcessID)
					{
						pFind->aResults.Add(hWnd);
						return FALSE; // we can stop now
					}
				}
				else
				{
					pFind->aResults.Add(hWnd);
				}
			}			
		}
	}

	return TRUE; // keep going to the end
}

BOOL CToDoListApp::WaitForInstanceToClose(DWORD dwProcID)
{
	int nTry = 10;

	while (nTry--)
	{
		HANDLE handle = ::OpenProcess(PROCESS_ALL_ACCESS, FALSE, dwProcID);

		if (!handle)
		{
			// no process having this ID so the
			// instance of TDL must have closed
			return TRUE;
		}

		CString sModule = FileMisc::GetProcessFilePath(handle);
		sModule.MakeUpper();

		if (sModule.Find(_T("TODOLIST.EXE")) == -1)
		{
			// the process ID has been reallocated
			// so the instance of TDL must have closed
			return TRUE;
		}

 		::CloseHandle(handle);

		// try again
		Sleep(200);
	}

	TRACE(_T("%d hasn't closed..."), dwProcID);


	return FALSE;
}

void CToDoListApp::OnHelpCheckForUpdates() 
{
	TDL_WEBUPDATE_CHECK nCheck = CheckForUpdates(TRUE);

	switch (nCheck)
	{
	case TDLWUC_WANTUPDATE:
	case TDLWUC_WANTPRERELEASEUPDATE:
		RunUpdater(nCheck == TDLWUC_WANTPRERELEASEUPDATE);
		break;

	case TDLWUC_CANCELLED:
		return;
		
	case TDLWUC_NOTCONNECTED:
		AfxMessageBox(CEnString(IDS_NO_WEBUPDATE_CONNECTION));
		break;

	case TDLWUC_NOUPDATES:
		AfxMessageBox(CEnString(IDS_NO_WEBUPDATE_AVAIL));
		break;

	case TDLWUC_FAILED:
	default:
		// TODO
		break;
	}
}

TDL_WEBUPDATE_CHECK CToDoListApp::CheckForUpdates(BOOL bManual)
{
//  FOR DEBUGGING
// 	return TDLWUC_WANTUPDATE;

	CPreferences prefs;

	// only auto-check once a day
	int nLastUpdate = prefs.GetProfileInt(_T("Updates"), _T("LastUpdate"), 0);
	int nToday = (int)CDateHelper::GetDate(DHD_TODAY);

	if (!bManual && (nLastUpdate >= nToday))
		return TDLWUC_NOUPDATES;

	prefs.WriteProfileInt(_T("Updates"), _T("LastUpdate"), nToday);

	// download the update script to temp file
	return CTDLWebUpdatePromptDlg::CheckForUpdates(m_bUseStaging);
} 

/////////////////////////////////////////////////////////////////////////////

void CToDoListApp::OnDebugTaskDialogInfo() 
{
	LPCTSTR szTestMsg = _T("This the optional caption of the message box|")
						_T("This the title of the Info message box|")
						_T("This is a paragraph of text with embedded carriage \n")
						_T("returns so that it looks ok on XP and below.\n\n")
						_T("This a second paragraph that ought to have a clear\n ")
						_T("line between it and the first para.");

	AfxMessageBox(szTestMsg, MB_ICONINFORMATION);
}

void CToDoListApp::OnDebugTaskDialogWarning() 
{
	LPCTSTR szTestMsg = _T("This the optional caption of the message box|")
						_T("This the title of the Warning message box|")
						_T("This is a paragraph of text with embedded carriage \n")
						_T("returns so that it looks ok on XP and below.\n\n")
						_T("This a second paragraph that ought to have a clear\n ")
						_T("line between it and the first para.");

	AfxMessageBox(szTestMsg, MB_ICONWARNING);
}

void CToDoListApp::OnDebugTaskDialogError() 
{
	LPCTSTR szTestMsg = _T("This the optional caption of the message box|")
						_T("This the title of the Error message box|")
						_T("This is a paragraph of text with embedded carriage \n")
						_T("returns so that it looks ok on XP and below.\n\n")
						_T("This a second paragraph that ought to have a clear\n ")
						_T("line between it and the first para.");
	
	AfxMessageBox(szTestMsg, MB_ICONERROR);
}

void CToDoListApp::OnDebugShowUpdateDlg() 
{
	CString sAppPath = (FileMisc::TerminatePath(FileMisc::GetAppFolder()) + _T("TDLUpdate.exe"));

	// pass our app id to app 
	CEnCommandLineInfo cmdLine;

	cmdLine.SetOption(SWITCH_APPID, TDLAPPID);
	cmdLine.SetOption(SWITCH_SHOWUI);

	// Pass the centroid of the main wnd so that the
	// updater appears in that same location
	if (GetSystemMetrics(SM_CMONITORS) > 1)
	{
		CRect rWindow;
		m_pMainWnd->GetWindowRect(rWindow);
		CPoint ptPos = rWindow.CenterPoint();

		cmdLine.SetOption(SWITCH_POSITION, MAKELPARAM(ptPos.x, ptPos.y));
	}

	DWORD dwRes = FileMisc::Run(NULL, sAppPath, cmdLine.GetCommandLine());
}

void CToDoListApp::OnDebugShowScriptDlg() 
{
	CTDLWebUpdatePromptDlg::CheckForUpdates(FALSE, _T("0.0.997.0"));
}

void CToDoListApp::OnHelpRecordBugReport() 
{
	FileMisc::Run(*m_pMainWnd, _T("psr.exe"));
}

BOOL CToDoListApp::CommandRequiresUI(UINT nCmdID)
{
	CMenu mainMenu;
	mainMenu.LoadMenu(IDR_MAINFRAME);

	HMENU hSubMenu = NULL;
	int nPos = CEnMenu::GetMenuItemPos(mainMenu, nCmdID, hSubMenu);

	ASSERT((nPos != -1) && hSubMenu);

	// We want the top-level sub-menu
	while ((nPos != -1) && (hSubMenu != mainMenu))
		nPos = CEnMenu::GetMenuItemPos(mainMenu, hSubMenu, hSubMenu);

	ASSERT(nPos != -1);
	ASSERT(hSubMenu);

	switch (nPos)
	{
	case AM_FILE: 
		switch (nCmdID)
		{
			// These items DO NOT require UI
		case ID_FILE_NEW:
		case ID_OPEN_RELOAD:
		case ID_CLOSE:
		case ID_CLOSEALL:
		case ID_CLOSEALLBUTTHIS:
		case ID_EXIT:
			return FALSE;
		}
		// All the rest DO require UI
		return TRUE;

	case AM_NEWTASK: 
		switch (nCmdID)
		{
			// These items DO require UI
		case ID_NEWTASK_ATTOP:
		case ID_NEWTASK_ATBOTTOM:
		case ID_NEWTASK_ATTOPSELECTED:
		case ID_NEWTASK_ATBOTTOMSELECTED:
		case ID_NEWTASK_BEFORESELECTEDTASK:
		case ID_NEWTASK_AFTERSELECTEDTASK:
		case ID_NEWSUBTASK_ATTOP:
		case ID_NEWSUBTASK_ATBOTTOM:
			return TRUE;
		}
		// All the rest DO NOT require UI
		return FALSE;

	case AM_EDIT: 
		switch (nCmdID)
		{
			// These items DO require UI
		case ID_EDIT_TASKTEXT:
		case ID_EDIT_TASKCOLOR:
		case ID_EDIT_SETTASKICON:
		case ID_ADDTIMETOLOGFILE:
		case ID_EDIT_OFFSETDATES:
		case ID_EDIT_SETFILEREF:
		case ID_EDIT_RECURRENCE:
			return TRUE;
		}
		// All the rest DO NOT require UI
		return FALSE;

	case AM_VIEW: 
		switch (nCmdID)
		{
			// These items DO require UI
		case ID_TASKLIST_SELECTCOLUMNS:
		case ID_TASKLIST_CUSTOMCOLUMNS:
		case ID_VIEW_FILTER:
			return TRUE;
		}
		// All the rest DO NOT require UI
		break;

	case AM_MOVE: 
		// NO items require UI
		return FALSE;

	case AM_SORT: 
		switch (nCmdID)
		{
			// These items DO require UI
		case ID_SORT_MULTI:
			return TRUE;
		}
		// All the rest DO NOT require UI
		return FALSE;

	case AM_TOOLS: 
		switch (nCmdID)
		{
			// These items DO require UI
		case ID_TOOLS_IMPORT:
		case ID_TOOLS_EXPORT:
		case ID_TOOLS_TRANSFORM:
		case ID_TOOLS_ANALYSELOGGEDTIME:
		case ID_PREFERENCES:
		case ID_TOOLS_IMPORTPREFS:
		case ID_TOOLS_EXPORTPREFS:
		case ID_TOOLS_SHOWTASKS_DUETODAY:
		case ID_TOOLS_SHOWTASKS_DUETOMORROW:
		case ID_TOOLS_SHOWTASKS_DUEENDTHISWEEK:
		case ID_TOOLS_SHOWTASKS_DUEENDNEXTWEEK:
		case ID_TOOLS_SHOWTASKS_DUEENDTHISMONTH:
		case ID_TOOLS_SHOWTASKS_DUEENDNEXTMONTH:
			return TRUE;
		}
		// All the rest DO NOT require UI
		return FALSE;

	case AM_SRCCONTROL: 
		// NO items require UI
		return FALSE;

	case AM_WINDOW: 
		// NO items require UI
		return FALSE;

	case AM_HELP: 
		switch (nCmdID)
		{
			// These items DO require UI
		case ID_HELP_UNINSTALL:
		case ID_HELP_COMMANDLINE:
		case ID_HELP_KEYBOARDSHORTCUTS:
		case ID_HELP_RECORDBUGREPORT:
		case ID_TOOLS_CHECKFORUPDATES:
		case ID_ABOUT:
			return TRUE;
		}
		// All the rest DO NOT require UI
		return FALSE;

	case AM_DEBUG:
		return TRUE; // always
	}

	// all else
	ASSERT(0);
	return FALSE;
}
