// filedialogex.cpp : implementation file
//

#include "stdafx.h"
#include "enfiledialog.h"
#include "osversion.h"
#include "enstring.h"
#include "filemisc.h"
#include "misc.h"

#include "..\Interfaces\IPreferences.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEnFileDialog

IMPLEMENT_DYNAMIC(CEnFileDialog, CFileDialog)

// Windows 2000 version of OPENFILENAME
#if _MSC_VER < 1300
struct OPENFILENAMEEX : public OPENFILENAME 
{ 
  void *        pvReserved;
  DWORD         dwReserved;
  DWORD         FlagsEx;
};
#endif		

CEnFileDialog::CEnFileDialog(BOOL bOpenFileDialog, 
							 LPCTSTR lpszTitle,
							 LPCTSTR lpszDefExt, 
							 LPCTSTR lpszFileName, 
							 DWORD dwFlags, 
							 LPCTSTR lpszFilter, 
							 CWnd* pParentWnd) 
	:
	CFileDialog(bOpenFileDialog, 
				lpszDefExt, 
				lpszFileName, 
				dwFlags, 
				lpszFilter, 
				pParentWnd), 
	m_sTitle(lpszTitle),
	m_sOrgFilters(lpszFilter)
{
#if _MSC_VER < 1300
	DWORD dwVersion = ::GetVersion();
	DWORD dwWinMajor = (DWORD)(LOBYTE(LOWORD(dwVersion)));

	if ((BYTE)dwWinMajor >= 5)
		m_ofn.lStructSize = sizeof(OPENFILENAMEEX);
#endif		

	ASSERT(dwFlags); // make sure caller sets this
}


BEGIN_MESSAGE_MAP(CEnFileDialog, CFileDialog)
	//{{AFX_MSG_MAP(CEnFileDialog)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

int CEnFileDialog::DoModal(IPreferences* pPrefs)
{
	// show the new Vista/Windows 7 file dialog
	if (COSVersion() >= OSV_VISTA)
	{
		m_ofn.lpfnHook = NULL;
		m_ofn.Flags &= ~OFN_ENABLEHOOK;
	}

	// determine the best file filter for the specified file
	// if not already specified
	if ((m_ofn.nFilterIndex <= 0) && !m_sOrgFilters.IsEmpty())
	{
		CStringArray aFilters;
		int nNumFilter = Misc::Split(m_sOrgFilters, aFilters, '|');

		// filter strings come in pairs
		if (nNumFilter && ((nNumFilter % 2) == 0))
		{
			CString sExt(m_ofn.lpstrDefExt);

			if (sExt.IsEmpty())
				sExt = FileMisc::GetExtension(m_ofn.lpstrFile);

			int nAllFilesFilter = -1; // backup filter index
			int nExtFilter = -1;

			for (int nFilter = 1; nFilter < nNumFilter; nFilter += 2)
			{
				const CString& sFilter = Misc::GetItem(aFilters, nFilter);

				if (sFilter.Find(_T(".*")) != -1)
					nAllFilesFilter = nFilter;
				
				if (!sExt.IsEmpty() && sFilter.Find(sExt) != -1)
					nExtFilter = nFilter;
			}

			if (nExtFilter != -1)
			{
				ASSERT(nExtFilter % 2);
				m_ofn.nFilterIndex = ((nExtFilter / 2) + 1); // 1-based indices
			}
			else if (nAllFilesFilter != -1)
			{
				ASSERT(nAllFilesFilter % 2);
				m_ofn.nFilterIndex = ((nAllFilesFilter / 2) + 1); // 1-based indices
			}
		}
	}

	if (!m_sTitle.IsEmpty())
		m_ofn.lpstrTitle = m_sTitle;

	// if a filename has been specified, split off the folder 
	CString sInitialFolder, sFilename;
	
	if (HasInitialFileName())
	{
		sFilename = FileMisc::GetFileNameFromPath(m_ofn.lpstrFile);
		sInitialFolder = FileMisc::GetFolderFromFilePath(m_ofn.lpstrFile);
	}

	// restore last saved folder in preference to initial dir
	if (!m_sTitle.IsEmpty())
	{
		CString sLastFolder;
      
      if (pPrefs)
         sLastFolder = pPrefs->GetProfileString(GetPreferenceKey(), m_sTitle);
      else
         sLastFolder = AfxGetApp()->GetProfileString(GetPreferenceKey(), m_sTitle);

		if (!sLastFolder.IsEmpty())
			sInitialFolder = sLastFolder;
	}

	// save off original values and apply new ones
	LPCTSTR szOrgFolder = NULL;

	if (!sInitialFolder.IsEmpty())
	{
		szOrgFolder = m_ofn.lpstrInitialDir;
		m_ofn.lpstrInitialDir = sInitialFolder;
	}

	if (!sFilename.IsEmpty())
	{
		// overwrite original filename
		lstrcpyn(m_szFileName, sFilename, (sizeof(m_szFileName) / sizeof(TCHAR)));
	}
	
	int nRet = CFileDialog::DoModal();

	// restore original folder
	if (szOrgFolder)
		m_ofn.lpstrInitialDir = szOrgFolder;
		
	// get and save off current working directory
	::GetCurrentDirectory(MAX_PATH, m_sLastFolder.GetBuffer(MAX_PATH + 1));
	m_sLastFolder.ReleaseBuffer();

	if ((nRet == IDOK) && !m_sTitle.IsEmpty())
	{
		if (pPrefs)
			pPrefs->WriteProfileString(GetPreferenceKey(), m_sTitle, m_sLastFolder);
		else
			AfxGetApp()->WriteProfileString(GetPreferenceKey(), m_sTitle, m_sLastFolder);
	}
	
	return nRet;
}

CString CEnFileDialog::GetPreferenceKey() const
{
	return (m_bOpenFileDialog ? _T("OpenFileDialog") : _T("SaveFileDialog"));
}

void CEnFileDialog::SetTitle(LPCTSTR szTitle)
{
	m_sTitle = szTitle;
}

BOOL CEnFileDialog::HasInitialFileName() const
{
	return !CString(m_ofn.lpstrFile).IsEmpty();
}

BOOL CEnFileDialog::HasInitialFolder() const
{
	return !CString(m_ofn.lpstrInitialDir).IsEmpty();
}

/////////////////////////////////////////////////////////////////////////////
// CFileOpenDialog

IMPLEMENT_DYNAMIC(CFileOpenDialog, CEnFileDialog)

CFileOpenDialog::CFileOpenDialog(LPCTSTR lpszTitle,
								LPCTSTR lpszDefExt, 
								LPCTSTR lpszFileName,
								DWORD dwFlags, 
								LPCTSTR lpszFilter, 
								CWnd* pParentWnd)
: 
	CEnFileDialog(TRUE, lpszTitle, lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd)
{
	// nothing else to do
}

CFileOpenDialog::CFileOpenDialog(UINT nIDTitle,
								LPCTSTR lpszDefExt, 
								LPCTSTR lpszFileName,
								DWORD dwFlags, 
								LPCTSTR lpszFilter, 
								CWnd* pParentWnd)
: 
	CEnFileDialog(TRUE, CEnString(nIDTitle), lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd)
{
	// nothing else to do
}


/////////////////////////////////////////////////////////////////////////////
// CEnFileDialog

IMPLEMENT_DYNAMIC(CFileSaveDialog, CEnFileDialog)

CFileSaveDialog::CFileSaveDialog(LPCTSTR lpszTitle,
								 LPCTSTR lpszDefExt, 
								 LPCTSTR lpszFileName,
								 DWORD dwFlags, 
								 LPCTSTR lpszFilter, 
								 CWnd* pParentWnd)
: 
	CEnFileDialog(FALSE, lpszTitle, lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd)
{
	// nothing else to do
}

CFileSaveDialog::CFileSaveDialog(UINT nIDTitle,
								 LPCTSTR lpszDefExt, 
								 LPCTSTR lpszFileName,
								 DWORD dwFlags, 
								 LPCTSTR lpszFilter, 
								 CWnd* pParentWnd)
: 
	CEnFileDialog(FALSE, CEnString(nIDTitle), lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd)
{
	// nothing else to do
}

