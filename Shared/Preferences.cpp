// Preferences.cpp: implementation of the CPreferences class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Preferences.h"
#include "misc.h"
#include "filemisc.h"
#include "driveinfo.h"
#include "regkey.h"
#include "autoflag.h"

#include "..\3rdparty\stdiofileex.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

static LPCTSTR ENDL = _T("\r\n");
static CString NULLSTR;

//////////////////////////////////////////////////////////////////////

CIniSectionArray	CPreferences::s_aIni;
BOOL				CPreferences::s_bDirty = FALSE;
BOOL				CPreferences::s_bIni = FALSE;
int					CPreferences::s_nRef = 0;
CCriticalSection	CPreferences::s_cs;
BOOL				CPreferences::s_bLocked = FALSE;

//////////////////////////////////////////////////////////////////////

#define LOCKPREFS()                           \
	ASSERT(!s_bLocked);                       \
	if (s_bLocked) return;                    \
	CSingleLock lock(&s_cs, TRUE);            \
	CAutoFlag af(s_bLocked, lock.IsLocked());

#define LOCKPREFSRET(ret)                     \
	ASSERT(!s_bLocked);                       \
	if (s_bLocked) return ret;                \
	CSingleLock lock(&s_cs, TRUE);            \
	CAutoFlag af(s_bLocked, lock.IsLocked());

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CPreferences::CPreferences()
{
	// prevent anyone else changing the shared resources
	// for the duration of this function
	LOCKPREFS();
			
	// if no other object is active we need to do the setup
	if (s_nRef == 0)
	{
		// figure out whether we're using an ini file or the registry
		s_bIni = (AfxGetApp()->m_pszRegistryKey == NULL);
		
		if (s_bIni)
		{
			// check for existing backup file first
			CString sIniPath(AfxGetApp()->m_pszProfileName);
			CString sBackupPath(sIniPath);
			
			FileMisc::ReplaceExtension(sBackupPath, _T("bak.ini"));
			
			if (FileMisc::FileExists(sBackupPath))
				::CopyFile(sBackupPath, sIniPath, FALSE);			
			
			// read the ini file
			CStdioFileEx file;
			
			if (file.Open(sIniPath, CFile::modeRead))
			{
				CString sLine;
				INISECTION* pCurSection = NULL;
				
				while (file.ReadString(sLine))
				{
					if (sLine.IsEmpty())
						continue;
					
					// is it a section ?
					else if (sLine[0] == '[')
					{
						CString sSection = sLine.Mid(1, sLine.GetLength() - 2);

						// assume (for speed) that the section is already unique
						pCurSection = new INISECTION(sSection);
						s_aIni.Add(pCurSection);

						ASSERT (pCurSection != NULL);
					}
					// else an entry
					else if (pCurSection)
					{
						INIENTRY ie;

						if (ie.Parse(sLine))
							SetEntryValue(*pCurSection, ie);
					}
				}
			}

			s_bDirty = FALSE;

			// delete backup
			::DeleteFile(sBackupPath);
		}
	}
				
	s_nRef++; // increment reference count
}

CPreferences::~CPreferences()
{
	// prevent anyone else changing the shared resources
	// for the duration of this function
	LOCKPREFS();
			
	s_nRef--; // decrement reference count
	
	// save ini?
	if (s_nRef == 0 && s_bIni)
	{
		SaveInternal();

		// cleanup
		int nSection = s_aIni.GetSize();
		
		while (nSection--)
			delete s_aIni[nSection];
		
		s_aIni.RemoveAll();
	}
}

BOOL CPreferences::Save()
{
	if (!s_bDirty)
		return TRUE; // nothing to do

	// prevent anyone else changing the shared resources
	// for the duration of this function
	LOCKPREFSRET(FALSE);

	return SaveInternal();
}

// THIS IS AN INTERNAL METHOD THAT ASSUMES CALLERS HAVE INITIALISED A LOCK ALREADY
BOOL CPreferences::SaveInternal()
{
	ASSERT(s_bLocked);

	if (!s_bLocked)
		return FALSE;

	if (!s_bDirty)
		return TRUE; // nothing to do

	// backup file first
	CString sIniPath(AfxGetApp()->m_pszProfileName);
	CString sBackupPath(sIniPath);
	
	FileMisc::ReplaceExtension(sBackupPath, _T("bak.ini"));
	FileMisc::CopyFile(sIniPath, sBackupPath, TRUE, TRUE);
	
	// write prefs
	CStdioFileEx file;
	
	if (!file.Open(sIniPath, CFile::modeWrite | CFile::modeCreate))
		return FALSE;
	
	// insert application version
	INISECTION* pSection = GetSection(_T("AppVer"), TRUE);
	ASSERT(pSection);
	
	if (pSection)
		SetEntryValue(*pSection, _T("Version"), FileMisc::GetAppVersion(), FALSE);
	
	for (int nSection = 0; nSection < s_aIni.GetSize(); nSection++)
	{
		// write section line
		INISECTION* pSection = s_aIni[nSection];
		
		CString sLine;
		sLine.Format(_T("[%s]%s"), pSection->sSection, ENDL);
		
		file.WriteString(sLine);
		
		// write entries to a CStringArray, then sort it and write it to file
		CStringArray aEntries;
		aEntries.SetSize(pSection->aEntries.GetCount(), 10);
		
		// save map to array
		POSITION pos = pSection->aEntries.GetStartPosition();
		int nEntry = 0;
		
		while (pos)
		{
			CString sDummy;
			INIENTRY ie;
			
			pSection->aEntries.GetNextAssoc(pos, sDummy, ie);

			aEntries.SetAt(nEntry++, ie.Format());
		}
		
		// sort array
		Misc::SortArray(aEntries);
		
		// format by newlines
		CString sSection = Misc::FormatArray(aEntries, ENDL);
		sSection += ENDL;
		sSection += ENDL;
		
		// save to file
		file.WriteString(sSection);
	}
	
	// Close/Flush the file BEFORE deleting the backup
	file.Close();
	::DeleteFile(sBackupPath);
	
	s_bDirty = FALSE;

	return TRUE;
}

CString CPreferences::ToString(int nValue)
{
	CString sValue;
	sValue.Format(_T("%ld"), nValue);
	return sValue;
}

CString CPreferences::ToString(double dValue)
{
	CString sValue;
	sValue.Format(_T("%.6f"), dValue);
	return sValue;
}

UINT CPreferences::GetProfileInt(LPCTSTR lpszSection, LPCTSTR lpszEntry, int nDefault) const
{
	if (s_bIni)
	{
		CString sValue = GetProfileString(lpszSection, lpszEntry);
		
		if (sValue.IsEmpty()) 
			return nDefault;
		else
			return _ttol(sValue);
	}
	else
	{
		return AfxGetApp()->GetProfileInt(lpszSection, lpszEntry, nDefault);
	}
}

BOOL CPreferences::WriteProfileInt(LPCTSTR lpszSection, LPCTSTR lpszEntry, int nValue)
{
	if (s_bIni)
		return WriteProfileString(lpszSection, lpszEntry, ToString(nValue), FALSE);
	else
		return AfxGetApp()->WriteProfileInt(lpszSection, lpszEntry, nValue);
}

LPCTSTR CPreferences::GetProfileString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszDefault) const
{
	if (s_bIni)
	{
		// prevent anyone else changing the shared resources
		// for the duration of this function
		LOCKPREFSRET(NULLSTR);

		return GetIniString(lpszSection, lpszEntry, lpszDefault);
	}

	// else registry
	return AfxGetApp()->GetProfileString(lpszSection, lpszEntry, lpszDefault);
}

BOOL CPreferences::WriteProfileString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszValue)
{
	// quote string values passed in from outside
	return WriteProfileString(lpszSection, lpszEntry, lpszValue, TRUE);
}

BOOL CPreferences::WriteProfileString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszValue, BOOL bQuoted)
{
	if (s_bIni)
	{
		// prevent anyone else changing the shared resources
		// for the duration of this function
		LOCKPREFSRET(FALSE);

		return WriteIniString(lpszSection, lpszEntry, lpszValue, bQuoted);
	}

	// else registry
	return AfxGetApp()->WriteProfileString(lpszSection, lpszEntry, lpszValue);
}

// THIS IS AN INTERNAL METHOD THAT ASSUMES CALLERS HAVE INITIALISED A LOCK ALREADY
LPCTSTR CPreferences::GetIniString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszDefault)
{
	ASSERT(s_bIni);
	ASSERT(s_bLocked);

	if (s_bIni && s_bLocked)
	{
		INISECTION* pSection = GetSection(lpszSection, FALSE);
		
		if (pSection)
		{
			CString sValue;
			
			if (GetEntryValue(*pSection, lpszEntry, sValue))
				return sValue;
		}
	
		// else
		return lpszDefault;
	}

	// else
	return NULLSTR;
}

// THIS IS AN INTERNAL METHOD THAT ASSUMES CALLERS HAVE INITIALISED A LOCK ALREADY
BOOL CPreferences::WriteIniString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszValue, BOOL bQuoted)
{
	ASSERT(s_bIni);
	ASSERT(s_bLocked);

	if (s_bIni && s_bLocked)
	{
		INISECTION* pSection = GetSection(lpszSection, TRUE);
		ASSERT(pSection);
		
		if (pSection)
		{
			SetEntryValue(*pSection, lpszEntry, lpszValue, bQuoted);
			return TRUE;
		}
	}
	
	// else
	return FALSE;
}

double CPreferences::GetProfileDouble(LPCTSTR lpszSection, LPCTSTR lpszEntry, double dDefault) const
{
	CString sValue = GetProfileString(lpszSection, lpszEntry, ToString(dDefault));
	
	if (sValue.IsEmpty())
		return dDefault;
	else
		return Misc::Atof(sValue);
}

BOOL CPreferences::WriteProfileDouble(LPCTSTR lpszSection, LPCTSTR lpszEntry, double dValue)
{
	return WriteProfileString(lpszSection, lpszEntry, ToString(dValue), FALSE);
}

int CPreferences::GetProfileArray(LPCTSTR lpszSection, CStringArray& aItems, BOOL bAllowEmpty) const
{
	aItems.RemoveAll();
	int nCount = GetProfileInt(lpszSection, _T("ItemCount"), 0);
	
	// items
	for (int nItem = 0; nItem < nCount; nItem++)
	{
		CString sItemKey, sItem;
		sItemKey.Format(_T("Item%d"), nItem);
		sItem = GetProfileString(lpszSection, sItemKey);
		
		if (bAllowEmpty || !sItem.IsEmpty())
			aItems.Add(sItem);
	}
	
	return aItems.GetSize();
}

void CPreferences::WriteProfileArray(LPCTSTR lpszSection, const CStringArray& aItems, BOOL bDelSection)
{
	// pre-delete?
	if (bDelSection)
		DeleteSection(lpszSection);

	int nCount = aItems.GetSize();
	
	// items
	for (int nItem = 0; nItem < nCount; nItem++)
	{
		CString sItemKey;
		sItemKey.Format(_T("Item%d"), nItem);
		WriteProfileString(lpszSection, sItemKey, aItems[nItem]);
	}
	
	// item count
	WriteProfileInt(lpszSection, _T("ItemCount"), nCount);
}

int CPreferences::GetProfileArray(LPCTSTR lpszSection, CDWordArray& aItems) const
{
	aItems.RemoveAll();
	int nCount = GetProfileInt(lpszSection, _T("ItemCount"), 0);
	
	// items
	for (int nItem = 0; nItem < nCount; nItem++)
	{
		CString sItemKey, sItem;

		sItemKey.Format(_T("Item%d"), nItem);
		aItems.Add(GetProfileInt(lpszSection, sItemKey));
	}
	
	return aItems.GetSize();
}

void CPreferences::WriteProfileArray(LPCTSTR lpszSection, const CDWordArray& aItems, BOOL bDelSection)
{
	// pre-delete?
	if (bDelSection)
		DeleteSection(lpszSection);

	int nCount = aItems.GetSize();
	
	// items
	for (int nItem = 0; nItem < nCount; nItem++)
	{
		CString sItemKey;
		sItemKey.Format(_T("Item%d"), nItem);
		WriteProfileInt(lpszSection, sItemKey, (int)aItems[nItem]);
	}
	
	// item count
	WriteProfileInt(lpszSection, _T("ItemCount"), nCount);
}

// THIS IS AN INTERNAL METHOD THAT ASSUMES CALLERS HAVE INITIALISED A LOCK ALREADY
BOOL CPreferences::GetEntryValue(const INISECTION& section, LPCTSTR lpszEntry, CString& sValue)
{
	ASSERT(s_bIni);
	ASSERT(s_bLocked);
	
	if (s_bIni && s_bLocked)
	{
		CString sKey(lpszEntry);
		sKey.MakeUpper();
		
		INIENTRY ie;
		
		if (section.aEntries.Lookup(sKey, ie))
		{
			sValue = ie.sValue;
			return TRUE;
		}
	}

	// else
	return FALSE;
}

void CPreferences::SetEntryValue(INISECTION& section, LPCTSTR lpszEntry, LPCTSTR szValue, BOOL bQuoted)
{
	SetEntryValue(section, INIENTRY(lpszEntry, szValue, bQuoted));
}

// THIS IS AN INTERNAL METHOD THAT ASSUMES CALLERS HAVE INITIALISED A LOCK ALREADY
void CPreferences::SetEntryValue(INISECTION& section, const INIENTRY& ie)
{
	ASSERT(s_bIni);
	ASSERT(s_bLocked);

	if (s_bIni && s_bLocked)
	{
		INIENTRY ieExist;
		CString sKey(ie.sName);
		sKey.MakeUpper();

		if (!section.aEntries.Lookup(sKey, ieExist) || !(ie == ieExist))
		{
			section.aEntries[sKey] = ie;
			s_bDirty = TRUE;
		}
	}
}

// THIS IS AN INTERNAL METHOD THAT ASSUMES CALLERS HAVE INITIALISED A LOCK ALREADY
INISECTION* CPreferences::GetSection(LPCTSTR lpszSection, BOOL bCreateNotExist)
{
	ASSERT(s_bIni);
	ASSERT(s_bLocked);

	if (s_bIni && s_bLocked)
	{
		int nSection = FindSection(lpszSection);

		if (nSection != -1)
			return s_aIni[nSection];
		
		// add a new section
		if (bCreateNotExist)
		{
			INISECTION* pSection = new INISECTION(lpszSection);
			
			s_aIni.Add(pSection);
			return pSection;
		}
	}
	
	// else
	return NULL;
}

// THIS IS AN INTERNAL METHOD THAT ASSUMES CALLERS HAVE INITIALISED A LOCK ALREADY
int CPreferences::FindSection(LPCTSTR lpszSection, BOOL bIncSubSections)
{
	ASSERT(s_bIni);
	ASSERT(s_bLocked);

	if (s_bIni && s_bLocked)
	{
		int nLenSection = lstrlen(lpszSection);
		int nSection = s_aIni.GetSize();
		
		while (nSection--)
		{
			const CString& sSection = s_aIni[nSection]->sSection;

			if (sSection.GetLength() < nLenSection)
				continue;
			
			else if (sSection.GetLength() == nLenSection)
			{
				if (sSection.CompareNoCase(lpszSection) == 0)
					return nSection;
			}
			else // sSection.GetLength() > nLenSection
			{
				if (bIncSubSections)
				{
					// look for parent section at head of subsection
					CString sTest = sSection.Left(nLenSection);

					if (sTest.CompareNoCase(lpszSection) == 0)
						return nSection;
				}
			}
		}
	}

	// not found
	return -1;
}

BOOL CPreferences::DeleteSection(LPCTSTR lpszSection, BOOL bIncSubSections)
{
	ASSERT(lpszSection && *lpszSection);

	if (s_bIni)
	{
		// prevent anyone else changing the shared resources
		// for the duration of this function
		LOCKPREFSRET(FALSE);
		
		int nSection = FindSection(lpszSection);
		 
		if (nSection != -1)
		{
			delete s_aIni[nSection];
			s_aIni.RemoveAt(nSection);
			s_bDirty = TRUE;
			
			if (bIncSubSections)
			{
				nSection = FindSection(lpszSection, TRUE);

				while (nSection != -1)
				{
					delete s_aIni[nSection];
					s_aIni.RemoveAt(nSection);
					nSection = FindSection(lpszSection, TRUE);
				} 
			}

			return TRUE;
		}

		// not found
		return FALSE;
	}

	// else registry
	CString sFullKey;
	sFullKey.Format(_T("%s\\%s"), AfxGetApp()->m_pszRegistryKey, lpszSection);

	return CRegKey2::DeleteKey(HKEY_CURRENT_USER, sFullKey);
}

BOOL CPreferences::DeleteEntry(LPCTSTR lpszSection, LPCTSTR lpszEntry)
{
	ASSERT(lpszSection && *lpszSection);
	ASSERT(lpszEntry && *lpszEntry);

	if (s_bIni)
	{
		// prevent anyone else changing the shared resources
		// for the duration of this function
		LOCKPREFSRET(FALSE);
		
		INISECTION* pSection = GetSection(lpszSection, FALSE);

		if (pSection)
		{
			CString sKey(lpszEntry);
			sKey.MakeUpper();

			pSection->aEntries.RemoveKey(sKey);
			s_bDirty = TRUE;

			return TRUE;
		}

		// not found
		return FALSE;
	}

	// else registry
	CString sFullKey;
	sFullKey.Format(_T("%s\\%s\\%s"), AfxGetApp()->m_pszRegistryKey, lpszSection, lpszEntry);

	return CRegKey2::DeleteKey(HKEY_CURRENT_USER, sFullKey);
}

BOOL CPreferences::HasSection(LPCTSTR lpszSection) const
{
	if (s_bIni)
	{
		// prevent anyone else changing the shared resources
		// for the duration of this function
		LOCKPREFSRET(FALSE);
		
		return (GetSection(lpszSection, FALSE) != NULL);
	}

	// else registry
	CString sFullKey;
	sFullKey.Format(_T("%s\\%s"), AfxGetApp()->m_pszRegistryKey, lpszSection);

	return CRegKey2::KeyExists(HKEY_CURRENT_USER, sFullKey);
}

int CPreferences::GetSectionNames(CStringArray& aSections)
{
	// prevent anyone else changing the shared resources
	// for the duration of this function
	LOCKPREFSRET(0);
	
	aSections.RemoveAll();

	if (s_bIni)
	{
		int nSection = s_aIni.GetSize();
		aSections.SetSize(nSection);
		
		while (nSection--)
			aSections[nSection] = s_aIni[nSection]->sSection;
	}

	return aSections.GetSize();
}

CString CPreferences::KeyFromFile(LPCTSTR szFilePath, BOOL bFileNameOnly)
{
	CString sKey = (bFileNameOnly ? FileMisc::GetFileNameFromPath(szFilePath) : szFilePath);
	sKey.Replace('\\', '_');

	// if the filepath is on a removable drive then we strip off the drive letter
	if (!bFileNameOnly)
	{
		int nType = CDriveInfo::GetPathType(szFilePath);

		if (nType == DRIVE_REMOVABLE)
			sKey = sKey.Mid(1);
	}

	return sKey;
}
