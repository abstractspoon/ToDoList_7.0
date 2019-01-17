// Preferences.h: interface for the CPreferences class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_PREFERENCES_H__DF763543_F9D5_4C94_BBD9_DF7E6E41B8C2__INCLUDED_)
#define AFX_PREFERENCES_H__DF763543_F9D5_4C94_BBD9_DF7E6E41B8C2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "..\Interfaces\IPreferences.h"

#include <afxtempl.h>
#include <afxmt.h>

struct INIENTRY
{
	INIENTRY(LPCTSTR szName = NULL, LPCTSTR szValue = NULL, BOOL bQuote = FALSE) 
		: sName(szName), sValue(szValue), bQuoted(bQuote) 
	{
	}

	CString Format() const
	{
		CString sEntry;

		if (bQuoted)
			sEntry.Format(_T("%s=\"%s\""), sName, sValue);
		else
			sEntry.Format(_T("%s=%s"), sName, sValue);

		return sEntry;
	}

	BOOL Parse(const CString& sEntry)
	{
		int nEquals = sEntry.Find('=');
		
		if (nEquals != -1)
		{
			sName = sEntry.Left(nEquals);
			sName.TrimRight();
			
			sValue = sEntry.Mid(nEquals + 1);
			sValue.TrimLeft();
			
			// remove quotes
			bQuoted = sValue.Replace(_T("\""), _T(""));

			return !sName.IsEmpty();
		}

		return FALSE;
	}

	BOOL operator==(const INIENTRY& ie) const
	{
		return ((sName == ie.sName) && 
				(sValue == ie.sValue) && 
				(bQuoted == ie.bQuoted));
	}
	
	CString sName;
	CString sValue;
	BOOL bQuoted;
};

typedef CMap<CString, LPCTSTR, INIENTRY, INIENTRY&> CIniEntryMap;

struct INISECTION
{
	INISECTION(LPCTSTR szName = NULL) : sSection(szName) 
	{
		aEntries.InitHashTable(199); // prime number closest to 200
	}
	
	CString sSection;
	CIniEntryMap aEntries;
};

typedef CArray<INISECTION*, INISECTION*> CIniSectionArray;

class CPreferences : public IPreferences
{
public:
	CPreferences(); 
	~CPreferences();

	BOOL UsesIni() const { return s_bIni; }

	// directly from CWinApp
	UINT GetProfileInt(LPCTSTR lpszSection, LPCTSTR lpszEntry, int nDefault = 0) const;
	BOOL WriteProfileInt(LPCTSTR lpszSection, LPCTSTR lpszEntry, int nValue);
	LPCTSTR GetProfileString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszDefault = NULL) const;
	BOOL WriteProfileString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszValue);

	// note: Binary not currently supported by ini file

	// extras
	double GetProfileDouble(LPCTSTR lpszSection, LPCTSTR lpszEntry, double dDefault = 0) const;
	BOOL WriteProfileDouble(LPCTSTR lpszSection, LPCTSTR lpszEntry, double dValue);
	
	int GetSectionNames(CStringArray& aSections);
	BOOL HasSection(LPCTSTR lpszSection) const;
	BOOL DeleteSection(LPCTSTR lpszSection, BOOL bIncSubSections = FALSE);
	BOOL DeleteEntry(LPCTSTR lpszSection, LPCTSTR lpszEntry);

	int GetProfileArray(LPCTSTR lpszSection, CStringArray& aItems, BOOL bAllowEmpty = FALSE) const;
	void WriteProfileArray(LPCTSTR lpszSection, const CStringArray& aItems, BOOL bDelSection = TRUE);
	int GetProfileArray(LPCTSTR lpszSection, CDWordArray& aItems) const;
	void WriteProfileArray(LPCTSTR lpszSection, const CDWordArray& aItems, BOOL bDelSection = TRUE);
	
	static CString KeyFromFile(LPCTSTR szFilePath, BOOL bFilenameOnly = TRUE);
	static BOOL Save();

protected:
	static CIniSectionArray s_aIni;
	static BOOL s_bDirty;
	static BOOL s_bIni;
	static int s_nRef;
	static CCriticalSection s_cs;
	static BOOL s_bLocked;
	
protected:
	BOOL WriteProfileString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszValue, BOOL bQuoted);

protected:
	// internal helpers that require pre-locking
	static BOOL SaveInternal();

	static LPCTSTR GetIniString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszDefault);
	static BOOL WriteIniString(LPCTSTR lpszSection, LPCTSTR lpszEntry, LPCTSTR lpszValue, BOOL bQuoted);
	
	static INISECTION* GetSection(LPCTSTR lpszSection, BOOL bCreateNotExist);
	static int FindSection(LPCTSTR lpszSection, BOOL bIncSubSections = FALSE);
	
	static POSITION GetEntry(INISECTION& section, LPCTSTR lpszEntry);
	static BOOL GetEntryValue(const INISECTION& section, LPCTSTR lpszEntry, CString& sValue);
	static void SetEntryValue(INISECTION& section, LPCTSTR lpszEntry, LPCTSTR szValue, BOOL bQuoted);
	static void SetEntryValue(INISECTION& section, const INIENTRY& ie);

	static CString ToString(int nValue);
	static CString ToString(double dValue);
};

#endif // !defined(AFX_PREFERENCES_H__DF763543_F9D5_4C94_BBD9_DF7E6E41B8C2__INCLUDED_)
