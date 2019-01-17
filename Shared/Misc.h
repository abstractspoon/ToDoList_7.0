// Misc.h: interface for the CMisc class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_MISC_H__4B2FDA3E_63C5_4F52_A139_9512105C3AD4__INCLUDED_)
#define AFX_MISC_H__4B2FDA3E_63C5_4F52_A139_9512105C3AD4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <afxtempl.h>

enum 
{ 
	MKS_NONE = 0x00,
	MKS_CTRL = 0x01, 
	MKS_SHIFT = 0x02, 
	MKS_ALT = 0x04, 
};

enum 
{ 
	MKC_LEFTRIGHT = 0x01, 
	MKC_UPDOWN = 0x02, 
	MKC_ANY = (MKC_LEFTRIGHT | MKC_UPDOWN) 
};

#ifndef _ttof
#define _ttof(str) Misc::Atof(str)
#endif

namespace Misc  
{
	CString FormatGetLastError(DWORD dwLastErr = -1);

	BOOL CopyTexttoClipboard(const CString& sText, HWND hwnd, UINT nFormat = 0, BOOL bClear = TRUE); 
	CString GetClipboardText(UINT nFormat = 0); 
	BOOL ClipboardHasFormat(UINT nFormat);
	BOOL ClipboardHasText();

	char* WideToMultiByte(const WCHAR* szFrom, UINT nCodePage = CP_ACP);
	char* WideToMultiByte(const WCHAR* szFrom, int& nLength, UINT nCodePage = CP_ACP);
	WCHAR* MultiByteToWide(const char* szFrom, UINT nCodepage = CP_ACP);
	WCHAR* MultiByteToWide(const char* szFrom, int& nLength, UINT nCodepage = CP_ACP);
	CString& EncodeAsMultiByte(CString& sUnicode, UINT nCodePage = CP_ACP);
	CString& EncodeAsUnicode(CString& sMultibyte, UINT nCodePage = CP_ACP);

	BOOL GuidFromString(LPCTSTR szGuid, GUID& guid);
	BOOL IsGuid(LPCTSTR szGuid);
	BOOL GuidToString(const GUID& guid, CString& sGuid);
	BOOL GuidIsNull(const GUID& guid);
	void NullGuid(GUID& guid);
	BOOL SameGuids(const GUID& guid1, const GUID& guid2);
	CString NewGuid(GUID* pGuid = NULL);

	template <class T> 
	BOOL MatchAllT(const T& array1, const T& array2, BOOL bOrderSensitive = FALSE)
	{
		int nItem1 = array1.GetSize();
		int nItem2 = array2.GetSize();
		
		if (nItem1 != nItem2)
			return 0;
		
		if (bOrderSensitive)
		{
			while (nItem1--)
			{
				// check for non-equality
				if (!(array1[nItem1] == array2[nItem1]))
					return FALSE;
			}
			
			return TRUE;
		}
		
		// else order not important
		while (nItem1--)
		{
			// look for matching item
			nItem2 = array2.GetSize();

			while (nItem2--)
			{
				if (array1[nItem1] == array2[nItem2])
					break;
			}
			
			// no-match found == not the same
			if (nItem2 == -1)
				return FALSE;
		}
		
		return TRUE;
	}

	CString FormatComputerNameAndUser(char cSeparator = ':');
	CString GetComputerName();
	CString GetUserName();
	CString GetListSeparator();
	CString GetDecimalSeparator();
	CString GetDefCharset();
	CString GetAM();
	CString GetPM();
	CString GetTimeSeparator();
	CString GetTimeFormat(BOOL bIncSeconds = TRUE);
	CString GetShortDateFormat(BOOL bIncDOW = FALSE);
	CString GetDateSeparator();

	BOOL MatchAll(const CStringArray& array1, const CStringArray& array2, 
					 BOOL bOrderSensitive = FALSE, BOOL bCaseSensitive = FALSE);
	BOOL MatchAny(const CStringArray& array1, const CStringArray& array2, 
					BOOL bCaseSensitive = FALSE, BOOL bPartialOK = TRUE);

	BOOL MatchAll(const CDWordArray& array1, const CDWordArray& array2, 
		BOOL bOrderSensitive = FALSE);
	BOOL MatchAny(const CDWordArray& array1, const CDWordArray& array2);
	
	void Trace(const CStringArray& array);
	int RemoveItems(const CStringArray& aItems, CStringArray& aFrom, BOOL bCaseSensitive = FALSE);
	int RemoveEmptyItems(CStringArray& aFrom);
	BOOL RemoveItem(LPCTSTR szItem, CStringArray& aFrom, BOOL bCaseSensitive = FALSE);
	int AddUniqueItems(const CStringArray& aItems, CStringArray& aTo, BOOL bCaseSensitive = FALSE);
	int AddUniqueItem(const CString& sItem, CStringArray& aTo, BOOL bCaseSensitive = FALSE);
	int GetTotalLength(const CStringArray& array);

	const CString& GetItem(const CStringArray& array, int nItem);

	template <class T, class S> 
	const S& GetItemT(const T& array, int nItem)
	{
		ASSERT(nItem >= 0 && nItem < array.GetSize());
		
		if (nItem < 0 || nItem >= array.GetSize())
		{
			static S dummy;
			return dummy;
		}
		
		return array.GetData()[nItem];
	}

	CString FormatArray(const CDWordArray& array, LPCTSTR szSep = NULL);
	CString FormatArray(const CDWordArray& array, TCHAR cSep);
	CString FormatArray(const CStringArray& array, LPCTSTR szSep = NULL);
	CString FormatArray(const CStringArray& array, TCHAR cSep);
	CString FormatArrayAsNumberedList(const CStringArray& array, LPCTSTR szDelim = _T(". "), int nStart = 1);

	int Split(const CString& sText, CStringArray& aValues, LPCTSTR szSep = _T(""), BOOL bAllowEmpty = FALSE);
 	int Split(const CString& sText, CStringArray& aValues, TCHAR cDelim, BOOL bAllowEmpty = FALSE);
 	BOOL Split(CString& sText, CString& sRest, TCHAR cDelim, BOOL bTrim = TRUE);

	template <class T, class S> 
	int FindT(const T& array, const S& toFind)
	{
		int nItem = array.GetSize();
		const S* pData = array.GetData();

		while (nItem--)
		{
			if (pData[nItem] == toFind)
				return nItem;
		}

		// not found
		return -1;
	}
	
	int Find(const CStringArray& array, LPCTSTR szItem, 
		BOOL bCaseSensitive = FALSE, BOOL bPartialOK = TRUE);
	
	typedef int (*STRINGSORTPROC)(const void* pStr1, const void* pStr2);
	void SortArray(CStringArray& array, STRINGSORTPROC pSortProc = NULL);

	typedef int (*DWORDSORTPROC)(const void* pDW1, const void* pDW2);
	void SortArray(CDWordArray& array, DWORDSORTPROC pSortProc = NULL);

	void MakeUpper(CString& sText);
	void MakeLower(CString& sText);
	CString ToUpper(LPCTSTR szText);
	CString ToLower(LPCTSTR szText);
	CString ToUpper(TCHAR cText);
	CString ToLower(TCHAR cText);
	int NaturalCompare(LPCTSTR szString1, LPCTSTR szString2);
	BOOL LCMapString(CString& sText, DWORD dwMapFlags);

	TCHAR First(const CString& sText);
	TCHAR Last(const CString& sText);
	TCHAR TrimFirst(CString& sText);
	TCHAR TrimLast(CString& sText);
	CString& Trim(CString& sText, LPCTSTR lpszTargets = NULL);

	double Round(double dValue);
	float Round(float fValue);
	double Atof(const CString& sValue);
	CString Format(double dVal, int nDecPlaces = -1, LPCTSTR szTrail = _T(""));
	CString Format(int nVal, LPCTSTR szTrail = _T(""));
	CString Format(DWORD dwVal, LPCTSTR szTrail = _T(""));
	CString FormatCost(double dCost);
	BOOL IsNumber(const CString& sValue);
	BOOL IsSymbol(const CString& sValue);

	const CString& GetLongest(const CString& str1, const CString& str2, BOOL bAsDouble = FALSE);

	BOOL IsWorkStationLocked();
	BOOL IsScreenSaverActive();
	BOOL IsScreenReaderActive(BOOL bCheckForMSNarrator = TRUE);
	BOOL IsMSNarratorActive();
	LANGID GetUserDefaultUILanguage();
	BOOL IsMetricMeasurementSystem();
	BOOL IsHighContrastActive();

	BOOL ShutdownBlockReasonCreate(HWND hWnd, LPCTSTR szReason);
	BOOL ShutdownBlockReasonDestroy(HWND hWnd);

	void ProcessMsgLoop();
	DWORD GetLastUserInputTick();
	DWORD GetTicksSinceLastUserInput();

	int ParseSearchString(LPCTSTR szLookFor, CStringArray& aWords);
	BOOL FindWord(LPCTSTR szWord, LPCTSTR szText, BOOL bCaseSensitive, BOOL bMatchWholeWord);
	int FilterString(CString& sText, const CString& sFilter);

	BOOL ModKeysArePressed(DWORD dwKeys); 
	BOOL IsKeyPressed(DWORD dwVirtKey);
	BOOL IsCursorKey(DWORD dwVirtKey, DWORD dwKeys = MKC_ANY);
	BOOL IsCursorKeyPressed(DWORD dwKeys = MKC_ANY);
	CString GetKeyName(WORD wVirtKey, BOOL bExtended = FALSE); 

	inline BOOL HasFlag(DWORD dwFlags, DWORD dwFlag) { return (((dwFlags & dwFlag) == dwFlag) ? TRUE : FALSE); }
	BOOL ModifyFlags(DWORD& dwFlags, DWORD dwRemove, DWORD dwAdd = 0);

	CString MakeKey(const CString& sFormat, int nKeyVal, LPCTSTR szParentKey = NULL);
	CString MakeKey(const CString& sFormat, LPCTSTR szKeyVal, LPCTSTR szParentKey = NULL);
};

#endif // !defined(AFX_MISC_H__4B2FDA3E_63C5_4F52_A139_9512105C3AD4__INCLUDED_)
