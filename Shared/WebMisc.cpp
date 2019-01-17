#include "stdafx.h"
#include "webmisc.h"
#include "misc.h"
#include "filemisc.h"
#include "xmlcharmap.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

#include <wininet.h>
#include <shlwapi.h>

#pragma comment(lib, "wininet.lib")
#pragma comment(lib, "shlwapi.lib")

///////////////////////////////////////////////////////////////////////////////////////////////////

// Missing from shlwapi.h
LWSTDAPI UrlCreateFromPathA(LPCSTR pszPath, LPSTR pszUrl, LPDWORD pcchUrl, DWORD dwFlags);
LWSTDAPI UrlCreateFromPathW(LPCWSTR pszPath, LPWSTR pszUrl, LPDWORD pcchUrl, DWORD dwFlags);
LWSTDAPI PathCreateFromUrlA(LPCSTR pszUrl, LPSTR pszPath, LPDWORD pcchPath, DWORD dwFlags);
LWSTDAPI PathCreateFromUrlW(LPCWSTR pszUrl, LPWSTR pszPath, LPDWORD pcchPath, DWORD dwFlags);

#ifdef UNICODE
#	define UrlCreateFromPath       UrlCreateFromPathW
#	define PathCreateFromUrl       PathCreateFromUrlW
#else //!UNICODE
#	define UrlCreateFromPath       UrlCreateFromPathA
#	define PathCreateFromUrl       PathCreateFromUrlA
#endif //UNICODE

///////////////////////////////////////////////////////////////////////////////////////////////////

const LPCTSTR FILEPROTOCOL = _T("file:///");

///////////////////////////////////////////////////////////////////////////////////////////////////

BOOL WebMisc::IsOnline()
{
    DWORD dwState = 0; 
    DWORD dwSize = sizeof(DWORD);
	
    return InternetQueryOption(NULL, INTERNET_OPTION_CONNECTED_STATE, &dwState, &dwSize) && 
		(dwState & INTERNET_STATE_CONNECTED);
}

BOOL WebMisc::DeleteCacheEntry(LPCTSTR szURI)
{
	BOOL bSuccess = FALSE;

#if _MSC_VER >= 1400
	bSuccess = DeleteUrlCacheEntry(szURI);
#elif _UNICODE
	LPSTR szAnsiPath = Misc::WideToMultiByte(szURI);
	bSuccess = DeleteUrlCacheEntry(szAnsiPath);
	delete [] szAnsiPath;
#else
	bSuccess = DeleteUrlCacheEntry(szURI);
#endif

	return bSuccess;
}

BOOL WebMisc::IsURL(LPCTSTR szURL)
{
	if (::PathIsURL(szURL))
		return TRUE;

	CString sURL(szURL);

	sURL.TrimLeft();
	sURL.MakeLower();

	return ((sURL.Find(_T("www.")) == 0) || (sURL.Find(_T("ftp.")) == 0));
}

int WebMisc::ExtractFirstHtmlLink(const CString& sHtml, CString& sLink, CString& sText)
{
	return ExtractNextHtmlLink(sHtml, 0, sLink, sText);
}

int WebMisc::ExtractNextHtmlLink(const CString& sHtml, int nFrom, CString& sLink, CString& sText)
{
	const CString LINK_START(_T("href=\"")), LINK_END(_T("\""));
	const CString TEXT_START(_T(">")), TEXT_END(_T("</a>"));

	int nStart = sHtml.Find(LINK_START, nFrom), nEnd = -1;

	if (nStart == -1)
		return -1;

	nStart += LINK_START.GetLength();
	nEnd = sHtml.Find(LINK_END, nStart);

	if (nEnd == -1)
		return -1;

	sLink = sHtml.Mid(nStart, (nEnd - nStart));

	// get text
	nStart = sHtml.Find(TEXT_START, nEnd + 1);

	if (nStart == -1)
		return -1;

	nStart += TEXT_START.GetLength();
	nEnd = sHtml.Find(TEXT_END, nStart);

	if (nEnd == -1)
		return -1;

	sText = sHtml.Mid(nStart, (nEnd - nStart));

	// cleanup
	Misc::Trim(sText, _T("\" "));

	// translate HTML represenations
	CXmlCharMap::ConvertFromRep(sText);

	return (nEnd + TEXT_END.GetLength());
}

int WebMisc::ExtractHtmlLinks(const CString& sHtml, CStringArray& aLinks, CStringArray& aLinkText)
{
	CString sLink, sText;
	int nFind = ExtractFirstHtmlLink(sHtml, sLink, sText);

	while (nFind != -1)
	{
		aLinks.Add(sLink);
		aLinkText.Add(sText);

		// next link
		nFind = ExtractNextHtmlLink(sHtml, nFind, sLink, sText);
	}
	ASSERT(aLinkText.GetSize() == aLinks.GetSize());

	return aLinks.GetSize();
}

BOOL WebMisc::IsAboutBlank(LPCTSTR szURL)
{
	CString sURL(szURL);
	sURL.MakeUpper();

	return (sURL.Find(_T("ABOUT:BLANK")) != -1);
}

BOOL WebMisc::FormatFileURI(LPCTSTR szFilePath, CString& sFileURI, BOOL bEncodeChars)
{
	if (!FileMisc::IsPath(szFilePath))
		return FALSE;

	if (bEncodeChars)
	{
		DWORD dwLen = MAX_PATH;
		HRESULT hr = ::UrlCreateFromPath(szFilePath, sFileURI.GetBuffer(MAX_PATH + 1), &dwLen, 0);
		sFileURI.ReleaseBuffer(dwLen);

		if (hr == S_OK)
			return TRUE;

		sFileURI.Empty();
		return FALSE;
	}

	// else just append file protocol and flip slashes
	sFileURI.Format(_T("%s%s"), FILEPROTOCOL, szFilePath);
	sFileURI.Replace('\\', '/');

	return TRUE;
}

BOOL WebMisc::DecodeFileURI(LPCTSTR szFileURI, CString& sFilePath)
{
	DWORD dwLen = MAX_PATH;
	HRESULT hr = ::PathCreateFromUrl(szFileURI, sFilePath.GetBuffer(MAX_PATH + 1), &dwLen, 0);
	sFilePath.ReleaseBuffer(dwLen);

	if (hr == S_OK)
		return TRUE;

	// else
	sFilePath.Empty();
	return FALSE;
}

BOOL WebMisc::IsFileURI(LPCTSTR szFilePath)
{
	if (_tcsstr(szFilePath, FILEPROTOCOL) != NULL)
		return TRUE;

	// shim for bad file protocol (one too few slashes)
	return (_tcsstr(szFilePath, _T("file://")) != NULL);
}
