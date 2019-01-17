#ifndef _MISCWEB_FUNCTIONS_H_
#define _MISCWEB_FUNCTIONS_H_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

///////////////////////////////////////////////////////////////////////////////////////////////////

// enum 
// {
// };

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace WebMisc
{
	BOOL IsOnline();
	BOOL DeleteCacheEntry(LPCTSTR szURI);
	BOOL IsURL(LPCTSTR szURL);
	BOOL IsAboutBlank(LPCTSTR szURL);

	int ExtractHtmlLinks(const CString& sHtml, CStringArray& aLinks, CStringArray& aLinkText);
	int ExtractFirstHtmlLink(const CString& sHtml, CString& sLink, CString& sText);
	int ExtractNextHtmlLink(const CString& sHtml, int nFrom, CString& sLink, CString& sText);

	BOOL FormatFileURI(LPCTSTR szFilePath, CString& sFileURI, BOOL bEncodeChars = TRUE);
	BOOL DecodeFileURI(LPCTSTR szFileURI, CString& sFilePath);
	BOOL IsFileURI(LPCTSTR szFilePath);
}

#endif // _MISCWEB_FUNCTIONS_H_