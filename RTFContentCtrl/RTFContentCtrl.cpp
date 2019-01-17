// RTFContentCtrl.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include "RTFContentCtrl.h"
#include "RTFContentControl.h"

#include "..\3rdparty\compression.h"
#include "..\3rdparty\stdiofileex.h"

#include "..\shared\filemisc.h"
#include "..\shared\misc.h"
#include "..\shared\mswordhelper.h"
#include "..\shared\localizer.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

////////////////////////////////////////////////////////////////////////////////////

static CRTFContentCtrlApp theApp;

////////////////////////////////////////////////////////////////////////////////////

DLL_DECLSPEC IContent* CreateContentInterface()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return &theApp;
}

DLL_DECLSPEC int GetInterfaceVersion() 
{ 
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return ICONTENTCTRL_VERSION; 
}

////////////////////////////////////////////////////////////////////////////////////

CRTFContentCtrlApp::CRTFContentCtrlApp()
{
}

void CRTFContentCtrlApp::SetLocalizer(ITransText* pTT)
{
	CLocalizer::Initialize(pTT);
}

IContentControl* CRTFContentCtrlApp::CreateCtrl(unsigned short nCtrlID, unsigned long nStyle, 
						long nLeft, long nTop, long nWidth, long nHeight, HWND hwndParent)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CRTFContentControl* pControl = new CRTFContentControl(m_rtfHtml);

	if (pControl)
	{
		CRect rect(nLeft, nTop, nLeft + nWidth, nTop + nHeight);

		if (pControl->Create(nStyle, rect, CWnd::FromHandle(hwndParent), nCtrlID, TRUE))
		{
			return pControl;
		}
	}

	delete pControl;
	return NULL;
}

void CRTFContentCtrlApp::Release()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	m_rtfHtml.Release();
}

int CRTFContentCtrlApp::ConvertToHtml(const unsigned char* pContent, int nLength, 
									  LPCTSTR szCharSet, LPTSTR& szHtml, LPCTSTR szImageDir)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (nLength == 0)
		return 0; // nothing to convert
	
	// we may have to decompress it first
	unsigned char* pDecompressed = NULL;
	
	if (!m_rtfHtml.IsRTF((const char*)pContent))
	{
		int nLenDecompressed = 0;
		
		if (Compression::Decompress(pContent, nLength, pDecompressed, nLenDecompressed))
		{
			pContent = pDecompressed;
			nLength = nLenDecompressed;

			// check again that it is RTF
			if (!m_rtfHtml.IsRTF((const char*)pContent))
			{
				nLength = 0;
			}
		}
		else
		{
			return 0;
		}
	}

	if (nLength)
	{
		CString sHtml;

		if (m_rtfHtml.ConvertRtfToHtml((const char*)pContent, szCharSet, sHtml, szImageDir))
		{
			nLength = sHtml.GetLength();
			ASSERT(nLength);

			szHtml = new TCHAR[nLength+1];
			lstrcpy(szHtml, sHtml);
			szHtml[nLength] = 0;
		}
	}


	// cleanup
	delete [] pDecompressed;

	return nLength;
}

void CRTFContentCtrlApp::FreeHtmlBuffer(LPTSTR& szHtml)
{
	delete [] szHtml;
	szHtml = NULL;
}
