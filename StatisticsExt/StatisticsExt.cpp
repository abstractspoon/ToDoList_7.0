// StatisticsExt.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include "resource.h"
#include "StatisticsExt.h"
#include "StatisticsWnd.h"

#include "..\shared\misc.h"
#include "..\shared\localizer.h"

#include "..\Interfaces\iuiextension.h"

#include <afxdllx.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


static AFX_EXTENSION_MODULE StatisticsExtDLL = { NULL, NULL };

extern "C" int APIENTRY
DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// Remove this if you use lpReserved
	UNREFERENCED_PARAMETER(lpReserved);

	if (dwReason == DLL_PROCESS_ATTACH)
	{
		TRACE0("STATISTICSEXT.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(StatisticsExtDLL, hInstance))
			return 0;

		// Insert this DLL into the resource chain
		// NOTE: If this Extension DLL is being implicitly linked to by
		//  an MFC Regular DLL (such as an ActiveX Control)
		//  instead of an MFC application, then you will want to
		//  remove this line from DllMain and put it in a separate
		//  function exported from this Extension DLL.  The Regular DLL
		//  that uses this Extension DLL should then explicitly call that
		//  function to initialize this Extension DLL.  Otherwise,
		//  the CDynLinkLibrary object will not be attached to the
		//  Regular DLL's resource chain, and serious problems will
		//  result.

		new CDynLinkLibrary(StatisticsExtDLL);
	}
	else if (dwReason == DLL_PROCESS_DETACH)
	{
		TRACE0("STATISTICSEXT.DLL Terminating!\n");
		// Terminate the library before destructors are called
		AfxTermExtensionModule(StatisticsExtDLL);
	}
	return 1;   // ok
}

//////////////////////////////////////////////////////////////////////
// CStatisticsExtApp Class
//////////////////////////////////////////////////////////////////////

DLL_DECLSPEC IUIExtension* CreateUIExtensionInterface()
{
	return new CStatisticsExtApp;
}

DLL_DECLSPEC int GetInterfaceVersion() 
{ 
	return IUIEXTENSION_VERSION; 
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CStatisticsExtApp::CStatisticsExtApp() : m_hIcon(NULL)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_STATISTICS);
}

CStatisticsExtApp::~CStatisticsExtApp()
{

}

void CStatisticsExtApp::Release()
{
	delete this;
}

IUIExtensionWindow* CStatisticsExtApp::CreateExtWindow(UINT nCtrlID, DWORD nStyle, 
													long nLeft, long nTop, long nWidth, long nHeight, 
													HWND hwndParent)
{
	CStatisticsWnd* pWindow = new CStatisticsWnd;

	if (pWindow)
	{
		CRect rect(nLeft, nTop, nLeft + nWidth, nTop + nHeight);

		if (pWindow->Create(nStyle, rect, CWnd::FromHandle(hwndParent), nCtrlID))
		{
			return pWindow;
		}
	}

	delete pWindow;
	return NULL;
}

void CStatisticsExtApp::SetLocalizer(ITransText* pTT)
{
	CLocalizer::Initialize(pTT);
}

