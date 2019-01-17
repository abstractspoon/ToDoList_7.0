// CalendarExt.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include "Resource.h"
#include "calendarext.h"
#include "calendarwnd.h"

#include <afxdllx.h>

#include "..\shared\misc.h"
#include "..\shared\localizer.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


static AFX_EXTENSION_MODULE CalendarExtDLL = { NULL, NULL };

extern "C" int APIENTRY
DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// Remove this if you use lpReserved
	UNREFERENCED_PARAMETER(lpReserved);

	if (dwReason == DLL_PROCESS_ATTACH)
	{
		TRACE0("CALENDAREXT.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(CalendarExtDLL, hInstance))
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

		new CDynLinkLibrary(CalendarExtDLL);
	}
	else if (dwReason == DLL_PROCESS_DETACH)
	{
		TRACE0("CALENDAREXT.DLL Terminating!\n");
		// Terminate the library before destructors are called
		AfxTermExtensionModule(CalendarExtDLL);
	}
	return 1;   // ok
}

DLL_DECLSPEC IUIExtension* CreateUIExtensionInterface()
{
	return new CCalendarExtApp;
}

DLL_DECLSPEC int GetInterfaceVersion() 
{ 
	return IUIEXTENSION_VERSION; 
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalendarExtApp::CCalendarExtApp() : m_hIcon(NULL)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_CALENDAR);
}

CCalendarExtApp::~CCalendarExtApp()
{

}

void CCalendarExtApp::Release()
{
	delete this;
}

IUIExtensionWindow* CCalendarExtApp::CreateExtWindow(UINT nCtrlID, DWORD nStyle, 
													long nLeft, long nTop, long nWidth, long nHeight, 
													HWND hwndParent)
{
	CCalendarWnd* pWindow = new CCalendarWnd;

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

void CCalendarExtApp::SetLocalizer(ITransText* pTT)
{
	CLocalizer::Initialize(pTT);
}
