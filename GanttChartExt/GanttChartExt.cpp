// GanttChartExt.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include "GanttChartExt.h"
#include "GanttChartWnd.h"
#include "resource.h"

//#include <afxdllx.h>

#include "..\shared\misc.h"
#include "..\shared\localizer.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/*
static AFX_EXTENSION_MODULE GanttChartExtDLL = { NULL, NULL };

extern "C" int APIENTRY
DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// Remove this if you use lpReserved
	UNREFERENCED_PARAMETER(lpReserved);

	if (dwReason == DLL_PROCESS_ATTACH)
	{
		TRACE0("GANTTCHARTEXT.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(GanttChartExtDLL, hInstance))
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

		new CDynLinkLibrary(GanttChartExtDLL);
	}
	else if (dwReason == DLL_PROCESS_DETACH)
	{
		TRACE0("GANTTCHARTEXT.DLL Terminating!\n");
		// Terminate the library before destructors are called
		AfxTermExtensionModule(GanttChartExtDLL);
	}
	return 1;   // ok
}
*/

//////////////////////////////////////////////////////////////////////
// CGanttChartExtApp Class
//////////////////////////////////////////////////////////////////////

static CGanttChartExtApp theApp;

DLL_DECLSPEC IUIExtension* CreateUIExtensionInterface()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
//	return new CGanttChartExtApp;
	return &theApp;
}

DLL_DECLSPEC int GetInterfaceVersion() 
{ 
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
	return IUIEXTENSION_VERSION; 
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CGanttChartExtApp::CGanttChartExtApp() : m_hIcon(NULL)
{
}

CGanttChartExtApp::~CGanttChartExtApp()
{

}

void CGanttChartExtApp::Release()
{
// 	AFX_MANAGE_STATE(AfxGetStaticModuleState());
// 	
// 	delete this;
}

IUIExtensionWindow* CGanttChartExtApp::CreateExtWindow(UINT nCtrlID, DWORD nStyle, 
													long nLeft, long nTop, long nWidth, long nHeight, 
													HWND hwndParent)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
	CGanttChartWnd* pWindow = new CGanttChartWnd;

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

void CGanttChartExtApp::SetLocalizer(ITransText* pTT)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
	CLocalizer::Initialize(pTT);
}

BOOL CGanttChartExtApp::InitInstance()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
	m_hIcon = LoadIcon(IDR_GANTTCHART);

	return CWinApp::InitInstance();
}
