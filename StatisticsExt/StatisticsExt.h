// StatisticsExt.h: interface for the CStatisticsExtApp class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_StatisticsEXT_H__DEE73DE1_C6EC_4648_9151_0FC2C75A806D__INCLUDED_)
#define AFX_StatisticsEXT_H__DEE73DE1_C6EC_4648_9151_0FC2C75A806D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "..\Interfaces\IUIExtension.h"

const LPCTSTR STATS_TYPEID = _T("0AA35779-4BB4-4151-BA8E-D471281B6A08");

class CStatisticsExtApp : public IUIExtension  
{
public:
	CStatisticsExtApp();
	virtual ~CStatisticsExtApp();

    void Release(); // releases the interface
	void SetLocalizer(ITransText* pTT);

	LPCTSTR GetMenuText() const { return _T("Statistics"); }
	HICON GetIcon() const { return m_hIcon; }
	LPCTSTR GetTypeID() const { return STATS_TYPEID; }

	IUIExtensionWindow* CreateExtWindow(UINT nCtrlID, DWORD nStyle, 
										long nLeft, long nTop, long nWidth, long nHeight, 
										HWND hwndParent);

protected:
	HICON m_hIcon;
};

#endif // !defined(AFX_StatisticsEXT_H__DEE73DE1_C6EC_4648_9151_0FC2C75A806D__INCLUDED_)
