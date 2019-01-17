// TDLContentMgr.cpp: implementation of the CTDLContentMgr class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TDLContentMgr.h"
#include "ToDoCommentsCtrl.h"

#include "..\shared\localizer.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

class CTDCDefaultContent : public IContent
{
	LPCTSTR GetTypeID() const { static LPCTSTR sID = _T("PLAIN_TEXT"); return sID; }
	LPCTSTR GetTypeDescription() const { static LPCTSTR sDesc = _T("Simple Text"); return sDesc; }

	IContentControl* CreateCtrl(unsigned short nCtrlID, unsigned long nStyle, 
						long nLeft, long nTop, long nWidth, long nHeight, HWND hwndParent)
	{
		CToDoCommentsCtrl* pComments = new CToDoCommentsCtrl;

		nStyle |= ES_MULTILINE | ES_WANTRETURN | WS_VSCROLL;
		CRect rect(nLeft, nTop, nLeft + nWidth, nTop + nHeight);

		if (pComments->Create(nStyle, rect, CWnd::FromHandle(hwndParent), nCtrlID))
			return pComments;

		// else
		delete pComments;
		return NULL;
	}

	void SetLocalizer(ITransText* pTT) { CLocalizer::Initialize(pTT); }
	void Release() { delete this; }

	int ConvertToHtml(const unsigned char* /*pContent*/, int /*nLength*/, LPCTSTR /*szCharset*/,
						LPTSTR& /*szHtml*/, LPCTSTR /*szImageDir*/) { return 0; } // not supported
	void FreeHtmlBuffer(LPTSTR& /*szHtml*/) {}
};

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CTDLContentMgr::CTDLContentMgr() 
{
}

CTDLContentMgr::~CTDLContentMgr()
{
}

void CTDLContentMgr::Initialize() const
{
	if (!m_bInitialized)
	{
		CContentMgr::Initialize();

		// we need a non-const pointer to update the array
		CTDLContentMgr* pMgr = const_cast<CTDLContentMgr*>(this);

		pMgr->m_aContent.InsertAt(0, new CTDCDefaultContent);
	}
}
