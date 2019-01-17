// TransTextMgr.cpp: implementation of the CTransTextMgr class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TransTextMgr.h"
#include "subclass.h"
#include "xmlfile.h"
#include "winclasses.h"
#include "filemisc.h"
#include "misc.h"
#include "enstring.h"
#include "dialoghelper.h"

#include "..\3rdparty\StdioFileEx.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////

#ifdef _DEBUG
	static UINT nDICreated = 0;
	static UINT nDIDestroyed = 0;
#endif

const UINT POPUPMENU_ID = 0xFFFFFFFF;

static const CString TRANSTEXT_HEADER	= _T("TRANSTEXT");
static const CString TEXTOUT			= _T("TEXTOUT");
static const CString TEXTIN				= _T("TEXTIN");
static const CString ALTERNATIVE		= _T("ALTERNATIVE");
static const CString CLASSID			= _T("CLASSID");
static const CString TRANSLATED			= _T("TRANSLATED");
static const CString DICTVER			= _T("VERSION");
static const CString NEED_TRANSLATION	= _T("NEED_TRANSLATION");
static const CString ITEM				= _T("ITEM");
static const CString CSVCOLUMN_HEADER	= _T("English Text\tTranslated Text\tItem Type");
static const CString PRIMARY_LANGID		= _T("PRIMARY_LANGID");
static const CString STATIC_ID			= _T("65535");
static const CString SPACECHAR			= _T(" ");
// static CString _T("");

#ifndef _countof
#define _countof(array) (sizeof(array)/sizeof(array[0]))
#endif

//////////////////////////////////////////////////////////////////////

// local helpers
namespace TransText
{
	CString& DecodeChars(CString& sText)
	{
		if (!sText.IsEmpty())
		{
			if (sText.Find(_T("\\t")) >= 0)
				sText.Replace(_T("\\t"), _T("\t"));
			
			if (sText.Find(_T("\\r")) >= 0)
				sText.Replace(_T("\\r"), _T("\r"));
			
			if (sText.Find(_T("\\n")) >= 0)
				sText.Replace(_T("\\n"), _T("\n"));
		}
		
		return sText;
	}
	
	CString& EncodeChars(CString& sText)
	{
		if (!sText.IsEmpty())
		{
			if (sText.Find('\t') >= 0)
				sText.Replace(_T("\t"), _T("\\t"));
			
			if (sText.Find('\r') >= 0)
				sText.Replace(_T("\r"), _T("\\r"));
			
			if (sText.Find('\n') >= 0)
				sText.Replace(_T("\n"), _T("\\n"));
		}
		
		return sText;
	}

	CString GetFriendlyClass(const CString& sClass, HWND hWndRef = NULL) 
	{
		if (CWinClasses::IsClass(sClass, WC_DIALOGBOX))
		{
			return _T("dialog");
		}	
		else if (CWinClasses::IsClass(sClass, WC_STATIC))
		{
			return _T("label");
		}
		else if (CWinClasses::IsClass(sClass, WC_MENU))
		{
			return _T("menu");
		}
		else if (CWinClasses::IsClass(sClass, WC_TOOLTIPS))
		{
			return _T("tooltip");
		}
		else if (CWinClasses::IsClass(sClass, WC_BUTTON) && hWndRef)
		{
			int nBtnType = CWinClasses::GetButtonType(hWndRef);
			
			switch (nBtnType)
			{
			case BS_CHECKBOX:         
			case BS_AUTOCHECKBOX:  
			case BS_3STATE:           
			case BS_AUTO3STATE:  
				return _T("checkbox");
				
			case BS_RADIOBUTTON:      
			case BS_AUTORADIOBUTTON:
				return _T("radiobutton");
				
			case BS_GROUPBOX:      
				return _T("groupbox");
			}
			// drop thru
		}
		
		return sClass;
	}

	CString GetClassIDName(HWND hWnd)
	{
		if (hWnd == NULL)
			return _T("text");
		
		CString sClass = GetFriendlyClass(CWinClasses::GetClassEx(hWnd), hWnd);
		
		// do we have a 'valid' id
		int nCtrlID = ::GetDlgCtrlID(hWnd);
		
		if (nCtrlID <= 0 || nCtrlID >= 0xffff)
			return sClass;
		
		// else
		CString sName;
		sName.Format(_T("%s.%d"), sClass, nCtrlID);
		
		return sName;
	}
	
	CString GetClassIDName(HMENU /*hMenu*/, int nMenuID)
	{
		CString sClass = GetFriendlyClass(WC_MENU);
		
		if (nMenuID < 0)
			return sClass;
		
		CString sName;
		sName.Format(_T("%s.%d"), sClass, nMenuID);	
		
		return sName;
	}
	
	BOOL PrepareLookupText(CString& sText)
	{
		// remove trailing/leading spaces
		Misc::Trim(sText, SPACECHAR).TrimRight(':');
		
		// remove accelerators
		if (sText.Find('&') >= 0)
			sText.Remove('&');
		
		// check for numbers and symbols
		if (Misc::IsNumber(sText) || Misc::IsSymbol(sText))
		{
			CTransTextMgr::IgnoreString(sText, FALSE);
			return FALSE;
		}
		
		// finally decode 'tricky' characters like tabs and newlines
		DecodeChars(sText);
		
		return !sText.IsEmpty();
	}
	
	BOOL IsOwnerDraw(int nCmdID, HMENU hMenu)
	{
		if (nCmdID > 0)
		{
			MENUITEMINFO mii;
			ZeroMemory(&mii, sizeof(mii));
			
			mii.cbSize = sizeof(mii);
			mii.fMask = MIIM_TYPE;
			
			if (GetMenuItemInfo(hMenu, nCmdID, FALSE, &mii))
				return (mii.fType & MFT_OWNERDRAW);
		}
		
		return FALSE;
	}
	
	BOOL IsPopup(HWND hWnd)
	{
		DWORD dwStyle = GetWindowLong(hWnd, GWL_STYLE);
		DWORD dwMask = (WS_POPUP | WS_CAPTION);
		
		return (dwStyle & dwMask);
	}
}

//////////////////////////////////////////////////////////////////////
// private helper class

enum
{
	TWS_HANDLENONE		= 0x00,

	//TWS_HANDLESETTEXT	= 0x01,
	TWS_HANDLETOOLTIPS	= 0x02,
	TWS_HANDLEMENUPOPUP	= 0x04,

	// more here

	TWS_HANDLEALL		= 0xff,
	//TWS_NOHANDLESETTEXT	= (TWS_HANDLEALL & ~TWS_HANDLESETTEXT),
};

class CTransWnd : public CSubclassWnd
{
public:
	static CTransWnd* NewTransWnd(const CString& sClass, DWORD dwStyle);

	CTransWnd(DWORD dwOptions = TWS_HANDLEALL);
	virtual ~CTransWnd();
	
	void UpdateMenu() { TranslateMenu(::GetMenu(GetHwnd()), FALSE); }
	void AllowTranslation(BOOL bAllow) { m_bAllowTranslate = bAllow; }
	void PostHookWindow();
	
protected:
	DWORD m_dwOptions;
	BOOL m_bInit;
	BOOL m_bAllowTranslate;

protected:
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);
	virtual BOOL HookWindow(HWND hRealWnd, CSubclasser* pSubclasser);
	virtual void Initialize();
	
	BOOL TranslateText(CString& sText);
	void TranslateMenu(HMENU hMenu, BOOL bToplevelOnly);
	BOOL HasFlag(DWORD dwFlag) { return Misc::HasFlag(m_dwOptions, dwFlag); }
};

//////////////////////////////////////////////////////////////////////

class CTransComboBox : public CTransWnd
{
public:
	CTransComboBox() : CTransWnd(TWS_HANDLENONE) {}

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);

	// needed by Initialize()
	struct CBITEMDATA
	{
		CString sText;
		DWORD dwData;
	};
	typedef CArray<CBITEMDATA, CBITEMDATA&> CCbItemArray;
};

class CTransComboBoxEx : public CTransWnd
{
public:
	CTransComboBoxEx() : CTransWnd(TWS_HANDLENONE) {}

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);
};

class CTransListBox : public CTransWnd
{
public:
	CTransListBox(BOOL bCheckListBox = FALSE) : CTransWnd(TWS_HANDLENONE), m_bCheckLB(bCheckListBox) { }

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);

protected:
	BOOL m_bCheckLB;

	// needed by Initialize()
	struct LBITEMDATA
	{
		CString sText;
		DWORD dwData;
		int nCheck;
	};
	typedef CArray<LBITEMDATA, LBITEMDATA&> CLbItemArray;
};

class CTransTabCtrl : public CTransWnd
{
public:
	CTransTabCtrl() : CTransWnd(TWS_HANDLENONE) {}

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);
};

class CTransHeaderCtrl : public CTransWnd
{
public:
	CTransHeaderCtrl() : CTransWnd(TWS_HANDLENONE) {}

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);
};

class CTransListCtrl : public CTransWnd
{
public:
	CTransListCtrl() : CTransWnd(TWS_HANDLENONE) {}

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);
};

class CTransTooltips : public CTransWnd
{
public:
	CTransTooltips() : CTransWnd(TWS_HANDLENONE) {}

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);
};

class CTransToolBar : public CTransWnd
{
public:
	CTransToolBar() : CTransWnd(TWS_HANDLETOOLTIPS) {}

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);
};

class CTransStatusBar : public CTransWnd
{
public:
	CTransStatusBar() : CTransWnd(TWS_HANDLETOOLTIPS) {}

protected:
	virtual void Initialize();
	virtual LRESULT WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp);
};

//////////////////////////////////////////////////////////////////////
// Construction/Destruction

ITT_TRANSLATEOPTION DICTITEM::s_nTranslationOption = ITTTO_TRANSLATEONLY;

void DICTITEM::SetTranslationOption(ITT_TRANSLATEOPTION nOption)
{
	s_nTranslationOption = nOption;
}

BOOL DICTITEM::WantAddToDictionary()
{
	return (s_nTranslationOption == ITTTO_ADD2DICTIONARY);
}

BOOL DICTITEM::WantTranslateOnly()
{
	return (s_nTranslationOption == ITTTO_TRANSLATEONLY);
}

BOOL DICTITEM::WantUppercase()
{
	return (s_nTranslationOption == ITTTO_UPPERCASE);
}

// --------------------------------

DICTITEM::DICTITEM()
{
#ifdef _DEBUG
	nDICreated++;
#endif
}

DICTITEM::DICTITEM(const DICTITEM& di)
{
	*this = di;

#ifdef _DEBUG
	nDICreated++;
#endif
}

DICTITEM::DICTITEM(const CXmlItem* pXI)
{
	FromXml(pXI);

#ifdef _DEBUG
	nDICreated++;
#endif
}

DICTITEM::DICTITEM(const CString& sText) : m_sTextIn(sText)
{
#ifdef _DEBUG
	nDICreated++;
#endif
}

DICTITEM::~DICTITEM()
{
#ifdef _DEBUG
	nDIDestroyed++;
#endif
}

DICTITEM& DICTITEM::operator= (const DICTITEM& di)
{
	m_sTextIn = di.m_sTextIn;
	m_sTextOut = di.m_sTextOut;
	m_sClassID = di.m_sClassID;

	m_mapAlternatives.RemoveAll();

	if (di.m_mapAlternatives.GetCount())
	{
		POSITION pos = di.m_mapAlternatives.GetStartPosition();

		while (pos)
		{
			CString sKey, sText;
			di.m_mapAlternatives.GetNextAssoc(pos, sKey, sText);

			m_mapAlternatives[sKey] = sText;
		}
	}

	return *this;
}

// --------------------------------

BOOL DICTITEM::IsTranslated() const 
{ 
	return !m_sTextOut.IsEmpty(); 
}

int DICTITEM::GetDlgCtrlID(HWND hWnd)
{
	int nCtrlID = ::GetDlgCtrlID(hWnd);

	return max(nCtrlID, -1);
}

void DICTITEM::ToXml(CXmlItem* pXI, const DICTITEM& di) 
{
	ASSERT(!di.m_sTextIn.IsEmpty());

	pXI->SetItemValue(TEXTIN, di.m_sTextIn);
	pXI->SetItemValue(TEXTOUT, di.m_sTextOut);
	pXI->SetItemValue(CLASSID, di.m_sClassID);
}

void DICTITEM::FixupFormatString(CString& sFormat)
{
	// check that the char following the % is valid
	static const CString FMTCHARS = _T("cdiouxXeEfgGnps.*012346789 \t\n%");
	int nFind = sFormat.Find('%', 0);

	if (nFind != -1)
	{
		int nLen = sFormat.GetLength();

		while (nFind != -1 && ((nFind + 1) < nLen))
		{
			TCHAR cNext = sFormat[nFind + 1];
			
			if (FMTCHARS.Find(cNext) == -1) // invalid char
			{
				// try flipping the char
				if (_istlower(cNext))
					cNext = _totupper(cNext);
				else
					cNext = _totlower(cNext);

				sFormat.SetAt(nFind + 1, cNext);
				ASSERT(FMTCHARS.Find(cNext) != -1);
			}

			// next
			nFind = sFormat.Find('%', nFind + 1);
		}
	}
}

BOOL DICTITEM::ToXml(CXmlItem* pXI) const
{
	if (pXI && pXI->NameIs(ITEM))
	{
		ToXml(pXI, *this);
		
		// alternatives
		if (m_mapAlternatives.GetCount())
		{
			CString sClassID, sAlternative;
			POSITION pos = m_mapAlternatives.GetStartPosition();

			while (pos)
			{
				m_mapAlternatives.GetNextAssoc(pos, sClassID, sAlternative);
				ASSERT (!sClassID.IsEmpty());

				if (!sClassID.IsEmpty())
				{
					CXmlItem* pXISub = pXI->AddItem(ALTERNATIVE);

					pXISub->SetItemValue(CLASSID, sClassID);
					pXISub->SetItemValue(TEXTOUT, sAlternative);
				}
			}
		}

		return TRUE;
	}

	return FALSE;
}

BOOL DICTITEM::FromXml(const CXmlItem* pXI)
{
	if (pXI && pXI->NameIs(ITEM))
	{
		m_sTextIn = pXI->GetItemValue(TEXTIN);
		m_sTextOut = pXI->GetItemValue(TEXTOUT);
		m_sClassID = pXI->GetItemValue(CLASSID);

		// mark text out as being not-translatable
		// else the translated text can itself be translated!
		if (!m_sTextOut.IsEmpty())
			CTransTextMgr::IgnoreString(m_sTextOut, TRUE);

		//  alternatives
		const CXmlItem* pXISub = pXI->GetItem(ALTERNATIVE);

		while (pXISub)
		{
			CString sClassID = pXISub->GetItemValue(CLASSID);
			CString sAlternative = pXISub->GetItemValue(TEXTOUT);

			if (!sClassID.IsEmpty())
			{
				m_mapAlternatives[sClassID] = sAlternative;

				// mark text out as being not-translatable
				// else the translated text can itself be translated!
				if (!sAlternative.IsEmpty())
					CTransTextMgr::IgnoreString(sAlternative, TRUE);
			}

			// next
			pXISub = pXISub->GetSibling();
		}

		return TRUE;
	}

	return FALSE;
}

BOOL DICTITEM::ToCsv(CStringArray& aTransLines, CStringArray& aNeedTransLines) const
{
	aTransLines.RemoveAll();
	aNeedTransLines.RemoveAll();

	if (!m_sTextIn.IsEmpty())
	{
		// replace certain chars in text else they'll trip up the dictionary when it's read back in
		CString sTextIn(m_sTextIn), sTextOut(m_sTextOut);

		TransText::EncodeChars(sTextIn);
		TransText::EncodeChars(sTextOut);
			
		CString sLine;
		sLine.Format(_T("\"%s\"\t\"%s\"\t\"%s\""), sTextIn, sTextOut, m_sClassID);

		if (m_sTextOut.IsEmpty())
			aNeedTransLines.Add(sLine);
		else
			aTransLines.Add(sLine);

		// add alternatives, indented 
		POSITION pos = m_mapAlternatives.GetStartPosition();

		while (pos)
		{
			CString sAltTextOut, sAltClassID;
			m_mapAlternatives.GetNextAssoc(pos, sAltClassID, sAltTextOut);

			if (!sAltClassID.IsEmpty())
			{
				TransText::EncodeChars(sAltTextOut);

				sLine.Format(_T("  \"%s\"\t\"%s\"\t\"%s\""), m_sTextIn, sAltTextOut, sAltClassID);

				if (sAltTextOut.IsEmpty())
					aNeedTransLines.Add(sLine);
				else
					aTransLines.Add(sLine);
			}
		}
	}

	return (aTransLines.GetSize() > 0 || aNeedTransLines.GetSize() > 0);
}

BOOL DICTITEM::FromCsv(const CStringArray& aLines, int& nLine)
{
	const CString& sLine = Misc::GetItem(aLines, nLine);

	if (sLine.Find(NEED_TRANSLATION) == 0 || sLine.Find(TRANSLATED) == 0)
		return FALSE;

	if (FromCsv(sLine, *this))
	{
		// mark text out as being not-translatable
		// else the translated text can itself be translated!
		CTransTextMgr::IgnoreString(m_sTextOut, TRUE);

		// check for alternatives
		int nNextLine = nLine + 1;

		while (nNextLine < aLines.GetSize())
		{
			const CString& sNextLine = Misc::GetItem(aLines, nNextLine);
			DICTITEM diAlt;

			if (FromCsv(sNextLine, diAlt) && diAlt.m_sTextIn == m_sTextIn)
			{
				ASSERT(!diAlt.m_sClassID.IsEmpty());

				if (!diAlt.m_sClassID.IsEmpty())
					m_mapAlternatives[diAlt.m_sClassID] = diAlt.m_sTextOut;

				// mark text out as being not-translatable
				// else the translated text can itself be translated!
				CTransTextMgr::IgnoreString(diAlt.m_sTextOut, TRUE);

				nLine++;
				nNextLine++;
			}
			else
				break;
		}

		return TRUE;
	}

	return FALSE;
}

// static helper
BOOL DICTITEM::FromCsv(const CString& sLine, DICTITEM& di)
{
	CStringArray aFields;
	int nNumFields = Misc::Split(sLine, aFields, '\t', TRUE);

	switch (nNumFields)
	{
	case 3:
		di.m_sClassID = aFields[2];
		// fall thru

	case 2:
		di.m_sTextOut = aFields[1];
		TransText::DecodeChars(di.m_sTextOut);
		// fall thru

	case 1:
		di.m_sTextIn = aFields[0];

		if (TransText::PrepareLookupText(di.m_sTextIn))
			return TRUE;
	}

	// all else
	return FALSE;
}

BOOL DICTITEM::Merge(const DICTITEM& di)
{
	ASSERT(di.m_sTextIn == m_sTextIn);

	if (di.m_sTextIn != m_sTextIn)
		return FALSE;

	// merge main ID if not exists
	if (!HasClassID(di.m_sClassID))
		m_mapAlternatives[di.m_sClassID] = di.m_sTextOut;

	// merge alternatives if not exist
	POSITION pos = di.m_mapAlternatives.GetStartPosition();
	
	while (pos)
	{
		CString sAltTextOut, sAltClassID;
		di.m_mapAlternatives.GetNextAssoc(pos, sAltClassID, sAltTextOut);
		
		if (!HasClassID(sAltClassID))
			m_mapAlternatives[sAltClassID] = sAltTextOut;
	}

	return TRUE;
}

BOOL DICTITEM::HasClassID(const CString& sClassID) const
{
	ASSERT(!sClassID.IsEmpty());

	if (m_sClassID == sClassID)
		return TRUE;

	// check alternatives
	CString sTemp;
	return m_mapAlternatives.Lookup(sClassID, sTemp);
}

void DICTITEM::ClearTextOut()
{
	m_sTextOut.Empty();

	// alternatives
	POSITION pos = m_mapAlternatives.GetStartPosition();
	
	while (pos)
	{
		CString sText, sClassID;

		m_mapAlternatives.GetNextAssoc(pos, sClassID, sText);
		m_mapAlternatives[sClassID] = _T("");
	}
}

BOOL DICTITEM::Fixup()
{
	CString sReplaceID, sReplaceText;
	BOOL bCleaned = FALSE;

	if (NeedFixup(m_sClassID, sReplaceID, sReplaceText))
	{
		// overwrite text if empty
		if (m_sTextOut.IsEmpty())
		{
			m_sTextOut = sReplaceText;
		}

		m_sClassID = sReplaceID;
		m_mapAlternatives.RemoveKey(sReplaceID);
		bCleaned = TRUE;
	}

	// alternatives
	POSITION pos = m_mapAlternatives.GetStartPosition();
	
	while (pos)
	{
		CString sText, sClassID, sReplaceText;
		m_mapAlternatives.GetNextAssoc(pos, sClassID, sText);
		
		if (NeedFixup(sClassID, sReplaceID, sReplaceText))
		{
			// save the text for sClassID before deleting
			if (!sText.IsEmpty() && sReplaceText.IsEmpty())
			{
				m_mapAlternatives[sReplaceID] = sText;
			}

			m_mapAlternatives.RemoveKey(sClassID);
			bCleaned = TRUE;
		}
	}

	return bCleaned;
}

BOOL DICTITEM::NeedFixup(const CString& sClassID, CString& sReplaceID, CString& sReplaceText) const
{
	CStringArray aParts;

	if (Misc::Split(sClassID, aParts, '.', TRUE) == 2)
	{
		const CString& sClass = Misc::GetItem(aParts, 0);
		const CString& sID = Misc::GetItem(aParts, 1);

		if (sClass.CompareNoCase(WC_BUTTON) == 0)
		{
			// cleanup any 'button' text if there is a corresponding
			// 'checkbox', 'radiobutton' or 'groupbox' item with the 
			// same control ID
			LPCTSTR szButtons[3] = { _T("checkbox"), _T("radiobutton"), _T("groupbox") };

			for (int nBtn = 0; nBtn < 3; nBtn++)
			{
				CString sTryID;
				sTryID.Format(_T("%s.%s"), szButtons[nBtn], sID);

				// try top level item
				if (m_sClassID == sTryID)
				{
					sReplaceID = m_sClassID;
					sReplaceText = m_sTextOut;
					return TRUE;
				}
				// try alternatives
				else if (m_mapAlternatives.Lookup(sTryID, sReplaceText))
				{
					sReplaceID = sTryID;
					return TRUE;
				}
			}
		}
		else if (sClass == TransText::GetFriendlyClass(WC_TOOLTIPS))
		{
			// Replace any 'tooltip.ID' with 'tooltip'

			// try top level item
			if (m_sClassID == sClass)
			{
				sReplaceID = sClass;
				sReplaceText = m_sTextOut;
				return TRUE;
			}
			// try alternatives
			else if (m_mapAlternatives.Lookup(sClass, sReplaceText))
			{
				sReplaceID = sClass;
				return TRUE;
			}
			// just strip off the ID part
			else if (!sID.IsEmpty())
			{
				sReplaceID = sClass;
				return TRUE;
			}
		}

		// cleanup any items with '.65535'
		if (sID == STATIC_ID)
		{
			// look for item with same class without ctrlID
			if (m_sClassID == sClass)
			{
				sReplaceID = m_sClassID;
				sReplaceText = m_sTextOut;
				return TRUE;
			}
			
			// try alternatives
			if (m_mapAlternatives.Lookup(sID, sReplaceText))
			{
				sReplaceID = sID;
				return TRUE;
			}

			// not found?
			// just remove the control ID
			sReplaceID = sClass;
			sReplaceText = m_sTextOut;
			return TRUE;
		}
	}

	return FALSE;
}

BOOL DICTITEM::Translate(CString& sText)
{
	ASSERT (!sText.IsEmpty() && sText == m_sTextIn);

	if (m_sTextOut.IsEmpty())
	{
		if (WantUppercase())
		{
			sText.MakeUpper();
			FixupFormatString(sText);
			return TRUE;
		}

		// else
		return FALSE;
	}

	// else
	sText = m_sTextOut;

	return TRUE;
}

BOOL DICTITEM::Translate(CString& sText, HWND hWndRef, LPCTSTR szClassID)
{
	if (szClassID && *szClassID)
		return Translate(sText, szClassID);
	 
	// else
	return Translate(sText, TransText::GetClassIDName(hWndRef));
}

BOOL DICTITEM::Translate(CString& sText, HMENU hMenu, int nMenuID)
{
	return Translate(sText, TransText::GetClassIDName(hMenu, nMenuID)); 
}

BOOL DICTITEM::Translate(CString& sText, const CString& sClassID)
{
#ifdef _DEBUG
// 	if (sClassID == _T("text") && sText == _T("Delete Attribute"))
// 	{
// 		int breakpoint = 0;
// 	}
#endif

	// 1. check for an 'alternative' entry
	if (!sClassID.IsEmpty() && !m_mapAlternatives.IsEmpty())
	{
		CString sTextOut;

		if (m_mapAlternatives.Lookup(sClassID, sTextOut) && !sTextOut.IsEmpty())
		{
			sText = sTextOut;
			return TRUE;
		}
	}

	// 2. check root item
	if (sClassID.IsEmpty() || m_sClassID.IsEmpty() || sClassID == m_sClassID)
	{
		BOOL bTrans = Translate(sText);

		// if the root item has no class ID then use this one
		if (WantAddToDictionary() && m_sClassID.IsEmpty() && !sClassID.IsEmpty())
			m_sClassID = sClassID;

		return bTrans;
	}

	// 3. No translation so add as an alternative and use the root translation
	if (WantAddToDictionary())
		m_mapAlternatives[sClassID] = _T("");

	return Translate(sText);
}

///////////////////////////////////////////////////////////////////////////////////

CTransDictionary::CTransDictionary()
{
	LANGID nLangID = Misc::GetUserDefaultUILanguage();
	m_wDictLanguageID = (WORD)PRIMARYLANGID(nLangID);

	// estimated size of dictionary
	InitHashTable(1991); // prime number closest to 2000

	// estimated size of ignore list
	m_mapStringIgnore.InitHashTable(1991); // prime number closest to 2000
}

CTransDictionary::~CTransDictionary()
{
	DeleteDictionary();
}

BOOL CTransDictionary::LoadXmlDictionary(LPCTSTR szDictPath)
{
	if (!FileMisc::HasExtension(szDictPath, _T("xml")))
		return FALSE;

	CXmlFile file(TRANSTEXT_HEADER);

	if (file.Load(szDictPath))
	{
		DeleteDictionary();

		m_sDictVersion = file.GetItemValue(DICTVER);

		// process all top level items
		const CXmlItem* pXITrans = file.GetItem(TRANSLATED);

		if (pXITrans)
		{
			POSITION pos = pXITrans->GetFirstItemPos();
			
			while (pos)
			{	
				const CXmlItem* pXI = pXITrans->GetNextItem(pos);
				LoadItem(pXI);
			}
		}

		const CXmlItem* pXINeedTrans = file.GetItem(NEED_TRANSLATION);

		if (pXINeedTrans)
		{
			POSITION pos = pXINeedTrans->GetFirstItemPos();
			
			while (pos)
			{	
				const CXmlItem* pXI = pXINeedTrans->GetNextItem(pos);
				LoadItem(pXI);
			}
		}

		return TRUE;
	}

	return FALSE;
}

BOOL CTransDictionary::LoadCsvDictionary(LPCTSTR szDictPath)
{
	if (!FileMisc::HasExtension(szDictPath, _T("csv")) &&
		!FileMisc::HasExtension(szDictPath, _T("txt")))
		return FALSE;

	CString sCsvContents;
	CStdioFileEx file;

	if (file.Open(szDictPath, CFile::modeRead) && file.ReadFile(sCsvContents) && !sCsvContents.IsEmpty())
	{
		file.Close();

		DeleteDictionary();

		CStringArray aLines;
		int nNumLines = Misc::Split(sCsvContents, aLines, '\n');

		if (nNumLines)
		{
			// first line is header
			const CString& sHeader = aLines[0];

			if (sHeader.Find(TRANSTEXT_HEADER) != 0)
				return FALSE;

			// extract dictionary version
			if (sHeader != TRANSTEXT_HEADER)
				m_sDictVersion = sHeader.Mid(TRANSTEXT_HEADER.GetLength() + 1);

			// start reading the lines
			int nStartLine = 1;

			// always check there's more to read
			if (nStartLine >= nNumLines)
				return TRUE;

			const CString& sLine = aLines[nStartLine];

			// extract language ID and skip
			if (sLine.Find(PRIMARY_LANGID) == 0)
			{
				m_wDictLanguageID = (WORD)_ttoi(sLine.Mid(PRIMARY_LANGID.GetLength() + 1));
				nStartLine++;
			}

			// always check there's more to read
			if (nStartLine >= nNumLines)
				return TRUE;

			// skip column header
			if (aLines[nStartLine] == CSVCOLUMN_HEADER)
				nStartLine++;

			// always check there's more to read
			if (nStartLine >= nNumLines)
				return TRUE;

			// the dictionary itself
			for (int nLine = nStartLine; nLine < nNumLines; nLine++)
			{
// #ifdef _DEBUG
// 				const CString& sLine = aLines[nLine];
// #endif
				DICTITEM diTemp;

				// this call will pull all consecutive lines having the same text
				if (diTemp.FromCsv(aLines, nLine))
				{
					// NOTE: because we separate translated and untranslated
					// text in the dictionary, and the untranslated comes first
					// we need to append the untranslated AFTER the translated
					// so that when we save the dictionary the untranslated 
					// strings get moved t0 the translated section
					CString sItem(diTemp.GetTextIn());
					DICTITEM* pDI = GetDictItem(sItem, FALSE);

					// if it has we merge the items together
					if (pDI)
					{
						// if our understanding is correct then the existing
						// dictionary should be untranslated and the current
						// one translated
						//	ASSERT(pDI->GetTextOut().IsEmpty());
						//	ASSERT(!diTemp.GetTextOut().IsEmpty());

						// merge the untranslated dictionary into the 
						// translated one and then delete the untranslated
						// before reassign to the dictionary map
						diTemp.Merge(*pDI);
						delete pDI;
					}

					// assign to map
					pDI = new DICTITEM(diTemp);
					SetAt(sItem, pDI);
				}			
			}
		}

		FixupDictionary();

		return TRUE;
	}

	return FALSE;
}

void CTransDictionary::FixupDictionary()
{
	POSITION pos = GetStartPosition();
	BOOL bCleaned = FALSE;
	
	while (pos)
	{
		DICTITEM* pDI = NULL;
		CString sKey;
		
		GetNextAssoc(pos, sKey, pDI);
		ASSERT(pDI && sKey == pDI->GetTextIn());

		if (pDI->Fixup())
			bCleaned = TRUE;
	}

	if (bCleaned)
		SaveDictionary();
}

BOOL CTransDictionary::LoadDictionaryItem(const CXmlItem* pXIDict)
{
	if (pXIDict && pXIDict->NameIs(_T("ITEM")))
	{
		DICTITEM diTemp(pXIDict);

		// make sure this item has not already been mapped
		CString sItem(diTemp.GetTextIn());
		DICTITEM* pDI = GetDictItem(sItem, FALSE);

		// if it has we merge the items together
		if (pDI)
		{
			pDI->Merge(diTemp);
		}
		else
		{
			pDI = new DICTITEM(diTemp);
			SetAt(sItem, pDI);
		}

		// next
		LoadDictionaryItem(pXIDict->GetSibling());
		return TRUE;
	}

	return FALSE;
}

void CTransDictionary::LoadItem(const CXmlItem* pXI)
{
	if (pXI == NULL)
		return;

	// trying loading a DICTITEM
	if (LoadDictionaryItem(pXI))
	{
		// DICTITEM will have loaded itself and all subitems
		// and all siblings
	}
	else // not a DICTITEM so process all children
	{
		POSITION pos = pXI->GetFirstItemPos();

		while (pos)
		{
			const CXmlItem* pXISub = pXI->GetNextItem(pos);
			LoadItem(pXISub);
		}
	}
}

void CTransDictionary::DeleteDictionary()
{
	POSITION pos = GetStartPosition();

	while (pos)
	{
		DICTITEM* pDI = NULL;
		CString sPath;

		GetNextAssoc(pos, sPath, pDI);
		delete pDI;
	}

#ifdef _DEBUG
	if (nDICreated || nDIDestroyed)
	{
		TRACE (_T("%d dictionary items created\n"), nDICreated);
		TRACE (_T("%d dictionary items destroyed\n"), nDIDestroyed);
	}
#endif

	RemoveAll();
	m_sDictFile.Empty();
	m_sDictVersion.Empty();
}

BOOL CTransDictionary::CleanupDictionary(const CTransDictionary& tdMaster, CTransDictionary& tdRemoved)
{
	if (tdMaster.IsEmpty() || IsEmpty())
		return FALSE;

	// build a list of all items not found in 'tdMaster'
	POSITION pos = GetStartPosition();
	CStringArray aMissing;

	DICTITEM* pDI = NULL;
	CString sItem;
	BOOL bCleaned = FALSE;

	while (pos)
	{
		GetNextAssoc(pos, sItem, pDI);

		if (!tdMaster.HasDictItem(sItem))
		{
			aMissing.Add(pDI->GetTextIn());
			bCleaned = TRUE;
		}
	}

	// and move them to 'dtRemoved'
	int nItem = aMissing.GetSize();

	while (nItem--)
	{
		sItem = aMissing[nItem];

		VERIFY(Lookup(sItem, pDI));

		if (RemoveKey(sItem))
			tdRemoved[sItem] = pDI;
	}

	// now add all 'tdMaster' items not found in 'this'
	pos = tdMaster.GetStartPosition();
	
	while (pos)
	{
		DICTITEM* pDIMaster = NULL;
		CString sItemMaster;
		
		tdMaster.GetNextAssoc(pos, sItemMaster, pDIMaster);
		
		if (!HasDictItem(sItemMaster))
		{
			DICTITEM* pDI = new DICTITEM(*pDIMaster);
	
			// remove text out from item before adding
			pDI->ClearTextOut();
			SetAt(sItemMaster, pDI);

			bCleaned = TRUE;
		}
	}

	return bCleaned;
}

void CTransDictionary::IgnoreString(const CString& sText, BOOL bPrepare)
{
	if (sText.IsEmpty())
		return;
	
	CString sTemp(sText);
	
	if (bPrepare && !TransText::PrepareLookupText(sTemp))
		return;
	
	m_mapStringIgnore[sTemp] = NULL;
}

BOOL CTransDictionary::WantIgnore(const CString& sText) const
{
	void* pDummy = 0;
	return m_mapStringIgnore.Lookup(sText, pDummy);
}

BOOL CTransDictionary::SaveDictionary(LPCTSTR szAltPath, BOOL bForce)
{
	if (!bForce && !DICTITEM::WantAddToDictionary())
		return TRUE; // nothing to do

	if (szAltPath == NULL)
	{
		if (m_sDictVersion.IsEmpty())
			return FALSE;
		else
			szAltPath = m_sDictFile;
	}

	if (SaveCsvDictionary(szAltPath) || SaveXmlDictionary(szAltPath))
	{
		m_sDictFile = szAltPath;
		return TRUE;
	}

	return FALSE;
}

BOOL CTransDictionary::SaveXmlDictionary(LPCTSTR szDictPath) const
{
	if (!FileMisc::HasExtension(szDictPath, _T("xml")))
		return FALSE;

	CXmlFile file(TRANSTEXT_HEADER);

	// dtActive version
	file.AddItem(DICTVER, FileMisc::GetAppVersion());

	CXmlItem* pXITrans = file.AddItem(TRANSLATED);
	CXmlItem* pXINeedTrans = file.AddItem(NEED_TRANSLATION);

	// build xml file
	POSITION pos = GetStartPosition();

	while (pos)
	{
		DICTITEM* pDI = NULL;
		CString sKey;

		GetNextAssoc(pos, sKey, pDI);
		ASSERT(pDI && sKey == pDI->GetTextIn());

		// separate translated and non-translated items
		// remember, in DEBUG we always override empty textout
		// with UPPERCASED text
		CXmlItem* pXItem = NULL;

#ifdef _DEBUG
		pXItem = pXITrans->AddItem(ITEM);
#else
		if (pDI->IsTranslated())
			pXItem = pXITrans->AddItem(ITEM);
		else
			pXItem = pXINeedTrans->AddItem(ITEM);
#endif

		if (pXItem)
			pDI->ToXml(pXItem);
	}

	// sort by original text to maintain some sort of order
	pXITrans->SortItems(ITEM, TEXTIN);
	pXINeedTrans->SortItems(ITEM, TEXTIN);

	// mess about with the output to make it easier to understand
	CString sFileContents;
	file.Export(sFileContents);

	sFileContents.Replace(_T("<ITEM"), _T("\t<ITEM"));
	sFileContents.Replace(_T("</ITEM"), _T("\t</ITEM"));
	sFileContents.Replace(_T("<ALTERNATIVE"), _T("\t\t<ALTERNATIVE"));

	return FileMisc::SaveFile(szDictPath, sFileContents, SFE_UNICODE);
}

BOOL CTransDictionary::SaveCsvDictionary(LPCTSTR szDictPath) const
{
	if (!FileMisc::HasExtension(szDictPath, _T("csv")) && 
		!FileMisc::HasExtension(szDictPath, _T("txt")))
		return FALSE;

	// build csv file
	CStringArray aLines, aTranslated, aNeedTranslation;

	// header
	CString sHeader;
	sHeader.Format(_T("%s %s"), TRANSTEXT_HEADER, FileMisc::GetAppVersion());
	aLines.Add(sHeader);

	// language identifier
	CString sLangID;
	sLangID.Format(_T("%s %d"), PRIMARY_LANGID, m_wDictLanguageID);
	aLines.Add(sLangID);

	// column header
	aLines.Add(CSVCOLUMN_HEADER);

	// dictionary
	POSITION pos = GetStartPosition();

	while (pos)
	{
		DICTITEM* pDI = NULL;
		CString sKey;
		CStringArray aTransLines, aNeedTransLines;

		GetNextAssoc(pos, sKey, pDI);
		ASSERT(pDI && sKey == pDI->GetTextIn());

		if (pDI->ToCsv(aTransLines, aNeedTransLines))
		{
			if (pDI->IsTranslated())
			{
				aTranslated.Append(aTransLines);
				aTranslated.Append(aNeedTransLines);
			}
			else
			{
				aNeedTranslation.Append(aTransLines);
				aNeedTranslation.Append(aNeedTransLines);
			}
		}
	}

	// sort by original text to maintain some sort of order
	Misc::SortArray(aTranslated, CompareProc);
	Misc::SortArray(aNeedTranslation, CompareProc);

	// put NEED_TRANSLATION first
	if (aNeedTranslation.GetSize() > 0)
	{
		aLines.Add(NEED_TRANSLATION);
		aLines.Append(aNeedTranslation);
	}

	if (aTranslated.GetSize() > 0)
	{
		aLines.Add(TRANSLATED);
		aLines.Append(aTranslated);
	}

	CString sFileContents = Misc::FormatArray(aLines, _T("\r\n"));

	return FileMisc::SaveFile(szDictPath, sFileContents, SFE_MULTIBYTE, CP_UTF8);
}

int CTransDictionary::CompareProc(const void* pFirst, const void* pSecond)
{
	ASSERT(pFirst && pSecond);

	const CString& sFirst = *((CString*)pFirst);
	const CString& sSecond = *((CString*)pSecond);

	// compare only between the first set of double-quote (ie the input text only)
	int nFirstStart = sFirst.Find(_T("\""));
	int nFirstEnd = sFirst.Find(_T("\"\t"), nFirstStart + 1);

	int nSecondStart = sSecond.Find(_T("\""));
	int nSecondEnd = sSecond.Find(_T("\"\t"), nSecondStart + 1);

	if ((nFirstEnd - nFirstStart) != (nSecondEnd - nSecondStart)) // not the same input string
		return _tcscmp((LPCTSTR)sFirst + nFirstStart, (LPCTSTR)sSecond + nSecondStart);

	// else only compare up to the tab chars
	int nCompare = _tcsncmp((LPCTSTR)sFirst + nFirstStart, (LPCTSTR)sSecond + nSecondStart, nFirstEnd - nFirstStart);

	// if the string is the same make sure the alternatives 
	// go beneath the primary entry
	if (nCompare == 0)
		nCompare = (nFirstStart - nSecondStart);

	return nCompare;
}

BOOL CTransDictionary::HasDictItem(CString& sText) const
{
	if (TransText::PrepareLookupText(sText) && !WantIgnore(sText))
	{
		DICTITEM* pDI = NULL;
		return Lookup(sText, pDI);
	}	

	// all else
	return FALSE;
}

DICTITEM* CTransDictionary::GetDictItem(CString& sText, BOOL bAutoCreate)
{
	// check for valid text
	// and that we're ignoring this item
	if (!TransText::PrepareLookupText(sText) || WantIgnore(sText))
		return NULL;

	// can't auto-create if translating only
	if (DICTITEM::WantTranslateOnly())
		bAutoCreate = FALSE;
	
	// check list of items to be ignored
	// NOTE: it is valid for (text.In == text.Out) so we
	// only check for text out if it would otherwise lead
	// to a new dictionary item
	DICTITEM* pDI = NULL;
	
	if (!Lookup(sText, pDI) && bAutoCreate)
	{
		pDI = new DICTITEM(sText);
		SetAt(sText, pDI);
	}
#ifdef DEBUG
	else if (pDI && !pDI->GetTextOut().IsEmpty())
	{
		int breakpoint = 0;
	}
#endif

	return pDI;
}

BOOL CTransDictionary::LoadDictionary(LPCTSTR szDictPath)
{
	if (LoadCsvDictionary(szDictPath) || LoadXmlDictionary(szDictPath))
	{
		m_sDictFile = szDictPath;
		return TRUE;
	}

	// else
	return FALSE;
}

BOOL CTransDictionary::Translate(CString& sText, HWND hWndRef, LPCTSTR szClassID)
{
	DICTITEM* pDI = GetDictItem(sText); 

	if (pDI && pDI->Translate(sText, hWndRef, szClassID))
	{
		ASSERT(!sText.IsEmpty());

		// mark text out as being not-translatable
		// else the translated text can itself be translated!
		IgnoreString(sText, FALSE);
		return TRUE;
	}
#ifdef DEBUG
	else
	{
		int breakpoint = 0;
	}
#endif

	return FALSE;
}

BOOL CTransDictionary::TranslateMenuShortcut(CString& sShortcut)
{
	if (sShortcut.IsEmpty())
		return FALSE;

	// split by '+'
	CStringArray aParts;
	int nNumParts = Misc::Split(sShortcut, aParts, '+');

	// translate each part and rebuild shortcut as we go
	CString sTransShortcut;

	for (int i = 0; i < nNumParts; i++)
	{
		CString& sPart = aParts[i];
		sPart.MakeUpper();

		if (sPart == "CTRL")
		{
			sPart = Misc::GetKeyName(VK_CONTROL);
		}
		else if (sPart == "SHIFT")
		{
			sPart = Misc::GetKeyName(VK_SHIFT);
		}
		else if (sPart == "ALT")
		{
			sPart = Misc::GetKeyName(VK_MENU);
		}
		else if (sPart == "SPACE" || sPart == "SPACEBAR")
		{
			sPart = Misc::GetKeyName(VK_SPACE);
		}
		else if (sPart == "DEL" || sPart == "DELETE")
		{
			sPart = Misc::GetKeyName(VK_DELETE);
		}
		else if (sPart == "ESC" || sPart == "ESCAPE")
		{
			sPart = Misc::GetKeyName(VK_ESCAPE);
		}
		else // must be the last item
		{
			ASSERT(i == (nNumParts - 1));
			ASSERT((sPart.GetLength() == 1) || ((sPart.GetLength() == 2) && (sPart[0] == 'F')));
		}

		// add delimiter
		if (!sTransShortcut.IsEmpty())
			sTransShortcut += '+';
		
		sTransShortcut += sPart;
	}

	// replace shortcut
	sShortcut = sTransShortcut;

	return (!sShortcut.IsEmpty());
}
			
BOOL CTransDictionary::Translate(CString& sItem, HMENU hMenu, int nMenuID)
{
	// trim off trailing shortcut
	CString sShortcut;
	int nTab = sItem.Find('\t');
				
	// remove it
	if (nTab >= 0)
	{
		sShortcut = sItem.Mid(nTab+1);
		sItem = sItem.Left(nTab);
	}

	DICTITEM* pDI = GetDictItem(sItem);
	
	if (pDI && pDI->Translate(sItem, hMenu, nMenuID))
	{
		// translate and reattach any shortcut
		if (TranslateMenuShortcut(sShortcut))
		{
			sItem += '\t';
			sItem += sShortcut;
		}
		
		// mark text out as being not-translatable
		// else the translated text can itself be translated!
		IgnoreString(sItem, FALSE);
		
		return TRUE;
	}
#ifdef DEBUG
	else
	{
		int breakpoint = 0;
	}
#endif
	
	return FALSE;
}

///////////////////////////////////////////////////////////////////////////////////

CTransTextMgr::CTransTextMgr()
{
}

CTransTextMgr::~CTransTextMgr()
{
	m_dictionary.DeleteDictionary();
}

BOOL CTransTextMgr::Initialize(LPCTSTR szDictPath, ITT_TRANSLATEOPTION nOption)
{
	return GetInstance().InitHooks(szDictPath, nOption);
}

void CTransTextMgr::Release()
{
	GetInstance().m_dictionary.SaveDictionary();
	GetInstance().ReleaseHooks();
}

BOOL CTransTextMgr::InitHooks(LPCTSTR szDictPath, ITT_TRANSLATEOPTION nOption)
{
	if (!m_dictionary.LoadDictionary(szDictPath))
	{
		// if the dictionary did not load and we're only
		// in translation mode then we fail
		if (nOption == ITTTO_TRANSLATEONLY)
			return FALSE;
	}

	DICTITEM::SetTranslationOption(nOption);

	return CHookWndMgr<CTransTextMgr>::InitHooks();
}

LPCTSTR CTransTextMgr::GetDictionaryFile() 
{ 
	return GetInstance().m_dictionary.GetDictionaryPath(); 
}

LPCTSTR CTransTextMgr::GetDictionaryVersion() 
{ 
	return GetInstance().m_dictionary.GetDictionaryVersion(); 
}

ITT_TRANSLATEOPTION CTransTextMgr::GetTranslationOption() 
{ 
	return DICTITEM::GetTranslationOption(); 
}

BOOL CTransTextMgr::CleanupDictionary(LPCTSTR szMasterDictPath)
{
	if (CleanupDictionary(GetDictionaryFile(), szMasterDictPath))
	{
		// reload dictionary
		GetInstance().m_dictionary.LoadDictionary(GetDictionaryFile());
		return TRUE;
	}

	return FALSE;
}
	
BOOL CTransTextMgr::CleanupDictionary(LPCTSTR szDictPath, LPCTSTR szMasterDictPath)
{
	CTransDictionary dtActive, dtMaster, dtRemoved;
	
	if (!dtActive.LoadDictionary(szDictPath) || !dtMaster.LoadDictionary(szMasterDictPath))
		return FALSE;
	
	if (dtActive.CleanupDictionary(dtMaster, dtRemoved))
	{
		dtActive.SaveDictionary(NULL, TRUE);
		
		// save removed items
		if (!dtRemoved.IsEmpty())
		{
			CString sRemovedPath = CFileBackup::BuildBackupPath(szDictPath, _T("backup"), FBS_TIMESTAMP, _T(".removed"));

			VERIFY(FileMisc::CreateFolderFromFilePath(sRemovedPath));
			VERIFY(dtRemoved.SaveDictionary(sRemovedPath));
		}

		return TRUE;
	}
	
	return FALSE;
}

void CTransTextMgr::SetTranslationOption(ITT_TRANSLATEOPTION nOption)
{
	// if adding to dictionary there must be a dictionary specified
	CTransTextMgr& ttm = GetInstance();

	if ((nOption == ITTTO_ADD2DICTIONARY) || !ttm.m_dictionary.GetDictionaryPath().IsEmpty())
	{
		DICTITEM::SetTranslationOption(nOption);
	}
}

void CTransTextMgr::ReleaseHooks()
{
	// cleanup
	m_dictionary.DeleteDictionary();

	CHookWndMgr<CTransTextMgr>::Release();
}

void CTransTextMgr::PostHookWnd(HWND hWnd)
{
	CTransWnd* pTWnd = (CTransWnd*)GetHookWnd(hWnd);
	ASSERT (pTWnd);

	if (pTWnd)
		pTWnd->PostHookWindow();
}

void CTransTextMgr::IgnoreString(const CString& sText, BOOL bPrepare)
{
	GetInstance().m_dictionary.IgnoreString(sText, bPrepare);
}

BOOL CTransTextMgr::TranslateText(CString& sText, HWND hWndRef, LPCTSTR szClassID)
{
	return GetInstance().m_dictionary.Translate(sText, hWndRef, szClassID);
}

void CTransTextMgr::EnableTranslation(HWND hWnd, BOOL bEnable)
{
	if (hWnd)
	{
		CTransTextMgr& ttm = GetInstance();

		if (bEnable)
		{
			// we don't hook it ourselves
			// just let the normal mechanism do its stuff

			// remove from map
			ttm.m_mapWndIgnore.RemoveKey(hWnd);
		}
		else // unhook window if already hooked
		{
			if (ttm.WantTranslation(hWnd))
				ttm.RemoveHookWnd(hWnd);

			// add to map
			ttm.m_mapWndIgnore[hWnd] = NULL;
		}
	}
}

BOOL CTransTextMgr::WantTranslation(HWND hWnd, UINT nMsg) const
{
	void* pDummy = NULL;

	if (m_mapWndIgnore.Lookup(hWnd, pDummy))
		return FALSE;

	// weed out various window types
	switch (nMsg)
	{
	case WM_SETTEXT:
		{
			CString sClass = CWinClasses::GetClass(hWnd);

			if (CWinClasses::IsEditControl(sClass))
			{
				return FALSE;
			}
			else if (CWinClasses::IsCommonDialog(hWnd))
			{
				return FALSE;
			}
			// if we are a combobox and our parent is a comboboxex
			// also check that window for ignoring
			else if (CWinClasses::IsComboBox(sClass))
			{
				HWND hParent = ::GetParent(hWnd);

				if (CWinClasses::IsClass(hParent, WC_COMBOBOXEX))
				{
					if (m_mapWndIgnore.Lookup(hParent, pDummy))
						return FALSE;
				}
			}
		}
		break;
	}

	// else
	return TRUE;
}

void CTransTextMgr::EnableTranslation(HMENU hMenu, BOOL bEnable)
{
	if (hMenu)
	{
		for (int nItem = 0; nItem < ::GetMenuItemCount(hMenu); nItem++)
		{
			EnableTranslation(::GetMenuItemID(hMenu, nItem), bEnable);
		}
	}
}

void CTransTextMgr::EnableTranslation(UINT nMenuID, BOOL bEnable)
{
	if (nMenuID && (nMenuID != POPUPMENU_ID))
	{
		CTransTextMgr& ttm = GetInstance();

		if (bEnable)
			ttm.m_mapMenuIgnore.RemoveKey(nMenuID);
		else
			ttm.m_mapMenuIgnore[nMenuID] = NULL;
	}
}

BOOL CTransTextMgr::WantTranslation(UINT nMenuID) const
{
	void* pDummy = NULL;
	
	if (nMenuID && m_mapMenuIgnore.Lookup(nMenuID, pDummy))
		return FALSE;
	
	// else
	return TRUE;
}

BOOL CTransTextMgr::RemoveHookWnd(HWND hWnd)
{
	if (hWnd)
		return UnhookWnd(hWnd);
	
	// else
	return FALSE;
}

CSubclassWnd* CTransTextMgr::NewHookWnd(HWND hWnd, const CString& sClass, DWORD dwStyle) const
{
	// pre-check
	if (!WantTranslation(hWnd))
		return NULL;

	return CTransWnd::NewTransWnd(sClass, dwStyle);
}

void CTransTextMgr::UpdateMenu(HWND hWnd)
{
	CTransTextMgr& ttm = GetInstance();
	CTransWnd* pTWnd = (CTransWnd*)ttm.GetHookWnd(hWnd);

	if (pTWnd)
		pTWnd->UpdateMenu();
}

BOOL CTransTextMgr::HandleInitMenuPopup(HWND hWnd, UINT nMsg, WPARAM wp, LPARAM /*lp*/)
{
	if (nMsg != WM_INITMENUPOPUP)
		return FALSE;
	
	if (!WantTranslation(hWnd, nMsg))
		return FALSE;
	
	TranslateMenu((HMENU)wp, hWnd);

	return TRUE; // handled
}

/*
BOOL CTransTextMgr::HandleSetText(HWND hWnd, UINT nMsg, WPARAM wp, LPARAM lp)
{
	if (nMsg != WM_SETTEXT || !lp)
		return FALSE;

	if (!WantTranslation(hWnd, nMsg))
		return FALSE;

	LPCTSTR szText = (LPCTSTR)lp;

	if (*szText == 0)
		return FALSE;

	CString sText(szText);
		
	if (TranslateText(sText, hWnd))
	{
		MSG msg = { hWnd, nMsg, wp, (LPARAM)(LPCTSTR)sText, 0, { 0, 0 } };
		CHookWndMgr<CTransTextMgr>::OnCallWndProc(msg);

		return TRUE; // handled
	}

	// else 
	return FALSE;
}
*/

BOOL CTransTextMgr::HandleTootipNeedText(HWND hWnd, UINT nMsg, WPARAM /*wp*/, LPARAM lp)
{
	if (nMsg != WM_NOTIFY)
		return FALSE;

	NMHDR* pNMHDR = (NMHDR*)lp;

	if (pNMHDR->code != TTN_NEEDTEXTA && pNMHDR->code != TTN_NEEDTEXTW)
		return FALSE;

	if (!WantTranslation(hWnd, nMsg))
		return FALSE;

	// need to handle both ANSI and UNICODE versions of the message
	TOOLTIPTEXTA* pTTTA = (TOOLTIPTEXTA*)pNMHDR;
	TOOLTIPTEXTW* pTTTW = (TOOLTIPTEXTW*)pNMHDR;
	
	UINT_PTR nID = pNMHDR->idFrom;
	
	if (pNMHDR->code == TTN_NEEDTEXTA && (pTTTA->uFlags & TTF_IDISHWND) ||
		pNMHDR->code == TTN_NEEDTEXTW && (pTTTW->uFlags & TTF_IDISHWND))
	{
		// idFrom is actually the HWND of the tool
		nID = ::GetDlgCtrlID((HWND)nID);
	}
	
	if (nID == 0) // will be zero on a separator
		return FALSE;

	// see if the tooltip has already been handled
	static CString strTipText;

	if (pNMHDR->code == TTN_NEEDTEXTA)
	{
		strTipText = pTTTA->szText;

		if (strTipText.IsEmpty())
			strTipText = pTTTA->lpszText;
	}
	else // TTN_NEEDTEXTW
	{
		strTipText = pTTTW->szText;

		if (strTipText.IsEmpty())
			strTipText = pTTTW->lpszText;
	}

	if (strTipText.IsEmpty())
	{
		CString sFullText;
		
		// don't handle the message if no string resource found
		if (!sFullText.LoadString(nID))
			return FALSE;
		
		// this is the command id, not the button index
		AfxExtractSubString(strTipText, sFullText, 1, '\n');
		
		if (strTipText.IsEmpty())
			return FALSE;
	}
	
	// override class id
	TranslateText(strTipText, hWnd, _T("tooltip"));
	
	// copy back to TOOLTIPTEXT
#ifndef _UNICODE
	if (pNMHDR->code == TTN_NEEDTEXTA)
	{
		strncpy(pTTTA->szText, strTipText, sizeof(pTTTA->szText));
		pTTTA->lpszText = (LPTSTR)(LPCTSTR)strTipText;
	}
	else // TTN_NEEDTEXTW
	{
		_mbstowcsz(pTTTW->szText, strTipText, _countof(pTTTW->szText));
	}
#else
	if (pNMHDR->code == TTN_NEEDTEXTA)
	{
		_wcstombsz(pTTTA->szText, strTipText, _countof(pTTTA->szText));
	}
	else // TTN_NEEDTEXTW
	{
#if _MSC_VER >= 1400
		wcsncpy_s(pTTTW->szText, 80, strTipText, _countof(pTTTW->szText));
#else
		wcsncpy(pTTTW->szText, strTipText, _countof(pTTTW->szText));
#endif
		pTTTW->lpszText = (LPTSTR)(LPCTSTR)strTipText;
	}
#endif
	
	return TRUE; // handled
}

BOOL CTransTextMgr::OnCallWndProc(const MSG& msg)
{   
#ifdef _USRDLL
	// If this is a DLL, need to set up MFC state
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
#endif
	
	BOOL bRes = FALSE;
	
	switch (msg.message)
	{
// 	case WM_SETTEXT:
// 		// let hook wnd override
// 		if (!IsHooked(msg.hwnd))
// 			bRes = HandleSetText(msg.hwnd, msg.message, msg.wParam, msg.lParam);
// 		break;

	case WM_INITMENUPOPUP:
		// let hook wnd override
		if (!IsHooked(msg.hwnd))
			bRes = HandleInitMenuPopup(msg.hwnd, msg.message, msg.wParam, msg.lParam);
		break;

	case WM_NOTIFY:
		// let hook wnd override
		if (!IsHooked(msg.hwnd))
			bRes = HandleTootipNeedText(msg.hwnd, msg.message, msg.wParam, msg.lParam);
		break;

	case WM_CONTEXTMENU:
		// we handle this directly
		break;

	default:
		bRes = CHookWndMgr<CTransTextMgr>::OnCallWndProc(msg);
		break;
	}
	
	return bRes;
}


BOOL CTransTextMgr::WantHookWnd(HWND hWnd, UINT nMsg, WPARAM wp, LPARAM lp) const
{
	// filter out marked windows
	if (!WantTranslation(hWnd))
	{
		return FALSE;
	}
	else if (CHookWndMgr<CTransTextMgr>::WantHookWnd(hWnd, nMsg, wp, lp))
	{
		// we hook dialog/popup only when then are about to be shown.
		// as a byproduct their children will be hooked also at that point
		if (nMsg != WM_PARENTNOTIFY)
		{
			if (CWinClasses::IsDialog(hWnd))
			{
				// only translate common dialogs if our dictionary
				// differs from the users GUI language
				if (CWinClasses::IsCommonDialog(hWnd))
				{
					//static LANGID nUILangID = Misc::GetUserDefaultUILanguage();
					
					//if (PRIMARYLANGID(nUILangID) == m_wDictLanguageID)
						return FALSE;
				}
#ifdef _DEBUG	// don't translate debug windows
				else
				{
					CString sDlgTitle;
					CWnd::FromHandle(hWnd)->GetWindowText(sDlgTitle);
					
					if (sDlgTitle.Find(_T("Microsoft Visual C++ Debug Library")) == 0)
						return FALSE;
				}
#endif
				return TRUE; 
			}
			else if (TransText::IsPopup(hWnd))
			{
				return TRUE; 
			}
		}
		else if (GetHookWnd(::GetParent(hWnd)))
			return TRUE;

		// else 
		return FALSE;
	}

	return FALSE;
}

BOOL CTransTextMgr::TranslateMenu(HMENU hMenu, HWND hWndRef, BOOL bRecursive)
{
	if (!hMenu || !::IsMenu(hMenu))
		return FALSE;

	int nCount = (int)::GetMenuItemCount(hMenu);
	CTransTextMgr& ttm = CTransTextMgr::GetInstance();

	for (int nPos = 0; nPos < nCount; nPos++)
	{
		int nCmdID = (int)::GetMenuItemID(hMenu, nPos);

		// we don't do separators or ownerdraw or menus tagged as not-translatable
		if (!nCmdID || TransText::IsOwnerDraw(nCmdID, hMenu) || !ttm.WantTranslation(nCmdID))
			continue;

		CString sItem;
		int nLen = ::GetMenuString(hMenu, nPos, NULL, 0, MF_BYPOSITION);

		::GetMenuString(hMenu, nPos, sItem.GetBuffer(nLen + 1), nLen + 1, MF_BYPOSITION);
		sItem.ReleaseBuffer();

		if (!sItem.IsEmpty())
		{
			if (ttm.m_dictionary.Translate(sItem, hMenu, nCmdID))
			{
				ASSERT (!sItem.IsEmpty());
			
				MENUITEMINFO minfo;
				minfo.cbSize = sizeof(minfo);
				minfo.fMask = MIIM_STRING;
				minfo.dwTypeData = (LPTSTR)(LPCTSTR)sItem;
				
				SetMenuItemInfo(hMenu, nPos, TRUE, &minfo);
			}
		
			// submenus?
			if (bRecursive && nCmdID == -1)
			{
				HMENU hSubMenu = ::GetSubMenu(hMenu, nPos);

				if (hSubMenu)
					TranslateMenu(hSubMenu, hWndRef, bRecursive);
			}
		}
	}

	return TRUE;
}

//////////////////////////////////////////////////////////////////////
// CTransWnd implementation

CTransWnd::CTransWnd(DWORD dwOptions) : m_bInit(FALSE), m_bAllowTranslate(TRUE), m_dwOptions(dwOptions)
{
}

CTransWnd::~CTransWnd() 
{
}
	
BOOL CTransWnd::HookWindow(HWND hRealWnd, CSubclasser* pSubclasser)
{
	return CSubclassWnd::HookWindow(hRealWnd, pSubclasser);
}

void CTransWnd::PostHookWindow()
{
	ASSERT(IsHooked());

	m_bInit = TRUE;
	Initialize();
	m_bInit = FALSE;
}

void CTransWnd::TranslateMenu(HMENU hMenu, BOOL bToplevelOnly)
{
	CTransTextMgr::TranslateMenu(hMenu, GetHwnd(), bToplevelOnly);
}

void CTransWnd::Initialize()
{
	CWnd* pThis = GetCWnd();
	
	// caption
	CString sText;
	pThis->GetWindowText(sText);

	if (TranslateText(sText))
		pThis->SetWindowText(sText);

	// menu - captioned windows only
	if (GetStyle() & WS_CAPTION)
	{
		HMENU hMenu = ::GetMenu(*pThis);

		if (hMenu && ::IsMenu(hMenu))	
			TranslateMenu(hMenu, TRUE);
	}
}

BOOL CTransWnd::TranslateText(CString& sText) 
{
	if (!m_bAllowTranslate)
		return FALSE;

	return CTransTextMgr::TranslateText(sText, GetHwnd());
}

LRESULT CTransWnd::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	// let others do their stuff first then we translate that
	// if required
	LRESULT lRes = CSubclassWnd::WindowProc(hRealWnd, msg, wp, lp);

	switch (msg)
	{
	case WM_NOTIFY:
		{
			NMHDR* pNMHDR = (NMHDR*)lp;
			
			if ((pNMHDR->code == TTN_NEEDTEXTA || pNMHDR->code == TTN_NEEDTEXTW))
			{
				if (HasFlag(TWS_HANDLETOOLTIPS))
				{
					CTransTextMgr::GetInstance().HandleTootipNeedText(hRealWnd, msg, wp, lp);
				}
			}
		}
		break;

// 	case WM_SETTEXT:
// 		if (HasFlag(TWS_HANDLESETTEXT))
// 		{
// 			// call default handling
// 			CTransTextMgr::GetInstance().HandleSetText(hRealWnd, msg, wp, lp);
// 		}
// 		break;
		
	case WM_INITMENUPOPUP:
		if (HasFlag(TWS_HANDLEMENUPOPUP))
		{
			// call default handling
			CTransTextMgr::GetInstance().HandleInitMenuPopup(hRealWnd, msg, wp, lp);
		}
		break;
	}
		
	return lRes;
}

CTransWnd* CTransWnd::NewTransWnd(const CString& sClass, DWORD dwStyle)
{
	if (CWinClasses::IsClass(sClass, WC_STATIC))
	{
		return new CTransWnd; // standard
	}
	else if (CWinClasses::IsClass(sClass, WC_BUTTON))
	{
		return new CTransWnd; // standard
	}
	else if (CWinClasses::IsClass(sClass, WC_COMBOBOX))
	{
		// we do not translate ownerdraw (unless CBS_HASSTRINGS) or dropdown
		BOOL bOwnerDraw = (dwStyle & (CBS_OWNERDRAWFIXED | CBS_OWNERDRAWVARIABLE));
		BOOL bHasStrings = (dwStyle & CBS_HASSTRINGS);

		UINT nStyle = (dwStyle & 0xf);
		if ((nStyle == CBS_DROPDOWNLIST) && (!bOwnerDraw || bHasStrings))
			return new CTransComboBox;
	}
	else if (CWinClasses::IsClass(sClass, WC_LISTBOX) || 
			CWinClasses::IsClass(sClass, WC_CHECKLISTBOX))
	{
		// we do not translate ownerdraw (unless LBS_HASSTRINGS) 
		BOOL bOwnerDraw = (dwStyle & (LBS_OWNERDRAWFIXED | LBS_OWNERDRAWVARIABLE));
		BOOL bHasStrings = (dwStyle & LBS_HASSTRINGS);

		if (!bOwnerDraw || bHasStrings)
		{
			BOOL bCheckLB = CWinClasses::IsClass(sClass, WC_CHECKLISTBOX);
			return new CTransListBox(bCheckLB);
		}
	}
	else if (CWinClasses::IsClass(sClass, WC_TOOLBAR))
	{
		return new CTransWnd;//ToolBar;
	}
	else if (CWinClasses::IsClass(sClass, WC_STATUSBAR))
	{
		return new CTransStatusBar;
	}
	else if (CWinClasses::IsClass(sClass, WC_TABCONTROL))
	{
		return new CTransTabCtrl; 
	}
	else if (CWinClasses::IsClass(sClass, WC_COMBOBOXEX))
	{
		// we do not translate dropdown
		if (dwStyle & CBS_SIMPLE) // this will also catch CBS_DROPDOWNLIST
			return new CTransComboBoxEx;
	}
	else if (CWinClasses::IsClass(sClass, WC_HEADER))
	{
		return new CTransHeaderCtrl; // we translate the item text
	}
	else if (CWinClasses::IsClass(sClass, WC_LISTVIEW))
	{
		return new CTransListCtrl; // we translate the header item text
	}
	else if (CWinClasses::IsClass(sClass, WC_DIALOGBOX))
	{
		return new CTransWnd; // standard
	}
	else if (CWinClasses::IsClass(sClass, WC_TOOLTIPS))
	{
		return new CTransTooltips; 
	}
	else if (CWinClasses::IsClassEx(sClass, WC_MFCMDIFRAME))
	{
		return new CTransWnd; 
	}
	else if (CWinClasses::IsClassEx(sClass, WC_MFCFRAME))
	{
		// don't translated application title because it's dynamic
		return new CTransWnd/*(TWS_NOHANDLESETTEXT)*/; 
	}

	// everything else
	return NULL;
}

//////////////////////////////////////////////////////////////////////
// CTransWnd derived classes

void CTransComboBox::Initialize() 
{
	if (!m_bAllowTranslate)
		return;

	CString sText;
	int nNumItem = SendMessage(CB_GETCOUNT);
	int nSel = SendMessage(CB_GETCURSEL);
// 	BOOL bSorted = HasStyle(CBS_SORT);

	// Build an array of items
	CCbItemArray aItems;
	aItems.SetSize(nNumItem);

	int nItem;
	for (nItem = 0; nItem < nNumItem; nItem++)
	{
		CBITEMDATA cbid;

		int nLen = SendMessage(CB_GETLBTEXTLEN, nItem);
		SendMessage(CB_GETLBTEXT, nItem, (LPARAM)cbid.sText.GetBuffer(nLen + 1));
		cbid.sText.ReleaseBuffer();

		cbid.dwData = SendMessage(CB_GETITEMDATA, nItem);

		aItems.SetAt(nItem, cbid);
	};

	// delete existing content
	SendMessage(CB_RESETCONTENT, 0, 0);

	// re-add, translating as we go
	for (nItem = 0; nItem < nNumItem; nItem++)
	{
		CBITEMDATA cbid = aItems.GetAt(nItem);

		TranslateText(cbid.sText);

		int nIndex = SendMessage(CB_ADDSTRING, nItem, (LPARAM)(LPCTSTR)cbid.sText);
		SendMessage(CB_SETITEMDATA, nIndex, cbid.dwData);
	}

	SendMessage(CB_SETCURSEL, nSel);
	
	// update combo drop width
	CComboBox* pCombo = (CComboBox*)GetCWnd();
	CDialogHelper::RefreshMaxDropWidth(*pCombo, NULL, 20);

	// send a selection change to update the window text
	// This is a DODGY HACK to make CCheckComboBoxes work
	pCombo->SendMessage(WM_COMMAND, MAKEWPARAM(GetDlgCtrlID(), CBN_SELENDOK), (LPARAM)GetHwnd());
}

LRESULT CTransComboBox::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	// we don't handle if this originated in Initialize() else
	// cos it'll already have been done
	if (!m_bInit && m_bAllowTranslate)
	{
		switch (msg)
		{
		case CB_ADDSTRING:
		case CB_INSERTSTRING:
			{
				CString sText((LPCTSTR)lp);

				if (TranslateText(sText))
					return CSubclassWnd::WindowProc(hRealWnd, msg, wp, (LPARAM)(LPCTSTR)sText);
			}
			break;
		}
	}
	
	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

void CTransComboBoxEx::Initialize() 
{
	if (!m_bAllowTranslate)
		return;

	TCHAR szText[255];
	
	COMBOBOXEXITEM cbi;
	cbi.mask = CBEIF_TEXT;
	cbi.pszText = szText;
	cbi.cchTextMax = 255;

	int nItem = SendMessage(CB_GETCOUNT);

	while (nItem--)
	{
		cbi.iItem = nItem;

		if (SendMessage(CBEM_GETITEM, 0, (LPARAM)&cbi))
			SendMessage(CBEM_SETITEM, 0, (LPARAM)&cbi); // will get handled in WindowProc
	}
}

LRESULT CTransComboBoxEx::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	if (m_bAllowTranslate)
	{
		switch (msg)
		{
		case CBEM_INSERTITEM:
		case CBEM_SETITEM:
			{
				COMBOBOXEXITEM* pItem = (COMBOBOXEXITEM*)lp;

				if ((pItem->mask & CBEIF_TEXT) && pItem->pszText)
				{
					CString sText(pItem->pszText);

					if (!sText.IsEmpty() && TranslateText(sText))
					{
						TCHAR* szOrgText = pItem->pszText;
						pItem->pszText = (LPTSTR)(LPCTSTR)sText;

						LRESULT lr = CTransWnd::WindowProc(hRealWnd, msg, wp, lp);

						pItem->pszText = szOrgText;
						return lr;
					}
				}
			}
			break;
		}
	}
	
	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

void CTransListBox::Initialize()
{
	if (!m_bAllowTranslate)
		return;

	CString sText;
	int nNumItem = SendMessage(LB_GETCOUNT);
	int nSel = SendMessage(LB_GETCURSEL);
//	BOOL bSorted = HasStyle(LBS_SORT);
	
	// Build an array of items
	CLbItemArray aItems;
	aItems.SetSize(nNumItem);

	CCheckListBox* pCLB = m_bCheckLB ? (CCheckListBox*)GetCWnd() : NULL;

	int nItem;
	for (nItem = 0; nItem < nNumItem; nItem++)
	{
		LBITEMDATA lbid;

		int nLen = SendMessage(LB_GETTEXTLEN, nItem);
		SendMessage(LB_GETTEXT, nItem, (LPARAM)lbid.sText.GetBuffer(nLen + 1));
		lbid.sText.ReleaseBuffer();

		lbid.dwData = SendMessage(LB_GETITEMDATA, nItem);
		lbid.nCheck = pCLB ? pCLB->GetCheck(nItem) : 0;

		aItems.SetAt(nItem, lbid);
	};

	// delete existing content
	SendMessage(LB_RESETCONTENT, 0, 0);

	// re-add, translating as we go
	for (nItem = 0; nItem < nNumItem; nItem++)
	{
		LBITEMDATA lbid = aItems.GetAt(nItem);

		TranslateText(lbid.sText);

		int nIndex = SendMessage(LB_ADDSTRING, nItem, (LPARAM)(LPCTSTR)lbid.sText);
		SendMessage(LB_SETITEMDATA, nIndex, lbid.dwData);

		if (pCLB)
			pCLB->SetCheck(nIndex, lbid.nCheck);
	}

	SendMessage(LB_SETCURSEL, nSel);
	
	// update column width
	CListBox* pList = (CListBox*)GetCWnd();
	CDialogHelper::RefreshMaxColumnWidth(*pList);
}

LRESULT CTransListBox::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	// we don't handle if this originated in Initialize() else
	// cos it'll already have been done
	if (!m_bInit && m_bAllowTranslate)
	{
		switch (msg)
		{
		case LB_ADDSTRING:
		case LB_INSERTSTRING:
			{
				CString sText((LPCTSTR)lp);

				if (TranslateText(sText))
					return CSubclassWnd::WindowProc(hRealWnd, msg, wp, (LPARAM)(LPCTSTR)sText);
			}
			break;
		}
	}
	
	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

void CTransTabCtrl::Initialize() 
{
	if (m_bAllowTranslate)
	{
		TCHAR szText[255];
		HWND hThis = GetHwnd();
		
		TCITEM tci;
		tci.mask = TCIF_TEXT;
		tci.pszText = szText;
		tci.cchTextMax = 255;

		int nItem = TabCtrl_GetItemCount(hThis);

		while (nItem--)
		{
			if (TabCtrl_GetItem(hThis, nItem, &tci))
				TabCtrl_SetItem(hThis, nItem, &tci); // will get handled in WindowProc
		}
	}
}

LRESULT CTransTabCtrl::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	switch (msg)
	{
	case TCM_INSERTITEM:
	case TCM_SETITEM:
		if (m_bAllowTranslate)
		{
			TCITEM* pItem = (TCITEM*)lp;

			if ((pItem->mask & TCIF_TEXT) && pItem->pszText)
			{
				CString sText(pItem->pszText);

				if (!sText.IsEmpty() && TranslateText(sText))
				{
					TCHAR* szOrgText = pItem->pszText;
					pItem->pszText = (LPTSTR)(LPCTSTR)sText;

					LRESULT lr = CTransWnd::WindowProc(hRealWnd, msg, wp, lp);

					pItem->pszText = szOrgText;
					return lr;
				}
			}
		}
		break;
	}
	
	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

void CTransHeaderCtrl::Initialize() 
{
	if (!m_bAllowTranslate)
		return;

	TCHAR szText[255];
	HWND hThis = GetHwnd();
	
	HDITEM hdi;
	hdi.mask = HDI_TEXT;
	hdi.pszText = szText;
	hdi.cchTextMax = 255;

	int nItem = Header_GetItemCount(hThis);

	while (nItem--)
	{
		if (Header_GetItem(hThis, nItem, &hdi))
			Header_SetItem(hThis, nItem, &hdi); // will get handled in WindowProc
	}
}

LRESULT CTransHeaderCtrl::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	switch (msg)
	{
	case HDM_INSERTITEM:
	case HDM_SETITEM:
		if (m_bAllowTranslate)
		{
			HDITEM* pItem = (HDITEM*)lp;

			if ((pItem->mask & HDI_TEXT) && pItem->pszText && *(pItem->pszText))
			{
				CString sText(pItem->pszText);

				if (TranslateText(sText))
				{
					TCHAR* szOrgText = pItem->pszText;
					pItem->pszText = (LPTSTR)(LPCTSTR)sText;

					LRESULT lr = CTransWnd::WindowProc(hRealWnd, msg, wp, lp);

					pItem->pszText = szOrgText;
					return lr;
				}
			}
		}
		break;
	}
	
	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

void CTransListCtrl::Initialize() 
{
	// do nothing. header control will handle itself
}

LRESULT CTransListCtrl::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	switch (msg)
	{
	case LVM_INSERTCOLUMN:
	case LVM_SETCOLUMN:
		if (m_bAllowTranslate)
		{
			LVCOLUMN* pItem = (LVCOLUMN*)lp;

			if ((pItem->mask & LVIF_TEXT) && pItem->pszText)
			{
				CString sText(pItem->pszText);

				if (!sText.IsEmpty() && TranslateText(sText))
				{
					TCHAR* szOrgText = pItem->pszText;
					pItem->pszText = (LPTSTR)(LPCTSTR)sText;

					LRESULT lr = CTransWnd::WindowProc(hRealWnd, msg, wp, lp);

					pItem->pszText = szOrgText;
					return lr;
				}
			}
		}
		break;
	}
	
	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

void CTransTooltips::Initialize()
{
	if (!m_bAllowTranslate)
		return;

	HWND hThis = GetHwnd();
	int nItem = ::SendMessage(hThis, TTM_GETTOOLCOUNT, 0, 0);

	while (nItem--)
	{
		TOOLINFO ti =  { 0 };
		ti.cbSize = sizeof(ti);

		if (::SendMessage(hThis, TTM_GETTEXT, nItem, (LPARAM)&ti))
			::SendMessage(hThis, TTM_SETTOOLINFO, nItem, (LPARAM)&ti);
	}
}

LRESULT CTransTooltips::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	switch (msg)
	{
	case TTM_UPDATETIPTEXT:
		{
			TOOLINFO* pTI = (TOOLINFO*)lp;
			CString sText(pTI->lpszText);
			
			if (!sText.IsEmpty() && TranslateText(sText))
			{
				TCHAR* szOrgText = pTI->lpszText;
				pTI->lpszText = (LPTSTR)(LPCTSTR)sText;
				
				LRESULT lr = CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
				
				pTI->lpszText = szOrgText;
				return lr;
			}
		}
		break;
	}

	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

void CTransToolBar::Initialize()
{
	// handle tooltips too
//	CToolBar* pToolBar = (CToolBar*)GetCWnd();
//	CToolBarCtrl& ttCtrl = pToolBar->GetToolBarCtrl();


}

LRESULT CTransToolBar::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
// 	switch (msg)
// 	{
// 
// 	}

	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

void CTransStatusBar::Initialize()
{
	if (!m_bAllowTranslate)
		return;

	TCHAR szText[255];
	HWND hThis = GetHwnd();
	
	int nItem = ::SendMessage(hThis, SB_GETPARTS, 0, NULL);

	while (nItem--)
	{
		if (::SendMessage(hThis, SB_GETTEXT, nItem, (LPARAM)(LPCTSTR)szText))
			::SendMessage(hThis, SB_SETTEXT, nItem, (LPARAM)(LPCTSTR)szText);

		if (::SendMessage(hThis, SB_GETTIPTEXT, MAKEWPARAM(nItem, 255), (LPARAM)(LPCTSTR)szText))
			::SendMessage(hThis, SB_SETTIPTEXT, nItem, (LPARAM)(LPCTSTR)szText);
	}
}

LRESULT CTransStatusBar::WindowProc(HWND hRealWnd, UINT msg, WPARAM wp, LPARAM lp)
{
	switch (msg)
	{
	case SB_SETTEXT:
	case SB_SETTIPTEXT:
		if (m_bAllowTranslate)
		{
			LPCTSTR szOrgText = (LPCTSTR)lp;

			// only translate if changed
			if (szOrgText)
			{
				CString sText(szOrgText);

				if (TranslateText(sText))
				{
					CSubclassWnd::WindowProc(hRealWnd, msg, wp, (LPARAM)(LPCTSTR)sText);

					// restore text ptr
					lp = (LPARAM)szOrgText;
				}
			}
		}
		break;
	}

	return CTransWnd::WindowProc(hRealWnd, msg, wp, lp);
}

//////////////////

