// TransTextMgr.h: interface for the CTransTextMgr class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_TRANSTEXTMGR_H__1F1C6D57_D2C6_4C06_9E7C_1E0C54B65D30__INCLUDED_)
#define AFX_TRANSTEXTMGR_H__1F1C6D57_D2C6_4C06_9E7C_1E0C54B65D30__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "hookwndmgr.h"

#include "..\Interfaces\itranstext.h"

#include <afxtempl.h>

class CXmlItem;
class CTransWnd;

struct DICTITEM
{
	DICTITEM();
	DICTITEM(const DICTITEM& di);
	DICTITEM(const CXmlItem* pXI);
	DICTITEM(const CString& sText);
	virtual ~DICTITEM();

	DICTITEM& operator= (const DICTITEM& di);

	BOOL Translate(CString& sText);
	BOOL Translate(CString& sText, HWND hWndRef, LPCTSTR szClassID = NULL);
	BOOL Translate(CString& sText, HMENU hMenu, int nMenuID);
	BOOL IsTranslated() const;

	BOOL ToXml(CXmlItem* pXI) const;
	BOOL FromXml(const CXmlItem* pXI);

	BOOL ToCsv(CStringArray& aTransLines, CStringArray& aNeedTransLines) const;
	BOOL FromCsv(const CStringArray& aLines, int& nLine);
	
	BOOL Merge(const DICTITEM& di);
	BOOL HasClassID(const CString& sClassID) const;
	BOOL Fixup();
	BOOL NeedFixup(const CString& sClassID, CString& sReplaceID, CString& sReplaceText) const;
	void ClearTextOut();

	const CString& GetTextIn() const { return m_sTextIn; }
	const CString& GetTextOut() const { return m_sTextOut; }
 	void SetTextIn(const CString& sText) { m_sTextIn = sText; }

	static void SetTranslationOption(ITT_TRANSLATEOPTION nOption);
	static ITT_TRANSLATEOPTION GetTranslationOption() { return s_nTranslationOption; }
	static BOOL WantAddToDictionary();
	static BOOL WantTranslateOnly();
	static BOOL WantUppercase();

protected:
	CString m_sTextIn;
	CString m_sTextOut;
	CString m_sClassID;

	static ITT_TRANSLATEOPTION s_nTranslationOption;

	// map alternatives by path
	CMapStringToString m_mapAlternatives;

	static void ToXml(CXmlItem* pXI, const DICTITEM& di);
	static int GetDlgCtrlID(HWND hWnd);
	static BOOL FromCsv(const CString& sLine, DICTITEM& di);

	static void FixupFormatString(CString& sFormat);

	BOOL Translate(CString& sText, const CString& sClassID);
};

class CTransDictionary : protected CMap<CString, LPCTSTR, DICTITEM*, DICTITEM*&>
{
public:
	CTransDictionary();
	~CTransDictionary();

	CString GetDictionaryPath() const { return m_sDictFile; }
	CString GetDictionaryVersion() const { return m_sDictVersion; }
	WORD GetDictionaryLanguageID() const { return m_wDictLanguageID; }

	BOOL SaveDictionary(LPCTSTR szAltPath = NULL, BOOL bForce = FALSE);
	BOOL SaveXmlDictionary(LPCTSTR szDictPath) const;
	BOOL SaveCsvDictionary(LPCTSTR szDictPath) const;

	BOOL LoadDictionary(LPCTSTR szDictPath);
	BOOL LoadXmlDictionary(LPCTSTR szDictPath);
	BOOL LoadCsvDictionary(LPCTSTR szDictPath);
	void LoadItem(const CXmlItem* pXI);
	BOOL LoadDictionaryItem(const CXmlItem* pXIDict);
	BOOL IsEmpty() const { return (GetCount() == 0); }

	BOOL Translate(CString& sText);
	BOOL Translate(CString& sText, HWND hWndRef, LPCTSTR szClassID = NULL);
	BOOL Translate(CString& sText, HMENU hMenu, int nMenuID);

	void DeleteDictionary();
	void FixupDictionary();
	BOOL CleanupDictionary(const CTransDictionary& tdMaster, CTransDictionary& tdRemoved);
	
	void IgnoreString(const CString& sText, BOOL bPrepare);
	BOOL WantIgnore(const CString& sText) const;
	
protected:
	CString m_sDictFile, m_sDictVersion;
	WORD m_wDictLanguageID;
	CMapStringToPtr m_mapStringIgnore;
	
protected:
	DICTITEM* GetDictItem(CString& sText, BOOL bAutoCreate = TRUE);
	BOOL HasDictItem(CString& sText) const;
	BOOL TranslateMenuShortcut(CString& sShortcut);

	static int CompareProc(const void* pFirst, const void* pSecond);
};

class CTransTextMgr : protected CHookWndMgr<CTransTextMgr>
{
	friend class CHookMgr<CTransTextMgr>; // so CHookMgr can access protected constructor
	friend class CHookWndMgr<CTransTextMgr>; // so CHookMgr can access protected constructor
	friend CTransWnd;

public:
	virtual ~CTransTextMgr();

	static BOOL Initialize(LPCTSTR szDictPath = _T(""), ITT_TRANSLATEOPTION nOption = ITTTO_TRANSLATEONLY);
	static void Release();

	static LPCTSTR GetDictionaryFile();
	static LPCTSTR GetDictionaryVersion();
	static void SetTranslationOption(ITT_TRANSLATEOPTION nOption);
	static ITT_TRANSLATEOPTION GetTranslationOption();
	
	static BOOL TranslateText(CString& sText, HWND hWndRef, LPCTSTR szClassID = NULL);
	static BOOL TranslateMenu(HMENU hMenu, HWND hWndRef = NULL, BOOL bRecursive = FALSE);

	static void UpdateMenu(HWND hWnd);
	static void UpdateMenu(HMENU hMenu);
	
	static void EnableTranslation(HWND hWnd, BOOL bEnable);
	static void EnableTranslation(HMENU hMenu, BOOL bEnable);
	static void EnableTranslation(UINT nMenuID, BOOL bEnable);

	static void IgnoreString(const CString& sText, BOOL bPrepare);

	static BOOL CleanupDictionary(LPCTSTR szMasterDictPath);
	static BOOL CleanupDictionary(LPCTSTR szDictPath, LPCTSTR szMasterDictPath);

protected:
	CTransDictionary m_dictionary;
	CMap<HWND, HWND&, void*, void*&> m_mapWndIgnore;
	CMap<UINT, UINT&, void*, void*&> m_mapMenuIgnore;

protected:
	CTransTextMgr();

	BOOL InitHooks(LPCTSTR szDictPath, ITT_TRANSLATEOPTION nOption);
	void ReleaseHooks();

	virtual CSubclassWnd* NewHookWnd(HWND hWnd, const CString& sClass, DWORD dwStyle) const;
	virtual BOOL WantHookWnd(HWND hWnd, UINT nMsg, WPARAM wp, LPARAM lp) const;
	virtual void PostHookWnd(HWND hWnd);
	virtual BOOL RemoveHookWnd(HWND hWnd);
	virtual BOOL OnCallWndProc(const MSG& msg);

	BOOL HandleInitMenuPopup(HWND hWnd, UINT nMsg, WPARAM wp, LPARAM lp);
	BOOL HandleTootipNeedText(HWND hWnd, UINT nMsg, WPARAM wp, LPARAM lp);
	BOOL HandleSetText(HWND hWnd, UINT nMsg, WPARAM wp, LPARAM lp);

	BOOL WantTranslation(HWND hWnd, UINT nMsg = 0) const;
	BOOL WantTranslation(UINT nMenuID) const;

	void AddToFile(const DICTITEM* pDI, CXmlItem* pXI);
};

#endif // !defined(AFX_TRANSTEXTMGR_H__1F1C6D57_D2C6_4C06_9E7C_1E0C54B65D30__INCLUDED_)
