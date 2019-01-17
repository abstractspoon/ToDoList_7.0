#pragma once

#include "..\shared\autocombobox.h"
#include "..\shared\fileedit.h"
#include "..\shared\iconcache.h"

// CFileComboBox

class CFileComboBox : public CAutoComboBox
{
public:
	CFileComboBox(int nEditStyle = FES_COMBOSTYLEBTN);
	virtual ~CFileComboBox();

	void SetCurrentFolder(LPCTSTR szFolder) { m_fileEdit.SetCurrentFolder(szFolder); }
	int GetFileList(CStringArray& aFiles);
	int SetFileList(const CStringArray& aFiles);
	int AddFile(const CString& sFile) { return AddUniqueItem(sFile); }
	int AddFiles(const CStringArray& aFiles);
	
protected:
	CFileEdit m_fileEdit;
	CIconCache m_imageIcons;

protected:
	void PreSubclassWindow();
	BOOL PreCreateWindow(CREATESTRUCT& cs);

protected:
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg LRESULT OnFileEditBrowseChange(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnFileEditGetFileIcon(WPARAM wp, LPARAM lp);
	afx_msg BOOL OnSelChange();

	DECLARE_MESSAGE_MAP()

	// pseudo message handler
	LRESULT OnEditboxMessage(UINT msg, WPARAM wp, LPARAM lp);

protected:
	void ResizeEdit();
	BOOL InitFileEdit();
	void DrawItemText(CDC& dc, const CRect& rect, int nItem, UINT nItemState,
						DWORD dwItemData, const CString& sItem, BOOL bList);	
	BOOL HasIcon() const { return TRUE; }
	int GetMaxDropWidth() const { return 400; }
	UINT GetDrawEllipsis() const { return DT_PATH_ELLIPSIS; }
	BOOL DeleteLBItem(int nItem);

};


