#if !defined(AFX_TDCMSG_H__5951FDE6_508A_4A9D_A55D_D16EB026AEF7__INCLUDED_)
#define AFX_TDCMSG_H__5951FDE6_508A_4A9D_A55D_D16EB026AEF7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// tdcmsg.h : header file
//
static LPCTSTR TDL_PROTOCOL = _T("tdl://");
static LPCTSTR TDL_EXTENSION = _T(".tdl");

// notification messages
const UINT WM_TDCN_MODIFY					= ::RegisterWindowMessage(_T("WM_TDCN_MODIFY")); // lParam == <TDC_ATTRIBUTE>
const UINT WM_TDCN_SORT						= ::RegisterWindowMessage(_T("WM_TDCN_SORT")); 
const UINT WM_TDCN_COMMENTSCHANGE			= ::RegisterWindowMessage(_T("WM_TDCN_COMMENTSCHANGE"));
const UINT WM_TDCN_COMMENTSKILLFOCUS		= ::RegisterWindowMessage(_T("WM_TDCN_COMMENTSKILLFOCUS"));
const UINT WM_TDCN_TIMETRACK				= ::RegisterWindowMessage(_T("WM_TDCN_TIMETRACK")); // lParam = 0/1 => stop/start
const UINT WM_TDCN_VIEWPRECHANGE			= ::RegisterWindowMessage(_T("WM_TDCN_VIEWPRECHANGE"));
const UINT WM_TDCN_VIEWPOSTCHANGE			= ::RegisterWindowMessage(_T("WM_TDCN_VIEWPOSTCHANGE"));
const UINT WM_TDCN_SELECTIONCHANGE			= ::RegisterWindowMessage(_T("WM_TDCN_SELECTIONCHANGE"));
const UINT WM_TDCN_SCROLLCHANGE				= ::RegisterWindowMessage(_T("WM_TDCN_SCROLLCHANGE"));
const UINT WM_TDCN_RECREATERECURRINGTASK	= ::RegisterWindowMessage(_T("WM_TDCN_RECREATERECURRINGTASK"));
const UINT WM_TDCN_DOUBLECLKREMINDERCOL		= ::RegisterWindowMessage(_T("WM_TDCN_DOUBLECLKREMINDERCOL"));
const UINT WM_TDCN_COLUMNEDITCLICK			= ::RegisterWindowMessage(_T("WM_TDCN_COLUMNEDITCLICK")); // lParam == <TDC_COLUMN>

// from the filterbar
const UINT WM_FBN_FILTERCHNG				= ::RegisterWindowMessage(_T("WM_FBN_FILTERCHNG")); 

// sent when one of the auto dropdown lists is changed
const UINT WM_TDCN_LISTCHANGE				= ::RegisterWindowMessage(_T("WM_TDCN_LISTCHANGE")); // lParam == <TDC_ATTRIBUTE>

// request messages
const UINT WM_TDCM_GETCLIPBOARD				= ::RegisterWindowMessage(_T("WM_TDCM_GETCLIPBOARD")); // lParam == match hwnd
const UINT WM_TDCM_HASCLIPBOARD				= ::RegisterWindowMessage(_T("WM_TDCM_HASCLIPBOARD")); // lParam == match hwnd
const UINT WM_TDCM_TASKISDONE				= ::RegisterWindowMessage(_T("WM_TDCM_TASKISDONE")); // format as WM_TDCM_TASKLINK
const UINT WM_TDCM_TASKHASREMINDER			= ::RegisterWindowMessage(_T("WM_TDCM_TASKHASREMINDER")); // wParam = TaskID, lParam = TDC* 

// instruction messages
// sent when a task outside 'this' todoctrl needs displaying
const UINT WM_TDCM_TASKLINK					= ::RegisterWindowMessage(_T("WM_TDCM_TASKLINK")); // wParam = taskID, lParam = taskfile
const UINT WM_TDCM_FAILEDLINK				= ::RegisterWindowMessage(_T("WM_TDCM_FAILEDLINK")); // wParam = 0, lParam = url
const UINT WM_TDCM_LENGTHYOPERATION			= ::RegisterWindowMessage(_T("WM_TDCM_LENGTHYOPERATION")); // wParam = start/stop, lParam = text to display

// find tasks dialog
const UINT WM_FTD_FIND						= ::RegisterWindowMessage(_T("WM_FTD_FIND"));
const UINT WM_FTD_SELECTRESULT				= ::RegisterWindowMessage(_T("WM_FTD_SELECTRESULT"));
const UINT WM_FTD_SELECTALL					= ::RegisterWindowMessage(_T("WM_FTD_SELECTALL"));
const UINT WM_FTD_CLOSE						= ::RegisterWindowMessage(_T("WM_FTD_CLOSE"));
const UINT WM_FTD_APPLYASFILTER				= ::RegisterWindowMessage(_T("WM_FTD_APPLYASFILTER"));
const UINT WM_FTD_ADDSEARCH					= ::RegisterWindowMessage(_T("WM_FTD_ADDSEARCH"));
const UINT WM_FTD_SAVESEARCH				= ::RegisterWindowMessage(_T("WM_FTD_SAVESEARCH"));
const UINT WM_FTD_DELETESEARCH				= ::RegisterWindowMessage(_T("WM_FTD_DELETESEARCH"));
const UINT WM_FTD_GETLISTITEMS				= ::RegisterWindowMessage(_T("WM_FTD_GETLISTITEMS")); // wParam = <TDC_ATTRIBUTE>, lParam = CStringArray*

/////////////////////////////////////////////////////////////////////////////


#endif // AFX_TDCMSG_H__5951FDE6_508A_4A9D_A55D_D16EB026AEF7__INCLUDED_