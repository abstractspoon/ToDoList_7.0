// ToDoitem.h: interface for the CToDoitem class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_TODOITEM_H__02C3C360_45AB_45DC_B1BF_BCBEA472F0C7__INCLUDED_)
#define AFX_TODOITEM_H__02C3C360_45AB_45DC_B1BF_BCBEA472F0C7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "tdcstruct.h"
#include "..\shared\binarydata.h"

#include <afxtempl.h>

class TODOITEM
{
	friend class CToDoCtrlData;
	
public:
	TODOITEM(LPCTSTR szTitle = NULL, LPCTSTR szComments = NULL); 
	TODOITEM(const TODOITEM& tdi); 
	
	const TODOITEM& operator=(const TODOITEM& tdi); 
	
	// helpers
	BOOL HasLastMod() const;
	BOOL HasCreation() const;
	BOOL HasStart() const;
	BOOL HasStartTime() const;
	BOOL HasDue() const;
	BOOL HasDueTime() const;
	BOOL IsDone() const;
	BOOL HasDoneTime() const;
	
	void ClearStart();
	void ClearDue();
	void ClearDone();
	
	BOOL IsDue() const;
	BOOL IsDue(const COleDateTime& dateDueBy) const;
	
	void SetModified();
	BOOL IsReference() const;
	
	CString GetCategory(int nCat) const;
	CString GetAllocTo(int nAllocTo) const;
	CString GetDependency(int Depends) const;
	CString GetTag(int nTag) const;
	CString GetFileRef(int nFile) const;

	BOOL GetNextOccurence(COleDateTime& dtNext, BOOL& bDue);
	BOOL IsRecurring() const;
	BOOL CanRecur() const;
	
	BOOL IsRecentlyEdited() const; // 1 hour default
	COleDateTimeSpan GetRemainingDueTime() const; // in days
	
	BOOL HasCustomAttribValue(const CString& sAttribID) const;
	CString GetCustomAttribValue(const CString& sAttribID) const;
	
	BOOL GetCustomAttribCalcValue(const CString& sAttribID, double& dValue) const;
	void SetCustomAttribCalcValue(const CString& sAttribID, double dValue) const;
	
	TCHAR GetTimeUnits(BOOL bTimeEst) const;
	double GetTime(BOOL bTimeEst, BOOL bCalc, BOOL bWeighted = FALSE) const;
	
	// only applies to dependencies within this tasklist
	BOOL RemoveLocalDependency(DWORD dwDependID);
	BOOL IsLocallyDependentOn(DWORD dwDependID) const;
	int GetLocalDependencies(CDWordArray& aDependIDs) const;
	DWORD GetFirstLocalDependency() const;
	
	static COleDateTimeSpan GetRemainingDueTime(const COleDateTime& date); // in days
	static BOOL HasTime(const COleDateTime& date);
	static BOOL ParseTaskLink(const CString& sLink, DWORD& dwTaskID, CString& sFile);
	static CString FormatTaskDependency(DWORD dwTaskID, const CString& sFile = _T(""));
	static void SetRecentlyModifiedPeriod(double dDays);
	
	// Attributes
	CString sTitle;
	CString sComments;
	CString sCommentsTypeID;
	CString sAllocBy;
	CString sStatus;
	CString sCreatedBy;
	CString sExternalID;
	CString sVersion;
	CString sIcon;
	
	CStringArray aAllocTo;
	CStringArray aCategories;
	CStringArray aDependencies;
	CStringArray aTags;
	CStringArray aFileRefs;
	
	COleDateTime dateStart, dateDue, dateDone, dateCreated, dateLastMod;
	
	int nPriority;
	int nPercentDone;
	int nRisk;
	
	double dCost;
	double dTimeEstimate, dTimeSpent;
	
	COLORREF color;
	BOOL bFlagged;
	TCHAR nTimeEstUnits, nTimeSpentUnits;
	CBinaryData customComments;
	TDIRECURRENCE trRecurrence;
	DWORD dwTaskRefID;
	
	// meta-data for 3rd-party applications only
	CMapStringToString mapMetaData; 
	
	// custom attributes
	CMapStringToString mapCustomData;
	
private:
	void ResetCalcs(TDC_ATTRIBUTE nAttrib = TDCA_ALL, BOOL bIncCustom = TRUE) const; 
	BOOL AttribNeedsRecalc(DWORD dwAttrib) const;
	void SetAttribNeedsRecalc(DWORD dwAttrib, BOOL bSet = TRUE) const;
	BOOL CustomAttribNeedsRecalc(const CString& sAttribID) const;
	void ResetCustomAttribCalc(const CString& sAttribID = _T("")) const;
	int FindLocalDependency(DWORD dwDependID, int nSearchFrom = 0) const;
	
	// cached calculations for drawing optimization
	// mutable so that they can be updated in const methods
	mutable DWORD dwCalculated;
	mutable int nCalcPriority;
	mutable int nCalcPriorityIncDue;
	mutable int nCalcPercent;
	mutable int nCalcRisk;
	mutable double dCalcTimeEstimate, dCalcWeightedTimeEstimate, dCalcTimeSpent;
	mutable double dCalcCost;
	mutable COleDateTime dateCalcDue, dateCalcStart;
	mutable BOOL bGoodAsDone;
	mutable int nSubtasksCount, nSubtasksDone;
	mutable CString sPath, sPosition;
	mutable CString sCategoryList, sAllocToList, sTagList;
	mutable CMap<CString, LPCTSTR, double, double&> mapCustomCalcs;
	
	static double NULL_VALUE;
	static COleDateTimeSpan s_dtsRecentModPeriod;
	
};

class TODOSTRUCTURE
{
	friend class CToDoCtrlStructure;
	
public:
	TODOSTRUCTURE() : m_dwID(0), m_pTDSParent(NULL) {}
	TODOSTRUCTURE(DWORD dwID);
	~TODOSTRUCTURE();
	
	DWORD GetTaskID() const { return m_dwID; }
	DWORD GetSubTaskID(int nPos) const;
	
	int GetSubTaskPosition(DWORD dwID) const;
	int GetPosition() const;
	
	TODOSTRUCTURE* GetParentTask() const;
	DWORD GetParentTaskID() const;
	BOOL HasSameParent(const TODOSTRUCTURE* pTDS) const;
	
	BOOL ParentIsRoot() const { return (GetParentTaskID() == 0); }
	BOOL IsRoot() const { return (GetTaskID() == 0); }
	
	DWORD GetPreviousSubTaskID(int nPos);
	
	int GetSubTaskCount() const { return m_aSubTasks.GetSize(); }
	BOOL HasSubTasks() const { return GetSubTaskCount() > 0; }
	int GetLeafCount() const;
	
	TODOSTRUCTURE* GetSubTask(int nPos) const;
	int GetSubTaskPosition(const TODOSTRUCTURE* pTDS) const;
	
	void DeleteAll() { CleanUp(); }
	
	int MoveSubTask(int nPos, TODOSTRUCTURE* pTDSDestParent, int nDestPos);
	
protected:
	DWORD m_dwID;
	TODOSTRUCTURE* m_pTDSParent;
	CArray<TODOSTRUCTURE*, TODOSTRUCTURE*&> m_aSubTasks; 
	
protected:
	TODOSTRUCTURE(const TODOSTRUCTURE& tds);
	const TODOSTRUCTURE& operator=(const TODOSTRUCTURE& tds); 
	
	void CleanUp();
	BOOL DeleteSubTask(int nPos);
	BOOL InsertSubTask(TODOSTRUCTURE* pTDS, int nPos);
	BOOL AddSubTask(TODOSTRUCTURE* pTDS);
	TODOSTRUCTURE* AddSubTask(DWORD dwID);
};

class CToDoCtrlStructure : public TODOSTRUCTURE
{
public:
	CToDoCtrlStructure() : TODOSTRUCTURE() {}
	CToDoCtrlStructure(const CToDoCtrlStructure& tds); 
	~CToDoCtrlStructure();
	
	TODOSTRUCTURE* GetParentTask(DWORD dwID) const;
	DWORD GetParentTaskID(DWORD dwID) const;
	
	DWORD GetPreviousTaskID(DWORD dwID) const;
	
	BOOL DeleteTask(DWORD dwID);
	
	TODOSTRUCTURE* AddTask(DWORD dwID, TODOSTRUCTURE* pTDSParent);
	BOOL InsertTask(const TODOSTRUCTURE& tds, TODOSTRUCTURE* pTDSParent, int nPos);
	BOOL InsertTask(DWORD dwID, TODOSTRUCTURE* pTDSParent, int nPos);
	BOOL InsertTask(TODOSTRUCTURE* pTDS, TODOSTRUCTURE* pTDSParent, int nPos);
	
	TODOSTRUCTURE* FindTask(DWORD dwID) const;
	BOOL FindTask(DWORD dwID, TODOSTRUCTURE*& pTDSParent, int& nPos) const;
	
protected:
	CMap<DWORD, DWORD, TODOSTRUCTURE*, TODOSTRUCTURE*&> m_mapStructure;
	
protected:
	void BuildMap();
	void AddToMap(const TODOSTRUCTURE* pTDS);
	void RemoveFromMap(const TODOSTRUCTURE* pTDS);
	const CToDoCtrlStructure& operator=(const CToDoCtrlStructure& tds); 
	
};


#endif // !defined(AFX_TODOITEM_H__02C3C360_45AB_45DC_B1BF_BCBEA472F0C7__INCLUDED_)
