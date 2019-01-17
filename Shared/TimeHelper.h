// TimeHelper.h: interface for the TimeHelper class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_TIMEHELPER_H__BA0C1E67_FAAA_4E65_8EF3_01B011ACFBBC__INCLUDED_)
#define AFX_TIMEHELPER_H__BA0C1E67_FAAA_4E65_8EF3_01B011ACFBBC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <afxtempl.h>

enum 
{
	THU_MINS	= 'I',
	THU_HOURS	= 'H',
	THU_DAYS	= 'D',
	THU_WEEKS	= 'W',
	THU_MONTHS	= 'M',
	THU_YEARS	= 'Y',
};

enum THU_WORKDAYPERIOD
{
	THU_BEFORE,
	THU_MORNING,
	THU_LUNCH,
	THU_AFTERNOON,
	THU_AFTER,
};

class CTimeHelper  
{
public:
	CTimeHelper(); // uses statically defined hours and days
	CTimeHelper(double dHoursInDay, double dDaysInWeek);
	
	double GetTime(double dTime, int nFromUnits, int nToUnits) const;
	CString FormatTimeHMS(double dTime, int nUnitsFrom, BOOL bDecPlaces = TRUE) const;
	CString FormatTime(double dTime, int nUnits, int nDecPlaces) const;
	CString FormatTime(double dTime, int nDecPlaces) const;

	void CalculatePartWorkdays(const COleDateTime& dtStart, const COleDateTime& dtEnd,
								double& dPartStartDay, double& dPartEndDay, BOOL bInDays = TRUE) const;

	THU_WORKDAYPERIOD GetWorkdayPeriod(const COleDateTime& date) const;
	double GetStartOfWorkday(BOOL bInDays = TRUE) const;
	double GetStartOfWorkdayLunch(BOOL bInDays = TRUE) const;
	double GetEndOfWorkday(BOOL bInDays = TRUE) const;
	double GetEndOfWorkdayLunch(BOOL bInDays = TRUE) const;

	double GetHoursInOneDay() const { return m_dHours2Days; }
	double GetDaysInOneWeek() const { return m_dDaysToWeeks; }

public:
	static BOOL SetHoursInOneDay(double dHours);
	static BOOL SetDaysInOneWeek(double dDays);
	static void SetUnits(int nUnits, LPCTSTR szUnits);
	static void SetUnits(int nUnits, TCHAR cUnits);
	static TCHAR GetUnits(int nUnits);
	static BOOL IsValidUnit(int nUnits);
	static BOOL DecodeOffset(LPCTSTR szTime, double& dAmount, int& nUnits, BOOL bMustHaveSign = TRUE);

	static CString FormatClockTime(const COleDateTime& dtTime, BOOL bIncSeconds = FALSE, BOOL bISO = FALSE);
	static CString FormatClockTime(int nHour, int nMin, int nSec = 0, BOOL bIncSeconds = FALSE, BOOL bISO = FALSE);

protected:
	double m_dHours2Days, m_dDaysToWeeks;
	
protected:
	static double HOURS2DAYS, DAYS2WEEKS; // user definable
	static CMap<int, int, TCHAR, TCHAR&> MAPUNIT2CH; // user definable
	
protected:
	static double GetTimeOnly(double dDate);
	static double GetTimeOnly(const COleDateTime& date);
	static BOOL Compare(int nFromUnits, int nToUnits); // 0=same, -1=nFrom < nTo else 1
	static CString FormatTimeHMS(double dTime, int nUnits, int nLeftOverUnits, 
								double dLeftOverMultiplier, BOOL bDecPlaces);
};

#endif // !defined(AFX_TIMEHELPER_H__BA0C1E67_FAAA_4E65_8EF3_01B011ACFBBC__INCLUDED_)
