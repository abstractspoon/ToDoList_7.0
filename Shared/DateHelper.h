// DateHelper.h: interface for the CDateHelper class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DATEHELPER_H__2A4E63F6_A106_4295_BCBA_06D03CD67AE7__INCLUDED_)
#define AFX_DATEHELPER_H__2A4E63F6_A106_4295_BCBA_06D03CD67AE7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

enum DH_DATE
{
	DHD_TODAY,
	DHD_TOMORROW,
	DHD_BEGINTHISWEEK, // DHD_ENDTHIDWEEK - 6
	DHD_ENDTHISWEEK,   
	DHD_ENDNEXTWEEK,   // DHD_ENDTHISWEEK + 7
	DHD_ENDTHISMONTH,  // beginning of next month - 1
	DHD_ENDNEXTMONTH,  // get's trickier :)
	DHD_ENDTHISYEAR,
	DHD_ENDNEXTYEAR,
};

enum 
{
	DHU_WEEKDAYS	= 'K',
	DHU_DAYS		= 'D',
	DHU_WEEKS		= 'W',
	DHU_MONTHS		= 'M',
	DHU_YEARS		= 'Y',
};

enum
{
	DHFD_ISO	= 0x0001,
	DHFD_DOW	= 0x0002,
	DHFD_TIME	= 0x0004,
	DHFD_NOSEC	= 0x0008,
};

enum // days of week
{
	DHW_SUNDAY		= 0X01,
	DHW_MONDAY		= 0X02,
	DHW_TUESDAY		= 0X04,
	DHW_WEDNESDAY	= 0X08,
	DHW_THURSDAY	= 0X10,
	DHW_FRIDAY		= 0X20,
	DHW_SATURDAY	= 0X40,

	DHW_EVERYDAY	= 0x7F
};

// Note: 
// 1 <= nMonth <= 12
// 1 <= nDay <= 31

typedef __int64 time64_t;

class CDateHelper  
{
public:
	static BOOL IsDateSet(const COleDateTime& date);
	static void ClearDate(COleDateTime& date);

	static BOOL IsValidDayInMonth(int nDay, int nMonth, int nYear);
	static BOOL IsValidDayOfMonth(int nDOW, int nWhich, int nMonth);

	static int CalcDaysFromTo(const COleDateTime& dateFrom, const COleDateTime& dateTo, BOOL bInclusive);
	static int CalcDaysFromTo(const COleDateTime& dateFrom, DH_DATE nTo, BOOL bInclusive);
	static int CalcDaysFromTo(DH_DATE nFrom, DH_DATE nTo, BOOL bInclusive);
	static int CalcWeekdaysFromTo(const COleDateTime& dateFrom, const COleDateTime& dateTo, BOOL bInclusive);

	static double GetDate(DH_DATE nDate); // 12am
	static BOOL OffsetDate(COleDateTime& date, int nAmount, int nUnits, BOOL bForceWeekday = FALSE);

	static CString FormatDate(const COleDateTime& date, DWORD dwFlags = 0);
	static CString FormatCurrentDate(DWORD dwFlags = 0);

	static int GetISODayOfWeek(const COleDateTime& date);   // 1=Mon, 2=Tue, ..., 7=Sun

	// DOW = 'day of week'
	static BOOL FormatDate(const COleDateTime& date, DWORD dwFlags, CString& sDate, CString& sTime, CString& sDow);
	static BOOL FormatCurrentDate(DWORD dwFlags, CString& sDate, CString& sTime, CString& sDow);

	static BOOL DecodeDate(const CString& sDate, COleDateTime& date, BOOL bAndTime = FALSE);
	static BOOL DecodeDate(const CString& sDate, double& date, BOOL bAndTime = FALSE);
#if _MSC_VER < 1400
	static BOOL DecodeDate(const CString& sDate, time_t& date, BOOL bAndTime = FALSE);
#endif
	static BOOL DecodeDate(const CString& sDate, time64_t& date, BOOL bAndTime = FALSE);

	static BOOL DecodeOffset(LPCTSTR szDate, double& dAmount, int& nUnits, BOOL bMustHaveSign = TRUE);
	static BOOL DecodeRelativeDate(LPCTSTR szDate, COleDateTime& date, BOOL bForceWeekday, BOOL bMustHaveSign = TRUE);

	static int GetFirstDayOfWeek();
	static int GetLastDayOfWeek();
	static int GetNextDayOfWeek(int nDOW);
	static int GetDaysInMonth(int nMonth, int nYear); 
	static int GetDaysInMonth(const COleDateTime& date); 
	static COleDateTime GetEndOfPreviousDay(const COleDateTime& date);

	static void GetNextMonth(int& nMonth, int& nYear, BOOL bNext = TRUE);
	static void IncrementMonth(int& nMonth, int& nYear, int nBy = 1);
	static void IncrementMonth(SYSTEMTIME& st, int nBy = 1);

	static COleDateTime CalcDate(int nDOW, int nWhich, int nMonth, int nYear);
	static int CalcDayOfMonth(int nDOW, int nWhich, int nMonth, int nYear);
	static int CalcWeekofYear(const COleDateTime& date);

	static CString GetWeekdayName(int nWeekday, BOOL bShort = FALSE); // 1-7, sun-sat
	static CString GetMonthName(int nMonth, BOOL bShort = FALSE); // 1-12, jan-nov
	static void GetWeekdayNames(BOOL bShort, CStringArray& aWeekDays); // sun-sat
	static void GetMonthNames(BOOL bShort, CStringArray& aMonths); // jan-dec
	static int CalcLongestWeekdayName(CDC* pDC, BOOL bShort = FALSE);

	static BOOL IsLeapYear(const COleDateTime& date = COleDateTime::GetCurrentTime());
	static BOOL IsLeapYear(int nYear);

	static void SplitDate(const COleDateTime& date, double& dDateOnly, double& dTimeOnly);
	static COleDateTime MakeDate(const COleDateTime& dtDateOnly, const COleDateTime& dtTimeOnly);

	static BOOL DateHasTime(const COleDateTime& date);
	static COleDateTime GetTimeOnly(const COleDateTime& date);
	static COleDateTime GetDateOnly(const COleDateTime& date);

	static void SetWeekendDays(DWORD dwDays = DHW_SATURDAY | DHW_SUNDAY);
	static DWORD GetWeekdays();

	static BOOL IsWeekend(int nDOW);
	static BOOL IsWeekend(const COleDateTime& date);
	static BOOL IsWeekend(double dDate);
	static int GetWeekendDuration();
	static BOOL HasWeekend();
	static BOOL MakeWeekday(COleDateTime& date, BOOL bForwards = TRUE);
	static COleDateTime ToWeekday(const COleDateTime& date, BOOL bForwards = TRUE);

	static BOOL GetTimeT(const COleDateTime& date, time_t& timeT);
	static BOOL GetTimeT64(const COleDateTime& date, time64_t& timeT);
	static COleDateTime GetDate(time64_t date);

	static COleDateTime GetNearestYear(const COleDateTime& date, BOOL bEnd);
	static COleDateTime GetNearestHalfYear(const COleDateTime& date, BOOL bEnd);
	static COleDateTime GetNearestQuarter(const COleDateTime& date, BOOL bEnd);
	static COleDateTime GetNearestMonth(const COleDateTime& date, BOOL bEnd);
	static COleDateTime GetNearestWeek(const COleDateTime& date, BOOL bEnd);
	static COleDateTime GetNearestDay(const COleDateTime& date, BOOL bEnd);
	static COleDateTime GetNearestHalfDay(const COleDateTime& date, BOOL bEnd);
	static COleDateTime GetNearestHour(const COleDateTime& date, BOOL bEnd);

	static BOOL Min(COleDateTime& date, const COleDateTime& dtOther);
	static BOOL Max(COleDateTime& date, const COleDateTime& dtOther);

protected:
	static DWORD s_dwWeekend; 

protected:
	static BOOL DecodeISODate(const CString& sDate, COleDateTime& date, BOOL bAndTime = FALSE);
	static BOOL DecodeLocalShortDate(const CString& sDate, COleDateTime& date);
	static BOOL GetTimeT(time64_t date, time_t& timeT);
	static BOOL IsValidUnit(int nUnits);
	static BOOL DecodeOffsetEx(LPCTSTR szDate, double& dAmount, int& nUnits, int nDefUnits, BOOL bMustHaveSign);

	// Copyright (c) Robert Walker, support@tunesmithy.co.uk
	static void T64ToFileTime(time64_t *pt, FILETIME *pft);
	static void FileTimeToT64(FILETIME *pft, time64_t *pt);
	static void SystemTimeToT64(SYSTEMTIME *pst, time64_t *pt);
	static void T64ToSystemTime(time64_t *pt, SYSTEMTIME *pst);
};

#endif // !defined(AFX_DATEHELPER_H__2A4E63F6_A106_4295_BCBA_06D03CD67AE7__INCLUDED_)
