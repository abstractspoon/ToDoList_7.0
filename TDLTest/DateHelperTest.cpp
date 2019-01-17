// DateHelperClass.cpp: implementation of the DateHelperTest class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DateHelperTest.h"

#include "..\Shared\DateHelper.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CDateHelperTest::CDateHelperTest(const CTestUtils& utils) : CTDLTestBase(utils)
{

}

CDateHelperTest::~CDateHelperTest()
{

}

void CDateHelperTest::Run()
{
	RunDecodeOffsetTest();
	RunDecodeRelativeDateTest();
}

void CDateHelperTest::RunDecodeOffsetTest()
{
	BeginTest(_T("CDateHelper::DecodeOffset"));

	double dAmount = 0.0;
	int nUnits = 0;
	
	// Positive offset with sign -------------------------------------------------------

	ExpectTrue(CDateHelper::DecodeOffset(_T("+5D"), dAmount, nUnits, FALSE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("+5D"), dAmount, nUnits, TRUE));
	ExpectEQ(dAmount, 5.0);
	ExpectEQ(nUnits, DHU_DAYS);
	
	ExpectTrue(CDateHelper::DecodeOffset(_T("+8W"), dAmount, nUnits, FALSE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("+8W"), dAmount, nUnits, TRUE));
	ExpectEQ(dAmount, 8.0);
	ExpectEQ(nUnits, DHU_WEEKS);
	
	ExpectTrue(CDateHelper::DecodeOffset(_T("+10M"), dAmount, nUnits, FALSE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("+10M"), dAmount, nUnits, TRUE));
	ExpectEQ(dAmount, 10.0);
	ExpectEQ(nUnits, DHU_MONTHS);
	
	ExpectTrue(CDateHelper::DecodeOffset(_T("+2Y"), dAmount, nUnits, FALSE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("+2Y"), dAmount, nUnits, TRUE));
	ExpectEQ(dAmount, 2.0);
	ExpectEQ(nUnits, DHU_YEARS);
	
	// Positive offset without sign ---------------------------------------------------

	ExpectFalse(CDateHelper::DecodeOffset(_T("5D"), dAmount, nUnits, TRUE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("5D"), dAmount, nUnits, FALSE));
	ExpectEQ(dAmount, 5.0);
	ExpectEQ(nUnits, DHU_DAYS);
	
	ExpectFalse(CDateHelper::DecodeOffset(_T("8W"), dAmount, nUnits, TRUE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("8W"), dAmount, nUnits, FALSE));
	ExpectEQ(dAmount, 8.0);
	ExpectEQ(nUnits, DHU_WEEKS);
	
	ExpectFalse(CDateHelper::DecodeOffset(_T("10M"), dAmount, nUnits, TRUE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("10M"), dAmount, nUnits, FALSE));
	ExpectEQ(dAmount, 10.0);
	ExpectEQ(nUnits, DHU_MONTHS);
	
	ExpectFalse(CDateHelper::DecodeOffset(_T("2Y"), dAmount, nUnits, TRUE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("2Y"), dAmount, nUnits, FALSE));
	ExpectEQ(dAmount, 2.0);
	ExpectEQ(nUnits, DHU_YEARS);
	
	// Negative offset with sign ------------------------------------------------------

	ExpectTrue(CDateHelper::DecodeOffset(_T("-5D"), dAmount, nUnits, FALSE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("-5D"), dAmount, nUnits, TRUE));
	ExpectEQ(dAmount, -5.0);
	ExpectEQ(nUnits, DHU_DAYS);
	
	ExpectTrue(CDateHelper::DecodeOffset(_T("-8W"), dAmount, nUnits, FALSE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("-8W"), dAmount, nUnits, TRUE));
	ExpectEQ(dAmount, -8.0);
	ExpectEQ(nUnits, DHU_WEEKS);
	
	ExpectTrue(CDateHelper::DecodeOffset(_T("-10M"), dAmount, nUnits, FALSE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("-10M"), dAmount, nUnits, TRUE));
	ExpectEQ(dAmount, -10.0);
	ExpectEQ(nUnits, DHU_MONTHS);
	
	ExpectTrue(CDateHelper::DecodeOffset(_T("-2Y"), dAmount, nUnits, FALSE));
	ExpectTrue(CDateHelper::DecodeOffset(_T("-2Y"), dAmount, nUnits, TRUE));
	ExpectEQ(dAmount, -2.0);
	ExpectEQ(nUnits, DHU_YEARS);
	
	// -------------------------------------------------------------------------

	EndTest();
}

void CDateHelperTest::RunDecodeRelativeDateTest()
{
	BeginTest(_T("CDateHelper::DecodeRelativeDate"));
	
	const COleDateTime TODAY	= CDateHelper::GetDate(DHD_TODAY);
	const COleDateTime ENDWEEK	= CDateHelper::GetDate(DHD_ENDTHISWEEK);
	const COleDateTime ENDMONTH = CDateHelper::GetDate(DHD_ENDTHISMONTH);
	const COleDateTime ENDYEAR	= CDateHelper::GetDate(DHD_ENDTHISYEAR);

	COleDateTime date;	

	// No Offset ---------------------------------------------------------------

	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T"), date, FALSE));
	ExpectEQ(TODAY, date);

	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("W"), date, FALSE));
	ExpectEQ(ENDWEEK, date);
	
	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("M"), date, FALSE));
	ExpectEQ(ENDMONTH, date);
	
	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("Y"), date, FALSE));
	ExpectEQ(ENDYEAR, date);
		
	// Positive offset with prefixed date --------------------------------------

	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T+5D"), date, FALSE));
	ExpectEQ(TODAY.m_dt + 5, date.m_dt);
	
	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T+8W"), date, FALSE));
	ExpectEQ(TODAY.GetDayOfWeek(), date.GetDayOfWeek());
	ExpectEQ(TODAY.m_dt + 56, date.m_dt);
	
	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T+10M"), date, FALSE));
	ExpectEQ(TODAY.GetDay(), date.GetDay());

	if (TODAY.GetMonth() > 1)
	{
		ExpectEQ(TODAY.GetYear() + 1, date.GetYear());
		ExpectEQ(TODAY.GetMonth() - 2, date.GetMonth());
	}
	else
	{
		ExpectEQ(TODAY.GetYear(), date.GetYear());
		ExpectEQ(TODAY.GetMonth() + 10, date.GetMonth());
	}
	
	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T+2Y"), date, FALSE));
	ExpectEQ(TODAY.GetDay(), date.GetDay());
	ExpectEQ(TODAY.GetMonth(), date.GetMonth());
	ExpectEQ(TODAY.GetYear() + 2, date.GetYear());
	
	// Positive offset without prefixed date ------------------------------------

	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("+5D"), date, FALSE));
	ExpectEQ(TODAY.m_dt + 5, date.m_dt);

	// Positive offset without prefixed sign or date-----------------------------

	ExpectFalse(CDateHelper::DecodeRelativeDate(_T("5D"), date, FALSE));
	
	// Negative offset without prefixed date ------------------------------------

	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("-5D"), date, FALSE));
	ExpectEQ(TODAY.m_dt - 5, date.m_dt);

	// Negative offset with prefixed date ---------------------------------------

	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T-5D"), date, FALSE));
	ExpectEQ(TODAY.m_dt - 5, date.m_dt);
	
	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T-8W"), date, FALSE));
	ExpectEQ(TODAY.GetDayOfWeek(), date.GetDayOfWeek());
	ExpectEQ(TODAY.m_dt - 56, date.m_dt);
	
	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T-10M"), date, FALSE));
	ExpectEQ(TODAY.GetDay(), date.GetDay());

	if (TODAY.GetMonth() < 10)
	{
		ExpectEQ(TODAY.GetYear() - 1, date.GetYear());
		ExpectEQ(TODAY.GetMonth() + 2, date.GetMonth());
	}
	else
	{
		ExpectEQ(TODAY.GetYear(), date.GetYear());
		ExpectEQ(TODAY.GetMonth() - 10, date.GetMonth());
	}
	
	ExpectTrue(CDateHelper::DecodeRelativeDate(_T("T-2Y"), date, FALSE));
	ExpectEQ(TODAY.GetDay(), date.GetDay());
	ExpectEQ(TODAY.GetMonth(), date.GetMonth());
	ExpectEQ(TODAY.GetYear() - 2, date.GetYear());
		
	// -------------------------------------------------------------------------

	EndTest();
}
