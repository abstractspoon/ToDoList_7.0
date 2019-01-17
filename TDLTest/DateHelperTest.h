// DateHelperClass.h: interface for the CDateHelperClass class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DATEHELPERTEST_H__79858864_666A_44C5_954A_F15F4849D2D0__INCLUDED_)
#define AFX_DATEHELPERTEST_H__79858864_666A_44C5_954A_F15F4849D2D0__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "TDLTestBase.h"

class CDateHelperTest : public CTDLTestBase    
{
public:
	CDateHelperTest(const CTestUtils& utils);
	virtual ~CDateHelperTest();
	
	void Run();

protected:
	void RunDecodeOffsetTest();
	void RunDecodeRelativeDateTest();

};

#endif // !defined(AFX_DATEHELPERTEST_H__79858864_666A_44C5_954A_F15F4849D2D0__INCLUDED_)
