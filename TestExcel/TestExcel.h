// TestExcel.h : main header file for the TESTEXCEL application
//

#if !defined(AFX_TESTEXCEL_H__D0533EAE_A3F9_4213_AC49_803FA6199788__INCLUDED_)
#define AFX_TESTEXCEL_H__D0533EAE_A3F9_4213_AC49_803FA6199788__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CTestExcelApp:
// See TestExcel.cpp for the implementation of this class
//

class CTestExcelApp : public CWinApp
{
public:
	CTestExcelApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTestExcelApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CTestExcelApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TESTEXCEL_H__D0533EAE_A3F9_4213_AC49_803FA6199788__INCLUDED_)
