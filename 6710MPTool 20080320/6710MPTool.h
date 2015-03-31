// 6710MPTool.h : main header file for the 6710MPTOOL application
//

#if !defined(AFX_6710MPTOOL_H__704CC10C_1794_48C5_8018_D04816185380__INCLUDED_)
#define AFX_6710MPTOOL_H__704CC10C_1794_48C5_8018_D04816185380__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CMy6710MPToolApp:
// See 6710MPTool.cpp for the implementation of this class
//

class CMy6710MPToolApp : public CWinApp
{
public:
	CMy6710MPToolApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMy6710MPToolApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CMy6710MPToolApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_6710MPTOOL_H__704CC10C_1794_48C5_8018_D04816185380__INCLUDED_)
