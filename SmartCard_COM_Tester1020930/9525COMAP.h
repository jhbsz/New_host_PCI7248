// 9525COMAP.h : main header file for the 9525COMAP application
//

#if !defined(AFX_9525COMAP_H__0DA0EE60_F27F_469D_AB5E_2EA129B5319C__INCLUDED_)
#define AFX_9525COMAP_H__0DA0EE60_F27F_469D_AB5E_2EA129B5319C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

//define WM for USB(VB) and RS232(VC)
#define WM_READY_TIMER			0x10
#define WM_COM_START_TEST		WM_USER + 0x700
#define WM_COM_TEST_PASS		WM_USER + 0x710
#define WM_COM_TEST_FAIL		WM_USER + 0x720
#define WM_COM_DEVICE_UNKNOWN	WM_USER + 0x730
#define WM_COM_CLOSE			WM_USER + 0x740
#define WM_FT_READY				WM_USER + 0x800

/////////////////////////////////////////////////////////////////////////////
// CMy9525COMAPApp:
// See 9525COMAP.cpp for the implementation of this class
//

class CMy9525COMAPApp : public CWinApp
{
public:
	CMy9525COMAPApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMy9525COMAPApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CMy9525COMAPApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_9525COMAP_H__0DA0EE60_F27F_469D_AB5E_2EA129B5319C__INCLUDED_)
