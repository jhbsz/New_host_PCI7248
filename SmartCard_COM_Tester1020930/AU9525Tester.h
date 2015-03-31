//{{AFX_INCLUDES()

//}}AFX_INCLUDES
#if !defined(AFX_AU9525TESTER_H__ADB3F256_C01A_49DD_A30C_C41B56B59AC3__INCLUDED_)
#define AFX_AU9525TESTER_H__ADB3F256_C01A_49DD_A30C_C41B56B59AC3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AU9525Tester.h : header file
//
#include "9525COMAPDlg.h"
#include "mscomm.h"
#include "labelcontrol.h"
/////////////////////////////////////////////////////////////////////////////
// AU9525Tester dialog

class AU9525Tester : public CDialog
{
// Construction
public:
	virtual void TestInit2();
	static void TestInitial();
	static void StartTest();
	AU9525Tester(CWnd* pParent = NULL);   // standard constructor
    CMy9525COMAPDlg AU9525COMAdlg; 
// Dialog Data
	//{{AFX_DATA(AU9525Tester)
	enum { IDD = IDD_AU9525TESTER_DIALOG };
	CMSComm	m_Comm;
	CLabelControl	m_Slot0;
	CLabelControl	m_Slot1;
	CLabelControl	m_TestResult;
	CLabelControl	m_UnknowDevice;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(AU9525Tester)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(AU9525Tester)
	virtual BOOL OnInitDialog();
	afx_msg void OnTimer(UINT nIDEvent);
	afx_msg void OnStartTest();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AU9525TESTER_H__ADB3F256_C01A_49DD_A30C_C41B56B59AC3__INCLUDED_)
