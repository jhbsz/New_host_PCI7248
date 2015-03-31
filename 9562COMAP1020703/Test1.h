#if !defined(AFX_TEST1_H__19C50809_8120_4060_830D_9DABD9A4B8A2__INCLUDED_)
#define AFX_TEST1_H__19C50809_8120_4060_830D_9DABD9A4B8A2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Test1.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// Test1 dialog

class Test1 : public CDialog
{
// Construction
public:
	Test1(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(Test1)
	enum { IDD = IDD_TEST1_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(Test1)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(Test1)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TEST1_H__19C50809_8120_4060_830D_9DABD9A4B8A2__INCLUDED_)
