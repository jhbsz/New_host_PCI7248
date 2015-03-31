#if !defined(AFX_AT24CDLG_H__F04F367B_9BF8_4D7A_A1A3_5199799CE6AB__INCLUDED_)
#define AFX_AT24CDLG_H__F04F367B_9BF8_4D7A_A1A3_5199799CE6AB__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AT24CDlg.h : header file
//
	
/////////////////////////////////////////////////////////////////////////////
// CAT24CDlg dialog

class CAT24CDlg : public CDialog
{
// Construction
public:
	BYTE SlotNum;
	CSerial CSerial9525;
	CAT24CDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAT24CDlg)
	enum { IDD = IDD_AT24C };
	int	m_strReadAddress;
	int	m_strPageSize;
	int	m_strReadLength;
	int	m_strWriteAddress;
	CString	m_strAccessData;
	CString	m_strEEPROMdata;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAT24CDlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAT24CDlg)
	afx_msg void OnWrite();
	afx_msg void OnButtonWrite();
	afx_msg void OnButtonRead();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CAT24CDlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AT24CDLG_H__F04F367B_9BF8_4D7A_A1A3_5199799CE6AB__INCLUDED_)
