#if !defined(AFX_SLE4442CMDDLG_H__2C75FE64_3681_4F3C_A739_918EA512C789__INCLUDED_)
#define AFX_SLE4442CMDDLG_H__2C75FE64_3681_4F3C_A739_918EA512C789__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// SLE4442CMDDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CSLE4442CMDDlg dialog

class CSLE4442CMDDlg : public CDialog
{
// Construction
public:
	CSerial CSerial9525;
	BYTE SlotNum;
	CSLE4442CMDDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CSLE4442CMDDlg)
	enum { IDD = IDD_SLE4442_CMD };
	CString	m_STRReadDataLength;
	CString	m_strProtectBitFlag;
	CString	m_strClockNumberFlag;
	CString	m_strCommand2;
	CString	m_strCommand1;
	CString	m_strCommand3;
	CString	m_strResponse;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSLE4442CMDDlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CSLE4442CMDDlg)
	afx_msg void OnButtonSle4442CardBreak();
	afx_msg void OnButtonSle4442CardCommand();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CSLE4442CMDDlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SLE4442CMDDLG_H__2C75FE64_3681_4F3C_A739_918EA512C789__INCLUDED_)
