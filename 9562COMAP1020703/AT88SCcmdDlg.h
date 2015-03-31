#if !defined(AFX_AT88SCCMDDLG_H__E2287597_879F_46A9_AAE1_AE6B994F344A__INCLUDED_)
#define AFX_AT88SCCMDDLG_H__E2287597_879F_46A9_AAE1_AE6B994F344A__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AT88SCcmdDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAT88SCcmdDlg dialog

class CAT88SCcmdDlg : public CDialog
{
// Construction
public:
	CSerial CSerial9525;
	BYTE SlotNum;
	CAT88SCcmdDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAT88SCcmdDlg)
	enum { IDD = IDD_AT88SC_CMD };
	CString	m_strReadDataLength;
	CString	m_strSendData;
	CString	m_strGetData;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAT88SCcmdDlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAT88SCcmdDlg)
	afx_msg void OnButtonSmcCommand();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CAT88SCcmdDlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AT88SCCMDDLG_H__E2287597_879F_46A9_AAE1_AE6B994F344A__INCLUDED_)
