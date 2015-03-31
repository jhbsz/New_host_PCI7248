#if !defined(AFX_INPHONECMDDLG_H__C9994F0C_E813_4B32_956E_8B76AAF0A093__INCLUDED_)
#define AFX_INPHONECMDDLG_H__C9994F0C_E813_4B32_956E_8B76AAF0A093__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// InphoneCmdDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CInphoneCmdDlg dialog

class CInphoneCmdDlg : public CDialog
{
// Construction
public:
	BYTE SlotNum;
	CSerial CSerial9525;
	CInphoneCmdDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CInphoneCmdDlg)
	enum { IDD = IDD_Inphone_CMD };
	CString	m_strINPHONE_CARD_Read;
	CString	m_strINPHONE_CARD_PROG;
	CString	m_strINPHONE_CARD_MOVE_ADDRESS;
	CString	m_strINPHONE_CARD_AUTHENTICATION_KEY1;
	CString	m_strINPHONE_CARD_AUTHENTICATION_KEY2;
	CString	m_strCardResponse;
	CString	m_strKey1_SendData;
	CString	m_strKey2_SendData;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CInphoneCmdDlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CInphoneCmdDlg)
	afx_msg void OnButtonInphoneCardReset();
	afx_msg void OnBUTTONINPHONECARDRead();
	afx_msg void OnButtonInphoneCardProg();
	afx_msg void OnButtonInphoneCardMoveAddress();
	afx_msg void OnButtoninphoneCardAuthenticationKey1();
	afx_msg void OnButtonInphoneCardAuthenticationKey2();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CInphoneCmdDlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_INPHONECMDDLG_H__C9994F0C_E813_4B32_956E_8B76AAF0A093__INCLUDED_)
