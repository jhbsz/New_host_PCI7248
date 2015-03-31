#if !defined(AFX_AT88SCDLG_H__CA602958_B66F_44FA_A8B2_5313EF190DA1__INCLUDED_)
#define AFX_AT88SCDLG_H__CA602958_B66F_44FA_A8B2_5313EF190DA1__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AT88SCDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAT88SCDlg dialog

class CAT88SCDlg : public CDialog
{
// Construction
public:
	CSerial CSerial9525;
	BYTE SlotNum;
	CAT88SCDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAT88SCDlg)
	enum { IDD = IDD_AT88SC };
	CComboBox	m_ctrlCOMBO_ZONE;
	CComboBox	m_ctrlCOMBO_R_W;
	int	m_strWriteUserZoneAddr;
	CString	m_strWriteUserZoneData;
	int	m_strWriteUserZoneReplyLen;
	int	m_strReadUserZoneAddr;
	CString	m_strReadUserZoneData;
	int	m_strReadUserZoneReplyLen;
	int	m_strWriteConfigurationZoneAddr;
	CString	m_strWriteConfigurationZoneData;
	int	m_strWriteConfigurationZoneReplyLen;
	int	m_strReadConfigurationZoneAddr;
	CString	m_strReadConfigurationZoneData;
	int	m_strReadConfigurationZoneReplyLen;
	int	m_strSetUserZoneAddressAddr;
	CString	m_strSetUserZoneAddressData;
	int	m_strSetUserZoneAddressReplyLen;
	CString	m_strCardResponse;
	int	m_strPassword1;
	int	m_strPassword2;
	int	m_strPassword3;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAT88SCDlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAT88SCDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnBUTTONWriteUserZone();
	afx_msg void OnBUTTONReadUserZone();
	afx_msg void OnBUTTONWriteConfigurationZone();
	afx_msg void OnBUTTONReadConfigurationZone();
	afx_msg void OnBUTTONSetUserZoneAddress();
	afx_msg void OnBUTTONVerifyPassword();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CAT88SCDlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AT88SCDLG_H__CA602958_B66F_44FA_A8B2_5313EF190DA1__INCLUDED_)
