#if !defined(AFX_SLE4442DLG_H__09495E59_86A5_46B1_94D4_6BDAFFEF4C87__INCLUDED_)
#define AFX_SLE4442DLG_H__09495E59_86A5_46B1_94D4_6BDAFFEF4C87__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// SLE4442Dlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CSLE4442Dlg dialog

class CSLE4442Dlg : public CDialog
{
// Construction
public:
	BYTE SlotNum;
	CSerial CSerial9525;
	CSLE4442Dlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CSLE4442Dlg)
	enum { IDD = IDD_SLE4442 };
	int	m_strReadMainMemoryAddr;
	int	m_strReadMainMemoryData;
	int	m_strReadMainMemoryReplyLen;
	int	m_strUpdateMainMemoryAddr;
	int	m_strUpdateMainMemoryData;
	int	m_strUpdateMainMemoryReplyLen;
	int	m_strReadProtectionMemoryAddr;
	int	m_strReadProtectionMemoryData;
	int	m_strReadProtectionMemoryReplyLen;
	int	m_strWriteProtectionMemoryAddr;
	int	m_strWriteProtectionMemoryData;
	int	m_strWriteProtectionMemoryReplyLen;
	int	m_strReadSecurityMemoryAddr;
	int	m_strReadSecurityMemoryData;
	int	m_strReadSecurityMemoryReplyLen;
	int	m_strUpdateSecurityMemoryAddr;
	int	m_strUpdateSecurityMemoryData;
	int	m_strUpdateSecurityMemoryReplyLen;
	int	m_strCompareVerificationDataAddr;
	int	m_strCompareVerificationDataData;
	int	m_strCompareVerificationDataReplyLen;
	CString	m_strCardResponse;
	int	m_strReferenceData1;
	int	m_strReferenceData2;
	int	m_strReferenceData3;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSLE4442Dlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CSLE4442Dlg)
	afx_msg void OnBUTTONReadMainMemory();
	afx_msg void OnBUTTONReadProtectionMemory();
	afx_msg void OnBUTTONReadSecurityMemory();
	afx_msg void OnVerify();
	afx_msg void OnBUTTONUpdateMainMemory();
	afx_msg void OnBUTTONWriteProtectionMemory();
	afx_msg void OnBUTTONUpdateSecurityMemory();
	afx_msg void OnBUTTONCompareVerificationData();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CSLE4442Dlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SLE4442DLG_H__09495E59_86A5_46B1_94D4_6BDAFFEF4C87__INCLUDED_)
