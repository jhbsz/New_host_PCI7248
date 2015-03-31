#if !defined(AFX_SLE4428DLG_H__C7A81AE2_2456_464D_A282_9EBECD71E45A__INCLUDED_)
#define AFX_SLE4428DLG_H__C7A81AE2_2456_464D_A282_9EBECD71E45A__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// SLE4428Dlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CSLE4428Dlg dialog

class CSLE4428Dlg : public CDialog
{
// Construction
public:
	BYTE SlotNum;
	CSerial CSerial9525;
	CSLE4428Dlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CSLE4428Dlg)
	enum { IDD = IDD_SLE4428 };
	int	m_strWriteEraseWithPBAddr;
	int	m_strWriteEraseWithPBData;
	int	m_strWriteEraseWithPBReplyLen;
	int	m_strWriteEraseWithoutPBAddr;
	int	m_strWriteEraseWithoutPBData;
	int	m_strWriteEraseWithoutPBReplyLen;
	int	m_strWritePBdataComparisonAddr;
	int	m_strWritePBdataComparisonData;
	int	m_strWritePBdataComparisonReplyLen;
	int	m_strRead9BitsAddr;
	int	m_strRead9BitsData;
	int	m_strRead9BitsReplyLen;
	int	m_strRead8BitsAddr;
	int	m_strRead8BitsData;
	int	m_strRead8BitsReplyLen;
	int	m_strPSC1;
	int	m_strPSC2;
	int	m_strWriteErrorCounterAddr;
	int	m_strWriteErrorCounterReplyLen;
	int	m_strVerify1stPSCAddr;
	int	m_strVerify1stPSCData;
	int	m_strVerify1stPSCReplyLen;
	int	m_strVerify2ndPSCAddr;
	int	m_strVerify2ndPSCData;
	int	m_strVerify2ndPSCReplyLen;
	int	m_strEraseErrorCountAddr;
	int	m_strEraseErrorCountData;
	int	m_strEraseErrorCountReplyLen;
	CString	m_strCardResponse;
	int	m_strWriteErrorCounterData;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSLE4428Dlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CSLE4428Dlg)
	afx_msg void OnBUTTONWriteAndEraseWithPB();
	afx_msg void OnBUTTONRead8Bits();
	afx_msg void OnBUTTONRead9Bits();
	afx_msg void OnBUTTONWriteErrorCounter();
	afx_msg void OnBUTTONVerify1stPSC();
	afx_msg void OnBUTTONVerify2ndPSC();
	afx_msg void OnBUTTONEraseErrorCount();
	afx_msg void OnBUTTONWriteAndEraseWithoutPB();
	afx_msg void OnBUTTONWritePBAndDataComparison();
	afx_msg void OnBUTTONVerifyPSCAndEraseErrorCount();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CSLE4428Dlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SLE4428DLG_H__C7A81AE2_2456_464D_A282_9EBECD71E45A__INCLUDED_)
