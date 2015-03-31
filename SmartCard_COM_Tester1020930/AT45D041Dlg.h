#if !defined(AFX_AT45D041DLG_H__976EA9DA_014F_4689_8EE9_A2E9CD94BB76__INCLUDED_)
#define AFX_AT45D041DLG_H__976EA9DA_014F_4689_8EE9_A2E9CD94BB76__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AT45D041Dlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAT45D041Dlg dialog

class CAT45D041Dlg : public CDialog
{
// Construction
public:
	BYTE SlotNum;
	CSerial CSerial9525;
	CAT45D041Dlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAT45D041Dlg)
	enum { IDD = IDD_AT45D041 };
	int	m_strPageNum;
	int	m_strAddressInBufferOrMemory; 
	int	m_strDataLengthToRead;
	CString	m_strDataToWrite;
	CString	m_strDataFromCard;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAT45D041Dlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAT45D041Dlg)
	afx_msg void OnBUTTONMainMemoryPageRead();
	afx_msg void OnBUTTONBuffer1Read();
	afx_msg void OnBUTTONBuffer2Read();
	afx_msg void OnBUTTONMainMemoryPagetoBuffer1Xfr();
	afx_msg void OnBUTTONMainMemoryPagetoBuffer2Xfr();
	afx_msg void OnBUTTONMainMemoryPagetoBuffer1Compare();
	afx_msg void OnBUTTONMainMemoryPagetoBuffer2Compare();
	afx_msg void OnBUTTONBuffer1Write();
	afx_msg void OnBUTTONBuffer2Write();
	afx_msg void OnBUTTONBuffer1toMemoryPageProgramwithErase();
	afx_msg void OnBUTTONBuffer2toMemoryPageProgramwithErase();
	afx_msg void OnBUTTONBuffer1toMemoryPageProgramwithoutErase();
	afx_msg void OnBUTTONBuffer2toMemoryPageProgramwithoutErase();
	afx_msg void OnBUTTONMemoryPageProgramthroughBuffer1();
	afx_msg void OnBUTTONMemoryPageProgramthroughBuffer2();
	afx_msg void OnBUTTONAutoPageProgramthroughBuffer1();
	afx_msg void OnBUTTONAutoPageProgramthroughBuffer2();
	afx_msg void OnBUTTONGetStatusRegister();
	afx_msg void OnBUTTONTestReadSpeed();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CAT45D041Dlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AT45D041DLG_H__976EA9DA_014F_4689_8EE9_A2E9CD94BB76__INCLUDED_)
