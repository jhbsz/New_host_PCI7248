// 6710MPToolDlg.h : header file
//
#include "Label.h"
#include "mscomm.h"
//{{AFX_INCLUDES()
#include "mscomm.h"
//}}AFX_INCLUDES
#if !defined(AFX_6710MPTOOLDLG_H__D913C92E_8E3E_43A6_8D81_86A7B03CB7EF__INCLUDED_)
#define AFX_6710MPTOOLDLG_H__D913C92E_8E3E_43A6_8D81_86A7B03CB7EF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CMy6710MPToolDlg dialog

class CMy6710MPToolDlg : public CDialog
{
// Construction
public:

    void PCI7248_bin(BYTE Channel, BYTE PCI7248bin);
	void Timer_1ms(int ms);
	void SetTimer_1ms();

	void LedFunctionPass();
	void LedFunctionFail();
	void RWDataPass();
	void RWDataFail();
	void SpeedErrorPass();
	void SpeedErrorFail();
	void UnknowDevicePass();
	void UnknowDeviceFail();
	void TestInit();
	void ClearHandlesOfTwoPorts();
	void ShowTestSuccess(void);
	void ShowTestFail(void);
	void ShowInformation(CString ShowData);
	void Delay(DWORD dwDelayTime);
	void StartTest(void);
	CMy6710MPToolDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CMy6710MPToolDlg)
	enum { IDD = IDD_MY6710MPTOOL_DIALOG };
	CLabel	m_SpeedError;
	CLabel	m_LedFail;
	CLabel	m_RWFail;
	CLabel	m_TestResult;
	CLabel	m_UnknowDevice;
	CEdit	m_Infor;
	CStatic	m_ctrlChipIcon;
	CString	m_strResult;
	int	m_strPassNum;
	int	m_strFailNum;
	CString	m_strTestInformation;
	CString	m_strPID;
	CString	m_strVID;
	CMSComm	m_Comm;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMy6710MPToolDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CMy6710MPToolDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnButtonStart();
	afx_msg void OnTimer(UINT nIDEvent);
	afx_msg void OnBUTTONExit();
	afx_msg void OnButtonClear();
	afx_msg void OnChangeVid();
	afx_msg void OnChangePid();
	afx_msg void OnBUTTONOpenFile();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_6710MPTOOLDLG_H__D913C92E_8E3E_43A6_8D81_86A7B03CB7EF__INCLUDED_)
