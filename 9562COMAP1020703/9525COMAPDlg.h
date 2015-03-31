// 9525COMAPDlg.h : header file
//
//{{AFX_INCLUDES()
#include "mscomm1.h"
#include "labelcontrol1.h"
//}}AFX_INCLUDES

#if !defined(AFX_9525COMAPDLG_H__E1D1EF90_D9A6_4112_B625_A890CB3F290F__INCLUDED_)
#define AFX_9525COMAPDLG_H__E1D1EF90_D9A6_4112_B625_A890CB3F290F__INCLUDED_

#include "InphoneCmdDlg.h"	// Added by ClassView
#include "AT88SCcmdDlg.h"	// Added by ClassView
#include "AT24CDlg.h"	// Added by ClassView
#include "AT88SCDlg.h"	// Added by ClassView
#include "LCM_KEYPAD_EEPROMDlg.h"	// Added by ClassView
#include "AT45D041Dlg.h"	// Added by ClassView
#include "SLE4442Dlg.h"	// Added by ClassView
#include "SLE4428Dlg.h"	// Added by ClassView
#include "SLE4442CMDDlg.h"	// Added by ClassView
#include "DASK.h"
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CMy9525COMAPDlg dialog

class CMy9525COMAPDlg : public CDialog
{
// Construction
public:
	virtual void Disp();
	virtual void TestInit();
	virtual void TestSub();
	CMy9525COMAPDlg(CWnd* pParent = NULL);	// standard constructor

	HWND winHnd;

// Dialog Data
	//{{AFX_DATA(CMy9525COMAPDlg)
	enum { IDD = IDD_MY9525COMAP_DIALOG };
	CComboBox	m_ctrlSlot1CardType;
	CComboBox	m_ctrlSlot0CardType;
	CComboBox	m_ctrlBaudRate;
	CComboBox	m_ctrlCOM;
	CString	m_strSlot0_CardType;
	CString	m_strSlot1_CardType;
	CString	m_strSlot0ATR;
	CString	m_strSlot1ATR;
	CString	m_strSlot0BlockData;
	CString	m_strSlot1BlockData;
	int		m_Slot0_Ttype;
	int		m_Slot1_Ttype;
	CString	m_strPID;
	CString	m_strVID;
	CString	m_strSlot0ResponseData;
	CString	m_strSlot1ResponseData;
	CString	m_strPIDNum;
	CString	m_strVIDNum;
	CString	m_strSerialNum;
	CString	m_strReleaseNum;
	CMSComm	m_Comm;
	CLabelControl	m_TestLabel;
	CLabelControl	m_NewChip;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMy9525COMAPDlg)
	public:
	virtual BOOL DestroyWindow();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	virtual LRESULT DefWindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CMy9525COMAPDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnButton1();
	afx_msg void OnBUTTONPoweron();
	afx_msg void OnBUTTONpoweroff();
	afx_msg void OnTimer(UINT nIDEvent);
	afx_msg void OnSlot0RADIOT0();
	afx_msg void OnSlot0RADIOT1();
	afx_msg void OnSlot1RADIOT0();
	afx_msg void OnSlot1RADIOT1();
	afx_msg void OnBUTTONSlot0Xfr();
	afx_msg void OnBUTTONConnect();
	afx_msg void OnSelchangeComboCom();
	afx_msg void OnSelchangeCOMBOBaudRate();
	afx_msg void OnSelchangeCOMBOSlot0CardType();
	afx_msg void OnBUTTONSlot1Xfr();  
	afx_msg void OnButtonLcmKeypadEeprom();
	afx_msg void OnBUTTONSlot0SLE4428();
	afx_msg void OnBUTTONSlot1SLE4428();
	afx_msg void OnBUTTONSlot0SLE4442();
	afx_msg void OnBUTTONSlot1SLE4442();
	afx_msg void OnBUTTONSlot0AT45D041();
	afx_msg void OnBUTTONSlot1AT45D041();
	afx_msg void OnBUTTONSlot0AT88SC();
	afx_msg void OnBUTTONSlot0EEPROMCardEdit();
	afx_msg void OnBUTTONSlot1EEPROMCardEdit();
	afx_msg void OnBUTTONSlot1AT88SC();
	afx_msg void OnBUTTONSlot0Inphone();
	afx_msg void OnBUTTONSlot1Inphone();
	afx_msg void OnBeginTest();
	afx_msg void OnEXit();
	afx_msg void OnClose();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	CInphoneCmdDlg m_dCInphoneCmdDlg;
	CAT88SCcmdDlg m_dAT88SCcmdDlg;
	CSLE4442CMDDlg m_dSLE4442CMDDlg;
	CAT24CDlg m_dAT24CDlg;
	CAT88SCDlg m_dAT88SCDlg;
	CLCM_KEYPAD_EEPROMDlg m_dLCM_KEYPAD_EEPROMDlg;
	CAT45D041Dlg m_dAT45D041Dlg;
	CSLE4442Dlg m_dSLE4442Dlg;
	CSLE4428Dlg m_dSLE4428Dlg;


//	CSLE4442Dlg m_dSLE4442Dlg;
//	CAT45D041Dlg m_dAT45D041Dlg;
//	CLCM_KEYPAD_EEPROMDlg m_dLCM_KEYPAD_EEPROMDlg;

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_9525COMAPDLG_H__E1D1EF90_D9A6_4112_B625_A890CB3F290F__INCLUDED_)
