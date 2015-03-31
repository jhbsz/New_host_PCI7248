#if !defined(AFX_LCM_KEYPAD_EEPROMDLG_H__A57DFBD8_B231_4EA0_B6CC_3897C2853671__INCLUDED_)
#define AFX_LCM_KEYPAD_EEPROMDLG_H__A57DFBD8_B231_4EA0_B6CC_3897C2853671__INCLUDED_

#include "9525RS232Lib.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000



#include "Label.h"
// LCM_KEYPAD_EEPROMDlg.h : header file
//
#define LOAD_FW

#define LCM_LIST_NUM 0x05
#define LCM_HD44780 0
#define LCM_KS0108 1
#define LCM_ST7920 2
#define LCM_KGM0053 3
#define LCM_PE12832 4
#define LCM_DEFAULT 5
/////////////////////////////////////////////////////////////////////////////
// CLCM_KEYPAD_EEPROMDlg dialog
class EepromModel
{
	public:
		EepromModel(CString Vid, 
					CString Pid, 
					CString VidString, 
					CString PidString, 
					CString SnString,
					BOOL bIsSupported)
		{	m_Vid = Vid; 
			m_Pid = Pid;
			m_VidString = VidString; 
			m_PidString = PidString;
			m_SnString = SnString; 
			m_IsSnSupported = bIsSupported;
			m_LCDType = LCM_DEFAULT;
			m_LCDAddr = 0;
			m_LCDLen = 0;
		};
		EepromModel(){};

		void SetVid(CString Vid) {m_Vid = Vid;};
		void SetPid(CString Pid) {m_Pid = Pid;};
		void SetVidString(CString VidString){m_VidString = VidString;};
		void SetPidString(CString PidString){m_PidString = PidString;};
		void SetSnString(CString SnString){m_SnString = SnString;};	
		void SetSnEnabled(BOOL bIsSupported){m_IsSnSupported = bIsSupported;};//BOOL
		
		CString GetVid() {return m_Vid;};		
		CString GetPid() {return m_Pid;};
		CString GetPidString() {return m_PidString;};
		CString GetVidString() {return m_VidString;};
		CString GetSnString() {return m_SnString;};
		BOOL GetIsSnSupported() {return m_IsSnSupported;};
		
		BOOL Binary2EepromModel(UCHAR *pBinData, UINT iBinLen);
		BOOL EepromModel2Bin(UCHAR *pBinData, UINT iBufLen, UINT *iBinLen);
	private:
		CString	m_Vid;
		CString	m_Pid;
		CString	m_VidString;
		CString	m_PidString;
		CString m_SnString;
		BOOL    m_IsSnSupported;//BOOL
		UCHAR   m_LCDType;
		UCHAR   m_LCDAddr;
		UINT    m_LCDLen;
		UCHAR   m_LCDData[5000];	
};


class CLCM_KEYPAD_EEPROMDlg : public CDialog
{
// Construction
public:
	BYTE ByteASCToOffset(BYTE ASCdata);
	LONG DisplayST7565Data(CSerial *m_ctrlCSerial, BYTE CharacterOFFSET, BYTE Row, BYTE Column);
	void LcdPE12832_init(CSerial *m_ctrlCSerial);
	void LcdKGM0053_init(CSerial *m_ctrlCSerial);
	LONG ClrLcd_ST7565(CSerial *m_ctrlCSerial);
	INT m_nLCMIndex;
	CSerial CSerial9525;
	void UpdateDataToEepromModel(EepromModel *pEepromModel);
	UCHAR m_abBinData[5000];
	//UCHAR m_abBinData[4000];//ÕÅÃ÷Ö®
	UINT m_iBinDataLen;
	CLabel m_Result;
	CProgressCtrl * m_Prog;
	CStatic		  *m_Status;


	EepromModel	*m_EepromModel;

	CLCM_KEYPAD_EEPROMDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CLCM_KEYPAD_EEPROMDlg)
	enum { IDD = IDD_LCM_KEYPAD_EEPROM };
	CButton	m_About;
	CButton	m_Close;
	CButton	m_WriteEeprom;
	CButton	m_SaveFile;
	CButton	m_OpenFile;
	CComboBox	m_ctlLCMList;
	UINT	m_LCMPosX;
	UINT	m_LCMPosY;
	CString	m_UsbVid;
	CString	m_UsbPid;
	CString	m_VidString;
	CString	m_PidString;
	BOOL	m_IsSupportSn;
	CString	m_SnString;
	CString	m_KeypadValue;
	CString	m_RangeX;
	UINT	m_DisplayLength;
	UINT	m_DisplayHigh;
	CString	m_RangeY;
	CString	m_DisplayString;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CLCM_KEYPAD_EEPROMDlg)
	public:
	virtual void OnFinalRelease();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CLCM_KEYPAD_EEPROMDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnHelp1();
	afx_msg void OnAbout();
	afx_msg void OnClose();
	afx_msg void OnOpenFile();
	afx_msg void OnSaveFile();
	afx_msg void OnWriteEeprom();
	afx_msg void OnTestLcmKeypad();
	afx_msg void OnDisplayGraph();
	afx_msg void OnSelchangeLcmSelect();
	afx_msg void OnDisplayText();
	afx_msg void OnDISPLAYClear();
	afx_msg void OnBacklightChange();
	afx_msg void OnUpdateFw();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CLCM_KEYPAD_EEPROMDlg)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
private:
	void UpdateDataFromEepromModel(EepromModel *pEepromModel);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_LCM_KEYPAD_EEPROMDLG_H__A57DFBD8_B231_4EA0_B6CC_3897C2853671__INCLUDED_)
