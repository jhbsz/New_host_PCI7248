// 9525COMAPDlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"

#include "SLE4428Dlg.h"
#include "SLE4442Dlg.h"
#include "AT45D041Dlg.h"
#include "LCM_KEYPAD_EEPROMDlg.h"

#include "9525COMAPDlg.h"
#include "9525RS232Lib.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

BYTE AllenStop=0;
I16 card=-1, card_number = 0;
BYTE TestFlag=1;
CSerial CSerial9525;
CATR_capacity Slot0_ATR, Slot1_ATR;
Reader_Descriptor Reader9525_Descriptor;
int BaudRate;
BYTE CurrentSlot = 0xFF;//0:Slot0, 1:Slot1, 0xFF不在Slot0,或1的状态。
/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

int ECHO;

CString Pattern="3B FD 12 00 FF 80 31 FE 22 43 61 72 64 54 30 54 31 5F 45 43 48 4F 16";
int m_TestResult;
class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
} 

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMy9525COMAPDlg dialog

CMy9525COMAPDlg::CMy9525COMAPDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMy9525COMAPDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CMy9525COMAPDlg)
	m_strSlot0_CardType = _T("");
	m_strSlot1_CardType = _T("");
	m_strSlot0ATR = _T("");
	m_strSlot1ATR = _T("");
	m_strSlot0BlockData = _T("");
	m_strSlot1BlockData = _T("");
	m_Slot0_Ttype = -1;
	m_Slot1_Ttype = -1;
	m_strPID = _T("");
	m_strVID = _T("");
	m_strSlot0ResponseData = _T("");
	m_strSlot1ResponseData = _T("");
	m_strPIDNum = _T("");
	m_strVIDNum = _T("");
	m_strSerialNum = _T("");
	m_strReleaseNum = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMy9525COMAPDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CMy9525COMAPDlg)
	DDX_Control(pDX, IDC_COMBO_Slot1CardType, m_ctrlSlot1CardType);
	DDX_Control(pDX, IDC_COMBO_Slot0CardType, m_ctrlSlot0CardType);
	DDX_Control(pDX, IDC_COMBO_BaudRate, m_ctrlBaudRate);
	DDX_Control(pDX, IDC_COMBO_COM, m_ctrlCOM);
	DDX_Text(pDX, IDC_Slot0_CardType, m_strSlot0_CardType);
	DDX_Text(pDX, IDC_Slot1_CardType, m_strSlot1_CardType);
	DDX_Text(pDX, IDC_Slot0_ATR, m_strSlot0ATR);
	DDX_Text(pDX, IDC_Slot1_ATR, m_strSlot1ATR);
	DDX_Text(pDX, IDC_Slot0_BlockData, m_strSlot0BlockData);
	DDX_Text(pDX, IDC_Slot1_BlockData, m_strSlot1BlockData);
	DDX_Radio(pDX, IDC_Slot0_RADIO_T0, m_Slot0_Ttype);
	DDX_Radio(pDX, IDC_Slot1_RADIO_T0, m_Slot1_Ttype);
	DDX_Text(pDX, IDC_PID, m_strPID);
	DDX_Text(pDX, IDC_VID, m_strVID);
	DDX_Text(pDX, IDC_Slot0_Response_Data, m_strSlot0ResponseData);
	DDX_Text(pDX, IDC_Slot1_Response_Data, m_strSlot1ResponseData);
	DDX_Text(pDX, IDC_PID_Num, m_strPIDNum);
	DDX_Text(pDX, IDC_VID_Num, m_strVIDNum);
	DDX_Text(pDX, IDC_SerialNum, m_strSerialNum);
	DDX_Text(pDX, IDC_ReleaseNum, m_strReleaseNum);
	DDX_Control(pDX, IDC_MSCOMM1, m_Comm);
	DDX_Control(pDX, IDC_TestLabel, m_TestLabel);
	DDX_Control(pDX, IDC_NewChip, m_NewChip);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CMy9525COMAPDlg, CDialog)
	//{{AFX_MSG_MAP(CMy9525COMAPDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	ON_BN_CLICKED(IDC_BUTTON_Power_on, OnBUTTONPoweron)
	ON_BN_CLICKED(IDC_BUTTON_power_off, OnBUTTONpoweroff)
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_Slot0_RADIO_T0, OnSlot0RADIOT0)
	ON_BN_CLICKED(IDC_Slot0_RADIO_T1, OnSlot0RADIOT1)
	ON_BN_CLICKED(IDC_Slot1_RADIO_T0, OnSlot1RADIOT0)
	ON_BN_CLICKED(IDC_Slot1_RADIO_T1, OnSlot1RADIOT1)
	ON_BN_CLICKED(IDC_BUTTON_Slot0_Xfr, OnBUTTONSlot0Xfr)
	ON_BN_CLICKED(IDC_BUTTON_Connect, OnBUTTONConnect)
	ON_CBN_SELCHANGE(IDC_COMBO_COM, OnSelchangeComboCom)
	ON_CBN_SELCHANGE(IDC_COMBO_BaudRate, OnSelchangeCOMBOBaudRate)
	ON_CBN_SELCHANGE(IDC_COMBO_Slot0CardType, OnSelchangeCOMBOSlot0CardType)
	ON_BN_CLICKED(IDC_BUTTON_Slot1_Xfr, OnBUTTONSlot1Xfr)
	ON_BN_CLICKED(IDC_BUTTON_LCM_KEYPAD_EEPROM, OnButtonLcmKeypadEeprom)
	ON_BN_CLICKED(IDC_BUTTONSlot0SLE4428, OnBUTTONSlot0SLE4428)
	ON_BN_CLICKED(IDC_BUTTONSlot1SLE4428, OnBUTTONSlot1SLE4428)
	ON_BN_CLICKED(IDC_BUTTONSlot0SLE4442, OnBUTTONSlot0SLE4442)
	ON_BN_CLICKED(IDC_BUTTONSlot1SLE4442, OnBUTTONSlot1SLE4442)
	ON_BN_CLICKED(IDC_BUTTONSlot0AT45D041, OnBUTTONSlot0AT45D041)
	ON_BN_CLICKED(IDC_BUTTONSlot1AT45D041, OnBUTTONSlot1AT45D041)
	ON_BN_CLICKED(IDC_BUTTONSlot0AT88SC, OnBUTTONSlot0AT88SC)
	ON_BN_CLICKED(IDC_BUTTONSlot0EEPROMCardEdit, OnBUTTONSlot0EEPROMCardEdit)
	ON_BN_CLICKED(IDC_BUTTONSlot1EEPROMCardEdit, OnBUTTONSlot1EEPROMCardEdit)
	ON_BN_CLICKED(IDC_BUTTONSlot1AT88SC, OnBUTTONSlot1AT88SC)
	ON_BN_CLICKED(IDC_BUTTONSlot0Inphone, OnBUTTONSlot0Inphone)
	ON_BN_CLICKED(IDC_BUTTONSlot1Inphone, OnBUTTONSlot1Inphone)
	ON_BN_CLICKED(IDC_BeginTest, OnBeginTest)
	ON_BN_CLICKED(IDC_EXit, OnEXit)
	ON_WM_CLOSE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMy9525COMAPDlg message handlers

BOOL CMy9525COMAPDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here

	if(!CSerial9525.Open(2,38400))
	{
		MessageBox("Com occupyed!");
	}
	//CSerial9525.ClosePort();
	m_ctrlCOM.AddString("COM1");
	m_ctrlCOM.AddString("COM2");
	m_ctrlCOM.AddString("COM3");
	m_ctrlCOM.AddString("COM4");
	m_ctrlCOM.SetCurSel(1);

	m_ctrlBaudRate.AddString("4800");
	m_ctrlBaudRate.AddString("9600");
	m_ctrlBaudRate.AddString("19200");
	m_ctrlBaudRate.AddString("38400");
	
//	m_ctrlBaudRate.AddString("57600");
//	m_ctrlBaudRate.AddString("76800");
//	m_ctrlBaudRate.AddString("115200");
	BaudRate = 38400;

	m_ctrlBaudRate.SetCurSel(3);

	m_ctrlSlot0CardType.AddString("Asynchronous card");
/*
	m_ctrlSlot0CardType.AddString("AT24C card");
	m_ctrlSlot0CardType.AddString("SLE4418/28");
	m_ctrlSlot0CardType.AddString("SLE4432/42");
	m_ctrlSlot0CardType.AddString("AT88SC card");
	m_ctrlSlot0CardType.AddString("INPHONE card");
	m_ctrlSlot0CardType.AddString("AT45D041 card");
*/
	m_ctrlSlot0CardType.SetCurSel(0);

	m_ctrlSlot1CardType.AddString("Asynchronous card");
/*
	m_ctrlSlot1CardType.AddString("AT24C card");
	m_ctrlSlot1CardType.AddString("SLE4418/28");
	m_ctrlSlot1CardType.AddString("SLE4432/42");
	m_ctrlSlot1CardType.AddString("AT88SC card");
	m_ctrlSlot1CardType.AddString("INPHONE card");
	m_ctrlSlot1CardType.AddString("AT45D041 card");
*/
	m_ctrlSlot1CardType.SetCurSel(0);

// 	SetTimer(ID_CLOCK_TIMER,300,NULL);   // Allen Test(2)

	GetDlgItem(IDC_Slot0_RADIO_T0)->EnableWindow(FALSE);
	GetDlgItem(IDC_Slot0_RADIO_T1)->EnableWindow(FALSE);
	GetDlgItem(IDC_Slot1_RADIO_T0)->EnableWindow(FALSE);
	GetDlgItem(IDC_Slot1_RADIO_T1)->EnableWindow(FALSE);

	GetDlgItem(IDC_Slot0_CardType)->EnableWindow(FALSE);
	GetDlgItem(IDC_Slot1_CardType)->EnableWindow(FALSE);

	GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(FALSE);
	GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(FALSE);

//	GetDlgItem(IDC_Slot0_Response_Data)->EnableWindow(FALSE);
//	GetDlgItem(IDC_Slot1_Response_Data)->EnableWindow(FALSE);

	GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(FALSE);

	GetDlgItem(IDC_BUTTONSlot0SLE4428)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot0SLE4442)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot0AT45D041)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot0AT88SC)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot0Inphone)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot0EEPROMCardEdit)->EnableWindow(FALSE);
	

	GetDlgItem(IDC_BUTTONSlot1SLE4428)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot1SLE4442)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot1AT45D041)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot1AT88SC)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot1Inphone)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTONSlot1EEPROMCardEdit)->EnableWindow(FALSE);
	
	m_strPID = "";
	m_strVID = "";
	
	TestInit();

 //   OnBUTTONConnect();
//	Sleep(1000);
 //	TestSub();
//	TestSub();
	TestFlag=3;
	SetTimer(20,1000,FALSE);
	UpdateData(FALSE);
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CMy9525COMAPDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMy9525COMAPDlg::OnPaint() 
{ 
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMy9525COMAPDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CMy9525COMAPDlg::OnButton1() 
{
	// TODO: Add your control notification handler code here
	//----------------test 1-------------------
	/*
	char Comsendbuffer[5];

	Comsendbuffer[0] = 30;
	Comsendbuffer[1] = 30;
	Comsendbuffer[2] = 30;
	Comsendbuffer[3] = 30;
	Comsendbuffer[4] = 35;

	CSerial9525.SendData(Comsendbuffer, 5);

	char ComGetbuffer[5];
	memset(ComGetbuffer, 0, 5);
	Sleep(1000);
	CSerial9525.ReadData(ComGetbuffer,5);*/
	//----------------test 1-------------------
	//----------------test 2-------------------
	/*UINT Result;
	BYTE SlotOut;
	BYTE SequenseOut;
	BYTE StatusOut;
	BYTE ErrorOut; 
	BYTE ClockStatusOut;
	Result = CMD_PC_to_RDR_GetSlotStatus(&CSerial9525,
									0, 2,
									&SlotOut, &SequenseOut, 
									&StatusOut, &ErrorOut, &ClockStatusOut);
	MessageBox(Bytes2CString(&SlotOut,1));
	MessageBox(Bytes2CString(&SequenseOut,1));
	MessageBox(Bytes2CString(&StatusOut,1));
	MessageBox(Bytes2CString(&ErrorOut,1));
	MessageBox(Bytes2CString(&ClockStatusOut,1));*/
    //----------------test 2-------------------
	/*LONG EepromCmdWrite(
		CSerial *m_ctrlCSerial, 
		UCHAR	bSlotNum,
		UCHAR   SequenseIn,
		ULONG	lngStartAddr,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		);*/

	/*LONG EepromCmdWrite(
		long	lngCard,
		UCHAR	bSlotNum,
		ULONG	lngStartAddr,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		)*/

	//---------------test start-------------------
/*	LPCVOID	pSendBuffer;	
	pSendBuffer	 = malloc( 6 );
	*((PUCHAR)pSendBuffer+0)=1;
	*((PUCHAR)pSendBuffer+1)=2;
	*((PUCHAR)pSendBuffer+2)=3;
	*((PUCHAR)pSendBuffer+3)=4;
	*((PUCHAR)pSendBuffer+4)=5;
	*((PUCHAR)pSendBuffer+5)=6;

//	BYTE StatusOut; 
//	BYTE ErrorOut; 
//	BYTE abDataOut[5]; 
//	BYTE abDataOutLen;	

	BOOL	lngStatus;
	lngStatus = CMD_PC_to_RDR_Escape(&CSerial9525,
									0,
									(PUCHAR)pSendBuffer, 6,
									NULL,NULL,
									NULL,NULL);

//									&StatusOut, &ErrorOut,
//									abDataOut, &abDataOutLen);
	free((LPVOID)pSendBuffer);*/
	//----------------test end-------------------

	//----------------test start-------------------
	/*BYTE Comsendbuffer[5];
	CString temp_CString;

	Comsendbuffer[0] = 0x30;
	Comsendbuffer[1] = 0x31;
	Comsendbuffer[2] = 0x32;
	Comsendbuffer[3] = 0x33;
	Comsendbuffer[4] = 0x34;
	temp_CString = Bytes2CString_ASCII(Comsendbuffer,5);
	MessageBox(temp_CString);*/
	//----------------test end-------------------

	//----------------test start-------------------
/*	char Comsendbuffer[5];

	Comsendbuffer[0] = 30;
	Comsendbuffer[1] = 31;
	Comsendbuffer[2] = 32;
	Comsendbuffer[3] = 33;
	Comsendbuffer[4] = 34;
	//EepromCmdWrite(hCardHandle, 0, iCnt , 1, m_abBinData + iCnt)
	if( EepromCmdWrite(&CSerial9525, 0, 0, 5, (unsigned char *)Comsendbuffer) != SCARD_S_SUCCESS )
	{
		Comsendbuffer[0] = 30;
		
	}
	Comsendbuffer[0] = 31;
	//----------------test end-------------------

	UCHAR bData[8];
	UCHAR bNum;
	ULONG lReturnLen;
	bNum = 8;
	if( EepromCmdRead(&CSerial9525, 0, 0, bNum, &bData, &lReturnLen) 
		             != SCARD_S_SUCCESS )
	{
		bNum++;
		return;
	}
	bNum--;*/

	/*unsigned char aaa = 0xff,bbb;
	bbb = aaa+1;
	if((aaa+1)  == bbb)
	{
		aaa++;
	}*/
	//----------------test start-------------------
	CString aaaa;
	int ccc = 123;
	aaaa.Format("%d",ccc);
	MessageBox(aaaa);
	//----------------test end-------------------*/

}

void CMy9525COMAPDlg::OnBUTTONPoweron() 
{
	// TODO: Add your control notification handler code here
	UINT Result;
	BYTE SlotOut;
	BYTE SequenseOut;
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE ChainParameterOut;
	BYTE ATRbuffer[100];   
	BYTE ATRbufferLen;
	Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
										0, 2,
										&StatusOut, &ErrorOut, &ChainParameterOut,
										ATRbuffer,&ATRbufferLen);	
	MessageBox(Bytes2CString(&SlotOut,1));
	MessageBox(Bytes2CString(&SequenseOut,1));
	MessageBox(Bytes2CString(&StatusOut,1));
	MessageBox(Bytes2CString(&ErrorOut,1));
	MessageBox(Bytes2CString(&ChainParameterOut,1));
	MessageBox(Bytes2CString(ATRbuffer,(UINT)ATRbufferLen));	
}

void CMy9525COMAPDlg::OnBUTTONpoweroff() 
{
	// TODO: Add your control notification handler code here
	UINT Result;
	BYTE SlotOut;
	BYTE SequenseOut;
	BYTE StatusOut;
	BYTE ErrorOut; 
	BYTE ClockStatusOut;
	Result = CMD_PC_to_RDR_IccPowerOff(&CSerial9525,
									0,
									&StatusOut, &ErrorOut, &ClockStatusOut);
	MessageBox(Bytes2CString(&SlotOut,1));
	MessageBox(Bytes2CString(&SequenseOut,1));
	MessageBox(Bytes2CString(&StatusOut,1));
	MessageBox(Bytes2CString(&ErrorOut,1));
	MessageBox(Bytes2CString(&ClockStatusOut,1));
}

//-------------------------------ATR data start--------------------------

//-------------------------------ATR data end--------------------------

void CMy9525COMAPDlg::OnTimer(UINT nIDEvent) 
{

	if (( nIDEvent == 55 ) && ( ECHO == 0 ))
	{
		 DO_WritePort(card, Channel_P1A, 0xff);
	}


	 if ( nIDEvent == 20 )
	{
		KillTimer(20);
	 //	AU9525COMAdlg.DoModal();
		
		/* nick marked for disable auto run
		OnBeginTest();*/
	}
/*  
    if ((TestFlag==1)  && (m_TestResult != "Bin2" ))
		{
			if(m_strSlot0ATR  == Pattern)
			{
				m_TestLabel.SetCaption("Slot 0 R/W Pass");
				m_TestLabel.SetBackColor(RGB(0,255,0));

					if(m_strSlot1ATR  == Pattern)
					{
						m_TestLabel.SetCaption("Test  Pass");
						m_TestLabel.SetBackColor(RGB(0,255,0));
						m_TestResult="PASS";
					}
					else
					{
       					m_TestLabel.SetCaption("Bin4 : Slot 1 R/W Fail");
						m_TestLabel.SetBackColor(RGB(255,0,0));
						m_TestResult="Bin4";
					}
			}
			else
			{
       			m_TestLabel.SetCaption("Bin3: Slot 0 R/W Fail");
				m_TestLabel.SetBackColor(RGB(255,0,0));
				m_TestResult="Bin3";
			}
			 TestFlag=2; // UI finish 
			  m_Comm.SetOutput(COleVariant(m_TestResult));
		}

	// TODO: Add your message handler code here and/or call default
	//KillTimer(ID_CLOCK_TIMER);

 
	if (TestFlag !=0)   // Test begin
	{
		return;
    }
	TestFlag=1;
    OnBUTTONConnect();

	BYTE bmSlotICCStateOut;
   
//	if(Check_RDR_to_PC_NotifySlotChange(&CSerial9525,&bmSlotICCStateOut)!=0)
	 if (1)  // Allen debug(1)
	{
	 //	if((bmSlotICCStateOut&0x02) == 0x02)//Slot0 change
		 if (1)  // Allen debug(2) 
		{
		//	if((bmSlotICCStateOut&0x01) == 0x00)//Card out
            if (0)
			{
				m_strSlot0ATR = "";
				m_strSlot0_CardType = "";
				m_strSlot0BlockData = "";
				m_strSlot0ResponseData = "";
				m_Slot0_Ttype = -1;
				GetDlgItem(IDC_Slot0_RADIO_T0)->EnableWindow(FALSE);
				GetDlgItem(IDC_Slot0_RADIO_T1)->EnableWindow(FALSE);
				CurrentSlot = 0xFF;
				GetDlgItem(IDC_BUTTONSlot0SLE4428)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0SLE4442)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0AT88SC)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0AT45D041)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0EEPROMCardEdit)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0Inphone)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(FALSE);
				GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(FALSE);
			}
			else//Card in
			{
				if(m_ctrlSlot0CardType.GetCurSel() == -1)
				{
				   	return;
				 
				}
				//Power off
				BYTE TryTimes = 3;
				UINT Result;
				BYTE StatusOut;
				BYTE ErrorOut; 
				BYTE ClockStatusOut;

				Result = CMD_PC_to_RDR_IccPowerOff(&CSerial9525,
												0,
												&StatusOut, &ErrorOut, &ClockStatusOut);
				if(Result != 1)
				{
					return;
				}
				Result = CMD_SWITCH_CARD_MODE(&CSerial9525,
										  0,
										  (m_ctrlSlot0CardType.GetCurSel() +1)
										  );	
				if(Result != 1)
				{
					MessageBox("SWITCH CARD MODE fail!");
					return;
				}
				CurrentSlot = 0;

				//-----Power on-----
				//UINT Result;
				//BYTE StatusOut;
				//BYTE ErrorOut;
				BYTE ChainParameterOut;
				BYTE ATRbuffer[100];   
				BYTE ATRbufferLen;
				TryTimes = 3;
//				while(TryTimes)
//				{
//					TryTimes--;	
					Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
													0, 2,
													&StatusOut, &ErrorOut, &ChainParameterOut,
													ATRbuffer,&ATRbufferLen);	
					if(Result == 1)
					{
						m_strSlot0ATR = Bytes2CString(ATRbuffer,(UINT)ATRbufferLen);
					 
						UpdateData(FALSE);
//						break;
					}
//				}
				if(Result != 1)
				{
//					TryTimes = 3;
//					while(TryTimes)
//					{
//						TryTimes--;	
						Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
														0, 1,
														&StatusOut, &ErrorOut, &ChainParameterOut,
														ATRbuffer,&ATRbufferLen);	
						if(Result == 1)
						{
							m_strSlot0ATR = Bytes2CString(ATRbuffer,(UINT)ATRbufferLen);
							UpdateData(FALSE);
//							break;
						}
//					}
					if(Result != 1)
					{
						return;
					}
				}	
				if(ATRbufferLen == 0)
				{
					return;
				}
				switch((m_ctrlSlot0CardType.GetCurSel()+1))
				{	
					case 2:
					{
						GetDlgItem(IDC_BUTTONSlot0EEPROMCardEdit)->EnableWindow(TRUE);
					}
					break;
					case 3:
					{
						GetDlgItem(IDC_BUTTONSlot0SLE4428)->EnableWindow(TRUE);	
					}
					break;
					case 4:
					{
						GetDlgItem(IDC_BUTTONSlot0SLE4442)->EnableWindow(TRUE);				
					}
					break;
					case 5:
					{
						GetDlgItem(IDC_BUTTONSlot0AT88SC)->EnableWindow(TRUE);				
					}
					break;
					case 6:
					{
						GetDlgItem(IDC_BUTTONSlot0Inphone)->EnableWindow(TRUE);			
					}
					break;				
					case 7:
					{
						GetDlgItem(IDC_BUTTONSlot0AT45D041)->EnableWindow(TRUE);				
					}
					break;
				}

				//只有Asynchronous Card 才分析ATR
				if(m_ctrlSlot0CardType.GetCurSel() != 0)
				{
					return;
				}
				//分析ATR
				Slot0_ATR.Do_ATR(ATRbuffer, ATRbufferLen);
					
				switch(Slot0_ATR.T_Type)
				{
					case 1:
					{
						m_strSlot0_CardType ="T0";
						GetDlgItem(IDC_Slot0_RADIO_T0)->EnableWindow(TRUE);
						m_Slot0_Ttype = 0;
						GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(TRUE);
						GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(TRUE);
					}
					break;
					case 2:
					{
						m_strSlot0_CardType ="T1";
						GetDlgItem(IDC_Slot0_RADIO_T1)->EnableWindow(TRUE);
						m_Slot0_Ttype = 1;						
						GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(TRUE);
						GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(TRUE);
					}
					break;
					case 3:
					{
						m_strSlot0_CardType ="T0/T1";
						GetDlgItem(IDC_Slot0_RADIO_T0)->EnableWindow(TRUE);
						GetDlgItem(IDC_Slot0_RADIO_T1)->EnableWindow(TRUE);
						m_Slot0_Ttype = -1;
					}
					break;
				}
				UpdateData(FALSE);
			}
		}
		if((bmSlotICCStateOut&0x08) == 0x08)//Slot1 change
		{
			if((bmSlotICCStateOut&0x04) == 0x00)//Card out
			{
				m_strSlot1ATR = "";
				m_strSlot1_CardType = "";
				m_strSlot1BlockData = "";
				m_strSlot1ResponseData = "";
				m_Slot1_Ttype = -1;
				GetDlgItem(IDC_Slot1_RADIO_T0)->EnableWindow(FALSE);
				GetDlgItem(IDC_Slot1_RADIO_T1)->EnableWindow(FALSE);
				CurrentSlot = 0xFF;
				GetDlgItem(IDC_BUTTONSlot1SLE4428)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1SLE4442)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1AT88SC)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1AT45D041)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1EEPROMCardEdit)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1Inphone)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(FALSE);	
				GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(FALSE);
			}
			else//Card in
			{
				if(m_ctrlSlot1CardType.GetCurSel() == -1)
				{
					return;
				}
				//Power off
				BYTE TryTimes = 3;
				UINT Result;
				BYTE StatusOut;
				BYTE ErrorOut; 
				BYTE ClockStatusOut;
				while(TryTimes)
				{
					TryTimes--;	
					Result = CMD_PC_to_RDR_IccPowerOff(&CSerial9525,
												1,
												&StatusOut, &ErrorOut, &ClockStatusOut);
					if(Result == 1)
					{
						break;
					}
				}
				if(Result != 1)
				{
					return;
				}
				Result = CMD_SWITCH_CARD_MODE(&CSerial9525,
										  1,
										  (m_ctrlSlot1CardType.GetCurSel() +1)
										  );	
				if(Result != 1)
				{
					MessageBox("SWITCH CARD MODE fail!");
				}
				CurrentSlot = 1;
				//Power on
				//UINT Result;
				//BYTE StatusOut;
				//BYTE ErrorOut;
				BYTE ChainParameterOut;
				BYTE ATRbuffer[100];   
				BYTE ATRbufferLen;
				TryTimes = 3;
				while(TryTimes)
				{
					TryTimes--;	
					Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
													1, 2,
													&StatusOut, &ErrorOut, &ChainParameterOut,
													ATRbuffer,&ATRbufferLen);	
					if(Result == 1)
					{
						m_strSlot1ATR = Bytes2CString(ATRbuffer,(UINT)ATRbufferLen);
						UpdateData(FALSE);
						break;
					}
				}
				if(Result != 1)
				{
					TryTimes = 3;
					while(TryTimes)
					{
						TryTimes--;	
						Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
														1, 1,
														&StatusOut, &ErrorOut, &ChainParameterOut,
														ATRbuffer,&ATRbufferLen);	
						if(Result == 1)
						{
							m_strSlot1ATR = Bytes2CString(ATRbuffer,(UINT)ATRbufferLen);
							UpdateData(FALSE);
							break;
						}
					}
					if(Result != 1)
					{
						return;
					}
				}
				if(ATRbufferLen == 0)
				{
					return;
				}
				switch((m_ctrlSlot1CardType.GetCurSel()+1))
				{	
					case 2:
					{
						GetDlgItem(IDC_BUTTONSlot1EEPROMCardEdit)->EnableWindow(TRUE);
					}
					break;
					case 3:
					{
						GetDlgItem(IDC_BUTTONSlot1SLE4428)->EnableWindow(TRUE);	
					}
					break;
					case 4:
					{
						GetDlgItem(IDC_BUTTONSlot1SLE4442)->EnableWindow(TRUE);				
					}
					break;
					case 5:
					{
						GetDlgItem(IDC_BUTTONSlot1AT88SC)->EnableWindow(TRUE);				
					}
					break;
					case 6:
					{
						GetDlgItem(IDC_BUTTONSlot1Inphone)->EnableWindow(TRUE);			
					}
					break;
					case 7:
					{
						GetDlgItem(IDC_BUTTONSlot1AT45D041)->EnableWindow(TRUE);				
					}
					break;
				}
				//只有Asynchronous Card 才分析ATR
				if(m_ctrlSlot1CardType.GetCurSel() != 0)
				{
					return;
				}
				//分析ATR
				Slot1_ATR.Do_ATR(ATRbuffer, ATRbufferLen);
					
				switch(Slot1_ATR.T_Type)
				{
					case 1:
					{
						m_strSlot1_CardType ="T0";
						GetDlgItem(IDC_Slot1_RADIO_T0)->EnableWindow(TRUE);
						m_Slot1_Ttype = 0;
						GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(TRUE);
						GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(TRUE);
					}
					break;
					case 2:
					{
						m_strSlot1_CardType ="T1";
						GetDlgItem(IDC_Slot1_RADIO_T1)->EnableWindow(TRUE);
						m_Slot1_Ttype = 1;
						GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(TRUE);
						GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(TRUE);
					}
					break;
					case 3:
					{
						m_strSlot1_CardType ="T0/T1";
						GetDlgItem(IDC_Slot1_RADIO_T0)->EnableWindow(TRUE);
						GetDlgItem(IDC_Slot1_RADIO_T1)->EnableWindow(TRUE);
						m_Slot1_Ttype = -1;
					}
					break;
				}
				UpdateData(FALSE);
			}			
		}
		UpdateData(FALSE);
	}
*/
 
	//SetTimer(ID_CLOCK_TIMER,1000,NULL);
	CDialog::OnTimer(nIDEvent);
}

void CMy9525COMAPDlg::OnSlot0RADIOT0() 
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(FALSE);
	//  T0/T1 card
	if(Slot0_ATR.T_Type == 3) 
	{
		LPCVOID	pSendBuffer;	
		pSendBuffer	 = malloc( 5 );
		*((PUCHAR)pSendBuffer+0)=Slot0_ATR.FI_DI;
		*((PUCHAR)pSendBuffer+1)=0;
		*((PUCHAR)pSendBuffer+2)=0;
		*((PUCHAR)pSendBuffer+3)=0;
		*((PUCHAR)pSendBuffer+4)=0;

		BYTE StatusOut; 
		BYTE ErrorOut; 
		BYTE abDataOut[5]; 
		ULONG abDataOutLen;	

		BOOL	lngStatus;
		lngStatus = CMD_PC_to_RDR_SetParameters(&CSerial9525,
									0,
									0,//T0
									(PUCHAR)pSendBuffer, 5,
									&StatusOut, &ErrorOut,
									abDataOut, &abDataOutLen);
		if(lngStatus != 1)
		{
			MessageBox("Error");
		}
	}
	GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(TRUE);
}

void CMy9525COMAPDlg::OnSlot0RADIOT1() 
{
	// TODO: Add your control notification handler code here
	//  T0/T1 card
	GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(FALSE);
	if(Slot0_ATR.T_Type == 3) 
	{
		LPCVOID	pSendBuffer;	
		pSendBuffer	 = malloc( 7 );
		*((PUCHAR)pSendBuffer+0)=Slot0_ATR.FI_DI;
		*((PUCHAR)pSendBuffer+1)=0;
		*((PUCHAR)pSendBuffer+2)=0;
		*((PUCHAR)pSendBuffer+3)=0;
		*((PUCHAR)pSendBuffer+4)=0;
		*((PUCHAR)pSendBuffer+5)=0;
		*((PUCHAR)pSendBuffer+6)=0;

		BYTE StatusOut; 
		BYTE ErrorOut; 
		BYTE abDataOut[7]; 
		ULONG abDataOutLen;	

		BOOL	lngStatus;
		lngStatus = CMD_PC_to_RDR_SetParameters(&CSerial9525,
									0,
									1,//T1
									(PUCHAR)pSendBuffer, 7,
									&StatusOut, &ErrorOut,
									abDataOut, &abDataOutLen);
		if(lngStatus != 1)
		{
			MessageBox("Error");
		}
	}
	GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(TRUE);
}

void CMy9525COMAPDlg::OnSlot1RADIOT0() 
{
	// TODO: Add your control notification handler code here
	//  T0/T1 card
	GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(FALSE);
	if(Slot1_ATR.T_Type == 3) 
	{
		LPCVOID	pSendBuffer;	
		pSendBuffer	 = malloc( 5 );
		*((PUCHAR)pSendBuffer+0)=Slot1_ATR.FI_DI;
		*((PUCHAR)pSendBuffer+1)=0;
		*((PUCHAR)pSendBuffer+2)=0;
		*((PUCHAR)pSendBuffer+3)=0;
		*((PUCHAR)pSendBuffer+4)=0;

		BYTE StatusOut; 
		BYTE ErrorOut; 
		BYTE abDataOut[5]; 
		ULONG abDataOutLen;	

		BOOL	lngStatus;
		lngStatus = CMD_PC_to_RDR_SetParameters(&CSerial9525,
									1,
									0,//T0
									(PUCHAR)pSendBuffer, 5,
									&StatusOut, &ErrorOut,
									abDataOut, &abDataOutLen);
		if(lngStatus != 1)
		{
			MessageBox("Error");
		}
	}
	GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(TRUE);
}

void CMy9525COMAPDlg::OnSlot1RADIOT1() 
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(FALSE);
	//  T0/T1 card
	if(Slot1_ATR.T_Type == 3) 
	{
		LPCVOID	pSendBuffer;	
		pSendBuffer	 = malloc( 7 );
		*((PUCHAR)pSendBuffer+0)=Slot1_ATR.FI_DI;
		*((PUCHAR)pSendBuffer+1)=0;
		*((PUCHAR)pSendBuffer+2)=0;
		*((PUCHAR)pSendBuffer+3)=0;
		*((PUCHAR)pSendBuffer+4)=0;
		*((PUCHAR)pSendBuffer+5)=0;
		*((PUCHAR)pSendBuffer+6)=0;

		BYTE StatusOut; 
		BYTE ErrorOut; 
		BYTE abDataOut[7]; 
		ULONG abDataOutLen;	

		BOOL	lngStatus;
		lngStatus = CMD_PC_to_RDR_SetParameters(&CSerial9525,
									1,
									1,//T1
									(PUCHAR)pSendBuffer, 7,
									&StatusOut, &ErrorOut,
									abDataOut, &abDataOutLen);
		if(lngStatus != 1)
		{
			MessageBox("Error");
		}
	}
	GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(TRUE);
}

void CMy9525COMAPDlg::OnBUTTONSlot0Xfr() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	if(CurrentSlot!=0)
	{
		BOOL Result;
		Result = CMD_SWITCH_CARD_MODE(&CSerial9525,
								  0,
								  (m_ctrlSlot0CardType.GetCurSel() +1)
								  );	
		if(Result != 1)
		{
			MessageBox("SWITCH CARD MODE fail!");
			return;
		}
		CurrentSlot = 0;   
	}

	BYTE DataXfr[2000];
	UINT Len;
	Len = CString2Bytes (DataXfr, m_strSlot0BlockData);

	UINT SendTimes = 0;
	UINT i = 0;
	UINT AlreadySend = 0;
	UINT SendLength;

	BOOL result;
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE ChainParameterOut;
	BYTE DataOut[300];
	UINT abDataOutLen;
	UINT wLevelParameter = 0;

/*
BOOL CMD_PC_to_RDR_XfrBlock(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE BWIIn, UINT LevelParameter,
								BYTE *abDataIn, UINT abDataInLen,
							    BYTE *StatusOut, BYTE *ErrorOut, 
								BYTE *ChainParameterOut,
							    BYTE *abDataOut, UINT *abDataOutLen)
*/
	//Short APDU
	if(Len <= (Reader9525_Descriptor.BufferSize - 10))
	{
		result = CMD_PC_to_RDR_XfrBlock(&CSerial9525,
										0,
										0,wLevelParameter,
										DataXfr,
										Len,
										&StatusOut,&ErrorOut,
										&ChainParameterOut,
										DataOut,&abDataOutLen);	
		if(result!=1)
		{
			MessageBox("Xfr Failed!");
			return;
		}
	}
	//extended APDU
	//-----------------------------SEND------------------------------
	else
	{
		while(1)
		{
			if(SendTimes == 0)
			{
				wLevelParameter = 1;
				SendLength = (Reader9525_Descriptor.BufferSize - 10);
			}
			else if ((Len - AlreadySend) <= (Reader9525_Descriptor.BufferSize -10))
			{
				wLevelParameter =2;
				SendLength = (Len - AlreadySend);
			}
			else 
			{
				wLevelParameter = 3;
				SendLength = (Reader9525_Descriptor.BufferSize - 10);			
			}

			if((SendTimes != 0) && (ChainParameterOut!=0x10))
			{
				MessageBox("ChainParameter Error!");
				return;
			}
			result = CMD_PC_to_RDR_XfrBlock(&CSerial9525,
											0,
											0,wLevelParameter,
											(DataXfr+SendTimes*(Reader9525_Descriptor.BufferSize -10)),
											SendLength,
											&StatusOut,&ErrorOut,
											&ChainParameterOut,
											DataOut,&abDataOutLen);	
			AlreadySend += SendLength;
			SendTimes++;

			if(result!=1)
			{
				MessageBox("Xfr Failed!");
				return;
			}
			if(wLevelParameter == 2)
			{
				break;
			}
		}
	}
	//-----------------------------Receive------------------------------
	m_strSlot0ResponseData = "";
	while(1)
	{
		if(ChainParameterOut == 0)
		{
			m_strSlot0ResponseData = Bytes2CString(DataOut,abDataOutLen);
			UpdateData(FALSE);
			return;
		}
		m_strSlot0ResponseData += Bytes2CString(DataOut,abDataOutLen);
		m_strSlot0ResponseData += " ";
		UpdateData(FALSE);

		result = CMD_PC_to_RDR_XfrBlock(&CSerial9525,
									0,
									0,0x10,
									DataXfr,
									0,
									&StatusOut,&ErrorOut,
									&ChainParameterOut,
									DataOut,&abDataOutLen);	
		if(result!=1)
		{
			MessageBox("Xfr Failed!");
			return;
		}
		if(ChainParameterOut == 2)
		{
			break; 
		} 
	}
}

void CMy9525COMAPDlg::OnBUTTONConnect() 
{
	// TODO: Add your control notification handler code here
	UINT i;
	BOOL result;
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE DataOut[300];
	ULONG abDataOutLen;

	//Device descriptor
	result = CMD_GetReaderDescriptor(&CSerial9525,
									0,
									0,
									&StatusOut,&ErrorOut,
									DataOut,&abDataOutLen);	
	if(result != 1)
	{   
		m_TestLabel.SetBackColor(RGB(255,0,0));
		m_TestLabel.SetCaption("Bin2 : CAN NOT FIND DEVICE");
		//m_TestResult="Bin2";
		m_TestResult = WM_COM_DEVICE_UNKNOWN;
		return;
	}
//---------------------new---------------------
	Reader9525_Descriptor.BufferSize = (DataOut[2]+DataOut[3]*256);
	Reader9525_Descriptor.APDU_Type = DataOut[1];
	Reader9525_Descriptor.VID = (DataOut[9]*256 + DataOut[8]);
	Reader9525_Descriptor.PID = (DataOut[11]*256 + DataOut[10]);
//	Reader9525_Descriptor.ReleaseNumber = (DataOut[13]*256 + DataOut[12]);

//---------------------------------------------

//	Reader9525_Descriptor.BufferSize = (DataOut[1]+DataOut[2]*256);
//	Reader9525_Descriptor.APDU_Type = DataOut[3];
//	Reader9525_Descriptor.VID = (DataOut[5]*256 + DataOut[4]);
//	Reader9525_Descriptor.PID = (DataOut[7]*256 + DataOut[6]);
//	Reader9525_Descriptor.ReleaseNumber = (DataOut[9]*256 + DataOut[8]);


	m_strVIDNum = Bytes2CString((DataOut + 9),1);
	m_strVIDNum += Bytes2CString((DataOut + 8),1);
 
	m_strPIDNum = Bytes2CString((DataOut + 11),1);
	m_strPIDNum += Bytes2CString((DataOut + 10),1);
    
	
    if ((m_strVIDNum != "058F") || (m_strPIDNum != "9540"))
	{
		m_TestLabel.SetBackColor(RGB(255,255,0));
		m_TestLabel.SetCaption("Bin2 : CAN NOT FIND DEVICE");
		//m_TestResult="Bin2";
		m_TestResult = WM_COM_DEVICE_UNKNOWN;
    }
	else
	{
		m_TestLabel.SetCaption("Get Device");
		m_TestLabel.SetBackColor(RGB(0,255,0));
	}

//	m_strReleaseNum = Bytes2CString((DataOut + 13),1);
//	m_strReleaseNum += Bytes2CString((DataOut + 12),1);

	BYTE manufactureIndex = DataOut[14];
	BYTE productIndex     = DataOut[15];
	BYTE serialIndex      = DataOut[16];


	//-------------manufacture String-------------
	result = CMD_GetReaderDescriptor(&CSerial9525,
									1,
									manufactureIndex,
									&StatusOut,&ErrorOut,
									DataOut,&abDataOutLen);	
	if(result != 1)
	{
		return;
	}
	BYTE ManufactureArray[255];
	for(i=0;i<(abDataOutLen-2);i++)
	{
		ManufactureArray[i] = DataOut[2+i];
	}
	Reader9525_Descriptor.ManufactureString = Bytes2CString_ASCII(ManufactureArray,(abDataOutLen-2));
	m_strVID = Reader9525_Descriptor.ManufactureString;
	//-------------product string-------------
	result = CMD_GetReaderDescriptor(&CSerial9525,
									1,
									productIndex,
									&StatusOut,&ErrorOut,
									DataOut,&abDataOutLen);	
	if(result != 1)
	{
		return;
	}
	BYTE ProductArray[255];
	for(i=0;i<(abDataOutLen-2);i++)
	{
		ProductArray[i] = DataOut[2+i];
	}
	Reader9525_Descriptor.ProductString = Bytes2CString_ASCII(ProductArray,(abDataOutLen-2));
	m_strPID = Reader9525_Descriptor.ProductString;
	//-------------serial number-------------
	if(serialIndex!=0)
	{
		result = CMD_GetReaderDescriptor(&CSerial9525,
										1,
										serialIndex,
										&StatusOut,&ErrorOut,
										DataOut,&abDataOutLen);	
		if(result != 1)
		{
			return;
		}
//		Reader9525_Descriptor.SerialNumber = DataOut[2]*256 + DataOut[3];
		CString strtemp = "";
		m_strSerialNum = "";
		strtemp.Format("%c",DataOut[2]); 
		m_strSerialNum += strtemp;
		strtemp.Format("%c",DataOut[4]); 
		m_strSerialNum += strtemp;
		strtemp.Format("%c",DataOut[6]); 
		m_strSerialNum += strtemp;
		strtemp.Format("%c",DataOut[8]); 
		m_strSerialNum += strtemp;

//		m_strSerialNum = Bytes2CString((DataOut + 2),1);
//		m_strSerialNum += Bytes2CString((DataOut + 3),1);
	}
/*
	BYTE VID_len;
	BYTE VID_Array[255];
	BYTE PID_len;	
	BYTE PID_Array[255];

	VID_len = GetData[13];
	for(i=0;i<VID_len;i++)
	{
		VID_Array[i] = GetData[14+i];
	}
	m_ctrlReader_attribute->PID = Bytes2CString_ASCII(VID_Array,VID_len);

	PID_len = GetData[14+VID_len];
	for(i=0;i<PID_len;i++)
	{

		PID_Array[i] = GetData[15+VID_len+i];
	}
	m_ctrlReader_attribute->VID = Bytes2CString_ASCII(PID_Array,PID_len);*/
	UpdateData(FALSE);	
}

void CMy9525COMAPDlg::OnSelchangeComboCom() 
{
	// TODO: Add your control notification handler code here
	if(CSerial9525.m_bOpened == TRUE)
	{
		CSerial9525.ClosePort();
	}

	switch(m_ctrlCOM.GetCurSel())
	{
		case 0:
		{
			if(!CSerial9525.Open(1,BaudRate))
			{
				MessageBox("Com occupyed!");
			}
		}
		break;
		case 1:
		{
			if(!CSerial9525.Open(2,BaudRate))
			{
				MessageBox("Com occupyed!");
			}
		}
		break;
		case 2:
		{
			if(!CSerial9525.Open(3,BaudRate))
			{
				MessageBox("Com occupyed!");
			}
		}
		break;
		case 3:
		{
			if(!CSerial9525.Open(4,BaudRate))
			{
				MessageBox("Com occupyed!");
			}
		}
		break;
	}	
}

void CMy9525COMAPDlg::OnSelchangeCOMBOBaudRate() 
{
	// TODO: Add your control notification handler code here
	switch(m_ctrlBaudRate.GetCurSel())
	{
		case 0:
		{
			BaudRate = 4800;
		}
		break;
		case 1:
		{
			BaudRate = 9600;
		}
		break;
		case 2:
		{
			BaudRate = 19200;
		}
		break;
		case 3:
		{
			BaudRate = 38400;
		}
		break;
/*		case 4:
		{
			BaudRate = 57600;
		}
		break;
		case 5:
		{
			BaudRate = 76800;
		}
		break;
		case 6:
		{
			BaudRate = 115200;
		}*/
		break;
	}	
	if(CMD_SetBaudRate(&CSerial9525,m_ctrlBaudRate.GetCurSel()))
	{
		if(CSerial9525.m_bOpened == TRUE)
		{
			CSerial9525.ClosePort();
		}
		CSerial9525.Open((m_ctrlCOM.GetCurSel() + 1), BaudRate);
	}
	else
	{
		MessageBox("Fail to change the Baud rate!");
	}
	
}

void CMy9525COMAPDlg::OnSelchangeCOMBOSlot0CardType() 
{
	// TODO: Add your control notification handler code here
/*	LONG Result;
	Result = CMD_SWITCH_CARD_MODE(&CSerial9525,
							  0,
							  (m_ctrlSlot0CardType.GetCurSel() +1)
							  );	
	if(Result != 1)
	{
		MessageBox("SWITCH CARD MODE fail!");
	}*/
}

void CMy9525COMAPDlg::OnBUTTONSlot1Xfr() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	if(CurrentSlot!=1)
	{
		BOOL Result;
		Result = CMD_SWITCH_CARD_MODE(&CSerial9525,
								  1,
								  (m_ctrlSlot1CardType.GetCurSel() +1)
								  );	
		if(Result != 1)
		{
			MessageBox("SWITCH CARD MODE fail!");
			return;
		}
		CurrentSlot = 1;
	}

	BYTE DataXfr[2000];
	UINT Len;
	Len = CString2Bytes (DataXfr, m_strSlot1BlockData);

	UINT SendTimes = 0;
	UINT i = 0;
	UINT AlreadySend = 0;
	UINT SendLength;

	BOOL result;
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE ChainParameterOut;
	BYTE DataOut[300];
	UINT abDataOutLen;
	UINT wLevelParameter = 0;
	//Short APDU
	if(Len <= (Reader9525_Descriptor.BufferSize - 10))
	{
		result = CMD_PC_to_RDR_XfrBlock(&CSerial9525,
										1,
										0,wLevelParameter,
										DataXfr,
										Len,
										&StatusOut,&ErrorOut,
										&ChainParameterOut,
										DataOut,&abDataOutLen);	
		if(result!=1)
		{
			MessageBox("Xfr Failed!");
			return;
		}
	}
	//extended APDU
	//-----------------------------SEND------------------------------
	else
	{
		while(1)
		{
			if(SendTimes == 0)
			{
				wLevelParameter = 1;
				SendLength = (Reader9525_Descriptor.BufferSize - 10);
			}
			else if ((Len - AlreadySend) <= (Reader9525_Descriptor.BufferSize -10))
			{
				wLevelParameter =2;
				SendLength = (Len - AlreadySend);
			}
			else 
			{
				wLevelParameter = 3;
				SendLength = (Reader9525_Descriptor.BufferSize - 10);			
			}

			if((SendTimes != 0) && (ChainParameterOut!=0x10))
			{
				MessageBox("ChainParameter Error!");
				return;
			}
			result = CMD_PC_to_RDR_XfrBlock(&CSerial9525,
											1,
											0,wLevelParameter,
											(DataXfr+SendTimes*(Reader9525_Descriptor.BufferSize - 10)),
											SendLength,
											&StatusOut,&ErrorOut,
											&ChainParameterOut,
											DataOut,&abDataOutLen);	
			AlreadySend += SendLength;
			SendTimes++;

			if(result!=1)
			{
				MessageBox("Xfr Failed!");
				return;
			}
			if(wLevelParameter == 2)
			{
				break;
			}
		}
	}
	//-----------------------------Receive------------------------------
	m_strSlot0ResponseData = "";
	while(1)
	{
		if(ChainParameterOut == 0)
		{
			m_strSlot1ResponseData = Bytes2CString(DataOut,abDataOutLen);
			UpdateData(FALSE);
			return;
		}
		m_strSlot1ResponseData += Bytes2CString(DataOut,abDataOutLen);
		m_strSlot1ResponseData += " ";
		UpdateData(FALSE);

		result = CMD_PC_to_RDR_XfrBlock(&CSerial9525,
									1,
									0,0x10,
									DataXfr,
									0,
									&StatusOut,&ErrorOut,
									&ChainParameterOut,
									DataOut,&abDataOutLen);	
		if(result!=1)
		{
			MessageBox("Xfr Failed!");
			return;
		}
		if(ChainParameterOut == 2)
		{
			break;
		}
	}	
}

void CMy9525COMAPDlg::OnButtonLcmKeypadEeprom() 
{
	// TODO: Add your control notification handler code here
	m_dLCM_KEYPAD_EEPROMDlg.CSerial9525 = CSerial9525;
	m_dLCM_KEYPAD_EEPROMDlg.DoModal();
}

void CMy9525COMAPDlg::OnBUTTONSlot0SLE4428() 
{
	// TODO: Add your control notification handler code here
	m_dSLE4428Dlg.SlotNum = 0;
	m_dSLE4428Dlg.CSerial9525 = CSerial9525;	
	m_dSLE4428Dlg.DoModal();
}

void CMy9525COMAPDlg::OnBUTTONSlot1SLE4428() 
{
	// TODO: Add your control notification handler code here
	m_dSLE4428Dlg.SlotNum = 1;
	m_dSLE4428Dlg.CSerial9525 = CSerial9525;
	m_dSLE4428Dlg.DoModal();
}

void CMy9525COMAPDlg::OnBUTTONSlot0SLE4442() 
{
	// TODO: Add your control notification handler code here
	m_dSLE4442Dlg.SlotNum = 0;
	m_dSLE4442Dlg.CSerial9525 = CSerial9525;
	m_dSLE4442Dlg.DoModal();
}

void CMy9525COMAPDlg::OnBUTTONSlot1SLE4442() 
{
	// TODO: Add your control notification handler code here
	m_dSLE4442Dlg.SlotNum = 1;
	m_dSLE4442Dlg.CSerial9525 = CSerial9525;
	m_dSLE4442Dlg.DoModal();
}

void CMy9525COMAPDlg::OnBUTTONSlot0AT45D041() 
{
	// TODO: Add your control notification handler code here
	m_dAT45D041Dlg.SlotNum = 0;
	m_dAT45D041Dlg.CSerial9525 = CSerial9525;
	m_dAT45D041Dlg.DoModal();
}

void CMy9525COMAPDlg::OnBUTTONSlot1AT45D041() 
{
	// TODO: Add your control notification handler code here
	m_dAT45D041Dlg.SlotNum = 1;
	m_dAT45D041Dlg.CSerial9525 = CSerial9525;
	m_dAT45D041Dlg.DoModal();
}

void CMy9525COMAPDlg::OnBUTTONSlot0AT88SC() 
{
	// TODO: Add your control notification handler code here	
	m_dAT88SCDlg.SlotNum = 0;
	m_dAT88SCDlg.CSerial9525 = CSerial9525;
	m_dAT88SCDlg.DoModal();	
}

void CMy9525COMAPDlg::OnBUTTONSlot0EEPROMCardEdit() 
{
	// TODO: Add your control notification handler code here
	m_dAT24CDlg.SlotNum = 0;
	m_dAT24CDlg.CSerial9525 = CSerial9525;
	m_dAT24CDlg.DoModal();		
}

void CMy9525COMAPDlg::OnBUTTONSlot1EEPROMCardEdit() 
{
	// TODO: Add your control notification handler code here
	m_dAT24CDlg.SlotNum = 1;
	m_dAT24CDlg.CSerial9525 = CSerial9525;
	m_dAT24CDlg.DoModal();		
}

void CMy9525COMAPDlg::OnBUTTONSlot1AT88SC() 
{
	// TODO: Add your control notification handler code here
	m_dAT88SCDlg.SlotNum = 1;
	m_dAT88SCDlg.CSerial9525 = CSerial9525;
	m_dAT88SCDlg.DoModal();		
}

void CMy9525COMAPDlg::OnBUTTONSlot0Inphone() 
{
	// TODO: Add your control notification handler code here
	m_dCInphoneCmdDlg.SlotNum = 0;
	m_dCInphoneCmdDlg.CSerial9525 = CSerial9525;
	m_dCInphoneCmdDlg.DoModal();		
}

void CMy9525COMAPDlg::OnBUTTONSlot1Inphone() 
{
	// TODO: Add your control notification handler code here
	m_dCInphoneCmdDlg.SlotNum = 1;
	m_dCInphoneCmdDlg.CSerial9525 = CSerial9525;
	m_dCInphoneCmdDlg.DoModal();	
}

void CMy9525COMAPDlg::TestSub()
{
	// TODO: Add your message handler code here and/or call default
	//KillTimer(ID_CLOCK_TIMER);
	BYTE bmSlotICCStateOut;
 //   MessageBox("qq");
//	if(Check_RDR_to_PC_NotifySlotChange(&CSerial9525,&bmSlotICCStateOut)!=0)
	 if (1)  // Allen debug(1)
	{
	 //	if((bmSlotICCStateOut&0x02) == 0x02)//Slot0 change
		 if (1)  // Allen debug(2) 
		{
		//	if((bmSlotICCStateOut&0x01) == 0x00)//Card out
            if (0)
			{
				m_strSlot0ATR = "";
				m_strSlot0_CardType = "";
				m_strSlot0BlockData = "";
				m_strSlot0ResponseData = "";
				m_Slot0_Ttype = -1;
				GetDlgItem(IDC_Slot0_RADIO_T0)->EnableWindow(FALSE);
				GetDlgItem(IDC_Slot0_RADIO_T1)->EnableWindow(FALSE);
				CurrentSlot = 0xFF;
				GetDlgItem(IDC_BUTTONSlot0SLE4428)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0SLE4442)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0AT88SC)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0AT45D041)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0EEPROMCardEdit)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot0Inphone)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(FALSE);
				GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(FALSE);
			}
			else//Card in
			{
				if(m_ctrlSlot0CardType.GetCurSel() == -1)
				{
					return;
				}
				//Power off
				BYTE TryTimes = 3;
				UINT Result;
				BYTE StatusOut;
				BYTE ErrorOut; 
				BYTE ClockStatusOut;

				Result = CMD_PC_to_RDR_IccPowerOff(&CSerial9525,
												0,
												&StatusOut, &ErrorOut, &ClockStatusOut);
				if(Result != 1)
				{
					return;
				}
				Result = CMD_SWITCH_CARD_MODE(&CSerial9525,
										  0,
										  (m_ctrlSlot0CardType.GetCurSel() +1)
										  );	
				if(Result != 1)
				{
					MessageBox("SWITCH CARD MODE fail!");
					return;
				}
				CurrentSlot = 0;

				//-----Power on-----
				//UINT Result;
				//BYTE StatusOut;
				//BYTE ErrorOut;
				BYTE ChainParameterOut;
				BYTE ATRbuffer[100];   
				BYTE ATRbufferLen;
				TryTimes = 3;
//				while(TryTimes)
//				{
//					TryTimes--;	
					Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
													0, 2,
													&StatusOut, &ErrorOut, &ChainParameterOut,
													ATRbuffer,&ATRbufferLen);	
					if(Result == 1)
					{
						m_strSlot0ATR = Bytes2CString(ATRbuffer,(UINT)ATRbufferLen);
						UpdateData(FALSE);
//						break;
					}
//				}
				if(Result != 1)
				{
//					TryTimes = 3;
//					while(TryTimes)
//					{
//						TryTimes--;	
						Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
														0, 1,
														&StatusOut, &ErrorOut, &ChainParameterOut,
														ATRbuffer,&ATRbufferLen);	
						if(Result == 1)
						{
							m_strSlot0ATR = Bytes2CString(ATRbuffer,(UINT)ATRbufferLen);
							UpdateData(FALSE);
//							break;
						}
//					}
					if(Result != 1)
					{
						return;
					}
				}	
				if(ATRbufferLen == 0)
				{
					return;
				}
				switch((m_ctrlSlot0CardType.GetCurSel()+1))
				{	
					case 2:
					{
						GetDlgItem(IDC_BUTTONSlot0EEPROMCardEdit)->EnableWindow(TRUE);
					}
					break;
					case 3:
					{
						GetDlgItem(IDC_BUTTONSlot0SLE4428)->EnableWindow(TRUE);	
					}
					break;
					case 4:
					{
						GetDlgItem(IDC_BUTTONSlot0SLE4442)->EnableWindow(TRUE);				
					}
					break;
					case 5:
					{
						GetDlgItem(IDC_BUTTONSlot0AT88SC)->EnableWindow(TRUE);				
					}
					break;
					case 6:
					{
						GetDlgItem(IDC_BUTTONSlot0Inphone)->EnableWindow(TRUE);			
					}
					break;				
					case 7:
					{
						GetDlgItem(IDC_BUTTONSlot0AT45D041)->EnableWindow(TRUE);				
					}
					break;
				}

				//只有Asynchronous Card 才分析ATR
				if(m_ctrlSlot0CardType.GetCurSel() != 0)
				{
					return;
				}
				//分析ATR
				Slot0_ATR.Do_ATR(ATRbuffer, ATRbufferLen);
					
				switch(Slot0_ATR.T_Type)
				{
					case 1:
					{
						m_strSlot0_CardType ="T0";
						GetDlgItem(IDC_Slot0_RADIO_T0)->EnableWindow(TRUE);
						m_Slot0_Ttype = 0;
						GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(TRUE);
						GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(TRUE);
					}
					break;
					case 2:
					{
						m_strSlot0_CardType ="T1";
						GetDlgItem(IDC_Slot0_RADIO_T1)->EnableWindow(TRUE);
						m_Slot0_Ttype = 1;						
						GetDlgItem(IDC_Slot0_BlockData)->EnableWindow(TRUE);
						GetDlgItem(IDC_BUTTON_Slot0_Xfr)->EnableWindow(TRUE);
					}
					break;
					case 3:
					{
						m_strSlot0_CardType ="T0/T1";
						GetDlgItem(IDC_Slot0_RADIO_T0)->EnableWindow(TRUE);
						GetDlgItem(IDC_Slot0_RADIO_T1)->EnableWindow(TRUE);
						m_Slot0_Ttype = -1;
					}
					break;
				}
				UpdateData(FALSE);
			}
		}

	//	if((bmSlotICCStateOut&0x08) == 0x08)//Slot1 change
	    
		// SKIP Slot1 test for AU9540 
		if (0)  //Allen
		{
	//		if((bmSlotICCStateOut&0x04) == 0x00)//Card out
			 if(0)  //Allen
			{
				m_strSlot1ATR = "";
				m_strSlot1_CardType = "";
				m_strSlot1BlockData = "";
				m_strSlot1ResponseData = "";
				m_Slot1_Ttype = -1;
				GetDlgItem(IDC_Slot1_RADIO_T0)->EnableWindow(FALSE);
				GetDlgItem(IDC_Slot1_RADIO_T1)->EnableWindow(FALSE);
				CurrentSlot = 0xFF;
				GetDlgItem(IDC_BUTTONSlot1SLE4428)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1SLE4442)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1AT88SC)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1AT45D041)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1EEPROMCardEdit)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTONSlot1Inphone)->EnableWindow(FALSE);
				GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(FALSE);	
				GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(FALSE);
			}
			else//Card in
			{
				if(m_ctrlSlot1CardType.GetCurSel() == -1)
				{
					return;
				}
				//Power off
				BYTE TryTimes = 3;
				UINT Result;
				BYTE StatusOut;
				BYTE ErrorOut; 
				BYTE ClockStatusOut;
				while(TryTimes)
				{
					TryTimes--;	
					Result = CMD_PC_to_RDR_IccPowerOff(&CSerial9525,
												1,
												&StatusOut, &ErrorOut, &ClockStatusOut);
					if(Result == 1)
					{
						break;
					}
				}
				if(Result != 1)
				{
					return;
				}
				Result = CMD_SWITCH_CARD_MODE(&CSerial9525,
										  1,
										  (m_ctrlSlot1CardType.GetCurSel() +1)
										  );	
				if(Result != 1)
				{
					MessageBox("SWITCH CARD MODE fail!");
				}
				CurrentSlot = 1;
				//Power on
				//UINT Result;
				//BYTE StatusOut;
				//BYTE ErrorOut;
				BYTE ChainParameterOut;
				BYTE ATRbuffer[100];   
				BYTE ATRbufferLen;
				TryTimes = 3;
				while(TryTimes)
				{
					TryTimes--;	
					Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
													1, 2,
													&StatusOut, &ErrorOut, &ChainParameterOut,
													ATRbuffer,&ATRbufferLen);	
					if(Result == 1)
					{
						m_strSlot1ATR = Bytes2CString(ATRbuffer,(UINT)ATRbufferLen);
						UpdateData(FALSE);
						break;
					}
				}
				if(Result != 1)
				{
					TryTimes = 3;
					while(TryTimes)
					{
						TryTimes--;	
						Result = CMD_PC_to_RDR_IccPowerOn(&CSerial9525,
														1, 1,
														&StatusOut, &ErrorOut, &ChainParameterOut,
														ATRbuffer,&ATRbufferLen);	
						if(Result == 1)
						{
							m_strSlot1ATR = Bytes2CString(ATRbuffer,(UINT)ATRbufferLen);
							UpdateData(FALSE);
							break;
						}
					}
					if(Result != 1)
					{
						return;
					}
				}
				if(ATRbufferLen == 0)
				{
					return;
				}
				switch((m_ctrlSlot1CardType.GetCurSel()+1))
				{	
					case 2:
					{
						GetDlgItem(IDC_BUTTONSlot1EEPROMCardEdit)->EnableWindow(TRUE);
					}
					break;
					case 3:
					{
						GetDlgItem(IDC_BUTTONSlot1SLE4428)->EnableWindow(TRUE);	
					}
					break;
					case 4:
					{
						GetDlgItem(IDC_BUTTONSlot1SLE4442)->EnableWindow(TRUE);				
					}
					break;
					case 5:
					{
						GetDlgItem(IDC_BUTTONSlot1AT88SC)->EnableWindow(TRUE);				
					}
					break;
					case 6:
					{
						GetDlgItem(IDC_BUTTONSlot1Inphone)->EnableWindow(TRUE);			
					}
					break;
					case 7:
					{
						GetDlgItem(IDC_BUTTONSlot1AT45D041)->EnableWindow(TRUE);				
					}
					break;
				}
				//只有Asynchronous Card 才分析ATR
				if(m_ctrlSlot1CardType.GetCurSel() != 0)
				{
					return;
				}
				//分析ATR
				Slot1_ATR.Do_ATR(ATRbuffer, ATRbufferLen);
					
				switch(Slot1_ATR.T_Type)
				{
					case 1:
					{
						m_strSlot1_CardType ="T0";
						GetDlgItem(IDC_Slot1_RADIO_T0)->EnableWindow(TRUE);
						m_Slot1_Ttype = 0;
						GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(TRUE);
						GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(TRUE);
					}
					break;
					case 2:
					{
						m_strSlot1_CardType ="T1";
						GetDlgItem(IDC_Slot1_RADIO_T1)->EnableWindow(TRUE);
						m_Slot1_Ttype = 1;
						GetDlgItem(IDC_Slot1_BlockData)->EnableWindow(TRUE);
						GetDlgItem(IDC_BUTTON_Slot1_Xfr)->EnableWindow(TRUE);
					}
					break;
					case 3:
					{
						m_strSlot1_CardType ="T0/T1";
						GetDlgItem(IDC_Slot1_RADIO_T0)->EnableWindow(TRUE);
						GetDlgItem(IDC_Slot1_RADIO_T1)->EnableWindow(TRUE);
						m_Slot1_Ttype = -1;
					}
					break;
				}
				UpdateData(FALSE);
			}			
		}
		UpdateData(FALSE);
	}
}

void CMy9525COMAPDlg::OnBeginTest() 
{
	// TODO: Add your control notification handler code here

	UpdateData(FALSE);
	m_TestResult = 0;
	OnBUTTONConnect();
	Disp();
	Sleep(1000);
	if (m_TestResult == WM_COM_DEVICE_UNKNOWN)
	{
		winHnd = ::FindWindow(NULL, "ALCOR TESTER");
		if (winHnd != NULL)
			::PostMessage(winHnd, WM_COM_DEVICE_UNKNOWN, 0, 0);
	}
	else
	{
		//TestFlag=0;
		TestSub();

		if(m_strSlot0ATR  == Pattern)
		{
			m_TestLabel.SetCaption("Slot 0 R/W Pass");
			m_TestLabel.SetBackColor(RGB(0,255,0));

			winHnd = ::FindWindow(NULL, "ALCOR TESTER");
			if (winHnd != NULL)
				::PostMessage(winHnd, WM_COM_TEST_PASS, 0, 0);
		}
		else
		{
       		m_TestLabel.SetCaption("Bin3: Slot 0 R/W Fail");
			m_TestLabel.SetBackColor(RGB(255,0,0));

			winHnd = ::FindWindow(NULL, "ALCOR TESTER");
			if (winHnd != NULL)
				::PostMessage(winHnd, WM_COM_TEST_FAIL, 0, 0);

		}
	}
	ECHO = 0;
	SetTimer(55, 12000, NULL);
}

void CMy9525COMAPDlg::TestInit()
{
  

	    /*if (m_Comm.GetPortOpen())
		{
       		m_Comm.SetPortOpen(0);
		}

		m_Comm.SetCommPort((short)1);    // set com2 to test
		m_Comm.SetSettings("9600,n,8,1");
 
		if (!m_Comm.GetPortOpen())
		{
	 
			m_Comm.SetPortOpen(1);

		}
		 

    	m_Comm.SetInBufferCount(0);
	    m_Comm.SetOutBufferCount(0);*/
    

          
	   		//PCI_7248 initialize
 
		if ((card=Register_Card(PCI_7248, card_number)) < 0) 
		{
			MessageBox("Error Register Card");
		}
		DIO_PortConfig(card,  Channel_P1A, OUTPUT_PORT);
		DIO_PortConfig(card,  Channel_P1B, INPUT_PORT);
		DIO_PortConfig(card,  Channel_P1CL, OUTPUT_PORT);
		DIO_PortConfig(card,  Channel_P1CH, INPUT_PORT);
		DIO_PortConfig(card,  Channel_P2A, OUTPUT_PORT);
		DIO_PortConfig(card,  Channel_P2B, INPUT_PORT);
		DIO_PortConfig(card,  Channel_P2CL, OUTPUT_PORT);
		DIO_PortConfig(card,  Channel_P2CH, INPUT_PORT);
}

BOOL CMy9525COMAPDlg::DestroyWindow() 
{
	// TODO: Add your specialized code here and/or call the base class
	    AllenStop=1;
		EndDialog(IDOK);
	return CDialog::DestroyWindow();
}

void CMy9525COMAPDlg::OnEXit() 
{
	// TODO: Add your control notification handler code here
	 AllenStop=1;
		EndDialog(IDOK);
	
}

void CMy9525COMAPDlg::Disp()
{
	m_strVIDNum=_T("");
	m_strPIDNum=_T("");
	UpdateData(FALSE);

	m_TestLabel.SetCaption("Begin Test");
	m_TestLabel.SetBackColor(RGB(255,255,255));
	m_strSlot0ATR="";
	m_strSlot1ATR="";
	UpdateData(FALSE);
	UpdateWindow();
	//Sleep(100);
	//OnPaint();
}

LRESULT CMy9525COMAPDlg::DefWindowProc(UINT message, WPARAM wParam, LPARAM lParam) 
{
	// TODO: Add your specialized code here and/or call the base class

	//ECHO = 0;

	if(message == WM_COM_START_TEST)
	{
		ECHO = 1;
		OnBeginTest();
	}

	if(message == WM_COM_CLOSE)
	{
		ECHO = 1;
		OnEXit();
	}

	return CDialog::DefWindowProc(message, wParam, lParam);
}

void CMy9525COMAPDlg::OnClose() 
{
	// TODO: Add your message handler code here and/or call default
	
	CDialog::OnClose();
}
