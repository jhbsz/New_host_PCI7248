// AU9525Tester.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "AU9525Tester.h"
#include "Dask.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// AU9525Tester dialog

BYTE AllenStop=0;

I16 card=-1, card_number = 0;
int BINTimer = 0;

BYTE PCI7248_EOT =	0x01;		//for 7248 card
BYTE PCI7248_PASS = 0xFD;		//for 7248 card  11111101
BYTE PCI7248_BIN2 = 0xFB;		//for 7248 card  11111011
BYTE PCI7248_BIN3 = 0xF7;		//for 7248 card  11110111
BYTE PCI7248_BIN4 = 0xEF;		//for 7248 card  11101111
BYTE PCI7248_BIN5 = 0xDF;		//for 7248 card  11011111

BYTE EOT1 = 0x20; //for old card
BYTE EOT2 = 0x10; //for old card
BYTE RESET_LATCH = 0x3F;


AU9525Tester::AU9525Tester(CWnd* pParent /*=NULL*/)
	: CDialog(AU9525Tester::IDD, pParent)
{
	//{{AFX_DATA_INIT(AU9525Tester)
	//}}AFX_DATA_INIT
}


void AU9525Tester::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(AU9525Tester)
	DDX_Control(pDX, IDC_MSCOMM1, m_Comm);
	DDX_Control(pDX, IDC_Slot0, m_Slot0);
	DDX_Control(pDX, IDC_Slot1, m_Slot1);
	DDX_Control(pDX, IDC_TestResult, m_TestResult);
	DDX_Control(pDX, IDC_UnknowDevice, m_UnknowDevice);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(AU9525Tester, CDialog)
	//{{AFX_MSG_MAP(AU9525Tester)
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_StartTest, OnStartTest)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// AU9525Tester message handlers

BOOL AU9525Tester::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	/* nick marked for disable com port and 7248
	TestInit2();*/

	// TODO: Add extra initialization here
	SetTimer(20,500,FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void AU9525Tester::OnTimer(UINT nIDEvent) 
{
	// TODO: Add your message handler code here and/or call default
	
    if ( nIDEvent == 20 )
	{
		KillTimer(20);
		AU9525COMAdlg.DoModal();
		
	}
 

	CDialog::OnTimer(nIDEvent);
}

void AU9525Tester::OnStartTest() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
 
//	m_strResult = "";
//	m_strTestInformation = "";	
	UpdateData(FALSE);
//	GetDlgItem(IDC_BUTTON_START)->EnableWindow(FALSE);
 
//	SetTimer(ID_CLOCK,10000,NULL);
 
	long t1;   	// test cycle
 
	// EndLess contorl loop
    CString StrOutput;
	StrOutput="Ready";
	CString ChipName,StrInput;
	 MSG msg;
	COleVariant Buf;
	do{ 
		time(&t1);  // get time
      
		ChipName="";

		//============================================
		//     Get ChipName
		//============================================
		
		do
		{

				m_Comm.SetOutput(COleVariant(StrOutput)); // Send Ready 
				Sleep(10);
				Buf=m_Comm.GetInput();  // get chip name  
				ChipName=ChipName + (CString)(Buf.bstrVal);

				// DoEvents
				while( ::PeekMessage( &msg, NULL, 0, 0, PM_NOREMOVE ) ) 
				{
					  ::GetMessage( &msg, NULL, 0, 0 );
					  ::TranslateMessage( &msg );
					  ::DispatchMessage( &msg );
				} 
		} while ((ChipName!="AU6710ASF22") && (AllenStop==0));
      
	//	m_SpeedError.SetBkColor(RGB(255,255,255));   // initail interface
	//	m_LedFail.SetBkColor(RGB(255,255,255));
		m_UnknowDevice.SetBackColor(RGB(255,255,255));
		m_Slot0.SetBackColor(RGB(255,255,255));
		m_Slot1.SetBackColor(RGB(255,255,255));
		m_TestResult.SetBackColor(RGB(255,255,255));
		

		 
		m_TestResult.SetWindowText("Wait for Testing");  // set label UI
//	 	m_strResult = "";
 //    	m_strTestInformation = "";
 //  
         DO_WritePort(card, Channel_P1A, 0x0);  // set power on
         Sleep(2000);
  
		 StartTest();
//	     m_Infor.LineScroll(m_Infor.GetLineCount());
         DO_WritePort(card, Channel_P1A, 0xFF);

   
		 while( ::PeekMessage( &msg, NULL, 0, 0, PM_NOREMOVE ) ) 
		 {
		  ::GetMessage( &msg, NULL, 0, 0 );
		  ::TranslateMessage( &msg );
		  ::DispatchMessage( &msg );
		 } 
	
	}  while (AllenStop==0);
	
}

void AU9525Tester::StartTest()
{

}



void AU9525Tester::TestInitial()
{
        //set  com port


}

void AU9525Tester::TestInit2()
{
    
	m_Comm.SetCommPort((short)1);
	m_Comm.SetSettings("9600,n,8,1");
//	m_Comm.SetPortOpen(0);
//	Sleep(1000);
	if (!m_Comm.GetPortOpen())
	{
		m_Comm.SetPortOpen(1);
	}
	 

	m_Comm.SetInBufferCount(0);
	m_Comm.SetOutBufferCount(0);
    

          
	   		//PCI_7248 initialize
	/*
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
*/
//		SetTimer_1ms();
		
   //		m_SpeedError.SetBkColor(RGB(255,255,255));
//		m_LedFail.SetBkColor(RGB(255,255,255));
		m_Slot0.SetBackColor(RGB(255,255,255));
		m_Slot1.SetBackColor(RGB(255,255,255));
		m_TestResult.SetBackColor(RGB(255,255,255));
		m_UnknowDevice.SetBackColor(RGB(255,255,255));
}
