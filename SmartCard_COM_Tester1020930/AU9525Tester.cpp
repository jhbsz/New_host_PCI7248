// AU9525Tester.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "AU9525Tester.h"
//#include "Dask.h"
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
//	SetTimer(20,500,FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void AU9525Tester::OnTimer(UINT nIDEvent) 
{
	// TODO: Add your message handler code here and/or call default
	
    if ( nIDEvent == 20 )
	{
		//KillTimer(20);
		AU9525COMAdlg.DoModal();
		
	}
 

	CDialog::OnTimer(nIDEvent);
}
