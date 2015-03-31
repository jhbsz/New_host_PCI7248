// SLE4442CMDDlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "SLE4442CMDDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSLE4442CMDDlg dialog


CSLE4442CMDDlg::CSLE4442CMDDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CSLE4442CMDDlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CSLE4442CMDDlg)
	m_STRReadDataLength = _T("00 00");
	m_strProtectBitFlag = _T("00");
	m_strClockNumberFlag = _T("00");
	m_strCommand2 = _T("00");
	m_strCommand1 = _T("00");
	m_strCommand3 = _T("00");
	m_strResponse = _T("00");
	//}}AFX_DATA_INIT
}


void CSLE4442CMDDlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CSLE4442CMDDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CSLE4442CMDDlg)
	DDX_Text(pDX, IDC_EDIT_ReadDataLength, m_STRReadDataLength);
	DDX_Text(pDX, IDC_EDIT_ProtectBitFlag, m_strProtectBitFlag);
	DDX_Text(pDX, IDC_EDIT_ClockNumberFlag, m_strClockNumberFlag);
	DDX_Text(pDX, IDC_EDIT_Command2, m_strCommand2);
	DDX_Text(pDX, IDC_EDIT_Command1, m_strCommand1);
	DDX_Text(pDX, IDC_EDIT_Command3, m_strCommand3);
	DDX_Text(pDX, IDC_EDIT_Response, m_strResponse);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CSLE4442CMDDlg, CDialog)
	//{{AFX_MSG_MAP(CSLE4442CMDDlg)
	ON_BN_CLICKED(IDC_BUTTON_SLE4442_CARD_BREAK, OnButtonSle4442CardBreak)
	ON_BN_CLICKED(IDC_BUTTON_SLE4442_CARD_COMMAND, OnButtonSle4442CardCommand)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CSLE4442CMDDlg, CDialog)
	//{{AFX_DISPATCH_MAP(CSLE4442CMDDlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ISLE4442CMDDlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {B522F5CB-B697-4B54-9873-6D85AD016A79}
static const IID IID_ISLE4442CMDDlg =
{ 0xb522f5cb, 0xb697, 0x4b54, { 0x98, 0x73, 0x6d, 0x85, 0xad, 0x1, 0x6a, 0x79 } };

BEGIN_INTERFACE_MAP(CSLE4442CMDDlg, CDialog)
	INTERFACE_PART(CSLE4442CMDDlg, IID_ISLE4442CMDDlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSLE4442CMDDlg message handlers

void CSLE4442CMDDlg::OnButtonSle4442CardBreak() 
{
	// TODO: Add your control notification handler code here
	LONG	lngStatus;
	lngStatus = CMD_SLE4442_CARD_BREAK(	IN	&CSerial9525,
										IN	SlotNum);
	if(lngStatus!=1)
	{
		MessageBox("SLE4442 CARD BREAK fail!");
	}	
}

void CSLE4442CMDDlg::OnButtonSle4442CardCommand() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	LONG	lngStatus;
	
	BYTE ReadingLengthArray[2];
	UINT Len;
	Len = CString2Bytes(ReadingLengthArray, m_STRReadDataLength);
	UINT ReadingLength; 
	ReadingLength = (ReadingLengthArray[0]*256 + ReadingLengthArray[1]);

	BYTE ProtectBit;
	Len = CString2Bytes(&ProtectBit, m_strProtectBitFlag);

	BYTE AskForClkNum;
	Len = CString2Bytes(&AskForClkNum, m_strClockNumberFlag);
	BYTE CMD_1;
	Len = CString2Bytes(&CMD_1, m_strCommand1);
	BYTE CMD_2;
	Len = CString2Bytes(&CMD_2, m_strCommand2);
	BYTE CMD_3;
	Len = CString2Bytes(&CMD_3, m_strCommand3);

	BYTE GetData[300];
	ULONG GetDataLen;
	lngStatus = CMD_SLE4442_CARD_COMMAND(IN	&CSerial9525,
									     IN	SlotNum,
										 IN	ReadingLength,
										 IN ProtectBit,
										 IN AskForClkNum,
										 IN CMD_1,
										 IN CMD_2,
										 IN CMD_3,
										 OUT GetData,
										 OUT &GetDataLen);	
	if(lngStatus!=1)
	{
		MessageBox("SLE4442 CARD Command fail!");
	}
	m_strResponse = Bytes2CString(GetData,(UINT)GetDataLen);
}
