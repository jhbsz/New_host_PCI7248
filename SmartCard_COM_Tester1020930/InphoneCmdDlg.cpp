// InphoneCmdDlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "InphoneCmdDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CInphoneCmdDlg dialog


CInphoneCmdDlg::CInphoneCmdDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CInphoneCmdDlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CInphoneCmdDlg)
	m_strINPHONE_CARD_Read = _T("00 00");
	m_strINPHONE_CARD_PROG = _T("00 00");
	m_strINPHONE_CARD_MOVE_ADDRESS = _T("00 00");
	m_strINPHONE_CARD_AUTHENTICATION_KEY1 = _T("00 00");
	m_strINPHONE_CARD_AUTHENTICATION_KEY2 = _T("00 00");
	m_strCardResponse = _T("");
	m_strKey1_SendData = _T("");
	m_strKey2_SendData = _T("");
	//}}AFX_DATA_INIT
}


void CInphoneCmdDlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CInphoneCmdDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CInphoneCmdDlg)
	DDX_Text(pDX, IDC_EDIT_INPHONE_CARD_Read, m_strINPHONE_CARD_Read);
	DDX_Text(pDX, IDC_EDIT_INPHONE_CARD_PROG, m_strINPHONE_CARD_PROG);
	DDX_Text(pDX, IDC_EDIT_INPHONE_CARD_MOVE_ADDRESS, m_strINPHONE_CARD_MOVE_ADDRESS);
	DDX_Text(pDX, IDC_EDIT_INPHONE_CARD_AUTHENTICATION_KEY1, m_strINPHONE_CARD_AUTHENTICATION_KEY1);
	DDX_Text(pDX, IDC_EDIT_INPHONE_CARD_AUTHENTICATION_KEY2, m_strINPHONE_CARD_AUTHENTICATION_KEY2);
	DDX_Text(pDX, IDC_EDIT_CardResponse, m_strCardResponse);
	DDX_Text(pDX, IDC_EDIT_Key1_SendData, m_strKey1_SendData);
	DDX_Text(pDX, IDC_EDIT_Key2_SendData, m_strKey2_SendData);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CInphoneCmdDlg, CDialog)
	//{{AFX_MSG_MAP(CInphoneCmdDlg)
	ON_BN_CLICKED(IDC_BUTTON_INPHONE_CARD_RESET, OnButtonInphoneCardReset)
	ON_BN_CLICKED(IDC_BUTTON_INPHONE_CARD_Read, OnBUTTONINPHONECARDRead)
	ON_BN_CLICKED(IDC_BUTTON_INPHONE_CARD_PROG, OnButtonInphoneCardProg)
	ON_BN_CLICKED(IDC_BUTTON_INPHONE_CARD_MOVE_ADDRESS, OnButtonInphoneCardMoveAddress)
	ON_BN_CLICKED(IDC_BUTTONINPHONE_CARD_AUTHENTICATION_KEY1, OnButtoninphoneCardAuthenticationKey1)
	ON_BN_CLICKED(IDC_BUTTON_INPHONE_CARD_AUTHENTICATION_KEY2, OnButtonInphoneCardAuthenticationKey2)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CInphoneCmdDlg, CDialog)
	//{{AFX_DISPATCH_MAP(CInphoneCmdDlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IInphoneCmdDlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {249A0480-866C-4A68-BBD2-22C8DC2E44D1}
static const IID IID_IInphoneCmdDlg =
{ 0x249a0480, 0x866c, 0x4a68, { 0xbb, 0xd2, 0x22, 0xc8, 0xdc, 0x2e, 0x44, 0xd1 } };

BEGIN_INTERFACE_MAP(CInphoneCmdDlg, CDialog)
	INTERFACE_PART(CInphoneCmdDlg, IID_IInphoneCmdDlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CInphoneCmdDlg message handlers

void CInphoneCmdDlg::OnButtonInphoneCardReset() 
{
	// TODO: Add your control notification handler code here
	LONG	lngStatus;
	lngStatus = CMD_INPHONE_CARD_RESET(	IN	&CSerial9525,
										IN	SlotNum);
	if(lngStatus!=1)
	{
		MessageBox("INPHONE CARD RESET fail!");
	}	
}
/*
LONG CMD_INPHONE_CARD_Read(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  IN	ULONG	lngReadLen,
						  OUT	LPVOID	pReadData,
						  OUT	ULONG	*plngReturnLen);


	UINT Len;

	BYTE ReadingLengthArray[2];
	Len = CString2Bytes(ReadingLengthArray, m_strReadDataLength);
	UINT ReadingLength; 
	ReadingLength = (ReadingLengthArray[0]*256 + ReadingLengthArray[1]);

	m_strINPHONE_CARD_Read = _T("00 00");
	m_strINPHONE_CARD_PROG = _T("00 00");
	m_strINPHONE_CARD_MOVE_ADDRESS = _T("00 00");
	m_strINPHONE_CARD_AUTHENTICATION_KEY1 = _T("00 00");
	m_strINPHONE_CARD_AUTHENTICATION_KEY2 = _T("00 00");
	m_strCardResponse = _T("");
	m_strKey1_SendData = _T("");
	m_strKey2_SendData = _T("");
*/

void CInphoneCmdDlg::OnBUTTONINPHONECARDRead()
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	LONG	lngStatus;
	UINT Len;

	BYTE ReadingLengthArray[2];
	Len = CString2Bytes(ReadingLengthArray, m_strINPHONE_CARD_Read);
	UINT ReadingLength; 
	ReadingLength = (ReadingLengthArray[0]*256 + ReadingLengthArray[1]);

	BYTE  GetData[300];
	ULONG GetDataLen = 0;

	lngStatus = CMD_INPHONE_CARD_Read(	IN	&CSerial9525,
										IN	SlotNum,
										IN  ReadingLength,
										OUT GetData,
										OUT &GetDataLen);	
	if(lngStatus!=1)
	{
		MessageBox("Inphone CARD read fail!");
	}
	m_strCardResponse = Bytes2CString(GetData,(UINT)GetDataLen);

	UpdateData(FALSE);
}

void CInphoneCmdDlg::OnButtonInphoneCardProg() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	LONG	lngStatus;
	UINT Len;

	BYTE ReadingLengthArray[2];
	Len = CString2Bytes(ReadingLengthArray, m_strINPHONE_CARD_PROG);
	UINT ReadingLength; 
	ReadingLength = (ReadingLengthArray[0]*256 + ReadingLengthArray[1]);

	BYTE  GetData[300];
	ULONG GetDataLen = 0;

	lngStatus = CMD_INPHONE_CARD_PROG(	IN	&CSerial9525,
										IN	SlotNum,
										IN  ReadingLength,
										OUT GetData,
										OUT &GetDataLen);	
	if(lngStatus!=1)
	{
		MessageBox("Inphone CARD program fail!");
	}
	m_strCardResponse = Bytes2CString(GetData,(UINT)GetDataLen);

	UpdateData(FALSE);	
}

void CInphoneCmdDlg::OnButtonInphoneCardMoveAddress() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	LONG	lngStatus;
	UINT Len;

	BYTE ReadingLengthArray[2];
	Len = CString2Bytes(ReadingLengthArray, m_strINPHONE_CARD_MOVE_ADDRESS);
	UINT ClockNums; 
	ClockNums = (ReadingLengthArray[0]*256 + ReadingLengthArray[1]);

	lngStatus = CMD_INPHONE_CARD_MOVE_ADDRESS(	IN	&CSerial9525,
												IN	SlotNum,
												IN  ClockNums);	
	if(lngStatus!=1)
	{
		MessageBox("Inphone CARD Move Address fail!");
	}
	UpdateData(FALSE);	
}



/*
	UpdateData(TRUE);
	LONG	lngStatus;
	UINT Len;

	BYTE ReadingLengthArray[2];
	Len = CString2Bytes(ReadingLengthArray, m_strReadDataLength);
	UINT ReadingLength; 
	ReadingLength = (ReadingLengthArray[0]*256 + ReadingLengthArray[1]);

	BYTE SendData[300];
	UINT SendDataLen = 0;
	SendDataLen = CString2Bytes(SendData, m_strSendData);

	BYTE GetData[300];
	ULONG ReturnLen = 0;

	lngStatus = CMD_SMC_COMMAND(IN	&CSerial9525,
								IN	SlotNum,
								IN	SendDataLen,
								IN	SendData,
								IN	ReadingLength,
						        OUT	GetData,
								OUT	&ReturnLen);
	if(lngStatus!=1)
	{
		MessageBox("SMC CARD Command fail!");
	}	

	m_strGetData = Bytes2CString(GetData,(UINT)ReturnLen);
*/
void CInphoneCmdDlg::OnButtoninphoneCardAuthenticationKey1() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	LONG	lngStatus;
	UINT Len;

	BYTE AUTHENTICATION_KEY1_Len_Array[2];
	Len = CString2Bytes(AUTHENTICATION_KEY1_Len_Array, m_strINPHONE_CARD_AUTHENTICATION_KEY1);
	UINT AUTHENTICATION_KEY1_Len; 
	AUTHENTICATION_KEY1_Len = (AUTHENTICATION_KEY1_Len_Array[0]*256 + AUTHENTICATION_KEY1_Len_Array[1]);	

	BYTE SendData[300];
	UINT SendDataLen = 0;
	SendDataLen = CString2Bytes(SendData, m_strKey1_SendData);

	BYTE GetData[300];
	ULONG ReturnLen = 0;

	lngStatus = CMD_INPHONE_CARD_AUTHENTICATION_KEY1(IN	&CSerial9525,
													IN	SlotNum,
													IN  SendDataLen,
													IN	SendData,
													IN	AUTHENTICATION_KEY1_Len,
													OUT	GetData,
													OUT	&ReturnLen);
	if(lngStatus!=1)
	{
		MessageBox("INPHONE CARD AUTHENTICATION KEY1 fail!");
	}	

	m_strCardResponse = Bytes2CString(GetData,(UINT)ReturnLen);
	UpdateData(FALSE);
}

void CInphoneCmdDlg::OnButtonInphoneCardAuthenticationKey2() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	LONG	lngStatus;
	UINT Len;

	BYTE AUTHENTICATION_KEY2_Len_Array[2];
	Len = CString2Bytes(AUTHENTICATION_KEY2_Len_Array, m_strINPHONE_CARD_AUTHENTICATION_KEY2);
	UINT AUTHENTICATION_KEY2_Len; 
	AUTHENTICATION_KEY2_Len = (AUTHENTICATION_KEY2_Len_Array[0]*256 + AUTHENTICATION_KEY2_Len_Array[1]);	

	BYTE SendData[300];
	UINT SendDataLen = 0;
	SendDataLen = CString2Bytes(SendData, m_strKey2_SendData);

	BYTE GetData[300];
	ULONG ReturnLen = 0;

	lngStatus = CMD_INPHONE_CARD_AUTHENTICATION_KEY2(IN	&CSerial9525,
													IN	SlotNum,
													IN  SendDataLen,
													IN	SendData,
													IN	AUTHENTICATION_KEY2_Len,
													OUT	GetData,
													OUT	&ReturnLen);
	if(lngStatus!=1)
	{
		MessageBox("INPHONE CARD AUTHENTICATION KEY2 fail!");
	}	

	m_strCardResponse = Bytes2CString(GetData,(UINT)ReturnLen);
	UpdateData(FALSE);	
}
