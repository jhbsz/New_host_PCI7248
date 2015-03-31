// AT88SCcmdDlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "AT88SCcmdDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAT88SCcmdDlg dialog


CAT88SCcmdDlg::CAT88SCcmdDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAT88SCcmdDlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CAT88SCcmdDlg)
	m_strReadDataLength = _T("00 00");
	m_strSendData = _T("");
	m_strGetData = _T("");
	//}}AFX_DATA_INIT
}


void CAT88SCcmdDlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CAT88SCcmdDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAT88SCcmdDlg)
	DDX_Text(pDX, IDC_EDIT_ReadDataLength, m_strReadDataLength);
	DDX_Text(pDX, IDC_EDIT_SendData, m_strSendData);
	DDX_Text(pDX, IDC_EDIT_GetData, m_strGetData);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAT88SCcmdDlg, CDialog)
	//{{AFX_MSG_MAP(CAT88SCcmdDlg)
	ON_BN_CLICKED(IDC_BUTTON_SMC_COMMAND, OnButtonSmcCommand)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CAT88SCcmdDlg, CDialog)
	//{{AFX_DISPATCH_MAP(CAT88SCcmdDlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IAT88SCcmdDlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {DD041B54-E897-4BA1-B1E1-B958ACBF4FED}
static const IID IID_IAT88SCcmdDlg =
{ 0xdd041b54, 0xe897, 0x4ba1, { 0xb1, 0xe1, 0xb9, 0x58, 0xac, 0xbf, 0x4f, 0xed } };

BEGIN_INTERFACE_MAP(CAT88SCcmdDlg, CDialog)
	INTERFACE_PART(CAT88SCcmdDlg, IID_IAT88SCcmdDlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAT88SCcmdDlg message handlers
/*
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



	m_strReadDataLength = _T("00 00");
	m_strSendData = _T("");
	m_strGetData = _T("");




LONG CMD_SMC_COMMAND(CSerial *m_ctrlCSerial,
								BYTE bSlotNum,
								ULONG	lngWriteLen,
								UCHAR	*pWriteData,
								IN	ULONG	lngReadLen,
						        OUT	LPVOID	pReadData,
								OUT	ULONG	*plngReturnLen);
*/
void CAT88SCcmdDlg::OnButtonSmcCommand() 
{
	// TODO: Add your control notification handler code here
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

}
