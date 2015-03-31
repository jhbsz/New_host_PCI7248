// AT88SCDlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "AT88SCDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CString AT88SCBYTE2CString(BYTE *ReadData, UINT Len)
{
	CString OutCString = "";
	CString Temp_BYTE_CString = "";
	UINT i;
	for(i=0;i<Len;i++)
	{
		Temp_BYTE_CString.Format("%d",ReadData[i]);
		OutCString += Temp_BYTE_CString;
		OutCString += " ";
	}
	return OutCString;
}
//-----------------------------------------
//"1 2 3 4 57" -> [1],[2],[3],[4],[57]
//Return Value: 1, success; 0, data illegal
//-----------------------------------------
BOOL CString2BytesArray(IN  CString OriginalString, 
						OUT BYTE*   DataArray,
						OUT int*    DataLen
						)
{
	if(OriginalString == "")
	{
		return 0;
	}	
	CByteArray hexdata;
	*DataLen = 0;

	CString OriginalStringBackup = OriginalString;

	OriginalStringBackup.TrimLeft(" ");
	OriginalStringBackup.TrimRight(" ");
	OriginalStringBackup.Replace("  ", " ");
	OriginalStringBackup.Replace("   ", " ");
	OriginalStringBackup.Replace("     ", " ");

	char* SingalStringCharacter = (LPSTR)(LPCTSTR) OriginalStringBackup;

	int DataInArray = 0;

	UINT i = 0;
	while(1)
	{
		if(SingalStringCharacter[i] == 0x00)
		{
			if(DataInArray<=255)
			{
				DataArray[(*DataLen)] = DataInArray;
				DataInArray = 0;
			}
			else
			{
				return 0;
			}
			(*DataLen) ++;
			break;
		}
		else if(SingalStringCharacter[i] == ' ')
		{
			if(DataInArray<=255)
			{
				DataArray[(*DataLen)] = DataInArray;
				DataInArray = 0;
			}
			else
			{
				return 0;
			}
			(*DataLen) ++;
			i++;
		}
		else
		{
			if((SingalStringCharacter[i] <= '9')||(SingalStringCharacter[i] >= '0'))
			{
				DataInArray = ((DataInArray * 10) + (SingalStringCharacter[i] - '0'));
			}
			else
			{
				return 0;
			}
			i++;
		}
	
	}
	return 1;
}


/////////////////////////////////////////////////////////////////////////////
// CAT88SCDlg dialog


CAT88SCDlg::CAT88SCDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAT88SCDlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CAT88SCDlg)
	m_strWriteUserZoneAddr = 0;
	m_strWriteUserZoneData = _T("1 2 3 4 5");
	m_strWriteUserZoneReplyLen = 0;
	m_strReadUserZoneAddr = 0;
	m_strReadUserZoneData = _T("");
	m_strReadUserZoneReplyLen = 1;
	m_strWriteConfigurationZoneAddr = 56;
	m_strWriteConfigurationZoneData = _T("1 2 3 4 5");
	m_strWriteConfigurationZoneReplyLen = 0;
	m_strReadConfigurationZoneAddr = 56;
	m_strReadConfigurationZoneData = _T("");
	m_strReadConfigurationZoneReplyLen = 5;
	m_strSetUserZoneAddressAddr = 0;
	m_strSetUserZoneAddressData = _T("");
	m_strSetUserZoneAddressReplyLen = 0;
	m_strCardResponse = _T("");
	m_strPassword1 = 255;
	m_strPassword2 = 255;
	m_strPassword3 = 255;
	//}}AFX_DATA_INIT
}


void CAT88SCDlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CAT88SCDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAT88SCDlg)
	DDX_Control(pDX, IDC_COMBO_ZONE, m_ctrlCOMBO_ZONE);
	DDX_Control(pDX, IDC_COMBO_R_W, m_ctrlCOMBO_R_W);
	DDX_Text(pDX, IDC_WriteUserZoneAddr, m_strWriteUserZoneAddr);
	DDX_Text(pDX, IDC_WriteUserZoneData, m_strWriteUserZoneData);
	DDX_Text(pDX, IDC_WriteUserZoneReplyLen, m_strWriteUserZoneReplyLen);
	DDX_Text(pDX, IDC_ReadUserZoneAddr, m_strReadUserZoneAddr);
	DDX_Text(pDX, IDC_ReadUserZoneData, m_strReadUserZoneData);
	DDX_Text(pDX, IDC_ReadUserZoneReplyLen, m_strReadUserZoneReplyLen);
	DDX_Text(pDX, IDC_WriteConfigurationZoneAddr, m_strWriteConfigurationZoneAddr);
	DDX_Text(pDX, IDC_WriteConfigurationZoneData, m_strWriteConfigurationZoneData);
	DDX_Text(pDX, IDC_WriteConfigurationZoneReplyLen, m_strWriteConfigurationZoneReplyLen);
	DDX_Text(pDX, IDC_ReadConfigurationZoneAddr, m_strReadConfigurationZoneAddr);
	DDX_Text(pDX, IDC_ReadConfigurationZoneData, m_strReadConfigurationZoneData);
	DDX_Text(pDX, IDC_ReadConfigurationZoneReplyLen, m_strReadConfigurationZoneReplyLen);
	DDX_Text(pDX, IDC_SetUserZoneAddressAddr, m_strSetUserZoneAddressAddr);
	DDX_Text(pDX, IDC_SetUserZoneAddressData, m_strSetUserZoneAddressData);
	DDX_Text(pDX, IDC_SetUserZoneAddressReplyLen, m_strSetUserZoneAddressReplyLen);
	DDX_Text(pDX, IDC_CardResponse, m_strCardResponse);
	DDX_Text(pDX, IDC_Password1, m_strPassword1);
	DDX_Text(pDX, IDC_Password2, m_strPassword2);
	DDX_Text(pDX, IDC_Password3, m_strPassword3);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAT88SCDlg, CDialog)
	//{{AFX_MSG_MAP(CAT88SCDlg)
	ON_BN_CLICKED(IDC_BUTTONWriteUserZone, OnBUTTONWriteUserZone)
	ON_BN_CLICKED(IDC_BUTTONReadUserZone, OnBUTTONReadUserZone)
	ON_BN_CLICKED(IDC_BUTTONWriteConfigurationZone, OnBUTTONWriteConfigurationZone)
	ON_BN_CLICKED(IDC_BUTTONReadConfigurationZone, OnBUTTONReadConfigurationZone)
	ON_BN_CLICKED(IDC_BUTTONSetUserZoneAddress, OnBUTTONSetUserZoneAddress)
	ON_BN_CLICKED(IDC_BUTTONVerifyPassword, OnBUTTONVerifyPassword)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CAT88SCDlg, CDialog)
	//{{AFX_DISPATCH_MAP(CAT88SCDlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IAT88SCDlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {13173CDC-7127-4705-BB15-151360609021}
static const IID IID_IAT88SCDlg =
{ 0x13173cdc, 0x7127, 0x4705, { 0xbb, 0x15, 0x15, 0x13, 0x60, 0x60, 0x90, 0x21 } };

BEGIN_INTERFACE_MAP(CAT88SCDlg, CDialog)
	INTERFACE_PART(CAT88SCDlg, IID_IAT88SCDlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAT88SCDlg message handlers

BOOL CAT88SCDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	m_ctrlCOMBO_R_W.AddString("Write");
	m_ctrlCOMBO_R_W.AddString("Read");
	m_ctrlCOMBO_R_W.SetCurSel(0);

	m_ctrlCOMBO_ZONE.AddString("0");
	m_ctrlCOMBO_ZONE.AddString("1");
	m_ctrlCOMBO_ZONE.AddString("2");
	m_ctrlCOMBO_ZONE.AddString("3");
	m_ctrlCOMBO_ZONE.AddString("4");
	m_ctrlCOMBO_ZONE.AddString("5");
	m_ctrlCOMBO_ZONE.AddString("6");
	m_ctrlCOMBO_ZONE.AddString("7");
	m_ctrlCOMBO_ZONE.SetCurSel(0);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CAT88SCDlg::OnBUTTONWriteUserZone() 
{
	// TODO: Add your control notification handler code here
/*
LONG APIENTRY AT88SC1608Cmd_WriteUserZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bWriteLen,
		IN	LPVOID	pWriteBuffer
		)	
*/
	int Result;

	UpdateData(TRUE);
	int WriteLen;
	BYTE WriteBuffer[300];
	if(!CString2BytesArray(m_strWriteUserZoneData, WriteBuffer, &WriteLen))
	{
		MessageBox("Data Format is illegal");
		return;
	}
	Result = AT88SC1608Cmd_WriteUserZone(
					IN	&CSerial9525,
					IN	SlotNum,
					IN	m_strWriteUserZoneAddr,
					IN	WriteLen,
					IN	WriteBuffer
					);
	if(Result == 0)
	{
		m_strCardResponse = "Write User Zone success";
	}
	else
	{
		m_strCardResponse = "Write User Zone fail";
	}
	UpdateData(FALSE);
}

void CAT88SCDlg::OnBUTTONReadUserZone() 
{
	// TODO: Add your control notification handler code here
/*LONG APIENTRY AT88SC1608Cmd_ReadUserZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadBuffer,
		OUT	UCHAR	*pReturnLen
		)*/	
	int Result;

	BYTE ReadData[300];
	BYTE ReadDataLen;

	UpdateData(TRUE);
	Result = AT88SC1608Cmd_ReadUserZone(
									IN	&CSerial9525,
									IN	SlotNum,
									IN	m_strReadUserZoneAddr,
									IN	m_strReadUserZoneReplyLen,
									OUT	ReadData,
									&ReadDataLen
									);
	if(Result == 0)
	{
		m_strCardResponse = AT88SCBYTE2CString(ReadData,(UINT)ReadDataLen);
	}
	else
	{
		m_strCardResponse = "Read User Zone fail";
	}
	UpdateData(FALSE);
}

void CAT88SCDlg::OnBUTTONWriteConfigurationZone() 
{
	// TODO: Add your control notification handler code here
/*
LONG APIENTRY AT88SC1608Cmd_WriteConfigurationZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bWriteLen,
		IN	LPVOID	pWriteBuffer
		);
*/	
	int Result;

	UpdateData(TRUE);
	int WriteLen;
	BYTE WriteBuffer[300];
	if(!CString2BytesArray(m_strWriteConfigurationZoneData, WriteBuffer, &WriteLen))
	{
		MessageBox("Data Format is illegal");
		return;
	}
	Result = AT88SC1608Cmd_WriteConfigurationZone(
					IN	&CSerial9525,
					IN	SlotNum,
					IN	m_strWriteConfigurationZoneAddr,
					IN	WriteLen,
					IN	WriteBuffer
					);
	if(Result == 0)
	{
		m_strCardResponse = "Write Configuration Zone success";
	}
	else
	{
		m_strCardResponse = "Write Configuration Zone fail";
	}
	UpdateData(FALSE);
}

void CAT88SCDlg::OnBUTTONReadConfigurationZone() 
{
	// TODO: Add your control notification handler code here
/*
LONG APIENTRY AT88SC1608Cmd_ReadConfigurationZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadBuffer,
		OUT	UCHAR	*pReturnLen
		);	
*/
	int Result;

	BYTE ReadData[300];
	BYTE ReadDataLen;

	UpdateData(TRUE);
	Result = AT88SC1608Cmd_ReadConfigurationZone(
									IN	&CSerial9525,
									IN	SlotNum,
									IN	m_strReadConfigurationZoneAddr,
									IN	m_strReadConfigurationZoneReplyLen,
									OUT	ReadData,
									&ReadDataLen
									);
	if(Result == 0)
	{
		m_strCardResponse = AT88SCBYTE2CString(ReadData,(UINT)ReadDataLen);
	}
	else
	{
		m_strCardResponse = "Read Configuration Zone fail";
	}
	UpdateData(FALSE);
}

void CAT88SCDlg::OnBUTTONSetUserZoneAddress() 
{
	// TODO: Add your control notification handler code here
/*
LONG APIENTRY AT88SC1608Cmd_SetUserZoneAddress(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress
		);	
*/
	int Result;

	UpdateData(TRUE);

	Result = AT88SC1608Cmd_SetUserZoneAddress(
					IN	&CSerial9525,
					IN	SlotNum,
					IN	m_strSetUserZoneAddressAddr
					);
	if(Result == 0)
	{
		m_strCardResponse = "Set User Zone Address success";
	}
	else
	{
		m_strCardResponse = "Set User Zone Address fail";
	}
	UpdateData(FALSE);	
}

void CAT88SCDlg::OnBUTTONVerifyPassword() 
{
	// TODO: Add your control notification handler code here
/*
LONG APIENTRY AT88SC1608Cmd_VerifyPassword(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bZoneNo,
		IN	BOOL 	bIsReadAccess,
		IN	UCHAR	bPW1,
		IN	UCHAR	bPW2,
		IN	UCHAR	bPW3
		);	
*/
	int Result;

	UpdateData(TRUE);

	Result = AT88SC1608Cmd_VerifyPassword(
					IN	&CSerial9525,
					IN	SlotNum,
					IN	m_ctrlCOMBO_ZONE.GetCurSel(),
					IN	m_ctrlCOMBO_R_W.GetCurSel(),
					IN	m_strPassword1,
					IN	m_strPassword2,
					IN	m_strPassword3
					);
	if(Result == 0)
	{
		m_strCardResponse = "Verify Password success";
	}
	else
	{
		m_strCardResponse = "Verify Password fail";
	}
	UpdateData(FALSE);	

}
