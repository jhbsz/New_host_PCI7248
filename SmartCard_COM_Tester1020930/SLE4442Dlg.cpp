// SLE4442Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "SLE4442Dlg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSLE4442Dlg dialog


CSLE4442Dlg::CSLE4442Dlg(CWnd* pParent /*=NULL*/)
	: CDialog(CSLE4442Dlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CSLE4442Dlg)
	m_strReadMainMemoryAddr = 0;
	m_strReadMainMemoryData = 0;
	m_strReadMainMemoryReplyLen = 255;
	m_strUpdateMainMemoryAddr = 0;
	m_strUpdateMainMemoryData = 0;
	m_strUpdateMainMemoryReplyLen = 0;
	m_strReadProtectionMemoryAddr = 0;
	m_strReadProtectionMemoryData = 0;
	m_strReadProtectionMemoryReplyLen = 4;
	m_strWriteProtectionMemoryAddr = 0;
	m_strWriteProtectionMemoryData = 0;
	m_strWriteProtectionMemoryReplyLen = 0;
	m_strReadSecurityMemoryAddr = 0;
	m_strReadSecurityMemoryData = 0;
	m_strReadSecurityMemoryReplyLen = 4;
	m_strUpdateSecurityMemoryAddr = 1;
	m_strUpdateSecurityMemoryData = 255;
	m_strUpdateSecurityMemoryReplyLen = 0;
	m_strCompareVerificationDataAddr = 0;
	m_strCompareVerificationDataData = 0;
	m_strCompareVerificationDataReplyLen = 0;
	m_strCardResponse = _T("");
	m_strReferenceData1 = 255;
	m_strReferenceData2 = 255;
	m_strReferenceData3 = 255;
	//}}AFX_DATA_INIT
}


void CSLE4442Dlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CSLE4442Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CSLE4442Dlg)
	DDX_Text(pDX, IDC_ReadMainMemoryAddr, m_strReadMainMemoryAddr);
	DDX_Text(pDX, IDC_ReadMainMemoryData, m_strReadMainMemoryData);
	DDX_Text(pDX, IDC_ReadMainMemoryReplyLen, m_strReadMainMemoryReplyLen);
	DDX_Text(pDX, IDC_UpdateMainMemoryAddr, m_strUpdateMainMemoryAddr);
	DDX_Text(pDX, IDC_UpdateMainMemoryData, m_strUpdateMainMemoryData);
	DDX_Text(pDX, IDC_UpdateMainMemoryReplyLen, m_strUpdateMainMemoryReplyLen);
	DDX_Text(pDX, IDC_ReadProtectionMemoryAddr, m_strReadProtectionMemoryAddr);
	DDX_Text(pDX, IDC_ReadProtectionMemoryData, m_strReadProtectionMemoryData);
	DDX_Text(pDX, IDC_ReadProtectionMemoryReplyLen, m_strReadProtectionMemoryReplyLen);
	DDX_Text(pDX, IDC_WriteProtectionMemoryAddr, m_strWriteProtectionMemoryAddr);
	DDX_Text(pDX, IDC_WriteProtectionMemoryData, m_strWriteProtectionMemoryData);
	DDX_Text(pDX, IDC_WriteProtectionMemoryReplyLen, m_strWriteProtectionMemoryReplyLen);
	DDX_Text(pDX, IDC_ReadSecurityMemoryAddr, m_strReadSecurityMemoryAddr);
	DDX_Text(pDX, IDC_ReadSecurityMemoryData, m_strReadSecurityMemoryData);
	DDX_Text(pDX, IDC_ReadSecurityMemoryReplyLen, m_strReadSecurityMemoryReplyLen);
	DDX_Text(pDX, IDC_UpdateSecurityMemoryAddr, m_strUpdateSecurityMemoryAddr);
	DDX_Text(pDX, IDC_UpdateSecurityMemoryData, m_strUpdateSecurityMemoryData);
	DDX_Text(pDX, IDC_UpdateSecurityMemoryReplyLen, m_strUpdateSecurityMemoryReplyLen);
	DDX_Text(pDX, IDC_CompareVerificationDataAddr, m_strCompareVerificationDataAddr);
	DDX_Text(pDX, IDC_CompareVerificationDataData, m_strCompareVerificationDataData);
	DDX_Text(pDX, IDC_CompareVerificationDataReplyLen, m_strCompareVerificationDataReplyLen);
	DDX_Text(pDX, IDC_CardResponse, m_strCardResponse);
	DDX_Text(pDX, IDC_PinNumber1, m_strReferenceData1);
	DDX_Text(pDX, IDC_PinNumber2, m_strReferenceData2);
	DDX_Text(pDX, IDC_PinNumber3, m_strReferenceData3);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CSLE4442Dlg, CDialog)
	//{{AFX_MSG_MAP(CSLE4442Dlg)
	ON_BN_CLICKED(IDC_BUTTONReadMainMemory, OnBUTTONReadMainMemory)
	ON_BN_CLICKED(IDC_BUTTONReadProtectionMemory, OnBUTTONReadProtectionMemory)
	ON_BN_CLICKED(IDC_BUTTONReadSecurityMemory, OnBUTTONReadSecurityMemory)
	ON_BN_CLICKED(IDC_Verify, OnVerify)
	ON_BN_CLICKED(IDC_BUTTONUpdateMainMemory, OnBUTTONUpdateMainMemory)
	ON_BN_CLICKED(IDC_BUTTONWriteProtectionMemory, OnBUTTONWriteProtectionMemory)
	ON_BN_CLICKED(IDC_BUTTONUpdateSecurityMemory, OnBUTTONUpdateSecurityMemory)
	ON_BN_CLICKED(IDC_BUTTONCompareVerificationData, OnBUTTONCompareVerificationData)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CSLE4442Dlg, CDialog)
	//{{AFX_DISPATCH_MAP(CSLE4442Dlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ISLE4442Dlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {4C5AE205-66B2-4629-B4EF-3DD2896192D5}
static const IID IID_ISLE4442Dlg =
{ 0x4c5ae205, 0x66b2, 0x4629, { 0xb4, 0xef, 0x3d, 0xd2, 0x89, 0x61, 0x92, 0xd5 } };

BEGIN_INTERFACE_MAP(CSLE4442Dlg, CDialog)
	INTERFACE_PART(CSLE4442Dlg, IID_ISLE4442Dlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSLE4442Dlg message handlers
CString ReadDataArray4442Format(BYTE *ReadData, UINT Len)
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

CString ReadPBDataArray4442Format(BYTE *ReadData, BYTE *PBData, UINT Len)
{
	CString OutCString = "";
	CString Temp_BYTE_CString = "";
	UINT i;
	for(i=0;i<Len;i++)
	{
		Temp_BYTE_CString.Format("%d",ReadData[i]);
		OutCString += Temp_BYTE_CString;
		OutCString += "|";
		Temp_BYTE_CString.Format("%d",PBData[i]);
		OutCString += Temp_BYTE_CString;
		OutCString += " ";
	}
	return OutCString;
}

void CSLE4442Dlg::OnBUTTONReadMainMemory() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BYTE ResponseData[1030];	
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4442Cmd_ReadMainMemory(
			IN	&CSerial9525,
			IN	SlotNum,
			IN  m_strReadMainMemoryAddr,
			IN	m_strReadMainMemoryReplyLen,
			OUT	&ResponseData,
			OUT  &Len
			);
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Read success\r\n";		
			m_strCardResponse += ReadDataArray4442Format(ResponseData, (UINT)Len);
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Read fail";			
		}
		break;
	}
	UpdateData(FALSE);	
}

void CSLE4442Dlg::OnBUTTONReadProtectionMemory() 
{
//	m_strReadProtectionMemoryAddr = 0;
//	m_strReadProtectionMemoryData = 0;
//	m_strReadProtectionMemoryReplyLen = 4;
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BYTE ResponseData[1030];	
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4442Cmd_ReadProtectionMemory(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strReadProtectionMemoryReplyLen,
			OUT	&ResponseData,
			OUT  &Len
			);
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Read Protection Memory success\r\n";		
			m_strCardResponse += ReadDataArray4442Format(ResponseData, (UINT)Len);
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Read fail";			
		}
		break;
	}
	UpdateData(FALSE);	
}

void CSLE4442Dlg::OnBUTTONReadSecurityMemory() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BYTE ResponseData[1030];	
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4442Cmd_ReadSecurityMemory(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strReadSecurityMemoryReplyLen,
			OUT	&ResponseData,
			OUT  &Len
			); 
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Read Security Memory success\r\n";		
			m_strCardResponse += ReadDataArray4442Format(ResponseData, (UINT)Len);
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Read fail";			
		}
		break;
	}
	UpdateData(FALSE);		
}

void CSLE4442Dlg::OnVerify() 
{
	// TODO: Add your control notification handler code here
	//-------------------------------------
	//Read Error Count
	//-------------------------------------
	UpdateData(TRUE);
	BYTE ResponseData[1030];	
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4442Cmd_ReadSecurityMemory(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	4,
						OUT	&ResponseData,
						OUT  &Len
						);
	if(Result == 1)
	{
		m_strCardResponse = "Read fail";
		UpdateData(FALSE);	
		return;	
	}
	if(ResponseData[0] == 0)
	{
		m_strCardResponse = "Error Count is 0";
		UpdateData(FALSE);	
		return;
	}
	//-------------------------------------
	//Write one bit of the Error Count
	//-------------------------------------
	BYTE NewErrorCount = 0;
	if((ResponseData[0] & 0x01) != 0)
	{
		NewErrorCount = ResponseData[0] & 0xfe;
	}
	else if((ResponseData[0] & 0x02) != 0)
	{
		NewErrorCount = ResponseData[0]&0xfd;
	}
	else if((ResponseData[0] & 0x04) != 0)
	{
		NewErrorCount = ResponseData[0] & 0xfb;
	}
	Result = SLE4442Cmd_UpdateSecurityMemory(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	0,
						IN  NewErrorCount
						);
	if(Result == 1)
	{
		m_strCardResponse = "Write Error Count fail";
		UpdateData(FALSE);	
		return;	
	}
	//-------------------------------------
	//Compare verification data 1
	//-------------------------------------
/*LONG APIENTRY SLE4442Cmd_CompareVerificationData(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bData
		);*/
	Result = SLE4442Cmd_CompareVerificationData(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	1,
						IN  m_strReferenceData1
						);
	if(Result == 1)
	{
		m_strCardResponse = "Compare verification data fail";
		UpdateData(FALSE);	
		return;	
	}
	//-------------------------------------
	//Compare verification data 2
	//-------------------------------------
	Result = SLE4442Cmd_CompareVerificationData(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	2,
						IN  m_strReferenceData2
						);
	if(Result == 1)
	{
		m_strCardResponse = "Compare verification data fail";
		UpdateData(FALSE);	
		return;	
	}
	//-------------------------------------
	//Compare verification data 3
	//-------------------------------------
	Result = SLE4442Cmd_CompareVerificationData(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	3,
						IN  m_strReferenceData3
						);
	if(Result == 1)
	{
		m_strCardResponse = "Compare verification data fail";
		UpdateData(FALSE);	
		return;	
	}
	//-------------------------------------
	//write Error Count to 3 times (0x07)
	//-------------------------------------
	Result = SLE4442Cmd_UpdateSecurityMemory(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	0,
						IN  0x07
						);
	if(Result == 1)
	{
		m_strCardResponse = "Write Error Count fail";
		UpdateData(FALSE);	
		return;	
	}
	m_strCardResponse = "Verify Success";
	UpdateData(FALSE);
}

void CSLE4442Dlg::OnBUTTONUpdateMainMemory() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4442Cmd_UpdateMainMemory(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	m_strUpdateMainMemoryAddr,
						IN	m_strUpdateMainMemoryData
						);	
	if(Result == 1)
	{
		m_strCardResponse = "Update Main Memory fail";		
	}
	else
	{
		m_strCardResponse = "Update Main Memory success";	
	}
	UpdateData(FALSE);	
}

void CSLE4442Dlg::OnBUTTONWriteProtectionMemory() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4442Cmd_WriteProtectionMemory(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	m_strWriteProtectionMemoryAddr,
						IN	m_strWriteProtectionMemoryData
						);	
	if(Result == 1)
	{
		m_strCardResponse = "Write Protection Memory fail";		
	}
	else
	{
		m_strCardResponse = "Write Protection Memory success";	
	}
	UpdateData(FALSE);	
}

void CSLE4442Dlg::OnBUTTONUpdateSecurityMemory() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4442Cmd_UpdateSecurityMemory(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	m_strUpdateSecurityMemoryAddr,
						IN	m_strUpdateSecurityMemoryData
						);	
	if(Result == 1)
	{
		m_strCardResponse = "Update Security Memory fail";		
	}
	else
	{
		m_strCardResponse = "Update Security Memory success";	
	}
	UpdateData(FALSE);	
}

void CSLE4442Dlg::OnBUTTONCompareVerificationData() 
{
	// TODO: Add your control notification handler code here
/*LONG APIENTRY SLE4442Cmd_CompareVerificationData(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bData
		)*/
	UpdateData(TRUE);
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4442Cmd_CompareVerificationData(
						IN	&CSerial9525,
						IN	SlotNum,
						IN	m_strCompareVerificationDataAddr,
						IN	m_strCompareVerificationDataData
						);	
	if(Result == 1)
	{
		m_strCardResponse = "Compare Verification Data fail";		
	}
	else
	{
		m_strCardResponse = "Compare Verification Data success";	
	}
	UpdateData(FALSE);	
}
