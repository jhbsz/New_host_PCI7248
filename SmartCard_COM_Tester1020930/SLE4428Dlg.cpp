// SLE4428Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "SLE4428Dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSLE4428Dlg dialog


CSLE4428Dlg::CSLE4428Dlg(CWnd* pParent /*=NULL*/)
	: CDialog(CSLE4428Dlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CSLE4428Dlg)
	m_strWriteEraseWithPBAddr = 0;
	m_strWriteEraseWithPBData = 0;
	m_strWriteEraseWithPBReplyLen = 0;
	m_strWriteEraseWithoutPBAddr = 0;
	m_strWriteEraseWithoutPBData = 0;
	m_strWriteEraseWithoutPBReplyLen = 0;
	m_strWritePBdataComparisonAddr = 0;
	m_strWritePBdataComparisonData = 0;
	m_strWritePBdataComparisonReplyLen = 0;
	m_strRead9BitsAddr = 0;
	m_strRead9BitsData = 0;
	m_strRead9BitsReplyLen = 1;
	m_strRead8BitsAddr = 0;
	m_strRead8BitsData = 0;
	m_strRead8BitsReplyLen = 1;
	m_strPSC1 = 255;
	m_strPSC2 = 255;
	m_strWriteErrorCounterAddr = 1021;
	m_strWriteErrorCounterData = 254;
	m_strWriteErrorCounterReplyLen = 0;
	m_strVerify1stPSCAddr = 1022;
	m_strVerify1stPSCData = 255;
	m_strVerify1stPSCReplyLen = 0;
	m_strVerify2ndPSCAddr = 1023;
	m_strVerify2ndPSCData = 255;
	m_strVerify2ndPSCReplyLen = 0;
	m_strEraseErrorCountAddr = 1021;
	m_strEraseErrorCountData = 255;
	m_strEraseErrorCountReplyLen = 0;
	m_strCardResponse = _T("");	
	//}}AFX_DATA_INIT
}


void CSLE4428Dlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CSLE4428Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CSLE4428Dlg)
	DDX_Text(pDX, IDC_WriteEraseWithPBAddr, m_strWriteEraseWithPBAddr);
	DDX_Text(pDX, IDC_WriteEraseWithPBData, m_strWriteEraseWithPBData);
	DDX_Text(pDX, IDC_WriteEraseWithPBReplyLen, m_strWriteEraseWithPBReplyLen);
	DDX_Text(pDX, IDC_WriteEraseWithoutPBAddr, m_strWriteEraseWithoutPBAddr);
	DDX_Text(pDX, IDC_WriteEraseWithoutPBData, m_strWriteEraseWithoutPBData);
	DDX_Text(pDX, IDC_WriteEraseWithoutPBReplyLen, m_strWriteEraseWithoutPBReplyLen);
	DDX_Text(pDX, IDC_WritePBdataComparisonAddr, m_strWritePBdataComparisonAddr);
	DDX_Text(pDX, IDC_WritePBdataComparisonData, m_strWritePBdataComparisonData);
	DDX_Text(pDX, IDC_WritePBdataComparisonReplyLen, m_strWritePBdataComparisonReplyLen);
	DDX_Text(pDX, IDC_Read9BitsAddr, m_strRead9BitsAddr);
	DDX_Text(pDX, IDC_Read9BitsData, m_strRead9BitsData);
	DDX_Text(pDX, IDC_Read9BitsReplyLen, m_strRead9BitsReplyLen);
	DDX_Text(pDX, IDC_Read8BitsAddr, m_strRead8BitsAddr);
	DDX_Text(pDX, IDC_Read8BitsData, m_strRead8BitsData);
	DDX_Text(pDX, IDC_Read8BitsReplyLen, m_strRead8BitsReplyLen);
	DDX_Text(pDX, IDC_PSC1, m_strPSC1);
	DDX_Text(pDX, IDC_PSC2, m_strPSC2);
	DDX_Text(pDX, IDC_WriteErrorCounterAddr, m_strWriteErrorCounterAddr);
	DDX_Text(pDX, IDC_WriteErrorCounterReplyLen, m_strWriteErrorCounterReplyLen);
	DDX_Text(pDX, IDC_Verify1stPSCAddr, m_strVerify1stPSCAddr);
	DDX_Text(pDX, IDC_Verify1stPSCData, m_strVerify1stPSCData);
	DDX_Text(pDX, IDC_Verify1stPSCReplyLen, m_strVerify1stPSCReplyLen);
	DDX_Text(pDX, IDC_Verify2ndPSCAddr, m_strVerify2ndPSCAddr);
	DDX_Text(pDX, IDC_Verify2ndPSCData, m_strVerify2ndPSCData);
	DDX_Text(pDX, IDC_Verify2ndPSCReplyLen, m_strVerify2ndPSCReplyLen);
	DDX_Text(pDX, IDC_EraseErrorCountAddr, m_strEraseErrorCountAddr);
	DDX_Text(pDX, IDC_EraseErrorCountData, m_strEraseErrorCountData);
	DDX_Text(pDX, IDC_EraseErrorCountReplyLen, m_strEraseErrorCountReplyLen);
	DDX_Text(pDX, IDC_CardResponse, m_strCardResponse);
	DDX_Text(pDX, IDC_WriteErrorCounterData, m_strWriteErrorCounterData);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CSLE4428Dlg, CDialog)
	//{{AFX_MSG_MAP(CSLE4428Dlg)
	ON_BN_CLICKED(IDC_BUTTONWriteAndEraseWithPB, OnBUTTONWriteAndEraseWithPB)
	ON_BN_CLICKED(IDC_BUTTONRead8Bits, OnBUTTONRead8Bits)
	ON_BN_CLICKED(IDC_BUTTONRead9Bits, OnBUTTONRead9Bits)
	ON_BN_CLICKED(IDC_BUTTONWriteErrorCounter, OnBUTTONWriteErrorCounter)
	ON_BN_CLICKED(IDC_BUTTONVerify1stPSC, OnBUTTONVerify1stPSC)
	ON_BN_CLICKED(IDC_BUTTONVerify2ndPSC, OnBUTTONVerify2ndPSC)
	ON_BN_CLICKED(IDC_BUTTONEraseErrorCount, OnBUTTONEraseErrorCount)
	ON_BN_CLICKED(IDC_BUTTONWriteAndEraseWithoutPB, OnBUTTONWriteAndEraseWithoutPB)
	ON_BN_CLICKED(IDC_BUTTONWritePBAndDataComparison, OnBUTTONWritePBAndDataComparison)
	ON_BN_CLICKED(IDC_BUTTONVerifyPSCAndEraseErrorCount, OnBUTTONVerifyPSCAndEraseErrorCount)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CSLE4428Dlg, CDialog)
	//{{AFX_DISPATCH_MAP(CSLE4428Dlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ISLE4428Dlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {60F68F19-88C5-4AF6-8417-D3C6813F1577}
static const IID IID_ISLE4428Dlg =
{ 0x60f68f19, 0x88c5, 0x4af6, { 0x84, 0x17, 0xd3, 0xc6, 0x81, 0x3f, 0x15, 0x77 } };

BEGIN_INTERFACE_MAP(CSLE4428Dlg, CDialog)
	INTERFACE_PART(CSLE4428Dlg, IID_ISLE4428Dlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSLE4428Dlg message handlers

CString ReadDataArray4428Format(BYTE *ReadData, UINT Len)
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

CString ReadPBDataArray4428Format(BYTE *ReadData, BYTE *PBData, UINT Len)
{
	CString OutCString = "";
	CString Temp_BYTE_CString = "";
	UINT i;
	for(i=0;i<Len;i++)
	{
		Temp_BYTE_CString.Format("%d",ReadData[i]);
		OutCString += Temp_BYTE_CString;
		OutCString += "/";
		Temp_BYTE_CString.Format("%d",PBData[i]);
		OutCString += Temp_BYTE_CString;
		OutCString += " ";
	}
	return OutCString;
}

void CSLE4428Dlg::OnBUTTONWriteAndEraseWithPB() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
//	int	m_strWriteEraseWithPBAddr;
//	int	m_strWriteEraseWithPBData;
//	int	m_strWriteEraseWithPBReplyLen;	
	BOOL Result;
	Result = SLE4428Cmd_WriteEraseWithPB(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strWriteEraseWithPBAddr,
			IN	m_strWriteEraseWithPBData
			);

	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Write success";
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Write fail";			
		}
		break;
	}
	UpdateData(FALSE);
}

void CSLE4428Dlg::OnBUTTONRead8Bits() 
{
	// TODO: Add your control notification handler code here
/*	UpdateData(TRUE);
	BYTE ResponseData[1030];	
	ULONG Len = 0;
	BOOL Result;
	Result = SLE4428Cmd_Read8Bits(
			IN	&CSerial9525,
			IN	SlotNum,
			IN  m_strRead8BitsAddr,
			IN	m_strRead8BitsReplyLen,
			OUT	&ResponseData,
			OUT  &Len
			);
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Read success\r\n";		
			m_strCardResponse += ReadDataArrayFormat(ResponseData, (UINT)Len);
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Read fail";			
		}
		break;
	}
	UpdateData(FALSE);*/
//要求数据过50时，一次读50个，直到读完
	UpdateData(TRUE);
	if((m_strRead8BitsAddr+m_strRead8BitsReplyLen)>1024)
	{
		MessageBox("start address + length can't excess 1024!");
		return;
	}
	ULONG GetDataLen;

	BYTE ResponseData[1030];//Total
	ULONG Len = 0;//Total
	BOOL Result;

	ULONG ToReplyLen;//还需发送的数据长

	int temp_Len = 0;//每次要的数据长
	int Offset = 0;

	ToReplyLen = m_strRead8BitsReplyLen;
	while(1)
	{
		if(ToReplyLen==0)
		{
			break;
		}		
		if(ToReplyLen > 50)
		{
			temp_Len = 50;
			ToReplyLen = ToReplyLen - 50;
		}
		else
		{
			temp_Len = ToReplyLen;
			ToReplyLen = 0;
		}
		Result = SLE4428Cmd_Read8Bits(
				IN	&CSerial9525,
				IN	SlotNum,
				IN  (m_strRead8BitsAddr+Offset),
				//IN	m_strRead8BitsReplyLen,
				IN	temp_Len,
				OUT	(&ResponseData[Offset]),
				OUT  &GetDataLen
				);
		Offset += temp_Len;
		if(Result)
		{
			m_strCardResponse = "Read fail";
			break;
		}
	}
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Read success\r\n";		
			m_strCardResponse += ReadDataArray4428Format(ResponseData, (UINT)m_strRead8BitsReplyLen);
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
/*
LONG APIENTRY SLE4428Cmd_Read9Bits(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	LPVOID	pReadPB,
		OUT	ULONG	*plngReturnLen
		)
*/
void CSLE4428Dlg::OnBUTTONRead9Bits() 
{
//	int	m_strRead9BitsAddr;
//	int	m_strRead9BitsData;
//	int	m_strRead9BitsReplyLen;

	// TODO: Add your control notification handler code here
/*	UpdateData(TRUE);
	BOOL Result;
	ULONG Len = 0;	
	BYTE ResponseData[1030];		
	BYTE ReadPB[1030];
	Result = SLE4428Cmd_Read9Bits(
			IN	&CSerial9525,
			IN	SlotNum,
			IN  m_strRead9BitsAddr,
			IN	m_strRead9BitsReplyLen,
			OUT	&ResponseData,
			OUT	&ReadPB,
			OUT &Len
			);

	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Read success\r\nDATA: ";
			m_strCardResponse += ReadPBDataArrayFormat(ResponseData, ReadPB, (UINT)(Len/2));
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Read fail";			
		}
		break;
	}
	UpdateData(FALSE);*/
	UpdateData(TRUE);
	if((m_strRead9BitsAddr+m_strRead9BitsReplyLen)>1024)
	{
		MessageBox("start address + length can't excess 1024!");
		return;
	}
	ULONG GetDataLen;


	BYTE ResponseData[1030];//Total
	BYTE ReadPB[1030];
//	BYTE ResponseData[1130];//Total
//	BYTE ReadPB[1130];
	ULONG Len = 0;//Total
	BOOL Result;

	ULONG ToReplyLen;//还需发送的数据长

	int temp_Len = 0;//每次要的数据长
	int Offset = 0;

	ToReplyLen = m_strRead9BitsReplyLen;
	while(1)
	{
		if(ToReplyLen==0)
		{
			break;
		}		
		if(ToReplyLen > 50)
		{
			temp_Len = 50;
			ToReplyLen = ToReplyLen - 50;
		}
		else
		{
			temp_Len = ToReplyLen;
			ToReplyLen = 0;
		}
		Result = SLE4428Cmd_Read9Bits(
				IN	&CSerial9525,
				IN	SlotNum,
				IN  (m_strRead9BitsAddr+Offset),
				IN	temp_Len,
				OUT	(&ResponseData[Offset]),
				OUT	(&ReadPB[Offset]),
				OUT  &GetDataLen
				);
		Offset += temp_Len;
		if(Result)
		{
			m_strCardResponse = "Read fail";
			break;
		}
	}
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Read success\r\nDATA:";
			m_strCardResponse += ReadPBDataArray4428Format(ResponseData, ReadPB, (UINT)(m_strRead9BitsReplyLen));
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

void CSLE4428Dlg::OnBUTTONWriteErrorCounter() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result;
	Result = SLE4428Cmd_WriteErrorCounter(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strWriteErrorCounterData
		);
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Write Error Counter success";
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Write Error Counter fail";			
		}
		break;
	}
	UpdateData(FALSE);	
}

void CSLE4428Dlg::OnBUTTONVerify1stPSC() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result;
	Result = SLE4428Cmd_Verify1stPSC(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strVerify1stPSCData
		);	
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Verify 1st PSC success";
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Verify 1st PSC fail";			
		}
		break;
	}
	UpdateData(FALSE);	
}

void CSLE4428Dlg::OnBUTTONVerify2ndPSC() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result;
	Result = SLE4428Cmd_Verify2ndPSC(
		
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strVerify2ndPSCData
		);	
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Verify 2nd PSC success";
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Verify 2nd PSC fail";			
		}
		break;
	}
	UpdateData(FALSE);	
}

void CSLE4428Dlg::OnBUTTONEraseErrorCount() 
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);
	BOOL Result;
	Result = SLE4428Cmd_WriteEraseWithoutPB(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strEraseErrorCountAddr,
			IN	m_strEraseErrorCountData
			);

	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Erase Error success";
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Erase Error fail";			
		}
		break;
	}
	UpdateData(FALSE);	
}

void CSLE4428Dlg::OnBUTTONWriteAndEraseWithoutPB() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result;
	Result = SLE4428Cmd_WriteEraseWithoutPB(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strWriteEraseWithoutPBAddr,
			IN	m_strWriteEraseWithoutPBData
			);

	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Write success";
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Write fail";			
		}
		break;
	}
	UpdateData(FALSE);	
}

void CSLE4428Dlg::OnBUTTONWritePBAndDataComparison() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
//	int	m_strWritePBdataComparisonAddr;
//	int	m_strWritePBdataComparisonData;
//	int	m_strWritePBdataComparisonReplyLen;
	BOOL Result;
	Result = SLE4428Cmd_WritePBWithDataComparison(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strWritePBdataComparisonAddr,
			IN	m_strWritePBdataComparisonData
			);

	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Write success";
		}
		break;
		case 1://fail
		{
			m_strCardResponse = "Write fail";			
		}
		break;
	}
	UpdateData(FALSE);	
}

void CSLE4428Dlg::OnBUTTONVerifyPSCAndEraseErrorCount() 
{
	// TODO: Add your control notification handler code here
//Read Error Count
	UpdateData(TRUE);
	BYTE ResponseData[1030];	
	BYTE ReadPB[1030];
	ULONG Len = 0;
	BOOL Result;

	BYTE ErrorCount;
	Result = SLE4428Cmd_Read9Bits(
			IN	&CSerial9525,
			IN	SlotNum,
			IN  1021,
			IN	1,
			OUT	&ResponseData,
			OUT	ReadPB,
			OUT  &Len
			);
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse = "Read Error Count success\r\n";		
		}
		break;
		case 1://fail
		{
			m_strCardResponse += "Read Error Count fail";	
			UpdateData(FALSE);
			return;
		}
		break;
	}

	//-debug
	//-ResponseData[0] = 0xfe;
	//-debug

	if(ResponseData[0] == 0)
	{
		m_strCardResponse += "Error Count = 0!";	
		UpdateData(FALSE);
		return;
	}


	if((ResponseData[0]&0x01) != 0)
	{
		ErrorCount = ResponseData[0]&0xfe;
	}
	else if((ResponseData[0]&0x02) != 0)
	{
		ErrorCount = ResponseData[0]&0xfd;
	}
	else if((ResponseData[0]&0x04) != 0)
	{
		ErrorCount = ResponseData[0]&0xfb;
	}
	else if((ResponseData[0]&0x08) != 0)
	{
		ErrorCount = ResponseData[0]&0xf7;
	}
	else if((ResponseData[0]&0x1f) != 0)
	{
		ErrorCount = ResponseData[0]&0xef;
	}
	else if((ResponseData[0]&0x2f) != 0)
	{
		ErrorCount = ResponseData[0]&0xdf;
	}
	else if((ResponseData[0]&0x4f) != 0)
	{
		ErrorCount = ResponseData[0]&0xbf;
	}
	else if((ResponseData[0]&0x8f) != 0)
	{
		ErrorCount = ResponseData[0]&0x7f;
	}
	if(ErrorCount == 0)
	{
		MessageBox("ErrorCount == 0! please don't operate! That may cause a eternal damage!");
	}
//Write Error Count
	Result = SLE4428Cmd_WriteErrorCounter(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	ErrorCount
		);
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse += "Write Error Counter success\r\n";
		}
		break;
		case 1://fail
		{
			m_strCardResponse += "Write Error Counter fail";	
			UpdateData(FALSE);	
			return;
		}
		break;
	}
//verify PSC1
	Result = SLE4428Cmd_Verify1stPSC(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strPSC1
		);	
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse += "Verify 1st PSC success\r\n";
		}
		break;
		case 1://fail
		{
			m_strCardResponse += "Verify 1st PSC fail";		
			UpdateData(FALSE);	
			return;
		}
		break;
	}	
//verify PSC2
	Result = SLE4428Cmd_Verify2ndPSC(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	m_strPSC2
		);	
	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse += "Verify 2nd PSC success\r\n";
		}
		break;
		case 1://fail
		{
			m_strCardResponse += "Verify 2nd PSC fail";		
			UpdateData(FALSE);	
			return;
		}
		break;
	}	
//Write Error Count
	Result = SLE4428Cmd_WriteEraseWithoutPB(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	1021,
			IN	255
			);

	switch (Result)
	{
		case 0://成功
		{
			m_strCardResponse += "Erase Error Counter success\r\n";
		}
		break;
		case 1://fail
		{
			m_strCardResponse += "Erase Error Counter fail";
			UpdateData(FALSE);
			return;
		}
		break;
	}
	UpdateData(FALSE);
}
