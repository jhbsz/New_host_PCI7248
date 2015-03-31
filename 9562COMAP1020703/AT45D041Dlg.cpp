// AT45D041Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "AT45D041Dlg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAT45D041Dlg dialog


CAT45D041Dlg::CAT45D041Dlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAT45D041Dlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CAT45D041Dlg)
	m_strPageNum = 0;
	m_strAddressInBufferOrMemory = 0;
	m_strDataLengthToRead = 21;
	m_strDataToWrite = _T("");
	m_strDataFromCard = _T("");
	//}}AFX_DATA_INIT
}


void CAT45D041Dlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CAT45D041Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAT45D041Dlg)
	DDX_Text(pDX, IDC_PageNum, m_strPageNum);
	DDX_Text(pDX, IDC_AddressInBufferOrMemory, m_strAddressInBufferOrMemory);
	DDX_Text(pDX, IDC_DataLengthToRead, m_strDataLengthToRead);
	DDX_Text(pDX, IDC_DataToWrite, m_strDataToWrite);
	DDX_Text(pDX, IDC_DataFromCard, m_strDataFromCard);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAT45D041Dlg, CDialog)
	//{{AFX_MSG_MAP(CAT45D041Dlg)
	ON_BN_CLICKED(IDC_BUTTONMainMemoryPageRead, OnBUTTONMainMemoryPageRead)
	ON_BN_CLICKED(IDC_BUTTONBuffer1Read, OnBUTTONBuffer1Read)
	ON_BN_CLICKED(IDC_BUTTONBuffer2Read, OnBUTTONBuffer2Read)
	ON_BN_CLICKED(IDC_BUTTONMainMemoryPagetoBuffer1Xfr, OnBUTTONMainMemoryPagetoBuffer1Xfr)
	ON_BN_CLICKED(IDC_BUTTONMainMemoryPagetoBuffer2Xfr, OnBUTTONMainMemoryPagetoBuffer2Xfr)
	ON_BN_CLICKED(IDC_BUTTONMainMemoryPagetoBuffer1Compare, OnBUTTONMainMemoryPagetoBuffer1Compare)
	ON_BN_CLICKED(IDC_BUTTONMainMemoryPagetoBuffer2Compare, OnBUTTONMainMemoryPagetoBuffer2Compare)
	ON_BN_CLICKED(IDC_BUTTONBuffer1Write, OnBUTTONBuffer1Write)
	ON_BN_CLICKED(IDC_BUTTONBuffer2Write, OnBUTTONBuffer2Write)
	ON_BN_CLICKED(IDC_BUTTONBuffer1toMemoryPageProgramwithErase, OnBUTTONBuffer1toMemoryPageProgramwithErase)
	ON_BN_CLICKED(IDC_BUTTONBuffer2toMemoryPageProgramwithErase, OnBUTTONBuffer2toMemoryPageProgramwithErase)
	ON_BN_CLICKED(IDC_BUTTONBuffer1toMemoryPageProgramwithoutErase, OnBUTTONBuffer1toMemoryPageProgramwithoutErase)
	ON_BN_CLICKED(IDC_BUTTONBuffer2toMemoryPageProgramwithoutErase, OnBUTTONBuffer2toMemoryPageProgramwithoutErase)
	ON_BN_CLICKED(IDC_BUTTONMemoryPageProgramthroughBuffer1, OnBUTTONMemoryPageProgramthroughBuffer1)
	ON_BN_CLICKED(IDC_BUTTONMemoryPageProgramthroughBuffer2, OnBUTTONMemoryPageProgramthroughBuffer2)
	ON_BN_CLICKED(IDC_BUTTONAutoPageProgramthroughBuffer1, OnBUTTONAutoPageProgramthroughBuffer1)
	ON_BN_CLICKED(IDC_BUTTONAutoPageProgramthroughBuffer2, OnBUTTONAutoPageProgramthroughBuffer2)
	ON_BN_CLICKED(IDC_BUTTONGetStatusRegister, OnBUTTONGetStatusRegister)
	ON_BN_CLICKED(IDC_BUTTONTestReadSpeed, OnBUTTONTestReadSpeed)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CAT45D041Dlg, CDialog)
	//{{AFX_DISPATCH_MAP(CAT45D041Dlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IAT45D041Dlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {58118E7C-A581-4134-98F2-8325D059D184}
static const IID IID_IAT45D041Dlg =
{ 0x58118e7c, 0xa581, 0x4134, { 0x98, 0xf2, 0x83, 0x25, 0xd0, 0x59, 0xd1, 0x84 } };

BEGIN_INTERFACE_MAP(CAT45D041Dlg, CDialog)
	INTERFACE_PART(CAT45D041Dlg, IID_IAT45D041Dlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAT45D041Dlg message handlers

//将1字节字符转为1字节数.ie."9"->9;"a"->0x10
char AT45str2char(char para_str)
{
	if('A'<=para_str && para_str<='F')
	{
		para_str = para_str - 'A' + 0x0A;
	}
	else if ('a'<=para_str && para_str<='f')
	{
		para_str = para_str - 'a' + 0x0a;
	}
	else if ('0'<=para_str && para_str<='9')
	{
		para_str = para_str - '0';
	}
	return para_str;
}

int CString2int(CString DataIN)
{
	if(DataIN == "")
	{
		return 0;
	}
	int Result = 0;
	CByteArray hexdata;
	UINT i = 0;
	UINT DataOUT_len = 0;			
	char DataOUT_HL = 1;//1:转换1位，2:转换2位	
	//========去掉左,右边空格========
	CString temp_m_str = DataIN;
	temp_m_str.TrimLeft(" ");
	temp_m_str.TrimRight(" ");
	char* temp_CMD = (LPSTR)(LPCTSTR) temp_m_str;
	while(1)
	{
		if(temp_CMD[i] == 0x00)
		{
			break;
		}
		Result = (Result*16) + AT45str2char(temp_CMD[i]);
		i++;
	}
	return Result;
}

void CAT45D041Dlg::OnBUTTONMainMemoryPageRead() 
{
	// TODO: Add your control notification handler code here
//	if (CString2int("ffe") == 4094)
//	MessageBox("good");
/*
	CString	m_strPageNum;
	CString	m_strAddressInBufferOrMemory;
	CString	m_strDataLengthToRead;
	CString	m_strDataToWrite;
	CString	m_strDataFromCard;
*/

//	int PageNum = CString2int(m_strPageNum);
//	int AddressInBufferOrMemory = CString2int(m_strAddressInBufferOrMemory);
//	int DataLengthToRead = CString2int()

/*	LONG APIENTRY AT45D041Cmd(
		IN	UCHAR	OPcode,
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	int	PageNo,
		IN	int	lngStartAddr,
		IN	ULONG	lngWriteLen,
		IN	LPVOID	pWriteData,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		)*/
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x52,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  m_strDataLengthToRead,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Read Main Memory Page fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = Bytes2CString(GetData,(UINT)ReturnLen);
	UpdateData(FALSE);
}

void CAT45D041Dlg::OnBUTTONBuffer1Read() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x54,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  m_strDataLengthToRead,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Read Buffer1 fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = Bytes2CString(GetData,(UINT)ReturnLen);
	UpdateData(FALSE);	
}

void CAT45D041Dlg::OnBUTTONBuffer2Read() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x56,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  m_strDataLengthToRead,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Read Buffer1 fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = Bytes2CString(GetData,(UINT)ReturnLen);
	UpdateData(FALSE);		
}

void CAT45D041Dlg::OnBUTTONMainMemoryPagetoBuffer1Xfr() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x53,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Main Memory Page to Buffer1 Xfr fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Main Memory Page to Buffer1 Xfr Success";
	UpdateData(FALSE);		
}

void CAT45D041Dlg::OnBUTTONMainMemoryPagetoBuffer2Xfr() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x55,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Main Memory Page to Buffer2 Xfr fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Main Memory Page to Buffer2 Xfr Success";
	UpdateData(FALSE);	
}

void CAT45D041Dlg::OnBUTTONMainMemoryPagetoBuffer1Compare() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x60,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Main Memory Page to Buffer1 Compare fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Main Memory Page to Buffer1 Compare Success";
	UpdateData(FALSE);	
}

void CAT45D041Dlg::OnBUTTONMainMemoryPagetoBuffer2Compare() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x61,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Main Memory Page to Buffer2 Compare fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Main Memory Page to Buffer2 Compare Success";
	UpdateData(FALSE);		
}

void CAT45D041Dlg::OnBUTTONBuffer1Write() 
{
	// TODO: Add your control notification handler code here
/*	UpdateData(TRUE);
    //转换发送数据
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strDataToWrite);

	BOOL Result = 0;	
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x84,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  Len,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Buffer1 Write fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Buffer1 Write Success";
	UpdateData(FALSE);	*/	

	UpdateData(TRUE);
    //转换发送数据
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strDataToWrite);

	BOOL Result = 0;	
	BYTE GetData[264];

	ULONG ReturnLen = 0;
	if(Len<180)
	{
		Result = AT45D041Cmd(
						IN  0x84,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  m_strAddressInBufferOrMemory,
						IN  Len,
						IN  WriteData,
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Buffer1 Write fail!";
			UpdateData(FALSE);	
			return;
		}
		m_strDataFromCard = "Buffer1 Write Success";
		UpdateData(FALSE);	
	}
	else
	{
		Result = AT45D041Cmd(
						IN  0x84,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  m_strAddressInBufferOrMemory,
						IN  179,
						IN  WriteData,
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Buffer1 Write fail!";
			UpdateData(FALSE);	
			return;
		}
		Result = AT45D041Cmd(
						IN  0x84,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  (m_strAddressInBufferOrMemory+179),
						IN  (Len - 179),
						IN  (WriteData+179),
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Buffer1 Write fail!";
			UpdateData(FALSE);	
			return;
		}
		m_strDataFromCard = "Buffer1 Write Success";
		UpdateData(FALSE);	
	}
}

void CAT45D041Dlg::OnBUTTONBuffer2Write() 
{
	// TODO: Add your control notification handler code here
/*	UpdateData(TRUE);
    //转换发送数据
	//BYTE WriteData[264];
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strDataToWrite);

	BOOL Result = 0;	
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x87,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  Len,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Buffer1 Write fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Buffer1 Write Success";
	UpdateData(FALSE);	*/	
	UpdateData(TRUE);
    //转换发送数据
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strDataToWrite);

	BOOL Result = 0;	
	BYTE GetData[264];

	ULONG ReturnLen = 0;
	if(Len<180)
	{
		Result = AT45D041Cmd(
						IN  0x87,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  m_strAddressInBufferOrMemory,
						IN  Len,
						IN  WriteData,
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Buffer2 Write fail!";
			UpdateData(FALSE);	
			return;
		}
		m_strDataFromCard = "Buffer2 Write Success";
		UpdateData(FALSE);	
	}
	else
	{
		Result = AT45D041Cmd(
						IN  0x87,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  m_strAddressInBufferOrMemory,
						IN  179,
						IN  WriteData,
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Buffer2 Write fail!";
			UpdateData(FALSE);	
			return;
		}
		Result = AT45D041Cmd(
						IN  0x87,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  (m_strAddressInBufferOrMemory+179),
						IN  (Len - 179),
						IN  (WriteData+179),
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Buffer2 Write fail!";
			UpdateData(FALSE);	
			return;
		}
		m_strDataFromCard = "Buffer2 Write Success";
		UpdateData(FALSE);	
	}
}

void CAT45D041Dlg::OnBUTTONBuffer1toMemoryPageProgramwithErase() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x83,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Buffer1 to Memory Page Program with Erase fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Buffer1 to Memory Page Program with Erase Success";
	UpdateData(FALSE);	
}

void CAT45D041Dlg::OnBUTTONBuffer2toMemoryPageProgramwithErase() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x86,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Buffer2 to Memory Page Program with Erase fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Buffer2 to Memory Page Program with Erase Success";
	UpdateData(FALSE);	
}

void CAT45D041Dlg::OnBUTTONBuffer1toMemoryPageProgramwithoutErase() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x88,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Buffer1 to Memory Page Program without Erase fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Buffer1 to Memory Page Program without Erase Success";
	UpdateData(FALSE);		
}

void CAT45D041Dlg::OnBUTTONBuffer2toMemoryPageProgramwithoutErase() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x89,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Buffer2 to Memory Page Program without Erase fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Buffer2 to Memory Page Program without Erase Success";
	UpdateData(FALSE);		
}

void CAT45D041Dlg::OnBUTTONMemoryPageProgramthroughBuffer1() 
{
	// TODO: Add your control notification handler code here
/*	UpdateData(TRUE);
    //转换发送数据
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strDataToWrite);

	BOOL Result = 0;	
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x82,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  Len,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Memory Page Program through Buffer1 fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Memory Page Program through Buffer1 Success";
	UpdateData(FALSE);	*/
	UpdateData(TRUE);
    //转换发送数据
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strDataToWrite);

	BOOL Result = 0;	
	BYTE GetData[264];

	ULONG ReturnLen = 0;
	if(Len<180)
	{
		Result = AT45D041Cmd(
						IN  0x82,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  m_strAddressInBufferOrMemory,
						IN  Len,
						IN  WriteData,
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Memory Page Program through Buffer1 fail!";
			UpdateData(FALSE);	
			return;
		}
		m_strDataFromCard = "Memory Page Program through Buffer1 Success";
		UpdateData(FALSE);	
	}
	else
	{
		Result = AT45D041Cmd(
						IN  0x82,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  m_strAddressInBufferOrMemory,
						IN  179,
						IN  WriteData,
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Memory Page Program through Buffer1 fail!";
			UpdateData(FALSE);	
			return;
		}
		Result = AT45D041Cmd(
						IN  0x82,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  (m_strAddressInBufferOrMemory+179),
						IN  (Len - 179),
						IN  (WriteData+179),
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Memory Page Program through Buffer1 fail!";
			UpdateData(FALSE);	
			return;
		}
		m_strDataFromCard = "Memory Page Program through Buffer1 Success";
		UpdateData(FALSE);	
	}
}

void CAT45D041Dlg::OnBUTTONMemoryPageProgramthroughBuffer2() 
{
	// TODO: Add your control notification handler code here
/*	UpdateData(TRUE);
    //转换发送数据
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strDataToWrite);

	BOOL Result = 0;	
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x85,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  Len,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Memory Page Program through Buffer2 fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Memory Page Program through Buffer2 Success";
	UpdateData(FALSE);		*/
	UpdateData(TRUE);
    //转换发送数据
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strDataToWrite);

	BOOL Result = 0;	
	BYTE GetData[264];

	ULONG ReturnLen = 0;
	if(Len<180)
	{
		Result = AT45D041Cmd(
						IN  0x85,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  m_strAddressInBufferOrMemory,
						IN  Len,
						IN  WriteData,
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Memory Page Program through Buffer2 fail!";
			UpdateData(FALSE);	
			return;
		}
		m_strDataFromCard = "Memory Page Program through Buffer2 Success";
		UpdateData(FALSE);	
	}
	else
	{
		Result = AT45D041Cmd(
						IN  0x85,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  m_strAddressInBufferOrMemory,
						IN  179,
						IN  WriteData,
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Memory Page Program through Buffer2 fail!";
			UpdateData(FALSE);	
			return;
		}
		Result = AT45D041Cmd(
						IN  0x85,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  m_strPageNum,
						IN  (m_strAddressInBufferOrMemory+179),
						IN  (Len - 179),
						IN  (WriteData+179),
						IN  0,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard = "Memory Page Program through Buffer2 fail!";
			UpdateData(FALSE);	
			return;
		}
		m_strDataFromCard = "Memory Page Program through Buffer2 Success";
		UpdateData(FALSE);	
	}
}

void CAT45D041Dlg::OnBUTTONAutoPageProgramthroughBuffer1() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x58,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Auto Page Program through Buffer1 fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Auto Page Program through Buffer1 Success";
	UpdateData(FALSE);		
}

void CAT45D041Dlg::OnBUTTONAutoPageProgramthroughBuffer2() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x59,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  0,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Auto Page Program through Buffer2 fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Auto Page Program through Buffer2 Success";
	UpdateData(FALSE);		
}

void CAT45D041Dlg::OnBUTTONGetStatusRegister() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 0;
	BYTE WriteData[264];
	BYTE GetData[264];
	ULONG ReturnLen = 0;
	Result = AT45D041Cmd(
					IN  0x57,
					IN	&CSerial9525,
					IN	SlotNum,
					IN  m_strPageNum,
					IN  m_strAddressInBufferOrMemory,
					IN  0,
					IN  WriteData,
					IN  1,
					OUT GetData,
					OUT &ReturnLen);
	if(Result == 1)
	{
		m_strDataFromCard = "Get Status Register fail!";
		UpdateData(FALSE);	
		return;
	}
	m_strDataFromCard = "Status Register: 0x";
	m_strDataFromCard += Bytes2CString(GetData,(UINT)ReturnLen);
	UpdateData(FALSE);	
}

void CAT45D041Dlg::OnBUTTONTestReadSpeed() 
{
	// TODO: Add your control notification handler code here
	UINT i;
	BOOL Result;
	BYTE WriteData[264];
	BYTE GetData[264];
	CString PageNo = "";
	m_strDataFromCard = "";
//	for(i=0;i<2048;i++)
	for(i=0;i<2048;i++)
	{
		Result = 0;
		ULONG ReturnLen = 0;
		if(i==8)
		{
			i=8;
		}
		Result = AT45D041Cmd(
						IN  0x52,
						IN	&CSerial9525,
						IN	SlotNum,
						IN  i,
						IN  0,
						IN  0,
						IN  WriteData,
						IN  264,
						OUT GetData,
						OUT &ReturnLen);
		if(Result == 1)
		{
			m_strDataFromCard += "Test fail!";
			UpdateData(FALSE);	
			return;
		}
		PageNo.Format("Page: %d",i);
		GetDlgItem(IDC_STATIC_00)->SetWindowText(PageNo);
	}
}
