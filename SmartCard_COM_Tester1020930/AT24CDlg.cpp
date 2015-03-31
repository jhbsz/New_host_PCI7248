// AT24CDlg.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "AT24CDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAT24CDlg dialog


CAT24CDlg::CAT24CDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAT24CDlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CAT24CDlg)
	m_strReadAddress = 0;
	m_strPageSize = 16;
	m_strReadLength = 1;
	m_strWriteAddress = 0;
	m_strAccessData = _T("");
	m_strEEPROMdata = _T("");
	//}}AFX_DATA_INIT
}


void CAT24CDlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CAT24CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAT24CDlg)
	DDX_Text(pDX, IDC_ReadAddress, m_strReadAddress);
	DDX_Text(pDX, IDC_PageSize, m_strPageSize);
	DDX_Text(pDX, IDC_ReadLength, m_strReadLength);
	DDX_Text(pDX, IDC_WriteAddress, m_strWriteAddress);
	DDX_Text(pDX, IDC_AccessData, m_strAccessData);
	DDX_Text(pDX, IDC_EEPROMdata, m_strEEPROMdata);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAT24CDlg, CDialog)
	//{{AFX_MSG_MAP(CAT24CDlg)
	ON_BN_CLICKED(IDC_ButtonWrite, OnButtonWrite)
	ON_BN_CLICKED(IDC_ButtonRead, OnButtonRead)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CAT24CDlg, CDialog)
	//{{AFX_DISPATCH_MAP(CAT24CDlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IAT24CDlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {A9BE795A-FD0E-470C-9785-F486A47BF9AC}
static const IID IID_IAT24CDlg =
{ 0xa9be795a, 0xfd0e, 0x470c, { 0x97, 0x85, 0xf4, 0x86, 0xa4, 0x7b, 0xf9, 0xac } };

BEGIN_INTERFACE_MAP(CAT24CDlg, CDialog)
	INTERFACE_PART(CAT24CDlg, IID_IAT24CDlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAT24CDlg message handlers

void CAT24CDlg::OnButtonWrite() 
{
	// TODO: Add your control notification handler code here
/*
	m_strReadAddress = _T("");
	m_strPageSize = _T("");
	m_strReadLength = _T("");
	m_strWriteAddress = _T("");
	m_strAccessData = _T("");
	m_strEEPROMdata = _T("");
*/
	UpdateData(TRUE);
    //转换发送数据
	BYTE WriteData[300];
	UINT Len;
	Len = CString2Bytes (WriteData, m_strAccessData);
	BOOL Result;

	Result = AT24CxxCmd_Write(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	0x50,
			IN	m_strWriteAddress,
			IN	m_strPageSize,
			IN	Len,
			IN	WriteData
			);
	if(Result)
	{
		m_strAccessData = "Write fail";	
		UpdateData(FALSE);
		return;
	}

	BYTE GetData[300];
	ULONG GetDataLen = 0;
	Result = AT24CxxCmd_Read(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	0x50,
			IN	0,
			IN	256,
			OUT	GetData,
			OUT	&GetDataLen
			);
	switch (Result)
	{
		case 0://成功
		{
			m_strEEPROMdata = Bytes2CString(GetData,(UINT)GetDataLen);
		}
		break;
		case 1://fail
		{
			m_strEEPROMdata = "Get EEPROM data fail!";			
		}
		break;
	}
	UpdateData(FALSE);	

}

void CAT24CDlg::OnButtonRead() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	BOOL Result = 1;
	BYTE GetData[300];
	ULONG GetDataLen = 0;
	Result = AT24CxxCmd_Read(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	0x50,
			IN	m_strReadAddress,
			IN	m_strReadLength,
			OUT	GetData,
			OUT	&GetDataLen
			);
	if(Result)
	{
		m_strAccessData = "Read fail!";	
		UpdateData(FALSE);
		return;
	}
	m_strAccessData = Bytes2CString(GetData,(UINT)GetDataLen);

	Result = AT24CxxCmd_Read(
			IN	&CSerial9525,
			IN	SlotNum,
			IN	0x50,
			IN	0,
			IN	256,
			OUT	GetData,
			OUT	&GetDataLen
			);
	switch (Result)
	{
		case 0://成功
		{
			m_strEEPROMdata = Bytes2CString(GetData,(UINT)GetDataLen);
		}
		break;
		case 1://fail
		{
			m_strEEPROMdata = "Get EEPROM data fail!";			
		}
		break;
	}
	UpdateData(FALSE);		
}
