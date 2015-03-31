// 6710MPToolDlg.cpp : implementation file
//

#include "stdafx.h"
#include "6710MPTool.h"
#include "6710MPToolDlg.h"

#include "mmsystem.h"

#include "Dask.h"
#include <wtypes.h>
#include <initguid.h>
#include "guids.h"
#include "usbusr.h"
#include <tchar.h>  // Allen 20080319

//#include <usbioctl.h>
#include <winusb.h>
//#include <usbprint.h>

#define MAX_LOADSTRING 256

extern "C" {

// This file is in the Windows DDK available from Microsoft.

#include <setupapi.h>
#include <dbt.h>
}


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
//==============User define zone start==============
//-------USB API Part-------
ULONG								Length;
PSP_DEVICE_INTERFACE_DETAIL_DATA	detailData;
HANDLE								DeviceHandle;
LPGUID								WetGuid;
HANDLE								hDevInfo;
ULONG								Required;
LPOVERLAPPED						lpOverLap;

typedef void (*LPfnScsi2usb2K_KillEXE)(HANDLE);
HINSTANCE hDll=NULL;

LPfnScsi2usb2K_KillEXE lpfnScsi2usb2K_KillEXE;

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




//-------AP Part-------
class PortParameter
{
public:

	BYTE Flag_Port;//1:active, 0: inactive
	BYTE PortSpeed;//3:High speed
	HANDLE PortDevInfoHandle;
	HANDLE PortDeviceHandle;
	HANDLE PortInterfaceHandle;
	void InitiateParameter(void);
};
void PortParameter::InitiateParameter(void)
{
	Flag_Port = 0;
	PortDevInfoHandle = 0;
	PortDevInfoHandle = 0;
	PortInterfaceHandle = 0;
}
PortParameter PortParameter1, PortParameter2;
BYTE PortNum = 0;//record how much Port is active

LPVOID	m_pPortParam;

BYTE SendDataBuffer_0x55[2048];
BYTE SendDataBuffer_0xAA[2048];
BYTE ReceiveDataBuffer[2048];

BYTE Flag_Timer = 0;//1: the timer is running, 0: the timer is not running. 

BYTE DeviceDescriptor[18];
BYTE String0Descriptor[256];
BYTE String1Descriptor[256];
BYTE String2Descriptor[256];
BYTE String3PortADescriptor[256];
BYTE String3PortBDescriptor[256];
BYTE StringEE[256];
BYTE StringExtendConfigurationDescriptor[256];
BYTE stringExtendPropertiesDescriptor[256];

BYTE TestWritePipeData[2048];

BYTE TestData_AA55AA55AA55AA55[2048];
BYTE TestData_AAAA5555AAAA5555[2048];
BYTE TestData_5555AAAA5555AAAA[2048];
BYTE TestData_55AA55AA55AA55AA[2048];
BYTE TestData_AAAAAAAA55555555[2048];
BYTE TestData_55555555AAAAAAAA[2048];
//==============User define zone end==============

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMy6710MPToolDlg dialog

CMy6710MPToolDlg::CMy6710MPToolDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMy6710MPToolDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CMy6710MPToolDlg)
	m_strResult = _T("");
	m_strPassNum = 0;
	m_strFailNum = 0;
	m_strTestInformation = _T("");
	m_strPID = _T("2795");
	m_strVID = _T("0576");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMy6710MPToolDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CMy6710MPToolDlg)
	DDX_Control(pDX, IDC_SpeedError, m_SpeedError);
	DDX_Control(pDX, IDC_LedFail, m_LedFail);
	DDX_Control(pDX, IDC_RWFail, m_RWFail);
	DDX_Control(pDX, IDC_TestResult, m_TestResult);
	DDX_Control(pDX, IDC_UnknowDevice, m_UnknowDevice);
	DDX_Control(pDX, IDC_TEST_INFO, m_Infor);
	DDX_Control(pDX, IDC_CHIP_STATUS, m_ctrlChipIcon);
	DDX_Text(pDX, IDC_Result, m_strResult);
	DDX_Text(pDX, IDC_PASS_NUM, m_strPassNum);
	DDX_Text(pDX, IDC_FAIL_NUM, m_strFailNum);
	DDX_Text(pDX, IDC_TEST_INFO, m_strTestInformation);
	DDX_Text(pDX, IDC_PID, m_strPID);
	DDX_Text(pDX, IDC_VID, m_strVID);
	DDX_Control(pDX, IDC_MSCOMM1, m_Comm);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CMy6710MPToolDlg, CDialog)
	//{{AFX_MSG_MAP(CMy6710MPToolDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_START, OnButtonStart)
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_BUTTON_Exit, OnBUTTONExit)
	ON_BN_CLICKED(IDC_BUTTON_CLEAR, OnButtonClear)
	ON_EN_CHANGE(IDC_VID, OnChangeVid)
	ON_EN_CHANGE(IDC_PID, OnChangePid)
	ON_BN_CLICKED(IDC_BUTTON_OpenFile, OnBUTTONOpenFile)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMy6710MPToolDlg message handlers

BOOL CMy6710MPToolDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	m_ctrlChipIcon.SetIcon(AfxGetApp()->LoadIcon(IDI_ICON_NULL));
	m_strResult = "";
	m_strPassNum = 0;
	m_strFailNum = 0;
	m_strTestInformation = "";	
	UpdateData(FALSE);


	PortParameter1.InitiateParameter();
	PortParameter2.InitiateParameter();

	memset(SendDataBuffer_0x55,0x55,2048);
	memset(SendDataBuffer_0xAA,0xAA,2048);
//debug start
//	MessageBox("1");
//	Delay(10000);//10 s
//	MessageBox("2");
//start end

	UINT i = 0;
	for(i=0;i<2048;i+=8)
	{
		TestData_AA55AA55AA55AA55[i]   = 0xAA;
		TestData_AA55AA55AA55AA55[i+1] = 0x55;
		TestData_AA55AA55AA55AA55[i+2] = 0xAA;
		TestData_AA55AA55AA55AA55[i+3] = 0x55;
		TestData_AA55AA55AA55AA55[i+4] = 0xAA;
		TestData_AA55AA55AA55AA55[i+5] = 0x55;
		TestData_AA55AA55AA55AA55[i+6] = 0xAA;
		TestData_AA55AA55AA55AA55[i+7] = 0x55;

		TestData_AAAA5555AAAA5555[i]   = 0xAA;
		TestData_AAAA5555AAAA5555[i+1] = 0xAA;
		TestData_AAAA5555AAAA5555[i+2] = 0x55;
		TestData_AAAA5555AAAA5555[i+3] = 0x55;
		TestData_AAAA5555AAAA5555[i+4] = 0xAA;
		TestData_AAAA5555AAAA5555[i+5] = 0xAA;
		TestData_AAAA5555AAAA5555[i+6] = 0x55;
		TestData_AAAA5555AAAA5555[i+7] = 0x55;

		TestData_5555AAAA5555AAAA[i]   = 0x55;
		TestData_5555AAAA5555AAAA[i+1] = 0x55;
		TestData_5555AAAA5555AAAA[i+2] = 0xAA;
		TestData_5555AAAA5555AAAA[i+3] = 0xAA;
		TestData_5555AAAA5555AAAA[i+4] = 0x55;
		TestData_5555AAAA5555AAAA[i+5] = 0x55;
		TestData_5555AAAA5555AAAA[i+6] = 0xAA;
		TestData_5555AAAA5555AAAA[i+7] = 0xAA;

		TestData_55AA55AA55AA55AA[i]   = 0x55;
		TestData_55AA55AA55AA55AA[i+1] = 0x55;
		TestData_55AA55AA55AA55AA[i+2] = 0xAA;
		TestData_55AA55AA55AA55AA[i+3] = 0xAA;
		TestData_55AA55AA55AA55AA[i+4] = 0x55;
		TestData_55AA55AA55AA55AA[i+5] = 0x55;
		TestData_55AA55AA55AA55AA[i+6] = 0xAA;
		TestData_55AA55AA55AA55AA[i+7] = 0xAA;

		TestData_AAAAAAAA55555555[i]   = 0xAA;
		TestData_AAAAAAAA55555555[i+1] = 0xAA;
		TestData_AAAAAAAA55555555[i+2] = 0xAA;
		TestData_AAAAAAAA55555555[i+3] = 0xAA;
		TestData_AAAAAAAA55555555[i+4] = 0x55;
		TestData_AAAAAAAA55555555[i+5] = 0x55;
		TestData_AAAAAAAA55555555[i+6] = 0x55;
		TestData_AAAAAAAA55555555[i+7] = 0x55;

		TestData_55555555AAAAAAAA[i]   = 0x55;
		TestData_55555555AAAAAAAA[i+1] = 0x55;
		TestData_55555555AAAAAAAA[i+2] = 0x55;
		TestData_55555555AAAAAAAA[i+3] = 0x55;
		TestData_55555555AAAAAAAA[i+4] = 0xAA;
		TestData_55555555AAAAAAAA[i+5] = 0xAA;
		TestData_55555555AAAAAAAA[i+6] = 0xAA;
		TestData_55555555AAAAAAAA[i+7] = 0xAA;
	}
	PortNum = 0;

	TestInit();
	SetTimer(20,500, FALSE);
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CMy6710MPToolDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMy6710MPToolDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMy6710MPToolDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CMy6710MPToolDlg::OnButtonStart() 
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);
	m_ctrlChipIcon.SetIcon(AfxGetApp()->LoadIcon(IDI_ICON_NULL));
	m_strResult = "";
	m_strTestInformation = "";	
	UpdateData(FALSE);
	GetDlgItem(IDC_BUTTON_START)->EnableWindow(FALSE);
	Flag_Timer = 1;
	SetTimer(ID_CLOCK,10000,NULL);
   // begin
      // load usb remove clear dll
 
	hDll=AfxLoadLibrary("USBRESET.DLL");

    if (hDll == NULL)
	{
		AfxMessageBox("Could not load Dll");
	}
	else
	{
		lpfnScsi2usb2K_KillEXE=(LPfnScsi2usb2K_KillEXE)GetProcAddress(hDll,"fnScsi2usb2K_KillEXE");
	}
 

	long t1;   	// test cycle



	// EndLess contorl loop
    CString StrOutput;
	StrOutput="Ready";
	CString ChipName,StrInput;
	 MSG msg;
	COleVariant Buf;
	do{ 
		time(&t1);  // get time
       // Clear  information
	   //	 OnButtonClear();
       // Initial UI
      
		ChipName="";
		do
		{

		m_Comm.SetOutput(COleVariant(StrOutput));
		
		Sleep(10);
		Buf=m_Comm.GetInput();
		 
		lpfnScsi2usb2K_KillEXE;
		ChipName=ChipName + (CString)(Buf.bstrVal);
	//	m_Unknow_Device.SetWindowText(ChipName);
		
		 while( ::PeekMessage( &msg, NULL, 0, 0, PM_NOREMOVE ) ) 
		 {
		  ::GetMessage( &msg, NULL, 0, 0 );
		  ::TranslateMessage( &msg );
		  ::DispatchMessage( &msg );
		 } 
		} while ((ChipName!="AU6710ASF22") && (AllenStop==0));
     //     UpdateData(TRUE);
		m_SpeedError.SetBkColor(RGB(255,255,255));
		m_LedFail.SetBkColor(RGB(255,255,255));
		m_RWFail.SetBkColor(RGB(255,255,255));
		m_TestResult.SetBkColor(RGB(255,255,255));
		m_UnknowDevice.SetBkColor(RGB(255,255,255));

		 
		m_TestResult.SetWindowText("Wait for Testing");
      
	 	m_strResult = "";
 
     	m_strTestInformation = "";
    //    UpdateData(FALSE);
        // PCI power off
        
      //   DO_WritePort(card, Channel_P2A, 0xFF);

		// PCI power on
		 
	    
         
         DO_WritePort(card, Channel_P1A, 0x0);
         Sleep(2000);
	   

         
		 StartTest();
	     m_Infor.LineScroll(m_Infor.GetLineCount());
         DO_WritePort(card, Channel_P1A, 0xFF);

   //  MSG msg;
	 while( ::PeekMessage( &msg, NULL, 0, 0, PM_NOREMOVE ) ) 
	 {
	  ::GetMessage( &msg, NULL, 0, 0 );
	  ::TranslateMessage( &msg );
	  ::DispatchMessage( &msg );
	 } 
	
	}  while (AllenStop==0);
//	timeGetTime();



  // end


 

    
//	timeGetTime();
}

void CMy6710MPToolDlg::StartTest()
{
	//Use a series of API calls to find a WET with a specified Vendor IF and Product ID.
	SP_DEVICE_INTERFACE_DATA			devInfoData;
	bool								LastDevice = FALSE;
	int									MemberIndex = 0;
	BOOL								Result;	
	CString								UsageDescription;
//	LPCTSTR								strVendorID, strProductID;
	
    U32  DI_P;  // PCI input value
		
    DO_ReadPort(card, Channel_P1B, &DI_P );	
      // load bin file descriptor

	PortNum = 0;
	Length = 0;
	hDevInfo = NULL;
	detailData = NULL;
	DeviceHandle=NULL;	
    
	WetGuid = (LPGUID) &GUID_CLASS_WET;	
	/*
	API function: SetupDiGetClassDevs
	Returns: a handle to a device information set for all installed devices.
	Requires: the WET GUID.
	*/
	hDevInfo=SetupDiGetClassDevs 
		(WetGuid, 
		NULL, 
		NULL, 
		DIGCF_PRESENT|DIGCF_INTERFACEDEVICE);
		
	devInfoData.cbSize = sizeof(devInfoData);

	//Step through the available devices looking for the one we want. 
	//Quit on detecting the desired device or checking all available devices without success.
	MemberIndex = 0;
	LastDevice = FALSE;
	do
	{
		/*
		API function: SetupDiEnumDeviceInterfaces
		On return, MyDeviceInterfaceData contains the handle to a
		SP_DEVICE_INTERFACE_DATA structure for a detected device.
		Requires:
		The DeviceInfoSet returned in SetupDiGetClassDevs.
		The WET Guid.
		An index to specify a device.
		*/
		Result = FALSE; 

		Result=SetupDiEnumDeviceInterfaces 
			(hDevInfo, 
			0, 
			WetGuid, 
			MemberIndex, 
			&devInfoData);

		if (Result != 0)
		{
			//A device has been detected, so get more information about it.

			/*
			API function: SetupDiGetDeviceInterfaceDetail
			Returns: an SP_DEVICE_INTERFACE_DETAIL_DATA structure
			containing information about a device.
			To retrieve the information, call this function twice.
			The first time returns the size of the structure in Length.
			The second time returns a pointer to the data in DeviceInfoSet.
			Requires:
			A DeviceInfoSet returned by SetupDiGetClassDevs
			The SP_DEVICE_INTERFACE_DATA structure returned by SetupDiEnumDeviceInterfaces.
			
			The final parameter is an optional pointer to an SP_DEV_INFO_DATA structure.
			This application doesn't retrieve or use the structure.			
			If retrieving the structure, set 
			MyDeviceInfoData.cbSize = length of MyDeviceInfoData.
			and pass the structure's address.
			*/
			
			//Get the Length value.
			//The call will return with a "buffer too small" error which can be ignored.
			Result = SetupDiGetDeviceInterfaceDetail 
				(hDevInfo, 
				&devInfoData, 
				NULL, 
				0, 
				&Length, 
				NULL);

			//Allocate memory for the hDevInfo structure, using the returned Length.
			detailData = (PSP_DEVICE_INTERFACE_DETAIL_DATA)malloc(Length);
			
			//Set cbSize in the detailData structure.
			detailData -> cbSize = sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA);

			//Call the function again, this time passing it the returned buffer size.
			Result = SetupDiGetDeviceInterfaceDetail 
				(hDevInfo, 
				&devInfoData, 
				detailData, 
				Length, 
				&Required, 
				NULL);

			//Is it the desired device?
//			strVendorID = "058f";//m_VendorIDString;
//			strProductID = "6710";//m_ProductIDString;
			m_strVID.MakeLower();
			m_strPID.MakeLower();
//			MessageBox(detailData->DevicePath);
			if(strstr(detailData->DevicePath, m_strVID))//m_strVID
			{
				if(strstr(detailData->DevicePath, m_strPID))//m_strPID
				{
					//Both the Vendor ID and Product ID match.
					//---------------------------------------------------------------------------
					// Open a handle to the device.
					// To enable retrieving information about a system mouse or keyboard,
					// don't request Read or Write access for this handle.

					/*
					API function: CreateFile
					Returns: a handle that enables reading and writing to the device.
					Requires:
					The DevicePath in the detailData structure
					returned by SetupDiGetDeviceInterfaceDetail.
					*/
					DeviceHandle=CreateFile 
						(detailData->DevicePath, 
						GENERIC_WRITE|GENERIC_READ, 
						FILE_SHARE_READ|FILE_SHARE_WRITE, 
						NULL,
						OPEN_EXISTING, 
						0, 
						NULL);

					HANDLE InterfaceHandle;
					Result = WinUsb_Initialize(DeviceHandle,&InterfaceHandle);
					
					BYTE WriteBuffer[0x100];
					ULONG ReturnLen;

					ULONG  InformationType;
					InformationType = DEVICE_SPEED;
					ReturnLen = 1;
					Result = WinUsb_QueryDeviceInformation(InterfaceHandle,
												InformationType,
												&ReturnLen,
												WriteBuffer);
					//用这个Handle建立一个进程
					//Handle功能: 发送数据后，接收到数据退出。
					PortNum++;
					if(PortNum == 1)
					{
						PortParameter1.Flag_Port = 1;
						PortParameter1.PortSpeed = WriteBuffer[0];
						PortParameter1.PortDevInfoHandle = hDevInfo;
						PortParameter1.PortDeviceHandle = DeviceHandle;
						PortParameter1.PortInterfaceHandle = InterfaceHandle;
					}
					else if(PortNum == 2)
					{
						PortParameter2.Flag_Port = 1;
						PortParameter2.PortSpeed = WriteBuffer[0];
						PortParameter2.PortDevInfoHandle = hDevInfo;
						PortParameter2.PortDeviceHandle = DeviceHandle;
						PortParameter2.PortInterfaceHandle = InterfaceHandle;
					}	
					else if(PortNum >= 3)
					{
						WinUsb_Free(InterfaceHandle);
						CloseHandle(DeviceHandle);
						SetupDiDestroyDeviceInfoList(hDevInfo);					
					}
				} 
			}
			else
			{
				SetupDiDestroyDeviceInfoList(hDevInfo);			
			}
			//Free the memory used by the detailData structure (no longer needed).
			free(detailData);
		}
		//if (Result != 0)
		else
		{
			//SetupDiEnumDeviceInterfaces returned 0, so there are no more devices to check.
			LastDevice=TRUE;
		}
		//If we haven't found the device yet, and haven't tried every available device,
		//try the next one.
		MemberIndex = MemberIndex + 1;
	} //do
	while (LastDevice == FALSE);

	//----------------------------------------------------------
	//Show information about 2 ports' detection, 2ports' speed
	//----------------------------------------------------------
	switch (PortNum)
	{
		case 0:
		{
			ShowInformation("No Port detected!");
			ShowTestFail();
			UnknowDeviceFail();
			return;
		}
		break;
		case 1:
		{
			WinUsb_Free(PortParameter1.PortInterfaceHandle);
			CloseHandle(PortParameter1.PortDeviceHandle);
			SetupDiDestroyDeviceInfoList(PortParameter1.PortDevInfoHandle);

			ShowInformation("Only one Port detected!");
			ShowTestFail();
			UnknowDeviceFail();
			return;
		}
			break;
		case 2:
		{
			ShowInformation("Two port detected!");	
			UnknowDevicePass();
		}
		break;
		default:
		{
			ClearHandlesOfTwoPorts();
			ShowInformation("More than 2 devices' port is Detected");
			ShowTestFail();	
			UnknowDeviceFail();
			return;
		}
		break;
	}
	ShowInformation("-------------------------------------------------------------------------");
	if(PortParameter1.PortSpeed !=3)
	{
		ClearHandlesOfTwoPorts();

		ShowInformation("Port1 is not in High speed!");
		ShowTestFail();		
		SpeedErrorFail();
		return;	
	}
	else
	{
		ShowInformation("Port1 is in High speed!");	
	}
	if(PortParameter2.PortSpeed !=3)
	{
		ClearHandlesOfTwoPorts();

		ShowInformation("Port2 is not in High speed!");
		ShowTestFail();		
		SpeedErrorFail();
		return;	
	}
	else
	{
		ShowInformation("Port2 is in High speed!");	
		SpeedErrorPass();
	}
	ShowInformation("-------------------------------------------------------------------------");
	//---------------------------------------------
	//Debug for HongWu Start
	//---------------------------------------------
	//--------------------------------
	//port1 send DebugData to port2;
	//--------------------------------
//	BYTE DebugData[8] = {0x41,0x00,0x00,0x00,0x08,0x00,0x1A,0X7D};// ERROR

//	BYTE DebugData[8] = {0x55,0x55,0x55,0x55,0x55,0x55,0x55,0xFF};// ERROR
//	BYTE DebugData[8] = {0x55,0x55,0x55,0x5d,0x55,0x5d,0x55,0xFF};// ERROR Reply


//	BYTE DebugData[8] = {0x55,0x55,0x55,0x55,0x55,0x55,0x55,0x00};// RIGHT
//	BYTE DebugData[8] = {0x55,0x55,0x55,0x5d,0x55,0x5d,0x55,0xFF};// RIGHT
//	BYTE DebugData[8] = {0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF};// RIGHT
//	BYTE DebugData[8] = {0x55,0x55,0x55,0x55,0x55,0x55,0x55,0x55};// RIGHT
//	BYTE DebugData[8] = {0x41,0x00,0x00,0x00,0x08,0x00,0x1A,0X00};// RIGHT


//	BYTE DebugData[8] = {0x55,0x55,0x55,0x55,0x55,0x55,0x55,0x08};// ERROR
//	BYTE DebugData[8] = {0x55,0x55,0x55,0x5D,0x55,0x5D,0x55,0x08};// ERROR Reply

//	BYTE DebugData[8] = {0x55,0x55,0x55,0xAA,0x55,0x55,0x55,0x55};//ERROR
//	BYTE DebugData[8] = {0x55,0x55,0x55,0xA2,0x55,0x55,0x55,0x55};//ERROR Reply


//	BYTE DebugData[8] = {0x55,0x55,0x55,0xAA,0x55,0x55,0x55,0xAA};//ERROR
//	BYTE DebugData[8] = {0x55,0x55,0x55,0xAA,0x55,0x5D,0x55,0xAA};//ERROR Reply

//	BYTE DebugData[4] = {0x55,0x55,0x55,0X08};//ERROR

//	BYTE DebugData[4] = {0xAA,0x55,0xAA,0X55};//ERROR
//	BYTE DebugData[4] = {0xAA,0x55,0xAA,0X5D};//ERROR Reply
//	BYTE DebugData[4] = {0xAA,0xAA,0x55,0X55};
//	BYTE DebugData[4] = {0x55,0X55,0xAA,0xAA};

/*	BYTE DebugData[512] = {0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,
							0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA,0x55,0X55,0xAA,0xAA};
	UINT SendLength = 512;

	ULONG ReturnDebugDataLen;
	Result = WinUsb_WritePipe(PortParameter1.PortInterfaceHandle,
								1,
								DebugData,
								SendLength,
								&ReturnDebugDataLen,
								lpOverLap);
	if(Result != 1)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("port1 send DebugData to easy transfer's Buffer ---- fail");
		ShowTestFail();	
		return;
	}
	ShowInformation("port1 send DebugData to easy transfer's Buffer ----- success");

	Result = WinUsb_ReadPipe(PortParameter2.PortInterfaceHandle,
						0X82,
						ReceiveDataBuffer,
						SendLength,
						&ReturnDebugDataLen,
						lpOverLap);

	if(Result != 1)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 receive the debug data from easy transfer's Buffer ---- fail");
		ShowTestFail();	
		return;
	}
	ShowInformation("Port2 receive the debug data from easy transfer's Buffer ---- success");

	int CompareResultDebugData = 0;
	CompareResultDebugData = memcmp(DebugData, ReceiveDataBuffer, SendLength);
	if(CompareResultDebugData != 0)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("the data Port2 received is not identical to the debug data");
		ShowTestFail();	
		return;
	}
	ShowInformation("the data Port2 received is identical the debug data");
	ShowInformation("-------------------------------------------------------------------------");





	Result = WinUsb_WritePipe(PortParameter2.PortInterfaceHandle,
								1,
								DebugData,
								SendLength,
								&ReturnDebugDataLen,
								lpOverLap);
	if(Result != 1)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("port2 send DebugData to easy transfer's Buffer ---- fail");
		ShowTestFail();	
		return;
	}
	ShowInformation("port2 send DebugData to easy transfer's Buffer ----- success");

	Result = WinUsb_ReadPipe(PortParameter1.PortInterfaceHandle,
						0X82,
						ReceiveDataBuffer,
						SendLength,
						&ReturnDebugDataLen,
						lpOverLap);

	if(Result != 1)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 receive the debug data from easy transfer's Buffer ---- fail");
		ShowTestFail();	
		return;
	}
	ShowInformation("Port1 receive the debug data from easy transfer's Buffer ---- success");

	CompareResultDebugData = 0;
	CompareResultDebugData = memcmp(DebugData, ReceiveDataBuffer, SendLength);
	if(CompareResultDebugData != 0)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("the data Port1 received is not identical to the debug data");
		ShowTestFail();	
		return;
	}
	ShowInformation("the data Port1 received is identical the debug data");
	ShowInformation("-------------------------------------------------------------------------");



	ClearHandlesOfTwoPorts();
	ShowTestSuccess();
	return;*/
	//---------------------------------------------
	//Debug for HongWu End
	//---------------------------------------------

	//-------------------------------------------------------------
	//Test Control Pipe
	//-------------------------------------------------------------
	WINUSB_SETUP_PACKET  SetupPacket;
	ULONG ReturnLength = 0;
	BYTE DataLinkDescriptor[0x40];
	SetupPacket.Length = 0x40;

	//Get Device Descriptor from Port 1
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0100;
	SetupPacket.Index = 0x00;	
	Result = WinUsb_ControlTransfer(PortParameter1.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 get Device descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port1 get Device descriptor ----- success");
	if(0 != memcmp(DeviceDescriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare Device descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare Device descriptor ---- success");
	//Get Device Descriptor from Port 2
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0100;
	SetupPacket.Index = 0x00;	
	Result = WinUsb_ControlTransfer(PortParameter2.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 get Device descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port2 get Device descriptor ----- success");
	if(0 != memcmp(DeviceDescriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare Device descriptor ---- fail");
		ShowTestFail();	
		
        RWDataFail();
		return;	
	}
	ShowInformation("Compare Device descriptor ---- success");

	//Get String 0 Descriptor from port1
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0300;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter1.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 get string 0 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port1 get string 0 descriptor ---- success");
	if(0 != memcmp(String0Descriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string 0 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string 0 descriptor ---- success");
	//Get String0 Descriptor from port2
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0300;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter2.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 get string 0 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port2 get string 0 descriptor ---- success");
	if(0 != memcmp(String0Descriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string 0 descriptor ---- fail");
		ShowTestFail();
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string 0 descriptor ---- success");

	//Get String1 Descriptor from port1
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0301;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter1.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 get String 1 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port1 get String 1 descriptor ---- success");
	if(0 != memcmp(String1Descriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string 1 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string 1 descriptor ---- success");
	//Get String1 Descriptor from port2
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0301;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter2.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 get String 1 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port2 get String 1 descriptor ---- success");
	if(0 != memcmp(String1Descriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string 1 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string 1 descriptor ---- success");

	//Get String2 Descriptor from port1
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0302;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter1.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 get String 2 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port1 get String 2 descriptor ---- success");
	if(0 != memcmp(String2Descriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string 2 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string 2 descriptor ---- success");
	//Get String2 Descriptor from port2
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0302;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter2.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 get String 2 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port2 get String 2 descriptor ---- success");
	if(0 != memcmp(String2Descriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string 2 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string 2 descriptor ---- success");

	//Get String 3 Descriptor from port1
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0303;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter1.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 get String 3 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port1 get String 3 descriptor ---- success");
	if(0 != memcmp(String3PortADescriptor, DataLinkDescriptor, ReturnLength))
		
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string 3 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string 3 descriptor ---- success");
	//Get String 3 Descriptor from port2
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x0303;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter2.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 get String 3 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port2 get String 3 descriptor ---- success");
	if(0 != memcmp(String3PortADescriptor, DataLinkDescriptor, ReturnLength))	
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string 3 descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string 3 descriptor ---- success");

	//Get String EE Descriptor from port1
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x03EE;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter1.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 get string EE descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port1 get string EE descriptor ---- success");
	if(0 != memcmp(StringEE, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string EE descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string EE descriptor ---- success");
	//Get String EE Descriptor from port2
	SetupPacket.RequestType = 0x80;
	SetupPacket.Request = 0x06;
	SetupPacket.Value = 0x03EE;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter2.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 get string EE descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;
	}
	ShowInformation("Port2 get string EE descriptor ---- success");
	if(0 != memcmp(StringEE, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string EE descriptor ---- fail");
		ShowTestFail();	
		RWDataFail();
		return;	
	}
	ShowInformation("Compare string EE descriptor ---- success");

	//Get String extended configuration from port1
/*	SetupPacket.RequestType = 0xC0;
	SetupPacket.Request = 0x0C;
	SetupPacket.Value = 0x00;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter1.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 get string extended configuration descriptor ---- fail");
		ShowTestFail();	
		return;
	}
	ShowInformation("Port1 get string extended configuration descriptor ---- success");
	if(0 != memcmp(StringExtendConfigurationDescriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string extended configuration descriptor ---- fail");
		ShowTestFail();	
		return;	
	}
	ShowInformation("Compare string extended configuration descriptor ---- success");
	//Get String extended configuration from port2
	SetupPacket.RequestType = 0xC0;
	SetupPacket.Request = 0x0C;
	SetupPacket.Value = 0x00;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter2.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 get string extended configuration descriptor ---- fail");
		ShowTestFail();	
		return;
	}
	ShowInformation("Port2 get string extended configuration descriptor ---- success");
	if(0 != memcmp(StringExtendConfigurationDescriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string extended configuration descriptor ---- fail");
		ShowTestFail();	
		return;	
	}
	ShowInformation("Compare string extended configuration descriptor ---- success");

	//Get String extended properity from port1
	SetupPacket.RequestType = 0xC1;
	SetupPacket.Request = 0x0C;
	SetupPacket.Value = 0x00;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter1.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port1 get Device extended properity descriptor ---- fail");
		ShowTestFail();	
		return;
	}
	ShowInformation("Port2 get Device extended properity descriptor ---- success");
	if(0 != memcmp(stringExtendPropertiesDescriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string extended configuration descriptor ---- fail");
		ShowTestFail();	
		return;	
	}
	ShowInformation("Compare string extended configuration descriptor ---- success");
	//Get String extended properity from port2
	SetupPacket.RequestType = 0xC1;
	SetupPacket.Request = 0x0C;
	SetupPacket.Value = 0x00;
	SetupPacket.Index = 0x00;
	Result = WinUsb_ControlTransfer(PortParameter2.PortInterfaceHandle,
									SetupPacket,
									DataLinkDescriptor,
									0x40,
									&ReturnLength,
									lpOverLap);
	if(!Result)
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Port2 get Device extended properity descriptor ---- fail");
		ShowTestFail();	
		return;
	}
	ShowInformation("Port2 get Device extended properity descriptor ---- success");
	if(0 != memcmp(stringExtendPropertiesDescriptor, DataLinkDescriptor, ReturnLength))
	{
		ClearHandlesOfTwoPorts();
		ShowInformation("Compare string extended configuration descriptor ---- fail");
		ShowTestFail();	
		return;	
	}
	ShowInformation("Compare string extended configuration descriptor ---- success");
*/
	ShowInformation("-------------------------------------------------------------------------");

	//-------------------------------------------------------
	//Test Data Pipe
	//-------------------------------------------------------
/*
BYTE TestData_AA55AA55AA55AA55[2048];
BYTE TestData_AAAA5555AAAA5555[2048];
BYTE TestData_5555AAAA5555AAAA[2048];
BYTE TestData_55AA55AA55AA55AA[2048];
BYTE TestData_AAAAAAAA55555555[2048];
BYTE TestData_55555555AAAAAAAA[2048];
*/
	BYTE SendDataNum = 6;
	while(SendDataNum--)
	{
		switch(SendDataNum)
		{
			case 5:
			{
				memcpy(TestWritePipeData,TestData_AA55AA55AA55AA55,2048);
			}
			break;
			case 4:
			{
				memcpy(TestWritePipeData,TestData_AAAA5555AAAA5555,2048);
			}
			break;
			case 3:
			{
				memcpy(TestWritePipeData,TestData_5555AAAA5555AAAA,2048);
			}
			break;
			case 2:
			{
				memcpy(TestWritePipeData,TestData_55AA55AA55AA55AA,2048);
			}
			break;
			case 1:
			{
				memcpy(TestWritePipeData,TestData_AAAAAAAA55555555,2048);
			}
			break;
			case 0:
			{
				memcpy(TestWritePipeData,TestData_55555555AAAAAAAA,2048);
			}
			break;
		}
		CString TestNO;
		TestNO.Format("NO: %d",(6-SendDataNum));
		ShowInformation(TestNO);
		
		ULONG ReturnLen;
		//-----------------------------------------------
		//port1 send TestWritePipeData[2048] to port2
		//-----------------------------------------------
		Result = WinUsb_WritePipe(PortParameter1.PortInterfaceHandle,
									1,
									TestWritePipeData,
									2048,
									&ReturnLen,
									lpOverLap);
		if(Result != 1)
		{
			ClearHandlesOfTwoPorts();
			ShowInformation("Port1 transfer TestWritePipeData[2048] to easy transfer's Buffer ---- fail");
			ShowTestFail();	
			RWDataFail();
			return;
		}
		ShowInformation("Port1 send TestWritePipeData[2048] to easy transfer's Buffer ----- success");
		Result = WinUsb_ReadPipe(PortParameter2.PortInterfaceHandle,
						0X82,
						ReceiveDataBuffer,
						2048,
						&ReturnLen,
						lpOverLap);
		if(Result != 1)
		{
			ClearHandlesOfTwoPorts();
			ShowInformation("Port2 receive TestWritePipeData[2048] from easy transfer's Buffer ---- fail");
			ShowTestFail();	
			RWDataFail();
			return;
		}
		ShowInformation("Port2 receive TestWritePipeData[2048] from easy transfer's Buffer ---- success");
		int CompareResult = 0;
		CompareResult = memcmp(TestWritePipeData, ReceiveDataBuffer, 2048);
		if(CompareResult != 0)
		{
			ClearHandlesOfTwoPorts();
			ShowInformation("the data Port2 received is not identical to TestWritePipeData[2048]");
			ShowTestFail();	
			RWDataFail();
			return;
		}
		ShowInformation("the data Port2 received is identical to TestWritePipeData[2048]");
		//-----------------------------------------------
		//port2 send TestWritePipeData[2048] to port1
		//-----------------------------------------------
		Result = WinUsb_WritePipe(PortParameter2.PortInterfaceHandle,
									1,
									TestWritePipeData,
									2048,
									&ReturnLen,
									lpOverLap);
		if(Result != 1)
		{
			ClearHandlesOfTwoPorts();
			ShowInformation("Port2 transfer TestWritePipeData[2048] to easy transfer's Buffer ---- fail");
			ShowTestFail();	
			RWDataFail();
			return;
		}
		ShowInformation("Port2 send TestWritePipeData[2048] to easy transfer's Buffer ----- success");
		Result = WinUsb_ReadPipe(PortParameter1.PortInterfaceHandle,
						0X82,
						ReceiveDataBuffer,
						2048,
						&ReturnLen,
						lpOverLap);
		if(Result != 1)
		{
			ClearHandlesOfTwoPorts();
			ShowInformation("Port1 receive TestWritePipeData[2048] from easy transfer's Buffer ---- fail");
			ShowTestFail();	
			RWDataFail();
			return;
		}
		ShowInformation("Port1 receive TestWritePipeData[2048] from easy transfer's Buffer ---- success");
		CompareResult = 0;
		CompareResult = memcmp(TestWritePipeData, ReceiveDataBuffer, 2048);
		if(CompareResult != 0)
		{
			ClearHandlesOfTwoPorts();
			ShowInformation("the data Port1 received is not identical to TestWritePipeData[2048]");
			ShowTestFail();	
			RWDataFail();
			return;
		}
		ShowInformation("the data Port1 received is identical to TestWritePipeData[2048]");
		ShowInformation("-------------------------------------------------------------------------");
		RWDataPass();
	}

    // LED test
    if (DI_P !=0xFE) 
	{
		    ClearHandlesOfTwoPorts();
		    ShowInformation("LED Fail!");
		    ShowTestFail();	
		    LedFunctionFail();
			return;
		 
	}
	else
	{
	ShowInformation("LED PASS!");
	ShowInformation("-------------------------------------------------------------------------");
 
     LedFunctionPass();
    }
 


	ClearHandlesOfTwoPorts();
	ShowTestSuccess();

    m_TestResult.SetWindowText("Bin1 Pass");
	m_TestResult.SetBkColor(RGB(0,255,0));
	m_Comm.SetOutput(COleVariant("PASS"));
      


}

void CMy6710MPToolDlg::OnTimer(UINT nIDEvent) 
{
	// TODO: Add your message handler code here and/or call default
	if ( nIDEvent == 20 )
	{
		KillTimer(20);
		OnButtonStart();
	}
	Flag_Timer = 0;	
	
	CDialog::OnTimer(nIDEvent);
}
void CMy6710MPToolDlg::OnBUTTONExit() 
{
	// TODO: Add your control notification handler code here
	AllenStop=1;
	EndDialog(IDOK);
}

void CMy6710MPToolDlg::Delay(DWORD dwDelayTime)
{
	DWORD dwTimeBegin,dwTimeEnd;
    dwTimeBegin=timeGetTime();
    do
    {
		dwTimeEnd=timeGetTime();
    }
	while(dwTimeEnd-dwTimeBegin<dwDelayTime);
}

void CMy6710MPToolDlg::OnButtonClear() 
{
	// TODO: Add your control notification handler code here
	m_ctrlChipIcon.SetIcon(AfxGetApp()->LoadIcon(IDI_ICON_NULL));
	m_strResult = "";
//	m_strPassNum = 0;
//	m_strFailNum = 0;
	m_strTestInformation = "";	
	UpdateData(FALSE);
}

void CMy6710MPToolDlg::ShowInformation(CString ShowData)
{
	m_strTestInformation += ShowData;
	m_strTestInformation +="\r\n";
	UpdateData(FALSE);
	 

}

void CMy6710MPToolDlg::ShowTestFail()
{
	m_ctrlChipIcon.SetIcon(AfxGetApp()->LoadIcon(IDI_ICON_FAIL));
	m_strResult = "Fail";
	m_strFailNum ++ ;
	ShowInformation("Test Fail");
	GetDlgItem(IDC_BUTTON_START)->EnableWindow(TRUE);
	UpdateData(FALSE);
}

void CMy6710MPToolDlg::ShowTestSuccess()
{
	m_ctrlChipIcon.SetIcon(AfxGetApp()->LoadIcon(IDI_ICON_PASS));
	m_strResult = "PASS";
	m_strPassNum ++ ;
	ShowInformation("Test Success");
	GetDlgItem(IDC_BUTTON_START)->EnableWindow(TRUE);
	UpdateData(FALSE);
}		 

void CMy6710MPToolDlg::ClearHandlesOfTwoPorts()
{
	WinUsb_Free(PortParameter1.PortInterfaceHandle);
	CloseHandle(PortParameter1.PortDeviceHandle);
	SetupDiDestroyDeviceInfoList(PortParameter1.PortDevInfoHandle);

	WinUsb_Free(PortParameter2.PortInterfaceHandle);
	CloseHandle(PortParameter2.PortDeviceHandle);
	SetupDiDestroyDeviceInfoList(PortParameter2.PortDevInfoHandle);
}

void CMy6710MPToolDlg::OnChangeVid() 
{
	// TODO: If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.
	
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
}

void CMy6710MPToolDlg::OnChangePid() 
{
	// TODO: If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.
	
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);	
}

void CMy6710MPToolDlg::OnBUTTONOpenFile() 
{
	// TODO: Add your control notification handler code here
	UCHAR m_abBinData[5000];
	UINT m_iBinDataLen;

	CString strFilter;
	CFile	BinFile;
//	int	hResult;
//	int iActual;
/*
 	strFilter = "Binary File(*.bin)|*.bin||"; 

 	CFileDialog dlg(TRUE,NULL,NULL,OFN_FILEMUSTEXIST,strFilter); 
 	hResult = (int)dlg.DoModal(); 
 	if (hResult != IDOK) 
	{ 
 		return; 
	} 
*/   // Allen
	
   // Allen 20080320
   TCHAR FilePath[MAX_PATH];
   GetModuleFileName(NULL,FilePath,MAX_PATH);
   (_tcsrchr(FilePath,'\\'))[1]=0;
   lstrcat(FilePath,_T("BAFO.bin"));


	CFile myFile;
	CFileException fileException;
//	if ( !myFile.Open( dlg.GetFileName(), 
//			  CFile::modeReadWrite, &fileException ) )
 	if ( !myFile.Open( FilePath, 
 			  CFile::modeReadWrite, &fileException ) )
	{
		MessageBox("Can't open file", "Open file", MB_OK );
		return;
	}
	
	m_iBinDataLen = 0;
	m_iBinDataLen = myFile.Read( m_abBinData, sizeof( m_abBinData ) );	
	//------Get Device descriptor------
	memcpy(DeviceDescriptor, (m_abBinData + m_abBinData[2]),18);
	//------Get String0------
	memcpy(String0Descriptor, (m_abBinData + m_abBinData[4]),m_abBinData[5]);
	//------Get String1------
	memcpy(String1Descriptor, (m_abBinData + m_abBinData[6]),m_abBinData[7]);
	//------Get String2------
	memcpy(String2Descriptor, (m_abBinData + m_abBinData[8]),m_abBinData[9]);
	//------Get String3PortA------
	memcpy(String3PortADescriptor, (m_abBinData + m_abBinData[10]),m_abBinData[11]);
	//------Get String3PortB------
//	memcpy(String3PortBDescriptor, (m_abBinData + m_abBinData[18]),m_abBinData[19]);
	//------Get StringEE------
	memcpy(StringEE, (m_abBinData + m_abBinData[12]),m_abBinData[13]);
	//------Get String extended configuration------
	memcpy(StringExtendConfigurationDescriptor, (m_abBinData + m_abBinData[14]),m_abBinData[15]);
	//------Get String extended properity------
	memcpy(stringExtendPropertiesDescriptor, (m_abBinData + m_abBinData[16]),m_abBinData[17]);
}

void CMy6710MPToolDlg::TestInit()
{

        //set  com port
    
	m_Comm.SetCommPort(1);
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
		SetTimer_1ms();
		
   		m_SpeedError.SetBkColor(RGB(255,255,255));
		m_LedFail.SetBkColor(RGB(255,255,255));
		m_RWFail.SetBkColor(RGB(255,255,255));
		m_TestResult.SetBkColor(RGB(255,255,255));
		m_UnknowDevice.SetBkColor(RGB(255,255,255));

		 OnBUTTONOpenFile();
}

void CMy6710MPToolDlg::UnknowDeviceFail()
{
       
	  m_UnknowDevice.SetBkColor(RGB(255,0,0));
	  m_TestResult.SetBkColor(RGB(255,0,0));
      m_TestResult.SetWindowText("Bin2 Fail");
      m_Comm.SetOutput(COleVariant("Bin2"));
      
}

void CMy6710MPToolDlg::UnknowDevicePass()
{
      m_UnknowDevice.SetBkColor(RGB(0,255,0));
	  m_TestResult.SetBkColor(RGB(255,0,0));
}

void CMy6710MPToolDlg::SpeedErrorFail()
{
      m_SpeedError.SetBkColor(RGB(255,0,0));
	  m_TestResult.SetBkColor(RGB(255,0,0));
      m_TestResult.SetWindowText("Bin3 Fail");
	   m_Comm.SetOutput(COleVariant("Bin3"));
}

void CMy6710MPToolDlg::SpeedErrorPass()
{
      m_SpeedError.SetBkColor(RGB(0,255,0));
	  m_TestResult.SetBkColor(RGB(255,0,0));
}

void CMy6710MPToolDlg::RWDataFail()
{
      m_RWFail.SetBkColor(RGB(255,0,0));
	  m_TestResult.SetBkColor(RGB(255,0,0));
      m_TestResult.SetWindowText("Bin4 Fail");
	   m_Comm.SetOutput(COleVariant("Bin4"));
}

void CMy6710MPToolDlg::RWDataPass()
{
      m_RWFail.SetBkColor(RGB(0,255,0));
	  m_TestResult.SetBkColor(RGB(255,0,0));
}

void CMy6710MPToolDlg::LedFunctionFail()
{
      m_LedFail.SetBkColor(RGB(255,0,0));
	  m_TestResult.SetBkColor(RGB(255,0,0));
      m_TestResult.SetWindowText("Bin5 Fail");
	  
	  m_Comm.SetOutput(COleVariant("Bin5"));
}

void CMy6710MPToolDlg::LedFunctionPass()
{
      m_LedFail.SetBkColor(RGB(0,255,0));
	  m_TestResult.SetBkColor(RGB(255,0,0));
}

void CMy6710MPToolDlg::SetTimer_1ms()
{
	int err;
	err = CTR_Setup(card, 1, RATE_GENERATOR, 200, BINTimer);
	err = CTR_Setup(card, 2, RATE_GENERATOR, 10, BINTimer);
}

void CMy6710MPToolDlg::Timer_1ms(int ms)
{
	I16 result; 
	U32 old_value1, old_value2;
 
	result = CTR_Read(0, 2, &old_value1);
	for(int i = 1; i < ms; i++){
		do{
            result = CTR_Read(0, 2, &old_value2);
		}while(!(old_value1 != old_value2));
    
		do{           
            result = CTR_Read(0, 2, &old_value2);
		}while(!(old_value1 == old_value2));
	}
}

void CMy6710MPToolDlg::PCI7248_bin(BYTE Channel, BYTE PCI7248bin)
{
	I16 k;
	U32 DO_P;

	DO_P = PCI7248bin;
	Timer_1ms(7);  // Allen
    k = DO_WritePort(card, Channel, DO_P);
    Timer_1ms(7);
    //========================================
	DO_P = PCI7248bin - PCI7248_EOT;
    k = DO_WritePort(card, Channel, DO_P);
    Timer_1ms(7);
    //=======================================
	DO_P = PCI7248bin;
    k = DO_WritePort(card, Channel, DO_P);
    Timer_1ms(7);
    //========================================
	DO_P = 0xFF;
	k = DO_WritePort(card, Channel, DO_P);
        
}


