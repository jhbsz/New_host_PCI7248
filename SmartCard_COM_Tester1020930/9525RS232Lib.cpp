#include "stdafx.h"
#include "9525RS232Lib.h"

//"#define RS232ResendTimes  2" means that total send 3 times
#define RS232ResendTimes 2

#define SetBaudRate                  0x66
#define GetReaderDescriptor          0x67
#define	Response_SetBaudRate         0x86
#define Response_GetReaderDescriptor 0x87

#define LCDEnReg                    0xEA
#define LCDDataReg                  0xEB
#define LCDRWReg                    0xEC
#define LCDRSReg                    0xED
#define LCDBLReg                    0xEE
#define LCDCSReg                    0xEF
#define CMDBusy                     0xF0
#define CMDDelay                    0xF1

#define RESET                           0xFF
#define LOW                             0x00
#define HIGH                            0x01
#define CMD_CODE_POS                    0x05
#define BUSY_POS                        0x0A
#define CMD_CODE_DEFAULT                0x00
#define WRITE_DATA_POS                  0x05
#define WRITE_DATA_DEF                  0x00
#define	EEPROM_WRITE						0xC2
#define	EEPROM_READ							0xC3
#define	FIRMWARE_WRITE						0xC4
#define	FIRMWARE_READ						0xC5
#define LCD_SET_CURSOR                      0xE0
#define LCD_CLEAR_DISPLAY                   0xE1
#define LCD_DISPLAY_MESSAGE                 0xE2
#define LCD_DISPLAY_GRAPHIC                 0xE3
#define LCD_SET_CONTRAST                    0xE4
#define LCD_SET_BACKLIGHT                   0xE5
#define CLEAR_KEY_BUFFER                    0xD0
#define SET_KEY_SCAN_TIMER                  0xD1
#define GET_KEY_INPUT                       0xD2
#define LCD_WRITE_DATA                      0xD3
/* Alcor vendor command */
#define ALCOR_HEADER_LENGTH					0x08
#define ALCOR_OP_CODE						0x40
#define ALCOR_OP_CODE_LCM_KEYPAD			0x50
//Alcor vendor command for syn card
#define OP_SWITCH_CARD_MODE                 0x50
#define OP_POWER_ON                         0x51
#define OP_POWER_OFF                        0x52
#define OP_SET_I2C_ADD                      0x60
#define OP_WRITE_I2C                        0x61
#define OP_READ_I2C                         0x62
#define OP_AT45D041_CARD_COMMAND			0x64
#define OP_SMC_COMMAND						0x70
#define OP_SLE4442_CARD_COMMAND				0x80
#define OP_SLE4442_CARD_BREAK				0x81
#define OP_SLE4428_CARD_COMMAND				0x82
#define OP_INPHONE_CARD_RESET				0x90
#define OP_INPHONE_CARD_Read				0x91
#define OP_INPHONE_CARD_PROG				0x92
#define OP_INPHONE_CARD_MOVE_ADDRESS        0x93
#define OP_INPHONE_CARD_AUTHENTICATION_KEY1 0x94
#define OP_INPHONE_CARD_AUTHENTICATION_KEY2 0x95
#define OP_SET_CONFIG						0xC0
#define OP_SET_LED							0xC1

//CCID SendCommand code
#define PC_to_RDR_IccPowerOn      0x62
#define PC_to_RDR_IccPowerOff     0x63
#define PC_to_RDR_GetSlotStatus   0x65
#define PC_to_RDR_XfrBlock        0x6F
#define PC_to_RDR_GetParameters   0x6C
#define PC_to_RDR_ResetParameters 0x6D
#define PC_to_RDR_SetParameters   0x61
#define PC_to_RDR_Escape          0x6B
#define PC_to_RDR_IccClock        0x6E
#define PC_to_RDR_T0APDU          0x6A
//CCID ResponseCommand code
#define RDR_to_PC_NotifySlotChange 0x85
#define RDR_to_PC_DataBlock        0x80
#define RDR_to_PC_SlotStatus       0x81
#define RDR_to_PC_Parameters       0X82
#define RDR_to_PC_Escape           0x83
//Error Code
#define CMD_ABORTED             0xFF
#define ICC_MUTE                0xFE
#define XFR_PARITY_ERROR        0xFD
#define XFR_OVERRUN             0xFC
#define HW_ERROR                0xFB
#define BUSY_WITH_AUTO_SEQUENCE 0xF2
#define CMD_SLOT_BUSY           0xE0
#define Command_not_supported   0x00

int CSerial::Open( int nPort, int nBaud )
{
	if( m_bOpened ) return( TRUE );

	char szPort[15];
	DCB dcb;

	wsprintf( szPort, "COM%d", nPort );
	m_hComDev = CreateFile( szPort, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED, NULL );
	if( m_hComDev == NULL ) return( FALSE );

	memset( &m_OverlappedRead, 0, sizeof( OVERLAPPED ) );
	memset( &m_OverlappedWrite, 0, sizeof( OVERLAPPED ) );

	COMMTIMEOUTS CommTimeOuts;
	CommTimeOuts.ReadIntervalTimeout = 0xFFFFFFFF;
	CommTimeOuts.ReadTotalTimeoutMultiplier = 0;
	CommTimeOuts.ReadTotalTimeoutConstant = 0;
	CommTimeOuts.WriteTotalTimeoutMultiplier = 0;
	CommTimeOuts.WriteTotalTimeoutConstant = 5000;
	SetCommTimeouts( m_hComDev, &CommTimeOuts );

	m_OverlappedRead.hEvent = CreateEvent( NULL, TRUE, FALSE, NULL );
	m_OverlappedWrite.hEvent = CreateEvent( NULL, TRUE, FALSE, NULL );

	dcb.DCBlength = sizeof( DCB );
	GetCommState( m_hComDev, &dcb );
	dcb.BaudRate = nBaud;
	dcb.ByteSize = 8;
	if( !SetCommState( m_hComDev, &dcb ) ||
		!SetupComm( m_hComDev, 10000, 10000 ) ||
		m_OverlappedRead.hEvent == NULL ||
		m_OverlappedWrite.hEvent == NULL ){
		long dwError = GetLastError();
		if( m_OverlappedRead.hEvent != NULL ) CloseHandle( m_OverlappedRead.hEvent );
		if( m_OverlappedWrite.hEvent != NULL ) CloseHandle( m_OverlappedWrite.hEvent );
		CloseHandle( m_hComDev );
		return FALSE;
	}
	m_bOpened = TRUE;

	return m_bOpened;
}

void CSerial::ClosePort(void)
{
	m_bOpened = FALSE;
	CloseHandle( m_hComDev );
}

BOOL CATR_capacity::Do_ATR(BYTE *DataIN, UINT Len)
{
	FI_DI = 0;
	T_Type = 0;

	BYTE Offset = 0;
	BYTE T0_Value;
	BYTE TD1,TD2,TD3;
	T0_Value = DataIN[1];
	Offset = 2;
	//-----------------------
	//判断TA1是否存在
	//-----------------------
	if((T0_Value & 0x10) == 0)
	{
		FI_DI = 0x11;
	}
	else
	{
		FI_DI = DataIN[2];
	}
	//-----------------------
	//判断TD1是否存在
	//-----------------------
	if((T0_Value & 0x80) == 0)
	{
		T_Type = T0;
	}
	//分析TD1
	else
	{
		BYTE i = 0;
		if((T0_Value & 0x10) != 0)//TA1
		{
			Offset++;
		}
		if((T0_Value & 0x20) != 0)//TB1
		{
			Offset++;
		}
		if((T0_Value & 0x40) != 0)//TC1
		{
			Offset++;
		}
		TD1 = DataIN[Offset];
		Offset++;
		switch(TD1 & 0x0F)
		{
			case 0:
			{
				T_Type = T_Type | T0;
			}
			break;
			case 1:
			{
				T_Type = T_Type | T1;
			}
			break;
		}

		//-----------------------
		//判断TD2是否存在
		//-----------------------
		if((TD1 & 0x80) == 0)
		{
			;	
		}
		//分析TD2
		else
		{
			if((TD1 & 0x10) != 0)//TA2
			{
				Offset++;
			}
			if((TD1 & 0x20) != 0)//TB2
			{
				Offset++;
			}
			if((TD1 & 0x40) != 0)//TC2
			{
				Offset++;
			}
			TD2 = DataIN[Offset];
			Offset++;
			switch(TD2 & 0x0F)
			{
				case 0:
				{
					T_Type = T_Type | T0;
				}
				break;
				case 1:
				{
					T_Type = T_Type | T1;
				}
				break;
			}
			//-----------------------
			//判断TD3是否存在
			//-----------------------
			if((TD2 & 0x80) == 0)
			{
				;	
			}
			//分析TD3
			else
			{
				if((TD2 & 0x10) != 0)//TA3
				{
					Offset++;
				}
				if((TD2 & 0x20) != 0)//TB3
				{
					Offset++;
				}
				if((TD2 & 0x40) != 0)//TC3
				{
					Offset++;
				}
				TD3 = DataIN[Offset];
				switch(TD3 & 0x0F)
				{
					case 0:
					{
						T_Type = T_Type | T0;
					}
					break;
					case 1:
					{
						T_Type = T_Type | T1;
					}
					break;
				}
			}
		}	
	}
	//pCATR_capacity->T_Type的值为01 or 10 or 11

	return 1;
}
int CSerial::InBufferCount( void )
{

	if( !m_bOpened || m_hComDev == NULL ) return( 0 );

	unsigned long dwErrorFlags;
	COMSTAT ComStat;

	ClearCommError( m_hIDComDev, &dwErrorFlags, &ComStat );

	return (int)ComStat.cbInQue;
	
}

long CSerial::SendData( const char *buffer, unsigned long dwBytesWritten)
{
	if( !m_bOpened || m_hComDev == NULL ) return( 0 );
	BOOL bWriteStat;
	bWriteStat = WriteFile( m_hComDev, buffer, dwBytesWritten, &dwBytesWritten, &m_OverlappedWrite );
	if( !bWriteStat){
		if ( GetLastError() == ERROR_IO_PENDING ) {
			WaitForSingleObject( m_OverlappedWrite.hEvent, 1000 );
			return dwBytesWritten;
		}
		return 0;
	}
	return dwBytesWritten;
}

//等2.5秒，收数中30ms无数据退出。
long CSerial::ReadData( void *buffer, unsigned long dwBytesRead)
{
	Sleep(200);
	ULONG Offset = 0;
	dwBytesRead = 0;
	BOOL AlreadyReceiveData = 0;
	if( !m_bOpened || m_hComDev == NULL ) return 0;
	UINT iDataLength = 0;

	BYTE GetDataPer10ms[300];
	UINT i;
	ULONG DataLenPer10ms = 0;
	for(i=0;i<65535;i++)//330
	{	
//		Sleep(2);//待修改。 
//		Sleep(1);
		BOOL bReadStatus;
		unsigned long dwErrorFlags;
		COMSTAT ComStat;

		ClearCommError( m_hComDev, &dwErrorFlags, &ComStat );
		if( ComStat.cbInQue == 0 ) 
		{
			if(AlreadyReceiveData == 1)
			{
				if(dwBytesRead>2)
				{
					iDataLength = *((PUCHAR)buffer+1) + (*((PUCHAR)buffer+2))*256;
					if(iDataLength == dwBytesRead - 11)
					{
						return dwBytesRead;
					}
				}

			}
			else
			{		
				continue;
			}
		}
		else
		{
			AlreadyReceiveData = 1;
			//DEBUG
//			BYTE TempSenddata[5] = {0x5a, 0xa5, 0x11, 0x22, 0x33};
//			SendData((char *)TempSenddata,5);
			//DEBUG
		}
		
		DataLenPer10ms = ComStat.cbInQue;
		DataLenPer10ms = min(300, DataLenPer10ms);

		bReadStatus = ReadFile( m_hComDev, GetDataPer10ms, DataLenPer10ms, &DataLenPer10ms, &m_OverlappedRead );
		if( !bReadStatus )
		{
			if( GetLastError() == ERROR_IO_PENDING )
			{
				WaitForSingleObject( m_OverlappedRead.hEvent, 2000 );
				return dwBytesRead;
			}
			return 0;
		}
		dwBytesRead += DataLenPer10ms;

		memcpy( ((PUCHAR)buffer + Offset),
				GetDataPer10ms, 
				DataLenPer10ms);
		Offset += DataLenPer10ms;
		if(dwBytesRead>2)
		{
			iDataLength = *((PUCHAR)buffer+1) + (*((PUCHAR)buffer+2))*256;
			if(iDataLength == dwBytesRead - 11)
			{
				return dwBytesRead;
			}
		}
	}
	return dwBytesRead;
}

long CSerial::ReadDataWithoutWait(void *buffer, unsigned long dwBytesRead)
{
	if( !m_bOpened || m_hComDev == NULL ) return 0;

	BOOL bReadStatus;
	unsigned long dwErrorFlags;
	COMSTAT ComStat;  

	ClearCommError( m_hComDev, &dwErrorFlags, &ComStat );
	if( !ComStat.cbInQue ) return 0;

	dwBytesRead = min(dwBytesRead, ComStat.cbInQue);

	bReadStatus = ReadFile( m_hComDev, buffer, dwBytesRead, &dwBytesRead, &m_OverlappedRead );
	if( !bReadStatus ){
		if( GetLastError() == ERROR_IO_PENDING ){
			WaitForSingleObject( m_OverlappedRead.hEvent, 2000 );
			return dwBytesRead;
		}
		return 0;
	}
	return dwBytesRead;
}

//=========================
//-----math function-----
//=========================
BYTE LRC(char * DataToLRC, UINT DataLength)
{
	BYTE temp_LRC = 0;
	while(DataLength--)
	{
		temp_LRC = temp_LRC ^ DataToLRC[DataLength];
	}
	return temp_LRC;
}

//ie.[11],[22]->"11 22"
CString Bytes2CString(BYTE *DataIN, UINT Len)
{
	CString OUT_CString = "";
	CString str_temp;
	UINT i;
	for(i=0;i<Len;i++)
	{
		str_temp.Format("%02X ",DataIN[i]);//只16进制
		OUT_CString+=str_temp; //加入接收编辑框对应字符串 
	}
	OUT_CString.TrimRight(" ");//去掉右边空格
	return OUT_CString;
}

CString Bytes2CString_ASCII(BYTE *DataIN, UINT Len)
{
	CString OUT_CString = "";
	CString str_temp;
	UINT i;
	for(i=0;i<Len;i++)
	{
		str_temp.Format("%c",DataIN[i]);//只16进制
		OUT_CString+=str_temp; //加入接收编辑框对应字符串 
	}
	OUT_CString.TrimRight(" ");//去掉右边空格
	return OUT_CString;
}

//将1字节字符转为1字节数.ie."9"->9;"a"->0x10
char str2char(char para_str)
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

//ie."00 99 88"->[0x00],[0x99],[0x88]
UINT CString2Bytes (BYTE *DataOUT, CString DataIN)
{
	if(DataIN == "")
	{
		return 0;
	}
	CByteArray hexdata;
	UINT i = 0;
	UINT DataOUT_len = 0;			
	char DataOUT_HL = 1;//1:转换1位，2:转换2位	
	//========去掉左,右边空格========
	CString temp_m_str = DataIN;
	temp_m_str.TrimLeft(" ");
	temp_m_str.TrimRight(" ");
	char* temp_CMD = (LPSTR)(LPCTSTR) temp_m_str;
	//====temp_CMD[] —> DataOUT[], 长度:[0]-[DataOUT_len]====
	
	//建议在此加入合并空格
	
	i=0;
	while(1)
	{
		if(temp_CMD[i] == 0X20)//空格
		{
			DataOUT_len ++;
			DataOUT_HL = 1;
			i++;
			continue;
		}
		else if (temp_CMD[i] != 0x00)
		{
			if(DataOUT_HL == 1)
			{
				DataOUT_HL = 2;
				DataOUT[DataOUT_len] = str2char(temp_CMD[i]);
			}
			else if(DataOUT_HL == 2)
			{
				DataOUT[DataOUT_len] = DataOUT[DataOUT_len]*16 + str2char(temp_CMD[i]);
			}
		}
		else if (temp_CMD[i] == 0x00)
		{
			break;
		}
		i++;
	}
	DataOUT_len++;
	return DataOUT_len;
}

//=====================================================================
//-----9525 CMD----------
//===================================================================== 
BOOL CMD_PC_to_RDR_IccPowerOn(CSerial *m_ctrlCSerial, 
							  BYTE SlotIn, BYTE PowerSelect,
							  BYTE *StatusOut, BYTE *ErrorOut, BYTE *ChainParameterOut,
							  BYTE *ATRbuffer, BYTE *ATRbufferLen)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_IccPowerOn_Resend:
	//=============================Send==========================
	char SendData[10];
	SendData[0] = PC_to_RDR_IccPowerOn;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = PowerSelect;
	SendData[8] = 0;
	SendData[9] = 0;
	SendData[10] = LRC(SendData,10);
	m_ctrlCSerial->SendData(SendData,11);
	//=============================Receive=========================
//	Sleep(1000);
	unsigned char GetData[50];
	UINT len;
	len = m_ctrlCSerial->ReadData(GetData,50);
	UINT i;
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOn_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != RDR_to_PC_DataBlock)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOn_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOn_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOn_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOn_Resend;
		}
		return FALSE;
	}
	//=======================Output========================
	for(i=0;i<(len - 11);i++)
	{
		ATRbuffer[i] = GetData[10+i];
	}
	*ATRbufferLen = i;
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	*ChainParameterOut = GetData[9];
	return TRUE;
}

BOOL CMD_PC_to_RDR_IccPowerOff(CSerial *m_ctrlCSerial, 
							  BYTE SlotIn,
							  BYTE *StatusOut, BYTE *ErrorOut, BYTE *ClockStatus)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_IccPowerOff_Resend:
	//==================Send=======================
	char SendData[10];
	SendData[0] = PC_to_RDR_IccPowerOff;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = 0;
	SendData[8] = 0;
	SendData[9] = 0;
	SendData[10] = LRC(SendData,10);
	m_ctrlCSerial->SendData(SendData,11);
	//====================Receive====================
//	Sleep(1000);
	unsigned char GetData[50];
	UINT len;
	len = m_ctrlCSerial->ReadData(GetData,50);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOff_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != RDR_to_PC_SlotStatus)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOff_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOff_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOff_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccPowerOff_Resend;
		}
		return FALSE;
	}
    //==================Output=====================
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	*ClockStatus = GetData[9];
	return 1;
}

BOOL CMD_PC_to_RDR_GetSlotStatus(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE *StatusOut, BYTE *ErrorOut, BYTE *ClockStatus)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_GetSlotStatus_Resend:
	//=============================Send==========================
	char SendData[10];
	SendData[0] = PC_to_RDR_GetSlotStatus;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = 0;
	SendData[8] = 0;
	SendData[9] = 0;
	SendData[10] = LRC(SendData,10);
	m_ctrlCSerial->SendData(SendData,11);
	//====================Receive====================
//	Sleep(1000);
	unsigned char GetData[50];
	UINT len;
	len = m_ctrlCSerial->ReadData(GetData,50);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetSlotStatus_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != RDR_to_PC_SlotStatus)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetSlotStatus_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetSlotStatus_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetSlotStatus_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetSlotStatus_Resend;
		}
		return FALSE;
	}
    //==================Output=====================
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	*ClockStatus = GetData[9];
	return 1;
}

BOOL CMD_PC_to_RDR_XfrBlock(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE BWIIn, UINT LevelParameter,
								BYTE *abDataIn, UINT abDataInLen,
							    BYTE *StatusOut, BYTE *ErrorOut, 
								BYTE *ChainParameterOut,
							    BYTE *abDataOut, UINT *abDataOutLen)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_XfrBlock_Resend:
	//Max number of abDataLen in 9525 is 262
	//=============================Send==========================
	char SendData[400];
	SendData[0] = PC_to_RDR_XfrBlock;
	SendData[1] = abDataInLen%256;
	SendData[2] = abDataInLen/256;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = BWIIn;
	SendData[8] = LevelParameter%256;
	SendData[9] = LevelParameter/256;
	UINT i;
	for(i=0;i<abDataInLen;i++)
	{
		SendData[10+i] = abDataIn[i];
	}
	SendData[10+abDataInLen] = LRC(SendData,(10+abDataInLen));
	m_ctrlCSerial->SendData(SendData,(11+abDataInLen));
	//====================Receive====================
//	Sleep(1000);
	unsigned char GetData[300];
	UINT len;
	len = m_ctrlCSerial->ReadData(GetData,300);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_XfrBlock_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != RDR_to_PC_DataBlock)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_XfrBlock_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_XfrBlock_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_XfrBlock_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_XfrBlock_Resend;
		}
		return FALSE;
	}
	//=======================Output========================
	for(i=0;i<(len - 11);i++)
	{
		abDataOut[i] = GetData[10+i];
	}
	*abDataOutLen = i;
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	*ChainParameterOut = GetData[9];

	return 1;
}

BOOL CMD_PC_to_RDR_GetParameters(CSerial *m_ctrlCSerial,
								BYTE SlotIn,
								BYTE *StatusOut, BYTE *ErrorOut,
							    BYTE *abDataOut, ULONG *abDataOutLen)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_GetParameters_Resend:
	//Max number of abDataLen in 9525 is 262
	//=============================Send==========================
	char SendData[400];
	SendData[0] = PC_to_RDR_GetParameters;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = 0;
	SendData[8] = 0;
	SendData[9] = 0;
	SendData[10] = LRC(SendData,10);
	ULONG i;
	m_ctrlCSerial->SendData(SendData,11);	
	//====================Receive====================
//	Sleep(1000);
	unsigned char GetData[300];
	ULONG len;
	len = m_ctrlCSerial->ReadData(GetData,300);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetParameters_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != RDR_to_PC_Parameters)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetParameters_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetParameters_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetParameters_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_GetParameters_Resend;
		}
		return FALSE;
	}
	//=======================Output========================
	for(i=0;i<(len - 11);i++)
	{
		abDataOut[i] = GetData[10+i];
	}
	*abDataOutLen = i;
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	return TRUE;	
}

BOOL CMD_PC_to_RDR_SetParameters(CSerial *m_ctrlCSerial,
								BYTE SlotIn,
								BYTE bProtocolNumIn,
								BYTE *abDataIn, ULONG abDataInLen,
							    BYTE *StatusOut, BYTE *ErrorOut,
							    BYTE *abDataOut, ULONG *abDataOutLen)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_SetParameters_Resend:
	//Max number of abDataLen in 9525 is 262
	//=============================Send==========================
	char SendData[400];
	SendData[0] = PC_to_RDR_SetParameters;
	SendData[1] = ((UINT)abDataInLen)%256;
	SendData[2] = ((UINT)abDataInLen)/256;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = bProtocolNumIn;
	SendData[8] = 0;
	SendData[9] = 0;
	ULONG i;
	for(i=0;i<abDataInLen;i++)
	{
		SendData[10+i] = abDataIn[i];
	}
	SendData[10+abDataInLen] = LRC(SendData,(10+abDataInLen));
	m_ctrlCSerial->SendData(SendData,(11+abDataInLen));	
	//====================Receive====================
//	Sleep(1000);
	unsigned char GetData[300];
	ULONG len;
	len = m_ctrlCSerial->ReadData(GetData,300);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_SetParameters_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != RDR_to_PC_Parameters)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_SetParameters_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_SetParameters_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_SetParameters_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_SetParameters_Resend;
		}
		return FALSE;
	}
	//=======================Output========================
	for(i=0;i<(len - 11);i++)
	{
		abDataOut[i] = GetData[10+i];
	}
	*abDataOutLen = i;
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	return TRUE;
}

BOOL CMD_PC_to_RDR_Escape(CSerial *m_ctrlCSerial,
								BYTE SlotIn,
								BYTE *abDataIn, ULONG abDataInLen,
							    BYTE *StatusOut, BYTE *ErrorOut,
							    BYTE *abDataOut, ULONG *abDataOutLen)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_Escape_Resend:
	//Max number of abDataLen in 9525 is 262
	//=============================Send==========================
	char SendData[400];
	SendData[0] = PC_to_RDR_Escape;
	SendData[1] = ((UINT)abDataInLen)%256;
	SendData[2] = ((UINT)abDataInLen)/256;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = 0;
	SendData[8] = 0;
	SendData[9] = 0;
	ULONG i;
	for(i=0;i<abDataInLen;i++)
	{
		SendData[10+i] = abDataIn[i];
	}
	SendData[10+abDataInLen] = LRC(SendData,(10+abDataInLen));
	m_ctrlCSerial->SendData(SendData,(11+abDataInLen));	
	//====================Receive====================
	unsigned char GetData[300];
	ULONG len;
	len = m_ctrlCSerial->ReadData(GetData,300);
	//debug
//	m_ctrlCSerial->SendData(SendData,(11+abDataInLen));	
//	while(1)
//	{
//	;
//	}
	//debug

	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_Escape_Resend;
		}
		return FALSE;
	}

	//CHECK message
	if(GetData[0] != RDR_to_PC_Escape)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_Escape_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_Escape_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_Escape_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_Escape_Resend;
		}
		return FALSE;
	}
	//CHECK Status
	if((GetData[7]&0xf0)!=0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_Escape_Resend;
		}
		return FALSE;
	}
	//=======================Output========================
	for(i=0;i<(len - 11);i++)
	{
		abDataOut[i] = GetData[10+i];
	}
	*abDataOutLen = i;
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	return TRUE;
}
BOOL CMD_PC_to_RDR_IccClock(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE ClockCommand,
							    BYTE *StatusOut, BYTE *ErrorOut, 
							    BYTE *ClockStatus)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_IccClock_Resend:
	//=============================Send==========================
	char SendData[10];
	SendData[0] = PC_to_RDR_IccClock;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = ClockCommand;
	SendData[8] = 0;
	SendData[9] = 0;
	SendData[10] = LRC(SendData,10);
	m_ctrlCSerial->SendData(SendData,11);
	//====================Receive====================
//	Sleep(1000);
	unsigned char GetData[50];
	UINT len;
	len = m_ctrlCSerial->ReadData(GetData,50);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccClock_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != RDR_to_PC_SlotStatus)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccClock_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccClock_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccClock_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_IccClock_Resend;
		}
		return FALSE;
	}
	//=======================Output========================
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	*ClockStatus = GetData[9];
	return 1;
}

BOOL CMD_PC_to_RDR_T0APDU(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE bmChangesIn, BYTE bClassGetResponseIn,
								BYTE bClassEnvelopeIn,
							    BYTE *StatusOut, BYTE *ErrorOut, 
							    BYTE *ClockStatus)
{ 
	BYTE ResendTimes = RS232ResendTimes;
CMD_PC_to_RDR_T0APDU_Resend:
	//=============================Send==========================
	char SendData[10];
	SendData[0] = PC_to_RDR_IccClock;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = SlotIn;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = bmChangesIn;
	SendData[8] = bClassGetResponseIn;
	SendData[9] = bClassEnvelopeIn;
	SendData[10] = LRC(SendData,10);
	m_ctrlCSerial->SendData(SendData,11);
	//====================Receive====================
//	Sleep(1000);
	unsigned char GetData[50];
	UINT len;
	len = m_ctrlCSerial->ReadData(GetData,50);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_T0APDU_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != RDR_to_PC_SlotStatus)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_T0APDU_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_T0APDU_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_T0APDU_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != SlotIn) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_PC_to_RDR_T0APDU_Resend;
		}
		return FALSE;
	}
	//=======================Output========================
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	*ClockStatus = GetData[9];
	return 1;
}
BOOL Check_RDR_to_PC_NotifySlotChange(CSerial *m_ctrlCSerial, BYTE *bmSlotICCStateOut)
{
	//====================Receive====================
	unsigned char GetData[50];
	UINT len;
	len = m_ctrlCSerial->ReadDataWithoutWait(GetData,50);
	if(len == 0)
	{
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != 0X85)
	{
		goto Give_Error_info;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		goto Give_Error_info;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		goto Give_Error_info;
	}
	*bmSlotICCStateOut = GetData[11];
	//================Send===================
	char SendData[10];
	SendData[0] = 0x64;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = 0;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = 0;
	SendData[8] = 0;
	SendData[9] = 0;
	SendData[10] = LRC(SendData,10);
	m_ctrlCSerial->SendData(SendData,11);	
	//====================Receive====================
	len = m_ctrlCSerial->ReadData(GetData,50);
	if(len == 0)
	{
		return FALSE;
	}
	return TRUE;

Give_Error_info:
	SendData[0] = 0;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = 0;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = 0;
	SendData[8] = 0;
	SendData[9] = 0;
	SendData[10] = 0;
	m_ctrlCSerial->SendData(SendData,11);	
	return FALSE;

}

//=====================================================================
//-----9525 Vendor CMD----------
//===================================================================== 
LONG CMD_SetBaudRate(CSerial *m_ctrlCSerial,  BYTE BaudRate)
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_SetBaudRate_Resend:
	char SendData[10];
	SendData[0] = SetBaudRate;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = 0;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = BaudRate;
	SendData[8] = 0;
	SendData[9] = 0;
	SendData[10] = LRC(SendData,10);
	m_ctrlCSerial->SendData(SendData,11);	
	//====================Receive====================
//	Sleep(500);
	unsigned char GetData[300];
	ULONG len;
	len = m_ctrlCSerial->ReadData(GetData,300);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_SetBaudRate_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != Response_SetBaudRate)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_SetBaudRate_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_SetBaudRate_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_SetBaudRate_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != 0) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_SetBaudRate_Resend;
		}
		return FALSE;
	}

	return TRUE;
}

//Buffer size: 0x23, 0x01->0x123; 
//extended or short APDU: Always 1(extended)
//len of manufacture string
//manufacture string
//len of product string
//manufacture string
LONG CMD_GetReaderDescriptor(CSerial *m_ctrlCSerial,
							 BYTE bDevDescIn,
							 BYTE bStrDescIn,
							 BYTE *StatusOut, BYTE *ErrorOut,
							 BYTE *abDataOut, ULONG *abDataOutLen
							 )	   
{
	BYTE ResendTimes = RS232ResendTimes;
CMD_GetReaderDescriptor_Resend:
	char SendData[10];
	SendData[0] = GetReaderDescriptor;
	SendData[1] = 0;
	SendData[2] = 0;
	SendData[3] = 0;
	SendData[4] = 0;
	SendData[5] = 0;
	SendData[6] = m_ctrlCSerial->SendDataSequenceNum++;
	SendData[7] = bDevDescIn;
	SendData[8] = bStrDescIn;
	SendData[9] = 0;
	SendData[10] = LRC(SendData,10);
	ULONG i;
	m_ctrlCSerial->SendData(SendData,11);	
	//====================Receive====================
	unsigned char GetData[300];
	ULONG len;
	len = m_ctrlCSerial->ReadData(GetData,300);
	if(len == 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_GetReaderDescriptor_Resend;
		}
		return FALSE;
	}
	//CHECK message
	if(GetData[0] != Response_GetReaderDescriptor)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_GetReaderDescriptor_Resend;
		}
		return FALSE;
	}
	//CHECK Data length
	UINT Data_Length;
	Data_Length = GetData[1]+GetData[2]*256;
	if(Data_Length != (len - 11))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_GetReaderDescriptor_Resend;
		}
		return FALSE;
	}
	//CHECK LRC
	if(LRC((char *)GetData,len) != 0)
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_GetReaderDescriptor_Resend;
		}
		return FALSE;
	}
	//CHECK Slot and Sequense
	unsigned char Sequense_Num;
	Sequense_Num = GetData[6]+1;
	if((GetData[5] != 0) || (Sequense_Num != m_ctrlCSerial->SendDataSequenceNum))
	{
		if((ResendTimes) != 0)
		{
			ResendTimes -- ;
			goto CMD_GetReaderDescriptor_Resend;
		}
		return FALSE;
	}
	//=======================Output========================
	for(i=0;i<(len - 11);i++)
	{
		abDataOut[i] = GetData[10+i];
	}
	*abDataOutLen = i;
	*StatusOut = GetData[7];
	*ErrorOut = GetData[8];
	return TRUE;
}

//-----------LCM,Keypad,EE2prom-----------------------------------
LONG EepromCmdWrite(
		CSerial *m_ctrlCSerial, 
		UCHAR	bSlotNum,
		ULONG	lngStartAddr,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 128
	if( lngWriteLen > 128 )
	{
		return 0xFF;
	}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=EEPROM_WRITE | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngStartAddr & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngStartAddr>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)= (UCHAR)lngWriteLen;
	*((PUCHAR)pSendBuffer+7)=0x00;
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG EepromCmdRead(
		CSerial *m_ctrlCSerial, 
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngStartAddr,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	UCHAR	bSlotOffset;
	UCHAR	bCmdHdrSize;
	ULONG	lngReturnLen;

	// Read maximun length can not exceed 200
	if( lngReadLen > 256 )
	{
		return 0xFF;
	}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize=8;
	pSendBuffer	 = malloc(bCmdHdrSize);

	// send read command
	dwSendLength=bCmdHdrSize;

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=EEPROM_READ | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngStartAddr & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngStartAddr>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=(UCHAR)(lngReadLen & 0xFF);
	*((PUCHAR)pSendBuffer+7)=0x00;
	
	BYTE StatusOut;
	BYTE ErrorOut;
    //send command to get data from EEPROM card
	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);	
	free((LPVOID)pSendBuffer);
	return lngStatus;		
}

LONG CmdFirmwareWrite(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bIndex,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 256 //128
	if( lngWriteLen > 256 )//128 )
	{
		return 0xFF;
	}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=FIRMWARE_WRITE | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=bIndex;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);

	free((LPVOID)pSendBuffer);

	return lngStatus;	
}

LONG APIENTRY CmdFirmwareRead(
		CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bIndex,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	UCHAR	bSlotOffset;
	UCHAR	bCmdHdrSize;
	ULONG	lngReturnLen;

	// Read maximun length can not exceed 200
	if( lngReadLen > 256 )
	{
		return 0xFF;
	}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize=8;
	pSendBuffer	 = malloc(bCmdHdrSize);

	// send read command
	dwSendLength=bCmdHdrSize;

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=FIRMWARE_READ | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=bIndex;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
    //send command to get data from EEPROM card
	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);	

	*plngReturnLen = min(lngReadLen, *plngReturnLen);
	free((LPVOID)pSendBuffer);
	return lngStatus;		
}

LONG CmdLcmSetCursor(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bXL,
		UCHAR	bXH,
		UCHAR	bTP
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=LCD_SET_CURSOR | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=bXL;
	*((PUCHAR)pSendBuffer+3)=bXH;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=bTP;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}

//not use, not declared in 9525RS232Lib.cpp
LONG CmdLcmClearDisplay(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bXL,
		UCHAR	bXH,
		UCHAR	bTP
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=LCD_CLEAR_DISPLAY | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=bXL;
	*((PUCHAR)pSendBuffer+3)=bXH;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=bTP;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);
	return lngStatus;	
}

//not use, not declared in 9525RS232Lib.cpp
LONG CmdLcmDisplayMessage(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bXL,
		UCHAR	bXH,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 256 //128
	if( lngWriteLen > 256 )//128 )
	{
		return 0xFF;
	}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=LCD_DISPLAY_MESSAGE | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=bXL;
	*((PUCHAR)pSendBuffer+3)=bXH;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}

//not use, not declared in 9525RS232Lib.cpp
LONG CmdLcmDisplayGraphic(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 256 //128
	if( lngWriteLen > 256 )//128 )
	{
		return 0xFF;
	}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=LCD_DISPLAY_GRAPHIC | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=0x00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}

LONG CmdLcmWriteData(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 256 //128
	if( lngWriteLen > 256 )//128 )
	{
		return 0xFF;
	}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=LCD_WRITE_DATA | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)lngWriteLen;
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngWriteLen>>0x08);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}

LONG CmdTestFwUpDate(
		CSerial *m_ctrlCSerial
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=LCD_SET_CONTRAST;
	*((PUCHAR)pSendBuffer+2)=0x00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									0,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}


LONG CmdLcmSetBacklight(
		CSerial *m_ctrlCSerial
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=LCD_SET_BACKLIGHT;
	*((PUCHAR)pSendBuffer+2)=0x00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									0,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}

LONG CmdClearKeyBuffer(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=CLEAR_KEY_BUFFER | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=0x00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}

/*LONG CmdSetKeyScanTimer(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bXL,
		UCHAR	bXH,
		UCHAR	bTimerFlag
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=SET_KEY_SCAN_TIMER | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=bXL;
	*((PUCHAR)pSendBuffer+3)=bXH;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=bTimerFlag;
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}*/
LONG CmdSetKeyScanTimer(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bXTimes,
//		UCHAR	bXH,
		UCHAR	bTimerFlag
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=SET_KEY_SCAN_TIMER | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=bXTimes;
	*((PUCHAR)pSendBuffer+3)=bTimerFlag;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);
	free((LPVOID)pSendBuffer);

	return lngStatus;	
}

LONG APIENTRY CmdGetKeyInput(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	UCHAR	bSlotOffset;
	UCHAR	bCmdHdrSize;
	ULONG	lngReturnLen;

	// Read maximun length can not exceed 200
	if( lngReadLen > 256 )
	{
		return 0xFF;
	}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize=8;
	pSendBuffer	 = malloc(bCmdHdrSize);

	// send read command
	dwSendLength=bCmdHdrSize;

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE_LCM_KEYPAD;
	*((PUCHAR)pSendBuffer+1)=GET_KEY_INPUT | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=0x00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
    //send command to get data from EEPROM card
	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);	

	*plngReturnLen = min(lngReadLen, *plngReturnLen);

	free((LPVOID)pSendBuffer);
	return lngStatus;		
}
//no use
LONG CheckBusy(
		CSerial *m_ctrlCSerial,
		UCHAR bBusyValue
		)
{
	UCHAR aCheckBusy[] = 
	{
		LCDRSReg, LOW, LCDRWReg, HIGH, LCDDataReg, RESET, LCDEnReg, HIGH, CMDBusy, LCDDataReg,	0x00, 0x01,
		LCDEnReg, LOW
	};

	aCheckBusy[BUSY_POS] = bBusyValue;
	return CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(aCheckBusy), (unsigned char *)aCheckBusy);
}

LONG WriteCommandNocheck(
		CSerial *m_ctrlCSerial,
		UCHAR bCmd
		)
{
	UCHAR aWriteCommandNocheck[] = 
	{
		LCDRSReg, LOW, LCDRWReg, LOW, LCDDataReg, CMD_CODE_DEFAULT, LCDEnReg, HIGH, LCDEnReg, LOW,
	};

	aWriteCommandNocheck[CMD_CODE_POS] = bCmd;
	return CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(aWriteCommandNocheck), (unsigned char *)aWriteCommandNocheck);
}
 
LONG WriteData(
		CSerial *m_ctrlCSerial,
		UCHAR	bData
		)
{
	UCHAR aWriteData[] = 
	{
		LCDRSReg, HIGH, LCDRWReg, LOW, LCDDataReg, WRITE_DATA_DEF, LCDEnReg, HIGH, LCDEnReg, LOW
	};

	aWriteData[WRITE_DATA_POS] = bData;

	return CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(aWriteData), (unsigned char *)aWriteData);
}

LONG SelectScreen(
		CSerial *m_ctrlCSerial,
		UCHAR	bSelectScreen
		)
{
	UCHAR aSelectScreen[] = {LCDCSReg,0x00};

	aSelectScreen[0x01] = bSelectScreen;

	return CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(aSelectScreen), (unsigned char *)aSelectScreen);
}

LONG SetBackLight(
		CSerial *m_ctrlCSerial,
		UCHAR	bSelectBacklight
		)
{
	UCHAR aSetBackLight[] = {LCDBLReg,0x00};
	aSetBackLight[0x01] = bSelectBacklight;

	return CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(aSetBackLight), (unsigned char *)aSetBackLight);
}
//--------------------for synchronous card----------------------------
LONG CMD_SWITCH_CARD_MODE(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  BYTE Card_Mode_Switch
						  )
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_SWITCH_CARD_MODE | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=Card_Mode_Switch;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_POWER_ON(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_POWER_ON | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=0x00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;	
}
LONG CMD_POWER_OFF(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_POWER_OFF | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=0x00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;	
}

LONG CMD_SET_I2C_ADD(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  BYTE DeviceAddress,
						  UINT WordAddr,
						  BYTE PageSize)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_SET_I2C_ADD | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=DeviceAddress|(WordAddr/256);
	*((PUCHAR)pSendBuffer+3)=WordAddr%256;
	*((PUCHAR)pSendBuffer+4)=PageSize;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;	
}

LONG CMD_WRITE_I2C(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  ULONG	lngWriteLen,
						  UCHAR	*pWriteData)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 128
	//if( lngWriteLen > 128 )
	//{
	//	return 0xFF;
	//}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}
	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_WRITE_I2C | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngWriteLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngWriteLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0X00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_READ_I2C(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  IN	ULONG	lngReadLen,
						  OUT	LPVOID	pReadData,
						  OUT	ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	UCHAR	bSlotOffset;
	UCHAR	bCmdHdrSize;
	ULONG	lngReturnLen;
	// Read maximun length can not exceed 200
	//if( lngReadLen > 256 )
	//{
	//	return 0xFF;
	//}
	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize=8;
	pSendBuffer	 = malloc(bCmdHdrSize);

	// send read command
	dwSendLength=bCmdHdrSize;

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_READ_I2C | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngReadLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngReadLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0X00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	
	BYTE StatusOut;
	BYTE ErrorOut;
    //send command to get data from EEPROM card
	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);	
	free((LPVOID)pSendBuffer);
	return lngStatus;	
}

LONG CMD_AT45D041_CARD_COMMAND(CSerial *m_ctrlCSerial,
								BYTE bSlotNum,
								ULONG	lngWriteLen,
								UCHAR	*pWriteData,
								IN	ULONG	lngReadLen,
						        OUT	LPVOID	pReadData,
								OUT	ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 128
	//if( lngWriteLen > 128 )
	//{
	//	return 0xFF;
	//}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}
	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_AT45D041_CARD_COMMAND | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngReadLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngReadLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=(UCHAR)(lngWriteLen & 0xFF);
	*((PUCHAR)pSendBuffer+7)=(UCHAR)(lngWriteLen>>8);
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	//BYTE pReadData[400];
	//ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_SMC_COMMAND(CSerial *m_ctrlCSerial,
								BYTE bSlotNum,
								ULONG	lngWriteLen,
								UCHAR	*pWriteData,
								IN	ULONG	lngReadLen,
						        OUT	LPVOID	pReadData,
								OUT	ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 128
	//if( lngWriteLen > 128 )
	//{
	//	return 0xFF;
	//}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}
	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_SMC_COMMAND | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngReadLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngReadLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=(UCHAR)(lngWriteLen & 0xFF);
	*((PUCHAR)pSendBuffer+7)=(UCHAR)(lngWriteLen>>8);
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	//BYTE pReadData[400];
	//ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_SLE4442_CARD_COMMAND(CSerial *m_ctrlCSerial,
							  BYTE bSlotNum,
						IN	  UINT ReadingLength,
							  BYTE ProtectBit,
							  BYTE AskForClkNum,
							  BYTE CMD_1,
							  BYTE CMD_2,
							  BYTE CMD_3,
						OUT	  LPVOID pReadData,
						OUT	  ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}
	bCmdHdrSize	 = 8;
	dwSendLength = 11;
	pSendBuffer	 = malloc( 11 );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_SLE4442_CARD_COMMAND | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=ReadingLength%256;
	*((PUCHAR)pSendBuffer+3)=ReadingLength/256;
	*((PUCHAR)pSendBuffer+4)=ProtectBit;
	*((PUCHAR)pSendBuffer+5)=AskForClkNum;
	*((PUCHAR)pSendBuffer+6)=0x03;
	*((PUCHAR)pSendBuffer+7)=0x00;

	*((PUCHAR)pSendBuffer+8)=CMD_1;
	*((PUCHAR)pSendBuffer+9)=CMD_2;
	*((PUCHAR)pSendBuffer+10)=CMD_3;
	BYTE StatusOut;
	BYTE ErrorOut;
//  LONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}


LONG CMD_SLE4442_CARD_BREAK(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_SLE4442_CARD_BREAK | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=0X00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_SLE4428_CARD_COMMAND(CSerial *m_ctrlCSerial,
							  BYTE bSlotNum,
						IN	  UINT ReadingLength,
							  BYTE ProtectBit,
							  BYTE AskForClkNum,
							  BYTE CMD_1,
							  BYTE CMD_2,
							  BYTE CMD_3,
						OUT	  LPVOID pReadData,
						OUT	  ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}
	bCmdHdrSize	 = 8;
	dwSendLength = 11;
	pSendBuffer	 = malloc( 11 );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_SLE4428_CARD_COMMAND | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=ReadingLength%256;
	*((PUCHAR)pSendBuffer+3)=ReadingLength/256;
	*((PUCHAR)pSendBuffer+4)=ProtectBit;
	*((PUCHAR)pSendBuffer+5)=AskForClkNum;
	*((PUCHAR)pSendBuffer+6)=0x03;
	*((PUCHAR)pSendBuffer+7)=0x00;

	*((PUCHAR)pSendBuffer+8)=CMD_1;
	*((PUCHAR)pSendBuffer+9)=CMD_2;
	*((PUCHAR)pSendBuffer+10)=CMD_3;
	BYTE StatusOut;
	BYTE ErrorOut;
//  LONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_INPHONE_CARD_RESET(CSerial *m_ctrlCSerial,
						    BYTE bSlotNum)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_INPHONE_CARD_RESET | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=0X00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}
LONG CMD_INPHONE_CARD_Read(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  IN	ULONG	lngReadLen,
						  OUT	LPVOID	pReadData,
						  OUT	ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	UCHAR	bSlotOffset;
	UCHAR	bCmdHdrSize;
	ULONG	lngReturnLen;
	// Read maximun length can not exceed 200
	//if( lngReadLen > 256 )
	//{
	//	return 0xFF;
	//}
	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize=8;
	pSendBuffer	 = malloc(bCmdHdrSize);

	// send read command
	dwSendLength=bCmdHdrSize;

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_INPHONE_CARD_Read | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngReadLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngReadLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0X00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	
	BYTE StatusOut;
	BYTE ErrorOut;
    //send command to get data from EEPROM card
	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);	
	free((LPVOID)pSendBuffer);
	return lngStatus;	
}

/*
LONG CMD_INPHONE_CARD_PROG(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  ULONG	lngWriteLen,
						  UCHAR	*pWriteData)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 128
	//if( lngWriteLen > 128 )
	//{
	//	return 0xFF;
	//}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}
	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_INPHONE_CARD_PROG | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngWriteLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngWriteLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0X00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}*/
LONG CMD_INPHONE_CARD_PROG(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  IN	ULONG	lngWriteLen,
						  OUT	LPVOID	pReadData,
						  OUT	ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	UCHAR	bSlotOffset;
	UCHAR	bCmdHdrSize;
	ULONG	lngReturnLen;
	// Read maximun length can not exceed 200
	//if( lngReadLen > 256 )
	//{
	//	return 0xFF;
	//}
	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize=8;
	pSendBuffer	 = malloc(bCmdHdrSize);

	// send read command
	dwSendLength=bCmdHdrSize;

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_INPHONE_CARD_PROG | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngWriteLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngWriteLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0X00;
	*((PUCHAR)pSendBuffer+7)=0x00;
	
	BYTE StatusOut;
	BYTE ErrorOut;
    //send command to get data from EEPROM card
	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);	
	free((LPVOID)pSendBuffer);
	return lngStatus;	
}


LONG CMD_INPHONE_CARD_MOVE_ADDRESS(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  UINT ClockNums)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_INPHONE_CARD_MOVE_ADDRESS | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=0x00;
	*((PUCHAR)pSendBuffer+3)=0x00;
	*((PUCHAR)pSendBuffer+4)=(UCHAR)(ClockNums & 0xFF);
	*((PUCHAR)pSendBuffer+5)=(UCHAR)(ClockNums>>8);
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;	
}

LONG CMD_INPHONE_CARD_AUTHENTICATION_KEY1(CSerial *m_ctrlCSerial,
										BYTE bSlotNum,
										ULONG	lngWriteLen,
										UCHAR	*pWriteData,
										IN	ULONG	lngReadLen,
										OUT	LPVOID	pReadData,
										OUT	ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 128
	//if( lngWriteLen > 128 )
	//{
	//	return 0xFF;
	//}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}
	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_INPHONE_CARD_AUTHENTICATION_KEY1 | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngReadLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngReadLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=(UCHAR)(lngWriteLen & 0xFF);
	*((PUCHAR)pSendBuffer+7)=(UCHAR)(lngWriteLen>>8);
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	//BYTE pReadData[400];
	//ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_INPHONE_CARD_AUTHENTICATION_KEY2(CSerial *m_ctrlCSerial,
										BYTE bSlotNum,
										ULONG	lngWriteLen,
										UCHAR	*pWriteData,
										IN	ULONG	lngReadLen,
										OUT	LPVOID	pReadData,
										OUT	ULONG	*plngReturnLen)
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	UCHAR	bSlotOffset;

	// Write maximun length can not exceed 128
	//if( lngWriteLen > 128 )
	//{
	//	return 0xFF;
	//}

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}
	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = bCmdHdrSize + lngWriteLen;
	pSendBuffer	 = malloc( bCmdHdrSize + lngWriteLen );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_INPHONE_CARD_AUTHENTICATION_KEY2 | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=(UCHAR)(lngReadLen & 0xFF);
	*((PUCHAR)pSendBuffer+3)=(UCHAR)(lngReadLen>>8);
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=(UCHAR)(lngWriteLen & 0xFF);
	*((PUCHAR)pSendBuffer+7)=(UCHAR)(lngWriteLen>>8);
	memcpy( (PUCHAR)pSendBuffer + 8, pWriteData, lngWriteLen);

	BYTE StatusOut;
	BYTE ErrorOut;
	//BYTE pReadData[400];
	//ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									(PUCHAR)pReadData,
									plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_SET_CONFIG(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  BYTE xcheckOverCurrent
						  )
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_SET_CONFIG | bSlotOffset;
	switch(xcheckOverCurrent)
	{
		case 1://set xcheckOverCurrent
		{
			*((PUCHAR)pSendBuffer+2)=0x01;
			*((PUCHAR)pSendBuffer+3)=0x01;		
		}
		break;
		case 0://clear xcheckOverCurrent
		{
			*((PUCHAR)pSendBuffer+2)=0x00;
			*((PUCHAR)pSendBuffer+3)=0x01;		
		}
		break;
	}

	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}

LONG CMD_SET_LED(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  BYTE LED_Num,
						  BYTE LED_Switch
						  )
{
	LPCVOID	pSendBuffer;
	DWORD	dwSendLength;
	LONG	lngStatus;
	ULONG	lngReturnLen;
	UCHAR	bCmdHdrSize;
	ULONG	lngReadLen=0;
	UCHAR	bSlotOffset;

	if( bSlotNum == 0 )
	{
		bSlotOffset=0;
	}
	else
	{
		bSlotOffset=0x08;
	}

	lngReturnLen=0;

	bCmdHdrSize	 = 8;
	dwSendLength = 8;
	pSendBuffer	 = malloc( bCmdHdrSize );

	*((PUCHAR)pSendBuffer+0)=ALCOR_OP_CODE;
	*((PUCHAR)pSendBuffer+1)=OP_SET_LED | bSlotOffset;
	*((PUCHAR)pSendBuffer+2)=LED_Num;
	*((PUCHAR)pSendBuffer+3)=LED_Switch;		
	*((PUCHAR)pSendBuffer+4)=0x00;
	*((PUCHAR)pSendBuffer+5)=0x00;
	*((PUCHAR)pSendBuffer+6)=0x00;
	*((PUCHAR)pSendBuffer+7)=0x00;

	BYTE StatusOut;
	BYTE ErrorOut;
	BYTE pReadData[400];
	ULONG plngReturnLen;

	lngStatus = CMD_PC_to_RDR_Escape(m_ctrlCSerial,
									bSlotNum,
									(BYTE *)pSendBuffer,dwSendLength,
									&StatusOut,
									&ErrorOut,
									pReadData,
									&plngReturnLen);															
	free((LPVOID)pSendBuffer);
	return lngStatus;
}
//=============================================
//同步卡命令
//=============================================
//-------------4418/28 Card------------------
//suceess 0; fail 1.

//
//return value
//SLE4428_SUCCESS : No error.
//SLE4428_ERR_COMMAND_FAILED : PSC byte incorrect or other error.
//SLE4428_ERR_ZERO_ERR_CNT : The error counter is zero. 
//
LONG APIENTRY SLE4428Cmd_VerifyPSCAndEraseErrorCounter (
                               IN           CSerial *m_ctrlCSerial,
                               IN           UCHAR bSlotNum,
                               IN           UCHAR b1stPsc,
                               IN           UCHAR b2ndPsc,
                               OUT ULONG   *plngErrorReason)
{
return 0;
}


LONG SLE4428Cmd_WriteEraseWithPB(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	UCHAR	bData
		)
{
	LONG	lngStatus;
	BYTE  pReadData[400];
	ULONG plngReturnLen;


	BYTE TEMP_CMD1_b7tob6 = (BYTE)((lngAddress&0x0300)>>2);
	BYTE TEMP_CMD1_b5tob0 = 0x31;
	BYTE TEMP_CMD1 = TEMP_CMD1_b7tob6 | TEMP_CMD1_b5tob0;

	BYTE lngAddress_b7tob0 = (BYTE)(lngAddress&0x00FF);

	lngStatus = CMD_SLE4428_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										1,
										0,
										1,
										TEMP_CMD1,
										lngAddress_b7tob0,
										bData,
										pReadData,
										&plngReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if((pReadData[0] != 0x67) && (pReadData[0] != 0xcb))//clock number
	{
		return 1;
	}
	return 0;
}

LONG SLE4428Cmd_WriteEraseWithoutPB(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	UCHAR	bData
		)
{
	LONG	lngStatus;
	BYTE  pReadData[400];
	ULONG plngReturnLen;


	BYTE TEMP_CMD1_b7tob6 = (BYTE)((lngAddress&0x0300)>>2);
	BYTE TEMP_CMD1_b5tob0 = 0x33;
	BYTE TEMP_CMD1 = TEMP_CMD1_b7tob6 | TEMP_CMD1_b5tob0;

	BYTE lngAddress_b7tob0 = (BYTE)(lngAddress&0x00FF);

	lngStatus = CMD_SLE4428_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										1,
										0,
										1,
										TEMP_CMD1,
										lngAddress_b7tob0,
										bData,
										pReadData,
										&plngReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if((pReadData[0] != 0x67) && (pReadData[0] != 0xcb))//clock number
	{
		return 1;
	}
	return 0;
}
/*
LONG CMD_SLE4428_CARD_COMMAND(CSerial *m_ctrlCSerial,
							  BYTE bSlotNum,
						IN	  UINT ReadingLength,
							  BYTE ProtectBit,
							  BYTE AskForClkNum,
							  BYTE CMD_1,
							  BYTE CMD_2,
							  BYTE CMD_3,
						OUT	  LPVOID pReadData,
						OUT	  ULONG	*plngReturnLen)
*/
LONG SLE4428Cmd_WritePBWithDataComparison(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	UCHAR	bData
		)
{
	LONG	lngStatus;
	BYTE  pReadData[400];
	ULONG plngReturnLen;

	BYTE TEMP_CMD1_b7tob6 = (BYTE)((lngAddress&0x0300)>>2);
	BYTE TEMP_CMD1_b5tob0 = 0x30;
	BYTE TEMP_CMD1 = TEMP_CMD1_b7tob6 | TEMP_CMD1_b5tob0;

	BYTE lngAddress_b7tob0 = (BYTE)(lngAddress&0x00FF);

	lngStatus = CMD_SLE4428_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										1,
										0,
										1,
										TEMP_CMD1,
										lngAddress_b7tob0,
										bData,
										pReadData,
										&plngReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if((pReadData[0] != 0x67) && (pReadData[0] != 0xcb))//clock number
	{
		return 1;
	}
	return 0;
}

LONG SLE4428Cmd_Read9Bits(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	LPVOID	pReadPB,
		OUT	ULONG	*plngReturnLen
		)
{
	LONG	lngStatus;
	UINT i;
	BYTE ReceiveData[2060];

	BYTE TEMP_CMD1_b7tob6 = (BYTE)((lngAddress&0x0300)>>2);
	BYTE TEMP_CMD1_b5tob0 = 0x0C;
	BYTE TEMP_CMD1 = TEMP_CMD1_b7tob6 | TEMP_CMD1_b5tob0;

	BYTE lngAddress_b7tob0 = (BYTE)(lngAddress&0x00FF);

	lngStatus = CMD_SLE4428_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										(lngReadLen*2),
										1,
										0,
										TEMP_CMD1,
										lngAddress_b7tob0,
										0,
										ReceiveData,
										plngReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	for (i=0;i<(*plngReturnLen)/2;i++)
	{
		*((PCHAR)pReadData+i) = ReceiveData[i*2];
		*((PCHAR)pReadPB+i) = ReceiveData[(i*2+1)];
	}
	//pReadPB = (PCHAR)pReadData +1;
	return 0;
}

LONG SLE4428Cmd_Read8Bits(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		)
{
	LONG	lngStatus;

	BYTE TEMP_CMD1_b7tob6 = (BYTE)((lngAddress&0x0300)>>2);
	BYTE TEMP_CMD1_b5tob0 = 0x0e;
	BYTE TEMP_CMD1 = TEMP_CMD1_b7tob6 | TEMP_CMD1_b5tob0;

	BYTE lngAddress_b7tob0 = (BYTE)(lngAddress&0x00FF);

	lngStatus = CMD_SLE4428_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										lngReadLen,
										0,
										0,
										TEMP_CMD1,
										lngAddress_b7tob0,
										0,
										pReadData,
										plngReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	return 0;
}

LONG SLE4428Cmd_WriteErrorCounter(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bData
		)
{
	LONG	lngStatus;

	BYTE  pReadData[400];
	ULONG plngReturnLen;

	BYTE TEMP_CMD1_b7tob6 = (BYTE)((1021&0x0300)>>2);
	BYTE TEMP_CMD1_b5tob0 = 0x32;
	BYTE TEMP_CMD1 = TEMP_CMD1_b7tob6 | TEMP_CMD1_b5tob0;

	BYTE lngAddress_b7tob0 = (BYTE)(1021&0x00FF);
	lngStatus = CMD_SLE4428_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										1,
										0,
										1,
										TEMP_CMD1,
										lngAddress_b7tob0,
										bData,
										pReadData,
										&plngReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if(pReadData[0] != 0x67)
	{
		return 1;	
	}
	return 0;
}
 
LONG SLE4428Cmd_Verify1stPSC(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bData
		)
{
	LONG  lngStatus;

	BYTE  pReadData[400];
	ULONG plngReturnLen;

	BYTE TEMP_CMD1_b7tob6 = (BYTE)((1022&0x0300)>>2);
	BYTE TEMP_CMD1_b5tob0 = 0x0d;
	BYTE TEMP_CMD1 = TEMP_CMD1_b7tob6 | TEMP_CMD1_b5tob0;

	BYTE lngAddress_b7tob0 = (BYTE)(1022&0x00FF);
	lngStatus = CMD_SLE4428_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										1,
										0,
										1,
										TEMP_CMD1,
										lngAddress_b7tob0,
										bData,
										pReadData,
										&plngReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if(pReadData[0]!=0x02)
	{
		return 1;	
	}
	return 0;
}
 
LONG SLE4428Cmd_Verify2ndPSC(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bData
		)
{
	LONG  lngStatus;

	BYTE  pReadData[400];
	ULONG plngReturnLen;

	BYTE TEMP_CMD1_b7tob6 = (BYTE)((1023&0x0300)>>2);
	BYTE TEMP_CMD1_b5tob0 = 0x0d;
	BYTE TEMP_CMD1 = TEMP_CMD1_b7tob6 | TEMP_CMD1_b5tob0;

	BYTE lngAddress_b7tob0 = (BYTE)(1023&0x00FF);
	lngStatus = CMD_SLE4428_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										1,
										0,
										1,
										TEMP_CMD1,
										lngAddress_b7tob0,
										bData,
										pReadData,
										&plngReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if(pReadData[0]!=0x02)
	{
		return 1;	
	}
	return 0;	
}

//-------------4432/42 Card------------------
/*LONG CMD_SLE4442_CARD_COMMAND(CSerial *m_ctrlCSerial,
							  BYTE bSlotNum,
						IN	  UINT ReadingLength,
							  BYTE ProtectBit,
							  BYTE AskForClkNum,
							  BYTE CMD_1,
							  BYTE CMD_2,
							  BYTE CMD_3,
						OUT	  LPVOID pReadData,
						OUT	  ULONG	*plngReturnLen)
LONG CMD_SLE4442_CARD_BREAK(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum)*/
//suceess 0; fail 1.
LONG APIENTRY SLE4442Cmd_ReadMainMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*pbReturnLen
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SLE4442_CARD_BREAK(m_ctrlCSerial,bSlotNum);
	if(lngStatus!=1)
	{
		return 1;
	}						  
	BYTE TEMP_CMD1 = 0x30;
/*
LONG CMD_SLE4442_CARD_COMMAND(CSerial *m_ctrlCSerial,
							  BYTE bSlotNum,
						IN	  UINT ReadingLength,
							  BYTE ProtectBit,
							  BYTE AskForClkNum,
							  BYTE CMD_1,
							  BYTE CMD_2,
							  BYTE CMD_3,
						OUT	  LPVOID pReadData,
						OUT	  ULONG	*plngReturnLen)
*/
	lngStatus = CMD_SLE4442_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										bReadLen,
										0,
										0,
										TEMP_CMD1,
										bAddress,
										0,
										pReadData,
										pbReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	return 0;
}

LONG APIENTRY SLE4442Cmd_UpdateMainMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bData
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SLE4442_CARD_BREAK(m_ctrlCSerial,bSlotNum);
	if(lngStatus!=1)
	{
		return 1;
	}	
	BYTE TEMP_CMD1 = 0x38;
	BYTE  pReadData[400];
	ULONG pbReturnLen;
	lngStatus = CMD_SLE4442_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										0,
										0,
										0,
										TEMP_CMD1,
										bAddress,
										bData,
										pReadData,
										&pbReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if((pReadData[0] != 0x7b) && (pReadData[0] != 0xf4))//clock number
	{
		return 1;
	}
	return 0;
}
 
LONG APIENTRY SLE4442Cmd_ReadProtectionMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*pbReturnLen
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SLE4442_CARD_BREAK(m_ctrlCSerial,bSlotNum);
	if(lngStatus!=1)
	{
		return 1;
	}						  
	BYTE TEMP_CMD1 = 0x34;
	lngStatus = CMD_SLE4442_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										bReadLen,
										0,
										0,
										TEMP_CMD1,
										0,
										0,
										pReadData,
										pbReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	return 0;
}

LONG APIENTRY SLE4442Cmd_WriteProtectionMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bData
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SLE4442_CARD_BREAK(m_ctrlCSerial,bSlotNum);
	if(lngStatus!=1)
	{
		return 1;
	}	
	BYTE TEMP_CMD1 = 0x3C;
	BYTE  pReadData[400];
	ULONG pbReturnLen;
	lngStatus = CMD_SLE4442_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										0,
										0,
										0,
										TEMP_CMD1,
										bAddress,
										bData,
										pReadData,
										&pbReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if((pReadData[0] != 0x7b) && (pReadData[0] != 0xf4))//clock number
	{
		return 1;
	}
	return 0;	
} 

LONG APIENTRY SLE4442Cmd_ReadSecurityMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*pbReturnLen
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SLE4442_CARD_BREAK(m_ctrlCSerial,bSlotNum);
	if(lngStatus!=1)
	{
		return 1;
	}						  
	BYTE TEMP_CMD1 = 0x31;
	lngStatus = CMD_SLE4442_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										bReadLen,
										0,
										0,
										TEMP_CMD1,
										0,
										0,
										pReadData,
										pbReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	return 0;
}

LONG APIENTRY SLE4442Cmd_UpdateSecurityMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN  UCHAR	bData
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SLE4442_CARD_BREAK(m_ctrlCSerial,bSlotNum);
	if(lngStatus!=1)
	{
		return 1;
	}	
	BYTE TEMP_CMD1 = 0x39;
	BYTE  pReadData[400];
	ULONG pbReturnLen;
	lngStatus = CMD_SLE4442_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										0,
										0,
										0,
										TEMP_CMD1,
										bAddress,
										bData,
										pReadData,
										&pbReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if((pReadData[0] != 0x7b) && (pReadData[0] != 0xf4))//clock number
	{
		return 1;
	}
	return 0;	
}

LONG APIENTRY SLE4442Cmd_CompareVerificationData(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bData
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SLE4442_CARD_BREAK(m_ctrlCSerial,bSlotNum);
	if(lngStatus!=1)
	{
		return 1;
	}	
	BYTE TEMP_CMD1 = 0x33;
	BYTE  pReadData[400];
	ULONG pbReturnLen;
	lngStatus = CMD_SLE4442_CARD_COMMAND(
										m_ctrlCSerial,
										bSlotNum,
										0,
										0,
										0,
										TEMP_CMD1,
										bAddress,
										bData,
										pReadData,
										&pbReturnLen
										);
	if(lngStatus!=1)
	{
		return 1;
	}
	if(pReadData[0] != 0x02)//clock number
	{
		return 1;
	}
	return 0;	
}


//-------------AT45D041 Card------------------
LONG APIENTRY AT45D041Cmd(
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
		)
{
	BYTE AddressByte1 = 0;
	BYTE AddressByte2 = 0;
	BYTE AddressByte3 = 0;

	BOOL Result = 0;

	AddressByte1 = (PageNo>>7);
	AddressByte2 = ((PageNo<<1)&0xFF)|(lngStartAddr>>8);
	AddressByte3 = lngStartAddr&0xFF;

	LPCVOID	pSendBuffer;

	UINT SendLength = 0;

	switch(OPcode)
	{
		//main memory page read
		case 0x52:
		{
			SendLength = 8+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;
			*((PUCHAR)pSendBuffer+4)=0;
			*((PUCHAR)pSendBuffer+5)=0;
			*((PUCHAR)pSendBuffer+6)=0;
			*((PUCHAR)pSendBuffer+7)=0;
		}
		break;
		//Buffer1 read
		case 0x54:
		{
			SendLength = 5+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;
			*((PUCHAR)pSendBuffer+4)=0;
		}
		break;
		//Buffer2 read
		case 0x56:
		{
			SendLength = 5+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;
			*((PUCHAR)pSendBuffer+4)=0;		
		}
		break;
		//Main Memory page to buffer1 xfr
		case 0x53:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;
		}
		break;
		//Main Memory page to buffer2 xfr
		case 0x55:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;
		}
		break;
		//Main Memory Page to Buffer1 comp.
		case 0x60:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;			
		}
		break;
		//Main Memory Page to Buffer2 comp.
		case 0x61:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;			
		}
		break;
		//Buffer1 write
		case 0x84:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;	
			memcpy( (PUCHAR)pSendBuffer + 4, pWriteData, lngWriteLen);
		}
		break;
		//Buffer2 write
		case 0x87:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;	
			memcpy( (PUCHAR)pSendBuffer + 4, pWriteData, lngWriteLen);		
		}
		break;
		//B1 to Mem.Page prog.with erase 
		case 0x83:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;			
		}
		break;
		//B2 to Mem.Page prog.with erase 
		case 0x86:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;			
		}
		break;
		//B1 to Mem.Page prog.without erase
		case 0x88:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;			
		}
		break;
		//B2 to Mem.Page prog.without erase
		case 0x89:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;			
		}
		break;
		//mem page prog through b1.
		case 0x82:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;	
			memcpy( (PUCHAR)pSendBuffer + 4, pWriteData, lngWriteLen);	
		}
		break;
		//mem page prog through b2.
		case 0x85:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;	
			memcpy( (PUCHAR)pSendBuffer + 4, pWriteData, lngWriteLen);	
		}
		break;
		//auto page rewrite through b1.
		case 0x58:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;			
		}
		break;
		//auto page rewrite through b2.
		case 0x59:
		{
			SendLength = 4+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
			*((PUCHAR)pSendBuffer+1)=AddressByte1;
			*((PUCHAR)pSendBuffer+2)=AddressByte2;
			*((PUCHAR)pSendBuffer+3)=AddressByte3;			
		}
		break;
		//get status register
		case 0x57:
		{
			SendLength = 1+lngWriteLen;
			pSendBuffer	 = malloc(SendLength);	

			*((PUCHAR)pSendBuffer+0)=OPcode;
		}
		break;
		default:
		{
			return 1;//Unknown Opcode
		}
	}
/*
LONG CMD_AT45D041_CARD_COMMAND(CSerial *m_ctrlCSerial,
								BYTE bSlotNum,
								ULONG	lngWriteLen,
								UCHAR	*pWriteData,
								IN	ULONG	lngReadLen,
						        OUT	LPVOID	pReadData,
								OUT	ULONG	*plngReturnLen)
*/
	Result = CMD_AT45D041_CARD_COMMAND(m_ctrlCSerial,
										bSlotNum,
										SendLength,
										(PUCHAR)pSendBuffer,
										lngReadLen,
										(PUCHAR)pReadData,
										plngReturnLen);	
	free((LPVOID)pSendBuffer);
	if(Result!=1)
	{
		return 1;
	}
	return 0;
}

//-------------AT88SC1608 Card------------------
/*LONG CMD_SMC_COMMAND(CSerial *m_ctrlCSerial,
								BYTE bSlotNum,
								ULONG	lngWriteLen,
								UCHAR	*pWriteData,
								IN	ULONG	lngReadLen,
						        OUT	LPVOID	pReadData,
								OUT	ULONG	*plngReturnLen)*/
//suceess 0; fail 1.
LONG APIENTRY AT88SC1608Cmd_WriteUserZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bWriteLen,
		IN	LPVOID	pWriteBuffer
		)
{
	LONG	lngStatus;

	BYTE SendData[300];
	SendData[0] = 0xb0;
	SendData[1] = bAddress;
	memcpy((SendData + 2), (BYTE *)pWriteBuffer, bWriteLen);

	BYTE GetData[300];
	ULONG ReturnLen = 0;

	lngStatus = CMD_SMC_COMMAND(
								m_ctrlCSerial,
								bSlotNum,
								(bWriteLen + 2),
								SendData,
								0,
						        OUT	GetData,
								OUT	&ReturnLen);
	if(lngStatus != 1)
	{
		return 1;
	}
	else
	{
		return 0;
	}
}

LONG APIENTRY AT88SC1608Cmd_ReadUserZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadBuffer,
		OUT	UCHAR	*pReturnLen
		)
{
	LONG	lngStatus;

	BYTE SendData[300];
	SendData[0] = 0xb1;
	SendData[1] = bAddress;

	lngStatus = CMD_SMC_COMMAND(
								m_ctrlCSerial,
								bSlotNum,
								2,
								SendData,
								bReadLen,
						        OUT	pReadBuffer,
								OUT	(ULONG *)pReturnLen);
	if(lngStatus != 1)
	{
		return 1;
	}
	else
	{
		return 0;
	}
}
LONG APIENTRY AT88SC1608Cmd_WriteConfigurationZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bWriteLen,
		IN	LPVOID	pWriteBuffer
		)
{
	LONG	lngStatus;

	BYTE SendData[300];
	SendData[0] = 0xb4;
	SendData[1] = bAddress;
	memcpy((SendData + 2), (BYTE *)pWriteBuffer, bWriteLen);

	BYTE GetData[300];
	ULONG ReturnLen = 0;

	lngStatus = CMD_SMC_COMMAND(
								m_ctrlCSerial,
								bSlotNum,
								(bWriteLen + 2),
								SendData,
								0,
						        OUT	GetData,
								OUT	&ReturnLen);
	if(lngStatus != 1)
	{
		return 1;
	}
	else
	{
		return 0;
	}
}
LONG APIENTRY AT88SC1608Cmd_ReadConfigurationZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadBuffer,
		OUT	UCHAR	*pReturnLen
		)
{
	LONG	lngStatus;

	BYTE SendData[300];
	SendData[0] = 0xb5;
	SendData[1] = bAddress;

	lngStatus = CMD_SMC_COMMAND(
								m_ctrlCSerial,
								bSlotNum,
								2,
								SendData,
								bReadLen,
						        OUT	pReadBuffer,
								OUT	(ULONG *)pReturnLen);
	if(lngStatus != 1)
	{
		return 1;
	}
	else
	{
		return 0;
	}
}

LONG APIENTRY AT88SC1608Cmd_SetUserZoneAddress(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress
		)
{
	LONG	lngStatus;

	BYTE SendData[300];
	SendData[0] = 0xb2;
	SendData[1] = bAddress;

	BYTE GetData[300];
	ULONG ReturnLen = 0;

	lngStatus = CMD_SMC_COMMAND(
								m_ctrlCSerial,
								bSlotNum,
								2,
								SendData,
								0,
						        OUT	GetData,
								OUT	&ReturnLen);
	if(lngStatus != 1)
	{
		return 1;
	}
	else
	{
		return 0;
	}
} 

LONG APIENTRY AT88SC1608Cmd_VerifyPassword(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bZoneNo,
		IN	BOOL 	bIsReadAccess,
		IN	UCHAR	bPW1,
		IN	UCHAR	bPW2,
		IN	UCHAR	bPW3
		)
{
	LONG	lngStatus;

	BYTE SendData[300];
	SendData[0] = 0xb3;
	SendData[1] = bZoneNo;
	if(bIsReadAccess)
	{
		SendData[1] = (SendData[1] | 0x08);
	}
	SendData[2] = bPW1;
	SendData[3] = bPW2;
	SendData[4] = bPW3;


	BYTE GetData[300];
	ULONG ReturnLen = 0;

	lngStatus = CMD_SMC_COMMAND(
								m_ctrlCSerial,
								bSlotNum,
								5,
								SendData,
								0,
						        OUT	GetData,
								OUT	&ReturnLen);
	if(lngStatus != 1)
	{
		return 1;
	}
	else
	{
		return 0;
	}
}
 

 
 













//-------------Memory Card-------------
//suceess 0; fail 1.
/*
LONG CMD_SET_I2C_ADD(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  BYTE DeviceAddress,
						  UINT WordAddr,
						  BYTE PageSize);
LONG CMD_WRITE_I2C(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  ULONG	lngWriteLen,
						  UCHAR	*pWriteData);
LONG CMD_READ_I2C(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  IN	ULONG	lngReadLen,
						  OUT	LPVOID	pReadData,
						  OUT	ULONG	*plngReturnLen);
*/
LONG APIENTRY AT24CxxCmd_Write(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngDeviceAddr,
		IN	ULONG	lngStartAddr,
		IN	ULONG	lngWordPageSize,
		IN	ULONG	lngWriteLen,
		IN	UCHAR	*pWriteData
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SET_I2C_ADD(m_ctrlCSerial,
								bSlotNum,
								(BYTE)lngDeviceAddr,
								lngStartAddr,
								(BYTE)lngWordPageSize);
	if(lngStatus!=1)
	{
		return 1;
	}
	lngStatus = CMD_WRITE_I2C(m_ctrlCSerial,
							  bSlotNum,
							  lngWriteLen,
							  pWriteData);
	if(lngStatus!=1)
	{
		return 1;
	}
	return 0;
}

LONG APIENTRY AT24CxxCmd_Read(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngDeviceAddr,
		IN	ULONG	lngStartAddr,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		)
{
	LONG	lngStatus;
	lngStatus = CMD_SET_I2C_ADD(m_ctrlCSerial,
								bSlotNum,
								(BYTE)lngDeviceAddr,
								lngStartAddr,
								0);
	if(lngStatus!=1)
	{
		return 1;
	}
	lngStatus = CMD_READ_I2C(m_ctrlCSerial,
							 bSlotNum,
						  IN	lngReadLen,
						  OUT	pReadData,
						  OUT	plngReturnLen);
	if(lngStatus!=1)
	{
		return 1;
	}
	return 0;
}
 

