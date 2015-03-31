#ifndef AU9525_RS232_LIB_H
#define AU9525_RS232_LIB_H
#include "AFX.h"
#include <windows.h>
#include <Winbase.h>



class CSerial
{
public:
	CSerial()
	{
		m_bOpened = 0;
		SendDataSequenceNum = 0;
	}
	HANDLE m_hComDev;
	BOOL m_bOpened;
	unsigned char SendDataSequenceNum;
	OVERLAPPED m_OverlappedRead;
	OVERLAPPED m_OverlappedWrite;
	LPDWORD m_hIDComDev;
	int Open( int nPort, int nBaud );
	void ClosePort(void);
	int InBufferCount( void );
	long SendData( const char *buffer, unsigned long dwBytesWritten);
	long ReadData( void *buffer, unsigned long dwBytesRead);
	long ReadDataWithoutWait(void *buffer, unsigned long dwBytesRead);
};

class CATR_capacity
{
#define T_unknown 0
#define T0        1
#define T1        2
#define T0_T1     3
public:
	CATR_capacity()
	{
		T_Type = T_unknown;
	}
	BYTE FI_DI;
	BYTE T_Type;
	BOOL Do_ATR(BYTE *DataIN, UINT Len);	
};

class Reader_Descriptor
{
public:
	Reader_Descriptor()
	{
		BufferSize = 0;
		APDU_Type = 0;
		PID = 0;
		VID = 0;
		ReleaseNumber = 0;
		ManufactureString = "";
		ProductString = "";
		SerialNumber = "";	
	}
	UINT	BufferSize;
	BOOL	APDU_Type;
	UINT	PID;
	UINT	VID;
	UINT	ReleaseNumber;
	CString ManufactureString;
	CString ProductString;
	CString SerialNumber;
};
//=======================
//-----9525 CMD----------
//======================= 
BOOL CMD_PC_to_RDR_IccPowerOn(CSerial *m_ctrlCSerial, 
							  BYTE SlotIn, BYTE PowerSelect,
							  BYTE *StatusOut, BYTE *ErrorOut, BYTE *ChainParameterOut,
							  BYTE *ATRbuffer, BYTE *ATRbufferLen);
BOOL CMD_PC_to_RDR_IccPowerOff(CSerial *m_ctrlCSerial, 
							  BYTE SlotIn,
							  BYTE *StatusOut, BYTE *ErrorOut, BYTE *ClockStatus);
BOOL CMD_PC_to_RDR_GetSlotStatus(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE *StatusOut, BYTE *ErrorOut, BYTE *ClockStatus);
BOOL CMD_PC_to_RDR_XfrBlock(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE BWIIn, UINT LevelParameter,
								BYTE *abDataIn, UINT abDataInLen,
							    BYTE *StatusOut, BYTE *ErrorOut, 
								BYTE *ChainParameterOut,
							    BYTE *abDataOut, UINT *abDataOutLen);
BOOL CMD_PC_to_RDR_SetParameters(CSerial *m_ctrlCSerial,
								BYTE SlotIn,
								BYTE bProtocolNumIn,
								BYTE *abDataIn, ULONG abDataInLen,
							    BYTE *StatusOut, BYTE *ErrorOut,
							    BYTE *abDataOut, ULONG *abDataOutLen);
BOOL CMD_PC_to_RDR_Escape(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE *abDataIn, ULONG abDataInLen,
							    BYTE *StatusOut, BYTE *ErrorOut, 
							    BYTE *abDataOut, ULONG *abDataOutLen);
BOOL CMD_PC_to_RDR_IccClock(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE ClockCommand,
							    BYTE *StatusOut, BYTE *ErrorOut, 
							    BYTE *ClockStatus);
BOOL CMD_PC_to_RDR_T0APDU(CSerial *m_ctrlCSerial, 
								BYTE SlotIn,
								BYTE bmChangesIn, BYTE bClassGetResponseIn,
								BYTE bClassEnvelopeIn,
							    BYTE *StatusOut, BYTE *ErrorOut, 
							    BYTE *ClockStatus);
BOOL Check_RDR_to_PC_NotifySlotChange(CSerial *m_ctrlCSerial, BYTE *bmSlotICCStateOut);
//===========math functions==========
CString Bytes2CString(BYTE *DataIN, UINT Len);
CString Bytes2CString_ASCII(BYTE *DataIN, UINT Len);
char str2char(char para_str);
UINT CString2Bytes (BYTE *DataOUT, CString DataIN);
//=====================================================================
//-----9525 Vendor CMD----------
//===================================================================== 
LONG CMD_SetBaudRate(CSerial *m_ctrlCSerial,BYTE BaudRate);

LONG CMD_GetReaderDescriptor(CSerial *m_ctrlCSerial,
							 BYTE bDevDescIn,
							 BYTE bStrDescIn,
							 BYTE *StatusOut, BYTE *ErrorOut,
							 BYTE *abDataOut, ULONG *abDataOutLen
							 );
//-----------LCM,Keypad,EE2prom-----------------------------------
LONG EepromCmdWrite(
		CSerial *m_ctrlCSerial, 
		UCHAR	bSlotNum,
		ULONG	lngStartAddr,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		);
LONG EepromCmdRead(
		CSerial *m_ctrlCSerial, 
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngStartAddr,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		);
LONG CmdFirmwareWrite(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bIndex,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		);

LONG APIENTRY CmdFirmwareRead(
		CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bIndex,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		);

LONG CmdLcmWriteData(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		ULONG	lngWriteLen,
		UCHAR	*pWriteData
		);

LONG CmdTestFwUpDate(
		CSerial *m_ctrlCSerial
		);
LONG CmdLcmSetBacklight(
		CSerial *m_ctrlCSerial
		);
LONG CmdClearKeyBuffer(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum
		);
LONG CmdSetKeyScanTimer(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bXTimes,
		UCHAR	bTimerFlag
		);
/*LONG CmdSetKeyScanTimer(
		CSerial *m_ctrlCSerial,
		UCHAR	bSlotNum,
		UCHAR	bXL,
		UCHAR	bXH,
		UCHAR	bTimerFlag
		);*/
LONG APIENTRY CmdGetKeyInput(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		);
LONG WriteCommandNocheck(
		CSerial *m_ctrlCSerial,
		UCHAR bCmd
		);
LONG WriteData(
		CSerial *m_ctrlCSerial,
		UCHAR	bData
		);
LONG SelectScreen(
		CSerial *m_ctrlCSerial,
		UCHAR	bSelectScreen
		);
LONG SetBackLight(
		CSerial *m_ctrlCSerial,
		UCHAR	bSelectBacklight
		);

//--------------------for synchronous card----------------------------
LONG CMD_SWITCH_CARD_MODE(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  BYTE Card_Mode_Switch
						  );
LONG CMD_POWER_ON(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum);
LONG CMD_POWER_OFF(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum);
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
LONG CMD_AT45D041_CARD_COMMAND(CSerial *m_ctrlCSerial,
								BYTE bSlotNum,
								ULONG	lngWriteLen,
								UCHAR	*pWriteData,
								IN	ULONG	lngReadLen,
						        OUT	LPVOID	pReadData,
								OUT	ULONG	*plngReturnLen);
LONG CMD_SMC_COMMAND(CSerial *m_ctrlCSerial,
								BYTE bSlotNum,
								ULONG	lngWriteLen,
								UCHAR	*pWriteData,
								IN	ULONG	lngReadLen,
						        OUT	LPVOID	pReadData,
								OUT	ULONG	*plngReturnLen);
LONG CMD_SLE4442_CARD_COMMAND(CSerial *m_ctrlCSerial,
							  BYTE bSlotNum,
						IN	  UINT ReadingLength,
							  BYTE ProtectBit,
							  BYTE AskForClkNum,
							  BYTE CMD_1,
							  BYTE CMD_2,
							  BYTE CMD_3,
						OUT	  LPVOID pReadData,
						OUT	  ULONG	*plngReturnLen);
LONG CMD_SLE4442_CARD_BREAK(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum);
LONG CMD_SLE4428_CARD_COMMAND(CSerial *m_ctrlCSerial,
							  BYTE bSlotNum,
						IN	  UINT ReadingLength,
							  BYTE ProtectBit,
							  BYTE AskForClkNum,
							  BYTE CMD_1,
							  BYTE CMD_2,
							  BYTE CMD_3,
						OUT	  LPVOID pReadData,
						OUT	  ULONG	*plngReturnLen);

LONG CMD_INPHONE_CARD_RESET(CSerial *m_ctrlCSerial,
						    BYTE bSlotNum);
LONG CMD_INPHONE_CARD_Read(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  IN	ULONG	lngReadLen,
						  OUT	LPVOID	pReadData,
						  OUT	ULONG	*plngReturnLen);
//LONG CMD_INPHONE_CARD_PROG(CSerial *m_ctrlCSerial,
//						  BYTE bSlotNum,
//						  ULONG	lngWriteLen,
//						  UCHAR	*pWriteData);
LONG CMD_INPHONE_CARD_PROG(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  IN	ULONG	lngWriteLen,
						  OUT	LPVOID	pReadData,
						  OUT	ULONG	*plngReturnLen);
LONG CMD_INPHONE_CARD_MOVE_ADDRESS(CSerial *m_ctrlCSerial,
						  BYTE bSlotNum,
						  UINT ClockNums);
LONG CMD_INPHONE_CARD_AUTHENTICATION_KEY1(CSerial *m_ctrlCSerial,
										BYTE bSlotNum,
										ULONG	lngWriteLen,
										UCHAR	*pWriteData,
										IN	ULONG	lngReadLen,
										OUT	LPVOID	pReadData,
										OUT	ULONG	*plngReturnLen);
LONG CMD_INPHONE_CARD_AUTHENTICATION_KEY2(CSerial *m_ctrlCSerial,
										BYTE bSlotNum,
										ULONG	lngWriteLen,
										UCHAR	*pWriteData,
										IN	ULONG	lngReadLen,
										OUT	LPVOID	pReadData,
										OUT	ULONG	*plngReturnLen);

//=============================================
//Õ¨≤Ωø®√¸¡Ó
//=============================================
//-------------4418/28 Card------------------
LONG SLE4428Cmd_WriteEraseWithPB(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	UCHAR	bData
		);
LONG SLE4428Cmd_WriteEraseWithoutPB(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	UCHAR	bData
		);
LONG SLE4428Cmd_WritePBWithDataComparison(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	UCHAR	bData
		);
LONG SLE4428Cmd_Read9Bits(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	LPVOID	pReadPB,
		OUT	ULONG	*plngReturnLen
		);
LONG  SLE4428Cmd_Read8Bits(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngAddress,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		);
LONG SLE4428Cmd_WriteErrorCounter(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bData
		);
LONG SLE4428Cmd_Verify1stPSC(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bData
		);
LONG SLE4428Cmd_Verify2ndPSC(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bData
		);
//-------------4432/42 Card------------------
//suceess 0; fail 1.
LONG APIENTRY SLE4442Cmd_ReadMainMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*pbReturnLen
		);
LONG APIENTRY SLE4442Cmd_UpdateMainMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bData
		);
LONG APIENTRY SLE4442Cmd_ReadProtectionMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*pbReturnLen
		);
LONG APIENTRY SLE4442Cmd_WriteProtectionMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bData
		);
LONG APIENTRY SLE4442Cmd_ReadSecurityMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*pbReturnLen
		);
LONG APIENTRY SLE4442Cmd_UpdateSecurityMemory(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN  UCHAR	bData
		);
LONG APIENTRY SLE4442Cmd_CompareVerificationData(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR	bAddress,
		IN	UCHAR	bData
		);
//-------------AT45D041 Card------------------
//suceess 0; fail 1.
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
		);
//-------------AT88SC1608-------------
//suceess 0; fail 1.
LONG APIENTRY AT88SC1608Cmd_WriteUserZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bWriteLen,
		IN	LPVOID	pWriteBuffer
		);
LONG APIENTRY AT88SC1608Cmd_ReadUserZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadBuffer,
		OUT	UCHAR	*pReturnLen
		);
LONG APIENTRY AT88SC1608Cmd_WriteConfigurationZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bWriteLen,
		IN	LPVOID	pWriteBuffer
		);
LONG APIENTRY AT88SC1608Cmd_ReadConfigurationZone(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress,
		IN	UCHAR	bReadLen,
		OUT	LPVOID	pReadBuffer,
		OUT	UCHAR	*pReturnLen
		);
LONG APIENTRY AT88SC1608Cmd_SetUserZoneAddress(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bAddress
		);
LONG APIENTRY AT88SC1608Cmd_VerifyPassword(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	UCHAR   bZoneNo,
		IN	BOOL 	bIsReadAccess,
		IN	UCHAR	bPW1,
		IN	UCHAR	bPW2,
		IN	UCHAR	bPW3
		);


//-------------Memory Card-------------
//suceess 0; fail 1.
LONG APIENTRY AT24CxxCmd_Write(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngDeviceAddr,
		IN	ULONG	lngStartAddr,
		IN	ULONG	lngWordPageSize,
		IN	ULONG	lngWriteLen,
		IN	UCHAR	*pWriteData
		);
LONG APIENTRY AT24CxxCmd_Read(
		IN	CSerial *m_ctrlCSerial,
		IN	UCHAR	bSlotNum,
		IN	ULONG	lngDeviceAddr,
		IN	ULONG	lngStartAddr,
		IN	ULONG	lngReadLen,
		OUT	LPVOID	pReadData,
		OUT	ULONG	*plngReturnLen
		);


#endif