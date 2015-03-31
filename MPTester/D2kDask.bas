Attribute VB_Name = "D2K_DASK"
Option Explicit

'ADLink PCI Card Type
Global Const DAQ_2010 = 1
Global Const DAQ_2205 = 2
Global Const DAQ_2206 = 3
Global Const DAQ_2005 = 4
Global Const DAQ_2204 = 5
Global Const DAQ_2006 = 6
Global Const DAQ_2501 = 7
Global Const DAQ_2502 = 8
Global Const DAQ_2208 = 9
Global Const DAQ_2213 = 10
Global Const DAQ_2214 = 11
Global Const DAQ_2016 = 12

Global Const MAX_CARD = 32

'Error Code
Global Const NoError = 0
Global Const ErrorUnknownCardType = -1
Global Const ErrorInvalidCardNumber = -2
Global Const ErrorTooManyCardRegistered = -3
Global Const ErrorCardNotRegistered = -4
Global Const ErrorFuncNotSupport = -5
Global Const ErrorInvalidIoChannel = -6
Global Const ErrorInvalidAdRange = -7
Global Const ErrorContIoNotAllowed = -8
Global Const ErrorDiffRangeNotSupport = -9
Global Const ErrorLastChannelNotZero = -10
Global Const ErrorChannelNotDescending = -11
Global Const ErrorChannelNotAscending = -12
Global Const ErrorOpenDriverFailed = -13
Global Const ErrorOpenEventFailed = -14
Global Const ErrorTransferCountTooLarge = -15
Global Const ErrorNotDoubleBufferMode = -16
Global Const ErrorInvalidSampleRate = -17
Global Const ErrorInvalidCounterMode = -18
Global Const ErrorInvalidCounter = -19
Global Const ErrorInvalidCounterState = -20
Global Const ErrorInvalidBinBcdParam = -21
Global Const ErrorBadCardType = -22
Global Const ErrorInvalidDaRange = -23
Global Const ErrorAdTimeOut = -24
Global Const ErrorNoAsyncAI = -25
Global Const ErrorNoAsyncAO = -26
Global Const ErrorNoAsyncDI = -27
Global Const ErrorNoAsyncDO = -28
Global Const ErrorNotInputPort = -29
Global Const ErrorNotOutputPort = -30
Global Const ErrorInvalidDioPort = -31
Global Const ErrorInvalidDioLine = -32
Global Const ErrorContIoActive = -33
Global Const ErrorDblBufModeNotAllowed = -34
Global Const ErrorConfigFailed = -35
Global Const ErrorInvalidPortDirection = -36
Global Const ErrorBeginThreadError = -37
Global Const ErrorInvalidPortWidth = -38
Global Const ErrorInvalidCtrSource = -39
Global Const ErrorOpenFile = -40
Global Const ErrorAllocateMemory = -41
Global Const ErrorDaVoltageOutOfRange = -42
Global Const ErrorInvalidSyncMode = -43
Global Const ErrorInvalidBufferID = -44
Global Const ErrorInvalidCNTInterval = -45
Global Const ErrorReTrigModeNotAllowed = -46
Global Const ErrorResetBufferNotAllowed = -47
Global Const ErrorAnaTriggerLevel = -48
Global Const ErrorDAQEvent = -49

'Error code for driver API
Global Const ErrorConfigIoctl = -201
Global Const ErrorAsyncSetIoctl = -202
Global Const ErrorDBSetIoctl = -203
Global Const ErrorDBHalfReadyIoctl = -204
Global Const ErrorContOPIoctl = -205
Global Const ErrorContStatusIoctl = -206
Global Const ErrorPIOIoctl = -207
Global Const ErrorDIntSetIoctl = -208
Global Const ErrorWaitEvtIoctl = -209
Global Const ErrorOpenEvtIoctl = -210
Global Const ErrorCOSIntSetIoctl = -211
Global Const ErrorMemMapIoctl = -212
Global Const ErrorMemUMapSetIoctl = -213
Global Const ErrorCTRIoctl = -214
Global Const ErrorGetResIoctl = -215

'Synchronous Mode
Global Const SYNCH_OP = 1
Global Const ASYNCH_OP = 2

'AD Range
Global Const AD_B_10_V = 1
Global Const AD_B_5_V = 2
Global Const AD_B_2_5_V = 3
Global Const AD_B_1_25_V = 4
Global Const AD_B_0_625_V = 5
Global Const AD_B_0_3125_V = 6
Global Const AD_B_0_5_V = 7
Global Const AD_B_0_05_V = 8
Global Const AD_B_0_005_V = 9
Global Const AD_B_1_V = 10
Global Const AD_B_0_1_V = 11
Global Const AD_B_0_01_V = 12
Global Const AD_B_0_001_V = 13
Global Const AD_U_20_V = 14
Global Const AD_U_10_V = 15
Global Const AD_U_5_V = 16
Global Const AD_U_2_5_V = 17
Global Const AD_U_1_25_V = 18
Global Const AD_U_1_V = 19
Global Const AD_U_0_1_V = 20
Global Const AD_U_0_01_V = 21
Global Const AD_U_0_001_V = 22
Global Const AD_B_2_V = 23
Global Const AD_B_0_25_V = 24
Global Const AD_B_0_2_V = 25
Global Const AD_U_4_V = 26
Global Const AD_U_2_V = 27
Global Const AD_U_0_5_V = 28
Global Const AD_U_0_4_V = 29

'Constants for DAQ2000
Global Const All_Channels = -1
Global Const BufferNotUsed = -1

Global Const DAQ2K_AI_ADSTARTSRC_Int = &H0
Global Const DAQ2K_AI_ADSTARTSRC_AFI0 = &H10
Global Const DAQ2K_AI_ADSTARTSRC_SSI = &H20

Global Const DAQ2K_AI_ADCONVSRC_Int = &H0
Global Const DAQ2K_AI_ADCONVSRC_AFI0 = &H4
Global Const DAQ2K_AI_ADCONVSRC_SSI = &H8
Global Const DAQ2K_AI_ADCONVSRC_AFI1 = &HC

'Constants for AI Delay Counter SRC: only available for DAQ-250X
Global Const DAQ2K_AI_DTSRC_Int = &H0
Global Const DAQ2K_AI_DTSRC_AFI1 = &H10
Global Const DAQ2K_AI_DTSRC_GPTC0 = &H20
Global Const DAQ2K_AI_DTSRC_GPTC1 = &H30

Global Const DAQ2K_AI_TRGSRC_SOFT = &H0
Global Const DAQ2K_AI_TRGSRC_ANA = &H1
Global Const DAQ2K_AI_TRGSRC_ExtD = &H2
Global Const DAQ2K_AI_TRSRC_SSI = &H3
Global Const DAQ2K_AI_TRGMOD_POST = &H0    'Post Trigger Mode
Global Const DAQ2K_AI_TRGMOD_DELAY = &H8   'Delay Trigger Mode
Global Const DAQ2K_AI_TRGMOD_PRE = &H10    'Pre-Trigger Mode
Global Const DAQ2K_AI_TRGMOD_MIDL = &H18   'Middle Trigger Mode
Global Const DAQ2K_AI_ReTrigEn = &H80
Global Const DAQ2K_AI_Dly1InSamples = &H100
Global Const DAQ2K_AI_Dly1InTimebase = &H0
Global Const DAQ2K_AI_MCounterEn = &H400
Global Const DAQ2K_AI_TrgPositive = &H0
Global Const DAQ2K_AI_TrgNegative = &H1000

'AI Reference Ground (input mode)
Global Const AI_RSE = &H0
Global Const AI_DIFF = &H100
Global Const AI_NRSE = &H200

Global Const DAQ2K_DA_BiPolar = &H1
Global Const DAQ2K_DA_UniPolar = &H0
Global Const DAQ2K_DA_Int_REF = &H0
Global Const DAQ2K_DA_Ext_REF = &H1

Global Const DAQ2K_DA_WRSRC_Int = &H0
Global Const DAQ2K_DA_WRSRC_AFI0 = &H1
Global Const DAQ2K_DA_WRSRC_AFI1 = &H1
Global Const DAQ2K_DA_WRSRC_SSI = &H2

'DA group
Global Const DA_Group_A = &H0
Global Const DA_Group_B = &H4
Global Const DA_Group_AB = &H8

'DA TD Counter SRC: only available for DAQ-250X
Global Const DAQ2K_DA_TDSRC_Int = &H0
Global Const DAQ2K_DA_TDSRC_AFI0 = &H10
Global Const DAQ2K_DA_TDSRC_GPTC0 = &H20
Global Const DAQ2K_DA_TDSRC_GPTC1 = &H30

'DA BD Counter SRC: only available for DAQ-250X
Global Const DAQ2K_DA_BDSRC_Int = &H0
Global Const DAQ2K_DA_BDSRC_AFI0 = &H40
Global Const DAQ2K_DA_BDSRC_GPTC0 = &H80
Global Const DAQ2K_DA_BDSRC_GPTC1 = &HC0

'DA trigger constant
Global Const DAQ2K_DA_TRGSRC_SOFT = &H0    'Software Trigger Mode
Global Const DAQ2K_DA_TRGSRC_ANA = &H1    'Post Trigger Mode
Global Const DAQ2K_DA_TRGSRC_ExtD = &H2     'Delay Trigger Mode
Global Const DAQ2K_DA_TRGSRC_SSI = &H3
Global Const DAQ2K_DA_TRGMOD_POST = &H0
Global Const DAQ2K_DA_TRGMOD_DELAY = &H4
Global Const DAQ2K_DA_ReTrigEn = &H20
Global Const DAQ2K_DA_Dly1InUI = &H40
Global Const DAQ2K_DA_Dly1InTimebase = &H0
Global Const DAQ2K_DA_Dly2InUI = &H80
Global Const DAQ2K_DA_Dly2InTimebase = &H0
Global Const DAQ2K_DA_DLY2En = &H100
Global Const DAQ2K_DA_TrgPositive = &H0
Global Const DAQ2K_DA_TrgNegative = &H200
'DA stop mode
Global Const DAQ2K_DA_TerminateImmediate = 0
Global Const DAQ2K_DA_TerminateUC = 1
Global Const DAQ2K_DA_TerminateFIFORC = 2
Global Const DAQ2K_DA_TerminateIC = 2
'DA stop source : only available for DAQ-250X
Global Const DAQ2K_DA_STOPSRC_SOFT = 0
Global Const DAQ2K_DA_STOPSRC_AFI0 = 1
Global Const DAQ2K_DA_STOPSRC_ATrig = 2
Global Const DAQ2K_DA_STOPSRC_AFI1 = 3

'DIO Port Direction
Global Const INPUT_PORT = 1
Global Const OUTPUT_PORT = 2

'Channel&Port
Global Const Channel_P1A = 0
Global Const Channel_P1B = 1
Global Const Channel_P1C = 2
Global Const Channel_P1CL = 3
Global Const Channel_P1CH = 4
Global Const Channel_P1AE = 10
Global Const Channel_P1BE = 11
Global Const Channel_P1CE = 12
Global Const Channel_P2A = 5
Global Const Channel_P2B = 6
Global Const Channel_P2C = 7
Global Const Channel_P2CL = 8
Global Const Channel_P2CH = 9
Global Const Channel_P2AE = 15
Global Const Channel_P2BE = 16
Global Const Channel_P2CE = 17
Global Const Channel_P3A = 10
Global Const Channel_P3B = 11
Global Const Channel_P3C = 12
Global Const Channel_P3CL = 13
Global Const Channel_P3CH = 14
Global Const Channel_P4A = 15
Global Const Channel_P4B = 16
Global Const Channel_P4C = 17
Global Const Channel_P4CL = 18
Global Const Channel_P4CH = 19
Global Const Channel_P5A = 20
Global Const Channel_P5B = 21
Global Const Channel_P5C = 22
Global Const Channel_P5CL = 23
Global Const Channel_P5CH = 24
Global Const Channel_P6A = 25
Global Const Channel_P6B = 26
Global Const Channel_P6C = 27
Global Const Channel_P6CL = 28
Global Const Channel_P6CH = 29
Global Const Channel_P1 = 30
Global Const Channel_P2 = 31
Global Const Channel_P3 = 32
Global Const Channel_P4 = 33
Global Const Channel_P1E = 34
Global Const Channel_P2E = 35
Global Const Channel_P3E = 36
Global Const Channel_P4E = 37

'--------- Constants for Timer/Counter --------------
'Counter Mode (8254)
Global Const TOGGLE_OUTPUT = 0             'Toggle output from low to high on terminal count
Global Const PROG_ONE_SHOT = 1             'Programmable one-shot
Global Const RATE_GENERATOR = 2            'Rate generator
Global Const SQ_WAVE_RATE_GENERATOR = 3    'Square wave rate generator
Global Const SOFT_TRIG = 4                 'Software-triggered strobe
Global Const HARD_TRIG = 5                 'Hardware-triggered strobe

'---------- Constants for Analog trigger ------------
'define analog trigger condition constants
Global Const Below_Low_level = &H0
Global Const Above_High_Level = &H100
Global Const Inside_Region = &H200
Global Const High_Hysteresis = &H300
Global Const Low_Hysteresis = &H400

'define analog trigger Dedicated Channel
Global Const CH0ATRIG = &H0
Global Const CH1ATRIG = &H2
Global Const CH2ATRIG = &H4
Global Const CH3ATRIG = &H6
Global Const EXTATRIG = &H1
Global Const ADCATRIG = &H0

'----------- Time Base -------------------
Global Const DAQ2K_IntTimeBase = &H0
Global Const DAQ2K_ExtTimeBase = &H1
Global Const DAQ2K_SSITimeBase = &H2

'------- General Purpose Timer/Counter -----------------
'Counter Mode
Global Const SimpleGatedEventCNT = 1
Global Const SinglePeriodMSR = 2
Global Const SinglePulseWidthMSR = 3
Global Const SingleGatedPulseGen = 4
Global Const SingleTrigPulseGen = 5
Global Const RetrigSinglePulseGen = 6
Global Const SingleTrigContPulseGen = 7
Global Const ContGatedPulseGen = 8

'GPTC clock source
Global Const GPTC_GATESRC_EXT = &H4
Global Const GPTC_GATESRC_INT = &H0
Global Const GPTC_CLKSRC_EXT = &H8
Global Const GPTC_CLKSRC_INT = &H0
Global Const GPTC_UPDOWN_SEL_EXT = &H10
Global Const GPTC_UPDOWN_SEL_INT = &H0

'GPTC clock polarity
Global Const GPTC_CLKEN_LACTIVE = &H1
Global Const GPTC_CLKEN_HACTIVE = &H0
Global Const GPTC_GATE_LACTIVE = &H2
Global Const GPTC_GATE_HACTIVE = &H0
Global Const GPTC_UPDOWN_LACTIVE = &H4
Global Const GPTC_UPDOWN_HACTIVE = &H0
Global Const GPTC_OUTPUT_LACTIVE = &H8
Global Const GPTC_OUTPUT_HACTIVE = &H0
Global Const GPTC_INT_LACTIVE = &H10
Global Const GPTC_INT_HACTIVE = &H0

'GPTC paramID
Global Const GPTC_IntGATE = &H0
Global Const GPTC_IntUpDnCTR = &H1
Global Const GPTC_IntENABLE = &H2

'SSI signal code
Global Const SSI_TIME = 1
Global Const SSI_CONV = 2
Global Const SSI_WR = 4
Global Const SSI_ADSTART = 8
Global Const SSI_ADTRIG = &H20
Global Const SSI_DATRIG = &H40

'DAQ Event type for the event message
Global Const DAQEnd = 0
Global Const DBEvent = 1
Global Const TrigEvent = 2
Global Const DAQEnd_A = 0
Global Const DAQEnd_B = 2
Global Const DAQEnd_AB = 3
Global Const DATrigEvent = 4
Global Const DATrigEvent_A = 4
Global Const DATrigEvent_B = 5
Global Const DATrigEvent_AB = 6

'16-bit binary or 4-decade BCD counter
Global Const BIN = 0
Global Const BCD = 1

Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'-------------------------------------------------------------------
'  D2K-DASK Function prototype
'-----------------------------------------------------------------*/
Declare Function D2K_Register_Card Lib "D2K-Dask.dll" (ByVal cardType As Integer, ByVal card_num As Integer) As Integer
Declare Function D2K_Release_Card Lib "D2K-Dask.dll" (ByVal CardNumber As Integer) As Integer
Declare Function D2K_AIO_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal TimerBase As Integer, ByVal AnaTrigCtrl As Integer, ByVal H_TrgLevel As Integer, ByVal L_TrgLevel As Integer) As Integer

'AI Functions
Declare Function D2K_AI_CH_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, ByVal AdRange_RefGnd As Integer) As Integer
Declare Function D2K_AI_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal ConfigCtrl As Integer, ByVal TrigCtrl As Integer, ByVal MidOrDlyScans As Long, ByVal MCnt As Integer, ByVal ReTrgCnt As Integer, ByVal AutoResetBuf As Byte) As Integer
Declare Function D2K_AI_PostTrig_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal ClkSrc As Integer, ByVal TrigSrcCtrl As Integer, ByVal ReTrgEn As Integer, ByVal ReTrgCnt As Integer, ByVal AutoResetBuf As Byte) As Integer
Declare Function D2K_AI_DelayTrig_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal ClkSrc As Integer, ByVal TrigSrcCtrl As Integer, ByVal DlyScans As Long, ByVal ReTrgEn As Integer, ByVal ReTrgCnt As Integer, ByVal AutoResetBuf As Byte) As Integer
Declare Function D2K_AI_PreTrig_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal ClkSrc As Integer, ByVal TrigSrcCtrl As Integer, ByVal MCtrEn As Integer, ByVal MCnt As Integer, ByVal AutoResetBuf As Byte) As Integer
Declare Function D2K_AI_MiddleTrig_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal ClkSrc As Integer, ByVal TrigSrcCtrl As Integer, ByVal MiddleScans As Long, ByVal MCtrEn As Integer, ByVal MCnt As Integer, ByVal AutoResetBuf As Byte) As Integer
Declare Function D2K_AI_AsyncCheck Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, Stopped As Byte, AccessCnt As Long) As Integer
Declare Function D2K_AI_AsyncClear Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, startPos As Long, AccessCnt As Long) As Integer
Declare Function D2K_AI_AsyncDblBufferHalfReady Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, HalfReady As Byte, StopFlag As Byte) As Integer
Declare Function D2K_AI_AsyncDblBufferMode Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Enable As Byte) As Integer
Declare Function D2K_AI_AsyncDblBufferToFile Lib "D2K-Dask.dll" (ByVal CardNumber As Integer) As Integer
Declare Function D2K_AI_ContReadChannel Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, ByVal BufId As Integer, ByVal ReadScans As Long, ByVal ScanIntrv As Long, ByVal SampIntrv As Long, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AI_ContScanChannels Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal wChannel As Integer, ByVal BufId As Integer, ByVal ReadScans As Long, ByVal ScanIntrv As Long, ByVal SampIntrv As Long, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AI_ContReadMultiChannels Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal NumChans As Integer, chans As Integer, ByVal BufId As Integer, ByVal ReadScans As Long, ByVal ScanIntrv As Long, ByVal SampIntrv As Long, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AI_ContReadChannelToFile Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, ByVal BufId As Integer, ByVal FileName As String, ByVal ReadScans As Long, ByVal ScanIntrv As Long, ByVal SampIntrv As Long, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AI_ContScanChannelsToFile Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal wChannel As Integer, ByVal BufId As Integer, ByVal FileName As String, ByVal dwReadScans As Long, ByVal ScanIntrv As Long, ByVal SampIntrv As Long, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AI_ContReadMultiChannelsToFile Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal NumChans As Integer, chans As Integer, ByVal BufId As Integer, ByVal FileName As String, ByVal ReadScans As Long, ByVal ScanIntrv As Long, ByVal SampIntrv As Long, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AI_ContStatus Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, status As Integer) As Integer
Declare Function D2K_AI_InitialMemoryAllocated Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, MemSize As Long) As Integer
Declare Function D2K_AI_ReadChannel Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, Value As Integer) As Integer
Declare Function D2K_AI_VReadChannel Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, Voltage As Double) As Integer
Declare Function D2K_AI_SimuReadChannel Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal NumChans As Integer, chans As Integer, Buffer As Integer) As Integer
Declare Function D2K_AI_ScanReadChannels Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal NumChans As Integer, chans As Integer, Buffer As Integer) As Integer
Declare Function D2K_AI_VoltScale Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal AdRange As Integer, ByVal reading As Integer, Voltage As Double) As Integer
Declare Function D2K_AI_ContVScale Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal AdRange As Integer, readingArray As Integer, voltageArray As Double, ByVal Count As Long) As Integer
Declare Function D2K_AI_ContBufferSetup Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, Buffer As Any, ByVal ReadCount As Long, BufferId As Integer) As Integer
Declare Function D2K_AI_ContBufferReset Lib "D2K-Dask.dll" (ByVal CardNumber As Integer) As Integer
Declare Function D2K_AI_MuxScanSetup Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal wNumChans As Integer, chans As Integer, AdRange_RefGnds As Integer) As Integer
Declare Function D2K_AI_ReadMuxScan Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, Buffer As Integer) As Integer
Declare Function D2K_AI_ContMuxScan Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal BufId As Integer, ByVal ReadScans As Long, ByVal ScanIntrv As Long, ByVal SampIntrv As Long, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AI_ContMuxScanToFile Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal BufId As Integer, ByVal FileName As String, ByVal ReadScans As Long, ByVal ScanIntrv As Long, ByVal SampIntrv As Long, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AI_EventCallBack Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Mode As Integer, ByVal EventType As Integer, ByVal callbackAddr As Long) As Integer
Declare Function D2K_AI_AsyncReTrigNextReady Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, trgReady As Byte, StopFlag As Byte, RdyTrigCnt As Integer) As Integer
Declare Function D2K_AI_AsyncDblBufferHandled Lib "D2K-Dask.dll" (ByVal CardNumber As Integer) As Integer
Declare Function D2K_AI_AsyncDblBufferOverrun Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal op As Integer, overrunFlag As Integer) As Integer

'AO Functions
Declare Function D2K_AO_CH_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, ByVal OutputPolarity As Integer, ByVal wIntOrExtRef As Integer, ByVal refVoltage As Double) As Integer
Declare Function D2K_AO_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal ConfigCtrl As Integer, ByVal TrigCtrl As Integer, ByVal ReTrgCnt As Integer, ByVal DLY1Cnt As Integer, ByVal DLY2Cnt As Integer, ByVal AutoResetBuf As Byte) As Integer
Declare Function D2K_AO_PostTrig_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal ClkSrc As Integer, ByVal TrigSrcCtrl As Integer, ByVal DLY2Ctrl As Integer, ByVal DLY2Cnt As Integer, ByVal ReTrgEn As Integer, ByVal ReTrgCnt As Integer, ByVal AutoResetBuf As Byte) As Integer
Declare Function D2K_AO_DelayTrig_Config Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal ClkSrc As Integer, ByVal TrigSrcCtrl As Integer, ByVal DLY1Cnt As Integer, ByVal DLY2Ctrl As Integer, ByVal DLY2Cnt As Integer, ByVal ReTrgEn As Integer, ByVal ReTrgCnt As Integer, ByVal AutoResetBuf As Byte) As Integer
Declare Function D2K_AO_InitialMemoryAllocated Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, MemSize As Long) As Integer
Declare Function D2K_AO_AsyncCheck Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, Stopped As Byte, WriteCnt As Long) As Integer
Declare Function D2K_AO_AsyncClear Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, WriteCnt As Long, ByVal stop_mode As Integer) As Integer
Declare Function D2K_AO_AsyncDblBufferHalfReady Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, HalfReady As Byte) As Integer
Declare Function D2K_AO_AsyncDblBufferMode Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Enable As Byte) As Integer
Declare Function D2K_AO_ContWriteChannel Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, ByVal BufId As Integer, ByVal UpdateCount As Long, ByVal Iterations As Long, ByVal CHUI As Long, ByVal definite As Integer, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AO_ContWriteMultiChannels Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal NumChans As Integer, chans As Integer, ByVal BufId As Integer, ByVal UpdateCount As Long, ByVal Iterations As Long, ByVal CHUI As Long, ByVal definite As Integer, ByVal SyncMode As Integer) As Integer
Declare Function D2K_AO_SimuWriteChannel Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal NumChans As Integer, Buffer As Integer) As Integer
Declare Function D2K_AO_ContBufferSetup Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, Buffer As Any, ByVal WriteCount As Long, BufferId As Integer) As Integer
Declare Function D2K_AO_ContBufferReset Lib "D2K-Dask.dll" (ByVal CardNumber As Integer) As Integer
Declare Function D2K_AO_WriteChannel Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, ByVal Value As Integer) As Integer
Declare Function D2K_AO_VWriteChannel Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, ByVal Voltage As Double) As Integer
Declare Function D2K_AO_VoltScale Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Channel As Integer, ByVal Voltage As Double, binValue As Integer) As Integer
Declare Function D2K_AO_ContBufferComposeAll Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, ByVal UpdateCount As Long, ConBuffer As Any, Buffer As Any, ByVal fifoload As Byte) As Integer
Declare Function D2K_AO_ContBufferCompose Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, ByVal Channel As Integer, ByVal UpdateCount As Long, ConBuffer As Any, Buffer As Any, ByVal fifoload As Byte) As Integer
Declare Function D2K_AO_EventCallBack Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Mode As Integer, ByVal EventType As Integer, ByVal callbackAddr As Long) As Integer

'AO Group Functions
Declare Function D2K_AO_Group_Setup Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, ByVal wNumChans As Integer, wChans As Integer) As Integer
Declare Function D2K_AO_Group_Update Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, wBuffer As Integer) As Integer
Declare Function D2K_AO_Group_VUpdate Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, Voltage As Double) As Integer
Declare Function D2K_AO_Group_FIFOLoad Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, ByVal BufId As Integer, ByVal dwWriteCount As Long) As Integer
Declare Function D2K_AO_Group_FIFOLoad_2 Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, ByVal BufId As Integer, ByVal dwWriteCount As Long) As Integer
Declare Function D2K_AO_Group_WFM_StopConfig Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, ByVal stopSrc As Integer, ByVal stopMode As Integer) As Integer
Declare Function D2K_AO_Group_WFM_Start Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, ByVal fstBufIdOrNotUsed As Integer, ByVal sndBufId As Integer, ByVal UpdateCount As Long, ByVal Iterations As Long, ByVal CHUI As Long, ByVal definite As Integer) As Integer
Declare Function D2K_AO_Group_WFM_AsyncCheck Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, Stopped As Byte, WriteCnt As Long) As Integer
Declare Function D2K_AO_Group_WFM_AsyncClear Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal group As Integer, WriteCnt As Long, ByVal stop_mode As Integer) As Integer

'DI Functions
Declare Function D2K_DI_ReadPort Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Port As Integer, Value As Long) As Integer
Declare Function D2K_DI_ReadLine Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Port As Integer, ByVal Line As Integer, Value As Integer) As Integer

'DO Functions
Declare Function D2K_DO_WritePort Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Port As Integer, ByVal Value As Long) As Integer
Declare Function D2K_DO_WriteLine Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Port As Integer, ByVal Line As Integer, ByVal Value As Integer) As Integer
Declare Function D2K_DO_ReadLine Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Port As Integer, ByVal Line As Integer, Value As Integer) As Integer
Declare Function D2K_DO_ReadPort Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Port As Integer, Value As Long) As Integer

'DIO Functions
Declare Function D2K_DIO_PortConfig Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal Port As Integer, ByVal Direction As Integer) As Integer

'Counter Functions
Declare Function D2K_GCTR_Setup Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal wGCtr As Integer, ByVal wMode As Integer, ByVal SrcCtrl As Byte, ByVal PolCtrl As Byte, ByVal LReg1_Val As Integer, ByVal LReg2_Val As Integer) As Integer
Declare Function D2K_GCTR_Control Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal wGCtr As Integer, ByVal ParamID As Integer, ByVal Value As Integer) As Integer
Declare Function D2K_GCTR_Reset Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal wGCtr As Integer) As Integer
Declare Function D2K_GCTR_Read Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal wGCtr As Integer, pValue As Long) As Integer
Declare Function D2K_GCTR_Status Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal wGCtr As Integer, pValue As Integer) As Integer

'SSI Functions
Declare Function D2K_SSI_SourceConn Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal sigCode As Integer) As Integer
Declare Function D2K_SSI_SourceDisConn Lib "D2K-Dask.dll" (ByVal CardNumber As Integer, ByVal sigCode As Integer) As Integer
Declare Function D2K_SSI_SourceClear Lib "D2K-Dask.dll" (ByVal CardNumber As Integer) As Integer

'Calibration Functions
Declare Function DAQ2204_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, gain_err As Single, bioffset_err As Single, unioffset_err As Single, hg_bios_err As Single) As Integer
Declare Function DAQ2204_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, da0v_err As Single, da5v_err As Single) As Integer
Declare Function DAQ2205_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, gain_err As Single, bioffset_err As Single, unioffset_err As Single, hg_bios_err As Single) As Integer
Declare Function DAQ2205_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, da0v_err As Single, da5v_err As Single) As Integer
Declare Function DAQ2206_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, gain_err As Single, bioffset_err As Single, unioffset_err As Single, hg_bios_err As Single) As Integer
Declare Function DAQ2206_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, da0v_err As Single, da5v_err As Single) As Integer
Declare Function DAQ2010_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ2010_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ2005_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ2005_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ2006_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ2006_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ2016_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ2016_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ2208_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, gain_err As Single, bioffset_err As Single, unioffset_err As Single, hg_bios_err As Single) As Integer
Declare Function DAQ2213_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, gain_err As Single, bioffset_err As Single, unioffset_err As Single, hg_bios_err As Single) As Integer
Declare Function DAQ2214_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, gain_err As Single, bioffset_err As Single, unioffset_err As Single, hg_bios_err As Single) As Integer
Declare Function DAQ2214_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, da0v_err As Single, da5v_err As Single) As Integer
Declare Function DAQ250X_Acquire_AD_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function DAQ250X_Acquire_DA_Error Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal Channel As Integer, ByVal polarity As Integer, gain_err As Single, offset_err As Single) As Integer
Declare Function D2K_DB_Auto_Calibration_ALL Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer) As Integer
Declare Function D2K_EEPROM_CAL_Constant_Update Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal bank As Integer) As Integer
Declare Function D2K_Load_CAL_Data Lib "D2K-Dask.dll" (ByVal wCardNumber As Integer, ByVal bank As Integer) As Integer
