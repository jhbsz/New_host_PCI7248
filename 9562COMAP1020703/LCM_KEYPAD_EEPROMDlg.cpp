// LCM_KEYPAD_EEPROMDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Label.h"

#define LOAD_FW

#include "9525COMAP.h"
#include "LCM_KEYPAD_EEPROMDlg.h"
#include "9525RS232Lib.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define KEYPAD_ROW_NUM 0x05
#define KEYPAD_COL_NUM 0x06

PCHAR g_strLCMList[LCM_LIST_NUM] =
{
	"HD44780",
	"KS0108",
	"ST7920",
	"KGM0053",
	"PE12832"
};
BOOL HexString2Integer(CString HexStr, UINT *iValue);
void Integer2HexString(UINT iValue, CString *HexString);
//-----------------------------
//LCM ST7565 Functions start
//-----------------------------
#define LCDEnReg                    0xEA
#define LCDDataReg                0xEB
#define LCDRWReg                   0xEC
#define LCDRSReg                    0xED
#define LCDBLReg                    0xEE
#define LCDCSReg                    0xEF
#define CMDBusy                     0xF0
#define CMDDelay                    0xF1
#define LCDRepeatDataReg			0xF5




	
#define RESET                          0xFF
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
#define LCD_WRITE_DATA						0xD3

#define OFFSET_A 0
#define OFFSET_B 1
#define OFFSET_C 2
#define OFFSET_D 3
#define OFFSET_E 4
#define OFFSET_F 5
#define OFFSET_G 6
#define OFFSET_H 7
#define OFFSET_I 8
#define OFFSET_J 9
#define OFFSET_K 10
#define OFFSET_L 11
#define OFFSET_M 12
#define OFFSET_N 13
#define OFFSET_O 14
#define OFFSET_P 15
#define OFFSET_Q 16
#define OFFSET_R 17
#define OFFSET_S 18
#define OFFSET_T 19
#define OFFSET_U 20
#define OFFSET_V 21
#define OFFSET_W 22
#define OFFSET_X 23
#define OFFSET_Y 24
#define OFFSET_Z 25
#define OFFSET_0 26
#define OFFSET_1 27
#define OFFSET_2 28
#define OFFSET_3 29
#define OFFSET_4 30
#define OFFSET_5 31
#define OFFSET_6 32
#define OFFSET_7 33
#define OFFSET_8 34
#define OFFSET_9 35
//'='
#define OFFSET_EqualMark 36
//'*'
#define OFFSET_asterisk 37
//'.'
#define OFFSET_period  38
//','
#define OFFSET_comma 39
//' '
#define OFFSET_Space 40

//16 * 8 pixels per character (16 bytes each)
UCHAR aCharST7565KeyPadKey[] = 
{
	//'A'
	0x00,0x00,0xC0,0x38,0xE0,0x00,0x00,0x00,0x20,0x3C,0x23,0x02,0x02,0x27,0x38,0x20,
	//'B'
	0x08,0xF8,0x88,0x88,0x88,0x70,0x00,0x00,0x20,0x3F,0x20,0x20,0x20,0x11,0x0E,0x00,
	//'C'
	0xC0,0x30,0x08,0x08,0x08,0x08,0x38,0x00,0x07,0x18,0x20,0x20,0x20,0x10,0x08,0x00,
	//'D'
	0x08,0xF8,0x08,0x08,0x08,0x10,0xE0,0x00,0x20,0x3F,0x20,0x20,0x20,0x10,0x0F,0x00,
	//'E'
	0x08,0xF8,0x88,0x88,0xE8,0x08,0x10,0x00,0x20,0x3F,0x20,0x20,0x23,0x20,0x18,0x00,
	//'F'
	0x08,0xF8,0x88,0x88,0xE8,0x08,0x10,0x00,0x20,0x3F,0x20,0x00,0x03,0x00,0x00,0x00,
	//'G'
	0xC0,0x30,0x08,0x08,0x08,0x38,0x00,0x00,0x07,0x18,0x20,0x20,0x22,0x1E,0x02,0x00,
	//'H'
	0x08,0xF8,0x08,0x00,0x00,0x08,0xF8,0x08,0x20,0x3F,0x21,0x01,0x01,0x21,0x3F,0x20,
	//'I'
	0x00,0x08,0x08,0xF8,0x08,0x08,0x00,0x00,0x00,0x20,0x20,0x3F,0x20,0x20,0x00,0x00,
	//'J'
	0x00,0x00,0x08,0x08,0xF8,0x08,0x08,0x00,0xC0,0x80,0x80,0x80,0x7F,0x00,0x00,0x00,
	//'K'
	0x08,0xF8,0x88,0xC0,0x28,0x18,0x08,0x00,0x20,0x3F,0x20,0x01,0x26,0x38,0x20,0x00,
	//'L'
	0x08,0xF8,0x08,0x00,0x00,0x00,0x00,0x00,0x20,0x3F,0x20,0x20,0x20,0x20,0x30,0x00,
	//'M'
	0x08,0xF8,0xF8,0x00,0xF8,0xF8,0x08,0x00,0x20,0x3F,0x00,0x3F,0x00,0x3F,0x20,0x00,
	//'N'
	0x08,0xF8,0x30,0xC0,0x00,0x08,0xF8,0x08,0x20,0x3F,0x20,0x00,0x07,0x18,0x3F,0x00,
	//'O'
	0xE0,0x10,0x08,0x08,0x08,0x10,0xE0,0x00,0x0F,0x10,0x20,0x20,0x20,0x10,0x0F,0x00,
	//'P'
	0x08,0xF8,0x08,0x08,0x08,0x08,0xF0,0x00,0x20,0x3F,0x21,0x01,0x01,0x01,0x00,0x00,
	//'Q'
	0xE0,0x10,0x08,0x08,0x08,0x10,0xE0,0x00,0x0F,0x18,0x24,0x24,0x38,0x50,0x4F,0x00,
	//'R'
	0x08,0xF8,0x88,0x88,0x88,0x88,0x70,0x00,0x20,0x3F,0x20,0x00,0x03,0x0C,0x30,0x20,
	//'S'
	0x00,0x70,0x88,0x08,0x08,0x08,0x38,0x00,0x00,0x38,0x20,0x21,0x21,0x22,0x1C,0x00,
	//'T'
	0x18,0x08,0x08,0xF8,0x08,0x08,0x18,0x00,0x00,0x00,0x20,0x3F,0x20,0x00,0x00,0x00,
	//'U'
	0x08,0xF8,0x08,0x00,0x00,0x08,0xF8,0x08,0x00,0x1F,0x20,0x20,0x20,0x20,0x1F,0x00,
	//'V'
	0x08,0x78,0x88,0x00,0x00,0xC8,0x38,0x08,0x00,0x00,0x07,0x38,0x0E,0x01,0x00,0x00,
	//'W'
	0xF8,0x08,0x00,0xF8,0x00,0x08,0xF8,0x00,0x03,0x3C,0x07,0x00,0x07,0x3C,0x03,0x00,
	//'X'
	0x08,0x18,0x68,0x80,0x80,0x68,0x18,0x08,0x20,0x30,0x2C,0x03,0x03,0x2C,0x30,0x20,
	//'Y'
	0x08,0x38,0xC8,0x00,0xC8,0x38,0x08,0x00,0x00,0x00,0x20,0x3F,0x20,0x00,0x00,0x00,
	//'Z'
	0x10,0x08,0x08,0x08,0xC8,0x38,0x08,0x00,0x20,0x38,0x26,0x21,0x20,0x20,0x18,0x00,
	//'0'
	0x00,0xE0,0x10,0x08,0x08,0x10,0xE0,0x00,0x00,0x0F,0x10,0x20,0x20,0x10,0x0F,0x00,
	//'1'
	0x00,0x10,0x10,0xF8,0x00,0x00,0x00,0x00,0x00,0x20,0x20,0x3F,0x20,0x20,0x00,0x00,
	//'2'
	0x00,0x70,0x08,0x08,0x08,0x88,0x70,0x00,0x00,0x30,0x28,0x24,0x22,0x21,0x30,0x00,
	//'3'
	0x00,0x30,0x08,0x88,0x88,0x48,0x30,0x00,0x00,0x18,0x20,0x20,0x20,0x11,0x0E,0x00,
	//'4'
	0x00,0x00,0xC0,0x20,0x10,0xF8,0x00,0x00,0x00,0x07,0x04,0x24,0x24,0x3F,0x24,0x00,
	//'5'
	0x00,0xF8,0x08,0x88,0x88,0x08,0x08,0x00,0x00,0x19,0x21,0x20,0x20,0x11,0x0E,0x00,
	//'6'
	0x00,0xE0,0x10,0x88,0x88,0x18,0x00,0x00,0x00,0x0F,0x11,0x20,0x20,0x11,0x0E,0x00,
	//'7'
	0x00,0x38,0x08,0x08,0xC8,0x38,0x08,0x00,0x00,0x00,0x00,0x3F,0x00,0x00,0x00,0x00,
	//'8'
	0x00,0x70,0x88,0x08,0x08,0x88,0x70,0x00,0x00,0x1C,0x22,0x21,0x21,0x22,0x1C,0x00,
	//'9'
	0x00,0xE0,0x10,0x08,0x08,0x10,0xE0,0x00,0x00,0x00,0x31,0x22,0x22,0x11,0x0F,0x00,
	//'='
	0x40,0x40,0x40,0x40,0x40,0x40,0x40,0x00,0x04,0x04,0x04,0x04,0x04,0x04,0x04,0x00,
	//'*'
	0x40,0x40,0x80,0xF0,0x80,0x40,0x40,0x00,0x02,0x02,0x01,0x0F,0x01,0x02,0x02,0x00,
	//'.'
	0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x30,0x30,0x00,0x00,0x00,0x00,0x00,
	//','	
	0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x80,0xB0,0x70,0x00,0x00,0x00,0x00,0x00,
	//' '
	0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00
};
LONG lcd_wcmd(CSerial *m_ctrlCSerial, BYTE CMDcode)
{
	UCHAR aWriteCommandNocheck[] = 
	{
		LCDDataReg, CMDcode, LCDRSReg, LOW, LCDEnReg, HIGH, LCDRWReg, LOW, LCDEnReg, LOW,
	};
	return CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(aWriteCommandNocheck), (unsigned char *)aWriteCommandNocheck);
}
LONG lcd_wdata(CSerial *m_ctrlCSerial, BYTE data)
{
	UCHAR aWriteCommandNocheck[] = 
	{
		LCDDataReg, data, LCDRSReg, 1, LCDRWReg, LOW, LCDEnReg, HIGH, LCDEnReg, LOW,
	};
	return CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(aWriteCommandNocheck), (unsigned char *)aWriteCommandNocheck);
}

LONG CLCM_KEYPAD_EEPROMDlg::ClrLcd_ST7565(CSerial *m_ctrlCSerial)
{
	//display OFF
	if(lcd_wcmd(m_ctrlCSerial, 0XAE)!= 1)
	{
		MessageBox("Error! Failed to display off!", "Error", MB_OK);
		return 1;	
	}
    UINT i;    
    for (i=0;i<4;i++)
    {
		if(lcd_wcmd(m_ctrlCSerial, (0xb0+i))!= 1)//Page address set (0-3)
		{
			MessageBox("Error! Failed to set Page address!", "Error", MB_OK);
			return 1;
		}
		if(lcd_wcmd(m_ctrlCSerial, 0x40)!= 1)//Display start line set
		{
			MessageBox("Error! Failed to set start line!", "Error", MB_OK);
			return 1;
		}
		if(lcd_wcmd(m_ctrlCSerial, 0x10)!= 1)//Column address set(upper bit)
		{
			MessageBox("Error! Failed to set Column address!", "Error", MB_OK);
			return 1;
		}
		if(lcd_wcmd(m_ctrlCSerial, 0x00)!= 1)//Column address set(lower bit)
		{
			MessageBox("Error! Failed to set Column address!", "Error", MB_OK);
			return 1;
		}
		UCHAR aWriteCommandNocheck[] = 
		{
			LCDRepeatDataReg, 0, 132, LCDRSReg, 1, LCDRWReg, LOW, LCDEnReg, HIGH, LCDEnReg, LOW,
		};
		if((CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(aWriteCommandNocheck), (unsigned char *)aWriteCommandNocheck) != 1))
		{
			MessageBox("Error! Failed to Write data!", "Error", MB_OK);
			return 1;
		}        
	} 	
	//display ON
	if(lcd_wcmd(m_ctrlCSerial, 0XAF) != 1)
	{
		MessageBox("Error! Failed to display ON!", "Error", MB_OK);
		return 1;
	}
	return 0;
}
void CLCM_KEYPAD_EEPROMDlg::LcdKGM0053_init(CSerial *m_ctrlCSerial)
{
	UCHAR ResetCMDarray[] =	
	{
		LCDCSReg, 1
	};
	CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(ResetCMDarray), (unsigned char *)ResetCMDarray);
	Sleep(20);
	ResetCMDarray[1] = 0;
	CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(ResetCMDarray), (unsigned char *)ResetCMDarray);
	Sleep(20);
	ResetCMDarray[1] = 1;
	CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(ResetCMDarray), (unsigned char *)ResetCMDarray);
	Sleep(8);
	//Internal reset
	lcd_wcmd(m_ctrlCSerial, 0xE2);
	//LCD bias set
	lcd_wcmd(m_ctrlCSerial, 0XA2);
	//Sets the display RAM address SEG output correspondence (reverse)
	lcd_wcmd(m_ctrlCSerial, 0XA1);
	//COM output scan direction (reverse direction)
	lcd_wcmd(m_ctrlCSerial, 0XC8);
	//Booster ratio set
	lcd_wcmd(m_ctrlCSerial, 0XF8);
	//least significant 4 bits of the display RAM column address.
	lcd_wcmd(m_ctrlCSerial, 0X00);
	//Set the V0 output voltage electronic volume register
	lcd_wcmd(m_ctrlCSerial, 0X81);
	//V0 voltage
	lcd_wcmd(m_ctrlCSerial, 0X20);
	//Sleep mode
	lcd_wcmd(m_ctrlCSerial, 0XAC);
	//least significant 4 bits of the display RAM column address.
	lcd_wcmd(m_ctrlCSerial, 0X00);
	//Power control set
	lcd_wcmd(m_ctrlCSerial, 0X2C);
	//Power control set
	lcd_wcmd(m_ctrlCSerial, 0X2E);
	//Power control set
	lcd_wcmd(m_ctrlCSerial, 0X2F);
	//clear screen
	ClrLcd_ST7565(m_ctrlCSerial);
}
void CLCM_KEYPAD_EEPROMDlg::LcdPE12832_init(CSerial *m_ctrlCSerial)
{
	UCHAR ResetCMDarray[] =	
	{
		LCDCSReg, 1
	};
	CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(ResetCMDarray), (unsigned char *)ResetCMDarray);	
	Sleep(20);
	ResetCMDarray[1] = 0;
	CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(ResetCMDarray), (unsigned char *)ResetCMDarray);
	Sleep(20);
	ResetCMDarray[1] = 1;
	CmdLcmWriteData(m_ctrlCSerial, 0, sizeof(ResetCMDarray), (unsigned char *)ResetCMDarray);
	Sleep(8);
	//Internal reset
	lcd_wcmd(m_ctrlCSerial, 0xE2);
	//LCD bias set
	lcd_wcmd(m_ctrlCSerial, 0XA2);
	//Sets the display RAM address SEG output correspondence (reverse)
	lcd_wcmd(m_ctrlCSerial, 0XA0);
	//COM output scan direction (reverse direction)
	lcd_wcmd(m_ctrlCSerial, 0XC8);
	//Booster ratio set
	lcd_wcmd(m_ctrlCSerial, 0XF8);
	//least significant 4 bits of the display RAM column address.
	lcd_wcmd(m_ctrlCSerial, 0X00);
	//Set the V0 output voltage electronic volume register
	//lcd_wcmd(hCardHandle, 0X81);
	//V0 voltage
	lcd_wcmd(m_ctrlCSerial, 0X20);
	//Sleep mode
	lcd_wcmd(m_ctrlCSerial, 0XAC);
	//least significant 4 bits of the display RAM column address.
	lcd_wcmd(m_ctrlCSerial, 0X00);
	//Power control set
	lcd_wcmd(m_ctrlCSerial, 0X2C);
	//Power control set
	lcd_wcmd(m_ctrlCSerial, 0X2E);
	//Power control set
	lcd_wcmd(m_ctrlCSerial, 0X2F);
	//clear screen
	ClrLcd_ST7565(m_ctrlCSerial);
}
int ST7565CurentColumn = 0;
int ST7565CurrentRow = 1;
BOOL Flag_TheFirstKey = 0;
BYTE ST7565KeypadCharacterOFFSET;

LONG CLCM_KEYPAD_EEPROMDlg::DisplayST7565Data(CSerial *m_ctrlCSerial, BYTE CharacterOFFSET, BYTE Row, BYTE Column)
{
	BYTE PageAddress;//See ST7565 spec (0-3)
	switch(Row)
	{
	case 1:
		PageAddress = 0;
		break;
	case 2:
		PageAddress = 2;
		break;
	default:
		PageAddress = 0;//error
		break;
	}
	int Offset;
	Offset = CharacterOFFSET * 16;

	if(lcd_wcmd(m_ctrlCSerial, (0xb0 + PageAddress))!= 1)//Page address set (0-3)
	{
		MessageBox("Error! Failed to set Page address!", "Error", MB_OK);
		return 1;
	}
	if(lcd_wcmd(m_ctrlCSerial, 0x40)!= 1)//Display start line set
	{
		MessageBox("Error! Failed to set display start line!", "Error", MB_OK);
		return 1;
	}
	if(lcd_wcmd(m_ctrlCSerial, (0x10 + (Column >> 4)))!= 1)//Column address set(upper bit)
	{
		MessageBox("Error! Failed to set column address!", "Error", MB_OK);
		return 1;
	}
	if(lcd_wcmd(m_ctrlCSerial, (0x00 + (Column & 0x0F)))!= 1)//Column address set(lower bit)
	{
		MessageBox("Error! Failed to set column address!", "Error", MB_OK);
		return 1;
	}
	int i;
	for(i=0;i<7;i++)
	{
		if(lcd_wdata(m_ctrlCSerial, aCharST7565KeyPadKey[Offset + i])!= 1)
		{
			MessageBox("Error! Failed to write data!", "Error", MB_OK);
			return 1;
		}
	}	
	PageAddress++;
	if(lcd_wcmd(m_ctrlCSerial, (0xb0 + PageAddress))!= 1)//Page address set (0-3)
	{
		MessageBox("Error! Failed to set page address!", "Error", MB_OK);
		return 1;
	}
	if(lcd_wcmd(m_ctrlCSerial, 0x40)!= 1)//Display start line set
	{
		MessageBox("Error! Failed to set display start line!", "Error", MB_OK);
		return 1;
	}
	if(lcd_wcmd(m_ctrlCSerial, (0x10 + (Column >> 4)))!= 1)//Column address set(upper bit)
	{
		MessageBox("Error! Failed to set column address!", "Error", MB_OK);
		return 1;
	}
	if(lcd_wcmd(m_ctrlCSerial, (0x00 + (Column & 0x0F)))!= 1)//Column address set(lower bit)
	{
		MessageBox("Error! Failed to set column address!", "Error", MB_OK);
		return 1;
	}
	for(i=0;i<7;i++)
	{
		if(lcd_wdata(m_ctrlCSerial, aCharST7565KeyPadKey[Offset + i + 8])!= 1)
		{
			MessageBox("Error! Failed to write data!", "Error", MB_OK);
			return 1;
		}
	}
	return 0;
}
BYTE CLCM_KEYPAD_EEPROMDlg::ByteASCToOffset(BYTE ASCdata)
{
	if(('A' <= ASCdata) && (ASCdata <= 'Z'))
	{
		return (ASCdata - 'A' + OFFSET_A);
	}
	else if(('0' <= ASCdata) && (ASCdata <= '9'))
	{
		return (ASCdata - '0' + OFFSET_0);	
	}
	else if(ASCdata == '=')
	{
		return OFFSET_EqualMark;
	}
	else if(ASCdata == '*')
	{
		return OFFSET_asterisk;
	}
	else if(ASCdata == '.')
	{
		return OFFSET_period;
	}
	else if(ASCdata == ',')
	{
		return OFFSET_comma;
	}
	else if(ASCdata == ' ')
	{
		return OFFSET_Space;
	}
	else
	{
		MessageBox("there is illegal character in the text");
		return 0xff;
	}
}

//-----------------------------
//LCM ST7565 Functions end
//-----------------------------


#define HD44780_BUSY                    0x80
#define KS0108_BUSY                     0x90
#define CLEAR_DISPLAY                   0x01
#define ADDR_RESET                      0x10
#define RAM_ADDR_RIGHT_INCREASE         0x06
#define ST7920_DISPLAY_ON               0x0C
#define KS0108_DISPLAY_OFF              0x3E
#define KS0108_DISPLAY_ON               0x3F

#define SET_CURSOR_EN                   0x80
#define ST7920_SET_LINE_EN              0x80
#define KS0108_SET_LINE_EN              0xB8
#define HD44780_LINE_SELECT             0x40
#define ST7920_LINE_SELECT              0x10
#define SET_COLUMN_EN                   0x40
#define SET_COLUMN_MASK                 0x3F
#define SET_START_LINE_EN               0xC0

#define HD44780_COL_CHAR_NUM            0x10
#define ST7920_COL_CHAR_NUM             0x10
#define KS0108_COL_CHAR_NUM             0x08
#define HD44780_ROW_CHAR_NUM            0x02
#define KS0108_ROW_CHAR_NUM             0x04
#define KS0108_SUB_COL_NUM              0x10
#define GRAPH_MODE_CLEAR_FULL_SCREEN        0xF3
UCHAR	abDevDesc[18] = {0x12,0x01, 0x10, 0x01, 0x00, 0x00, 0x00, 0x08, 0x8F, 0x05, 0x20, 0x95, 0xFF, 0xFF, 0x01, 0x02, 0x00, 0x01};
/////////////////////////////////////////////////////////////////////////////
// CLCM_KEYPAD_EEPROMDlg dialog


CLCM_KEYPAD_EEPROMDlg::CLCM_KEYPAD_EEPROMDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CLCM_KEYPAD_EEPROMDlg::IDD, pParent)
{
	EnableAutomation();

	//{{AFX_DATA_INIT(CLCM_KEYPAD_EEPROMDlg)
	m_LCMPosX = 0;
	m_LCMPosY = 0;
	m_UsbVid = _T("");
	m_UsbPid = _T("");
	m_VidString = _T("");
	m_PidString = _T("");
	m_IsSupportSn = FALSE;
	m_SnString = _T("");
	m_KeypadValue = _T("");
	m_RangeX = _T("");
	m_DisplayLength = 0;
	m_DisplayHigh = 0;
	m_RangeY = _T("");
	m_DisplayString = _T("");
	//}}AFX_DATA_INIT
}


void CLCM_KEYPAD_EEPROMDlg::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CDialog::OnFinalRelease();
}

void CLCM_KEYPAD_EEPROMDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CLCM_KEYPAD_EEPROMDlg)
	DDX_Control(pDX, IDC_ABOUT, m_About);
	DDX_Control(pDX, IDC_CLOSE, m_Close);
	DDX_Control(pDX, IDC_WRITE_EEPROM, m_WriteEeprom);
	DDX_Control(pDX, IDC_SAVE_FILE, m_SaveFile);
	DDX_Control(pDX, IDC_OPEN_FILE, m_OpenFile);
	DDX_Control(pDX, IDC_LCM_SELECT, m_ctlLCMList);
	DDX_Text(pDX, IDC_DISPLAY_POS_X, m_LCMPosX);
	DDX_Text(pDX, IDC_DISPLAY_POS_Y, m_LCMPosY);
	DDX_Text(pDX, IDC_VID, m_UsbVid);
	DDX_Text(pDX, IDC_PID, m_UsbPid);
	DDX_Text(pDX, IDC_VID_STRING, m_VidString);
	DDX_Text(pDX, IDC_PID_STRING, m_PidString);
	DDX_Check(pDX, IDC_CHECK_SUPPORT_SN, m_IsSupportSn);
	DDX_Text(pDX, IDC_SN, m_SnString);
	DDX_Text(pDX, IDC_KEYPAD_VALUE, m_KeypadValue);
	DDX_Text(pDX, IDC_RANGE_X, m_RangeX);
	DDX_Text(pDX, IDC_DISPLAY_LENGTH, m_DisplayLength);
	DDX_Text(pDX, IDC_DISPLAY_HIGH, m_DisplayHigh);
	DDX_Text(pDX, IDC_RANGE_Y, m_RangeY);
	DDX_Text(pDX, IDC_DISPLAY_STRING, m_DisplayString);
	DDX_Control(pDX, IDC_RESULT, m_Result);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CLCM_KEYPAD_EEPROMDlg, CDialog)
	//{{AFX_MSG_MAP(CLCM_KEYPAD_EEPROMDlg)
	ON_BN_CLICKED(IDC_HELP1, OnHelp1)
	ON_BN_CLICKED(IDC_ABOUT, OnAbout)
	ON_BN_CLICKED(IDC_CLOSE, OnClose)
	ON_BN_CLICKED(IDC_OPEN_FILE, OnOpenFile)
	ON_BN_CLICKED(IDC_SAVE_FILE, OnSaveFile)
	ON_BN_CLICKED(IDC_WRITE_EEPROM, OnWriteEeprom)
	ON_BN_CLICKED(IDC_TEST_LCM_KEYPAD, OnTestLcmKeypad)
	ON_BN_CLICKED(IDC_DISPLAY_GRAPH, OnDisplayGraph)
	ON_CBN_SELCHANGE(IDC_LCM_SELECT, OnSelchangeLcmSelect)
	ON_BN_CLICKED(IDC_DISPLAY_TEXT, OnDisplayText)
	ON_BN_CLICKED(IDC_DISPLAY_Clear, OnDISPLAYClear)
	ON_BN_CLICKED(IDC_BACKLIGHT_CHANGE, OnBacklightChange)
	ON_BN_CLICKED(IDC_UPDATE_FW, OnUpdateFw)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CLCM_KEYPAD_EEPROMDlg, CDialog)
	//{{AFX_DISPATCH_MAP(CLCM_KEYPAD_EEPROMDlg)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ILCM_KEYPAD_EEPROMDlg to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {87603988-29D9-4CE1-BD4E-B7BD217C064B}
static const IID IID_ILCM_KEYPAD_EEPROMDlg =
{ 0x87603988, 0x29d9, 0x4ce1, { 0xbd, 0x4e, 0xb7, 0xbd, 0x21, 0x7c, 0x6, 0x4b } };

BEGIN_INTERFACE_MAP(CLCM_KEYPAD_EEPROMDlg, CDialog)
	INTERFACE_PART(CLCM_KEYPAD_EEPROMDlg, IID_ILCM_KEYPAD_EEPROMDlg, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CLCM_KEYPAD_EEPROMDlg message handlers

void CLCM_KEYPAD_EEPROMDlg::OnHelp1() 
{
	// TODO: Add your control notification handler code here
	CString   appPath;
	//LPTSTR pAppPath;
	//GetCurrentDirectory(1000,pAppPath);
	TCHAR   exeFullPath[256];       
	GetModuleFileName(NULL,exeFullPath,256);     
	appPath=(CString)exeFullPath; 
	// TODO: Add your control notification handler code here
	ShellExecute(NULL,"open","AU9525DemoHelp.chm",appPath,NULL,SW_SHOWNORMAL);	
}

void CLCM_KEYPAD_EEPROMDlg::OnAbout() 
{
	// TODO: Add your control notification handler code here
//		UpdateData(TRUE);
//		m_LCMPosY++;
//		UpdateData(FALSE);
}

void CLCM_KEYPAD_EEPROMDlg::OnClose() 
{
	// TODO: Add your control notification handler code here
	EndDialog(IDOK);
}

BOOL CLCM_KEYPAD_EEPROMDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	for(int idx=0; idx<LCM_LIST_NUM; idx++)
	{
		m_ctlLCMList.AddString(g_strLCMList[idx]);
	}
	m_ctlLCMList.SetCurSel(0);

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
//	SetIcon(m_hIcon, TRUE);			// Set big icon
//	SetIcon(m_hIcon, FALSE);		// Set small icon

	m_UsbVid = _T("058F");

	m_UsbPid = _T("9525");

	m_VidString = _T("Alcor Micro, Corp.");
	m_PidString = _T("USB Smart Card Reader");
	
	m_SnString = _T(""); 
	              
	m_OpenFile.EnableWindow(TRUE);
	m_SaveFile.EnableWindow(TRUE);   
	m_WriteEeprom.EnableWindow(TRUE);
	m_Close.EnableWindow(TRUE);      
	m_About.EnableWindow(TRUE);//TRUE
	              
	m_IsSupportSn = FALSE;//TRUE
	
	((CEdit *)GetDlgItem(IDC_VID))->SetLimitText(4);
	//((CEdit *)GetDlgItem(IDC_FW_START_POS))->SetLimitText(4);
	((CEdit *)GetDlgItem(IDC_PID))->SetLimitText(4);
	//((CEdit *)GetDlgItem(IDC_FW_END_POS))->SetLimitText(4);
	((CEdit *)GetDlgItem(IDC_VID_STRING))->SetLimitText(30);
	((CEdit *)GetDlgItem(IDC_PID_STRING))->SetLimitText(30);
	((CEdit *)GetDlgItem(IDC_SN))->SetLimitText(30);
	
	m_Prog = ((CProgressCtrl *)GetDlgItem(IDC_PROGRESS));
	m_Prog->SetPos(0);

	m_Status = ((CStatic *)GetDlgItem(IDC_STATUS));
	m_Status->SetWindowText("");

	m_Result.SetWindowText("");
	//GetDlgItem(IDC_STATUS)
		
	// TODO: Add extra initialization here
	m_nLCMIndex = m_ctlLCMList.GetCurSel();
	m_EepromModel = new EepromModel(m_UsbVid, m_UsbPid, m_VidString, m_PidString, m_SnString, m_IsSupportSn);
	UpdateData(FALSE);
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CLCM_KEYPAD_EEPROMDlg::OnOpenFile() 
{
	// TODO: Add your control notification handler code here
	CString strFilter;
	CFile	BinFile;
	int	hResult;
//	int iActual;
	strFilter = "Binary File(*.bin)|*.bin||"; 

	CFileDialog dlg(TRUE,NULL,NULL,OFN_FILEMUSTEXIST,strFilter); 
	hResult = (int)dlg.DoModal(); 
	if (hResult != IDOK) 
	{ 
		return; 
	} 
	
	CFile myFile;
	CFileException fileException;
	if ( !myFile.Open( dlg.GetFileName(), 
			  CFile::modeReadWrite, &fileException ) )
	{
		MessageBox("Can't open file", "Open file", MB_OK );
		return;
	}

	m_iBinDataLen = 0;
	m_iBinDataLen = myFile.Read( m_abBinData, sizeof( m_abBinData ) );
	
	if( m_iBinDataLen )
	{
		GetDlgItem(IDC_SAVE_FILE)->EnableWindow(TRUE);
	}

	if( m_EepromModel->Binary2EepromModel(m_abBinData, m_iBinDataLen) == FALSE )
	{
		MessageBox("Error! The file format is incorrect or is too large for the EEPROM!", " Verify File ", MB_OK );
		return;
	}
	UpdateDataFromEepromModel(m_EepromModel);
	UpdateData(FALSE);
	return;	
}

BOOL EepromModel::Binary2EepromModel(UCHAR *pBinData, UINT iBinLen)
{
	UCHAR bAddrStart;
	UCHAR bAddrEnd;
	UINT	iTmp;
	UCHAR i;
	CString Vid;
	CString Pid;
	CString VidString;
	CString PidString;
	CString SnString;
	BOOL    IsSupportSn; //BOOL
	
	if( iBinLen < 12 )
	{
		return FALSE;//FALSE
	}
		
	if( pBinData[0] != 0x99 || pBinData[1] != 0x07 )
	{
		return FALSE;//FALSE
	}
	
	/* Device  Descriptors */
	bAddrStart = pBinData[2];
	bAddrEnd = pBinData[3];
	if( (bAddrEnd - bAddrStart) != 18 )
	{
		return FALSE;//FALSE
	}
	
	/* Get VID/PID */
	iTmp = pBinData[bAddrStart + 8] | ((UINT)pBinData[bAddrStart + 9]<<8);
	Integer2HexString( iTmp, &Vid );
	
	iTmp = pBinData[bAddrStart + 10] | ((UINT)pBinData[bAddrStart + 11]<<8);
	Integer2HexString( iTmp, &Pid );
	
	/* Get Serial number supported? */
	if( pBinData[bAddrStart + 16] == 0 )
		IsSupportSn = FALSE;
	else
		IsSupportSn = TRUE;//FALSE
	

	/* String 0  Descriptors */
	bAddrStart = pBinData[4];
	bAddrEnd = pBinData[5];
	if( (bAddrEnd - bAddrStart) != 4 )
	{
		return FALSE;//FALSE
	}
	if( pBinData[bAddrStart] != 0x04 || 
		pBinData[bAddrStart+1] != 0x03 || 
		pBinData[bAddrStart+2] != 0x09 || 
		pBinData[bAddrStart+3] != 0x04 )
	{
		return FALSE;//FALSE
	}
	
	/* String 1  Descriptors */
	bAddrStart = pBinData[6];
	bAddrEnd = pBinData[7];
	if( bAddrStart >= bAddrEnd )
	{
		return FALSE;//FALSE
	}
	if( pBinData[bAddrStart] != (bAddrEnd-bAddrStart) )
	{
		return FALSE;//FALSE
	}
	VidString = "";
	for( i = bAddrStart+2; i < bAddrEnd ;i+=2 )
	{
		VidString = VidString + (CHAR)pBinData[i];
	}
	
	/* String 2  Descriptors */
	bAddrStart = pBinData[8];
	bAddrEnd = pBinData[9];
	if( bAddrStart >= bAddrEnd )
	{
		return FALSE;//FALSE
	}
	if( pBinData[bAddrStart] != (bAddrEnd-bAddrStart) )
	{
		return FALSE;//FALSE
	}
	PidString = "";
	for( i = bAddrStart+2; i < bAddrEnd ;i+=2 )
	{
		PidString = PidString + (CHAR)pBinData[i];
	}
	
	if( IsSupportSn == FALSE) //FALSE
	{
		SnString = "";
		goto OK_LEAVE;
	}	

	/* String 3  Descriptors */
	bAddrStart = pBinData[10];
	bAddrEnd = pBinData[11];
	if( bAddrStart >= bAddrEnd )
	{
		return FALSE;//FALSE
	}
	if( pBinData[bAddrStart] != (bAddrEnd-bAddrStart) )
	{
		return FALSE;//FALSE
	}
	SnString = "";
	for( i = bAddrStart+2; i < bAddrEnd ;i+=2 )
	{
		SnString = SnString + (CHAR)pBinData[i];
	}	
OK_LEAVE:
	/* LCD */
	m_LCDType = pBinData[12];
	m_LCDAddr = pBinData[13];
	m_LCDLen  = ((UINT)pBinData[14]<<8)+pBinData[15];

	if(m_LCDLen>=2000)//超出EEPROM容量
	{
//		MessageBox("Picture File is too large to copy in EEPROM");	
		return FALSE;
	}

	if( (m_LCDAddr != 0) && (m_LCDLen != 0) )
	{
		for(iTmp=0;iTmp<m_LCDLen;iTmp++)
		{
			m_LCDData[iTmp] = pBinData[m_LCDAddr+iTmp];
		}
	}
	
	m_Vid = Vid;
	m_Pid = Pid;
	m_VidString = VidString;
	m_PidString = PidString;
	m_SnString = SnString;
	m_IsSnSupported = IsSupportSn; //BOOL

	return TRUE;//TRUE
}

BOOL EepromModel::EepromModel2Bin(UCHAR *pBinData, UINT iBufLen, UINT *iBinLen)
{
	UCHAR bAddrStart;
	UCHAR bAddrEnd;
	UINT iTmp;
	UCHAR i;
	UCHAR	bHeaderOffSet;


	*iBinLen = 0;
	if( m_Vid.GetLength() != 4
		|| m_Pid.GetLength() != 4
		|| m_VidString.GetLength() == 0
		|| m_PidString.GetLength() == 0 )
	{
		return FALSE;//FALSE
	}
	if( m_IsSnSupported == TRUE && m_SnString.GetLength() == 0 ) //TRUE
	{
		return FALSE;//FALSE
	}		
	
	bHeaderOffSet = 0;
	/* Signature */
	pBinData[bHeaderOffSet++] = 0x99;
	pBinData[bHeaderOffSet++] = 0x07;
	
	/* Device Descriptor */
	bAddrStart = 16; //12+4 for LCM
	bAddrEnd = bAddrStart + 0x12;// The length of device descriptor is fixed. 12 bytes.
	pBinData[bHeaderOffSet++] = bAddrStart;
	pBinData[bHeaderOffSet++] = bAddrEnd;
	memcpy( &pBinData[bAddrStart], abDevDesc, 0x12 );
	// VID
	if( HexString2Integer(m_Vid, &iTmp) != TRUE)//TRUE
	{
		return FALSE;//FALSE
	}
	pBinData[bAddrStart + 8] = (UCHAR)iTmp;
	pBinData[bAddrStart + 9] = iTmp >> 8;
	// PID
	if( HexString2Integer(m_Pid, &iTmp) != TRUE)//TRUE
	{
		return FALSE;//FALSE
	}
	pBinData[bAddrStart + 10] = (UCHAR)iTmp;
	pBinData[bAddrStart + 11] = iTmp >> 8;
	// Serial Number
	if( m_IsSnSupported == TRUE )//TRUE
		pBinData[bAddrStart + 0x10] = 3;
	else
		pBinData[bAddrStart + 0x10] = 0;
	
	/* String 0 Descriptor */
	bAddrStart = bAddrEnd;
	bAddrEnd += 4;
	pBinData[bHeaderOffSet++] = bAddrStart;
	pBinData[bHeaderOffSet++] = bAddrEnd;
	pBinData[bAddrStart] = 0x04;
	pBinData[bAddrStart + 1] = 0x03;
	pBinData[bAddrStart + 2] = 0x09;
	pBinData[bAddrStart + 3] = 0x04;

	/* String 1 Descriptor */
	UCHAR j;
	for( j = 1; j <= 3 ; j++ )
	{
		CString *CurString;
		
		if( j == 1 )
			CurString = &m_VidString;
		else if( j == 2 )
			CurString = &m_PidString;
		else
		{
			if( m_IsSnSupported == TRUE )
				CurString = &m_SnString;
			else
				break;
		} 
		
		//if( ((2*CurString->GetLength()+2) % 8 ) == 0 )
		//{
		//	MessageBox(NULL,"Sorry! The string length is not acceptable by the reader", "Error", MB_OK);
		//	return  FALSE;//FALSE;
		//} 
		bAddrStart = bAddrEnd;
		bAddrEnd += (2 + 2* CurString->GetLength());
		if( bAddrEnd > iBufLen ) return FALSE;//FALSE
		pBinData[bHeaderOffSet++] = bAddrStart;
		pBinData[bHeaderOffSet++] = bAddrEnd;
		pBinData[bAddrStart] = 2+2* CurString->GetLength();
		pBinData[bAddrStart + 1] = 0x03;
		for( i = 0; i < (CurString->GetLength()); i++ )
		{
			pBinData[bAddrStart + 2*i + 2] = CurString->GetAt(i);
			pBinData[bAddrStart + 2*i + 3] = 0x00;
		}
	}

//	UCHAR  x,y,bLastX = 0xFF,bLastY = 0xFF,bLastScreen = 0xFF, bScreenChangeFlag;
//	UINT   wLCDAddr, wLCDLen;
 
	/* LCD */
	pBinData[12] = m_LCDType;
	pBinData[13] = m_LCDAddr;
	pBinData[14] = m_LCDLen>>8;
	pBinData[15] = m_LCDLen;
	if( (m_LCDAddr != 0) && (m_LCDLen != 0) )
	{
		if( (m_LCDLen + m_LCDAddr ) > iBufLen ) return FALSE;
/*		if( m_LCDType == 0x01 ) //KS0108
		{
			wLCDAddr = m_LCDAddr;
			wLCDLen = 0;
			bScreenChangeFlag = 0;
			for(iTmp=0;iTmp<m_LCDLen;iTmp+=3)
			{

    			y = m_LCDData[iTmp+1]&0x3F|0x40;		//  col.and.0x3f.or.setx
				x = m_LCDData[iTmp]&0x07|0xb8;		//  row.and.0x07.or.sety
				if(bLastScreen!=(m_LCDData[iTmp+1]&0xc0))
				{
					bLastScreen = m_LCDData[iTmp+1]&0xc0;
					bScreenChangeFlag = 1;
					switch (bLastScreen)		//  col.and.0xC0
					{			
					case 0:	  
						pBinData[wLCDAddr++] =0xEF; //LCDCSReg = 1;// left screen
						pBinData[wLCDAddr++] =0x01;
						wLCDLen += 2;
						break;
					case 0x40:
						pBinData[wLCDAddr++] =0xEF; //LCDCSReg = 0;// right screen
						pBinData[wLCDAddr++] =0x00;
						wLCDLen += 2;
						break;
					}
				}
				if(bLastX!=x||bScreenChangeFlag)
				{
					bLastX = x;
					pBinData[wLCDAddr++] =0xF2; //Lcd1602_Write_Command(x);
					pBinData[wLCDAddr++] =x;
					wLCDLen += 2;
					bScreenChangeFlag = 0;
				}

				if((bLastY+1)!=y)
				{
					pBinData[wLCDAddr++] =0xF2; //Lcd1602_Write_Command(y);
					pBinData[wLCDAddr++] =y;
					wLCDLen += 2;
				}
				bLastY=y;
				pBinData[wLCDAddr++] =0xF1; //Lcd1602_Write_Data();
				pBinData[wLCDAddr++] =m_LCDData[iTmp+2];
				wLCDLen += 2;
			}
			m_LCDLen = wLCDLen;//(m_LCDLen/3)*8;
			pBinData[14] = m_LCDLen>>8;
			pBinData[15] = m_LCDLen;
		}
		else*/
		{
			for(iTmp=0;iTmp<m_LCDLen;iTmp++)
			{
				pBinData[m_LCDAddr+iTmp] = m_LCDData[iTmp];
			}
		}
		*iBinLen = m_LCDAddr+m_LCDLen;
	}
	else
	*iBinLen = bAddrEnd;
	return TRUE;//TRUE
}

BOOL HexString2Integer(CString HexStr, UINT *iValue)  //BOOL
{
	UINT iTmpValue;

	iTmpValue = 0;
	if( HexStr.GetLength() != 4 )
	{
		MessageBox(NULL, "Error! Buffer size not equals to 4!", "HexString2Integer()", MB_OK);
		return FALSE;
	}
	HexStr.MakeUpper();
	
	UCHAR i;
	for( i = 0; i < 4; i++ )
	{
		iTmpValue <<= 4;
		if( HexStr.GetAt(i) >= '0' && HexStr.GetAt(i) <= '9' )
		{
			iTmpValue |= ( HexStr.GetAt(i) - '0' );
		}
		else if( HexStr.GetAt(i) >= 'A' && HexStr.GetAt(i) <= 'F' )
		{
			iTmpValue |= ( HexStr.GetAt(i) - 'A' + 10 );
		}
		else
		{
			return FALSE;//FALSE
		}

	}

	*iValue = iTmpValue;
	return TRUE;//TRUE
}

void Integer2HexString(UINT iValue, CString *HexString)
{
	CString TmpHexString;
	UCHAR i;
	char bTmp;

	TmpHexString = "";
	for( i = 0; i < 4; i++ )
	{
		bTmp = (CHAR) (iValue % 16);
		iValue = iValue/16;
		if( bTmp >= 10 )
		{
			TmpHexString = (CHAR)('A'+ bTmp - 10 ) + TmpHexString;
		}
		else
		{
			TmpHexString = (CHAR)('0'+ bTmp ) + TmpHexString;
		}
	}
	
	*HexString = TmpHexString;
	return;
	
}

void CLCM_KEYPAD_EEPROMDlg::UpdateDataFromEepromModel(EepromModel *pEepromModel)
{
	m_UsbVid 	= pEepromModel->GetVid();	
	m_UsbPid 	= pEepromModel->GetPid();
	m_PidString = pEepromModel->GetPidString();
	m_VidString = pEepromModel->GetVidString();
	m_SnString 	= pEepromModel->GetSnString();
	m_IsSupportSn = pEepromModel->GetIsSnSupported();
	if( m_IsSupportSn )
		GetDlgItem(IDC_SN)->EnableWindow(TRUE);
	else
		GetDlgItem(IDC_SN)->EnableWindow(FALSE);//FALSE	
	return;
}

void CLCM_KEYPAD_EEPROMDlg::OnSaveFile() 
{
	// TODO: Add your control notification handler code here
	CString strFilter;
	
	UpdateData(TRUE); //TRUE
	UpdateDataToEepromModel(m_EepromModel);
	
	if( m_EepromModel->EepromModel2Bin(m_abBinData, sizeof(m_abBinData), &m_iBinDataLen) == FALSE ) //FALSE
	{
		MessageBox("The data is illegal!", "Error", MB_OK);
		return;
	}
	
//	int iActual;
	strFilter = "Binary File(*.bin)|*.bin||"; 
	CFileDialog opendlg (FALSE,"bin","*.bin",OFN_FILEMUSTEXIST|OFN_HIDEREADONLY,strFilter, this);

	if(opendlg.DoModal() != IDOK)
	{
		return;
	}

	CFile hFile;
	CFileException e;
	hFile.Open( opendlg.GetFileName(), CFile::modeCreate | CFile::modeWrite, &e );
	hFile.Write(m_abBinData,m_iBinDataLen);
	hFile.Close();
	
	return;	
}

void CLCM_KEYPAD_EEPROMDlg::UpdateDataToEepromModel(EepromModel *pEepromModel)
{
	pEepromModel->SetVid(m_UsbVid);	
	pEepromModel->SetPid(m_UsbPid);
	pEepromModel->SetPidString(m_PidString);
	pEepromModel->SetVidString(m_VidString);
	pEepromModel->SetSnString(m_SnString);
	pEepromModel->SetSnEnabled(m_IsSupportSn);
	return;
}

void CLCM_KEYPAD_EEPROMDlg::OnWriteEeprom() 
{
	COLORREF		ResultColor;
	// TODO: Add your control notification handler code here
	ResultColor = RGB(0,0,255);
	m_Result.SetTextColor(ResultColor);
	m_Result.SetWindowText("");
	m_Prog->SetPos(0);
	m_Status->SetWindowText("");	

	UpdateData(TRUE); //TRUE
	UpdateDataToEepromModel(m_EepromModel);

	if( m_EepromModel->EepromModel2Bin(m_abBinData, sizeof(m_abBinData), &m_iBinDataLen) == FALSE ) //FALSE
	{
		MessageBox("The data is illegal!", "Error", MB_OK);
		return;
	}

	UINT iCnt;
	/* Write to Eeprom */
	ResultColor = RGB(0,0,255);
	m_Result.SetTextColor(ResultColor);//zmz
	m_Result.SetWindowText("Wait ....");
	m_Status->SetWindowText("Write Eeprom");
	m_Prog->SetPos(0);

	for( iCnt = 0; iCnt < m_iBinDataLen; iCnt++ )
	{
		//Sleep(10);
		if( EepromCmdWrite(&CSerial9525, 0, iCnt , 1, m_abBinData + iCnt) != 1 )
		{
			ResultColor = RGB(255,0,0);
			m_Result.SetTextColor(ResultColor);
			m_Result.SetWindowText("FAILED !");
			MessageBox("Error! Failed to write to eeprom!", "Error", MB_OK);
			return;
		}
//		m_Result_SMC.SetFontSize(16);
//		m_Result_SMC.SetWindowText(m_Result);
		if( iCnt & 0x0C )
			m_Result.SetWindowText("Wait ....");
		else
			m_Result.SetWindowText("");
		m_Prog->SetPos(100*iCnt/m_iBinDataLen);
	}
	Sleep(10);
	/* Verify */
	ResultColor = RGB(0,0,255);
	m_Result.SetTextColor(ResultColor);
	m_Result.SetWindowText("Wait ....");

	m_Status->SetWindowText("Read/Verify Eeprom");
	m_Prog->SetPos(0);

	for( iCnt = 0; iCnt < m_iBinDataLen; iCnt+=8 )
	{
		UCHAR bData[8];
		UCHAR bNum;
		UCHAR i;
		ULONG lReturnLen;
		if( (m_iBinDataLen - iCnt) < 8 )		
			bNum = m_iBinDataLen - iCnt;
		else
			bNum = 8;

		if( EepromCmdRead(&CSerial9525, 0, iCnt, bNum, &bData, &lReturnLen) != 1 )
		{
			ResultColor = RGB(255,0,0);
			m_Result.SetTextColor(ResultColor);
			m_Result.SetWindowText("FAILED !");

			MessageBox("Error! Failed to read from eeprom!", "Error", MB_OK);
			return;			
		}
		for(i=0;i<bNum;i++)
		{
			if( bData[i] != m_abBinData[iCnt+i] )
			{
				ResultColor = RGB(255,0,0);
				m_Result.SetTextColor(ResultColor);
				m_Result.SetWindowText("FAILED !");

				MessageBox("Error! W/R compare failed!", "Error", MB_OK);
				return;			
			}
		}
		if( iCnt*8 & 0x0C )
			m_Result.SetWindowText("Wait ....");
		else
			m_Result.SetWindowText("");

		m_Prog->SetPos(100*iCnt*8/m_iBinDataLen);
	}
	m_Result.SetWindowText("success");		
	ResultColor = RGB(0,255,0);
	m_Result.SetTextColor(ResultColor);

	MessageBox("SUCCESS","DemoTool",MB_OK|MB_ICONINFORMATION);
	//m_Result.SetWindowText("SUCCESS");

	return;
}


unsigned char ReadCharData(FILE *fR)
{
	int index;
	size_t n;
	unsigned char returnVal,val[2];

	index = 0;
	n = 1;
	while((index < 2) && n)
	{
		n = fread(&val[index], sizeof(char), 1, fR);
		if(('0'<=val[index]&&'9'>=val[index])||('a'<=val[index]&&'f'>=val[index])||('A'<=val[index]&&'F'>=val[index]))
		{
			index++;
		}
	}

	returnVal = 0;
	if('0'<=val[0]&&'9'>=val[0])
	{
		returnVal |= (val[0] - '0') << 4;
	}
	else if('a'<=val[0]&&'f'>=val[0])
	{
		returnVal |= ((val[0] - 'a')+10) << 4;
	}
	else if('A'<=val[0]&&'F'>=val[0])
	{
		returnVal |= ((val[0] - 'A')+10) << 4;
	}

	if('0'<=val[1]&&'9'>=val[1])
	{
		returnVal |= (val[1] - '0');
	}
	else if('a'<=val[1]&&'f'>=val[1])
	{
		returnVal |= (val[1] - 'a')+10;
	}
	else if('A'<=val[1]&&'F'>=val[1])
	{
		returnVal |= (val[1] - 'A')+10;
	}

	return returnVal;
}

unsigned int  ReadIntData(FILE *fR)
{
	int index;
	size_t n;
	unsigned char val[4];
	unsigned int returnVal;

	index = 0;
	n = 1;
	while((index < 4) && n)
	{
		n = fread(&val[index], sizeof(char), 1, fR);
		if(('0'<=val[index]&&'9'>=val[index])||('a'<=val[index]&&'f'>=val[index])||('A'<=val[index]&&'F'>=val[index]))
		{
			index++;
		}
	}

	returnVal = 0;
	if('0'<=val[0]&&'9'>=val[0])
	{
		returnVal = (val[0] - '0') << 12;
	}
	else if('a'<=val[0]&&'f'>=val[0])
	{
		returnVal = ((val[0] - 'a')+10) << 12;
	}
	else if('A'<=val[0]&&'F'>=val[0])
	{
		returnVal = ((val[0] - 'A')+10) << 12;
	}

	if('0'<=val[1]&&'9'>=val[1])
	{
		returnVal |= (val[1] - '0') << 8;
	}
	else if('a'<=val[1]&&'f'>=val[1])
	{
		returnVal |= ((val[1] - 'a')+10) << 8;
	}
	else if('A'<=val[1]&&'F'>=val[1])
	{
		returnVal |= ((val[1] - 'A')+10) << 8;
	}

	if('0'<=val[2]&&'9'>=val[2])
	{
		returnVal |= (val[2] - '0') << 4;
	}
	else if('a'<=val[2]&&'f'>=val[2])
	{
		returnVal |= ((val[2] - 'a')+10) << 4;
	}
	else if('A'<=val[2]&&'F'>=val[2])
	{
		returnVal |= ((val[2] - 'A')+10) << 4;
	}

	if('0'<=val[3]&&'9'>=val[3])
	{
		returnVal |= (val[3] - '0');
	}
	else if('a'<=val[3]&&'f'>=val[3])
	{
		returnVal |= (val[3] - 'a')+10;
	}
	else if('A'<=val[3]&&'F'>=val[3])
	{
		returnVal |= (val[3] - 'A')+10;
	}

	return returnVal;
}

void CLCM_KEYPAD_EEPROMDlg::OnTestLcmKeypad() 
{
	CString			ReaderName;	
	ULONG           dwAP = 0;
	PCHAR strSetKey = "Set Key";
	PCHAR strSetKeyCn = "请按键";
	PCHAR strStar = "********************************";
	UCHAR aWriteCommandClear[] ={GRAPH_MODE_CLEAR_FULL_SCREEN,LCM_KS0108};
	UCHAR bSelectScreen = 1;
	bool  bClearScreenForKeyFlags = true;
	BYTE  bChange, i, j;
	BYTE  abKeyImageBuffer[KEYPAD_ROW_NUM];
	UCHAR bLCMData[] = 
	{
		0x00,0x40,0x47,0xFC,0x30,0x40,0x23,0xF8,0x00,0x40,0x07,0xFE,0xF0,0x00,0x13,0xF8,
		0x12,0x08,0x13,0xF8,0x12,0x08,0x13,0xF8,0x16,0x08,0x1A,0x08,0x12,0x28,0x02,0x10,

		0x20,0x40,0x20,0x40,0xFC,0xA0,0x21,0x18,0x43,0xF6,0x54,0x00,0xFC,0x04,0x53,0xD4,
		0x12,0x54,0x1F,0xD4,0xF2,0x54,0x13,0xD4,0x12,0x54,0x12,0x54,0x13,0x44,0x12,0x8C,

		0x0C,0x00,0x06,0x00,0x02,0x00,0x01,0x00,0x03,0x00,0x02,0x80,0x02,0x80,0x04,0x40,
		0x04,0x20,0x08,0x20,0x08,0x10,0x10,0x08,0x20,0x0E,0x40,0x04,0x80,0x00,0x00,0x00,

		0x02,0x00,0x01,0x00,0x3F,0xFE,0x42,0x24,0x49,0x50,0x29,0x48,0x48,0xA4,0x0B,0x34,
		0x1F,0xE0,0xE0,0x00,0x41,0x00,0x11,0x08,0x11,0x08,0x11,0x08,0x1F,0xF8,0x00,0x00,

		0x00,0x00,0xFD,0xF8,0x10,0x08,0x10,0x88,0x10,0x88,0x20,0x88,0x3C,0x88,0x64,0xFC,
		0xA4,0x04,0x24,0x04,0x25,0xF4,0x24,0x04,0x3C,0x04,0x24,0x04,0x20,0x28,0x00,0x10
	};
	UCHAR aCharSetSpaceKey[] =
	{
		0x00,0x70,0x88,0x08,0x08,0x08,0x38,0x00,0x00,0x38,0x20,0x21,0x21,0x22,0x1C,0x00,
		0x00,0x00,0x80,0x80,0x80,0x80,0x00,0x00,0x00,0x1F,0x22,0x22,0x22,0x22,0x13,0x00,
		0x00,0x80,0x80,0xE0,0x80,0x80,0x00,0x00,0x00,0x00,0x00,0x1F,0x20,0x20,0x00,0x00,
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
		0x08,0xF8,0x88,0xC0,0x28,0x18,0x08,0x00,0x20,0x3F,0x20,0x01,0x26,0x38,0x20,0x00,
		0x00,0x00,0x80,0x80,0x80,0x80,0x00,0x00,0x00,0x1F,0x22,0x22,0x22,0x22,0x13,0x00,
		0x80,0x80,0x80,0x00,0x00,0x80,0x80,0x80,0x80,0x81,0x8E,0x70,0x18,0x06,0x01,0x00,
	};
	UCHAR aCharKeyPadKey[] = 
	{
		0x08,0xF8,0x08,0x00,0x00,0x00,0x00,0x00,0x20,0x3F,0x20,0x20,0x20,0x20,0x30,0x00,
		0x08,0xF8,0xF8,0x00,0xF8,0xF8,0x08,0x00,0x20,0x3F,0x00,0x3F,0x00,0x3F,0x20,0x00,
		0x08,0xF8,0x30,0xC0,0x00,0x08,0xF8,0x08,0x20,0x3F,0x20,0x00,0x07,0x18,0x3F,0x00,
		0xE0,0x10,0x08,0x08,0x08,0x10,0xE0,0x00,0x0F,0x10,0x20,0x20,0x20,0x10,0x0F,0x00,
		0x08,0xF8,0x08,0x08,0x08,0x08,0xF0,0x00,0x20,0x3F,0x21,0x01,0x01,0x01,0x00,0x00,
		0xC0,0x30,0x08,0x08,0x08,0x38,0x00,0x00,0x07,0x18,0x20,0x20,0x22,0x1E,0x02,0x00,
		0x08,0xF8,0x08,0x00,0x00,0x08,0xF8,0x08,0x20,0x3F,0x21,0x01,0x01,0x21,0x3F,0x20,
		0x00,0x08,0x08,0xF8,0x08,0x08,0x00,0x00,0x00,0x20,0x20,0x3F,0x20,0x20,0x00,0x00,
		0x00,0x00,0x08,0x08,0xF8,0x08,0x08,0x00,0xC0,0x80,0x80,0x80,0x7F,0x00,0x00,0x00,
		0x08,0xF8,0x88,0xC0,0x28,0x18,0x08,0x00,0x20,0x3F,0x20,0x01,0x26,0x38,0x20,0x00,
		0x00,0xE0,0x10,0x08,0x08,0x10,0xE0,0x00,0x00,0x0F,0x10,0x20,0x20,0x10,0x0F,0x00,
		0x00,0x10,0x10,0xF8,0x00,0x00,0x00,0x00,0x00,0x20,0x20,0x3F,0x20,0x20,0x00,0x00,
		0x00,0x70,0x08,0x08,0x08,0x88,0x70,0x00,0x00,0x30,0x28,0x24,0x22,0x21,0x30,0x00,
		0x00,0x30,0x08,0x88,0x88,0x48,0x30,0x00,0x00,0x18,0x20,0x20,0x20,0x11,0x0E,0x00,
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x80,0xB0,0x70,0x00,0x00,0x00,0x00,0x00,
		0x00,0x00,0xC0,0x20,0x10,0xF8,0x00,0x00,0x00,0x07,0x04,0x24,0x24,0x3F,0x24,0x00,
		0x00,0xF8,0x08,0x88,0x88,0x08,0x08,0x00,0x00,0x19,0x21,0x20,0x20,0x11,0x0E,0x00,
		0x00,0xE0,0x10,0x88,0x88,0x18,0x00,0x00,0x00,0x0F,0x11,0x20,0x20,0x11,0x0E,0x00,
		0x00,0x38,0x08,0x08,0xC8,0x38,0x08,0x00,0x00,0x00,0x00,0x3F,0x00,0x00,0x00,0x00,
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x30,0x30,0x00,0x00,0x00,0x00,0x00,
		0x00,0x70,0x88,0x08,0x08,0x88,0x70,0x00,0x00,0x1C,0x22,0x21,0x21,0x22,0x1C,0x00,
		0x00,0xE0,0x10,0x08,0x08,0x10,0xE0,0x00,0x00,0x00,0x31,0x22,0x22,0x11,0x0F,0x00,
		0x00,0x00,0xC0,0x38,0xE0,0x00,0x00,0x00,0x20,0x3C,0x23,0x02,0x02,0x27,0x38,0x20,
		0x08,0xF8,0x88,0x88,0x88,0x70,0x00,0x00,0x20,0x3F,0x20,0x20,0x20,0x11,0x0E,0x00,
		0x40,0x40,0x80,0xF0,0x80,0x40,0x40,0x00,0x02,0x02,0x01,0x0F,0x01,0x02,0x02,0x00,
		0xC0,0x30,0x08,0x08,0x08,0x08,0x38,0x00,0x07,0x18,0x20,0x20,0x20,0x10,0x08,0x00,
		0x08,0xF8,0x08,0x08,0x08,0x10,0xE0,0x00,0x20,0x3F,0x20,0x20,0x20,0x10,0x0F,0x00,
		0x08,0xF8,0x88,0x88,0xE8,0x08,0x10,0x00,0x20,0x3F,0x20,0x20,0x23,0x20,0x18,0x00,
		0x08,0xF8,0x88,0x88,0xE8,0x08,0x10,0x00,0x20,0x3F,0x20,0x00,0x03,0x00,0x00,0x00,
		0x40,0x40,0x40,0x40,0x40,0x40,0x40,0x00,0x04,0x04,0x04,0x04,0x04,0x04,0x04,0x00,
	};
	BYTE KeyRom[KEYPAD_ROW_NUM][KEYPAD_COL_NUM] = 
	{
		'C', '8', '4', '0', 'G', 'L',
		'D', '9', '5', '1', 'H', 'M',
		'E', 'A', '6', '2', 'I', 'N',
		'F', 'B', '7', '3', 'J', 'O',
		'=', '*', '.', ',', 'K', 'P',
	};
	BYTE ST7565KeypadCharacterOFFSETArray[KEYPAD_ROW_NUM][KEYPAD_COL_NUM] = 
	{
		OFFSET_C, OFFSET_8, OFFSET_4, OFFSET_0, OFFSET_G, OFFSET_L,
		OFFSET_D, OFFSET_9, OFFSET_5, OFFSET_1, OFFSET_H, OFFSET_M,
		OFFSET_E, OFFSET_A, OFFSET_6, OFFSET_2, OFFSET_I, OFFSET_N,
		OFFSET_F, OFFSET_B, OFFSET_7, OFFSET_3, OFFSET_J, OFFSET_O,
		OFFSET_EqualMark, OFFSET_asterisk, OFFSET_period, OFFSET_comma, OFFSET_K, OFFSET_P,
	};
	UCHAR bData[0x10];
	UCHAR bDisplayData[0x40];
	ULONG lReturnLen, lKeyLength;
	UCHAR cRow, cCol, cSubCol;
	bool bSetKeyFlags;
	CString	strKeypadValue="";
//	UCHAR bWriteData[0x100];

	UpdateData(TRUE);
	m_KeypadValue = "";
	UpdateData(FALSE); 
	// Establish the context.
	switch(m_nLCMIndex)
	{
		case LCM_HD44780:
		{
			if( WriteCommandNocheck(&CSerial9525, 0x38) != 1 ) //row 0, column 0
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}

			//lcd clear display
			if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}

			//lcd set cursor
			if( WriteCommandNocheck(&CSerial9525, SET_CURSOR_EN+HD44780_LINE_SELECT*0+0) != 1 ) //row 0, column 0
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}

			if( WriteCommandNocheck(&CSerial9525, 0x06) != 1 ) //row 0, column 0
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}

			if( WriteCommandNocheck(&CSerial9525, 0x0C) != 1 ) //row 0, column 0
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}	
			//lcd display message
			for(cCol=0, cRow=0;cCol<0x07;cCol++)
			{
				if( WriteData(&CSerial9525, strSetKey[cCol]) != 1 )
				{
					MessageBox("Error! Failed to display char!", "Error", MB_OK);
					return;
				}
				
				if((HD44780_COL_CHAR_NUM-1) == (cCol%HD44780_COL_CHAR_NUM))
				{
					cRow++;
					if(HD44780_ROW_CHAR_NUM == cRow)
					{
						cRow = 0;
					}

					//lcd set cursor
					if( WriteCommandNocheck(&CSerial9525, SET_CURSOR_EN+HD44780_LINE_SELECT*cRow) != 1 )
					{
						MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
						return;
					}
				}
			}			
		}
		break;
		case LCM_KS0108:
		{
			SelectScreen(&CSerial9525, 1);
			if( WriteCommandNocheck(&CSerial9525, KS0108_DISPLAY_OFF) != 1 )
			{
				MessageBox("Error! Failed to off show!", "Error", MB_OK);
				return;
			}
			SelectScreen(&CSerial9525, 0);
			if( WriteCommandNocheck(&CSerial9525, KS0108_DISPLAY_OFF) != 1 )
			{
				MessageBox("Error! Failed to off show!", "Error", MB_OK);
				return;
			}
			SelectScreen(&CSerial9525, 1);
			//SetOnOff(1); //开显示
			if( WriteCommandNocheck(&CSerial9525, KS0108_DISPLAY_ON) != 1 )
			{
				MessageBox("Error! Failed to on show!", "Error", MB_OK);
				return;
			}
			SelectScreen(&CSerial9525, 0);
			//SetOnOff(1); //开显示
			if( WriteCommandNocheck(&CSerial9525, KS0108_DISPLAY_ON) != 1 )
			{
				MessageBox("Error! Failed to on show!", "Error", MB_OK);
				return;
			}
			//Clear Full Screen
			CmdLcmWriteData(&CSerial9525, 0,sizeof(aWriteCommandClear), (unsigned char *)aWriteCommandClear);
			if( WriteCommandNocheck(&CSerial9525, KS0108_SET_LINE_EN) != 1 )
			{
				MessageBox("Error! Failed to set line!", "Error", MB_OK);
				return;
			}
			//SetColumn
			//column=column &0x3f; // 0=<column<=63
			//column=column | 0x40; //01xx xxxx
			//SendCommandToLCD(column);
			if( WriteCommandNocheck(&CSerial9525, SET_COLUMN_EN) != 1 )
			{
				MessageBox("Error! Failed to set column!", "Error", MB_OK);
				return;
			}
			if( SelectScreen(&CSerial9525, 1) != 1 )
			{
				MessageBox("Error! Failed to select screen!", "Error", MB_OK);
				return;
			}
			for(cCol=0;cCol<112;cCol++)
			{
				if(0==(cCol%16))
				{
					if( WriteCommandNocheck(&CSerial9525, KS0108_SET_LINE_EN) != 1 )
					{
						MessageBox("Error! Failed to set line!", "Error", MB_OK);
						return;
					}

					//SetColumn
					//column=column &0x3f; // 0=<column<=63
					//column=column | 0x40; //01xx xxxx
					//SendCommandToLCD(column);
					if( WriteCommandNocheck(&CSerial9525, SET_COLUMN_EN|(cCol/2)) != 1 )
					{
						MessageBox("Error! Failed to set column!", "Error", MB_OK);
						return;
					}
				}
				else if(0==(cCol%8))
				{
					if( WriteCommandNocheck(&CSerial9525, KS0108_SET_LINE_EN+1) != 1 )
					{
						MessageBox("Error! Failed to set line!", "Error", MB_OK);
						return;
					}

					if( WriteCommandNocheck(&CSerial9525, SET_COLUMN_EN|((cCol-8)/2)) != 1 )
					{
						MessageBox("Error! Failed to set column!", "Error", MB_OK);
						return;
					}
				}

				if( WriteData(&CSerial9525, aCharSetSpaceKey[cCol]) != 1 )
				{
					MessageBox("Error! Failed to show set key!", "Error", MB_OK);
					return;
				}
			}
		}
		break;
		case LCM_ST7920:
		{
			//lcd clear display
			if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}

			//lcd set cursor
			if( WriteCommandNocheck(&CSerial9525, ADDR_RESET) != 1 )
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}

			if( WriteCommandNocheck(&CSerial9525, RAM_ADDR_RIGHT_INCREASE) != 1 )//set ram address increasing direction
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}

			if( WriteCommandNocheck(&CSerial9525, ST7920_DISPLAY_ON) != 1 )
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}

			if( WriteCommandNocheck(&CSerial9525, ST7920_SET_LINE_EN) != 1 )//line 0
			{
				MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
				return;
			}
			//lcd display message
			for(cCol=0, cRow=0;cCol<0x06;cCol++)
			{
				if( WriteData(&CSerial9525, strSetKeyCn[cCol]) != 1 )
				{
					MessageBox("Error! Failed to display char!", "Error", MB_OK);
					return;
				}
				
				if((ST7920_COL_CHAR_NUM-1) == (cCol%ST7920_COL_CHAR_NUM))
				{
					cRow++;
					if(0x04 == cRow)
					{
						cRow = 0;
					}
					//else
					{
						//lcd set cursor
						if( WriteCommandNocheck(&CSerial9525, ST7920_SET_LINE_EN+ST7920_LINE_SELECT*(cRow%2)+(cRow/2)*8) != 1 )
						{
							MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
							return;
						}
					}
				}
			}			
		}
		break;
		case LCM_KGM0053:
		{
			//lcd initiate
			LcdKGM0053_init(&CSerial9525);
			lcd_wcmd(&CSerial9525, 0xA5);//All points on
			lcd_wcmd(&CSerial9525, 0xA4);//All points off
			Sleep(1);
			//Display "Set key"
			if(DisplayST7565Data(&CSerial9525, OFFSET_S, 1, 0))
			{
				return;
			}
			DisplayST7565Data(&CSerial9525, OFFSET_E, 1, 8);
			DisplayST7565Data(&CSerial9525, OFFSET_T, 1, 16);
			DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 24);
			DisplayST7565Data(&CSerial9525, OFFSET_K, 1, 32);
			DisplayST7565Data(&CSerial9525, OFFSET_E, 1, 40);
			DisplayST7565Data(&CSerial9525, OFFSET_Y, 1, 48);
			//initiate the variable
			ST7565CurentColumn = 0;
			ST7565CurrentRow = 1;
			Flag_TheFirstKey = 0;		
		}
		break;
		case LCM_PE12832:
		{
			//lcd initiate
			LcdPE12832_init(&CSerial9525);
			lcd_wcmd(&CSerial9525, 0xA5);//All points on
			lcd_wcmd(&CSerial9525, 0xA4);//All points off
			Sleep(1);
			//Display "Set key"
			if(DisplayST7565Data(&CSerial9525, OFFSET_S, 1, 0))
			{
				return;
			}
			DisplayST7565Data(&CSerial9525, OFFSET_E, 1, 8);
			DisplayST7565Data(&CSerial9525, OFFSET_T, 1, 16);
			DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 24);
			DisplayST7565Data(&CSerial9525, OFFSET_K, 1, 32);
			DisplayST7565Data(&CSerial9525, OFFSET_E, 1, 40);
			DisplayST7565Data(&CSerial9525, OFFSET_Y, 1, 48);
			//initiate the variable
			ST7565CurentColumn = 0;
			ST7565CurrentRow = 1;
			Flag_TheFirstKey = 0;		
		}
		break;
		default:
		{}
		break;
	}
	if( CmdClearKeyBuffer(&CSerial9525, 0) != 1 )
	{
		MessageBox("Error! Failed to clear key buffer!", "Error", MB_OK);
		return;
	}

	if( CmdSetKeyScanTimer(&CSerial9525, 0, 1, 1) != 1 )
	{
		MessageBox("Error! Failed to set timer!", "Error", MB_OK);
		return;
	}
	
	lKeyLength = 0;
	bSetKeyFlags = false;
	for(i=0;i<KEYPAD_ROW_NUM;i++)
	{
		abKeyImageBuffer[i] = 0xFF;
	}
	while(1)
	{
//GetKey:
		Sleep(100);

		if( CmdGetKeyInput(&CSerial9525, 0, 0x10, bData, &lReturnLen) != 1 )
		{
			MessageBox("Error! Failed to get key!", "Error", MB_OK);
			return;
		}
		else
		{
			for(i=0;i<KEYPAD_ROW_NUM;i++)
			{
				if(abKeyImageBuffer[i]!=bData[i])
				{
					bChange = abKeyImageBuffer[i]^bData[i];
					for(j=0;j<KEYPAD_COL_NUM;j++)
					{
						if(bChange&0x01)
						{
							if((abKeyImageBuffer[i]>>j)&0x01)//press
							{
								strKeypadValue.Insert(lKeyLength++, KeyRom[i][j]);
								bDisplayData[lKeyLength-1] = i+(KEYPAD_COL_NUM-j-1)*KEYPAD_ROW_NUM;
								ST7565KeypadCharacterOFFSET = ST7565KeypadCharacterOFFSETArray[i][j];

								if(LCM_KS0108!=m_nLCMIndex)
								{
									bDisplayData[lKeyLength-1]=strKeypadValue.GetAt(lKeyLength-1);
								}
								bSetKeyFlags = true;

								UpdateData(TRUE);
								m_KeypadValue = strKeypadValue;
								UpdateData(FALSE);
								//MessageBox(strKeypadValue, "Key", MB_OK);
								switch(m_nLCMIndex)
								{
								case LCM_HD44780:
									if(strSetKeyCn)
									{
										strSetKeyCn = false;
										//lcd set cursor
										if( WriteCommandNocheck(&CSerial9525, SET_CURSOR_EN+HD44780_LINE_SELECT*0+0) != 1 )//row 0, column 0
										{
											MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
											return;
										}

										//lcd clear display
										if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
										{
											MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
											return;
										}
										cRow = 0;
									}

									//lcd display message
									cCol = (unsigned char)(lKeyLength-1);//for(cCol=0, cRow=0;cCol<lKeyLength;cCol++)
									{
										if( WriteData(&CSerial9525, bDisplayData[cCol]) != 1 )
										{
											MessageBox("Error! Failed to display char!", "Error", MB_OK);
											return;
										}
							
										if((HD44780_COL_CHAR_NUM-1) == (cCol%HD44780_COL_CHAR_NUM))
										{
											cRow++;
											if(HD44780_ROW_CHAR_NUM == cRow)
											{
												cRow = 0;
											}
											//else
											{
												//lcd set cursor
												if( WriteCommandNocheck(&CSerial9525, SET_CURSOR_EN+HD44780_LINE_SELECT*cRow) != 1 )
												{
													MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
													return;
												}
											}
										}
									}			
									break;
								case LCM_KS0108:
									//ClearScreen
									if(bClearScreenForKeyFlags)
									{
										bClearScreenForKeyFlags = false;
										//Clear Full Screen
										CmdLcmWriteData(&CSerial9525, 0,sizeof(aWriteCommandClear), (unsigned char *)aWriteCommandClear);

										cRow=0;
										if( SelectScreen(&CSerial9525, 1) != 1 )
										{
											MessageBox("Error! Failed to select screen!", "Error", MB_OK);
											return;
										}
									}

									//SetOnOff(1); //开显示		
									cCol = (unsigned char)(lKeyLength-1);//for(cCol=0, cRow=0;cCol<lKeyLength;cCol++)
									{
										for(cSubCol=0;cSubCol<KS0108_SUB_COL_NUM;cSubCol++)
										{
											if(0==cSubCol)
											{
												if( WriteCommandNocheck(&CSerial9525, KS0108_SET_LINE_EN|(2*cRow)) != 1 )
												{
													MessageBox("Error! Failed to set line!", "Error", MB_OK);
													return;
												}

												if( WriteCommandNocheck(&CSerial9525, SET_COLUMN_EN|((cCol*8)&SET_COLUMN_MASK)) != 1 )
												{
													MessageBox("Error! Failed to set column!", "Error", MB_OK);
													return;
												}
											}
											else if(8==cSubCol)
											{
												if( WriteCommandNocheck(&CSerial9525, KS0108_SET_LINE_EN|(2*cRow+1)) != 1 )
												{
													MessageBox("Error! Failed to set line!", "Error", MB_OK);
													return;
												}

												if( WriteCommandNocheck(&CSerial9525, SET_COLUMN_EN|((cCol*8)&SET_COLUMN_MASK)) != 1 )
												{
													MessageBox("Error! Failed to set column!", "Error", MB_OK);
													return;
												}
											}

											if( WriteData(&CSerial9525, aCharKeyPadKey[bDisplayData[cCol]*KS0108_SUB_COL_NUM+cSubCol]) != 1 )
											{
												MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
												return;
											}
										}
							
										if((KS0108_COL_CHAR_NUM-1) == (cCol%KS0108_COL_CHAR_NUM))
										{
											if((KS0108_COL_CHAR_NUM*2-1) == (cCol%(KS0108_COL_CHAR_NUM*2)))
											{
												cRow++;
												bSelectScreen = 0x01;
											}
											else
											{
												bSelectScreen = 0x00;
											}

											if( SelectScreen(&CSerial9525, bSelectScreen) != 1 )
											{
												MessageBox("Error! Failed to select screen!", "Error", MB_OK);
												return;
											}
											
											if(KS0108_ROW_CHAR_NUM == cRow)
											{
												cRow = 0;
											}
										}
									}
									break;

								case LCM_ST7920:
									if(strSetKeyCn)
									{
										strSetKeyCn = false;
										//lcd set cursor
										if( WriteCommandNocheck(&CSerial9525, ADDR_RESET) != 1 )
										{
											MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
											return;
										}

										//lcd clear display
										if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
										{
											MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
											return;
										}
										cRow = 0;
									}

									//lcd display message
									cCol = (unsigned char)(lKeyLength-1);//for(cCol=0, cRow=0;cCol<lKeyLength;cCol++)
									{
										if( WriteData(&CSerial9525, bDisplayData[cCol]) != 1 )
										{
											MessageBox("Error! Failed to display char!", "Error", MB_OK);
											return;
										}
							
										if((ST7920_COL_CHAR_NUM-1) == (cCol%ST7920_COL_CHAR_NUM))
										{
											cRow++;
											if(0x04 == cRow)
											{
												cRow = 0;
											}
											//else
											{
												//lcd set cursor
												if( WriteCommandNocheck(&CSerial9525, ST7920_SET_LINE_EN+ST7920_LINE_SELECT*(cRow%2)+(cRow/2)*8) != 1 )
												{
													MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
													return;
												}
											}
										}
									}			
									break;
								case LCM_KGM0053:
									//Display the key
									if(!Flag_TheFirstKey)
									{
										//Display off
										if(lcd_wcmd(&CSerial9525, 0XAE)!= 1)
										{
											MessageBox("Error! Failed to display off!", "Error", MB_OK);
											return;	
										}
										Flag_TheFirstKey = 1;
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 0))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 8))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 16))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 32))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 40))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 48))
										{
											return;
										}
										//Display on
										if(lcd_wcmd(&CSerial9525, 0XAF)!= 1)
										{
											MessageBox("Error! Failed to display on!", "Error", MB_OK);
											return;	
										}
									}

									if(DisplayST7565Data(&CSerial9525, ST7565KeypadCharacterOFFSET, ST7565CurrentRow, ST7565CurentColumn))
									{
										return;
									}					
									if(ST7565CurentColumn == 120)
									{
										ST7565CurentColumn = 0;
										if(ST7565CurrentRow == 2)
										{
											ST7565CurrentRow = 1;
										}
										else
										{
											ST7565CurrentRow = 2;
										}
									}
									else
									{
										ST7565CurentColumn += 8;
									}
									break;
								case LCM_PE12832:
									//Display the key
									if(!Flag_TheFirstKey)
									{
										//Display off
										if(lcd_wcmd(&CSerial9525, 0XAE)!= 1)
										{
											MessageBox("Error! Failed to display off!", "Error", MB_OK);
											return;	
										}
										Flag_TheFirstKey = 1;

										//Clear "SET KEY"
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 0))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 8))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 16))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 32))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 40))
										{
											return;
										}
										if(DisplayST7565Data(&CSerial9525, OFFSET_Space, 1, 48))
										{
											return;
										}
										//Display on
										if(lcd_wcmd(&CSerial9525, 0XAF)!= 1)
										{
											MessageBox("Error! Failed to display on!", "Error", MB_OK);
											return;	
										}
									}

									if(DisplayST7565Data(&CSerial9525, ST7565KeypadCharacterOFFSET, ST7565CurrentRow, ST7565CurentColumn))
									{
										return;
									}					
									if(ST7565CurentColumn == 120)
									{
										ST7565CurentColumn = 0;
										if(ST7565CurrentRow == 2)
										{
											ST7565CurrentRow = 1;
										}
										else
										{
											ST7565CurrentRow = 2;
										}
									}
									else
									{
										ST7565CurentColumn += 8;
									}
									break;
								default:
									break;
								}

							}
						}
						bChange >>= 1;
					}
					abKeyImageBuffer[i]=bData[i];
					//OfflineGetKeyDisplay();
				}
			}

			bSetKeyFlags = false;

			if(lKeyLength>0&&'='==strKeypadValue.GetAt(lKeyLength-1))
			{
				MessageBox(strKeypadValue, "Key", MB_OK);
				if( CmdSetKeyScanTimer(&CSerial9525, 0, 0, 0) != 1 )
				{
					MessageBox("Error! Failed to clear timer!", "Error", MB_OK);
					return;
				}
				break;
			}
		}
	}
}



void CLCM_KEYPAD_EEPROMDlg::OnDisplayGraph() 
{
	// TODO: Add your control notification handler code here
	// TODO: Add your control notification handler code here
	CString			ReaderName;	
	ULONG           dwAP = 0;
	CString filterOpen, nameOpen;
	char name[128],buff[80];
	UCHAR cDisplayData[1024];
	filterOpen = "TEXT Files(*.txt;*.TXT)|*.txt;*.TXT";
	FILE *fR;
	UCHAR cRow, cCol,x,y;//,cPart;//, cSubCol;
	UINT  nLength,nHigh,i,h,l;
	long j,length;
	UCHAR aWriteCommandClear[] ={GRAPH_MODE_CLEAR_FULL_SCREEN,LCM_ST7920};

	char *strMessage;

	strMessage = (char *)malloc(0x200);
	if(NULL==strMessage)
	{
		MessageBox("allocating strMessage space fials.\n");
		return;
	}


	// Establish the context.

	// Get the first enumerated reader's name 

	// Get the reader's handle 

	CFileDialog alcorFileDlg(TRUE, NULL, NULL, OFN_HIDEREADONLY, filterOpen);

    if(m_nLCMIndex==LCM_HD44780)
    {
		MessageBox("Error! HD44780 does support graph displaying!", "Error", MB_OK);
		goto DisplayGraphEnd;
	}

	if( alcorFileDlg.DoModal() != IDOK )
	{
		MessageBox("Cannot open file");
		free(strMessage);
		return;
	}

	nameOpen = alcorFileDlg.GetPathName();
	length = (long) nameOpen.GetLength();
	for(j=0; j < length; j++)
	{
		name[j] = nameOpen.GetAt(j);
	}
	name[length] = '\0';

	if( (fR=fopen(name,"r"))==NULL )
	{
		sprintf(strMessage,"OpenTable: could not find %s\n", name);
		MessageBox(strMessage);
		free(strMessage);
		return;
	}

	length=0;

	while(!feof(fR))
	{
			fgets(buff,80,fR); 
			if( (buff[0]=='0')&&(buff[1]=='x') )
			{
				for(i=0;i<16;i++)
				{
					if( (buff[i*5+2] >= 'A')&&(buff[i*5+2] <= 'F') )
						cDisplayData[length]=(buff[i*5+2]-55)*16;
					else if( (buff[i*5+2] >= '0')&&(buff[i*5+2] <= '9') ) 
						cDisplayData[length]=(buff[i*5+2]-48)*16;
					else
						break;

					if( (buff[i*5+3] >= 'A')&&(buff[i*5+3] <= 'F') )
						cDisplayData[length]+=(buff[i*5+3]-55);
					else if( (buff[i*5+3] >= '0')&&(buff[i*5+3] <= '9') )
						cDisplayData[length]+=(buff[i*5+3]-48);	
					else
						break;
					
					length++;
				}
			}
			else 
			{
				continue;
			}

	}
			
	fclose(fR);

	UpdateData(TRUE); //TRUE

	switch(m_nLCMIndex)
	{
/*	case LCM_HD44780:

		MessageBox("Error! HD44780 does support graph displaying!", "Error", MB_OK);

		break;*/
	

	case LCM_KS0108:

		nLength = m_DisplayLength;
		nHigh = m_DisplayHigh;
		cRow=m_LCMPosX;
		cCol=m_LCMPosY;
		if( ((cCol+nLength) > 128) || ((cRow*8+nHigh)>64) )
		{
			MessageBox("Error! Failed to set length and high!", "Error", MB_OK);
			return;
		}
		for( h=0;h<nHigh;h++)
		{

			for( l=0;l<nLength;l++ )
			{
				y = cCol&0x3F|SET_COLUMN_EN;		// col.and.0x3f.or.setx	
				x = cRow&0x07|KS0108_SET_LINE_EN;		    //  row.and.0x07.or.sety	
				if( (cCol&0xc0)==0 )		//  col.and.0xC0
				{			
					// left screen
					if( SelectScreen(&CSerial9525, 1) != 1 )
					{
						MessageBox("Error! Failed to select screen!", "Error", MB_OK);
						return;
					}

				}
				else if( (cCol&0xc0)==0x40 )
				{
					// right screen
					if( SelectScreen(&CSerial9525, 0) != 1 )
					{
						MessageBox("Error! Failed to select screen!", "Error", MB_OK);
						return;
					}
				}
				if( WriteCommandNocheck(&CSerial9525, x) != 1 )
				{
					MessageBox("Error! Failed to set line!", "Error", MB_OK);
					return;
				}

				if( WriteCommandNocheck(&CSerial9525, y) != 1 )
				{
					MessageBox("Error! Failed to set column!", "Error", MB_OK);
					return;
				}

				if( WriteData(&CSerial9525, cDisplayData[nLength*h+l]) != 1 )
				{
					MessageBox("Error! Failed to show data!", "Error", MB_OK);
					return;
				}

				cCol++;
				length--;
				if(length==0)
				{
					return;
				}
			}

			cRow++;
			cCol=m_LCMPosY;

		}

		break;

	case LCM_ST7920:

		if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		//lcd set cursor
		if( WriteCommandNocheck(&CSerial9525, ADDR_RESET) != 1 )
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		if( WriteCommandNocheck(&CSerial9525, RAM_ADDR_RIGHT_INCREASE) != 1 )//set ram address increasing direction
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		if( WriteCommandNocheck(&CSerial9525, ST7920_DISPLAY_ON) != 1 )
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		if( WriteCommandNocheck(&CSerial9525, ST7920_SET_LINE_EN) != 1 )//line 0
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}
	
		nLength = m_DisplayLength/16;
		nHigh = m_DisplayHigh;
		cRow=m_LCMPosX;
		cCol=m_LCMPosY/16;
		if( ((cCol+nLength) > 16) || ((cRow+nHigh)>64) )
		{
			MessageBox("Error! Failed to set length and high!", "Error", MB_OK);
			return;
		}

		//Clear Full Screen
		if( CmdLcmWriteData(&CSerial9525, 0,sizeof(aWriteCommandClear), (unsigned char *)aWriteCommandClear) != 1 )
		{
			MessageBox("Error! Failed to clear GDRAM!", "Error", MB_OK);
			return;
		}

		if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
		{
			MessageBox("Error! Failed to clear screen!", "Error", MB_OK);
			return;
		}

		j=0;
		i=0;	

		for( h=0;h<nHigh;h++)
		{
			if( (h+cRow) >= 32 )
				i=1;   
			else
				i=0;

			for( l=0;l<nLength;l++ )
			{
				if( WriteCommandNocheck(&CSerial9525, 0x36) != 1 )
				{
					MessageBox("Error! Failed to set graph display on!", "Error", MB_OK);
					return;
				}
				if( WriteCommandNocheck(&CSerial9525, 0x80+(cRow+h)%32) != 1 )
				{
					MessageBox("Error! Failed to set row address!", "Error", MB_OK);
					return;
				}

				if( WriteCommandNocheck(&CSerial9525, 0x80+cCol+l+i*8) != 1 )
				{
					MessageBox("Error! Failed to set column address!", "Error", MB_OK);
					return;
				}
				if( WriteCommandNocheck(&CSerial9525, 0x30) != 1 )
				{
					MessageBox("Error! Failed to set graph display off!", "Error", MB_OK);
					return;
				}

				if( WriteData(&CSerial9525, cDisplayData[j++]) != 1 )
				{
					MessageBox("Error! Failed to show data!", "Error", MB_OK);
					return;
				}
				length--;
				if(length==0)
				{
					return;
				}
				if( WriteData(&CSerial9525, cDisplayData[j++]) != 1 )
				{
					MessageBox("Error! Failed to show data!", "Error", MB_OK);
					return;
				}
				length--;
				if(length==0)
				{
					return;
				}
			}
		}	
		break;
		case LCM_KGM0053:
		{
			LcdKGM0053_init(&CSerial9525);

			lcd_wcmd(&CSerial9525, (0xb0 + 0));
			lcd_wcmd(&CSerial9525, (0x10));
			lcd_wcmd(&CSerial9525, (0x00));
			for(i=0;i<32;i++)
			{
				lcd_wdata(&CSerial9525, cDisplayData[i]);
			}

			lcd_wcmd(&CSerial9525, (0xb0 + 1));
			lcd_wcmd(&CSerial9525, (0x10));
			lcd_wcmd(&CSerial9525, (0x00));
			for(i=0;i<32;i++)
			{
				lcd_wdata(&CSerial9525, cDisplayData[i+32]);
			}

			lcd_wcmd(&CSerial9525, (0xb0 + 2));
			lcd_wcmd(&CSerial9525, (0x10));
			lcd_wcmd(&CSerial9525, (0x00));
			for(i=0;i<32;i++)
			{
				lcd_wdata(&CSerial9525, cDisplayData[i+64]);
			}

			lcd_wcmd(&CSerial9525, (0xb0 + 3));
			lcd_wcmd(&CSerial9525, (0x10));
			lcd_wcmd(&CSerial9525, (0x00));
			for(i=0;i<32;i++)
			{
				lcd_wdata(&CSerial9525, cDisplayData[i+96]);
			}
		}
		break;
		case LCM_PE12832:
		{
			LcdPE12832_init(&CSerial9525);
			lcd_wcmd(&CSerial9525, (0xb0 + 0));
			lcd_wcmd(&CSerial9525, (0x10));
			lcd_wcmd(&CSerial9525, (0x00));
			for(i=0;i<32;i++)
			{
				lcd_wdata(&CSerial9525, cDisplayData[i]);
			}

			lcd_wcmd(&CSerial9525, (0xb0 + 1));
			lcd_wcmd(&CSerial9525, (0x10));
			lcd_wcmd(&CSerial9525, (0x00));
			for(i=0;i<32;i++)
			{
				lcd_wdata(&CSerial9525, cDisplayData[i+32]);
			}

			lcd_wcmd(&CSerial9525, (0xb0 + 2));
			lcd_wcmd(&CSerial9525, (0x10));
			lcd_wcmd(&CSerial9525, (0x00));
			for(i=0;i<32;i++)
			{
				lcd_wdata(&CSerial9525, cDisplayData[i+64]);
			}

			lcd_wcmd(&CSerial9525, (0xb0 + 3));
			lcd_wcmd(&CSerial9525, (0x10));
			lcd_wcmd(&CSerial9525, (0x00));
			for(i=0;i<32;i++)
			{
				lcd_wdata(&CSerial9525, cDisplayData[i+96]);
			}
		}
		break;
	default:
		break;
	}
DisplayGraphEnd:
	UpdateData(FALSE); 
	
	return;		
}

void CLCM_KEYPAD_EEPROMDlg::OnSelchangeLcmSelect() 
{
	// TODO: Add your control notification handler code here
	m_nLCMIndex = m_ctlLCMList.GetCurSel();
	UpdateData(TRUE);
	switch(m_nLCMIndex)
	{
		case LCM_HD44780:
			m_RangeX = _T("(0-1)");
			m_RangeY = _T("(0-15)");		
			break;
		case LCM_KS0108:
			m_RangeX = _T("(0-7)");
			m_RangeY = _T("(0-127)");		
			break;	
		case LCM_ST7920:
			m_RangeX = _T("(0-3/63)");
			m_RangeY = _T("(0-7/127)");		
			break;	
		case LCM_KGM0053:
			m_RangeX = _T("");
			m_RangeY = _T("");
			break;
		case LCM_PE12832:
			m_RangeX = _T("");
			m_RangeY = _T("");
			break;
		default:
			break;
	}
	UpdateData(FALSE);
	return;
}

void CLCM_KEYPAD_EEPROMDlg::OnDisplayText() 
{
	// TODO: Add your control notification handler code here
	CString			ReaderName;	
	ULONG           dwAP = 0;
	UCHAR cRow, cCol,cLength,i;//, cSubCol;
    UCHAR aWriteCommandClear[] ={GRAPH_MODE_CLEAR_FULL_SCREEN,LCM_KS0108};
	CString filterOpen, nameOpen;
	char name[128],buff[80];
	UCHAR cDisplayData[1024];
	filterOpen = "TEXT Files(*.txt;*.TXT)|*.txt;*.TXT";
	FILE *fR;
	UCHAR x,y;//, cSubCol;
	UINT  nLength,nHigh,l;
	long j,length;

	// Establish the context.


	// Get the first enumerated reader's name 


	// Get the reader's handle 

	if( m_nLCMIndex == LCM_KS0108 )
	{
		char *strMessage;

		strMessage = (char *)malloc(0x200);
		if(NULL==strMessage)
		{
			MessageBox("allocating strMessage space fials.\n");
			return;
		}

		CFileDialog alcorFileDlg(TRUE, NULL, NULL, OFN_HIDEREADONLY, filterOpen);
		if( alcorFileDlg.DoModal() != IDOK )
		{
			MessageBox("Cannot open file");
			free(strMessage);
			return;
		}

		nameOpen = alcorFileDlg.GetPathName();
		length = (long) nameOpen.GetLength();
		for(j=0; j < length; j++)
		{
			name[j] = nameOpen.GetAt(j);
		}
		name[length] = '\0';

		if( (fR=fopen(name,"r"))==NULL )
		{
			sprintf(strMessage,"OpenTable: could not find %s\n", name);
			MessageBox(strMessage);
			free(strMessage);
			return;
		}
	}

	UpdateData(TRUE); //TRUE

	switch(m_nLCMIndex)
	{
	case LCM_HD44780:
        if( WriteCommandNocheck(&CSerial9525, 0x38) != 1 ) //设定LCD为16*2显示，5*7点阵，8位数据接口
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		//lcd clear display
		if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )//显示清屏
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		//lcd set cursor
		if( WriteCommandNocheck(&CSerial9525, SET_CURSOR_EN+HD44780_LINE_SELECT*0+0) != 1 ) //row 0, column 0
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}
        //显示光标自动右移，整屏不移动
		if( WriteCommandNocheck(&CSerial9525, 0x06) != 1 ) 
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		//开显示，不显示光标
		if( WriteCommandNocheck(&CSerial9525, 0x0C) != 1 ) 
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}



		cLength = m_DisplayString.GetLength();
		cRow=m_LCMPosX;
		cCol=m_LCMPosY;
		i=0;



		//lcd display message
		while( cLength )
		{
			if( WriteData(&CSerial9525, m_DisplayString.GetAt(i)) != 1 )
			{
				MessageBox("Error! Failed to display char!", "Error", MB_OK);
				return;
			}
			
			cLength--;
			i++;
			cCol++;

			if(0 == (cCol%HD44780_COL_CHAR_NUM))
			{
				cRow++;
				if(HD44780_ROW_CHAR_NUM == cRow)
				{
					cRow = 0;
				}

				if( WriteCommandNocheck(&CSerial9525, SET_CURSOR_EN+HD44780_LINE_SELECT*cRow) != 1 )
				{
					MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
					return;
				}
				
			}
		}
		
		break;
	case LCM_KS0108:
        SelectScreen(&CSerial9525, 1);
		if( WriteCommandNocheck(&CSerial9525, KS0108_DISPLAY_OFF) != 1 )
		{
			MessageBox("Error! Failed to off show!", "Error", MB_OK);
			return;
		}
		SelectScreen(&CSerial9525, 0);
		if( WriteCommandNocheck(&CSerial9525, KS0108_DISPLAY_OFF) != 1 )
		{
			MessageBox("Error! Failed to off show!", "Error", MB_OK);
			return;
		}
		SelectScreen(&CSerial9525, 1);
		//SetOnOff(1); //开显示
		if( WriteCommandNocheck(&CSerial9525, KS0108_DISPLAY_ON) != 1 )
		{
			MessageBox("Error! Failed to on show!", "Error", MB_OK);
			return;
		}
		SelectScreen(&CSerial9525, 0);
		//SetOnOff(1); //开显示
		if( WriteCommandNocheck(&CSerial9525, KS0108_DISPLAY_ON) != 1 )
		{
			MessageBox("Error! Failed to on show!", "Error", MB_OK);
			return;
		}
        //Clear Full Screen
		CmdLcmWriteData(&CSerial9525, 0,sizeof(aWriteCommandClear), (unsigned char *)aWriteCommandClear);

		if( WriteCommandNocheck(&CSerial9525, KS0108_SET_LINE_EN) != 1 )
		{
			MessageBox("Error! Failed to set line!", "Error", MB_OK);
			return;
		}
		if( WriteCommandNocheck(&CSerial9525, SET_COLUMN_EN) != 1 )
		{
			MessageBox("Error! Failed to set column!", "Error", MB_OK);
			return;
		}  

		if( SelectScreen(&CSerial9525, 1) != 1 )
		{
			MessageBox("Error! Failed to select screen!", "Error", MB_OK);
			return;
		} 


		nLength = m_DisplayLength;
		nHigh = m_DisplayHigh;
		cRow=m_LCMPosX;
		cCol=m_LCMPosY;
		length=0;
		while(!feof(fR))
		{
			fread(buff, sizeof(unsigned char), 0x01, fR); 
			if( buff[0]=='0' )
			{
				fread(buff, sizeof(unsigned char), 0x01, fR);
				if( buff[0]=='x' )
				{
					fread(buff, sizeof(unsigned char), 0x02, fR);
					if( (buff[0] >= 'A')&&(buff[0] <= 'F') )
						cDisplayData[length]=(buff[0]-55)*16;
					else if( (buff[0] >= '0')&&(buff[0] <= '9') ) 
						cDisplayData[length]=(buff[0]-48)*16;
					else
						break;

					if( (buff[1] >= 'A')&&(buff[1] <= 'F') )
						cDisplayData[length]+=(buff[1]-55);
					else if( (buff[1] >= '0')&&(buff[1] <= '9') )
						cDisplayData[length]+=(buff[1]-48);	
					else
						break;
					
					length++;
				}
				else
				{
					continue;
				}
			}
			else 
			{
				continue;
			}

		}
			
		fclose(fR);

		
		if( ((cCol+nLength) > 128) || ((cRow*8+nHigh)>64) )
		{
			MessageBox("Error! Failed to set length and high!", "Error", MB_OK);
			return;
		}

		j=0;

		while( length )
		{
			for( l=0;l<nLength;l++ )
			{
				y = cCol&0x3F|SET_COLUMN_EN;		// col.and.0x3f.or.setx	
				x = cRow&0x07|KS0108_SET_LINE_EN;		    //  row.and.0x07.or.sety	
				if( (cCol&0xc0)==0 )		//  col.and.0xC0
				{			
					// left screen
					if( SelectScreen(&CSerial9525, 1) != 1 )
					{
						MessageBox("Error! Failed to select screen!", "Error", MB_OK);
						return;
					}

				}
				else if( (cCol&0xc0)==0x40 )
				{
					// right screen
					if( SelectScreen(&CSerial9525, 0) != 1 )
					{
						MessageBox("Error! Failed to select screen!", "Error", MB_OK);
						return;
					}
				}
				if( WriteCommandNocheck(&CSerial9525, x) != 1 )
				{
					MessageBox("Error! Failed to set line!", "Error", MB_OK);
					return;
				}

				if( WriteCommandNocheck(&CSerial9525, y) != 1 )
				{
					MessageBox("Error! Failed to set column!", "Error", MB_OK);
					return;
				}

				if( WriteData(&CSerial9525, cDisplayData[j++]) != 1 )
				{
					MessageBox("Error! Failed to show set key!", "Error", MB_OK);
					return;
				}

				cCol++;
				length--;
			}
			if( nHigh > 8 )
			{
				cRow++;
				if( cRow == 8 )
				{					
					MessageBox("Error! Screen is full!", "Error", MB_OK);
					return;
				}
				nHigh -= 8;
				cCol-=nLength;
			}
			else
			{
				if( cCol == 128 )
				{
					cRow++;
					m_LCMPosX+=2;
					if( cRow == 8 )
					{
						cRow = 0;
						m_LCMPosX=0;
					}
					cCol = 0;
				}
				else
				{			
					cRow=m_LCMPosX;
				}
				nHigh = m_DisplayHigh;
			}

		}
		break;

	case LCM_ST7920:

		if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		//lcd set cursor
		if( WriteCommandNocheck(&CSerial9525, ADDR_RESET) != 1 )
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		if( WriteCommandNocheck(&CSerial9525, RAM_ADDR_RIGHT_INCREASE) != 1 )//set ram address increasing direction
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		if( WriteCommandNocheck(&CSerial9525, ST7920_DISPLAY_ON) != 1 )
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		if( WriteCommandNocheck(&CSerial9525, ST7920_SET_LINE_EN) != 1 )//line 0
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}
	
		cLength = m_DisplayString.GetLength();
		cRow=m_LCMPosX;
		cCol=m_LCMPosY*2;
		i=0;

		if( WriteCommandNocheck(&CSerial9525, ST7920_SET_LINE_EN+ST7920_LINE_SELECT*(cRow%2)+(cRow/2)*8+m_LCMPosY) != 1 )
		{
			MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
			return;
		}

		//lcd display message
		while( cLength )
		{

			if( WriteData(&CSerial9525, m_DisplayString.GetAt(i)) != 1 )
			{
				MessageBox("Error! Failed to display char!", "Error", MB_OK);
				return;
			}

			i++;
			cCol++;
			cLength--;
			
			if( 0 == cCol%ST7920_COL_CHAR_NUM )
			{
				cRow++;
				if(0x04 == cRow)
				{
					cRow = 0;
				}

				if( WriteCommandNocheck(&CSerial9525, ST7920_SET_LINE_EN+ST7920_LINE_SELECT*(cRow%2)+(cRow/2)*8) != 1 )
				{
					MessageBox("Error! Failed to set cursor!", "Error", MB_OK);
					return;
				}
			}

		}		
		break;
		case LCM_KGM0053:
		{
			cLength = m_DisplayString.GetLength();
			if(cLength > 32)
			{
				MessageBox("the length of the text is too long", "Error", MB_OK);
				return;
			}
			else
			{
				m_DisplayString.MakeUpper();
				LcdKGM0053_init(&CSerial9525);
				BYTE DisplayRow = 0;
				BYTE DisplayColumn = 0;
				BYTE OFFSET[32];
				for(i=0;i<cLength;i++)
				{
					OFFSET[i] = ByteASCToOffset(m_DisplayString.GetAt(i));
					if(OFFSET[i] == 0xff)
					{
						return;
					}
				}
				for(i=0;i<cLength;i++)
				{
					if(i<16)
					{
						DisplayRow = 1;
						DisplayColumn = (i * 8);
					}
					else
					{
						DisplayRow = 2;
						DisplayColumn = ((i - 16) * 8);				
					}
					if(DisplayST7565Data(&CSerial9525, OFFSET[i], DisplayRow, DisplayColumn))
					{
						return;
					}
				}
			}	
		}
		break;
		case LCM_PE12832:
		{
			cLength = m_DisplayString.GetLength();
			if(cLength > 32)
			{
				MessageBox("the length of the text is too long", "Error", MB_OK);
				return;
			}
			else
			{
				m_DisplayString.MakeUpper();
				LcdPE12832_init(&CSerial9525);
				BYTE DisplayRow = 0;
				BYTE DisplayColumn = 0;
				BYTE OFFSET[32];
				for(i=0;i<cLength;i++)
				{
					OFFSET[i] = ByteASCToOffset(m_DisplayString.GetAt(i));
					if(OFFSET[i] == 0xff)
					{
						return;
					}
				}
				for(i=0;i<cLength;i++)
				{
					if(i<16)
					{
						DisplayRow = 1;
						DisplayColumn = (i * 8);
					}
					else
					{
						DisplayRow = 2;
						DisplayColumn = ((i - 16) * 8);				
					}
					if(DisplayST7565Data(&CSerial9525, OFFSET[i], DisplayRow, DisplayColumn))
					{
						return;
					}
				}
			}
		}
		break;
		default:
		{
		}
		break;
	}

//	UpdateData(FALSE); 
	return;	
}

void CLCM_KEYPAD_EEPROMDlg::OnDISPLAYClear() 
{
	// TODO: Add your control notification handler code here

	CString			ReaderName;	

	ULONG           dwAP = 0;
	UCHAR aWriteCommandClear[] ={GRAPH_MODE_CLEAR_FULL_SCREEN,LCM_KS0108};
	//UCHAR cRow, cCol,i;//, cSubCol;
	//bool bSetKeyFlags;
	//CString	strKeypadValue=_T("");
//	UCHAR bWriteData[0x100];

	// Establish the context.
//-	if ( SCardEstablishContext(SCARD_SCOPE_SYSTEM, NULL, NULL, &hSC) != 1 )
//-	{
//-		MessageBox("Error! Failed to connect to smartcard service", "Error", MB_OK);
//-		return;
//-	}

	// Get the first enumerated reader's name 
//-	if( GetReaderName(hSC, &ReaderName) != 1 )
//-	{
//-		SCardReleaseContext(hSC);
//-		MessageBox("Error! Failed to get the reader's name", "Error", MB_OK);
//-		return;
//-	}

	// Get the reader's handle 
//-	if( SCardConnect( hSC, ReaderName, SCARD_SHARE_DIRECT, NULL, &&CSerial9525, &dwAP ) != 1 )
//-	{
//-		SCardReleaseContext(hSC);
//-		MessageBox("Error! Failed to connect to smartcard reader", "Error", MB_OK);
//-		return;
//-	}

	switch(m_nLCMIndex)
	{
	case LCM_HD44780:

		//lcd set cursor
		if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
		{
			MessageBox("Error! Failed to clear screen!", "Error", MB_OK);
			return;
		}
	
		break;

	case LCM_KS0108:
	
		//Clear Full Screen
		if( CmdLcmWriteData(&CSerial9525, 0,sizeof(aWriteCommandClear), (unsigned char *)aWriteCommandClear) != 1 )
		{
			MessageBox("Error! Failed to clear screen!", "Error", MB_OK);
			return;
		}
		break;

	case LCM_ST7920:
	
		if( WriteCommandNocheck(&CSerial9525, CLEAR_DISPLAY) != 1 )
		{
			MessageBox("Error! Failed to clear screen!", "Error", MB_OK);
			return;
		}
		
		break;
	case LCM_KGM0053:
		LcdKGM0053_init(&CSerial9525);
		break;
	case LCM_PE12832:
		LcdPE12832_init(&CSerial9525);
		break;
	default:
		break;
	}
	return;		
}

void CLCM_KEYPAD_EEPROMDlg::OnBacklightChange() 
{
	// TODO: Add your control notification handler code here
	// TODO: Add your control notification handler code here
	CString			ReaderName;	
	ULONG           dwAP = 0;

	// Establish the context.


	// Get the first enumerated reader's name 


	// Get the reader's handle 
		
	if( CmdLcmSetBacklight(&CSerial9525) != 1 )
	{
		MessageBox("Error! Failed to set backlight!", "Error", MB_OK);
		return;
	}
	return;	
}

void CLCM_KEYPAD_EEPROMDlg::OnUpdateFw() 
{
	// TODO: Add your control notification handler code here
	CString filterOpen, nameOpen;
	char name[128];
	filterOpen = "HEX Files(*.hex;*.HEX)|*.hex;*.HEX|BIN Files(*.bin;*.BIN)|*.bin;*.BIN";
	FILE *fR;
	unsigned char cBuffer[0x10000];
	memset(cBuffer, 0, 0x10000);
	long i,length;
	unsigned char cHeader,cLength;
	long  iAddr,iAddrMax;
	long  lFWStartPos, lFWEndPos;
	int iFilesType;
	bool bByteflags;

	CString			ReaderName;	
	ULONG           dwAP = 0;
	COLORREF		ResultColor;

	char *strMessage;

	strMessage = (char *)malloc(0x200);
	if(NULL==strMessage)
	{
		MessageBox("allocating strMessage space fials.\n");
		return;
	}

	ResultColor = RGB(0,0,255);
	m_Result.SetTextColor(ResultColor);
	m_Result.SetWindowText("");
	m_Prog->SetPos(0);
	m_Status->SetWindowText("");

	UpdateData(TRUE); //TRUE

	lFWStartPos = 0xA000;              

	lFWEndPos = 0xA7FF;


	// Establish the context.

	// Get the first enumerated reader's name 
	
	// Get the reader's handle 

	// Write to Eeprom 
	ResultColor = RGB(0,0,255);
	m_Result.SetTextColor(ResultColor);
	m_Result.SetWindowText("Wait ....");
	m_Status->SetWindowText("Write Firmware");
	m_Prog->SetPos(0);
#ifdef LOAD_FW
	CFileDialog alcorFileDlg(TRUE, NULL, NULL, OFN_HIDEREADONLY, filterOpen);
	if( alcorFileDlg.DoModal() != IDOK )
	{
		MessageBox("Cannot open file");
		free(strMessage);
		return;
	}
	nameOpen = alcorFileDlg.GetPathName();
	length = (long) nameOpen.GetLength();
	for(i=0; i < length; i++)
	{
		name[i] = nameOpen.GetAt(i);
	}
	name[length] = '\0';
	iFilesType = 0;
	if('x'==name[length-1]||'X'==name[length-1])
	{
		iFilesType = 1;
	}
	else if('n'==name[length-1]||'N'==name[length-1])
	{
		iFilesType = 2;
	}
	if( (fR=fopen(name,"rb"))==NULL )
	{
		sprintf(strMessage,"OpenTable: could not find %s\n", name);
		MessageBox(strMessage);
		free(strMessage);
		return;
	}

	length = 0x0000;
	iAddrMax = 0x0000;
	bByteflags = true;
	while(!feof(fR))
	{
		fread(&cHeader, sizeof(unsigned char), 0x01, fR);

		if(1==iFilesType)
		{
			if(0x3A!=cHeader)
			{
				if(0x00==cHeader&&0x0000==ReadIntData(fR)&&0x01FFF==ReadIntData(fR))
				{
					break;
				}
				else
				{
					continue;
				}
			}

			cLength = ReadCharData(fR);
			iAddr = ReadIntData(fR);
			if(iAddr>iAddrMax)
			{
				iAddrMax = iAddr;
				length = iAddrMax+cLength;
			}
			ReadCharData(fR);

			for(i=0;i<cLength;i++)
			{
				cBuffer[iAddr+i] = ReadCharData(fR);
			}
		}
		else
		{
			cBuffer[length++] = cHeader;
		}
	}
	fclose(fR);
	while(length<lFWEndPos)
	{
		cBuffer[length++] = 0x00;
	}

	for( i = lFWStartPos; i < lFWEndPos; i += 0x100 )
	{
		if( CmdFirmwareWrite(&CSerial9525, 0, (i-lFWStartPos)/0x100, 0x100, cBuffer + i) != 1 )
		{
			ResultColor = RGB(255,0,0);
			m_Result.SetTextColor(ResultColor);
			m_Result.SetWindowText("FAILED !");
			MessageBox("Error! Failed to write the firmware!", "Error", MB_OK);
			free(strMessage);
			return;
		}
		if( i & 0x1800 )
			m_Result.SetWindowText("Wait ....");
		else
			m_Result.SetWindowText("");
		m_Prog->SetPos(100*(i-lFWStartPos)/(lFWEndPos-lFWStartPos));
	}
	Sleep(10);
	//Verify 
	ResultColor = RGB(0,0,255);
	m_Result.SetTextColor(ResultColor);
	m_Result.SetWindowText("Wait ....");

	m_Status->SetWindowText("Read/Verify Firmware");
	m_Prog->SetPos(0);
	for( i = lFWStartPos; i < lFWEndPos; i += 0x100 )
	{
		UCHAR bData[0x100];
		ULONG lReturnLen;
		if( CmdFirmwareRead(&CSerial9525, 0, (i-lFWStartPos)/0x100, 0x100, bData, &lReturnLen) != 1 )
		{
			ResultColor = RGB(255,0,0);
			m_Result.SetTextColor(ResultColor);
			m_Result.SetWindowText("FAILED !");

			MessageBox("Error! Failed to read the firmware!", "Error", MB_OK);
			free(strMessage);
			return;			
		}
		
		for(iAddr=0; iAddr<0x100; iAddr++)
		{
			if(bData[iAddr] != cBuffer[i+iAddr] )
			{
				break;
			}
		}
		if(0x100!=iAddr)
		{
			ResultColor = RGB(255,0,0);
			m_Result.SetTextColor(ResultColor);
			m_Result.SetWindowText("FAILED !");

			MessageBox("Error! W/R compare failed!", "Error", MB_OK);
			free(strMessage);
			return;			
		}
		if( i & 0x1800 )
			m_Result.SetWindowText("Wait ....");
		else
			m_Result.SetWindowText("");

		m_Prog->SetPos(100*(i-lFWStartPos)*0x100/(lFWEndPos-lFWStartPos));
	}

	m_Result.SetWindowText("success");		
	ResultColor = RGB(0,255,0);
	m_Result.SetTextColor(ResultColor);
#endif
	MessageBox("SUCCESS","DemoTool",MB_OK|MB_ICONINFORMATION);

	if( CmdTestFwUpDate(&CSerial9525) != 1 )
	{
		MessageBox("Error! Failed to run on coderam !", "Error", MB_OK);
		free(strMessage);
		return;			
	}
	MessageBox("Success to run on coderam!","DemoTool",MB_OK|MB_ICONINFORMATION);

	free(strMessage);
	return;

}









