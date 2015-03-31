Attribute VB_Name = "DeclMdl"
Option Explicit
Public Declare Function LowLevelFormat Lib "CustomMp.dll" (var1 As Any, var2 As Any, var3 As Any) As Byte
Public Declare Function CustomMP Lib "CustomMp.dll" (ByRef Inv As Byte, ByRef var1 As Byte, ByRef var2 As Byte, ByRef var3 As Byte) As Byte
Public Declare Function fnScsi2usb2K_KillEXE Lib "usbreset.dll" () As Integer
Public Declare Function SetDebugLevel Lib "CustomMp.dll" (ByRef DebugLevel As Integer) 'Public Declare Function fnScsi2usbME_KillEXE Lib "usbreset" () As Integer
Public Declare Function ReaderTester Lib "TestReader.dll" _
(ByRef CBW As Byte, ByRef Data As Byte, ByRef CSW As Byte) As Integer
Public Declare Function ReaderTester2 Lib "TestReader.dll" _
(ByRef CBW As Byte, ByRef Data As Byte, ByRef CSW As Byte) As Integer



