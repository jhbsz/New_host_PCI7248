Attribute VB_Name = "Module1"
Global card As Integer
Global card_number As Integer
Global AlarmCtrl As Byte

Public Declare Function CAndValue Lib "DLL7.dll" (ByVal In1 As Byte, ByVal In2 As Byte) As Integer
