VERSION 5.00
Begin VB.Form OpenShortFrm 
   AutoRedraw      =   -1  'True
   Caption         =   "OpenShortResult"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton SaveBtn 
      Caption         =   "Save"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "OpenShortFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SaveBtn_Click()
Open App.Path & "\OSStandard\OSDebug.txt" For Output As #5
Print #5, OpenShortPinNo
Dim i As Integer
For i = 0 To 127
Print #5, i, OSValue(i)
Next i
Close #5
MsgBox "Save OSDebug.txt Finish"

End Sub
