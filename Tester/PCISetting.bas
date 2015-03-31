Attribute VB_Name = "PCISetting"
Option Explicit
Public Function PCI7248_AU6476_ALPSExist() As Integer
Dim result As Integer

  card = Register_Card(PCI_7248, 0) 'FOR PCI_7248
     
    
    If card < 0 Then 'FOR PCI_7248
       MsgBox "Register Card Failed" 'FOR PCI_7248
    '   End 'FOR PCI_7248
    End If 'FOR PCI_7248
    
    

result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1A as output card fail"
    End
End If


result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1B as output card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1CH as output card fail"
    End
End If

result = DIO_PortConfig(card, Channel_P1CL, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1CL as output card fail"
    End
End If



result = DIO_PortConfig(card, Channel_P2A, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2A as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2B, INPUT_PORT)

If result <> 0 Then
    MsgBox " config PCI_P2A as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2CH, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2CH as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2CL, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2CL as input card fail"
    End
End If
PCI7248InitFinish = 1
PCI7248_AU6476_ALPSExist = 1
End Function
Public Function PCI7248ExistAU6256() As Integer
Dim result As Integer

  card = Register_Card(PCI_7248, 0) 'FOR PCI_7248
     
    
    If card < 0 Then 'FOR PCI_7248
       MsgBox "Register Card Failed" 'FOR PCI_7248
    '   End 'FOR PCI_7248
    End If 'FOR PCI_7248
    
    

result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1A as output card fail"
    End
End If


result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1B as output card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P1C, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1CH as output card fail"
    End
End If

 


result = DIO_PortConfig(card, Channel_P2A, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2A as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2B, INPUT_PORT)

If result <> 0 Then
    MsgBox " config PCI_P2A as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2CH, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2CH as input card fail"
    End
End If
 
PCI7248InitFinish = 1
PCI7248ExistAU6256 = 1
End Function

Public Function PCI7248Exist() As Integer
Dim result As Integer

  card = Register_Card(PCI_7248, 0) 'FOR PCI_7248
     
    
    If card < 0 Then 'FOR PCI_7248
       MsgBox "Register Card Failed" 'FOR PCI_7248
    '   End 'FOR PCI_7248
    End If 'FOR PCI_7248
    
    

result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1A as output card fail"
    End
End If


result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1B as output card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1CH as output card fail"
    End
End If

result = DIO_PortConfig(card, Channel_P1CL, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1CL as output card fail"
    End
End If



result = DIO_PortConfig(card, Channel_P2A, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2A as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2B, INPUT_PORT)

If result <> 0 Then
    MsgBox " config PCI_P2A as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2CH, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2CH as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2CL, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2CL as input card fail"
    End
End If
PCI7248InitFinish = 1
PCI7248Exist = 1
End Function

Public Function PCI7248Exist_P1C_Sync() As Integer
Dim result As Integer


    card = Register_Card(PCI_7248, 0)   'FOR PCI_7248
    
    If card < 0 Then                    'FOR PCI_7248
       MsgBox "Register Card Failed"    'FOR PCI_7248
    End If                              'FOR PCI_7248
    
    result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P1A as output card fail"
        End
    End If

    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P1B as input card fail"
        End
    End If

    If GPIBCard_Exist Then
        result = DIO_PortConfig(card, Channel_P1CH, OUTPUT_PORT)
        If result <> 0 Then
            MsgBox " config PCI_P1CH as output card fail"
            End
        End If
        
        result = DIO_PortConfig(card, Channel_P1CL, INPUT_PORT)
        If result <> 0 Then
            MsgBox " config PCI_P1CL as input card fail"
            End
        End If
    Else
        result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
        If result <> 0 Then
            MsgBox " config PCI_P1CH as input card fail"
            End
        End If
        
        result = DIO_PortConfig(card, Channel_P1CL, OUTPUT_PORT)
        If result <> 0 Then
            MsgBox " config PCI_P1CL as output card fail"
            End
        End If
    End If



    result = DIO_PortConfig(card, Channel_P2A, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P2A as input card fail"
        End
    End If
    result = DIO_PortConfig(card, Channel_P2B, INPUT_PORT)

    If result <> 0 Then
        MsgBox " config PCI_P2A as input card fail"
        End
    End If
    result = DIO_PortConfig(card, Channel_P2CH, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P2CH as input card fail"
        End
    End If
    result = DIO_PortConfig(card, Channel_P2CL, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P2CL as input card fail"
        End
    End If
    PCI7248InitFinish_Sync = 1
    PCI7248Exist_P1C_Sync = 1
    
End Function

Public Function PCI7248ExistAU6254() As Integer
Dim result As Integer

  card = Register_Card(PCI_7248, 0) 'FOR PCI_7248
     
    
    If card < 0 Then 'FOR PCI_7248
       MsgBox "Register Card Failed" 'FOR PCI_7248
    '   End 'FOR PCI_7248
    End If 'FOR PCI_7248
    
    

result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1A as output card fail"
    End
End If


result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1B as output card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P1CH, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1CH as output card fail"
    End
End If

result = DIO_PortConfig(card, Channel_P1CL, OUTPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P1CL as output card fail"
    End
End If



result = DIO_PortConfig(card, Channel_P2A, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2A as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2B, INPUT_PORT)

If result <> 0 Then
    MsgBox " config PCI_P2A as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2CH, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2CH as input card fail"
    End
End If
result = DIO_PortConfig(card, Channel_P2CL, INPUT_PORT)
If result <> 0 Then
    MsgBox " config PCI_P2CL as input card fail"
    End
End If
PCI7248InitFinish = 1
PCI7248ExistAU6254 = 1
End Function

