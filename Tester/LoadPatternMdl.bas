Attribute VB_Name = "LoadPatternMdl"
Sub loadPatternForMp3()


Open App.Path & "\Pattern.txt" For Input As #2

    Do While Not EOF(2)
    
    
        Input #2, Pattern(i)
         
        i = i + 1
    Loop

Close #2
 
'================== 64k ,128 sector
 i = 0
Open App.Path & "\AU6982_f1_s2.txt" For Input As #2

    Do While Not EOF(2)
    
         Pattern_AU6377(i) = &HFF
        Input #2, Pattern_AU6982(i)
        
         
        i = i + 1
    Loop

Close #2



'================== 64k ,128 sector
 i = 0
Open App.Path & "\AU6375_ram_unstable.txt" For Input As #2

    Do While Not EOF(2)
    
         Pattern_AU6375(i) = &HFF
        Input #2, Pattern_AU6375(i)
        
         
        i = i + 1
    Loop

Close #2
 
 
 i = 0
 Open App.Path & "\AU6981Pattern.txt" For Input As #2

    Do While Not EOF(2)
    
    
        Input #2, AU6981Pattern(i)
         
        i = i + 1
    Loop

Close #2
 
 
 
 
 
'==================== AU3130A
 
Open App.Path & "\MP3_A.txt" For Input As #3
Open App.Path & "\MP31_A.txt" For Input As #4
Open App.Path & "\MP32_A.txt" For Input As #5
Open App.Path & "\MP33_A.txt" For Input As #6
Open App.Path & "\MP34_A.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_A(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7


'======================= Au3130B43 1st  mode ===================================
Open App.Path & "\MP3_B.txt" For Input As #3
Open App.Path & "\MP31_B.txt" For Input As #4
Open App.Path & "\MP32_B.txt" For Input As #5
Open App.Path & "\MP33_B.txt" For Input As #6
Open App.Path & "\MP34_B.txt" For Input As #7

   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

'============================ AU3130B43 2nd mode ============================
Open App.Path & "\MP3_B1.txt" For Input As #3
Open App.Path & "\MP31_B1.txt" For Input As #4
Open App.Path & "\MP32_B1.txt" For Input As #5
Open App.Path & "\MP33_B1.txt" For Input As #6
Open App.Path & "\MP34_B1.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B1(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7
 
'======================== AU3130B43 3rd mode ====================

Open App.Path & "\MP3_B2.txt" For Input As #3
Open App.Path & "\MP31_B2.txt" For Input As #4
Open App.Path & "\MP32_B2.txt" For Input As #5
Open App.Path & "\MP33_B2.txt" For Input As #6
Open App.Path & "\MP34_B2.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B2(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

'====================================================================

Open App.Path & "\MP3_C.txt" For Input As #3
Open App.Path & "\MP31_C.txt" For Input As #4
Open App.Path & "\MP32_C.txt" For Input As #5
Open App.Path & "\MP33_C.txt" For Input As #6
Open App.Path & "\MP34_C.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

 
 '====================================================================

Open App.Path & "\MP3_C1.txt" For Input As #3
Open App.Path & "\MP31_C1.txt" For Input As #4
Open App.Path & "\MP32_C1.txt" For Input As #5
Open App.Path & "\MP33_C1.txt" For Input As #6
Open App.Path & "\MP34_C1.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C1(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7



  
 '====================================================================

Open App.Path & "\MP3_C2.txt" For Input As #3
Open App.Path & "\MP31_C2.txt" For Input As #4
Open App.Path & "\MP32_C2.txt" For Input As #5
Open App.Path & "\MP33_C2.txt" For Input As #6
Open App.Path & "\MP34_C2.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C2(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7


'=================================================================
'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_BL.txt" For Input As #3
 




   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_BL(i) = CInt(CSng(tmp0))
         
   Next i

Close #3
 
'=================================================================
'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_BL1.txt" For Input As #3
 




   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_BL1(i) = CInt(CSng(tmp0))
         
   Next i

Close #3


'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_CL.txt" For Input As #3

   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_CL(i) = CInt(CSng(tmp0))
         
   Next i

Close #3

'==================================================
  Open App.Path & "\MP3_CW1.txt" For Input As #3
   For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_CW1(i) = CInt(CSng(tmp0))
         
   Next i
Close #3

'============================================

   Open App.Path & "\MP3_3150J.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150J(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 '============================================

   Open App.Path & "\MP3_3150J1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150J1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3


'============================================

   Open App.Path & "\MP3_3152A1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3152A1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3152A2.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3152A2(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 '============================================
   Open App.Path & "\MP3_3152AL23.txt" For Input As #3
    Open App.Path & "\MP3_3152AL231.txt" For Input As #4
      Open App.Path & "\MP3_3152AL232.txt" For Input As #5
    For i = 0 To 99
    
        Input #3, tmp0
         Input #4, tmp1
          Input #5, tmp2
          MP3Data_3152A3(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp2)) * 0.25)
         
    Next i
    Close #5
 Close #4
  Close #3
'==============================================
' for AU3150ALF22,AU3150ALF22
'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A221(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A222(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3




' for AU3150ALF22,AU3150ALF22
'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A221(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A222(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

' for AU3150AKL ,
'============================================

   Open App.Path & "\MP3_3150KL1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150KL2.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL2(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

 'for AU3150BKL ,
'============================================

   Open App.Path & "\MP3_3150KL21.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL21(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150KL22.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL22(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

  Open App.Path & "\MP3_3150KL23.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL23(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================


 Open App.Path & "\MP3_AU62541.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_AU6254(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

 

'==============================================


 Open App.Path & "\MP3_AU62542.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_AU62541(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================
End Sub


Sub loadPatternForMp5()


Open App.Path & "\Pattern.txt" For Input As #2

    Do While Not EOF(2)
    
    
        Input #2, Pattern(i)
         
        i = i + 1
    Loop

Close #2
 
'================== 64k ,128 sector
 i = 0
Open App.Path & "\AU6982_f1_s2.txt" For Input As #2

    Do While Not EOF(2)
    
         Pattern_AU6377(i) = &HFF
        Input #2, Pattern_AU6982(i)
        
         
        i = i + 1
    Loop

Close #2



'================== 64k ,128 sector
 i = 0
Open App.Path & "\AU6375_ram_unstable.txt" For Input As #2

    Do While Not EOF(2)
    
         Pattern_AU6375(i) = &HFF
        Input #2, Pattern_AU6375(i)
        
         
        i = i + 1
    Loop

Close #2
 
 
 i = 0
 Open App.Path & "\AU6981Pattern.txt" For Input As #2

    Do While Not EOF(2)
    
    
        Input #2, AU6981Pattern(i)
         
        i = i + 1
    Loop

Close #2
 
 
 
 
 
'==================== AU3130A
 
Open App.Path & "\MP3_A.txt" For Input As #3
Open App.Path & "\MP31_A.txt" For Input As #4
Open App.Path & "\MP32_A.txt" For Input As #5
Open App.Path & "\MP33_A.txt" For Input As #6
Open App.Path & "\MP34_A.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_A(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7


'======================= Au3130B43 1st  mode ===================================
Open App.Path & "\MP3_B.txt" For Input As #3
Open App.Path & "\MP31_B.txt" For Input As #4
Open App.Path & "\MP32_B.txt" For Input As #5
Open App.Path & "\MP33_B.txt" For Input As #6
Open App.Path & "\MP34_B.txt" For Input As #7

   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

'============================ AU3130B43 2nd mode ============================
Open App.Path & "\MP3_B1.txt" For Input As #3
Open App.Path & "\MP31_B1.txt" For Input As #4
Open App.Path & "\MP32_B1.txt" For Input As #5
Open App.Path & "\MP33_B1.txt" For Input As #6
Open App.Path & "\MP34_B1.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B1(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7
 
'======================== AU3130B43 3rd mode ====================

Open App.Path & "\MP3_B2.txt" For Input As #3
Open App.Path & "\MP31_B2.txt" For Input As #4
Open App.Path & "\MP32_B2.txt" For Input As #5
Open App.Path & "\MP33_B2.txt" For Input As #6
Open App.Path & "\MP34_B2.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_B2(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

'====================================================================

Open App.Path & "\MP3_C.txt" For Input As #3
Open App.Path & "\MP31_C.txt" For Input As #4
Open App.Path & "\MP32_C.txt" For Input As #5
Open App.Path & "\MP33_C.txt" For Input As #6
Open App.Path & "\MP34_C.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7

 
 '====================================================================

Open App.Path & "\MP3_C1.txt" For Input As #3
Open App.Path & "\MP31_C1.txt" For Input As #4
Open App.Path & "\MP32_C1.txt" For Input As #5
Open App.Path & "\MP33_C1.txt" For Input As #6
Open App.Path & "\MP34_C1.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C1(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7



  
 '====================================================================

Open App.Path & "\MP3_C2.txt" For Input As #3
Open App.Path & "\MP31_C2.txt" For Input As #4
Open App.Path & "\MP32_C2.txt" For Input As #5
Open App.Path & "\MP33_C2.txt" For Input As #6
Open App.Path & "\MP34_C2.txt" For Input As #7




   For i = 0 To 99
    
        Input #3, tmp0
        Input #4, tmp1
        Input #5, tmp2
        Input #6, tmp3
        Input #7, tmp4
        MP3Data_C2(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp3) + CSng(tmp4)) * 0.2)
         
   Next i

Close #3
Close #4
Close #5
Close #6
Close #7


'=================================================================
'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_BL.txt" For Input As #3
 




   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_BL(i) = CInt(CSng(tmp0))
         
   Next i

Close #3
 
'=================================================================
'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_BL1.txt" For Input As #3
 




   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_BL1(i) = CInt(CSng(tmp0))
         
   Next i

Close #3


'========= AU3130BLF20 1st mode
Open App.Path & "\MP3_CL.txt" For Input As #3

   For i = 0 To 99
    
        Input #3, tmp0
         MP3Data_CL(i) = CInt(CSng(tmp0))
         
   Next i

Close #3

'==================================================
  Open App.Path & "\MP3_CW1.txt" For Input As #3
   For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_CW1(i) = CInt(CSng(tmp0))
         
   Next i
Close #3

'============================================

   Open App.Path & "\MP3_3150J.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150J(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 '============================================

   Open App.Path & "\MP3_3150J1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150J1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3


'============================================

   Open App.Path & "\MP3_3152A1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3152A1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3152A2.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3152A2(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
 
 '============================================
   Open App.Path & "\MP3_3152AL23.txt" For Input As #3
    Open App.Path & "\MP3_3152AL231.txt" For Input As #4
      Open App.Path & "\MP3_3152AL232.txt" For Input As #5
    For i = 0 To 99
    
        Input #3, tmp0
         Input #4, tmp1
          Input #5, tmp2
          MP3Data_3152A3(i) = CInt((CSng(tmp0) + CSng(tmp1) + CSng(tmp2) + CSng(tmp2)) * 0.25)
         
    Next i
    Close #5
 Close #4
  Close #3
'==============================================
' for AU3150ALF22,AU3150ALF22
'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A221(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A222(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3




' for AU3150ALF22,AU3150ALF22
'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A221(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150ALF221.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150A222(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

' for AU3150AKL ,
'============================================

   Open App.Path & "\MP3_3150KL1.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL1(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150KL2.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL2(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

 'for AU3150BKL ,
'============================================

   Open App.Path & "\MP3_3150KL21.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL21(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

'============================================

   Open App.Path & "\MP3_3150KL22.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL22(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================

  Open App.Path & "\MP3_3150KL23.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_3150KL23(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================


 Open App.Path & "\MP3_AU62541.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_AU6254(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3

 

'==============================================


 Open App.Path & "\MP3_AU62542.txt" For Input As #3
    For i = 0 To 99
    
        Input #3, tmp0

          MP3Data_AU62541(i) = CInt(CSng(tmp0))
         
    Next i
 Close #3
'==============================================
End Sub

