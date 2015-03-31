# Microsoft Developer Studio Project File - Name="9525COMAP" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=9525COMAP - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "9525COMAP.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "9525COMAP.mak" CFG="9525COMAP - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "9525COMAP - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE "9525COMAP - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "9525COMAP - Win32 Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_AFXDLL" /FR /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /machine:I386
# ADD LINK32 /nologo /subsystem:windows /machine:I386 /out:"Release/SmartCard_COM_Mode_Tester.exe"

!ELSEIF  "$(CFG)" == "9525COMAP - Win32 Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_AFXDLL" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 /nologo /subsystem:windows /debug /machine:I386 /out:"Debug/SmartCard_COM_Mode_Tester.exe" /pdbtype:sept

!ENDIF 

# Begin Target

# Name "9525COMAP - Win32 Release"
# Name "9525COMAP - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\9525COMAP.cpp
# End Source File
# Begin Source File

SOURCE=.\9525COMAP.rc
# End Source File
# Begin Source File

SOURCE=.\9525COMAPDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\9525RS232Lib.cpp
# End Source File
# Begin Source File

SOURCE="C:\Program Files\Microsoft Visual Studio\VC98\MFC\SRC\APPMODUL.CPP"
# End Source File
# Begin Source File

SOURCE=.\AT24CDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\AT45D041Dlg.cpp
# End Source File
# Begin Source File

SOURCE=.\AT88SCcmdDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\AT88SCDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\font1.cpp
# End Source File
# Begin Source File

SOURCE=.\InphoneCmdDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\Label.cpp
# End Source File
# Begin Source File

SOURCE=.\labelcontrol1.cpp
# End Source File
# Begin Source File

SOURCE=.\LCM_KEYPAD_EEPROMDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\mscomm1.cpp
# End Source File
# Begin Source File

SOURCE=.\picture1.cpp
# End Source File
# Begin Source File

SOURCE=.\SLE4428Dlg.cpp
# End Source File
# Begin Source File

SOURCE=.\SLE4442CMDDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\SLE4442Dlg.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\9525COMAP.h
# End Source File
# Begin Source File

SOURCE=.\9525COMAPDlg.h
# End Source File
# Begin Source File

SOURCE=.\9525RS232Lib.h
# End Source File
# Begin Source File

SOURCE=.\AT24CDlg.h
# End Source File
# Begin Source File

SOURCE=.\AT45D041Dlg.h
# End Source File
# Begin Source File

SOURCE=.\AT88SCcmdDlg.h
# End Source File
# Begin Source File

SOURCE=.\AT88SCDlg.h
# End Source File
# Begin Source File

SOURCE=.\font1.h
# End Source File
# Begin Source File

SOURCE=.\InphoneCmdDlg.h
# End Source File
# Begin Source File

SOURCE=.\Label.h
# End Source File
# Begin Source File

SOURCE=.\labelcontrol1.h
# End Source File
# Begin Source File

SOURCE=.\LCM_KEYPAD_EEPROMDlg.h
# End Source File
# Begin Source File

SOURCE=.\mscomm1.h
# End Source File
# Begin Source File

SOURCE=.\picture1.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\SLE4428Dlg.h
# End Source File
# Begin Source File

SOURCE=.\SLE4442CMDDlg.h
# End Source File
# Begin Source File

SOURCE=.\SLE4442Dlg.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE="C:\Program Files\Microsoft Visual Studio\VC98\Include\WINDEF.H"
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\res\9525COMAP.ico
# End Source File
# Begin Source File

SOURCE=.\res\9525COMAP.rc2
# End Source File
# Begin Source File

SOURCE=.\Alcor.ico
# End Source File
# End Group
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# End Target
# End Project
# Section 9525COMAP : {7BF80981-BF32-101A-8BBB-00AA00300CAB}
# 	2:5:Class:CPicture
# 	2:10:HeaderFile:picture1.h
# 	2:8:ImplFile:picture1.cpp
# End Section
# Section 9525COMAP : {04598FC1-866C-11CF-AB7C-00AA00C08FCF}
# 	2:5:Class:CLabelControl
# 	2:10:HeaderFile:labelcontrol1.h
# 	2:8:ImplFile:labelcontrol1.cpp
# End Section
# Section 9525COMAP : {648A5600-2C6E-101B-82B6-000000000014}
# 	2:21:DefaultSinkHeaderFile:mscomm1.h
# 	2:16:DefaultSinkClass:CMSComm
# End Section
# Section 9525COMAP : {E6E17E90-DF38-11CF-8E74-00A0C90F26F8}
# 	2:5:Class:CMSComm
# 	2:10:HeaderFile:mscomm1.h
# 	2:8:ImplFile:mscomm1.cpp
# End Section
# Section 9525COMAP : {978C9E23-D4B0-11CE-BF2D-00AA003F40D0}
# 	2:21:DefaultSinkHeaderFile:labelcontrol1.h
# 	2:16:DefaultSinkClass:CLabelControl
# End Section
# Section 9525COMAP : {BEF6E003-A874-101A-8BBA-00AA00300CAB}
# 	2:5:Class:COleFont
# 	2:10:HeaderFile:font1.h
# 	2:8:ImplFile:font1.cpp
# End Section
