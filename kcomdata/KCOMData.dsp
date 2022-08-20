# Microsoft Developer Studio Project File - Name="KCOMData" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=KCOMData - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "KCOMData.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "KCOMData.mak" CFG="KCOMData - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "KCOMData - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "KCOMData - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "KCOMData - Win32 Unicode Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "KCOMData - Win32 Unicode Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "KCOMData - Win32 Release"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 1
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Target_Ext "ocx"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /Od /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "_WINDLL" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /win32
# ADD BASE RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386
# SUBTRACT LINK32 /debug
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Desc=copying file to share directory
PostBuild_Cmds=copy release\*.ocx \temp
# End Special Build Tool

!ELSEIF  "$(CFG)" == "KCOMData - Win32 Debug"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Target_Ext "ocx"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_MBCS" /D "_USRDLL" /D "_GXDLL" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /win32
# ADD BASE RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\Debug
TargetPath=.\Debug\KCOMData.ocx
InputPath=.\Debug\KCOMData.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "KCOMData - Win32 Unicode Debug"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "DebugU"
# PROP BASE Intermediate_Dir "DebugU"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "DebugU"
# PROP Intermediate_Dir "DebugU"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_MBCS" /D "_USRDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_USRDLL" /D "_UNICODE" /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /win32
# ADD BASE RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\DebugU
TargetPath=.\DebugU\KCOMData.ocx
InputPath=.\DebugU\KCOMData.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "KCOMData - Win32 Unicode Release"

# PROP BASE Use_MFC 2
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "ReleaseU"
# PROP BASE Intermediate_Dir "ReleaseU"
# PROP BASE Target_Ext "ocx"
# PROP BASE Target_Dir ""
# PROP Use_MFC 2
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "ReleaseU"
# PROP Intermediate_Dir "ReleaseU"
# PROP Target_Ext "ocx"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_MBCS" /D "_USRDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_USRDLL" /D "_UNICODE" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /win32
# ADD BASE RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x804 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\ReleaseU
TargetPath=.\ReleaseU\KCOMData.ocx
InputPath=.\ReleaseU\KCOMData.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ENDIF 

# Begin Target

# Name "KCOMData - Win32 Release"
# Name "KCOMData - Win32 Debug"
# Name "KCOMData - Win32 Unicode Debug"
# Name "KCOMData - Win32 Unicode Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AddColumnDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\ComboEdit.cpp
# End Source File
# Begin Source File

SOURCE=.\DropDownColumn.cpp
# End Source File
# Begin Source File

SOURCE=.\DropDownColumns.cpp
# End Source File
# Begin Source File

SOURCE=.\DropDownRealWnd.cpp
# End Source File
# Begin Source File

SOURCE=.\FieldsDLg.cpp
# End Source File
# Begin Source File

SOURCE=.\GridColumn.cpp
# End Source File
# Begin Source File

SOURCE=.\GridColumns.cpp
# End Source File
# Begin Source File

SOURCE=.\GridInner.cpp
# End Source File
# Begin Source File

SOURCE=.\GridInPpg.cpp
# End Source File
# Begin Source File

SOURCE=.\GridInput.cpp
# End Source File
# Begin Source File

SOURCE=.\GridInputDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMComboBox.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMData.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMData.def
# End Source File
# Begin Source File

SOURCE=.\KCOMData.odl
# ADD MTL /h "KCOMData_i.h" /iid "KCOMData_i.c" /Oicf
# End Source File
# Begin Source File

SOURCE=.\KCOMData.rc
# End Source File
# Begin Source File

SOURCE=.\KCOMDateControl.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMDrawingTools.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMEditBtn.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMEditControl.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMMaskControl.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichComboCtl.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichComboPpg.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichDropDownCtl.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichDropDownPpg.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichGridColumnPpg.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichGridCtl.cpp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichGridPpg.cpp
# End Source File
# Begin Source File

SOURCE=.\SelBookmarks.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\AddColumnDlg.h
# End Source File
# Begin Source File

SOURCE=.\ComboEdit.h
# End Source File
# Begin Source File

SOURCE=.\DropDownColumn.h
# End Source File
# Begin Source File

SOURCE=.\DropDownColumns.h
# End Source File
# Begin Source File

SOURCE=.\DropDownRealWnd.h
# End Source File
# Begin Source File

SOURCE=.\FieldsDLg.h
# End Source File
# Begin Source File

SOURCE=.\GridColumn.h
# End Source File
# Begin Source File

SOURCE=.\GridColumns.h
# End Source File
# Begin Source File

SOURCE=.\GridInner.h
# End Source File
# Begin Source File

SOURCE=.\GridInPpg.h
# End Source File
# Begin Source File

SOURCE=.\GridInput.h
# End Source File
# Begin Source File

SOURCE=.\GridInputDlg.h
# End Source File
# Begin Source File

SOURCE=.\KCOMComboBox.h
# End Source File
# Begin Source File

SOURCE=.\KCOMData.h
# End Source File
# Begin Source File

SOURCE=.\KCOMDateControl.h
# End Source File
# Begin Source File

SOURCE=.\KCOMDrawingTools.h
# End Source File
# Begin Source File

SOURCE=.\KCOMEditBtn.h
# End Source File
# Begin Source File

SOURCE=.\KCOMEditControl.h
# End Source File
# Begin Source File

SOURCE=.\KCOMMaskControl.h
# End Source File
# Begin Source File

SOURCE=.\KCOMRichComboCtl.h
# End Source File
# Begin Source File

SOURCE=.\KCOMRichComboPpg.h
# End Source File
# Begin Source File

SOURCE=.\KCOMRichDropDownCtl.h
# End Source File
# Begin Source File

SOURCE=.\KCOMRichDropDownPpg.h
# End Source File
# Begin Source File

SOURCE=.\KCOMRichGridColumnPpg.h
# End Source File
# Begin Source File

SOURCE=.\KCOMRichGridCtl.h
# End Source File
# Begin Source File

SOURCE=.\KCOMRichGridPpg.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\SelBookmarks.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\KCOMData.rgs
# End Source File
# Begin Source File

SOURCE=.\KCOMRichComboCtl.bmp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichDropDownCtl.bmp
# End Source File
# Begin Source File

SOURCE=.\KCOMRichGridCtl.bmp
# End Source File
# End Group
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# End Target
# End Project
# Section KCOMData : {00000000-0000-0000-0000-000000000000}
# 	1:12:IDR_KCOMDATA:102
# End Section
