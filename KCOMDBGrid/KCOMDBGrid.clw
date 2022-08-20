; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CKCOMDBGridCtrl
LastTemplate=COlePropertyPage
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "kcomdbgrid.h"
LastPage=0
CDK=Y

ClassCount=8
Class1=CGridCtrl
Class2=CGridInner
Class3=CInPlaceEdit
Class4=CComboEdit
Class5=CInPlaceList
Class6=CKCOMDBGridCtrl
Class7=CKCOMDBGridPropPage

ResourceCount=3
Resource1=IDD_ABOUTBOX_KCOMDBGRID
Resource2=IDD_PROPPAGE_KCOMDBGRID
Class8=CKCOMDBGridColumnsPpg
Resource3=IDD_PROPPAGE_KCOMDBGRID_COLUMNS

[CLS:CGridCtrl]
Type=0
BaseClass=CWnd
HeaderFile=GridCtrl.h
ImplementationFile=GridCtrl.cpp
LastObject=CGridCtrl

[CLS:CGridInner]
Type=0
BaseClass=CGridCtrl
HeaderFile=GridInner.h
ImplementationFile=GridInner.cpp
LastObject=CGridInner

[CLS:CInPlaceEdit]
Type=0
BaseClass=CEdit
HeaderFile=InPlaceEdit.h
ImplementationFile=InPlaceEdit.cpp

[CLS:CComboEdit]
Type=0
BaseClass=CEdit
HeaderFile=InPlaceList.h
ImplementationFile=InPlaceList.cpp
LastObject=CComboEdit

[CLS:CInPlaceList]
Type=0
BaseClass=CComboBox
HeaderFile=InPlaceList.h
ImplementationFile=InPlaceList.cpp

[CLS:CKCOMDBGridCtrl]
Type=0
BaseClass=COleControl
HeaderFile=KCOMDBGridCtl.h
ImplementationFile=KCOMDBGridCtl.cpp
LastObject=CKCOMDBGridCtrl
Filter=W
VirtualFilter=wWC

[CLS:CKCOMDBGridPropPage]
Type=0
BaseClass=COlePropertyPage
HeaderFile=KCOMDBGridPpg.h
ImplementationFile=KCOMDBGridPpg.cpp

[DLG:IDD_PROPPAGE_KCOMDBGRID]
Type=1
Class=CKCOMDBGridPropPage
ControlCount=23
Control1=IDC_STATIC,button,1342177287
Control2=IDC_STATIC,static,1342308352
Control3=IDC_COMBO_DATAMODE,combobox,1344339971
Control4=IDC_CHECK_SHOWRECORDSELECTOR,button,1342242851
Control5=IDC_CHECK_ALLOWADDNEW,button,1342242851
Control6=IDC_CHECK_ALLOWDELETE,button,1342242851
Control7=IDC_CHECK_READONLY,button,1342242851
Control8=IDC_CHECK_ALLOWARROWKEYS,button,1342242851
Control9=IDC_CHECK_ALLOWAROWRESIZING,button,1342242851
Control10=IDC_STATIC,static,1342308352
Control11=IDC_EDIT_ROWS,edit,1350639746
Control12=IDC_SPIN_ROWS,msctls_updown32,1342177462
Control13=IDC_STATIC,static,1342308352
Control14=IDC_EDIT_COLS,edit,1350639746
Control15=IDC_SPIN_COLS,msctls_updown32,1342177462
Control16=IDC_STATIC,static,1342308352
Control17=IDC_EDIT_DEFCOLWIDTH,edit,1350639744
Control18=IDC_STATIC,static,1342308352
Control19=IDC_EDIT_HEADERHEIGHT,edit,1350639744
Control20=IDC_STATIC,static,1342308352
Control21=IDC_EDIT_ROWHEIGHT,edit,1350639744
Control22=IDC_STATIC,static,1342308352
Control23=IDC_COMBO_GRIDLINESTYLE,combobox,1344339971

[DLG:IDD_ABOUTBOX_KCOMDBGRID]
Type=1
Class=?
ControlCount=5
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308352
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889
Control5=IDC_STATIC,static,1342308352

[DLG:IDD_PROPPAGE_KCOMDBGRID_COLUMNS]
Type=1
Class=CKCOMDBGridColumnsPpg
ControlCount=7
Control1=IDC_STATIC,static,1342308352
Control2=IDC_COMBO_COLUMN,combobox,1344339971
Control3=IDC_STATIC,static,1342308352
Control4=IDC_EDIT_HEADER,edit,1350631552
Control5=IDC_STATIC,static,1342308352
Control6=IDC_COMBO_DATAFIELD,combobox,1344339971
Control7=IDC_BUTTON_RESET,button,1342242816

[CLS:CKCOMDBGridColumnsPpg]
Type=0
HeaderFile=KCOMDBGridColumnsPpg.h
ImplementationFile=KCOMDBGridColumnsPpg.cpp
BaseClass=COlePropertyPage
Filter=D
LastObject=CKCOMDBGridColumnsPpg
VirtualFilter=idWC

