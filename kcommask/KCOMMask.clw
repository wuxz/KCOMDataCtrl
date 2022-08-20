; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CKCOMMaskCtrl
LastTemplate=CEdit
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "KCOMMask.h"
CDK=Y

ClassCount=3
Class1=CKCOMMaskCtrl
Class2=CKCOMMaskPropPage

ResourceCount=2
LastPage=0
Resource1=IDD_ABOUTBOX_KCOMMASK
Class3=CMaskEdit
Resource2=IDD_PROPPAGE_KCOMMASK

[CLS:CKCOMMaskCtrl]
Type=0
HeaderFile=KCOMMaskCtl.h
ImplementationFile=KCOMMaskCtl.cpp
Filter=W
BaseClass=COleControl
VirtualFilter=wWC
LastObject=CKCOMMaskCtrl

[CLS:CKCOMMaskPropPage]
Type=0
HeaderFile=KCOMMaskPpg.h
ImplementationFile=KCOMMaskPpg.cpp
Filter=D
LastObject=CKCOMMaskPropPage
BaseClass=COlePropertyPage
VirtualFilter=idWC

[DLG:IDD_ABOUTBOX_KCOMMASK]
Type=1
Class=?
ControlCount=4
Control1=IDC_STATIC,static,1342177294
Control2=IDC_STATIC,static,1342308352
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889

[DLG:IDD_PROPPAGE_KCOMMASK]
Type=1
Class=CKCOMMaskPropPage
ControlCount=12
Control1=IDC_STATIC,static,1342308352
Control2=IDC_EDIT_MASK,edit,1350631552
Control3=IDC_STATIC,static,1342308352
Control4=IDC_COMBO_APPEARANCE,combobox,1344339971
Control5=IDC_STATIC,static,1342308352
Control6=IDC_COMBO_BORDERSTYLE,combobox,1344339971
Control7=IDC_CHECK_AUTOTATB,button,1342242819
Control8=IDC_CHECK_PROMPTINCLUDE,button,1342242819
Control9=IDC_STATIC,static,1342308352
Control10=IDC_EDIT_MAXLENGTH,edit,1350639744
Control11=IDC_STATIC,static,1342308352
Control12=IDC_EDIT_PROMPTCHAR,edit,1350631424

[CLS:CMaskEdit]
Type=0
HeaderFile=MaskEdit.h
ImplementationFile=MaskEdit.cpp
BaseClass=CEdit
Filter=W
VirtualFilter=WC
LastObject=CMaskEdit

