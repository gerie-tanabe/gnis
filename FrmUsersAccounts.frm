VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmUsersAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   Icon            =   "FrmUsersAccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab3 
      Height          =   7785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   13732
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   176
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmUsersAccounts.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Labelx(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(71)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(70)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ImageList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListViewUserAccounts"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmUsersAccounts.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label(72)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label(73)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TextBoxUserAccount"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "RaveCancelpage"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "RaveNextPage"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FrmUsersAccounts.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label(74)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "OptionButton1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "OptionButton2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Line2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Line3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "RaveCancelAccount"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "RaveCreateAccount"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "RaveBackPage"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "FrmUsersAccounts.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LabelAccountType"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "LabelUsername"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "ImgUser"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label(75)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "LabelEditAccount(0)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "LabelEditAccount(1)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "LabelEditAccount(2)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "LabelEditAccount(3)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "LabelEditAccount(4)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Image3"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Image4(0)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Image5"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Image8"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Image7"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Image1"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "LabelEditAccount(5)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "FrmUsersAccounts.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Line4"
      Tab(4).Control(1)=   "TextBoxNewName"
      Tab(4).Control(2)=   "Label(78)"
      Tab(4).Control(3)=   "Label(76)"
      Tab(4).Control(4)=   "RaveChangeCancel"
      Tab(4).Control(5)=   "RaveButtonsChangeName"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "FrmUsersAccounts.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Line1(1)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label(80)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "TextBoxConfirm"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label(79)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "TextBoxPassword"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label(77)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "TextBoxOldPassword"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label(0)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "RavePasswordCancel"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "RaveChangePassword"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "FrmUsersAccounts.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Line1(2)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label(81)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "RaveChangePicture"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "RaveButtonsChangePicture"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "ListViewPictures"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).ControlCount=   5
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "FrmUsersAccounts.frx":0D8E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Line6"
      Tab(7).Control(1)=   "Line5"
      Tab(7).Control(2)=   "OptionButtonEncode"
      Tab(7).Control(3)=   "OptionButtonAdministrator"
      Tab(7).Control(4)=   "Label(82)"
      Tab(7).Control(5)=   "RaveCancelAccountType"
      Tab(7).Control(6)=   "RaveButtonsChangeAccountType"
      Tab(7).ControlCount=   7
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "FrmUsersAccounts.frx":0DAA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Line7"
      Tab(8).Control(1)=   "Label(83)"
      Tab(8).Control(2)=   "RaveCancelDelete"
      Tab(8).Control(3)=   "RaveButtonsDeleteAccount"
      Tab(8).ControlCount=   4
      Begin Rave_Buttons.RaveButtons RaveNextPage 
         Height          =   255
         Left            =   -68880
         TabIndex        =   1
         Top             =   2190
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Next >"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":0DC6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView ListViewUserAccounts 
         Height          =   4905
         Left            =   660
         TabIndex        =   2
         Top             =   2250
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   8652
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
         SmallIcons      =   "ImageList"
         ColHdrIcons     =   "ImageList"
         ForeColor       =   4210752
         BackColor       =   16777215
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmUsersAccounts.frx":0DE2
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   8819
         EndProperty
      End
      Begin Rave_Buttons.RaveButtons RaveCancelpage 
         Height          =   255
         Left            =   -68070
         TabIndex        =   3
         Top             =   2190
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":10FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveBackPage 
         Height          =   285
         Left            =   -70440
         TabIndex        =   4
         Top             =   3480
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "< Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":1118
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveCreateAccount 
         Height          =   285
         Left            =   -69600
         TabIndex        =   5
         Top             =   3480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Create Account"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":1134
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveCancelAccount 
         Height          =   285
         Left            =   -67890
         TabIndex        =   6
         Top             =   3480
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":1150
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtonsChangeName 
         Height          =   345
         Left            =   -69570
         TabIndex        =   7
         Top             =   2400
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Change Name"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":116C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveChangeCancel 
         Height          =   345
         Left            =   -68130
         TabIndex        =   8
         Top             =   2400
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":1188
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveChangePassword 
         Height          =   345
         Left            =   -69870
         TabIndex        =   9
         Top             =   3780
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Change Password"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":11A4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RavePasswordCancel 
         Height          =   345
         Left            =   -68190
         TabIndex        =   10
         Top             =   3780
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":11C0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView ListViewPictures 
         Height          =   3165
         Left            =   -74490
         TabIndex        =   11
         Top             =   1230
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   5583
         View            =   1
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
         SmallIcons      =   "ImageList"
         ColHdrIcons     =   "ImageList"
         ForeColor       =   8421504
         BackColor       =   16777215
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmUsersAccounts.frx":11DC
         NumItems        =   0
      End
      Begin Rave_Buttons.RaveButtons RaveButtonsChangePicture 
         Height          =   285
         Left            =   -66990
         TabIndex        =   12
         Top             =   4710
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Change Picture"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":14F6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveChangePicture 
         Height          =   285
         Left            =   -65280
         TabIndex        =   13
         Top             =   4710
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":1512
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtonsChangeAccountType 
         Height          =   285
         Left            =   -69840
         TabIndex        =   14
         Top             =   3660
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Change Account Type"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":152E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveCancelAccountType 
         Height          =   285
         Left            =   -67860
         TabIndex        =   15
         Top             =   3660
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":154A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtonsDeleteAccount 
         Height          =   345
         Left            =   -69450
         TabIndex        =   16
         Top             =   1800
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Delete Account"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":1566
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveCancelDelete 
         Height          =   345
         Left            =   -68010
         TabIndex        =   17
         Top             =   1800
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUsersAccounts.frx":1582
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ImageList ImageList 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":159E
               Key             =   "snowflake"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":30F2
               Key             =   "skateboard"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":4C46
               Key             =   "sunflower"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":679A
               Key             =   "orchid"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":82EE
               Key             =   "palm"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":9E42
               Key             =   "shuttle"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":B996
               Key             =   "karate"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":D4EA
               Key             =   "horses"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":F03E
               Key             =   "guitar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":10B92
               Key             =   "frog"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":126E6
               Key             =   "fish"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":1423A
               Key             =   "duck"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":15D8E
               Key             =   "water"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":178E2
               Key             =   "dog"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":19436
               Key             =   "bike"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":1AF8A
               Key             =   "chess"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":1CADE
               Key             =   "cat"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":1E632
               Key             =   "car"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":20186
               Key             =   "butterfly"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":21CDA
               Key             =   "beach"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":2382E
               Key             =   "ball"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":25382
               Key             =   "astronaut"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmUsersAccounts.frx":26ED6
               Key             =   "jet"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type OLD password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   -74520
         TabIndex        =   50
         Top             =   960
         Width           =   1560
      End
      Begin MSForms.TextBox TextBoxOldPassword 
         Height          =   465
         Left            =   -74520
         TabIndex        =   26
         Top             =   1230
         Width           =   3615
         VariousPropertyBits=   746604571
         ForeColor       =   13461302
         BorderStyle     =   1
         Size            =   "6376;820"
         PasswordChar    =   61
         SpecialEffect   =   0
         FontName        =   "Webdings"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   2
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label LabelEditAccount 
         Height          =   195
         Index           =   5
         Left            =   -73830
         TabIndex        =   49
         Top             =   3720
         Width           =   4815
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Home"
         Size            =   "8493;344"
         MousePointer    =   99
         MouseIcon       =   "FrmUsersAccounts.frx":28A2A
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   -74490
         Picture         =   "FrmUsersAccounts.frx":28D44
         Top             =   3630
         Width           =   360
      End
      Begin VB.Image Image7 
         Height          =   360
         Left            =   -74490
         Picture         =   "FrmUsersAccounts.frx":294AE
         Top             =   3180
         Width           =   360
      End
      Begin VB.Image Image8 
         Height          =   360
         Left            =   -74490
         Picture         =   "FrmUsersAccounts.frx":29C18
         Top             =   2760
         Width           =   360
      End
      Begin VB.Image Image5 
         Height          =   360
         Left            =   -74490
         Picture         =   "FrmUsersAccounts.frx":2A382
         Top             =   2220
         Width           =   360
      End
      Begin VB.Image Image4 
         Height          =   360
         Index           =   0
         Left            =   -74490
         Picture         =   "FrmUsersAccounts.frx":2AAEC
         Top             =   1230
         Width           =   360
      End
      Begin VB.Image Image3 
         Height          =   360
         Left            =   -74490
         Picture         =   "FrmUsersAccounts.frx":2B256
         Top             =   1710
         Width           =   360
      End
      Begin MSForms.Label LabelEditAccount 
         Height          =   195
         Index           =   4
         Left            =   -73830
         TabIndex        =   48
         Top             =   3270
         Width           =   4815
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Delete the account"
         Size            =   "8493;344"
         MousePointer    =   99
         MouseIcon       =   "FrmUsersAccounts.frx":2B9C0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label LabelEditAccount 
         Height          =   195
         Index           =   3
         Left            =   -73830
         TabIndex        =   47
         Top             =   2790
         Width           =   4815
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Change the account type"
         Size            =   "8493;344"
         MousePointer    =   99
         MouseIcon       =   "FrmUsersAccounts.frx":2BCDA
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label LabelEditAccount 
         Height          =   195
         Index           =   2
         Left            =   -73830
         TabIndex        =   46
         Top             =   2310
         Width           =   4815
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Change the picture"
         Size            =   "8493;344"
         MousePointer    =   99
         MouseIcon       =   "FrmUsersAccounts.frx":2BFF4
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label LabelEditAccount 
         Height          =   195
         Index           =   1
         Left            =   -73830
         TabIndex        =   45
         Top             =   1830
         Width           =   4815
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Change the password"
         Size            =   "8493;344"
         MousePointer    =   99
         MouseIcon       =   "FrmUsersAccounts.frx":2C30E
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label LabelEditAccount 
         Height          =   195
         Index           =   0
         Left            =   -73830
         TabIndex        =   44
         Top             =   1290
         Width           =   4815
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Change the name"
         Size            =   "8493;344"
         MousePointer    =   99
         MouseIcon       =   "FrmUsersAccounts.frx":2C628
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   75
         Left            =   -74500
         TabIndex        =   43
         Top             =   500
         Width           =   660
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00F5E1D6&
         X1              =   -74460
         X2              =   -67290
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00F5E1D6&
         X1              =   -74490
         X2              =   -67320
         Y1              =   1590
         Y2              =   1590
      End
      Begin MSForms.OptionButton OptionButton2 
         Height          =   405
         Left            =   -71790
         TabIndex        =   42
         Top             =   1050
         Width           =   2745
         VariousPropertyBits=   746588179
         BackColor       =   16777215
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "4842;714"
         Value           =   "0"
         Caption         =   "Encoder"
         SpecialEffect   =   0
         GroupName       =   "access"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton OptionButton1 
         Height          =   405
         Left            =   -74520
         TabIndex        =   41
         Top             =   1050
         Width           =   2745
         VariousPropertyBits=   746588179
         BackColor       =   16777215
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "4842;714"
         MultiSelect     =   1
         Value           =   "1"
         Caption         =   "Administrator"
         SpecialEffect   =   0
         GroupName       =   "access"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick an account type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   74
         Left            =   -74500
         TabIndex        =   40
         Top             =   500
         Width           =   3030
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00F5E1D6&
         Index           =   0
         X1              =   -74490
         X2              =   -67320
         Y1              =   1920
         Y2              =   1920
      End
      Begin MSForms.TextBox TextBoxUserAccount 
         Height          =   345
         Left            =   -74490
         TabIndex        =   39
         Top             =   1380
         Width           =   3615
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         Size            =   "6376;609"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type a name for the new account:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   73
         Left            =   -74490
         TabIndex        =   38
         Top             =   1110
         Width           =   2505
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name the new account"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   72
         Left            =   -74500
         TabIndex        =   37
         Top             =   500
         Width           =   3330
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   870
         Picture         =   "FrmUsersAccounts.frx":2C942
         Top             =   990
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick a task..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   70
         Left            =   720
         TabIndex        =   36
         Top             =   495
         Width           =   1920
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "or pick an account to change"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   71
         Left            =   690
         TabIndex        =   35
         Top             =   1650
         Width           =   4215
      End
      Begin MSForms.Label Labelx 
         Height          =   390
         Index           =   2
         Left            =   1440
         TabIndex        =   34
         Top             =   1080
         Width           =   3180
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Create a new account"
         Size            =   "5609;688"
         MousePointer    =   99
         MouseIcon       =   "FrmUsersAccounts.frx":2D60C
         FontName        =   "Tahoma"
         FontEffects     =   1073741828
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Image ImgUser 
         Height          =   885
         Left            =   -68610
         Top             =   1290
         Width           =   1005
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   76
         Left            =   -74500
         TabIndex        =   33
         Top             =   500
         Width           =   660
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type a  new name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   78
         Left            =   -74520
         TabIndex        =   32
         Top             =   1320
         Width           =   1350
      End
      Begin MSForms.TextBox TextBoxNewName 
         Height          =   345
         Left            =   -74550
         TabIndex        =   31
         Top             =   1590
         Width           =   3615
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         Size            =   "6376;609"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00F5E1D6&
         X1              =   -74550
         X2              =   -67380
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   77
         Left            =   -74500
         TabIndex        =   30
         Top             =   500
         Width           =   660
      End
      Begin MSForms.TextBox TextBoxPassword 
         Height          =   465
         Left            =   -74520
         TabIndex        =   27
         Top             =   2100
         Width           =   3615
         VariousPropertyBits=   746604571
         ForeColor       =   13461302
         BorderStyle     =   1
         Size            =   "6376;820"
         PasswordChar    =   61
         SpecialEffect   =   0
         FontName        =   "Webdings"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   2
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type a NEW password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   79
         Left            =   -74490
         TabIndex        =   29
         Top             =   1830
         Width           =   1725
      End
      Begin MSForms.TextBox TextBoxConfirm 
         Height          =   435
         Left            =   -74490
         TabIndex        =   28
         Top             =   2970
         Width           =   3615
         VariousPropertyBits=   746604571
         ForeColor       =   13461302
         BorderStyle     =   1
         Size            =   "6376;767"
         PasswordChar    =   61
         SpecialEffect   =   0
         FontName        =   "Webdings"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   2
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type the new password again to confirm:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   80
         Left            =   -74490
         TabIndex        =   25
         Top             =   2700
         Width           =   3045
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00F5E1D6&
         Index           =   1
         X1              =   -74460
         X2              =   -67290
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   81
         Left            =   -74500
         TabIndex        =   24
         Top             =   500
         Width           =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00F5E1D6&
         Index           =   2
         X1              =   -74520
         X2              =   -64560
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick an account type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   82
         Left            =   -74500
         TabIndex        =   23
         Top             =   500
         Width           =   3030
      End
      Begin MSForms.OptionButton OptionButtonAdministrator 
         Height          =   405
         Left            =   -74490
         TabIndex        =   22
         Top             =   1230
         Width           =   2745
         VariousPropertyBits=   746588179
         BackColor       =   16777215
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "4842;714"
         MultiSelect     =   1
         Value           =   "1"
         Caption         =   "Administrator"
         SpecialEffect   =   0
         GroupName       =   "access"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton OptionButtonEncode 
         Height          =   405
         Left            =   -71760
         TabIndex        =   21
         Top             =   1230
         Width           =   2745
         VariousPropertyBits=   746588179
         BackColor       =   16777215
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "4842;714"
         Value           =   "0"
         Caption         =   "Encoder"
         SpecialEffect   =   0
         GroupName       =   "access"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00F5E1D6&
         X1              =   -74460
         X2              =   -67290
         Y1              =   1770
         Y2              =   1770
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00F5E1D6&
         X1              =   -74430
         X2              =   -67260
         Y1              =   3480
         Y2              =   3480
      End
      Begin MSForms.Label LabelUsername 
         Height          =   225
         Left            =   -67530
         TabIndex        =   20
         Top             =   1290
         Width           =   2175
         BackColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Username"
         Size            =   "3836;397"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label LabelAccountType 
         Height          =   375
         Left            =   -67530
         TabIndex        =   19
         Top             =   1500
         Width           =   2175
         ForeColor       =   4210752
         BackColor       =   0
         VariousPropertyBits=   8388627
         Caption         =   "Account Type"
         Size            =   "3836;661"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD6736&
         Height          =   405
         Index           =   83
         Left            =   -74500
         TabIndex        =   18
         Top             =   500
         Width           =   360
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00F5E1D6&
         X1              =   -74430
         X2              =   -67260
         Y1              =   1530
         Y2              =   1530
      End
   End
End
Attribute VB_Name = "FrmUsersAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Me.SSTab3.TabVisible(0) = False
Me.SSTab3.TabVisible(1) = False
Me.SSTab3.TabVisible(2) = False
Me.SSTab3.TabVisible(3) = False
Me.SSTab3.TabVisible(4) = False
Me.SSTab3.TabVisible(5) = False
Me.SSTab3.TabVisible(6) = False
Me.SSTab3.TabVisible(7) = False
Me.SSTab3.TabVisible(8) = False

LoadUserAccounts

End Sub




Public Sub LoadUserAccounts()
    Dim rst As New ADODB.Recordset
    rst.Open "Select * from useraccounts order by access,username", cnn, adOpenStatic
    
    Dim i  As Integer
    
    Me.ListViewUserAccounts.ListItems.Clear
    
    Dim varlist
    For i = 1 To rst.RecordCount
        
        Set varlist = Me.ListViewUserAccounts.ListItems.Add
        
        varlist.Text = "     " & rst("username")
        varlist.SubItems(1) = IIf(rst("access") = 1, "Administrator", "Encoder")
        varlist.SmallIcon = Me.ImageList.ListImages(Val(rst("picture"))).Key
        
        rst.MoveNext
    Next

End Sub

Private Sub LabelEditAccount_Click(Index As Integer)
If Index = 0 Then
        
        If Trim(Me.ListViewUserAccounts.SelectedItem.Text) = "Administrator" Then
            MsgBox "You cannot rename the Administrator account.", vbInformation, "User Accounts"
            Exit Sub
        End If
        
        If Trim(Me.ListViewUserAccounts.SelectedItem.Text) = Encoder Or AccessLevel = 1 Then
            Else
            MsgBox "You dont have administrative priviledge to change the account name.", vbInformation, "User Accounts"
            Exit Sub
        End If
        
    
        Me.SSTab3.TabVisible(4) = True
        Me.SSTab3.Tab = 4
        Me.SSTab3.TabVisible(4) = False
        Me.Label(76) = "Provide a new name for " & Me.ListViewUserAccounts.SelectedItem.Text & "'s account"
        Me.TextBoxNewName = Trim(Me.ListViewUserAccounts.SelectedItem.Text)
        Me.TextBoxNewName.SetFocus
        Me.TextBoxNewName.SelStart = 0
        Me.TextBoxNewName.SelLength = Len(Me.TextBoxNewName)
        
   End If
   
   If Index = 1 Then
       
        If Trim(Me.ListViewUserAccounts.SelectedItem.Text) = Encoder Or AccessLevel = 1 Then
            Else
            MsgBox "You dont have administrative priviledge to change the password.", vbInformation, "User Accounts"
            Exit Sub
        End If
   
        Me.SSTab3.TabVisible(5) = True
        Me.SSTab3.Tab = 5
        Me.SSTab3.TabVisible(5) = False
        Me.Label(77) = "Change " & Trim(Me.ListViewUserAccounts.SelectedItem.Text) & "'s password"
        Me.TextBoxPassword = ""
        Me.TextBoxConfirm = ""
        Me.TextBoxOldPassword.SetFocus
   End If
   
   
   If Index = 2 Then
        Me.SSTab3.TabVisible(6) = True
        Me.SSTab3.Tab = 6
        Me.SSTab3.TabVisible(6) = False
        Me.Label(81) = "Pick a new picture for " & Trim(Me.ListViewUserAccounts.SelectedItem.Text) & "'s account"
        LoadUserAccountPictures
   End If
   
   If Index = 3 Then
        
        If Trim(Me.ListViewUserAccounts.SelectedItem.Text) = "Administrator" Then
            MsgBox "You cannot change the account type of the Administrator", vbInformation, "User Accounts"
            Exit Sub
        End If
        
        If AccessLevel = 1 Then
            Else
            MsgBox "You cannot change the account type of the Administrator", vbInformation, "User Accounts"
            Exit Sub
        End If
   
        Me.SSTab3.TabVisible(7) = True
        Me.SSTab3.Tab = 7
        Me.SSTab3.TabVisible(7) = False
        If Trim(Me.LabelAccountType) = "Administrator" Then
            Me.OptionButtonAdministrator.Value = True
            Else
            Me.OptionButtonEncode.Value = True
        End If
        
        Me.Label(82) = "Pick a new account type " & Trim(Me.ListViewUserAccounts.SelectedItem.Text)
   End If
   
   If Index = 4 Then
        If Trim(Me.ListViewUserAccounts.SelectedItem.Text) = "Administrator" Then
            MsgBox "You cannot delete the Administrator account.", vbInformation, "User Accounts"
            Exit Sub
        End If
        If AccessLevel = 1 Then
            Else
            MsgBox "You cannot delete this account.", vbInformation, "User Accounts"
            Exit Sub
        End If
        
        Me.SSTab3.TabVisible(8) = True
        Me.SSTab3.Tab = 8
        Me.SSTab3.TabVisible(8) = False
        
        
        Me.Label(83) = "Are you sure you want to delete " & Trim(Me.ListViewUserAccounts.SelectedItem.Text) & "'s account?"
   End If
   
   If Index = 5 Then
        Me.SSTab3.TabVisible(0) = True
       Me.SSTab3.Tab = 0
       Me.SSTab3.TabVisible(0) = False
       LoadUserAccounts
   End If
End Sub

Private Sub LabelEditAccount_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
    For i = 1 To LabelEditAccount.UBound + 1
        If Index = i - 1 Then
            LabelEditAccount(i - 1).Font.Underline = True
            LabelEditAccount(i - 1).Font.Bold = True
        Else
            LabelEditAccount(i - 1).Font.Underline = False
            LabelEditAccount(i - 1).Font.Bold = False
        End If
   Next
End Sub

Private Sub Labelx_Click(Index As Integer)
If AccessLevel <> 1 Then
    MsgBox "Only users with administrative rights can create an account.", vbInformation, "User Accounts"
    Exit Sub
End If
        If Index = 2 Then
        Me.SSTab3.TabVisible(1) = True
        Me.SSTab3.Tab = 1
        Me.SSTab3.TabVisible(1) = False
        Me.TextBoxUserAccount = ""
        Me.TextBoxUserAccount.SetFocus
        End If
End Sub

Private Sub ListViewUserAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CurrentUserAccount = Trim(Me.ListViewUserAccounts.SelectedItem.Text)
    CurrentUserPassword = GetUserPassword(CurrentUserAccount)
    
        Me.SSTab3.TabVisible(3) = True
        Me.SSTab3.Tab = 3
        Me.SSTab3.TabVisible(3) = False

        Me.Label(75) = "What do you want to change about " & CurrentUserAccount & "'s account? "
        Me.LabelUsername = CurrentUserAccount
        Me.LabelAccountType = Trim(Me.ListViewUserAccounts.SelectedItem.SubItems(1))
        Me.ImgUser.Picture = Me.ImageList.ListImages(Me.ListViewUserAccounts.SelectedItem.SmallIcon).Picture
End Sub


Private Function GetUserPassword(username As String) As String
        Dim rst  As New ADODB.Recordset
        rst.Open "Select password from useraccounts where username='" & username & "'", cnn, adOpenStatic
        If rst.RecordCount > 0 Then
            GetUserPassword = rst!Password
        End If
        
End Function


Private Sub RaveBackPage_Click()
 Me.SSTab3.TabVisible(1) = True
        Me.SSTab3.Tab = 1
        Me.SSTab3.TabVisible(1) = False
End Sub

Private Sub RaveButtonsChangeAccountType_Click()
cnn.Execute "update useraccounts set access=" & IIf(Me.OptionButtonAdministrator.Value = True, 1, 2) & " where username='" & Replace(CurrentUserAccount, "'", "''") & "'"
    Me.SSTab3.TabVisible(3) = True
    Me.SSTab3.Tab = 3
    Me.SSTab3.TabVisible(3) = False
    Me.LabelAccountType = IIf(Me.OptionButtonAdministrator.Value = True, "Administrator", "Encoder")
End Sub

Private Sub RaveButtonsChangeName_Click()
If IfDuplicateUser(Trim(Me.TextBoxNewName)) Then
   MsgBox "An account name '" & Trim(Me.TextBoxNewName) & "' already exist. Type a different name.", vbInformation, "User Accounts"
   Exit Sub
End If
    cnn.Execute "update useraccounts set username='" & Trim(Replace(Me.TextBoxNewName, "'", "''")) & "' where username='" & CurrentUserAccount & "'"
    Me.SSTab3.TabVisible(3) = True
    Me.SSTab3.Tab = 3
    Me.SSTab3.TabVisible(3) = False
    
    CurrentUserAccount = Trim(Me.TextBoxNewName)
    Me.Label(75) = "What to you want to change about " & CurrentUserAccount & "'s account? "
    Me.LabelUsername = CurrentUserAccount
End Sub

Private Sub RaveButtonsChangePicture_Click()
cnn.Execute "update useraccounts set picture=" & Me.ListViewPictures.SelectedItem.Index & " where username='" & Replace(CurrentUserAccount, "'", "''") & "'"
    Me.SSTab3.TabVisible(3) = True
    Me.SSTab3.Tab = 3
    Me.SSTab3.TabVisible(3) = False
    
    Me.ImgUser.Picture = Me.ImageList.ListImages(Me.ListViewPictures.SelectedItem.SmallIcon).Picture
End Sub

Private Sub RaveButtonsDeleteAccount_Click()
cnn.Execute "delete from useraccounts where username='" & Replace(CurrentUserAccount, "'", "''") & "'"
    LoadUserAccounts
    Me.SSTab3.TabVisible(0) = True
    Me.SSTab3.Tab = 0
    Me.SSTab3.TabVisible(0) = False
End Sub

Private Sub RaveCancelAccount_Click()
Me.SSTab3.TabVisible(0) = True
        Me.SSTab3.Tab = 0
        Me.SSTab3.TabVisible(0) = False
End Sub

Private Sub RaveCancelAccountType_Click()
Me.SSTab3.TabVisible(3) = True
        Me.SSTab3.Tab = 3
        Me.SSTab3.TabVisible(3) = False
End Sub

Private Sub RaveCancelDelete_Click()
Me.SSTab3.TabVisible(3) = True
        Me.SSTab3.Tab = 3
        Me.SSTab3.TabVisible(3) = False
End Sub

Private Sub RaveCancelpage_Click()
    Me.SSTab3.TabVisible(0) = True
        Me.SSTab3.Tab = 0
        Me.SSTab3.TabVisible(0) = False
End Sub

Private Sub RaveChangeCancel_Click()
        Me.SSTab3.TabVisible(3) = True
        Me.SSTab3.Tab = 3
        Me.SSTab3.TabVisible(3) = False
End Sub

Private Sub RaveChangePassword_Click()

If Trim(Me.TextBoxOldPassword) = CurrentUserPassword Then




                    If Trim(Me.TextBoxPassword) = Trim(Me.TextBoxConfirm) Then
                        
                       cnn.Execute "update useraccounts set password='" & Trim(Replace(Me.TextBoxPassword, "'", "''")) & "' where username='" & Replace(CurrentUserAccount, "'", "''") & "'"
                        Me.SSTab3.TabVisible(3) = True
                        Me.SSTab3.Tab = 3
                        Me.SSTab3.TabVisible(3) = False
                        
                        Else
                        
                        MsgBox "The password you typed do not match. Please type the new password in both boxes.", vbExclamation, "User Account"
                    End If
                    
Else

                    MsgBox "The OLD password you typed is not correct.", vbExclamation, "User Account"

End If

End Sub

Private Sub RaveChangePicture_Click()
Me.SSTab3.TabVisible(3) = True
        Me.SSTab3.Tab = 3
        Me.SSTab3.TabVisible(3) = False
End Sub

Private Sub RaveCreateAccount_Click()
Dim R As Integer
  Randomize
  R = Int(Rnd(1) * 23 + 1)
 
  cnn.Execute "insert into useraccounts (username,password,access,picture) values('" & Trim(StrConv(Me.TextBoxUserAccount, vbProperCase)) & "'," & "''" & "," & IIf(Me.OptionButton1, 1, 2) & "," & R & ")"
  LoadUserAccounts
    
  Me.SSTab3.TabVisible(0) = True
  Me.SSTab3.Tab = 0
  Me.SSTab3.TabVisible(0) = False
End Sub

Private Sub RaveNextPage_Click()
If Trim(Me.TextBoxUserAccount) <> "" Then
        Me.SSTab3.TabVisible(2) = True
        Me.SSTab3.Tab = 2
        Me.SSTab3.TabVisible(2) = False
       
End If
End Sub

Public Sub LoadUserAccountPictures()
Dim i As Integer
Me.ListViewPictures.ListItems.Clear
    For i = 1 To Me.ImageList.ListImages.Count
        Me.ListViewPictures.ListItems.Add i, , , Me.ImageList.ListImages(i).Key, Me.ImageList.ListImages(i).Key
    Next
End Sub


Public Function IfDuplicateUser(username As String) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "Select username from useraccounts where username='" & Replace(username, "'", "''") & "'", cnn, adOpenStatic
    If rst.RecordCount > 0 Then
        IfDuplicateUser = True
        Else
        IfDuplicateUser = False
        End If
End Function

Private Sub RavePasswordCancel_Click()
Me.SSTab3.TabVisible(3) = True
        Me.SSTab3.Tab = 3
        Me.SSTab3.TabVisible(3) = False
End Sub

