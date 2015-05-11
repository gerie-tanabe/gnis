VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0460CA20-346F-11CF-8682-00805F7CED21}#1.1#0"; "mo10.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Object = "{DDA53BD0-2CD0-11D4-8ED4-00E07D815373}#1.0#0"; "mbmouse.ocx"
Begin VB.Form FrmGCPDS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "gnis"
   ClientHeight    =   10680
   ClientLeft      =   1260
   ClientTop       =   3105
   ClientWidth     =   14700
   ForeColor       =   &H00C00000&
   Icon            =   "FrmGCPDS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmGCPDS.frx":0CCA
   ScaleHeight     =   712
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   980
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveTab 
      Height          =   405
      Index           =   5
      Left            =   8280
      TabIndex        =   273
      Top             =   1125
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "Triangulation"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8260132
      BCOLO           =   8260132
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmGCPDS.frx":132B2
      PICN            =   "FrmGCPDS.frx":132CE
      PICH            =   "FrmGCPDS.frx":15B94
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   13800
      Top             =   120
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9165
      Left            =   0
      TabIndex        =   85
      Top             =   1560
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   16166
      _Version        =   393216
      Tabs            =   9
      Tab             =   5
      TabsPerRow      =   9
      TabHeight       =   132
      ShowFocusRect   =   0   'False
      BackColor       =   14737632
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmGCPDS.frx":18C67
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label(36)"
      Tab(0).Control(1)=   "Label(4)"
      Tab(0).Control(2)=   "Label(8)"
      Tab(0).Control(3)=   "TxtBarangay"
      Tab(0).Control(4)=   "TxtMunicipality"
      Tab(0).Control(5)=   "TxtProvince"
      Tab(0).Control(6)=   "TxtZone"
      Tab(0).Control(7)=   "TxtNorthing"
      Tab(0).Control(8)=   "TxtEasting"
      Tab(0).Control(9)=   "Label(7)"
      Tab(0).Control(10)=   "TxtName"
      Tab(0).Control(11)=   "TxtDateComputed"
      Tab(0).Control(12)=   "txtLastRecover"
      Tab(0).Control(13)=   "Shape2"
      Tab(0).Control(14)=   "TxtEllipsoidalH2"
      Tab(0).Control(15)=   "Label(12)"
      Tab(0).Control(16)=   "Label(11)"
      Tab(0).Control(17)=   "Label(1)"
      Tab(0).Control(18)=   "txtMarkStatus"
      Tab(0).Control(19)=   "txtMarkType"
      Tab(0).Control(20)=   "txtMarkPurpose"
      Tab(0).Control(21)=   "Label(30)"
      Tab(0).Control(22)=   "Label(29)"
      Tab(0).Control(23)=   "Label(10)"
      Tab(0).Control(24)=   "Label(9)"
      Tab(0).Control(25)=   "Label(31)"
      Tab(0).Control(26)=   "Label(32)"
      Tab(0).Control(27)=   "TxtEstablishedBy"
      Tab(0).Control(28)=   "TxtFixingMethod"
      Tab(0).Control(29)=   "Label(18)"
      Tab(0).Control(30)=   "Label(17)"
      Tab(0).Control(31)=   "Label(15)"
      Tab(0).Control(32)=   "Label(0)"
      Tab(0).Control(33)=   "Label(3)"
      Tab(0).Control(34)=   "TextBox1"
      Tab(0).Control(35)=   "Label(19)"
      Tab(0).Control(36)=   "TxtDateEntry"
      Tab(0).Control(37)=   "TxtDLat"
      Tab(0).Control(38)=   "TxtMLat"
      Tab(0).Control(39)=   "TxtSLat"
      Tab(0).Control(40)=   "TxtSLong"
      Tab(0).Control(41)=   "TxtMLong"
      Tab(0).Control(42)=   "TxtDLong"
      Tab(0).Control(43)=   "Label(21)"
      Tab(0).Control(44)=   "Label(16)"
      Tab(0).Control(45)=   "Label(14)"
      Tab(0).Control(46)=   "Label(28)"
      Tab(0).Control(47)=   "Label(33)"
      Tab(0).Control(48)=   "Label(34)"
      Tab(0).Control(49)=   "Label(35)"
      Tab(0).Control(50)=   "Label(37)"
      Tab(0).Control(51)=   "Label(38)"
      Tab(0).Control(52)=   "txtAuthority"
      Tab(0).Control(53)=   "FormCaption"
      Tab(0).Control(54)=   "Label(43)"
      Tab(0).Control(55)=   "TxtRegion"
      Tab(0).Control(56)=   "Label(13)"
      Tab(0).Control(57)=   "ComboBoxMonth"
      Tab(0).Control(58)=   "txtEstablished"
      Tab(0).Control(59)=   "ComboBoxDay"
      Tab(0).Control(60)=   "txtYear"
      Tab(0).Control(61)=   "Label(86)"
      Tab(0).Control(62)=   "LabelEncoder"
      Tab(0).Control(63)=   "LabelDateUpdated"
      Tab(0).Control(64)=   "Label1"
      Tab(0).Control(65)=   "Label2"
      Tab(0).Control(66)=   "Label5"
      Tab(0).Control(67)=   "Label(2)"
      Tab(0).Control(68)=   "Label(84)"
      Tab(0).Control(69)=   "Label(85)"
      Tab(0).Control(70)=   "Shape4"
      Tab(0).Control(71)=   "TextBoxUTMEasting"
      Tab(0).Control(72)=   "TextBoxUTMNorthing"
      Tab(0).Control(73)=   "TextBoxUTMZone"
      Tab(0).Control(74)=   "TextBox2"
      Tab(0).Control(75)=   "Label(42)"
      Tab(0).Control(76)=   "Label(27)"
      Tab(0).Control(77)=   "Label(25)"
      Tab(0).Control(78)=   "Label(24)"
      Tab(0).Control(79)=   "Label(23)"
      Tab(0).Control(80)=   "Label(22)"
      Tab(0).Control(81)=   "Label(20)"
      Tab(0).Control(82)=   "TxtLongD"
      Tab(0).Control(83)=   "TxtLongM"
      Tab(0).Control(84)=   "TxtLongS"
      Tab(0).Control(85)=   "TxtLatS"
      Tab(0).Control(86)=   "TxtLatM"
      Tab(0).Control(87)=   "TxtLatD"
      Tab(0).Control(88)=   "Label(26)"
      Tab(0).Control(89)=   "Label(6)"
      Tab(0).Control(90)=   "Label(5)"
      Tab(0).Control(91)=   "txtEllipsoidalH"
      Tab(0).Control(92)=   "Shape3"
      Tab(0).Control(93)=   "Shape1"
      Tab(0).Control(94)=   "TxtRef"
      Tab(0).Control(95)=   "TextBox5"
      Tab(0).Control(96)=   "TxtIsland"
      Tab(0).Control(97)=   "txtOrder"
      Tab(0).Control(98)=   "TxtMSL"
      Tab(0).Control(99)=   "RaveTranform2UTM"
      Tab(0).Control(100)=   "RaveTranformToPTM"
      Tab(0).Control(101)=   "RaveTranform2WGS84"
      Tab(0).Control(102)=   "RaveTranform2PRS92"
      Tab(0).Control(103)=   "cmdLocation"
      Tab(0).Control(104)=   "RaveImages"
      Tab(0).Control(105)=   "RaveInfo"
      Tab(0).Control(106)=   "RaveSearch"
      Tab(0).Control(107)=   "RaveGIS"
      Tab(0).Control(108)=   "RaveNext"
      Tab(0).Control(109)=   "RaveBack"
      Tab(0).Control(110)=   "RavePrint"
      Tab(0).Control(111)=   "RaveCancel"
      Tab(0).Control(112)=   "RaveSave"
      Tab(0).Control(113)=   "RaveDelete"
      Tab(0).Control(114)=   "RaveEdit"
      Tab(0).Control(115)=   "RaveAdd"
      Tab(0).ControlCount=   116
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmGCPDS.frx":18C83
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ResultLabel"
      Tab(1).Control(1)=   "PageCounterLabel"
      Tab(1).Control(2)=   "optGCPs"
      Tab(1).Control(3)=   "optBMs"
      Tab(1).Control(4)=   "RaveRemoveQuery"
      Tab(1).Control(5)=   "RaveEditQuery"
      Tab(1).Control(6)=   "RaveNextPage"
      Tab(1).Control(7)=   "RavePreviousPage"
      Tab(1).Control(8)=   "RaveDeleteQuery"
      Tab(1).Control(9)=   "RaveAddQuery"
      Tab(1).Control(10)=   "LstConditions"
      Tab(1).Control(11)=   "RaveFilter"
      Tab(1).Control(12)=   "LstResult"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FrmGCPDS.frx":18C9F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label(41)"
      Tab(2).Control(1)=   "RaveGoto"
      Tab(2).Control(2)=   "RaveExtent"
      Tab(2).Control(3)=   "RaveZoomOut"
      Tab(2).Control(4)=   "RaveZoomIn"
      Tab(2).Control(5)=   "StatusBar1"
      Tab(2).Control(6)=   "MyMap"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "FrmGCPDS.frx":18CBB
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "RaveButtons7"
      Tab(3).Control(1)=   "RaveDeleteDuplicateBMs"
      Tab(3).Control(2)=   "RaveButtons15"
      Tab(3).Control(3)=   "RaveDMS2Degree"
      Tab(3).Control(4)=   "RaveButtonsUploaded"
      Tab(3).Control(5)=   "RaveButtonsImport"
      Tab(3).Control(6)=   "RaveButtonsExport"
      Tab(3).Control(7)=   "RaveInventory"
      Tab(3).Control(8)=   "SummaryRave"
      Tab(3).Control(9)=   "RaveDeletedGCPs"
      Tab(3).Control(10)=   "RaveUserAccounts"
      Tab(3).Control(11)=   "RaveMarkStatus"
      Tab(3).Control(12)=   "RaveMarkPurpose"
      Tab(3).Control(13)=   "RaveMarkType"
      Tab(3).Control(14)=   "RequestingPartyLibRave"
      Tab(3).Control(15)=   "SignatoryLibRave"
      Tab(3).Control(16)=   "RaveButtons16"
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "FrmGCPDS.frx":18CD7
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "CommonDialogImport"
      Tab(4).Control(1)=   "RaveButtons2"
      Tab(4).Control(2)=   "RaveButtons1"
      Tab(4).Control(3)=   "ImportExcel"
      Tab(4).Control(4)=   "ExtractWGS84Rave"
      Tab(4).Control(5)=   "DescriptionCommonDialog"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "FrmGCPDS.frx":18CF3
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "TxtBMDateOfEntry"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label(39)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label(44)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label(45)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label(46)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label(47)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label(48)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label(49)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Label(50)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Label(51)"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Label(52)"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "Label(53)"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "TxtBMDateLastRecovered"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "TxtBMDateComputed"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "TxtEName"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Label(54)"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "TxtEMunicipality"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "TxtEBarangay"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "Label(55)"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "Label(56)"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "Label(57)"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "TxtElevation"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "Label(58)"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).Control(23)=   "Label(59)"
      Tab(5).Control(23).Enabled=   0   'False
      Tab(5).Control(24)=   "TxtBMDescription"
      Tab(5).Control(24).Enabled=   0   'False
      Tab(5).Control(25)=   "TxtEFix"
      Tab(5).Control(25).Enabled=   0   'False
      Tab(5).Control(26)=   "Label(60)"
      Tab(5).Control(26).Enabled=   0   'False
      Tab(5).Control(27)=   "TxtEProvince"
      Tab(5).Control(27).Enabled=   0   'False
      Tab(5).Control(28)=   "BMMarkStatus"
      Tab(5).Control(28).Enabled=   0   'False
      Tab(5).Control(29)=   "BMMarkPurpose"
      Tab(5).Control(29).Enabled=   0   'False
      Tab(5).Control(30)=   "BMMarkType"
      Tab(5).Control(30).Enabled=   0   'False
      Tab(5).Control(31)=   "TxtEDatum"
      Tab(5).Control(31).Enabled=   0   'False
      Tab(5).Control(32)=   "TxtBMOrder"
      Tab(5).Control(32).Enabled=   0   'False
      Tab(5).Control(33)=   "TxtBMDateEstablished"
      Tab(5).Control(33).Enabled=   0   'False
      Tab(5).Control(34)=   "TxtElevationAuthority"
      Tab(5).Control(34).Enabled=   0   'False
      Tab(5).Control(35)=   "TxtBMAuthority"
      Tab(5).Control(35).Enabled=   0   'False
      Tab(5).Control(36)=   "FormCaption2"
      Tab(5).Control(36).Enabled=   0   'False
      Tab(5).Control(37)=   "TxtERegion"
      Tab(5).Control(37).Enabled=   0   'False
      Tab(5).Control(38)=   "Label(61)"
      Tab(5).Control(38).Enabled=   0   'False
      Tab(5).Control(39)=   "Label(40)"
      Tab(5).Control(39).Enabled=   0   'False
      Tab(5).Control(40)=   "TextBoxLongitude"
      Tab(5).Control(40).Enabled=   0   'False
      Tab(5).Control(41)=   "TextBoxLatitude"
      Tab(5).Control(41).Enabled=   0   'False
      Tab(5).Control(42)=   "Label(62)"
      Tab(5).Control(42).Enabled=   0   'False
      Tab(5).Control(43)=   "Label(63)"
      Tab(5).Control(43).Enabled=   0   'False
      Tab(5).Control(44)=   "Label7"
      Tab(5).Control(44).Enabled=   0   'False
      Tab(5).Control(45)=   "Label8"
      Tab(5).Control(45).Enabled=   0   'False
      Tab(5).Control(46)=   "LabelDateUpdated2"
      Tab(5).Control(46).Enabled=   0   'False
      Tab(5).Control(47)=   "LabelEncoder2"
      Tab(5).Control(47).Enabled=   0   'False
      Tab(5).Control(48)=   "TxtEIsland"
      Tab(5).Control(48).Enabled=   0   'False
      Tab(5).Control(49)=   "Label(88)"
      Tab(5).Control(49).Enabled=   0   'False
      Tab(5).Control(50)=   "TxtBMPlus"
      Tab(5).Control(50).Enabled=   0   'False
      Tab(5).Control(51)=   "RaveDup"
      Tab(5).Control(51).Enabled=   0   'False
      Tab(5).Control(52)=   "RaveEditBM"
      Tab(5).Control(52).Enabled=   0   'False
      Tab(5).Control(53)=   "RaveMapBM"
      Tab(5).Control(53).Enabled=   0   'False
      Tab(5).Control(54)=   "RaveButtons17"
      Tab(5).Control(54).Enabled=   0   'False
      Tab(5).Control(55)=   "RaveSearchBM"
      Tab(5).Control(55).Enabled=   0   'False
      Tab(5).Control(56)=   "RaveNextBM"
      Tab(5).Control(56).Enabled=   0   'False
      Tab(5).Control(57)=   "RaveBackBM"
      Tab(5).Control(57).Enabled=   0   'False
      Tab(5).Control(58)=   "RavePrintBM"
      Tab(5).Control(58).Enabled=   0   'False
      Tab(5).Control(59)=   "RaveCancelBM"
      Tab(5).Control(59).Enabled=   0   'False
      Tab(5).Control(60)=   "RaveSaveBM"
      Tab(5).Control(60).Enabled=   0   'False
      Tab(5).Control(61)=   "RaveDeleteBM"
      Tab(5).Control(61).Enabled=   0   'False
      Tab(5).Control(62)=   "RaveAddBM"
      Tab(5).Control(62).Enabled=   0   'False
      Tab(5).ControlCount=   63
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "FrmGCPDS.frx":18D0F
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "RaveBM"
      Tab(6).Control(1)=   "StatusBar2"
      Tab(6).Control(2)=   "RaveExtentBM"
      Tab(6).Control(3)=   "RavePanBM"
      Tab(6).Control(4)=   "RaveZoomOutBM"
      Tab(6).Control(5)=   "RaveZoomInBM"
      Tab(6).Control(6)=   "MyMap2"
      Tab(6).ControlCount=   7
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "FrmGCPDS.frx":18D2B
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "TextBoxGravityDescription"
      Tab(7).Control(1)=   "Label(94)"
      Tab(7).Control(2)=   "TextBoxGravityName"
      Tab(7).Control(3)=   "TextBoxGravityRegion"
      Tab(7).Control(4)=   "TextBoxGravityProvince"
      Tab(7).Control(5)=   "TextBoxGravityMunicipality"
      Tab(7).Control(6)=   "TextBoxGravityBarangay"
      Tab(7).Control(7)=   "TextBoxGravityLatitude"
      Tab(7).Control(8)=   "TextBoxGravityLongitude"
      Tab(7).Control(9)=   "TextBoxObservedValues"
      Tab(7).Control(10)=   "TextBoxGravityElevation"
      Tab(7).Control(11)=   "Label(64)"
      Tab(7).Control(12)=   "Label(75)"
      Tab(7).Control(13)=   "Label(72)"
      Tab(7).Control(14)=   "Label(71)"
      Tab(7).Control(15)=   "Label(74)"
      Tab(7).Control(16)=   "Label(73)"
      Tab(7).Control(17)=   "Label(93)"
      Tab(7).Control(18)=   "Label(87)"
      Tab(7).Control(19)=   "Label(82)"
      Tab(7).Control(20)=   "Label(68)"
      Tab(7).Control(21)=   "ComboBoxOrderGravity"
      Tab(7).Control(22)=   "Label3"
      Tab(7).Control(23)=   "Label4"
      Tab(7).Control(24)=   "LabelUpdatedGravity"
      Tab(7).Control(25)=   "LabelencoderGravity"
      Tab(7).Control(26)=   "Label6"
      Tab(7).Control(27)=   "LabelGravityRecordStatus"
      Tab(7).Control(28)=   "RaveButtonsUnits"
      Tab(7).Control(29)=   "RaveSearchGravity"
      Tab(7).Control(30)=   "RaveLocation"
      Tab(7).Control(31)=   "RaveNextGravity"
      Tab(7).Control(32)=   "RaveBackGravity"
      Tab(7).Control(33)=   "RavePrintGravity"
      Tab(7).Control(34)=   "RaveCancelGravity"
      Tab(7).Control(35)=   "RaveSaveGravity"
      Tab(7).Control(36)=   "RaveDeleteGravity"
      Tab(7).Control(37)=   "RaveEditGravity"
      Tab(7).Control(38)=   "RaveAddGravity"
      Tab(7).ControlCount=   39
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "FrmGCPDS.frx":18D47
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "RaveButtons3"
      Tab(8).Control(1)=   "RaveButtons4"
      Tab(8).Control(2)=   "RaveButtons5"
      Tab(8).Control(3)=   "RaveButtons6"
      Tab(8).Control(4)=   "TextBoxOldBookmark"
      Tab(8).Control(5)=   "Label(79)"
      Tab(8).Control(6)=   "TextBoxOldOrder"
      Tab(8).Control(7)=   "TextBoxOldLongitude"
      Tab(8).Control(8)=   "TextBoxOldLatitude"
      Tab(8).Control(9)=   "TextBoxOldDescription"
      Tab(8).Control(10)=   "Label(83)"
      Tab(8).Control(11)=   "Label(81)"
      Tab(8).Control(12)=   "Label(80)"
      Tab(8).Control(13)=   "Label(78)"
      Tab(8).Control(14)=   "TextBoxOldDateEntry"
      Tab(8).Control(15)=   "Label(77)"
      Tab(8).Control(16)=   "TextBoxOldDateEstablished"
      Tab(8).Control(17)=   "Label(76)"
      Tab(8).Control(18)=   "Label(70)"
      Tab(8).Control(19)=   "Label(69)"
      Tab(8).Control(20)=   "TextBoxOldBarangay"
      Tab(8).Control(21)=   "TextBoxOldMunicipality"
      Tab(8).Control(22)=   "TextBoxOldProvince"
      Tab(8).Control(23)=   "Label(67)"
      Tab(8).Control(24)=   "TextBoxOldStationName"
      Tab(8).Control(25)=   "Label(66)"
      Tab(8).Control(26)=   "TextBoxOldRegion"
      Tab(8).Control(27)=   "Label(65)"
      Tab(8).ControlCount=   28
      Begin MapObjects.Map MyMap 
         Height          =   7665
         Left            =   -74775
         TabIndex        =   192
         Top             =   850
         Width           =   14025
         _Version        =   65537
         _ExtentX        =   24739
         _ExtentY        =   13520
         _StockProps     =   225
         BackColor       =   8207644
         BorderStyle     =   1
         Appearance      =   1
         ScrollBars      =   0   'False
         BackColor       =   8207644
         FullRedrawOnPan =   -1  'True
         Contents        =   "FrmGCPDS.frx":18D63
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   465
         Left            =   -74775
         TabIndex        =   184
         Top             =   8520
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   820
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   5
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Projection: WGS84  "
               TextSave        =   "Projection: WGS84  "
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   2646
               MinWidth        =   2646
               Text            =   "Tool: None"
               TextSave        =   "Tool: None"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   7056
               MinWidth        =   7056
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   7056
               MinWidth        =   7056
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog DescriptionCommonDialog 
         Left            =   -65400
         Top             =   4095
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         MaxFileSize     =   10000
      End
      Begin Rave_Buttons.RaveButtons RaveAdd 
         Height          =   885
         Left            =   -73830
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   145
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Add"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":18D7D
         PICN            =   "FrmGCPDS.frx":18D99
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveEdit 
         Height          =   885
         Left            =   -72750
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":19292
         PICN            =   "FrmGCPDS.frx":192AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveDelete 
         Height          =   885
         Left            =   -71670
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   125
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":197AB
         PICN            =   "FrmGCPDS.frx":197C7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveSave 
         Height          =   885
         Left            =   -70590
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":19C94
         PICN            =   "FrmGCPDS.frx":19CB0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveCancel 
         Height          =   885
         Left            =   -69510
         TabIndex        =   89
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":19EB3
         PICN            =   "FrmGCPDS.frx":19ECF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RavePrint 
         Height          =   885
         Left            =   -68430
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1A3E3
         PICN            =   "FrmGCPDS.frx":1A3FF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveBack 
         Height          =   885
         Left            =   -67350
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1A6EC
         PICN            =   "FrmGCPDS.frx":1A708
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveNext 
         Height          =   885
         Left            =   -66270
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Next"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1A9E6
         PICN            =   "FrmGCPDS.frx":1AA02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveGIS 
         Height          =   885
         Left            =   -65190
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Map"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1ACD8
         PICN            =   "FrmGCPDS.frx":1ACF4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveSearch 
         Height          =   885
         Left            =   -74910
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1B220
         PICN            =   "FrmGCPDS.frx":1B23C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveInfo 
         Height          =   885
         Left            =   -64110
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Info"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1B513
         PICN            =   "FrmGCPDS.frx":1B52F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveImages 
         Height          =   885
         Left            =   -63030
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Images"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1BA61
         PICN            =   "FrmGCPDS.frx":1BA7D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons cmdLocation 
         Height          =   315
         Left            =   -64770
         TabIndex        =   2
         Top             =   1360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   11110782
         BCOLO           =   11110782
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1BD91
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView LstResult 
         Height          =   3435
         Left            =   -74850
         TabIndex        =   117
         Top             =   3620
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   6059
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   8207644
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Station Name"
            Object.Width           =   8784
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin Rave_Buttons.RaveButtons RaveFilter 
         Height          =   555
         Left            =   -68880
         TabIndex        =   118
         Top             =   8290
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   979
         BTYPE           =   4
         TX              =   "Go to Record"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   32768
         BCOLO           =   32768
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1BDAD
         PICN            =   "FrmGCPDS.frx":1BDC9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveTranform2PRS92 
         Height          =   465
         Left            =   -73800
         TabIndex        =   121
         Top             =   6120
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "to PRS92"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   11292960
         BCOLO           =   11292960
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1CE5B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView LstConditions 
         Height          =   2250
         Left            =   -74880
         TabIndex        =   133
         Top             =   500
         Width           =   14115
         _ExtentX        =   24897
         _ExtentY        =   3969
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   8207644
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Field"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Operator"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Value"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Boolean"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Alias"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
      End
      Begin Rave_Buttons.RaveButtons RaveAddQuery 
         Height          =   465
         Left            =   -70200
         TabIndex        =   134
         Top             =   2945
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "Add"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1CE77
         PICN            =   "FrmGCPDS.frx":1CE93
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveDeleteQuery 
         Height          =   465
         Left            =   -64800
         TabIndex        =   135
         Top             =   2945
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "Reset"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1D017
         PICN            =   "FrmGCPDS.frx":1D033
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RavePreviousPage 
         Height          =   405
         Left            =   -68280
         TabIndex        =   136
         Top             =   7720
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1D2F0
         PICN            =   "FrmGCPDS.frx":1D30C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveNextPage 
         Height          =   405
         Left            =   -67710
         TabIndex        =   137
         Top             =   7720
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   714
         BTYPE           =   9
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1D5EA
         PICN            =   "FrmGCPDS.frx":1D606
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveZoomIn 
         Height          =   450
         Left            =   -74730
         TabIndex        =   140
         ToolTipText     =   "Zoom Box"
         Top             =   325
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1D8DC
         PICN            =   "FrmGCPDS.frx":1D8F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveZoomOut 
         Height          =   450
         Left            =   -74235
         TabIndex        =   141
         ToolTipText     =   "Zoom Out"
         Top             =   330
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1DF22
         PICN            =   "FrmGCPDS.frx":1DF3E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveAddBM 
         Height          =   885
         Left            =   1170
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Add"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1E568
         PICN            =   "FrmGCPDS.frx":1E584
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveDeleteBM 
         Height          =   885
         Left            =   3330
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1EA7D
         PICN            =   "FrmGCPDS.frx":1EA99
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveSaveBM 
         Height          =   885
         Left            =   4410
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1EF66
         PICN            =   "FrmGCPDS.frx":1EF82
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveCancelBM 
         Height          =   885
         Left            =   5490
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1F185
         PICN            =   "FrmGCPDS.frx":1F1A1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RavePrintBM 
         Height          =   885
         Left            =   8730
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   140
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1F6B5
         PICN            =   "FrmGCPDS.frx":1F6D1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveBackBM 
         Height          =   885
         Left            =   6570
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1F9BE
         PICN            =   "FrmGCPDS.frx":1F9DA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveNextBM 
         Height          =   885
         Left            =   7650
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Next"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1FCB8
         PICN            =   "FrmGCPDS.frx":1FCD4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveSearchBM 
         Height          =   885
         Left            =   120
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":1FFAA
         PICN            =   "FrmGCPDS.frx":1FFC6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons17 
         Height          =   315
         Left            =   4800
         TabIndex        =   39
         Top             =   2160
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   11110782
         BCOLO           =   11110782
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2029D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons16 
         Height          =   540
         Left            =   -67710
         TabIndex        =   175
         Top             =   1900
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Location Library            "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":202B9
         PICN            =   "FrmGCPDS.frx":202D5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons SignatoryLibRave 
         Height          =   540
         Left            =   -67725
         TabIndex        =   176
         Top             =   2530
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Signatory                     "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2083C
         PICN            =   "FrmGCPDS.frx":20858
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RequestingPartyLibRave 
         Height          =   540
         Left            =   -67710
         TabIndex        =   177
         Top             =   3160
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Requesting Party         "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":20D55
         PICN            =   "FrmGCPDS.frx":20D71
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveMarkType 
         Height          =   540
         Left            =   -67710
         TabIndex        =   178
         Top             =   3790
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Mark Type                     "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":21E03
         PICN            =   "FrmGCPDS.frx":21E1F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveMarkPurpose 
         Height          =   540
         Left            =   -67710
         TabIndex        =   179
         Top             =   4420
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Mark Purpose              "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":22345
         PICN            =   "FrmGCPDS.frx":22361
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveMarkStatus 
         Height          =   540
         Left            =   -67710
         TabIndex        =   180
         Top             =   5095
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Mark Status                 "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2303B
         PICN            =   "FrmGCPDS.frx":23057
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons ExtractWGS84Rave 
         Height          =   555
         Left            =   -65685
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   8245
         Visible         =   0   'False
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   979
         BTYPE           =   8
         TX              =   "Extract WGS84 from the description"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":23556
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons ImportExcel 
         Height          =   420
         Left            =   -65685
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   7750
         Visible         =   0   'False
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   741
         BTYPE           =   8
         TX              =   "Import From MS Excel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":23572
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveExtent 
         Height          =   450
         Left            =   -73740
         TabIndex        =   183
         ToolTipText     =   "Full Extent"
         Top             =   325
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2358E
         PICN            =   "FrmGCPDS.frx":235AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons1 
         Height          =   450
         Left            =   -65685
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   6670
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   794
         BTYPE           =   8
         TX              =   "Detailed Summary"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":23BD4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons2 
         Height          =   450
         Left            =   -65685
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   7210
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   794
         BTYPE           =   8
         TX              =   "Detect Region"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":23BF0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveEditQuery 
         Height          =   465
         Left            =   -68400
         TabIndex        =   189
         Top             =   2945
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "Edit  "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":23C0C
         PICN            =   "FrmGCPDS.frx":23C28
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveGoto 
         Height          =   450
         Left            =   -73245
         TabIndex        =   139
         ToolTipText     =   "Goto"
         Top             =   325
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":24125
         PICN            =   "FrmGCPDS.frx":24141
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComDlg.CommonDialog CommonDialogImport 
         Left            =   -65400
         Top             =   3615
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         MaxFileSize     =   10000
      End
      Begin Rave_Buttons.RaveButtons RaveMapBM 
         Height          =   885
         Left            =   9720
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Map"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":245B5
         PICN            =   "FrmGCPDS.frx":245D1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MapObjects.Map MyMap2 
         Height          =   7665
         Left            =   -74820
         TabIndex        =   197
         Top             =   700
         Width           =   14025
         _Version        =   65537
         _ExtentX        =   24739
         _ExtentY        =   13520
         _StockProps     =   225
         BackColor       =   8207644
         BorderStyle     =   1
         Appearance      =   1
         ScrollBars      =   0   'False
         BackColor       =   8207644
         FullRedrawOnPan =   -1  'True
         Contents        =   "FrmGCPDS.frx":24AFD
      End
      Begin Rave_Buttons.RaveButtons RaveZoomInBM 
         Height          =   450
         Left            =   -74760
         TabIndex        =   198
         ToolTipText     =   "Zoom Box"
         Top             =   190
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":24B17
         PICN            =   "FrmGCPDS.frx":24B33
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveZoomOutBM 
         Height          =   450
         Left            =   -74265
         TabIndex        =   199
         ToolTipText     =   "Zoom Out"
         Top             =   190
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2515D
         PICN            =   "FrmGCPDS.frx":25179
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RavePanBM 
         Height          =   450
         Left            =   -73770
         TabIndex        =   200
         ToolTipText     =   "Pan"
         Top             =   190
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":257A3
         PICN            =   "FrmGCPDS.frx":257BF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveExtentBM 
         Height          =   450
         Left            =   -73260
         TabIndex        =   201
         ToolTipText     =   "Full Extent"
         Top             =   190
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":25EB9
         PICN            =   "FrmGCPDS.frx":25ED5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.StatusBar StatusBar2 
         Height          =   465
         Left            =   -74790
         TabIndex        =   202
         Top             =   8440
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   820
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   5
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Projection: WGS84  "
               TextSave        =   "Projection: WGS84  "
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   2646
               MinWidth        =   2646
               Text            =   "Tool: None"
               TextSave        =   "Tool: None"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   7056
               MinWidth        =   7056
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   7056
               MinWidth        =   7056
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Rave_Buttons.RaveButtons RaveBM 
         Height          =   450
         Left            =   -72780
         TabIndex        =   203
         ToolTipText     =   "Assign Coordinates"
         Top             =   190
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         BTYPE           =   2
         TX              =   ""
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":264FF
         PICN            =   "FrmGCPDS.frx":2651B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveEditBM 
         Height          =   885
         Left            =   2250
         TabIndex        =   204
         TabStop         =   0   'False
         Top             =   140
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":268D3
         PICN            =   "FrmGCPDS.frx":268EF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveUserAccounts 
         Height          =   540
         Left            =   -71085
         TabIndex        =   205
         TabStop         =   0   'False
         Top             =   3735
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "User Accounts               "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":26DEC
         PICN            =   "FrmGCPDS.frx":26E08
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveDeletedGCPs 
         Height          =   540
         Left            =   -71085
         TabIndex        =   206
         TabStop         =   0   'False
         Top             =   4365
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Deleted GCP                  "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":272DB
         PICN            =   "FrmGCPDS.frx":272F7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons SummaryRave 
         Height          =   540
         Left            =   -71085
         TabIndex        =   207
         TabStop         =   0   'False
         Top             =   4995
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Summary of GCP           "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2783D
         PICN            =   "FrmGCPDS.frx":27859
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveInventory 
         Height          =   540
         Left            =   -71085
         TabIndex        =   208
         TabStop         =   0   'False
         Top             =   1900
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Inventory of Certificate"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":27D56
         PICN            =   "FrmGCPDS.frx":27D72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtonsExport 
         Height          =   540
         Left            =   -71085
         TabIndex        =   209
         TabStop         =   0   'False
         Top             =   5625
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Export GCP                     "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2827C
         PICN            =   "FrmGCPDS.frx":28298
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtonsImport 
         Height          =   540
         Left            =   -71085
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   6255
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Import GCP                     "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2932A
         PICN            =   "FrmGCPDS.frx":29346
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtonsUploaded 
         Height          =   540
         Left            =   -71085
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   6930
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Uploaded GCP              "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2A3D8
         PICN            =   "FrmGCPDS.frx":2A3F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveDMS2Degree 
         Height          =   540
         Left            =   -67665
         TabIndex        =   212
         Top             =   5725
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "DMS to Degree            "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2A70E
         PICN            =   "FrmGCPDS.frx":2A72A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveAddGravity 
         Height          =   885
         Left            =   -73695
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   545
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Add"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2AC29
         PICN            =   "FrmGCPDS.frx":2AC45
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveEditGravity 
         Height          =   885
         Left            =   -72615
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   545
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2B13E
         PICN            =   "FrmGCPDS.frx":2B15A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveDeleteGravity 
         Height          =   885
         Left            =   -71535
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   545
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2B657
         PICN            =   "FrmGCPDS.frx":2B673
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveSaveGravity 
         Height          =   885
         Left            =   -70500
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   545
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2BB40
         PICN            =   "FrmGCPDS.frx":2BB5C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveCancelGravity 
         Height          =   885
         Left            =   -69465
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   545
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2BD5F
         PICN            =   "FrmGCPDS.frx":2BD7B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RavePrintGravity 
         Height          =   885
         Left            =   -68430
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   545
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2C28F
         PICN            =   "FrmGCPDS.frx":2C2AB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveBackGravity 
         Height          =   885
         Left            =   -67350
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   545
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2C598
         PICN            =   "FrmGCPDS.frx":2C5B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveNextGravity 
         Height          =   885
         Left            =   -66315
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   545
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Next"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2C892
         PICN            =   "FrmGCPDS.frx":2C8AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveLocation 
         Height          =   315
         Left            =   -68880
         TabIndex        =   64
         Top             =   2465
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   11110782
         BCOLO           =   11110782
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2CB84
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveSearchGravity 
         Height          =   885
         Left            =   -74775
         TabIndex        =   229
         TabStop         =   0   'False
         Top             =   545
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2CBA0
         PICN            =   "FrmGCPDS.frx":2CBBC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtonsUnits 
         Height          =   390
         Left            =   -68880
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   6215
         Visible         =   0   'False
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   688
         BTYPE           =   5
         TX              =   "m"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2CE93
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveRemoveQuery 
         Height          =   465
         Left            =   -66600
         TabIndex        =   248
         Top             =   2945
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "Remove"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2CEAF
         PICN            =   "FrmGCPDS.frx":2CECB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons15 
         Height          =   450
         Left            =   -67680
         TabIndex        =   249
         TabStop         =   0   'False
         Top             =   6365
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   794
         BTYPE           =   8
         TX              =   "Import from Dbase IV"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2D3C8
         PICN            =   "FrmGCPDS.frx":2D3E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons3 
         Height          =   885
         Left            =   -71640
         TabIndex        =   260
         TabStop         =   0   'False
         Top             =   240
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2E476
         PICN            =   "FrmGCPDS.frx":2E492
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons4 
         Height          =   885
         Left            =   -73800
         TabIndex        =   261
         TabStop         =   0   'False
         Top             =   240
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2E77F
         PICN            =   "FrmGCPDS.frx":2E79B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons5 
         Height          =   885
         Left            =   -72720
         TabIndex        =   262
         TabStop         =   0   'False
         Top             =   240
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Next"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2EA79
         PICN            =   "FrmGCPDS.frx":2EA95
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons6 
         Height          =   885
         Left            =   -74880
         TabIndex        =   263
         TabStop         =   0   'False
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2ED6B
         PICN            =   "FrmGCPDS.frx":2ED87
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveDeleteDuplicateBMs 
         Height          =   435
         Left            =   -67680
         TabIndex        =   274
         TabStop         =   0   'False
         Top             =   6960
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   "Delete Duplicate BMs"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2F05E
         PICN            =   "FrmGCPDS.frx":2F07A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveDup 
         Height          =   885
         Left            =   10800
         TabIndex        =   275
         TabStop         =   0   'False
         Top             =   140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Dup"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2F2E2
         PICN            =   "FrmGCPDS.frx":2F2FE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveButtons7 
         Height          =   540
         Left            =   -71085
         TabIndex        =   282
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   953
         BTYPE           =   8
         TX              =   "Edit Print Records        "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   4210752
         FCOLO           =   8421504
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2F808
         PICN            =   "FrmGCPDS.frx":2F824
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveTranform2WGS84 
         Height          =   465
         Left            =   -70140
         TabIndex        =   287
         Top             =   6060
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "to WGS84"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   32768
         BCOLO           =   32768
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2FD2E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveTranformToPTM 
         Height          =   465
         Left            =   -71100
         TabIndex        =   288
         Top             =   6060
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "to PTM"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   128
         BCOLO           =   128
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2FD4A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Rave_Buttons.RaveButtons RaveTranform2UTM 
         Height          =   465
         Left            =   -69180
         TabIndex        =   289
         Top             =   6060
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "to UTM"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   8421376
         BCOLO           =   8421376
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmGCPDS.frx":2FD66
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.TextBox TxtBMPlus 
         Height          =   420
         Left            =   3840
         TabIndex        =   45
         Top             =   4200
         Width           =   855
         VariousPropertyBits=   1820346387
         BackColor       =   9136220
         BorderStyle     =   1
         Size            =   "1508;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+/- :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   88
         Left            =   3360
         TabIndex        =   302
         Top             =   4320
         Width           =   390
      End
      Begin MSForms.TextBox TxtMSL 
         Height          =   375
         Left            =   -72870
         TabIndex        =   301
         Top             =   2710
         Width           =   3015
         VariousPropertyBits=   1820346391
         BackColor       =   12648447
         ForeColor       =   0
         MaxLength       =   50
         BorderStyle     =   1
         Size            =   "5318;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox txtOrder 
         Height          =   390
         Left            =   -72870
         TabIndex        =   300
         Top             =   2230
         Width           =   3015
         VariousPropertyBits=   746604563
         BackColor       =   12648447
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "5318;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox TxtIsland 
         Height          =   390
         Left            =   -72870
         TabIndex        =   1
         Top             =   1770
         Width           =   3015
         VariousPropertyBits=   746604567
         BackColor       =   12648447
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5318;688"
         ListWidth       =   5291
         ListRows        =   11
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBox5 
         Height          =   375
         Left            =   -74760
         TabIndex        =   122
         Top             =   3300
         Width           =   3225
         VariousPropertyBits=   1820346399
         BackColor       =   9136220
         ForeColor       =   16777215
         Size            =   "5689;661"
         Value           =   "WGS84"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Trebuchet MS"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtRef 
         Height          =   375
         Left            =   -71280
         TabIndex        =   291
         Top             =   3300
         Width           =   3225
         VariousPropertyBits=   1820346399
         BackColor       =   9136220
         ForeColor       =   16777215
         Size            =   "5689;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Trebuchet MS"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00A9897E&
         Height          =   3315
         Left            =   -71280
         Top             =   3360
         Width           =   3225
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00A9897E&
         Height          =   3315
         Left            =   -74760
         Top             =   3360
         Width           =   3225
      End
      Begin MSForms.TextBox txtEllipsoidalH 
         Height          =   390
         Left            =   -71040
         TabIndex        =   20
         Top             =   5580
         Width           =   1155
         VariousPropertyBits=   1820346391
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2037;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         BackColor       =   &H00A9897E&
         BackStyle       =   0  'Transparent
         Caption         =   " Latitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   5
         Left            =   -71040
         TabIndex        =   80
         Top             =   3960
         Width           =   1755
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0ABA3&
         BackStyle       =   0  'Transparent
         Caption         =   " Longitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   6
         Left            =   -71040
         TabIndex        =   299
         Top             =   4680
         Width           =   1755
      End
      Begin VB.Label Label 
         BackColor       =   &H00A9897E&
         BackStyle       =   0  'Transparent
         Caption         =   " Ellip Hgt."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   26
         Left            =   -71040
         TabIndex        =   298
         Top             =   5400
         Width           =   1155
      End
      Begin MSForms.TextBox TxtLatD 
         Height          =   345
         Left            =   -71040
         TabIndex        =   14
         Top             =   4170
         Width           =   645
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         MaxLength       =   2
         BorderStyle     =   1
         Size            =   "1138;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtLatM 
         Height          =   345
         Left            =   -70260
         TabIndex        =   15
         Top             =   4170
         Width           =   525
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         MaxLength       =   2
         BorderStyle     =   1
         Size            =   "926;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtLatS 
         Height          =   345
         Left            =   -69660
         TabIndex        =   16
         Top             =   4170
         Width           =   1215
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2143;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtLongS 
         Height          =   345
         Left            =   -69660
         TabIndex        =   19
         Top             =   4920
         Width           =   1215
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2143;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtLongM 
         Height          =   345
         Left            =   -70260
         TabIndex        =   18
         Top             =   4920
         Width           =   525
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         MaxLength       =   2
         BorderStyle     =   1
         Size            =   "926;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtLongD 
         Height          =   345
         Left            =   -71040
         TabIndex        =   17
         Top             =   4920
         Width           =   645
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         MaxLength       =   3
         BorderStyle     =   1
         Size            =   "1138;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   20
         Left            =   -70380
         TabIndex        =   297
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "'"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   22
         Left            =   -69780
         TabIndex        =   296
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   """"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   23
         Left            =   -68460
         TabIndex        =   295
         Top             =   4080
         Width           =   195
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   """"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   24
         Left            =   -68400
         TabIndex        =   294
         Top             =   4800
         Width           =   195
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "'"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   25
         Left            =   -69780
         TabIndex        =   293
         Top             =   4800
         Width           =   135
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   27
         Left            =   -70380
         TabIndex        =   292
         Top             =   4800
         Width           =   135
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(meters)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   42
         Left            =   -69780
         TabIndex        =   290
         Top             =   5640
         Width           =   885
      End
      Begin MSForms.TextBox TextBox2 
         Height          =   375
         Left            =   -64440
         TabIndex        =   283
         Top             =   3300
         Width           =   3225
         VariousPropertyBits=   1820346399
         BackColor       =   9136220
         ForeColor       =   16777215
         Size            =   "5689;661"
         Value           =   "UTM (PRS92) (m.)"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Trebuchet MS"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxUTMZone 
         Height          =   345
         Left            =   -63090
         TabIndex        =   26
         Top             =   5235
         Width           =   1545
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2725;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxUTMNorthing 
         Height          =   345
         Left            =   -63090
         TabIndex        =   24
         Top             =   4170
         Width           =   1545
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2725;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxUTMEasting 
         Height          =   345
         Left            =   -63090
         TabIndex        =   25
         Top             =   4680
         Width           =   1545
         VariousPropertyBits=   1820346387
         BackColor       =   15855083
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2725;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00A9897E&
         Height          =   3435
         Left            =   -64440
         Top             =   3300
         Width           =   3225
      End
      Begin VB.Label Label 
         BackColor       =   &H00A9897E&
         BackStyle       =   0  'Transparent
         Caption         =   " Zone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   85
         Left            =   -64215
         TabIndex        =   286
         Top             =   5235
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00A9897E&
         BackStyle       =   0  'Transparent
         Caption         =   " Northing:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   84
         Left            =   -64215
         TabIndex        =   285
         Top             =   4170
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0ABA3&
         BackStyle       =   0  'Transparent
         Caption         =   " Easting:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   2
         Left            =   -64215
         TabIndex        =   284
         Top             =   4680
         Width           =   1095
      End
      Begin MSForms.ComboBox TxtEIsland 
         Height          =   420
         Left            =   7080
         TabIndex        =   52
         Top             =   4200
         Width           =   2265
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3995;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label LabelEncoder2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   240
         TabIndex        =   279
         Top             =   8520
         Width           =   360
      End
      Begin VB.Label LabelDateUpdated2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   6540
         TabIndex        =   278
         Top             =   8520
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encoder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   240
         TabIndex        =   277
         Top             =   8790
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Last Updated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   6540
         TabIndex        =   276
         Top             =   8790
         Width           =   1440
      End
      Begin MSForms.TextBox TextBoxOldBookmark 
         Height          =   405
         Left            =   -67560
         TabIndex        =   272
         Top             =   8400
         Width           =   5880
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   49152
         Size            =   "10372;714"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   79
         Left            =   -62400
         TabIndex        =   271
         Top             =   8640
         Width           =   615
      End
      Begin MSForms.TextBox TextBoxOldOrder 
         Height          =   390
         Left            =   -74400
         TabIndex        =   76
         Top             =   2400
         Width           =   4575
         VariousPropertyBits=   1820346391
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   50
         Size            =   "8070;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxOldLongitude 
         Height          =   390
         Left            =   -74400
         TabIndex        =   78
         Top             =   3840
         Width           =   4620
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         Size            =   "8149;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxOldLatitude 
         Height          =   390
         Left            =   -74400
         TabIndex        =   77
         Top             =   3120
         Width           =   4575
         VariousPropertyBits=   1820346391
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   50
         Size            =   "8070;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxOldDescription 
         Height          =   2235
         Left            =   -74400
         TabIndex        =   270
         Top             =   5760
         Width           =   12720
         VariousPropertyBits=   -326088673
         BackColor       =   16777215
         ForeColor       =   0
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "22437;3942"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   83
         Left            =   -74280
         TabIndex        =   269
         Top             =   8040
         Width           =   1515
      End
      Begin VB.Label Label 
         BackColor       =   &H00A9897E&
         BackStyle       =   0  'Transparent
         Caption         =   " Latitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   81
         Left            =   -74280
         TabIndex        =   268
         Top             =   3480
         Width           =   1755
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0ABA3&
         BackStyle       =   0  'Transparent
         Caption         =   " Longitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   80
         Left            =   -74280
         TabIndex        =   267
         Top             =   4200
         Width           =   1755
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Entry:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   78
         Left            =   -74280
         TabIndex        =   266
         Top             =   5040
         Width           =   990
      End
      Begin MSForms.TextBox TextBoxOldDateEntry 
         Height          =   390
         Left            =   -74400
         TabIndex        =   79
         Top             =   4680
         Width           =   4620
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         Size            =   "8149;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Established:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   77
         Left            =   -69000
         TabIndex        =   265
         Top             =   5040
         Width           =   1245
      End
      Begin MSForms.TextBox TextBoxOldDateEstablished 
         Height          =   405
         Left            =   -69120
         TabIndex        =   264
         Top             =   4680
         Width           =   5880
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         Size            =   "10372;714"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Province:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   76
         Left            =   -69000
         MouseIcon       =   "FrmGCPDS.frx":2FD82
         MousePointer    =   99  'Custom
         TabIndex        =   259
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Municipality:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   70
         Left            =   -69000
         MouseIcon       =   "FrmGCPDS.frx":3008C
         MousePointer    =   99  'Custom
         TabIndex        =   258
         Top             =   3480
         Width           =   870
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   69
         Left            =   -69000
         MouseIcon       =   "FrmGCPDS.frx":30396
         MousePointer    =   99  'Custom
         TabIndex        =   257
         Top             =   4200
         Width           =   750
      End
      Begin MSForms.TextBox TextBoxOldBarangay 
         Height          =   345
         Left            =   -69120
         TabIndex        =   256
         Top             =   3840
         Width           =   5955
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         Size            =   "10504;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxOldMunicipality 
         Height          =   345
         Left            =   -69120
         TabIndex        =   255
         Top             =   3120
         Width           =   5955
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         Size            =   "10504;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxOldProvince 
         Height          =   345
         Left            =   -69120
         TabIndex        =   254
         Top             =   2400
         Width           =   5880
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         Size            =   "10372;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Station Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   67
         Left            =   -74280
         MouseIcon       =   "FrmGCPDS.frx":306A0
         MousePointer    =   99  'Custom
         TabIndex        =   253
         Top             =   2040
         Width           =   990
      End
      Begin MSForms.TextBox TextBoxOldStationName 
         Height          =   390
         Left            =   -74400
         TabIndex        =   75
         Top             =   1680
         Width           =   4575
         VariousPropertyBits=   1820346391
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   50
         Size            =   "8070;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   66
         Left            =   -74280
         MouseIcon       =   "FrmGCPDS.frx":309AA
         MousePointer    =   99  'Custom
         TabIndex        =   252
         Top             =   2760
         Width           =   465
      End
      Begin MSForms.TextBox TextBoxOldRegion 
         Height          =   405
         Left            =   -69120
         TabIndex        =   251
         Top             =   1680
         Width           =   5895
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         Size            =   "10398;714"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   65
         Left            =   -69000
         MouseIcon       =   "FrmGCPDS.frx":30CB4
         MousePointer    =   99  'Custom
         TabIndex        =   250
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label LabelGravityRecordStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   -66405
         TabIndex        =   247
         Top             =   8465
         Width           =   5490
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   -61560
         TabIndex        =   246
         Top             =   8735
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   -61680
         TabIndex        =   245
         Top             =   8735
         Width           =   645
      End
      Begin VB.Label LabelencoderGravity 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   -74730
         TabIndex        =   244
         Top             =   8510
         Width           =   360
      End
      Begin VB.Label LabelUpdatedGravity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   -68430
         TabIndex        =   243
         Top             =   8510
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encoder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   -74730
         TabIndex        =   242
         Top             =   8780
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Last Updated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   -68430
         TabIndex        =   241
         Top             =   8780
         Width           =   1440
      End
      Begin MSForms.ComboBox ComboBoxOrderGravity 
         Height          =   345
         Left            =   -72165
         TabIndex        =   69
         Top             =   4505
         Width           =   3225
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         MaxLength       =   20
         DisplayStyle    =   7
         Size            =   "5689;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observed Values:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   68
         Left            =   -74370
         TabIndex        =   240
         Top             =   5900
         Width           =   1890
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   82
         Left            =   -74370
         TabIndex        =   239
         Top             =   4550
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Station Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   87
         Left            =   -74370
         TabIndex        =   238
         Top             =   1670
         Width           =   1470
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   93
         Left            =   -74370
         TabIndex        =   237
         Top             =   2480
         Width           =   825
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Province:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   73
         Left            =   -74370
         TabIndex        =   236
         Top             =   2930
         Width           =   1020
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Municipality"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   74
         Left            =   -74370
         TabIndex        =   235
         Top             =   3380
         Width           =   1260
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   71
         Left            =   -74370
         TabIndex        =   234
         Top             =   3830
         Width           =   1050
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   72
         Left            =   -74370
         TabIndex        =   233
         Top             =   5000
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   75
         Left            =   -74370
         TabIndex        =   232
         Top             =   5495
         Width           =   1140
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Elevation:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   64
         Left            =   -74370
         TabIndex        =   231
         Top             =   6260
         Width           =   1065
      End
      Begin MSForms.TextBox TextBoxGravityElevation 
         Height          =   375
         Left            =   -72165
         TabIndex        =   73
         Top             =   6305
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   20
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxObservedValues 
         Height          =   375
         Left            =   -72165
         TabIndex        =   72
         Top             =   5855
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   20
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxGravityLongitude 
         Height          =   375
         Left            =   -72165
         TabIndex        =   71
         Top             =   5405
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   20
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxGravityLatitude 
         Height          =   375
         Left            =   -72165
         TabIndex        =   70
         Top             =   4955
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   20
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxGravityBarangay 
         Height          =   375
         Left            =   -72165
         TabIndex        =   68
         Top             =   3785
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   50
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxGravityMunicipality 
         Height          =   375
         Left            =   -72165
         TabIndex        =   67
         Top             =   3335
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   50
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxGravityProvince 
         Height          =   375
         Left            =   -72165
         TabIndex        =   66
         Top             =   2885
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   50
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxGravityRegion 
         Height          =   375
         Left            =   -72165
         TabIndex        =   65
         Top             =   2435
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   50
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxGravityName 
         Height          =   375
         Left            =   -72165
         TabIndex        =   63
         Top             =   1625
         Width           =   3210
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         MaxLength       =   50
         Size            =   "5662;661"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   94
         Left            =   -67890
         TabIndex        =   228
         Top             =   1535
         Width           =   2265
      End
      Begin MSForms.TextBox TextBoxGravityDescription 
         Height          =   4245
         Left            =   -67890
         TabIndex        =   74
         Top             =   2435
         Width           =   6495
         VariousPropertyBits=   -326088673
         BackColor       =   16777215
         ForeColor       =   4210752
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "11456;7488"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Last Updated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   -68430
         TabIndex        =   218
         Top             =   8785
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encoder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   -74730
         TabIndex        =   217
         Top             =   8785
         Width           =   615
      End
      Begin VB.Label LabelDateUpdated 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   -68430
         TabIndex        =   216
         Top             =   8515
         Width           =   360
      End
      Begin VB.Label LabelEncoder 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   -74730
         TabIndex        =   215
         Top             =   8515
         Width           =   360
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MSL Elevation:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   86
         Left            =   -74610
         TabIndex        =   214
         Top             =   2730
         Width           =   1605
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   63
         Left            =   9570
         TabIndex        =   195
         Top             =   4245
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   62
         Left            =   9570
         TabIndex        =   194
         Top             =   4800
         Width           =   1140
      End
      Begin MSForms.TextBox TextBoxLatitude 
         Height          =   420
         Left            =   11520
         TabIndex        =   60
         Top             =   4125
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TextBoxLongitude 
         Height          =   420
         Left            =   11520
         TabIndex        =   61
         Top             =   4665
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   16777215
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.OptionButton optBMs 
         Height          =   375
         Left            =   -73110
         TabIndex        =   191
         Top             =   3075
         Width           =   1845
         VariousPropertyBits=   746588179
         BackColor       =   0
         ForeColor       =   32768
         DisplayStyle    =   5
         Size            =   "3254;661"
         Value           =   "0"
         Caption         =   "Benchmarks"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.OptionButton optGCPs 
         Height          =   330
         Left            =   -74460
         TabIndex        =   190
         Top             =   3075
         Width           =   1140
         VariousPropertyBits=   746588179
         BackColor       =   0
         ForeColor       =   32768
         DisplayStyle    =   5
         Size            =   "2011;582"
         Value           =   "1"
         Caption         =   "GCPs"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   40
         Left            =   5880
         TabIndex        =   186
         Top             =   5760
         Width           =   2265
      End
      Begin MSForms.TextBox txtYear 
         Height          =   405
         Left            =   -61995
         TabIndex        =   36
         Top             =   7410
         Width           =   1110
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "1958;714"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox ComboBoxDay 
         Height          =   405
         Left            =   -62805
         TabIndex        =   35
         Top             =   7410
         Width           =   765
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1349;714"
         ListWidth       =   5291
         ListRows        =   13
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
         Object.Width           =   "352"
      End
      Begin MSForms.TextBox txtEstablished 
         Height          =   405
         Left            =   -63960
         TabIndex        =   185
         Top             =   1800
         Visible         =   0   'False
         Width           =   1875
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "3307;714"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox ComboBoxMonth 
         Height          =   405
         Left            =   -63885
         TabIndex        =   34
         Top             =   7410
         Width           =   1035
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1826;714"
         ListWidth       =   5291
         ListRows        =   13
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
         Object.Width           =   "352"
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   61
         Left            =   405
         TabIndex        =   174
         Top             =   2130
         Width           =   825
      End
      Begin MSForms.TextBox TxtERegion 
         Height          =   420
         Left            =   2115
         TabIndex        =   40
         Top             =   2055
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   13
         Left            =   -69195
         MouseIcon       =   "FrmGCPDS.frx":30FBE
         MousePointer    =   99  'Custom
         TabIndex        =   173
         Top             =   1360
         Width           =   825
      End
      Begin MSForms.TextBox TxtRegion 
         Height          =   405
         Left            =   -67545
         TabIndex        =   3
         Top             =   1315
         Width           =   2655
         VariousPropertyBits=   1820346391
         BackColor       =   12648447
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4683;714"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label FormCaption2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   11175
         TabIndex        =   172
         Top             =   8565
         Width           =   3015
      End
      Begin MSForms.TextBox TxtBMAuthority 
         Height          =   420
         Left            =   11520
         TabIndex        =   59
         Top             =   3585
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtElevationAuthority 
         Height          =   420
         Left            =   11520
         TabIndex        =   58
         Top             =   3045
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtBMDateEstablished 
         Height          =   420
         Left            =   11520
         TabIndex        =   57
         Top             =   2505
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox TxtBMOrder 
         Height          =   420
         Left            =   7080
         TabIndex        =   50
         Top             =   3105
         Width           =   2265
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3995;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox TxtEDatum 
         Height          =   420
         Left            =   7080
         TabIndex        =   51
         Top             =   3645
         Width           =   2265
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3995;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox BMMarkType 
         Height          =   420
         Left            =   11520
         TabIndex        =   54
         Top             =   1425
         Width           =   2625
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox BMMarkPurpose 
         Height          =   420
         Left            =   7080
         TabIndex        =   53
         Top             =   4725
         Width           =   2265
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3995;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox BMMarkStatus 
         Height          =   420
         Left            =   11520
         TabIndex        =   56
         Top             =   1965
         Width           =   2625
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtEProvince 
         Height          =   420
         Left            =   2115
         TabIndex        =   41
         Top             =   2595
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical Datum:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   60
         Left            =   5130
         TabIndex        =   171
         Top             =   3720
         Width           =   1590
      End
      Begin MSForms.ComboBox TxtEFix 
         Height          =   420
         Left            =   2115
         TabIndex        =   46
         Top             =   4800
         Width           =   2625
         VariousPropertyBits=   746604567
         BackColor       =   16777215
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtBMDescription 
         Height          =   1905
         Left            =   240
         TabIndex        =   62
         Top             =   6240
         Width           =   13920
         VariousPropertyBits=   -326088673
         BackColor       =   16777215
         ForeColor       =   0
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "24553;3360"
         BorderColor     =   0
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fixing Method:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   59
         Left            =   405
         TabIndex        =   170
         Top             =   4830
         Width           =   1575
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Elevation (m):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   58
         Left            =   405
         TabIndex        =   169
         Top             =   4290
         Width           =   1455
      End
      Begin MSForms.TextBox TxtElevation 
         Height          =   420
         Left            =   2115
         TabIndex        =   44
         Top             =   4200
         Width           =   1230
         VariousPropertyBits=   1820346391
         BackColor       =   9136220
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2170;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Province:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   57
         Left            =   405
         TabIndex        =   168
         Top             =   2670
         Width           =   1020
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Municipality:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   56
         Left            =   405
         TabIndex        =   167
         Top             =   3210
         Width           =   1320
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   55
         Left            =   405
         TabIndex        =   166
         Top             =   3750
         Width           =   1050
      End
      Begin MSForms.TextBox TxtEBarangay 
         Height          =   420
         Left            =   2115
         TabIndex        =   43
         Top             =   3675
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtEMunicipality 
         Height          =   420
         Left            =   2115
         TabIndex        =   42
         Top             =   3135
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Station Name:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   300
         Index           =   54
         Left            =   390
         TabIndex        =   165
         Top             =   1560
         Width           =   1500
      End
      Begin MSForms.TextBox TxtEName 
         Height          =   420
         Left            =   2115
         TabIndex        =   38
         Top             =   1515
         Width           =   2625
         VariousPropertyBits=   1820346391
         BackColor       =   9136220
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4630;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtBMDateComputed 
         Height          =   420
         Left            =   7065
         TabIndex        =   48
         Top             =   1995
         Width           =   2265
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "3995;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtBMDateLastRecovered 
         Height          =   420
         Left            =   7065
         TabIndex        =   49
         Top             =   2580
         Width           =   2265
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "3995;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Purpose:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   53
         Left            =   5130
         TabIndex        =   164
         Top             =   4800
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   52
         Left            =   9570
         TabIndex        =   163
         Top             =   1500
         Width           =   1170
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Recovered:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   51
         Left            =   5160
         TabIndex        =   162
         Top             =   2655
         Width           =   1755
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Established:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   50
         Left            =   9570
         TabIndex        =   161
         Top             =   2580
         Width           =   1860
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resp. Authority:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   49
         Left            =   9570
         TabIndex        =   160
         Top             =   3660
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   48
         Left            =   9570
         TabIndex        =   159
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Computed:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   47
         Left            =   5160
         TabIndex        =   158
         Top             =   2070
         Width           =   1710
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Established by:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   46
         Left            =   9570
         TabIndex        =   157
         Top             =   3120
         Width           =   1605
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   45
         Left            =   5130
         TabIndex        =   156
         Top             =   3180
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Island:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   44
         Left            =   5130
         TabIndex        =   155
         Top             =   4260
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Entry:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   39
         Left            =   5160
         TabIndex        =   154
         Top             =   1530
         Width           =   1440
      End
      Begin MSForms.TextBox TxtBMDateOfEntry 
         Height          =   420
         Left            =   7065
         TabIndex        =   47
         Top             =   1455
         Width           =   2265
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "3995;741"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(meters)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Index           =   43
         Left            =   -73200
         TabIndex        =   144
         Top             =   5640
         Width           =   885
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   41
         Left            =   -60840
         TabIndex        =   143
         Top             =   8565
         Width           =   45
      End
      Begin VB.Label FormCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   -66525
         TabIndex        =   142
         Top             =   8465
         Width           =   5490
      End
      Begin MSForms.TextBox txtAuthority 
         Height          =   405
         Left            =   -63885
         TabIndex        =   37
         Top             =   7890
         Width           =   3000
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "5292;714"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label PageCounterLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   -68835
         TabIndex        =   138
         Top             =   7300
         Width           =   2055
      End
      Begin VB.Label ResultLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   -74865
         TabIndex        =   132
         Top             =   7165
         Width           =   75
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   38
         Left            =   -73800
         TabIndex        =   131
         Top             =   4020
         Width           =   135
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "'"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   37
         Left            =   -73200
         TabIndex        =   130
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   """"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   35
         Left            =   -71880
         TabIndex        =   129
         Top             =   4020
         Width           =   195
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   """"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   34
         Left            =   -71880
         TabIndex        =   128
         Top             =   4830
         Width           =   195
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "'"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   33
         Left            =   -73200
         TabIndex        =   127
         Top             =   4830
         Width           =   135
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   28
         Left            =   -73800
         TabIndex        =   126
         Top             =   4830
         Width           =   135
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   14
         Left            =   -74520
         TabIndex        =   125
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   16
         Left            =   -74520
         TabIndex        =   124
         Top             =   4680
         Width           =   750
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ellip Hgt."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   210
         Index           =   21
         Left            =   -74520
         TabIndex        =   123
         Top             =   5400
         Width           =   600
      End
      Begin MSForms.TextBox TxtDLong 
         Height          =   345
         Left            =   -74520
         TabIndex        =   10
         Top             =   4920
         Width           =   645
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         MaxLength       =   3
         BorderStyle     =   1
         Size            =   "1138;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtMLong 
         Height          =   345
         Left            =   -73680
         TabIndex        =   11
         Top             =   4920
         Width           =   525
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         MaxLength       =   2
         BorderStyle     =   1
         Size            =   "926;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtSLong 
         Height          =   345
         Left            =   -73080
         TabIndex        =   12
         Top             =   4920
         Width           =   1215
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2143;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtSLat 
         Height          =   345
         Left            =   -73080
         TabIndex        =   9
         Top             =   4170
         Width           =   1215
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2143;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtMLat 
         Height          =   345
         Left            =   -73680
         TabIndex        =   8
         Top             =   4170
         Width           =   525
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         MaxLength       =   2
         BorderStyle     =   1
         Size            =   "926;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtDLat 
         Height          =   345
         Left            =   -74520
         TabIndex        =   7
         Top             =   4200
         Width           =   645
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         MaxLength       =   2
         BorderStyle     =   1
         Size            =   "1138;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtDateEntry 
         Height          =   390
         Left            =   -72870
         TabIndex        =   27
         Top             =   6915
         Width           =   2340
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4128;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Entry:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   19
         Left            =   -74760
         TabIndex        =   116
         Top             =   6975
         Width           =   1440
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   375
         Left            =   -67860
         TabIndex        =   100
         Top             =   3300
         Width           =   3225
         VariousPropertyBits=   1820346399
         BackColor       =   9136220
         ForeColor       =   16777215
         Size            =   "5689;661"
         Value           =   "PTM (m.)"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Trebuchet MS"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Island:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   3
         Left            =   -74610
         MouseIcon       =   "FrmGCPDS.frx":312C8
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   1840
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   0
         Left            =   -74610
         MouseIcon       =   "FrmGCPDS.frx":315D2
         MousePointer    =   99  'Custom
         TabIndex        =   114
         Top             =   2260
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surveyed by:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   15
         Left            =   -65805
         MouseIcon       =   "FrmGCPDS.frx":318DC
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   6975
         Width           =   1380
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fixing Method:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   17
         Left            =   -63960
         TabIndex        =   112
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Computed:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   18
         Left            =   -74760
         TabIndex        =   111
         Top             =   7470
         Width           =   1710
      End
      Begin MSForms.ComboBox TxtFixingMethod 
         Height          =   390
         Left            =   -63960
         TabIndex        =   81
         Top             =   1440
         Visible         =   0   'False
         Width           =   2340
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4128;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtEstablishedBy 
         Height          =   405
         Left            =   -63885
         TabIndex        =   33
         Top             =   6915
         Width           =   3000
         VariousPropertyBits=   1820346391
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "5292;714"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   32
         Left            =   -70095
         TabIndex        =   110
         Top             =   7950
         Width           =   1320
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resp. Authority:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   31
         Left            =   -65790
         TabIndex        =   109
         Top             =   7950
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Established:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   9
         Left            =   -65775
         TabIndex        =   108
         Top             =   7470
         Width           =   1860
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Recover:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   10
         Left            =   -74760
         TabIndex        =   107
         Top             =   7950
         Width           =   1485
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   29
         Left            =   -70095
         TabIndex        =   106
         Top             =   7470
         Width           =   1170
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Purpose:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   30
         Left            =   -70095
         TabIndex        =   105
         Top             =   6975
         Width           =   1545
      End
      Begin MSForms.ComboBox txtMarkPurpose 
         Height          =   405
         Left            =   -68445
         TabIndex        =   30
         Top             =   6915
         Width           =   2475
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4366;714"
         ListWidth       =   5291
         ListRows        =   11
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox txtMarkType 
         Height          =   405
         Left            =   -68445
         TabIndex        =   31
         Top             =   7410
         Width           =   2475
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4366;714"
         ListWidth       =   5291
         ListRows        =   20
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox txtMarkStatus 
         Height          =   405
         Left            =   -68445
         TabIndex        =   32
         Top             =   7890
         Width           =   2475
         VariousPropertyBits=   746604567
         BackColor       =   14737632
         ForeColor       =   0
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4366;714"
         ListRows        =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         BackColor       =   &H00C0ABA3&
         BackStyle       =   0  'Transparent
         Caption         =   " Easting:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   1
         Left            =   -67755
         TabIndex        =   104
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00A9897E&
         BackStyle       =   0  'Transparent
         Caption         =   " Northing:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   11
         Left            =   -67755
         TabIndex        =   103
         Top             =   4170
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00A9897E&
         BackStyle       =   0  'Transparent
         Caption         =   " Zone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   345
         Index           =   12
         Left            =   -67755
         TabIndex        =   102
         Top             =   5235
         Width           =   1095
      End
      Begin MSForms.TextBox TxtEllipsoidalH2 
         Height          =   390
         Left            =   -74520
         TabIndex        =   13
         Top             =   5640
         Width           =   1260
         VariousPropertyBits=   1820346387
         BackColor       =   10972206
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2222;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00A9897E&
         Height          =   3315
         Left            =   -67860
         Top             =   3420
         Width           =   3225
      End
      Begin MSForms.TextBox txtLastRecover 
         Height          =   390
         Left            =   -72885
         TabIndex        =   29
         Top             =   7890
         Width           =   2340
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4128;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtDateComputed 
         Height          =   390
         Left            =   -72870
         TabIndex        =   28
         Top             =   7410
         Width           =   2340
         VariousPropertyBits=   1820346391
         BackColor       =   11110782
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4128;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtName 
         Height          =   390
         Left            =   -72870
         TabIndex        =   0
         Top             =   1315
         Width           =   3015
         VariousPropertyBits=   1820346391
         BackColor       =   12648447
         ForeColor       =   0
         MaxLength       =   50
         BorderStyle     =   1
         Size            =   "5318;688"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Station Name:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   300
         Index           =   7
         Left            =   -74610
         MouseIcon       =   "FrmGCPDS.frx":31BE6
         MousePointer    =   99  'Custom
         TabIndex        =   101
         Top             =   1360
         Width           =   1500
      End
      Begin MSForms.TextBox TxtEasting 
         Height          =   345
         Left            =   -66630
         TabIndex        =   22
         Top             =   4680
         Width           =   1545
         VariousPropertyBits=   1820346387
         BackColor       =   15855083
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2725;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtNorthing 
         Height          =   345
         Left            =   -66630
         TabIndex        =   21
         Top             =   4170
         Width           =   1545
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2725;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtZone 
         Height          =   345
         Left            =   -66630
         TabIndex        =   23
         Top             =   5235
         Width           =   1545
         VariousPropertyBits=   1820346387
         BackColor       =   15261917
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "2725;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtProvince 
         Height          =   345
         Left            =   -67545
         TabIndex        =   4
         Top             =   1795
         Width           =   2655
         VariousPropertyBits=   1820346391
         BackColor       =   12648447
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "4683;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtMunicipality 
         Height          =   345
         Left            =   -67545
         TabIndex        =   5
         Top             =   2260
         Width           =   5955
         VariousPropertyBits=   1820346391
         BackColor       =   12648447
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "10504;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TxtBarangay 
         Height          =   345
         Left            =   -67545
         TabIndex        =   6
         Top             =   2685
         Width           =   5835
         VariousPropertyBits=   1820346391
         BackColor       =   12648447
         ForeColor       =   0
         BorderStyle     =   1
         Size            =   "10292;609"
         BorderColor     =   11110782
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   8
         Left            =   -69195
         MouseIcon       =   "FrmGCPDS.frx":31EF0
         MousePointer    =   99  'Custom
         TabIndex        =   99
         Top             =   2710
         Width           =   1050
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Municipality:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   4
         Left            =   -69195
         MouseIcon       =   "FrmGCPDS.frx":321FA
         MousePointer    =   99  'Custom
         TabIndex        =   98
         Top             =   2290
         Width           =   1320
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Province:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C7C7C&
         Height          =   270
         Index           =   36
         Left            =   -69195
         MouseIcon       =   "FrmGCPDS.frx":32504
         MousePointer    =   99  'Custom
         TabIndex        =   97
         Top             =   1825
         Width           =   1020
      End
   End
   Begin Rave_Buttons.RaveButtons RaveTab 
      Height          =   405
      Index           =   1
      Left            =   10305
      TabIndex        =   82
      Top             =   1125
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "Query"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8260132
      BCOLO           =   8260132
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmGCPDS.frx":3280E
      PICN            =   "FrmGCPDS.frx":3282A
      PICH            =   "FrmGCPDS.frx":350F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveTab 
      Height          =   405
      Index           =   2
      Left            =   12375
      TabIndex        =   83
      Top             =   1125
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "Utilities"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8260132
      BCOLO           =   8260132
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmGCPDS.frx":381C3
      PICN            =   "FrmGCPDS.frx":381DF
      PICH            =   "FrmGCPDS.frx":3AAA5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveTab 
      Height          =   405
      Index           =   0
      Left            =   2145
      TabIndex        =   84
      Top             =   1125
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "GCPs"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8260132
      BCOLO           =   8260132
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmGCPDS.frx":3DB78
      PICN            =   "FrmGCPDS.frx":3DB94
      PICH            =   "FrmGCPDS.frx":4045A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   -1  'True
   End
   Begin Rave_Buttons.RaveButtons RaveTab 
      Height          =   405
      Index           =   3
      Left            =   4125
      TabIndex        =   145
      Top             =   1125
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "Benchmarks"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8260132
      BCOLO           =   8260132
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmGCPDS.frx":4352D
      PICN            =   "FrmGCPDS.frx":43549
      PICH            =   "FrmGCPDS.frx":45E0F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin CCRProgressBar6.ccrpProgressBar ProgressBar1 
      Height          =   375
      Left            =   11880
      Top             =   675
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      AutoCaption     =   1
      BackColor       =   4194304
      Caption         =   "0%"
      FillColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin Rave_Buttons.RaveButtons RaveTab 
      Height          =   405
      Index           =   4
      Left            =   6195
      TabIndex        =   219
      Top             =   1125
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "Gravity"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8260132
      BCOLO           =   8260132
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmGCPDS.frx":48EE2
      PICN            =   "FrmGCPDS.frx":48EFE
      PICH            =   "FrmGCPDS.frx":4B7C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.Label LabelLogout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      MouseIcon       =   "FrmGCPDS.frx":4E897
      MousePointer    =   99  'Custom
      TabIndex        =   281
      Top             =   480
      Width           =   495
   End
   Begin VB.Label LabelUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   280
      Top             =   240
      Width           =   75
   End
   Begin MBMouseHelper.MouseHelper MouseHelper1 
      Left            =   13080
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin VB.Label LabelProgress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   10800
      TabIndex        =   213
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label LabelBuild 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Build: GNIS 02-Jul-2014"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5910
      TabIndex        =   193
      Top             =   720
      Width           =   2505
   End
   Begin MSForms.TextBox TxtLongitude 
      Height          =   345
      Left            =   150
      TabIndex        =   120
      Top             =   480
      Visible         =   0   'False
      Width           =   2025
      VariousPropertyBits=   1820346391
      BackColor       =   15855083
      ForeColor       =   0
      Size            =   "3572;609"
      BorderColor     =   8260132
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TxtLatitude 
      Height          =   345
      Left            =   750
      TabIndex        =   119
      Top             =   960
      Visible         =   0   'False
      Width           =   2025
      VariousPropertyBits=   1820346391
      BackColor       =   15261917
      ForeColor       =   0
      Size            =   "3572;609"
      BorderColor     =   8260132
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image ImageBorder 
      Height          =   1485
      Left            =   30
      MousePointer    =   15  'Size All
      Top             =   60
      Width           =   14505
   End
End
Attribute VB_Name = "FrmGCPDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sngListViewX As Single
Private sngListViewY As Single
Private m_colEvents As Collection

Private Sub CmdAddUser_Click()
    FrmNewUser.Show 1
End Sub

Private Sub cmdLocation_Click()
    regionTextbox = FrmGCPDS.TxtRegion.name
    provinceTextbox = FrmGCPDS.TxtProvince.name
    municipalityTextbox = FrmGCPDS.TxtMunicipality.name
    barangayTextbox = FrmGCPDS.TxtBarangay.name
    FrmLocation.Show 1
End Sub



Private Sub DMStoDegree()

    Dim i
    Dim rst As New ADODB.Recordset
    rst.Open "select * from geoprov", cnn, adOpenStatic
    
    Me.ProgressBar1.Visible = True
    Me.ProgressBar1.Value = 0
    Me.LabelProgress.Visible = True
    Me.LabelProgress = "DMS to Degree"
    
        For i = 1 To rst.RecordCount
        
        Me.ProgressBar1.Value = i * (100 / rst.RecordCount)
        
        DoEvents
            
            If IsNull(rst("Wgs84ED")) = False Then
                
                cnn.Execute "update geoprov set latitude=" & rst("Wgs84ED") + (rst("Wgs84EM") / 60 + rst("Wgs84ES") / 3600) & " where stat_name='" & Replace(rst("stat_name"), "'", "''") & "'"
                cnn.Execute "update geoprov set longitude=" & rst("Wgs84ND") + (rst("Wgs84NM") / 60 + rst("Wgs84NS") / 3600) & " where stat_name='" & Replace(rst("stat_name"), "'", "''") & "'"
                
            End If
            
        rst.MoveNext
        
        
        Next
    
    
    Me.ProgressBar1.Visible = False
    Me.ProgressBar1.Value = 0
    Me.LabelProgress.Visible = False
    Me.LabelProgress = ""
    MsgBox "Done"


End Sub










Private Sub Command2_Click()

End Sub



Private Sub Command1_Click()
Me.ActiveControl.Type
End Sub

'Private Sub Command2_Click()
'  EditMode = False
'    AddMode = True
'    BlankForm
'    AddEditMode
'    EnableFields
'    Me.TxtDateEntry = Format(Date, "mm-dd-yyyy")
'    Me.txtAuthority = "NAMRIA"
'    Me.txtMarkPurpose = "Geodetic"
'    Me.TxtFixingMethod = "GPS"
'    Me.txtMarkStatus = "Existing"
'    Me.txtMarkType = "Copper Nail"
'    Me.TxtName.SetFocus
'    'FrmLocation.Show 1
'End Sub

Private Sub ExtractWGS84Rave_Click()
If MsgBox("This process will populate the wgs84 coordinates of all the prs92 stations base on the description.", vbOKCancel, "Extract?") = vbCancel Then
    Exit Sub
End If

FrmWGS84.Show 1
End Sub


Public Sub InitializeIsland()
    Me.TxtEIsland.AddItem ""
    Me.TxtEIsland.AddItem "Luzon"
    Me.TxtEIsland.AddItem "Visayas"
    Me.TxtEIsland.AddItem "Mindanao"
    Me.TxtIsland.AddItem ""
    Me.TxtIsland.AddItem "Luzon"
    Me.TxtIsland.AddItem "Visayas"
    Me.TxtIsland.AddItem "Mindanao"
End Sub


Private Sub Form_Initialize()
    Dim c As Control
    Dim ce As ControlEvents
    
    Set m_colEvents = New Collection
    
    
    For Each c In Me.Controls
   
        If TypeOf c Is MSForms.TextBox Or TypeOf c Is MSForms.ComboBox Then
            Set ce = New ControlEvents
            ce.Init c
            m_colEvents.Add ce
           
        End If
    Next c
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If (KeyCode = vbKeyPageUp And SSTab1.Tab = 0 And RaveNext.Enabled = True) Then
    RaveNext_Click
End If

If (KeyCode = vbKeyPageDown And SSTab1.Tab = 0 And RaveBack.Enabled = True) Then
    RaveBack_Click
End If

If (KeyCode = vbKeyEscape And SSTab1.Tab = 0 And RaveCancel.Enabled = True) Then
    RaveCancel_Click
End If

If (Shift = 2 And KeyCode = 83 And SSTab1.Tab = 0 And RaveSave.Enabled = True) Then
    RaveSave_Click
End If

If (Shift = 2 And KeyCode = vbKeyF And SSTab1.Tab = 0 And RaveSearch.Enabled = True) Then
    RaveSearch_Click
End If


If (KeyCode = vbKeyPageUp And SSTab1.Tab = 5 And RaveNextBM.Enabled = True) Then
    RaveNextBM_Click
End If

If (KeyCode = vbKeyPageDown And SSTab1.Tab = 5 And RaveBackBM.Enabled = True) Then
    RaveBackBM_Click
End If

If (KeyCode = vbKeyEscape And SSTab1.Tab = 5 And RaveCancelBM.Enabled = True) Then
    RaveCancelBM_Click
End If

If (Shift = 2 And KeyCode = 83 And SSTab1.Tab = 5 And RaveSaveBM.Enabled = True) Then
    RaveSaveBM_Click
End If

If (Shift = 2 And KeyCode = vbKeyF And SSTab1.Tab = 5 And RaveSearchBM.Enabled = True) Then
    RaveSearchBM_Click
End If
End Sub

Private Sub Form_Load()





UpdateTables
InitializeIsland
InitializeRaves
InitializeMap
InitializeMapBM
DisableFields
EnableBenchmarksFields False

TabVisibleFalse



        'GCPS
     LoadMonths
     LoadDays
     LoadOrder
     'LoadAdopters
     LoadMarkPurpose
     LoadMarkType
     LoadMarkStatus
     LoadHorizontalFixingMethod
     
     'Benchmarks
     
     LoadOrderBM
     LoadVerticalFixingMethod
     LoadVerticalDatum
     
     'Gravity
     
     LoadOrderGravity
     
     
     
rstRecords.CursorLocation = adUseClient
rstRecords.Open "Select stat_name,wgs84ED,wgs84EM,wgs84Es,wgs84ND,wgs84NM,wgs84NS,geoprov.h_order from geoprov where ltrim(h_ref)='PRS92' order by region,province,municipal,barangay,stat_name", cnn, adOpenStatic

rstBenchmarks.CursorLocation = adUseClient
rstBenchmarks.Open "Select stat_name,ucode from benchmarks order by region,province,municipal,barangay,stat_name", cnn, adOpenStatic
  
rstGravity.CursorLocation = adUseClient
rstGravity.Open "Select stat_name from gravity order by region,province,municipal,barangay,stat_name", cnn, adOpenStatic

rstTriangulation.CursorLocation = adUseClient
rstTriangulation.Open "Select * from triangulation order by region,province,municipal,barangay,stat_name", cnn, adOpenStatic
  
SetToolBarStatus
SetToolBarStatusBM
SetToolBarStatusGravity
FillUpTriangulation


 
End Sub


Public Sub ZeroMode()
    Me.RaveSearch.Enabled = 0
    Me.RaveAdd.Enabled = 1
    Me.RaveEdit.Enabled = 0
    Me.RaveDelete.Enabled = 0
    Me.RaveSave.Enabled = 0
    Me.RaveCancel.Enabled = 0
    Me.RaveNext.Enabled = 0
    Me.RaveBack.Enabled = 0
    Me.RavePrint.Enabled = 0
    Me.RaveGIS.Enabled = 0
    Me.RaveInfo.Enabled = 0
    Me.RaveImages.Enabled = 0
    
End Sub

Private Sub BrowseMode()
    Me.RaveSearch.Enabled = 1
    Me.RaveAdd.Enabled = 1
    Me.RaveEdit.Enabled = 1
    Me.RaveDelete.Enabled = 1
    Me.RaveSave.Enabled = 0
    Me.RaveCancel.Enabled = 0
    Me.RaveNext.Enabled = 1
    Me.RaveBack.Enabled = 1
    Me.RavePrint.Enabled = 1
    Me.RaveGIS.Enabled = 1
    Me.RaveInfo.Enabled = 1
    Me.RaveImages.Enabled = 1
    
End Sub

Private Sub AddEditMode()
    Me.RaveSearch.Enabled = 0
    Me.RaveAdd.Enabled = 0
    Me.RaveEdit.Enabled = 0
    Me.RaveDelete.Enabled = 0
    Me.RaveSave.Enabled = 1
    Me.RaveCancel.Enabled = 1
    Me.RaveNext.Enabled = 0
    Me.RaveBack.Enabled = 0
    Me.RavePrint.Enabled = 0
    Me.RaveInfo.Enabled = 0
    Me.RaveImages.Enabled = 0
    Me.RaveGIS.Enabled = 0
   
    
End Sub


Private Sub AddEditModeGravity()
    Me.RaveSearchGravity.Enabled = 0
    Me.RaveAddGravity.Enabled = 0
    Me.RaveEditGravity.Enabled = 0
    Me.RaveDeleteGravity.Enabled = 0
    Me.RaveSaveGravity.Enabled = 1
    Me.RaveCancelGravity.Enabled = 1
    Me.RaveNextGravity.Enabled = 0
    Me.RaveBackGravity.Enabled = 0
    Me.RavePrintGravity.Enabled = 0
      
End Sub


Public Sub FirstMode()
    Me.RaveSearch.Enabled = 1
    Me.RaveAdd.Enabled = 1
    Me.RaveEdit.Enabled = 1
    Me.RaveDelete.Enabled = 1
    Me.RaveSave.Enabled = 0
    Me.RaveCancel.Enabled = 0
    Me.RaveNext.Enabled = 1
    Me.RaveBack.Enabled = 0
    Me.RavePrint.Enabled = 1
    Me.RaveGIS.Enabled = 1
    Me.RaveInfo.Enabled = 1
    Me.RaveImages.Enabled = 1
    
End Sub

Public Sub SingleMode()
    Me.RaveSearch.Enabled = 1
    Me.RaveAdd.Enabled = 1
    Me.RaveEdit.Enabled = 1
    Me.RaveDelete.Enabled = 1
    Me.RaveSave.Enabled = 0
    Me.RaveCancel.Enabled = 0
    Me.RaveNext.Enabled = 0
    Me.RaveBack.Enabled = 0
    Me.RavePrint.Enabled = 1
    Me.RaveGIS.Enabled = 1
    Me.RaveInfo.Enabled = 1
    Me.RaveImages.Enabled = 1
   
End Sub

Private Sub LastMode()
    Me.RaveSearch.Enabled = 1
    Me.RaveAdd.Enabled = 1
    Me.RaveEdit.Enabled = 1
    Me.RaveDelete.Enabled = 1
    Me.RaveSave.Enabled = 0
    Me.RaveCancel.Enabled = 0
    Me.RaveNext.Enabled = 0
    Me.RaveBack.Enabled = 1
    Me.RavePrint.Enabled = 1
    Me.RaveGIS.Enabled = 1
    Me.RaveInfo.Enabled = 1
    Me.RaveImages.Enabled = 1
  
End Sub

Public Sub FirstModeGravity()
    Me.RaveSearchGravity.Enabled = 1
    Me.RaveAddGravity.Enabled = 1
    Me.RaveEditGravity.Enabled = 1
    Me.RaveDeleteGravity.Enabled = 1
    Me.RaveSaveGravity.Enabled = 0
    Me.RaveCancelGravity.Enabled = 0
    Me.RaveNextGravity.Enabled = 1
    Me.RaveBackGravity.Enabled = 0
    Me.RavePrintGravity.Enabled = 1
   
    
End Sub

Public Sub SingleModeGravity()
    Me.RaveSearchGravity.Enabled = 1
    Me.RaveAddGravity.Enabled = 1
    Me.RaveEditGravity.Enabled = 1
    Me.RaveDeleteGravity.Enabled = 1
    Me.RaveSaveGravity.Enabled = 0
    Me.RaveCancelGravity.Enabled = 0
    Me.RaveNextGravity.Enabled = 0
    Me.RaveBackGravity.Enabled = 0
    Me.RavePrintGravity.Enabled = 1
       
   
End Sub

Private Sub LastModeGravity()
    Me.RaveSearchGravity.Enabled = 1
    Me.RaveAddGravity.Enabled = 1
    Me.RaveEditGravity.Enabled = 1
    Me.RaveDeleteGravity.Enabled = 1
    Me.RaveSaveGravity.Enabled = 0
    Me.RaveCancelGravity.Enabled = 0
    Me.RaveNextGravity.Enabled = 0
    Me.RaveBackGravity.Enabled = 1
    Me.RavePrintGravity.Enabled = 1
  
End Sub


Public Sub ZeroModeGravity()
    Me.RaveSearchGravity.Enabled = 0
    Me.RaveAddGravity.Enabled = 1
    Me.RaveEditGravity.Enabled = 0
    Me.RaveDeleteGravity.Enabled = 0
    Me.RaveSaveGravity.Enabled = 0
    Me.RaveCancelGravity.Enabled = 0
    Me.RaveNextGravity.Enabled = 0
    Me.RaveBackGravity.Enabled = 0
    Me.RavePrintGravity.Enabled = 0
   
    
End Sub

Private Sub BrowseModeGravity()
    Me.RaveSearchGravity.Enabled = 1
    Me.RaveAddGravity.Enabled = 1
    Me.RaveEditGravity.Enabled = 1
    Me.RaveDeleteGravity.Enabled = 1
    Me.RaveSaveGravity.Enabled = 0
    Me.RaveCancelGravity.Enabled = 0
    Me.RaveNextGravity.Enabled = 1
    Me.RaveBackGravity.Enabled = 1
    Me.RavePrintGravity.Enabled = 1
    
    
End Sub


Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Image2_Click()

End Sub

Private Sub ImageBorder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long

If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

End If
End Sub

Private Sub Labelx_Click(Index As Integer)

End Sub







Private Sub LstRegion_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub LstProvince_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub LstMunicipality_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub LstBarangay_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ImportExcel_Click()
'Me.DescriptionCommonDialog.Filter = "Microsoft Excel | *.xls"
'Me.DescriptionCommonDialog.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
'Me.DescriptionCommonDialog.ShowOpen
FrmImportDesc.Show 1
End Sub



Private Sub Label_Click(Index As Integer)
   
    If AddMode = True Then
       If Index = 7 Then
        Me.TxtName = PreviousRecord.stationName
       End If
       If Index = 13 Then
        Me.TxtRegion = PreviousRecord.Region
       End If
       If Index = 36 Then
        Me.TxtProvince = PreviousRecord.Province
       End If
       If Index = 4 Then
        Me.TxtMunicipality = PreviousRecord.Municipality
       End If
       If Index = 8 Then
        Me.TxtBarangay = PreviousRecord.Barangay
       End If
        If Index = 3 Then
        Me.TxtIsland = PreviousRecord.Island
       End If
        If Index = 0 Then
        Me.txtOrder = PreviousRecord.Order
       End If
       If Index = 15 Then
        Me.TxtEstablishedBy = PreviousRecord.SurveyedBy
       End If
    End If
    
    
End Sub

Private Sub LabelBuild_Click()
    
    MsgBox "Bug Fixes:" & vbCrLf & _
    "Benchmark query result. Fixed - Mar-04-2010" & vbCrLf & _
    "Importing without description - Mar-18-2010" & vbCrLf & _
    "Inclusion of Road maps - May-11-2010" & vbCrLf & _
    "Benckmarks Coordinates - May-11-2010" & vbCrLf & _
    "Field Number is removed - Aug-23-2010" & vbCrLf & _
    "Bar Coding - Oct-19-2010" & vbCrLf & _
    "Uploaded GCPs - Oct-19-2010" & vbCrLf & _
    "Bar Coding - Oct-19-2010" & vbCrLf & _
    "Added AdoptedBy field - Nov-02-2010" & vbCrLf & _
    "Barcode for benchmarks - Nov-02-2010" & vbCrLf & _
    "Fix error on Print where station name has a apostrophe - Mar-02-2010" & vbCrLf & _
    "Added paper size in printing of certificates - Mar-02-2010" & vbCrLf & _
    "Fix the minimize window issue - Mar-02-2010" & vbCrLf & _
    "Fix gcps Info - April-26-2011" & vbCrLf & _
    "Fix negative value for ellipsoidal height - April-26-2011" & vbCrLf & _
    "Add Conversion from PRS92 to UTM - Sept-2011" & vbCrLf & _
    "Add UTM to Certificate - Sept-2011" & vbCrLf
End Sub

Private Sub LabelLogout_Click()
    FrmLogin.Show 1
End Sub

Private Sub MouseHelper1_MouseWheel(ctrl As Variant, Direction As MBMouseHelper.mbDirectionConstants, Button As Long, Shift As Long, Cancel As Boolean)
Dim speed, i As Integer
speed = Abs(Direction) / 120

If Me.ActiveControl.name = Me.MyMap.name Then
       If Direction > 0 Then

          WheelZoomOut speed

       ElseIf Direction < 0 Then


           WheelZoomIn speed

       End If

End If
End Sub


Public Sub WheelZoomIn(ByVal speed As Integer)
Dim R
   Set R = MyMap.Extent
        R.ScaleRectangle 1 - (speed * 0.1)
        MyMap.Extent = R
   
End Sub

Public Sub WheelZoomOut(ByVal speed As Integer)
Dim R
   Set R = MyMap.Extent
        R.ScaleRectangle 1 + (speed * 0.1)
        MyMap.Extent = R
   
End Sub



Private Sub LstResult_DblClick()
    GotoRecord
End Sub





Private Sub MyMap_AfterTrackingLayerDraw(ByVal hDC As Stdole.OLE_HANDLE)
On Error GoTo Hell



Dim i As Long
Dim txt As New MapObjects.TextSymbol
Dim pnt As New MapObjects.Point

txt.Font = "tahoma"
txt.Font.SIZE = 6
 
    If MyMap.Extent.Height > 0 And MyMap.Extent.Height < 0.5 Then
            txt.Font.SIZE = 8
     End If


     If MyMap.Extent.Height > 0.5 And MyMap.Extent.Height < 1 Then
            txt.Font.SIZE = 6
     End If

     If MyMap.Extent.Height > 1 And MyMap.Extent.Height < 2 Then
            txt.Font.SIZE = 6
     End If

     If MyMap.Extent.Height > 2 And MyMap.Extent.Height < 3 Then
            txt.Font.SIZE = 6
     End If
     If MyMap.Extent.Height > 3 And MyMap.Extent.Height < 4 Then
            txt.Font.SIZE = 6
     End If

 

txt.HorizontalAlignment = moAlignLeft

'Display the names of the gcps
 If MyMap.Extent.Height <= 0.06 Then
For i = 1 To UBound(gEventList)
        If gEventOrder(i) = 1 Then
            txt.Color = MyMap.TrackingLayer.Symbol(1).Color
            ElseIf gEventOrder(i) = 2 Then
            txt.Color = MyMap.TrackingLayer.Symbol(2).Color
            ElseIf gEventOrder(i) = 3 Then
            txt.Color = MyMap.TrackingLayer.Symbol(3).Color
            Else
            txt.Color = RGB(102, 192, 253)
        End If

        pnt.x = gEventList(i).x
        pnt.y = gEventList(i).y

        If MyMap.Extent.Height < 0.6 Then
            MyMap.DrawText "  " & gEventsTag(i), pnt, txt
        End If
       
Next
 End If


'MyMap.Refresh

'End If

Exit Sub
Hell:

MsgBox Err.Description

End Sub

Private Sub MyMap_BeforeLayerDraw(ByVal Index As Integer, ByVal hDC As Stdole.OLE_HANDLE)





If Index = 0 Then
If CurrentSpatialQuery <> "" Then
 SearchMunicipality CurrentSpatialQuery, CurrentSpatialQueryField
End If
End If


     If MyMap.Extent.Height <= 1 Then
            MyMap.Layers(0).Visible = True
            Else
            MyMap.Layers(0).Visible = False
     End If
     
     If MyMap.Extent.Height < 0.6 And MyMap.Extent.Height > 0.02 Then
            MyLabelRender.AllowDuplicates = False
            MyLabelRender.Field = "TOWN"
            MyLabelRender.Symbol(0).Font.Bold = True
            MyLabelRender.Symbol(0).Font.SIZE = 12
      ElseIf MyMap.Extent.Height <= 0.02 Then
            
        
            MyLabelRender.Symbol(0).Font.Bold = False
            MyLabelRender.Symbol(0).Font.SIZE = 8

           MyLabelRender.Symbol(0).Color = vbWhite
           MyLabelRender.AllowDuplicates = True
           MyLabelRender.Field = "Name"
   Else
   MyLabelRender.AllowDuplicates = False
   MyLabelRender.Field = "PROVINCE"
   MyLabelRender.Symbol(0).Font.Bold = True
            MyLabelRender.Symbol(0).Font.SIZE = 12
End If


    Set MyMap.Layers(0).Renderer = MyLabelRender
    
   



End Sub



Private Sub MyMap_BeforeTrackingLayerDraw(ByVal hDC As Stdole.OLE_HANDLE)
'Hide GCPS
     If MyMap.Extent.Height <= 0.5 Then
            MyMap.TrackingLayer.Visible = True
            Else
             MyMap.TrackingLayer.Visible = False
     End If

End Sub



Private Sub MyMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'mousedown = True
'mousemove = False

   Dim pnt As New MapObjects.Point
    pnt.x = MyMap.ToMapPoint(x, y).x
    pnt.y = MyMap.ToMapPoint(x, y).y
    Dim Index As Long
    
     Index = FindGeoEvent(pnt)
    
  
            If Index <> -1 And MyMap.Extent.Height < 0.5 Then
                
                CurrentStationToIdentify = gEventsTag(Index + 1)
                FrmIdentify.Show 1
                
                
               Else
                Me.MyMap.MousePointer = moPan
               Me.MyMap.Pan
                Me.MyMap.MousePointer = moArrow
               
End If
 

'If Me.RaveZoomIn.Value = True Then
'    DoZoom
'End If

'If Me.RavePan.Value = True Then
'    Me.MyMap.Pan
'End If

'If Me.RaveNearest.Value = True Then
'    FrmCurrentLocation.Show 1
'End If



End Sub



Private Sub MyMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim D As Integer
Dim M As Double
Dim s As Double



D = Int(MyMap.ToMapPoint(x, y).x)
M = (MyMap.ToMapPoint(x, y).x - D) * 60
s = Round((M - Int(M)) * 60, 2)
M = Int(M)

Me.StatusBar1.Panels(5).Text = "Longitude: " & D & "° " & M & "' " & s & "''"

D = Int(MyMap.ToMapPoint(x, y).y)
M = (MyMap.ToMapPoint(x, y).y - D) * 60
s = Round((M - Int(M)) * 60, 2)
M = Int(M)


Me.StatusBar1.Panels(4).Text = "Latitude: " & D & "° " & M & "' " & s & "''"


'If mousedown = True Then
'mousemove = True
'Me.MyMap.Pan
'End If
End Sub





Public Sub IdentifyGCP(ByVal x As Long, ByVal y As Long)
    
    Dim pnt As New MapObjects.Point
    pnt.x = MyMap.ToMapPoint(x, y).x
    pnt.y = MyMap.ToMapPoint(x, y).y
    MsgBox FindGeoEvent(pnt)
    
End Sub

Private Sub MyMap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Me.MyMap.MousePointer = moArrow
End Sub

Private Sub MyMap2_AfterTrackingLayerDraw(ByVal hDC As Stdole.OLE_HANDLE)
On Error GoTo Hell

If MyMap2.Extent.Height < 0.004 Then
    Dim i As Long
    Dim txtsym As New MapObjects.TextSymbol
    Dim pnt As New MapObjects.Point
    
    txtsym.Color = vbWhite
    txtsym.HorizontalAlignment = moAlignLeft
    txtsym.VerticalAlignment = moAlignTop
        
    
    For i = 1 To UBound(gEventBm)
        pnt.x = MyMap2.TrackingLayer.Event(i - 1).x
        pnt.y = MyMap2.TrackingLayer.Event(i - 1).y
        MyMap2.DrawText gEventBm(i), pnt, txtsym
    Next
End If
Exit Sub
Hell:
End Sub

Private Sub MyMap2_BeforeLayerDraw(ByVal Index As Integer, ByVal hDC As Stdole.OLE_HANDLE)
If MyMap2.Extent.Height < 0.02 Then
MyMap2.Layers(1).Visible = True
Else
MyMap2.Layers(1).Visible = False
End If


If MyMap2.Extent.Height < 0.6 And MyMap2.Extent.Height > 0.02 Then
            
            
            MyLabelRender2.AllowDuplicates = False
            MyLabelRender2.Field = "TOWN"
            MyLabelRender2.Symbol(0).Color = vbWhite
            MyLabelRender2.Symbol(0).Font.Bold = True
            MyLabelRender2.Symbol(0).Font.SIZE = 12
            Set MyMap2.Layers(0).Renderer = MyLabelRender2
           
   ElseIf MyMap2.Extent.Height <= 0.02 And MyMap2.Extent.Height > 0.002 Then
            
           MyLabelRender2.AllowDuplicates = True
           MyLabelRender2.Symbol(0).Font.Bold = False
           MyLabelRender2.Symbol(0).Font.SIZE = 8
           MyLabelRender2.Symbol(0).Color = vbWhite
           MyLabelRender2.AllowDuplicates = True
           MyLabelRender2.Field = "Name"
           MyLabelRender2.Symbol(0).HorizontalAlignment = moAlignLeft
           MyLabelRender2.Symbol(0).VerticalAlignment = moAlignTop
           Set MyMap2.Layers(2).Renderer = MyLabelRender2

           
           
                      
           
    ElseIf MyMap2.Extent.Height <= 0.002 Then
            
           MyLabelRender2.AllowDuplicates = True
           MyLabelRender2.Symbol(0).Font.Bold = False
           MyLabelRender2.Symbol(0).Font.SIZE = 10
           MyLabelRender2.Symbol(0).Color = vbWhite
           MyLabelRender2.AllowDuplicates = True
           MyLabelRender2.Field = "Namex"
           Set MyMap2.Layers(1).Renderer = MyLabelRender2
          
   ElseIf MyMap2.Extent.Height < 4.5 Then
   
            MyLabelRender2.AllowDuplicates = False
            MyLabelRender2.Field = "PROVINCE"
            MyLabelRender2.Symbol(0).Color = vbWhite
            MyLabelRender2.Symbol(0).Font.Bold = True
            MyLabelRender2.Symbol(0).Font.SIZE = 12
            Set MyMap2.Layers(0).Renderer = MyLabelRender2
   
   ElseIf MyMap2.Extent.Height > 4.5 Then
   
            Set MyMap2.Layers(0).Renderer = Nothing
 
End If


    
   
End Sub

Private Sub MyMap2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.RaveZoomInBM.Value = True Then
    DoZoom2
End If

If Me.RavePanBM.Value = True Then
    Me.MyMap2.Pan
End If

If Me.RaveBM.Value = True Then
   Call AssignCoordinates(x, y)
End If


End Sub

Private Sub AssignCoordinates(x As Single, y As Single)
    Dim pnt As New MapObjects.Point
    
    If MsgBox("Are you sure you want to update the coordinates of this benchmark?", vbYesNo, "Assign Coordinates") = vbNo Then
        Exit Sub
    End If
    
    pnt.x = MyMap2.ToMapPoint(x, y).x
    pnt.y = MyMap2.ToMapPoint(x, y).y
    cnn.Execute "Update benchmarks set Longitude=" & pnt.x & "," & "Latitude=" & pnt.y & " Where ucode=" & rstBenchmarks!ucode
    Me.TextBoxLongitude = pnt.x
    Me.TextBoxLatitude = pnt.y
    
    PlotBM
    MsgBox "Coordinates Updated!", vbInformation, "Update"
End Sub


Private Sub MyMap2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim D As Integer
Dim M As Double
Dim s As Double


D = Int(MyMap2.ToMapPoint(x, y).x)
M = (MyMap2.ToMapPoint(x, y).x - D) * 60
s = Round((M - Int(M)) * 60, 2)
M = Int(M)

Me.StatusBar2.Panels(5).Text = "Longitude: " & D & "° " & M & "' " & s & "''"

D = Int(MyMap2.ToMapPoint(x, y).y)
M = (MyMap2.ToMapPoint(x, y).y - D) * 60
s = Round((M - Int(M)) * 60, 2)
M = Int(M)


Me.StatusBar2.Panels(4).Text = "Latitude: " & D & "° " & M & "' " & s & "''"
End Sub

Private Sub optBMs_Click()
    If optBMs.Value = True Then
        BMQuery = 1
    Else
        BMQuery = 0
    End If
End Sub

Private Sub optGCPs_Click()
    If optGCPs.Value = True Then
        BMQuery = 0
    End If
End Sub

Private Sub RaveAdd_Click()
    EditMode = False
    AddMode = True
    BlankForm
    
    AddEditMode
    EnableFields
    Me.TxtDateEntry = Format(Date, "mm-dd-yyyy")
    Me.txtAuthority = "NAMRIA"
    Me.txtMarkPurpose = "Geodetic"
    Me.TxtFixingMethod = "GPS"
    Me.txtMarkStatus = "Existing"
    Me.txtMarkType = "Copper Nail"
    Me.TxtName.SetFocus
    'FrmLocation.Show 1
End Sub


Private Sub RaveAddBM_Click()
    AddModeBM = True
    BlankFormBM
    AddEditModeBM
    EnableBenchmarksFields True
    
    Me.TxtBMDateOfEntry = Format(Date, "mm-dd-yyyy")
    Me.TxtBMAuthority = "NAMRIA"
    Me.TxtEName.SetFocus
End Sub

Private Sub RaveAddGravity_Click()
    EditModegravity = False
    BlankFormGravity
    EnableFieldsGravity
    AddEditModeGravity
    Me.TextBoxGravityName.SetFocus
End Sub

Private Sub RaveAddQuery_Click()
AddQuery = True
FrmQuery.Show 1
End Sub

Private Sub RaveBack_Click()
    rstRecords.MovePrevious
    SetToolBarStatus
End Sub



Private Sub RaveBackBM_Click()
    rstBenchmarks.MovePrevious
    SetToolBarStatusBM
End Sub



Private Sub RaveBackGravity_Click()
    rstGravity.MovePrevious
    SetToolBarStatusGravity
End Sub

Private Sub RaveBM_Click()
If Me.RaveBM.Value = True Then
    Me.MyMap2.MousePointer = moCross
    Me.RavePanBM.Value = False
    Me.RaveZoomInBM.Value = False
     Me.StatusBar2.Panels(2).Text = "Tool: Assign Coordinates"
    Else
    Me.MyMap2.MousePointer = moArrow
     Me.StatusBar2.Panels(2).Text = "Tool: None"
 End If
End Sub

Private Sub RaveButtons15_Click()
FrmUtilities.Show 1
End Sub

Private Sub RaveButtons16_Click()
FrmPSGCLibrary.Show 1
End Sub

Private Sub RaveButtons17_Click()
    AddEditModeBM
    
    regionTextbox = FrmGCPDS.TxtERegion.name
    provinceTextbox = FrmGCPDS.TxtEProvince.name
    municipalityTextbox = FrmGCPDS.TxtEMunicipality.name
    barangayTextbox = FrmGCPDS.TxtEBarangay.name
    FrmLocation.Show 1
    
End Sub

Private Sub RaveButtons2_Click()
    FrmDetectLocation.Show 1
End Sub



Private Sub RaveButtons3_Click()
FrmTriangulationPrintOut.Show
End Sub

Private Sub RaveButtons4_Click()

    rstTriangulation.MovePrevious
    
If rstTriangulation.BOF <> True Then
    FillUpTriangulation
    Else
    rstTriangulation.MoveFirst
End If
End Sub

Private Sub RaveButtons5_Click()
rstTriangulation.MoveNext
If rstTriangulation.EOF <> True Then
    FillUpTriangulation
    Else
    rstTriangulation.MoveLast
End If
End Sub

Private Sub RaveButtons6_Click()
FrmSearchTriangulation.Show 1
End Sub

Private Sub RaveButtons7_Click()
    FrmInventory.Show 1
End Sub

Private Sub RaveButtonsExport_Click()
    FrmExport.Show 1
End Sub

Private Sub RaveButtonsImport_Click()
'On Error GoTo Hell:

 Me.CommonDialogImport.InitDir = App.Path
 Me.CommonDialogImport.Filter = "XML File|*.XML"
 Me.CommonDialogImport.ShowOpen
 
 Dim rst As New ADODB.Recordset
 Dim rstgeoprov As New ADODB.Recordset
 
 rst.Open CommonDialogImport.filename, , adOpenStatic, adLockOptimistic
 
Dim i As Integer


Me.ProgressBar1.Visible = True
Me.ProgressBar1.Value = 0
Me.LabelProgress.Visible = True
Me.LabelProgress = "Importing"


For i = 1 To rst.RecordCount

DoEvents
Me.ProgressBar1.Value = i * (100 / rst.RecordCount)
 

        'cnn.Execute "delete from geoprov where stat_name='" & Replace(rst!Stat_Name, "'", "''") & "'"
   
 If IsDuplicateStation(rst!Stat_Name) = False Then
    




    Dim strSQl As String
    Dim adoCommand As ADODB.Command

    Set adoCommand = New ADODB.Command
    strSQl = "INSERT INTO geoprov (stat_name,stat_new,region,province,municipal,barangay,island,h_order,h_ref,d_lat,m_lat,s_lat,d_long,m_long,s_long,ell_hgt,wgs84ND,wgs84NM,wgs84NS,wgs84ED,wgs84EM,wgs84ES,ellipz,h_date_ety,mark_const,mark_type,mark_stat,hor_authty,authority,date_est,h_date_com,date_las_r,h_fix,status,image,northing,easting,zone,description) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

    With adoCommand
        .ActiveConnection = cnn
        .CommandType = adCmdText
        .CommandText = strSQl
        .Prepared = True

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("stat_name").DefinedSize, rst("stat_name"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("stat_new").DefinedSize, rst("stat_new"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("region").DefinedSize, rst("region"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("province").DefinedSize, rst("province"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("municipal").DefinedSize, rst("municipal"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("barangay").DefinedSize, rst("barangay"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("island").DefinedSize, rst("island"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("h_order").DefinedSize, rst("h_order"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("h_ref").DefinedSize, rst("h_ref"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("d_lat").DefinedSize, rst("d_lat"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("m_lat").DefinedSize, rst("m_lat"))
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, rst("s_lat").DefinedSize, rst("s_lat"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("d_long").DefinedSize, rst("d_long"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("m_long").DefinedSize, rst("m_long"))
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, rst("s_long").DefinedSize, rst("s_long"))
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, rst("ell_hgt").DefinedSize, rst("ell_hgt"))
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("wgs84ND").DefinedSize, rst("wgs84ND"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("wgs84NM").DefinedSize, rst("wgs84NM"))
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, rst("wgs84NS").DefinedSize, rst("wgs84NS"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("wgs84ED").DefinedSize, rst("wgs84ED"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("wgs84EM").DefinedSize, rst("wgs84EM"))
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, rst("wgs84ES").DefinedSize, rst("wgs84ES"))
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, rst("ellipz").DefinedSize, rst("ellipz"))
        .Parameters.Append .CreateParameter(, adDBDate, adParamInput, rst("h_date_ety").DefinedSize, rst("h_date_ety"))
         
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("mark_const").DefinedSize, rst("mark_const"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("mark_type").DefinedSize, rst("mark_type"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("mark_stat").DefinedSize, rst("mark_stat"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("hor_authty").DefinedSize, rst("hor_authty"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("authority").DefinedSize, rst("authority"))
        
        .Parameters.Append .CreateParameter(, adDBDate, adParamInput, rst("date_est").DefinedSize, rst("date_est"))
        .Parameters.Append .CreateParameter(, adDBDate, adParamInput, rst("h_date_com").DefinedSize, rst("h_date_com"))
        .Parameters.Append .CreateParameter(, adDBDate, adParamInput, rst("date_las_r").DefinedSize, rst("date_las_r"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("h_fix").DefinedSize, rst("h_fix"))
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, rst("status").DefinedSize, rst("status"))
        
        .Parameters.Append .CreateParameter(, adLongVarBinary, adParamInput, rst("image").DefinedSize, rst("image"))
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, rst("northing").DefinedSize, rst("northing"))
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, rst("easting").DefinedSize, rst("easting"))
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, rst("zone").DefinedSize, rst("zone"))
        .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, rst("description").DefinedSize, rst("description"))
        .Execute , , adCmdText + adExecuteNoRecords
    End With
End If
    rst.MoveNext
Next

Me.ProgressBar1.Visible = False
Me.ProgressBar1.Value = 0
Me.LabelProgress.Visible = False
Me.LabelProgress = ""
    
MsgBox "Done"

rstRecords.Requery
SetToolBarStatus
 
rst.Close
Set rst = Nothing




Exit Sub
Hell:
End Sub

Private Sub RaveButtonsUnits_Click()
    If RaveButtonsUnits.Caption = "m" Then
     RaveButtonsUnits.Caption = "ft"
     Else
     RaveButtonsUnits.Caption = "m"
    End If
End Sub

Private Sub RaveButtonsUploaded_Click()
FrmUploadedGCP.Show 1
End Sub

Private Sub RaveCancel_Click()
    SetToolBarStatus
    DisableFields
    AddMode = False
    EditMode = False
    
    
  
End Sub

Private Sub RaveCancelBM_Click()
    SetToolBarStatusBM
    EnableBenchmarksFields False
    AddMode = False
    'EditMode = False
End Sub

Private Sub RaveCancelGravity_Click()
    
     SetToolBarStatusGravity
     DisableFieldsGravity
     
End Sub

Private Sub RaveDelete_Click()

If AccessLevel = 2 Then
   FrmAdmin.Show 1
   If TemporaryPass = False Then
        Exit Sub
   End If
End If

DataType = "GCP"

If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete") = vbYes Then
Dim bookmark As String
    
    cnn.Execute "Insert into deleted (stat_name,island,region,province,municipal,barangay,h_order,d_lat,m_lat,s_lat,d_long,m_long,s_long,ell_hgt,wgs84ED,wgs84EM,wgs84ES,wgs84ND,wgs84NM,wgs84NS,ellipz,h_date_ety,h_date_com,date_las_r,date_est,mark_const,mark_type,mark_stat,h_fix,hor_authty,authority,h_ref,Date_Deleted,Deleted_By)" _
             & " values('" & Replace(FrmGCPDS.TxtName, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtIsland, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtRegion, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtProvince, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtMunicipality, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtBarangay, "'", "''") & "'" & "," & IIf(Trim(FrmGCPDS.txtOrder) = "", "Null", Order(FrmGCPDS.txtOrder.ListIndex + 1)) & "," & IIf(Trim(FrmGCPDS.TxtLatD) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtLatD), FrmGCPDS.TxtLatD, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtLatM) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtLatM), FrmGCPDS.TxtLatM, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtLatS) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtLatS), FrmGCPDS.TxtLatS, "Null")) _
             & "," & IIf(Trim(FrmGCPDS.TxtLongD) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtLongD), FrmGCPDS.TxtLongD, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtLongM) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtLongM), FrmGCPDS.TxtLongM, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtLongS) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtLongS), FrmGCPDS.TxtLongS, "Null")) & "," & IIf(Trim(FrmGCPDS.txtEllipsoidalH) = "", "Null", IIf(IsNumeric(FrmGCPDS.txtEllipsoidalH), FrmGCPDS.txtEllipsoidalH, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtDLat) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtDLat), FrmGCPDS.TxtDLat, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtMLat) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtMLat), FrmGCPDS.TxtMLat, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtSLat) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtSLat), FrmGCPDS.TxtSLat, "Null")) _
             & "," & IIf(Trim(FrmGCPDS.TxtDLong) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtDLong), FrmGCPDS.TxtDLong, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtMLong) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtMLong), FrmGCPDS.TxtMLong, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtSLong) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtSLong), FrmGCPDS.TxtSLong, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtEllipsoidalH2) = "", "Null", IIf(IsNumeric(FrmGCPDS.TxtEllipsoidalH2), FrmGCPDS.TxtEllipsoidalH2, "Null")) & "," & IIf(Trim(FrmGCPDS.TxtDateEntry) = "", "Null", "'" & FrmGCPDS.TxtDateEntry & "'") & "," & IIf(Trim(FrmGCPDS.TxtDateComputed) = "", "Null", "'" & FrmGCPDS.TxtDateComputed & "'") & "," & IIf(Trim(FrmGCPDS.txtLastRecover) = "", "Null", "'" & FrmGCPDS.txtLastRecover & "'") & "," & IIf(Trim(FrmGCPDS.txtEstablished) = "", "Null", "'" & FrmGCPDS.txtEstablished & "'") & "," & IIf(Trim(FrmGCPDS.txtMarkPurpose) = "", "Null", MarkPurpose(FrmGCPDS.txtMarkPurpose.ListIndex + 1)) _
             & "," & IIf(Trim(FrmGCPDS.txtMarkType) = "", "Null", MarkType(FrmGCPDS.txtMarkType.ListIndex + 1)) & "," & IIf(Trim(FrmGCPDS.txtMarkStatus) = "", "Null", MarkStatus(FrmGCPDS.txtMarkStatus.ListIndex + 1)) & "," & IIf(Trim(FrmGCPDS.TxtFixingMethod) = "", "Null", HorizontalFixingMethod(FrmGCPDS.TxtFixingMethod.ListIndex + 1)) & "," & "'" & Replace(FrmGCPDS.TxtEstablishedBy, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.txtAuthority, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtRef, "'", "''") & "'" & "," & "'" & Now & "'" & "," & "'" & Replace(Encoder, "'", "''") & "'" & ")"
    
    cnn.Execute "Delete from geoprov where stat_name='" & Replace(FrmGCPDS.TxtName, "'", "''") & "'"
    
    rstRecords.MoveNext           'Move to the next record
    
    If rstRecords.EOF Then        'If current record is End Of File , then move to from EOF to the deleted record, then move again to the record prior the the deleted record (Record2)
        
        rstRecords.MovePrevious   'Diagram: BOF  Record1  Record2  DeletedRecord  EOF
        rstRecords.MovePrevious
        
            If rstRecords.BOF Then         'If after the move, the pointer points to BOF then there is no records left
                                   
                        rstRecords.Requery
                        BlankForm
                        ZeroMode
                        FrmGCPDS.FormCaption.Caption = "No Records"
                        Exit Sub
                    
                    Else
                        
                        bookmark = rstRecords!Stat_Name
                        rstRecords.Requery
                        rstRecords.Find "stat_name='" & bookmark & "'"
                    
        
             End If
            
            
            
        Else                      'IF NOT EOF. Bookmark the station name. Requery the recordset.  Go to bookmark
        
        bookmark = rstRecords!Stat_Name
        rstRecords.Requery
        rstRecords.Find "stat_name='" & bookmark & "'"
            
            
    End If
    
    
   
    
    
      
      FrmDelete.Show 1
     SetToolBarStatus
    
    
   
    
 

    
       FrmGCPDS.FormCaption.Caption = Format(rstRecords.AbsolutePosition, "#,##0") & " of " & Format(rstRecords.RecordCount, "#,##0") & " Records"
       Station = rstRecords("Stat_name").Value
       
     
        
End If

End Sub



Private Sub RaveDeleteBM_Click()


If AccessLevel = 2 Then
   FrmAdmin.Show 1
   If TemporaryPass = False Then
        Exit Sub
   End If
End If


DataType = "Benchmarks"

If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete") = vbYes Then
Dim bookmark As String
    
     
    cnn.Execute "Delete from benchmarks where ucode=" & rstBenchmarks!ucode
    
    rstBenchmarks.MoveNext           'Move to the next record
    
    If rstBenchmarks.EOF Then        'If current record is End Of File , then move to from EOF to the deleted record, then move again to the record prior the the deleted record (Record2)
        
        rstBenchmarks.MovePrevious   'Diagram: BOF  Record1  Record2  DeletedRecord  EOF
        rstBenchmarks.MovePrevious
        
            If rstBenchmarks.BOF Then         'If after the move, the pointer points to BOF then there is no records left
                                   
                        rstBenchmarks.Requery
                        BlankFormBM
                        ZeroModeBM
                        FrmGCPDS.FormCaption2.Caption = "No Records"
                        Exit Sub
                    
                    Else
                        
                        bookmark = rstBenchmarks!ucode
                        rstBenchmarks.Requery
                        rstBenchmarks.Find "ucode=" & bookmark
                    
        
             End If
            
            
            
        Else                      'IF NOT EOF. Bookmark the station name. Requery the recordset.  Go to bookmark
        
        bookmark = rstBenchmarks!ucode
        rstBenchmarks.Requery
        rstBenchmarks.Find "ucode=" & bookmark
            
            
    End If
    
    
   
    
    
      
      FrmDelete.Show 1
     SetToolBarStatusBM
    
    
   
    
 

    
       FrmGCPDS.FormCaption2.Caption = Format(rstBenchmarks.AbsolutePosition, "#,##0") & " of " & Format(rstBenchmarks.RecordCount, "#,##0") & " Records"
      
       
     
        
End If
       
       
   
       
End Sub

Private Sub RaveDeletedGCPs_Click()
FrmDeletedGCPs.Show 1
End Sub





Private Sub RaveDeleteProvince_Click()

End Sub

Private Sub RaveDeletemunicipality_Click()

End Sub

Private Sub RaveDeleteDuplicateBMs_Click()
'Dim rst As New ADODB.Recordset
'rst.Open "Select * from benchmarks", cnn, adOpenStatic
'
'Dim i As Long
'
'Me.ProgressBar1.Visible = True
'Me.ProgressBar1.Value = 0
'Me.LabelProgress.Visible = True
'Me.LabelProgress = "Deleting Dup BMs"
'
'For i = 1 To rst.RecordCount
'    Me.ProgressBar1.Value = i * (100 / rst.RecordCount)
'    DoEvents
'    rs
'
'    cnn.Execute "delete from benchmarks where stat_name='" & Replace(rst!Stat_Name, "'", "''") & "' AND Region = '" & rst!Region & "' AND Province ='" & rst!Province & "' AND municipal='" & rst!Municipal & "' AND Barangay='" & rst!Barangay & "' AND Elevation =" & IIf(IsNull(rst!Elevation) = True, 0, rst!Elevation) & " AND ucode<>" & rst!ucode
'    rst.MoveNext
'Next
'
'Me.ProgressBar1.Visible = False
'Me.ProgressBar1.Value = 0
'Me.LabelProgress.Visible = False
'Me.LabelProgress = ""


End Sub

Private Sub RaveDeleteGravity_Click()

If AccessLevel = 2 Then
   FrmAdmin.Show 1
   If TemporaryPass = False Then
        Exit Sub
   End If
End If

DataType = "Gravity"

If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete") = vbYes Then
Dim bookmark As String
    
     
    cnn.Execute "Delete from gravity where stat_name='" & rstGravity!Stat_Name & "'"
    
    rstGravity.MoveNext           'Move to the next record
    
    If rstGravity.EOF Then        'If current record is End Of File , then move to from EOF to the deleted record, then move again to the record prior the the deleted record (Record2)
        
        rstGravity.MovePrevious   'Diagram: BOF  Record1  Record2  DeletedRecord  EOF
        rstGravity.MovePrevious
        
            If rstGravity.BOF Then         'If after the move, the pointer points to BOF then there is no records left
                                   
                        rstGravity.Requery
                        BlankFormGravity
                        ZeroModeGravity
                        'FrmGCPDS.FormCaption2.Caption = "No Records"
                        Exit Sub
                    
                    Else
                        
                        bookmark = rstGravity!Stat_Name
                        rstGravity.Requery
                        rstGravity.Find "stat_name='" & bookmark & "'"
                    
        
             End If
            
            
            
        Else                      'IF NOT EOF. Bookmark the station name. Requery the recordset.  Go to bookmark
        
        bookmark = rstGravity!Stat_Name
        rstGravity.Requery
        rstGravity.Find "stat_name='" & bookmark & "'"
            
            
    End If
    
    
   
    
    
      
      FrmDelete.Show 1
     SetToolBarStatusGravity
    
    
   
    
 

    
      ' FrmGCPDS.FormCaption2.Caption = Format(rstGravity.AbsolutePosition, "#,##0") & " of " & Format(rstGravity.RecordCount, "#,##0") & " Records"
      
       
     
        
End If
End Sub

Private Sub RaveDeleteQuery_Click()
If Me.LstConditions.ListItems.Count > 0 Then

RstQuery.Close
Me.LstConditions.ListItems.Clear
Me.LstResult.ListItems.Clear
Me.ResultLabel = ""
Me.PageCounterLabel = ""
Me.RavePreviousPage.Enabled = False
Me.RaveNextPage.Enabled = False
strcondition = ""

If BMQuery = 0 Then
    SetToolBarStatus
Else
    SetToolBarStatusBM
End If

End If
End Sub

Private Sub RaveDMS2Degree_Click()
    
    Call DMStoDegree
    
End Sub

Private Sub RaveDup_Click()
FrmDup.Show 1
End Sub

Private Sub RaveEdit_Click()

If AccessLevel = 2 Then
   FrmAdmin.Show 1
   If TemporaryPass = False Then
        Exit Sub
   End If
End If


EditMode = True
AddMode = False
AddEditMode
EnableFields


Me.TxtName.SetFocus

If Trim(Me.TxtNorthing) <> "" Then
   Me.TxtNorthing = Format(Me.TxtNorthing, "#.####")
End If
If Trim(Me.TxtEasting) <> "" Then
   Me.TxtEasting = Format(Me.TxtEasting, "#.####")
End If
End Sub

Private Sub RaveEditProvince_Click()

End Sub

Private Sub RaveEditmunicipality_Click()

End Sub

Private Sub RaveEditBrgy_Click()

End Sub

Private Sub RaveEditBM_Click()
If AccessLevel = 2 Then
   FrmAdmin.Show 1
   If TemporaryPass = False Then
        Exit Sub
   End If
End If

AddModeBM = False
AddEditModeBM
Me.TxtEName.SetFocus
EnableBenchmarksFields True
End Sub

Private Sub RaveEditGravity_Click()
If AccessLevel = 2 Then
   FrmAdmin.Show 1
   If TemporaryPass = False Then
        Exit Sub
   End If
End If
 
EditModegravity = True
AddEditModeGravity
EnableFieldsGravity
 
    Me.TextBoxGravityLatitude.Text = Replace(Replace(Replace(Me.TextBoxGravityLatitude, "°", ""), "'", ""), Chr(34), "")
    Me.TextBoxGravityLongitude.Text = Replace(Replace(Replace(Me.TextBoxGravityLongitude, "°", ""), "'", ""), Chr(34), "")
   If Trim(Me.TextBoxGravityElevation.Text) <> "" Then
        Me.TextBoxGravityElevation.Text = Val(Me.TextBoxGravityElevation.Text)
   End If


End Sub

Private Sub RaveEditQuery_Click()
If Me.LstConditions.ListItems.Count > 0 Then
    AddQuery = False
    FrmQuery.Show 1
End If
End Sub




Private Sub RaveExtent_Click()
    
Set MyMap.Layers(0).Renderer = Nothing
Me.MyMap.Extent = Me.MyMap.FullExtent
Me.StatusBar1.Panels(3).Text = "Zoom: " & Round(MyMap.Extent.Height, 2)


End Sub

Private Sub RaveExtentBM_Click()
Set MyMap2.Layers(0).Renderer = Nothing
Set MyMap2.Layers(1).Renderer = Nothing
Set MyMap2.Layers(2).Renderer = Nothing
Me.MyMap2.Extent = Me.MyMap.FullExtent
End Sub

Private Sub RaveFilter_Click()
    GotoRecord
End Sub

Private Sub GotoRecord()
    If Me.LstResult.ListItems.Count = 0 Then
    Exit Sub
End If
    

    If BMQuery = 0 Then
       
        If rstRecords.RecordCount > 0 Then
            rstRecords.MoveFirst
            rstRecords.Find "stat_name='" & Replace(Me.LstResult.SelectedItem.Text, "'", "''") & "'"
            Else
            Me.FormCaption = "No Records."
        End If
    Else
        
        
        If rstBenchmarks.RecordCount > 0 Then
            rstBenchmarks.MoveFirst
            rstBenchmarks.Find "ucode=" & Me.LstResult.SelectedItem.SubItems(1)
            Else
            Me.FormCaption = "No Records."
        End If
    End If
    
  

    If BMQuery = 0 Then
        
        FillUp
        SetToolBarStatus
        DisableFields
        Me.SSTab1.TabVisible(0) = True
        Me.SSTab1.Tab = 0
        Me.SSTab1.TabVisible(0) = False
        RaveTab(0).Value = True
        RaveTab(1).Value = False
        RaveTab(2).Value = False
    Else
        FillUpBenchmarks
        SetToolBarStatusBM
        DisableFields
        Me.SSTab1.TabVisible(5) = True
        Me.SSTab1.Tab = 5
        Me.SSTab1.TabVisible(5) = False
        RaveTab(0).Value = False
        RaveTab(2).Value = False
    End If
        Exit Sub
Hell:
End Sub

Private Sub RaveGIS_Click()
        
        
        If Trim(TxtDLat) = "" Then
            MsgBox "No valid WGS84 coordinates."
            Exit Sub
        End If
        
        Me.SSTab1.TabVisible(2) = True
        Me.SSTab1.Tab = 2
        Me.SSTab1.TabVisible(2) = False
        If rstRecords.RecordCount > 0 Then
        PlotGCP
        CenterView
        End If

   MapType = "GCP"
End Sub

Private Sub RaveGoto_Click()
FrmMarkers.Show 1
End Sub



Private Sub RaveImages_Click()
FrmImages.Show 1
End Sub

Private Sub RaveInfo_Click()
    FrmDescription.Show 1
End Sub

Private Sub RaveLongLat_Click()
FrmLatLong.Show 1
End Sub

Private Sub RaveInventory_Click()
    FrmPrintInventory.Show 1
End Sub



Private Sub RaveLogin_Click()
FrmRequestingParty.Show 1
End Sub



Private Sub RaveLocation_Click()
    regionTextbox = Me.TextBoxGravityRegion.name
    provinceTextbox = Me.TextBoxGravityProvince.name
    municipalityTextbox = Me.TextBoxGravityMunicipality.name
    barangayTextbox = Me.TextBoxGravityBarangay.name
    FrmLocation.Show 1
End Sub

Private Sub RaveMapBM_Click()
        Me.SSTab1.TabVisible(6) = True
        Me.SSTab1.Tab = 6
        Me.SSTab1.TabVisible(6) = False '
       
        PlotBM
        CenterViewBM
End Sub

Private Sub RaveMarkPurpose_Click()
    FrmLibraryMarkPurpose.Show 1
End Sub

Private Sub RaveMarkStatus_Click()
    FrmLibraryStatus.Show 1
End Sub

Private Sub RaveMarkType_Click()
    FrmLibrary.Show 1
End Sub



Private Sub RaveNext_Click()
    rstRecords.MoveNext
    SetToolBarStatus
End Sub

Private Sub RaveNextBM_Click()
    rstBenchmarks.MoveNext
    SetToolBarStatusBM
End Sub

Private Sub RaveNextGravity_Click()
    rstGravity.MoveNext
    SetToolBarStatusGravity
End Sub

Private Sub RaveNextPage_Click()


Dim i As Integer
Dim varlist

PageCounter = PageCounter + 1
Me.RavePreviousPage.Enabled = True
RstQuery.AbsolutePage = PageCounter
Me.PageCounterLabel = "Page " & RstQuery.AbsolutePage & " of " & RstQuery.PageCount

Me.LstResult.ListItems.Clear

For i = 1 To 10
Set varlist = Me.LstResult.ListItems.Add
    varlist.Text = IIf(IsNull(RstQuery("Stat_name")), "", RstQuery("Stat_name"))
    If Me.optBMs.Value = True Then
        varlist.SubItems(1) = RstQuery!ucode
    End If
    
    RstQuery.MoveNext
    If RstQuery.EOF Then
        Me.RaveNextPage.Enabled = False
        Me.RavePreviousPage.Enabled = True
        Exit For
    End If
Next


End Sub



Private Sub RavePanBM_Click()
If Me.RavePanBM.Value = True Then
    Me.MyMap2.MousePointer = moPan
    Me.RaveZoomInBM.Value = False
    Me.RaveBM.Value = False
     Me.StatusBar2.Panels(2).Text = "Tool: Pan"
    Else
    Me.MyMap2.MousePointer = moArrow
    Me.StatusBar2.Panels(2).Text = "Tool: None"
   
 End If
End Sub

Private Sub RavePreviousPage_Click()
Dim i As Integer
Dim varlist

PageCounter = PageCounter - 1
Me.RaveNextPage.Enabled = True
RstQuery.AbsolutePage = PageCounter
Me.PageCounterLabel = "Page " & RstQuery.AbsolutePage & " of " & RstQuery.PageCount

Me.LstResult.ListItems.Clear

For i = 1 To 10
Set varlist = Me.LstResult.ListItems.Add
    varlist.Text = IIf(IsNull(RstQuery("Stat_name")), "", RstQuery("Stat_name"))
    If Me.optBMs.Value = True Then
        varlist.SubItems(1) = RstQuery!ucode
    End If
    RstQuery.MoveNext
Next

If RstQuery.AbsolutePage = 2 Then
   Me.RavePreviousPage.Enabled = False
End If

End Sub

Private Sub RavePrint_Click()
FrmCertificate.Show 1

End Sub

Private Sub RaveProjectControl_Click()
Me.SSTab1.TabVisible(1) = 1
Me.SSTab1.Tab = 1
Me.SSTab1.TabVisible(1) = 0
End Sub

Private Sub RavePrintBM_Click()
FrmBenchmarksCertificate.Show 1
End Sub

Private Sub RavePrintGravity_Click()
    FormReportGravity.Show
End Sub

Private Sub RaveRemoveQuery_Click()
strcondition = ""

If Me.LstConditions.ListItems.Count > 0 Then
If Me.LstConditions.SelectedItem.Index = Me.LstConditions.ListItems.Count Then
    
    
            If Me.LstConditions.ListItems.Count <> 1 Then
                Me.LstConditions.ListItems(Me.LstConditions.SelectedItem.Index - 1).SubItems(3) = ""
            End If
            
            Me.LstConditions.ListItems.Remove (Me.LstConditions.SelectedItem.Index)
            
            If Me.LstConditions.ListItems.Count <> 0 Then
                BuildExecuteQuery
                Else
                Me.LstResult.ListItems.Clear
                Me.ResultLabel.Caption = ""
            End If
    
    Else
    
            
            Me.LstConditions.ListItems.Remove (Me.LstConditions.SelectedItem.Index)
            BuildExecuteQuery
            
            
End If
End If
End Sub

Private Sub RaveSave_Click()


If Validation = False Then 'Check if fields are valid
    Exit Sub
End If


If AddMode = True Then

  If IsDuplicateStation(Trim(FrmGCPDS.TxtName)) = True Then
    MsgBox "Station already exist! :)", vbInformation, "Duplicate Entry"
    Exit Sub
 Else
 cnn.Execute "Insert into geoprov (stat_name) values('" & Trim(Replace(Me.TxtName, "'", "''")) & "')"
    
 End If
End If

If EditMode = True Then
    cnn.Execute "Update geoprov set Stat_Name='" & Trim(Replace(Me.TxtName, "'", "''")) & "' where stat_name='" & Trim(Replace(rstRecords("stat_name"), "'", "''")) & "'"
End If
          
    
     
     If Trim(Me.TxtIsland) = "" Then
        cnn.Execute "Update geoprov set Island=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set Island='" & Trim(Replace(Me.TxtIsland, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtRegion) = "" Then
        cnn.Execute "Update geoprov set region=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set region='" & Trim(Replace(Me.TxtRegion, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtProvince) = "" Then
        cnn.Execute "Update geoprov set Province=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set Province='" & Trim(Replace(Me.TxtProvince, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtMunicipality) = "" Then
        cnn.Execute "Update geoprov set Municipal=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set Municipal='" & Trim(Replace(Me.TxtMunicipality, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtBarangay) = "" Then
        cnn.Execute "Update geoprov set Barangay=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set Barangay='" & Trim(Replace(Me.TxtBarangay, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.txtOrder) = "" Then
        cnn.Execute "Update geoprov set h_order=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set h_order=" & Order(Me.txtOrder.ListIndex + 1) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If

     If Trim(Me.TxtRef) = "" Then
        cnn.Execute "Update geoprov set h_ref=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set h_ref='" & Trim(Replace(Me.TxtRef, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     'Latitude
     If IsNumeric(Me.TxtDLat) And IsNumeric(Me.TxtMLat) And IsNumeric(Me.TxtSLat) Then
        cnn.Execute "Update geoprov set latitude=" & CDec(Me.TxtDLat) + (CDec(Me.TxtMLat) / 60) + (CDec(Me.TxtSLat) / 3600) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set latitude=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     'Longitude
     If IsNumeric(Me.TxtDLong) And IsNumeric(Me.TxtMLong) And IsNumeric(Me.TxtSLong) Then
        cnn.Execute "Update geoprov set longitude=" & CDec(Me.TxtDLong) + (CDec(Me.TxtMLong) / 60) + (CDec(Me.TxtSLong) / 3600) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set longitude=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtDLat) = "" Then
        cnn.Execute "Update geoprov set wgs84ED=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set wgs84ED=" & Me.TxtDLat & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtMLat) = "" Then
        cnn.Execute "Update geoprov set wgs84EM=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set wgs84EM=" & Me.TxtMLat & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtSLat) = "" Then
        cnn.Execute "Update geoprov set wgs84ES=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set wgs84ES=" & Me.TxtSLat & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
    
    If Trim(Me.TxtDLong) = "" Then
        cnn.Execute "Update geoprov set wgs84nD=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set wgs84nD=" & Me.TxtDLong & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtMLong) = "" Then
        cnn.Execute "Update geoprov set wgs84nM=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set wgs84nM=" & Me.TxtMLong & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtSLong) = "" Then
        cnn.Execute "Update geoprov set wgs84nS=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set wgs84nS=" & Me.TxtSLong & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If


    If Trim(Me.TxtLatD) = "" Then
        cnn.Execute "Update geoprov set d_lat=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set d_lat=" & Me.TxtLatD & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtLatM) = "" Then
        cnn.Execute "Update geoprov set m_lat=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set m_lat=" & Me.TxtLatM & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtLatS) = "" Then
        cnn.Execute "Update geoprov set s_lat=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set s_lat=" & Me.TxtLatS & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtLongD) = "" Then
        cnn.Execute "Update geoprov set d_long=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set d_long=" & Me.TxtLongD & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtLongM) = "" Then
        cnn.Execute "Update geoprov set m_long=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set m_long=" & Me.TxtLongM & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtLongS) = "" Then
        cnn.Execute "Update geoprov set s_long=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set s_long=" & Me.TxtLongS & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     'PTM
     If Trim(Me.TxtNorthing) = "" Then
        cnn.Execute "Update geoprov set northing=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set northing=" & CDec(Me.TxtNorthing) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtEasting) = "" Then
        cnn.Execute "Update geoprov set easting=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set easting=" & CDec(Me.TxtEasting) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
      If Trim(Me.TxtZone) = "" Then
        cnn.Execute "Update geoprov set zone=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        'modified dec72009
        'cnn.Execute "Update geoprov set zone=" & CDec(Me.TxtZone) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        cnn.Execute "Update geoprov set zone='" & Me.TxtZone & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     
     'UTM
     If Trim(Me.TextBoxUTMNorthing) = "" Then
        cnn.Execute "Update geoprov set utmy=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set utmy=" & CDec(Me.TextBoxUTMNorthing) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TextBoxUTMEasting) = "" Then
        cnn.Execute "Update geoprov set utmx=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set utmx=" & CDec(Me.TextBoxUTMEasting) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
      If Trim(Me.TextBoxUTMZone) = "" Then
        cnn.Execute "Update geoprov set utmz=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        
        cnn.Execute "Update geoprov set utmz='" & TextBoxUTMZone & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     
     '___
     If Trim(Me.txtEllipsoidalH) = "" Then
        cnn.Execute "Update geoprov set ell_hgt=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set ell_hgt=" & Me.txtEllipsoidalH & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtEllipsoidalH2) = "" Then
        cnn.Execute "Update geoprov set ellipz=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set ellipz=" & Me.TxtEllipsoidalH2 & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtDateEntry) = "" Or IsDate(Me.TxtDateEntry) = False Then
        cnn.Execute "Update geoprov set h_date_ety=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set h_date_ety='" & Me.TxtDateEntry & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtDateComputed) = "" Or IsDate(Me.TxtDateComputed) = False Then
        cnn.Execute "Update geoprov set h_date_com=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set h_date_com='" & CDate(Me.TxtDateComputed) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.txtLastRecover) = "" Or IsDate(Me.TxtDateEntry) = False Then
        cnn.Execute "Update geoprov set date_las_r=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set date_las_r='" & CDate(Me.txtLastRecover) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If

     If Trim(Me.txtEstablished) = "" Or IsDate(Me.txtEstablished) = False Then
        cnn.Execute "Update geoprov set date_est=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set date_est='" & CDate(Me.txtEstablished) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     
     If Trim(Me.txtYear) <> "" And IsNumeric(Me.txtYear) = True Then
        cnn.Execute "Update geoprov set date_est_year=" & Val(Me.txtYear) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set date_est_year=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.ComboBoxMonth.Text) <> "" Then
        cnn.Execute "Update geoprov set date_est_month=" & Me.ComboBoxMonth.ListIndex & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set date_est_month=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.ComboBoxDay.Text) <> "" Then
        cnn.Execute "Update geoprov set date_est_day=" & Me.ComboBoxDay.ListIndex & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set date_est_day=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.TxtFixingMethod) = "" Then
        cnn.Execute "Update geoprov set h_fix=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set h_fix=" & HorizontalFixingMethod(Me.TxtFixingMethod.ListIndex + 1) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.txtMarkPurpose) = "" Then
        cnn.Execute "Update geoprov set mark_const=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set mark_const=" & MarkPurpose(Me.txtMarkPurpose.ListIndex + 1) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.txtMarkType) = "" Then
        cnn.Execute "Update geoprov set mark_type=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set mark_type=" & MarkType(Me.txtMarkType.ListIndex + 1) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.txtMarkStatus) = "" Then
        cnn.Execute "Update geoprov set mark_stat=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set mark_stat=" & MarkStatus(Me.txtMarkStatus.ListIndex + 1) & " where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
      If Trim(Me.TxtMSL) = "" Then
        cnn.Execute "Update geoprov set AdoptedBy=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        cnn.Execute "Update geoprov set IsAdopted=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set AdoptedBy='" & Me.TxtMSL.Text & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        cnn.Execute "Update geoprov set IsAdopted=1 where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
      End If
     
     If Trim(Me.TxtEstablishedBy) = "" Then
        cnn.Execute "Update geoprov set hor_authty=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set hor_authty='" & Trim(Replace(Me.TxtEstablishedBy, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
     End If
     
     If Trim(Me.txtAuthority) = "" Then
        cnn.Execute "Update geoprov set authority=NULL where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        Else
        cnn.Execute "Update geoprov set authority='" & Trim(Replace(Me.txtAuthority, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
        
     End If

cnn.Execute "Update geoprov set Encoder='" & Trim(Replace(Encoder, "'", "''")) & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
cnn.Execute "Update geoprov set DateUpdated='" & Format(Now, "mm/dd/yyyy") & "' where stat_name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'"
    
    rstRecords.Requery
    If rstRecords.RecordCount > 0 Then
        rstRecords.MoveFirst
    End If
    rstRecords.Find "Stat_Name='" & Trim(Replace(Me.TxtName, "'", "''")) & "'", , adSearchForward
    
    'LoadAdopters
    SetToolBarStatus
    DisableFields
    AddMode = False
    EditMode = False
    
    PreviousRecord.stationName = Me.TxtName
    PreviousRecord.Region = Me.TxtRegion
    PreviousRecord.Province = Me.TxtProvince
    PreviousRecord.Municipality = Me.TxtMunicipality
    PreviousRecord.Barangay = Me.TxtBarangay
    PreviousRecord.Island = Me.TxtIsland
    PreviousRecord.Order = Me.txtOrder
    PreviousRecord.SurveyedBy = Me.TxtEstablishedBy
    
    Exit Sub
    
Hell:

    'Debug.Print Err.Number
    
    If Err.Number = -2147467259 Then
        MsgBox "Station Name required."
    End If
    
    If Err.Number = -2147217900 Then
        MsgBox "Station name already exist."
        Else
        MsgBox Err.Description
    End If
    
    
    
End Sub

Private Function Validation() As Boolean

Validation = False

            If Trim(Me.TxtName) = "" Then
                    Me.TxtName.SetFocus
                    MsgBox "Station Name is required."
                    Exit Function
            End If
            
            If Trim(Me.TxtRegion) = "" Then
                    MsgBox "Region is required.", vbCritical, "GCPDS"
                    Exit Function
            End If
            
            If Trim(Me.TxtProvince) = "" Then
                    MsgBox "Province is required."
                    Exit Function
            End If
            
            If Trim(Me.txtOrder) = "" Then
                    MsgBox "Order is required.", vbCritical, "GCPDS"
                    Exit Function
            End If
            
            If Trim(Me.TxtDateEntry) <> "" And IsDate(Me.TxtDateEntry) = False Then
                    Me.TxtDateEntry.SetFocus
                    Me.TxtDateEntry.SelStart = 0
                    Me.TxtDateEntry.SelLength = Len(Me.TxtDateEntry)
                    MsgBox "Date of Entry is invalid."
                    Exit Function
            End If
            
            If Trim(Me.TxtDateComputed) <> "" And IsDate(Me.TxtDateComputed) = False Then
                    Me.TxtDateComputed.SetFocus
                    Me.TxtDateComputed.SelStart = 0
                    Me.TxtDateComputed.SelLength = Len(Me.TxtDateComputed)
                    MsgBox "Date Computed is invalid."
                    Exit Function
            End If
            
            If Trim(Me.TxtLatD) <> "" And (Me.TxtLatD > 22 Or Me.TxtLatD < 4) Then
                    Me.TxtLatD.SetFocus
                    Me.TxtLatD.SelStart = 0
                    Me.TxtLatD.SelLength = Len(Me.TxtLatD)
                    MsgBox "Invalid Coordinates"
                    Exit Function
            End If
            
            
            If Trim(Me.TxtLatS) >= 60 Then
                    Me.TxtLatS.SetFocus
                    Me.TxtLatS.SelStart = 0
                    Me.TxtLatS.SelLength = Len(Me.TxtLatS)
                    MsgBox "Invalid Coordinates"
                    Exit Function
            End If
            
            If Trim(Me.TxtLongS) >= 60 Then
                    Me.TxtLongS.SetFocus
                    Me.TxtLongS.SelStart = 0
                    Me.TxtLongS.SelLength = Len(Me.TxtLongS)
                    MsgBox "Invalid Coordinates"
                    Exit Function
            End If
            
            If Trim(Me.TxtSLat) >= 60 Then
                    Me.TxtSLat.SetFocus
                    Me.TxtSLat.SelStart = 0
                    Me.TxtSLat.SelLength = Len(Me.TxtSLat)
                    MsgBox "Invalid Coordinates"
                    Exit Function
            End If
            
            If Trim(Me.TxtSLong) >= 60 Then
                    Me.TxtSLong.SetFocus
                    Me.TxtSLong.SelStart = 0
                    Me.TxtSLong.SelLength = Len(Me.TxtSLong)
                    MsgBox "Invalid Coordinates"
                    Exit Function
            End If
            
            'modified dec72009
            'If Trim(Me.TxtZone) <> "" And IsNumeric(Me.TxtZone) = False Then
            '        Me.TxtZone.SetFocus
            '        MsgBox "PTM Zone should be numeric. 1-5"
            '        exit Function
            'End If
            '---
            If Trim(Me.TxtZone) <> "" And IsNumeric(Me.TxtZone) = True And Val(Me.TxtZone) > 5 Then
                    Me.TxtZone.SetFocus
                    MsgBox "PTM Zone should be numeric. 1-5"
                    Exit Function
            End If
            
            
             If IsNumeric(Me.txtEllipsoidalH.Text) = False Then
                    Me.txtEllipsoidalH.SetFocus
                    MsgBox "Ellipsoidal Height should be numeric."
                    Exit Function
            End If
            
          Validation = True
End Function

Private Sub RaveSaveBM_Click()


If Trim(Me.TxtEName) = "" Then
        Me.TxtEName.SetFocus
        MsgBox "Benchmark name is required."
        Exit Sub
End If

If Trim(Me.TxtBMDateOfEntry) <> "" And IsDate(Me.TxtBMDateOfEntry) = False Then
        Me.TxtBMDateOfEntry.SetFocus
        Me.TxtBMDateOfEntry.SelStart = 0
        Me.TxtBMDateOfEntry.SelLength = Len(Me.TxtBMDateOfEntry)
        MsgBox "Date of Entry is invalid."
        Exit Sub
End If

If Trim(Me.TxtBMDateComputed) <> "" And IsDate(Me.TxtBMDateComputed) = False Then
        Me.TxtBMDateComputed.SetFocus
        Me.TxtBMDateComputed.SelStart = 0
        Me.TxtBMDateComputed.SelLength = Len(Me.TxtBMDateComputed)
        MsgBox "Date Computed is invalid."
        Exit Sub
End If

If Trim(Me.TxtBMDateLastRecovered) <> "" And IsDate(Me.TxtBMDateLastRecovered) = False Then
        Me.TxtBMDateLastRecovered.SetFocus
        Me.TxtBMDateLastRecovered.SelStart = 0
        Me.TxtBMDateLastRecovered.SelLength = Len(Me.TxtBMDateLastRecovered)
        MsgBox "Date Last Recovered is invalid."
        Exit Sub
End If

If Trim(Me.TxtBMDateEstablished) <> "" And IsDate(Me.TxtBMDateEstablished) = False Then
        Me.TxtBMDateEstablished.SetFocus
        Me.TxtBMDateEstablished.SelStart = 0
        Me.TxtBMDateEstablished.SelLength = Len(Me.TxtBMDateEstablished)
        MsgBox "Date Established is invalid."
        Exit Sub
End If

 
 If IsNumeric(Me.TxtElevation) = False Then
        Me.TxtElevation.SetFocus
        MsgBox "Elevation should be numeric."
        Exit Sub
End If
 If IsNumeric(Me.TxtBMPlus) = False Then
        Me.TxtBMPlus.SetFocus
        MsgBox "Plus/Minus should be numeric."
        Exit Sub
End If
 
        If Trim(Me.TextBoxLatitude.Text) <> "" Then
                If isValidCoordinate(Me.TextBoxLatitude.Text) = False Then
                    MsgBox "Invalid Latitude", vbCritical, "Benchmarks"
                Exit Sub
                End If
        End If
        
         If Trim(Me.TextBoxLongitude.Text) <> "" Then
                If isValidCoordinate(Me.TextBoxLongitude.Text) = False Then
                    MsgBox "Invalid Longitude", vbCritical, "Benchmarks"
                Exit Sub
                End If
        End If
 
 If AddModeBM = True Then
 
    If DuplicateBenchmark(Trim(Me.TxtEName), Trim(Me.TxtEProvince)) = True Then
         MsgBox "Duplicate Benchmark", vbInformation, "Benchmarks"
    Exit Sub
    
    End If
          Dim rstidentity As New ADODB.Recordset
          
           ' ghelo 10/23/2014
           'rstidentity.Open "Select IDENT_CURRENT('benchmarks')", cnn, adOpenStatic, adLockOptimistic
           
         ''  rstidentity.Open "Select MAX(ucode) from benchmarks", cnn, adOpenStatic, adLockOptimistic
           
            'u_code = rstidentity.Fields(0).Value
          '' u_code = rstidentity.Fields(0).Value + 1
           
        
         ' cnn.Execute "Insert into benchmarks (stat_name,ucode,island,region,province,municipal,barangay,e_order,e_date_ety,e_date_com,date_las_r,date_est,mark_const,mark_type,mark_stat,e_fix,elv_authty,authority,description,encoder,elevation,bmplus,latitude,longitude,e_datum)" _
            ' & " values('" & Replace(FrmGCPDS.TxtEName, "'", "''") & "'" & "," & u_code & "," & "'" & Replace(FrmGCPDS.TxtEIsland, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtERegion, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtEProvince, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtEMunicipality, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtEBarangay, "'", "''") & "'" & "," & IIf(Trim(FrmGCPDS.TxtBMOrder) = "", "Null", OrderBM(FrmGCPDS.TxtBMOrder.ListIndex + 1)) _
            ' & "," & IIf(Trim(FrmGCPDS.TxtBMDateOfEntry) = "", "Null", "'" & FrmGCPDS.TxtBMDateOfEntry & "'") & "," & IIf(Trim(FrmGCPDS.TxtBMDateComputed) = "", "Null", "'" & FrmGCPDS.TxtBMDateComputed & "'") & "," & IIf(Trim(FrmGCPDS.TxtBMDateLastRecovered) = "", "Null", "'" & FrmGCPDS.TxtBMDateLastRecovered & "'") & "," & IIf(Trim(FrmGCPDS.TxtBMDateEstablished) = "", "Null", "'" & FrmGCPDS.TxtBMDateEstablished & "'") & "," & IIf(Trim(FrmGCPDS.BMMarkPurpose) = "", "Null", MarkPurpose(FrmGCPDS.BMMarkPurpose.ListIndex + 1)) _
            ' & "," & IIf(Trim(FrmGCPDS.BMMarkType) = "", "Null", MarkType(FrmGCPDS.BMMarkType.ListIndex + 1)) & "," & IIf(Trim(FrmGCPDS.BMMarkStatus) = "", "Null", MarkStatus(FrmGCPDS.BMMarkStatus.ListIndex + 1)) & "," & IIf(Trim(FrmGCPDS.TxtEFix) = "", "Null", VerticalFixingMethod(FrmGCPDS.TxtEFix.ListIndex + 1)) & "," & "'" & Replace(FrmGCPDS.TxtElevationAuthority, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtBMAuthority, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtBMDescription, "'", "''") & "'" & "," & "'" & Replace(Encoder, "'", "''") & "'" & "," & IIf(Trim(FrmGCPDS.TxtElevation) = "", "Null", FrmGCPDS.TxtElevation) & "," & IIf(Trim(FrmGCPDS.TxtBMPlus) = "", "Null", FrmGCPDS.TxtBMPlus) & "," & IIf(Trim(FrmGCPDS.TextBoxLatitude) = "", "Null", DMStoDD(FrmGCPDS.TextBoxLatitude)) & "," & IIf(Trim(FrmGCPDS.TextBoxLongitude) = "", "Null", DMStoDD(FrmGCPDS.TextBoxLongitude)) & "," & VerticalDatum(Me.TxtEDatum.ListIndex + 1) & ")"
             
           cnn.Execute "Insert into benchmarks (stat_name,island,region,province,municipal,barangay,e_order,e_date_ety,e_date_com,date_las_r,date_est,mark_const,mark_type,mark_stat,e_fix,elv_authty,authority,description,encoder,elevation,bmplus,latitude,longitude,e_datum)" _
             & " values('" & Replace(FrmGCPDS.TxtEName, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtEIsland, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtERegion, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtEProvince, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtEMunicipality, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtEBarangay, "'", "''") & "'" & "," & IIf(Trim(FrmGCPDS.TxtBMOrder) = "", "Null", OrderBM(FrmGCPDS.TxtBMOrder.ListIndex + 1)) _
             & "," & IIf(Trim(FrmGCPDS.TxtBMDateOfEntry) = "", "Null", "'" & FrmGCPDS.TxtBMDateOfEntry & "'") & "," & IIf(Trim(FrmGCPDS.TxtBMDateComputed) = "", "Null", "'" & FrmGCPDS.TxtBMDateComputed & "'") & "," & IIf(Trim(FrmGCPDS.TxtBMDateLastRecovered) = "", "Null", "'" & FrmGCPDS.TxtBMDateLastRecovered & "'") & "," & IIf(Trim(FrmGCPDS.TxtBMDateEstablished) = "", "Null", "'" & FrmGCPDS.TxtBMDateEstablished & "'") & "," & IIf(Trim(FrmGCPDS.BMMarkPurpose) = "", "Null", MarkPurpose(FrmGCPDS.BMMarkPurpose.ListIndex + 1)) _
             & "," & IIf(Trim(FrmGCPDS.BMMarkType) = "", "Null", MarkType(FrmGCPDS.BMMarkType.ListIndex + 1)) & "," & IIf(Trim(FrmGCPDS.BMMarkStatus) = "", "Null", MarkStatus(FrmGCPDS.BMMarkStatus.ListIndex + 1)) & "," & IIf(Trim(FrmGCPDS.TxtEFix) = "", "Null", VerticalFixingMethod(FrmGCPDS.TxtEFix.ListIndex + 1)) & "," & "'" & Replace(FrmGCPDS.TxtElevationAuthority, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtBMAuthority, "'", "''") & "'" & "," & "'" & Replace(FrmGCPDS.TxtBMDescription, "'", "''") & "'" & "," & "'" & Replace(Encoder, "'", "''") & "'" & "," & IIf(Trim(FrmGCPDS.TxtElevation) = "", "Null", FrmGCPDS.TxtElevation) & "," & IIf(Trim(FrmGCPDS.TxtBMPlus) = "", "Null", FrmGCPDS.TxtBMPlus) & "," & IIf(Trim(FrmGCPDS.TextBoxLatitude) = "", "Null", DMStoDD(FrmGCPDS.TextBoxLatitude)) & "," & IIf(Trim(FrmGCPDS.TextBoxLongitude) = "", "Null", DMStoDD(FrmGCPDS.TextBoxLongitude)) & "," & VerticalDatum(Me.TxtEDatum.ListIndex + 1) & ")"
             
           
           rstidentity.Open "Select IDENT_CURRENT('benchmarks')", cnn, adOpenStatic, adLockOptimistic
           u_code = rstidentity.Fields(0).Value
           cnn.Execute "Update benchmarks set dateupdated=GetDate() where ucode=" & u_code
             
             
             
             
    Else
    
            
        cnn.Execute "Update benchmarks set Stat_Name='" & Trim(Replace(Me.TxtEName, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode
        
        'Island
      
        cnn.Execute "Update benchmarks set Island='" & Trim(Replace(Me.TxtEIsland, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode
        
        cnn.Execute "Update benchmarks set Region='" & Trim(Replace(Me.TxtERegion, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode
        cnn.Execute "Update benchmarks set Province='" & Trim(Replace(Me.TxtEProvince, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode
        cnn.Execute "Update benchmarks set Municipal='" & Trim(Replace(Me.TxtEMunicipality, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode
        cnn.Execute "Update benchmarks set Barangay='" & Trim(Replace(Me.TxtEBarangay, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode

     
 
        cnn.Execute "Update benchmarks set E_order=" & OrderBM(Me.TxtBMOrder.ListIndex + 1) & " where ucode=" & rstBenchmarks!ucode
        
        
        'cnn.Execute "Update benchmarks set Elevation=" & Me.TxtElevation & " where stat_name='" & Trim(Replace(Me.TxtEName, "'", "''")) & "'"
   
        
        'ELEVATIOn
        If Trim(Me.TxtElevation) = "" Then
            cnn.Execute "Update benchmarks set  Elevation= Null  where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update benchmarks set  Elevation=" & Me.TxtElevation & " where ucode=" & rstBenchmarks!ucode
        End If
        
        'BMPlus
        If Trim(Me.TxtBMPlus) = "" Then
            cnn.Execute "Update benchmarks set  BMPlus= Null where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update benchmarks set  BMPlus=" & Me.TxtBMPlus & " where ucode=" & rstBenchmarks!ucode
        End If
        
              'Latitude
        If Trim(Me.TextBoxLatitude) = "" Then
            cnn.Execute "Update benchmarks set  latitude= Null  where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update benchmarks set  latitude=" & DMStoDD(Me.TextBoxLatitude) & " where ucode=" & rstBenchmarks!ucode
        End If
        
              'Longitude
        If Trim(Me.TextBoxLongitude) = "" Then
            cnn.Execute "Update benchmarks set  longitude= Null  where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update benchmarks set  longitude=" & DMStoDD(Me.TextBoxLongitude) & " where ucode=" & rstBenchmarks!ucode
        End If
        
        cnn.Execute "Update benchmarks set E_Fix=" & VerticalFixingMethod(Me.TxtEFix.ListIndex + 1) & " where ucode=" & rstBenchmarks!ucode

        
        
        
        
        'Date Computed
        If Trim(Me.TxtBMDateComputed) = "" Or IsDate(Me.TxtBMDateComputed) = False Then
            cnn.Execute "Update benchmarks set E_date_Com=NULL where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update benchmarks set E_date_Com='" & CDate(Me.TxtBMDateComputed) & "' where ucode=" & rstBenchmarks!ucode
        End If
        
        
        'Date Last Recovered
        If Trim(Me.TxtBMDateLastRecovered) = "" Or IsDate(Me.TxtBMDateLastRecovered) = False Then
            cnn.Execute "Update benchmarks set Date_Las_R=NULL where  ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update benchmarks set Date_Las_R='" & CDate(Me.TxtBMDateLastRecovered) & "' where ucode=" & rstBenchmarks!ucode
        End If
        
        'Date Established
        If Trim(Me.TxtBMDateEstablished) = "" Or IsDate(Me.TxtBMDateEstablished) = False Then
            cnn.Execute "Update benchmarks set Date_est=NULL where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update benchmarks set Date_est='" & CDate(Me.TxtBMDateEstablished) & "' where ucode=" & rstBenchmarks!ucode
        End If
       
        'Vertical Datum
        If Trim(Me.TxtEDatum) = "" Then
            cnn.Execute "Update Benchmarks set e_datum=NULL where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update Benchmarks set e_Datum=" & VerticalDatum(Me.TxtEDatum.ListIndex + 1) & " where ucode=" & rstBenchmarks!ucode
        End If
        
        'Mark Purpose
        If Trim(Me.BMMarkPurpose) = "" Then
            cnn.Execute "Update Benchmarks set mark_const=NULL where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update Benchmarks set mark_const=" & MarkPurpose(Me.BMMarkPurpose.ListIndex + 1) & " where ucode=" & rstBenchmarks!ucode
        End If
        
        'Mark Type
        If Trim(Me.BMMarkType) = "" Then
            cnn.Execute "Update Benchmarks set mark_type=NULL where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update Benchmarks set mark_type=" & MarkType(Me.BMMarkType.ListIndex + 1) & " where ucode=" & rstBenchmarks!ucode
        End If
        
        'Mark Status
        If Trim(Me.BMMarkStatus) = "" Then
            cnn.Execute "Update Benchmarks set mark_stat=NULL where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update Benchmarks set mark_stat=" & MarkStatus(Me.BMMarkStatus.ListIndex + 1) & " where ucode=" & rstBenchmarks!ucode
        End If
        
        'Elevation Authority
        If Trim(Me.TxtElevationAuthority) = "" Then
            cnn.Execute "Update Benchmarks set elv_authty=NULL where ucode=" & rstBenchmarks!ucode
        Else
            cnn.Execute "Update Benchmarks set elv_authty='" & Trim(Replace(Me.TxtElevationAuthority, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode
        End If
     
        cnn.Execute "Update benchmarks set description='" & Trim(Replace(Me.TxtBMDescription, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode
        cnn.Execute "Update benchmarks set Encoder='" & Trim(Replace(Encoder, "'", "''")) & "' where ucode=" & rstBenchmarks!ucode
        cnn.Execute "Update benchmarks set dateupdated=GetDate() where ucode=" & rstBenchmarks!ucode
        
        u_code = rstBenchmarks!ucode
  End If
  
  
rstBenchmarks.Requery
rstBenchmarks.MoveFirst
rstBenchmarks.Find "ucode=" & u_code

SetToolBarStatusBM
EnableBenchmarksFields False
   
    
   

    
End Sub

Private Sub RaveSaveGravity_Click()
    
    If GravityFieldValidation = True Then
    
        SaveGravity
        rstGravity.Requery
        rstGravity.MoveFirst
        rstGravity.Find "stat_name='" & Trim(Me.TextBoxGravityName.Text) & "'", , adSearchForward
        SetToolBarStatusGravity
        DisableFieldsGravity
        
   End If
    
End Sub

Private Sub RaveSearch_Click()
 FrmStationName.Show 1
End Sub

Private Sub RaveSearchBM_Click()
FrmBenchmarks.Show 1
End Sub



Private Sub RaveSearchGravity_Click()
FrmSearchGravity.Show 1
End Sub

Private Sub RaveTab_Click(Index As Integer)
MapType = "GCP"
    Dim i As Integer
    For i = 0 To 5
        If RaveTab(i).Index = Index Then
           RaveTab(i).Value = True
           Else
           RaveTab(i).Value = False
        End If
    Next
    
    If Index = 0 Then
        Me.SSTab1.TabVisible(0) = True
        Me.SSTab1.Tab = 0
        Me.SSTab1.TabVisible(0) = False
    End If
    If Index = 1 Then
        Me.SSTab1.TabVisible(1) = True
        Me.SSTab1.Tab = 1
        Me.SSTab1.TabVisible(1) = False
    End If
    
    If Index = 2 Then
        Me.SSTab1.TabVisible(3) = True
        Me.SSTab1.Tab = 3
        Me.SSTab1.TabVisible(3) = False
    End If

    
    If Index = 3 Then
        Me.SSTab1.TabVisible(5) = True
        Me.SSTab1.Tab = 5
        Me.SSTab1.TabVisible(5) = False
    End If
    
     If Index = 4 Then
        Me.SSTab1.TabVisible(7) = True
        Me.SSTab1.Tab = 7
        Me.SSTab1.TabVisible(7) = False
    End If
    
    If Index = 5 Then
        Me.SSTab1.TabVisible(8) = True
        Me.SSTab1.Tab = 8
        Me.SSTab1.TabVisible(8) = False
    End If
End Sub

Public Sub FlashCurrentGCP()
If IsNumeric(rstRecords("wgs84ED")) And IsNumeric(rstRecords("wgs84EM")) And IsNumeric(rstRecords("wgs84ES")) And IsNumeric(rstRecords("wgs84ND")) And IsNumeric(rstRecords("wgs84NM")) And IsNumeric(rstRecords("wgs84NS")) Then
           FlashY = rstRecords("wgs84ED") + (rstRecords("wgs84EM") / 60) + (rstRecords("wgs84ES") / 3600)
           FlashX = rstRecords("wgs84ND") + (rstRecords("wgs84NM") / 60) + (rstRecords("wgs84NS") / 3600)
End If
End Sub

Private Sub RaveTranform2PRS92_Click()
If IsNumeric(Me.TxtDLat) And IsNumeric(Me.TxtMLat) And IsNumeric(Me.TxtSLat) And IsNumeric(Me.TxtDLong) And IsNumeric(Me.TxtMLong) And IsNumeric(Me.TxtSLong) And IsNumeric(Me.TxtEllipsoidalH2) Then
    WGS84_TO_PRS92
    Else
    MsgBox "WGS84 data cannot be tranform to PRS92 due to invalid coordinates or ellipsoidal height."
End If
End Sub

Private Sub RaveTranform2UTM_Click()
       
       Dim latitude As Double
       Dim longitude As Double
       Dim MyUTM As UTM
       
       
       If IsNumeric(Me.TxtLatD) = True And IsNumeric(txtladm) = True And IsNumeric(TxtLatS) = True And IsNumeric(TxtLongD) = True And IsNumeric(TxtLongM) = True And IsNumeric(TxtLongS) = True Then
            latitude = Val(Me.TxtLatD.Text) + (Val(Me.TxtLatM.Text) / 60) + (Val(Me.TxtLatS.Text) / 3600)
            longitude = Val(Me.TxtLongD.Text) + (Val(Me.TxtLongM.Text) / 60) + (Val(Me.TxtLongS.Text) / 3600)
            MyUTM = LatLonToUTM(latitude, longitude)
            Me.TextBoxUTMNorthing = MyUTM.Northing
            Me.TextBoxUTMEasting = MyUTM.Easting
            Me.TextBoxUTMZone = MyUTM.Zone
       End If
       
End Sub

Private Sub RaveTranform2WGS84_Click()
If IsNumeric(Me.TxtLatD) And IsNumeric(Me.TxtLatM) And IsNumeric(Me.TxtLatS) And IsNumeric(Me.TxtLongD) And IsNumeric(Me.TxtLongM) And IsNumeric(Me.TxtLongS) And IsNumeric(Me.txtEllipsoidalH) Then
    PRS92_TO_WGS84
    Else
    MsgBox "PRS92 data cannot be tranform to WGS84 due to invalid coordinates or ellipsoidal height."
End If
End Sub

Private Sub RaveTranformToPTM_Click()
    Dim CM As Double
    If IsNumeric(Me.TxtLatD) = True And IsNumeric(txtladm) = True And IsNumeric(TxtLatS) = True And IsNumeric(TxtLongD) = True And IsNumeric(TxtLongM) = True And IsNumeric(TxtLongS) = True Then
      CM = GetCentralMeridian(Trim(Me.TxtRegion), Trim(Me.TxtProvince), Trim(Me.TxtMunicipality))
      If CM = 0 Then
        FrmCentralMeridian.Show 1
        CM = PTMZone

        If CM = 0 Then
            Exit Sub
        End If
    End If
      
      Call Compute(CDec(Me.TxtLatD), CDec(Me.TxtLatM), CDec(Me.TxtLatS), CDec(Me.TxtLongD), CDec(Me.TxtLongM), CDec(Me.TxtLongS), CM)
        FrmGCPDS.TxtNorthing = Format(North, "#,##0.####")
        FrmGCPDS.TxtEasting = Format(East, "#,##0.####")
        FrmGCPDS.TxtZone = Zone
        
    Else
        MsgBox "Invalid PRS92 Coordinates.", vbInformation
        
    End If
End Sub

Private Sub RaveUserAccounts_Click()
    FrmUsersAccounts.Show 1
End Sub

Private Sub RaveZoomIn_Click()

    Dim rect As New MapObjects.Rectangle '
    Set rect = Me.MyMap.Extent
    rect.ScaleRectangle (0.8)
    Set Me.MyMap.Extent = rect

End Sub

Private Sub RaveZoomInBM_Click()
If Me.RaveZoomInBM.Value = True Then
    Me.MyMap2.MousePointer = moZoomIn
    Me.RavePanBM.Value = False
    Me.RaveBM.Value = False
    Me.StatusBar2.Panels(2).Text = "Tool: Zoom Box"
    Else
    Me.MyMap2.MousePointer = moArrow
    Me.StatusBar2.Panels(2).Text = "Tool: None"
 End If
End Sub

Private Sub RaveZoomOut_Click()
Dim R

Dim i As Integer

   Set R = MyMap.Extent
        R.ScaleRectangle 5
        MyMap.Extent = R

        
Me.StatusBar1.Panels(3).Text = "Zoom: " & Round(MyMap.Extent.Height, 2)

End Sub

Private Sub RaveZoomOutBM_Click()
Dim R

Dim i As Integer

   Set R = MyMap2.Extent
        R.ScaleRectangle 1.1
        MyMap2.Extent = R
         Me.StatusBar2.Panels(3).Text = "Zoom: " & Round(MyMap2.Extent.Height, 3)
End Sub

Private Sub RequestingPartyLibRave_Click()
FrmRequestingPartyLib.Show 1
End Sub

Private Sub SignatoryLibRave_Click()
FrmSignatoryLib.Show 1
End Sub




Private Sub SummaryRave_Click()
FrmMasterlist.Show 1
End Sub



Private Sub TextBoxGravityElevation_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 45 Then
   If Me.ActiveControl.SelStart = 0 And Mid(Me.ActiveControl.Text, 1, 1) <> "-" Then
        Exit Sub
        
        
   Else
        KeyAscii = 0
      Exit Sub
   End If
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TextBoxObservedValues_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 45 Then
   If Me.ActiveControl.SelStart = 0 And Mid(Me.ActiveControl.Text, 1, 1) <> "-" Then
        Exit Sub
        
        
   Else
        KeyAscii = 0
      Exit Sub
   End If
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TextBoxTheoreticalValue_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 45 Then
   If Me.ActiveControl.SelStart = 0 And Mid(Me.ActiveControl.Text, 1, 1) <> "-" Then
        Exit Sub
        
        
   Else
        KeyAscii = 0
      Exit Sub
   End If
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Timer1_Timer()
'On Error GoTo hell
Dim ptx As New MapObjects.Point

Dim fnt As New StdFont
Dim latitude As Double
Dim longitude As Double



                ptx.y = FlashY
                ptx.x = FlashX
               Me.MyMap.FlashShape ptx, 1
        


Exit Sub

Hell:
End Sub



Private Sub txtAuthority_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   
   If (KeyCode = 9) Then
        Me.TxtName.SetFocus
   End If
   
End Sub

Private Sub TxtBarangay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.TxtBarangay.ToolTipText = Me.TxtBarangay
End Sub

Private Sub TxtDatum_Change()

End Sub











Private Sub TxtBMDescription_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If (KeyCode = 9) Then
        Me.TxtEName.SetFocus
   End If
End Sub

Private Sub TxtDLat_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtDLong_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtEasting_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub
















Private Sub txtEllipsoidalH_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


If (KeyAscii) = 45 Then
   If Me.txtEllipsoidalH.SelStart = 0 And Mid(Me.txtEllipsoidalH.Text, 1, 1) <> "-" Then
        Exit Sub
        
        
   Else
        KeyAscii = 0
      Exit Sub
   End If
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.txtEllipsoidalH, ".") = 0 And Trim(Me.txtEllipsoidalH) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub




Private Sub txtEOrder_KeyPress(KeyAscii As Integer)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtEllipsoidalH2_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Public Sub EnableBenchmarksFields(bool As Boolean)

Me.TxtEName.BorderStyle = IIf(bool = True, 1, 0)
Me.TxtEName.Locked = Not bool


Me.TxtERegion.BorderStyle = IIf(bool = True, 1, 0)
Me.TxtEProvince.BorderStyle = IIf(bool = True, 1, 0)
Me.TxtEMunicipality.BorderStyle = IIf(bool = True, 1, 0)
Me.TxtEBarangay.BorderStyle = IIf(bool = True, 1, 0)

Me.RaveButtons17.Visible = bool

TxtElevation.BorderStyle = IIf(bool = True, 1, 0)
TxtBMPlus.BorderStyle = IIf(bool = True, 1, 0)
TxtEFix.BorderStyle = IIf(bool = True, 1, 0)
TxtBMDateOfEntry.BorderStyle = IIf(bool = True, 1, 0)
TxtBMDateComputed.BorderStyle = IIf(bool = True, 1, 0)
TxtBMDateLastRecovered.BorderStyle = IIf(bool = True, 1, 0)
TxtBMOrder.BorderStyle = IIf(bool = True, 1, 0)
TxtEDatum.BorderStyle = IIf(bool = True, 1, 0)
TxtEIsland.BorderStyle = IIf(bool = True, 1, 0)
BMMarkPurpose.BorderStyle = IIf(bool = True, 1, 0)
BMMarkStatus.BorderStyle = IIf(bool = True, 1, 0)
BMMarkType.BorderStyle = IIf(bool = True, 1, 0)
TxtBMDateEstablished.BorderStyle = IIf(bool = True, 1, 0)
TxtElevationAuthority.BorderStyle = IIf(bool = True, 1, 0)
TxtBMAuthority.BorderStyle = IIf(bool = True, 1, 0)
TextBoxLatitude.BorderStyle = IIf(bool = True, 1, 0)
TextBoxLongitude.BorderStyle = IIf(bool = True, 1, 0)
'TxtBMDescription.BorderStyle = IIf(bool = True, 1, 0)

TxtElevation.Locked = Not bool
TxtBMPlus.Locked = Not bool
TxtEFix.Locked = Not bool
TxtBMDateOfEntry.Enabled = Not bool
TxtBMDateComputed.Locked = Not bool
TxtBMDateLastRecovered.Locked = Not bool
TxtBMOrder.Locked = Not bool
TxtEDatum.Locked = Not bool
TxtEIsland.Locked = Not bool
BMMarkPurpose.Locked = Not bool
BMMarkStatus.Locked = Not bool
BMMarkType.Locked = Not bool
TxtBMDateEstablished.Locked = Not bool
TxtElevationAuthority.Locked = Not bool
TxtBMAuthority.Locked = Not bool
TextBoxLatitude.Locked = Not bool
TextBoxLongitude.Locked = Not bool
TxtBMDescription.Locked = Not bool

If (bool = True) Then

TextBoxLatitude.Text = Replace(Replace(Replace(TextBoxLatitude.Text, "°", ""), "'", ""), """", "")
TextBoxLongitude.Text = Replace(Replace(Replace(TextBoxLongitude.Text, "°", ""), "'", ""), """", "")
End If

End Sub



Private Sub TxtFixingMethod_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 46 Then
   Me.ActiveControl.ListIndex = -1
End If
End Sub



Private Sub txtLatitude_Click()
If EditMode = True Or AddMode = True Then
    FrmLatLong.Show 1
End If
End Sub

Private Sub txtLongitude_Click()
If EditMode = True Or AddMode = True Then
    FrmLatLong.Show 1
End If
End Sub













Private Sub TxtLatD_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtLatitude_KeyPress(KeyAscii As MSForms.ReturnInteger)

'If (KeyAscii) = 8 Then
'    Exit Sub
'End If
'
'                If (KeyAscii) = 32 Or (KeyAscii) = 45 Then
'
'                   If InStr(1, Me.ActiveControl, "°") = 0 Then
'                      If Me.ActiveControl.Text <> "" Then
'                        Me.ActiveControl.Text = Me.ActiveControl.Text & "° "
'                      End If
'                      KeyAscii = 0
'                      Exit Sub
'                   End If
'
'                   If InStr(1, Me.ActiveControl, "'") = 0 Then
'
'                      If Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = "°" Then
'                         Me.ActiveControl.Text = Me.ActiveControl.Text & " "
'                         ElseIf Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = " " Then
'                             KeyAscii = 0
'                             Exit Sub
'                         Else
'                         Me.ActiveControl.Text = Me.ActiveControl.Text & "' "
'                      End If
'
'                      KeyAscii = 0
'                      Exit Sub
'                   End If
'
'                   If InStr(1, Me.ActiveControl, """") = 0 Then
'                      If Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = "'" Then
'                         Me.ActiveControl.Text = Me.ActiveControl.Text & " "
'                         ElseIf Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = " " Then
'                             KeyAscii = 0
'                             Exit Sub
'                         Else
'                         Me.ActiveControl.Text = Me.ActiveControl.Text & """"
'                      End If
'                      KeyAscii = 0
'                      Exit Sub
'                   End If
'
'                End If
'
'
'                If (KeyAscii) = 46 Then
'                    If InStr(1, (Me.ActiveControl), "'") > 0 And InStr(1, (Me.ActiveControl), ".") = 0 Then
'                     Exit Sub
'                   End If
'                End If
'
'
'
'
'    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
'       If Trim(Me.ActiveControl.Text) <> "" Then
'            If Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = """" Then
'                KeyAscii = 0
'                Exit Sub
'            End If
'       End If
'    Else
'        KeyAscii = 0
'        Exit Sub
'    End If
End Sub

Private Sub TxtLatM_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtLatS_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub TxtLongD_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtLongitude_KeyPress(KeyAscii As MSForms.ReturnInteger)
'If (KeyAscii) = 8 Then
'    Exit Sub
'End If
'
'                If (KeyAscii) = 32 Or (KeyAscii) = 45 Then
'
'                   If InStr(1, Me.ActiveControl, "°") = 0 Then
'                      If Me.ActiveControl.Text <> "" Then
'                        Me.ActiveControl.Text = Me.ActiveControl.Text & "° "
'                      End If
'                      KeyAscii = 0
'                      Exit Sub
'                   End If
'
'                   If InStr(1, Me.ActiveControl, "'") = 0 Then
'
'                      If Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = "°" Then
'                         Me.ActiveControl.Text = Me.ActiveControl.Text & " "
'                         ElseIf Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = " " Then
'                             KeyAscii = 0
'                             Exit Sub
'                         Else
'                         Me.ActiveControl.Text = Me.ActiveControl.Text & "' "
'                      End If
'
'                      KeyAscii = 0
'                      Exit Sub
'                   End If
'
'                   If InStr(1, Me.ActiveControl, """") = 0 Then
'                      If Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = "'" Then
'                         Me.ActiveControl.Text = Me.ActiveControl.Text & " "
'                         ElseIf Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = " " Then
'                             KeyAscii = 0
'                             Exit Sub
'                         Else
'                         Me.ActiveControl.Text = Me.ActiveControl.Text & """"
'                      End If
'                      KeyAscii = 0
'                      Exit Sub
'                   End If
'
'                End If
'
'
'                If (KeyAscii) = 46 Then
'                    If InStr(1, (Me.ActiveControl), "'") > 0 And InStr(1, (Me.ActiveControl), ".") = 0 Then
'                     Exit Sub
'                   End If
'                End If
'
'
'
'
'    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
'       If Trim(Me.ActiveControl.Text) <> "" Then
'            If Mid(Me.ActiveControl.Text, Len(Me.ActiveControl), 1) = """" Then
'                KeyAscii = 0
'                Exit Sub
'            End If
'       End If
'    Else
'        KeyAscii = 0
'        Exit Sub
'    End If
End Sub

Private Sub TxtLongM_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtLongS_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtM_long_Change()

End Sub

Private Sub TxtMarkPurpose_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 46 Then
   Me.ActiveControl.ListIndex = -1
End If
End Sub

Private Sub TxtMarkStatus_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 46 Then
   Me.ActiveControl.ListIndex = -1
End If
End Sub

Private Sub TxtMarkType_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 46 Then
   Me.ActiveControl.ListIndex = -1
End If
End Sub

Private Sub TxtMLat_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub TxtMLong_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtMonth_Change()

End Sub

Private Sub TxtMunicipality_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.TxtMunicipality.ToolTipText = Me.TxtMunicipality
End Sub





































Private Sub TxtName_Change()
Me.TxtName = UCase(Me.TxtName)
End Sub



Private Sub TxtNorthing_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub



Private Sub txtOrder_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 46 Then
   Me.ActiveControl.ListIndex = -1
End If
End Sub

Private Sub TxtProvince_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.TxtProvince.ToolTipText = Me.TxtProvince
End Sub









Sub DoZoom()
Dim R  ' get a rectangle from the user
Dim i As Integer
  Set R = Me.MyMap.TrackRectangle
  
  
  
  ' zoom to the rectangle if its valid
  If Not R Is Nothing Then MyMap.Extent = R
  
Me.StatusBar1.Panels(3).Text = "Zoom: " & Round(MyMap.Extent.Height, 2)









'If MyMap.Extent.Height > 1 And MyMap.Extent.Height < 17 Then
'   If 20 - (Round(MyMap.Extent.Height) * 3) >= 1 Then
'    MyMap.TrackingLayer.Symbol(1).Font.Bold = False
'    MyMap.TrackingLayer.Symbol(1).SIZE = 20 - (Round(MyMap.Extent.Height) * 3)
'    MyMap.TrackingLayer.Symbol(2).SIZE = 20 - (Round(MyMap.Extent.Height) * 3)
'    MyMap.TrackingLayer.Symbol(3).SIZE = 20 - (Round(MyMap.Extent.Height) * 3)
'    MyMap.TrackingLayer.Symbol(4).SIZE = 20 - (Round(MyMap.Extent.Height) * 3)
'   End If
'Else
'MyMap.TrackingLayer.Symbol(1).Font.Bold = True
'MyMap.TrackingLayer.Symbol(1).SIZE = 20
'MyMap.TrackingLayer.Symbol(2).SIZE = 20
'MyMap.TrackingLayer.Symbol(3).SIZE = 20
'MyMap.TrackingLayer.Symbol(4).SIZE = 20
'End If




End Sub

Sub DoZoom2()
Dim R  ' get a rectangle from the user
Dim i As Integer
  Set R = Me.MyMap2.TrackRectangle
  
  
  ' zoom to the rectangle if its valid
  If Not R Is Nothing Then MyMap2.Extent = R
  Me.StatusBar2.Panels(3).Text = "Zoom: " & Round(MyMap2.Extent.Height, 3)

End Sub

'Sub IdentifyGCP(X As Single, Y As Single)
'    Dim l
'    Dim P As New MapObjects.Point
'
'
'   get the layers
'  Set l = FrmGCPDS.MyMap.TrackingLayer
'  MsgBox l
'   transform the point to map coordinates
'  Set P = FrmGCPDS.MyMap.ToMapPoint(X, Y)
'
'
'
'
'   perform the search
'
'    Set recs = l.SearchByDistance(P, MyMap.ToMapDistance(100), "")
'
'
'  If Not recs.EOF Then
'    FrmIdentify.Show
'  End If
'
'  For Each fld In recs.Fields
'  MsgBox fld.name
'  Next
'End Sub






Public Sub SetToolBarStatus()
    
    If rstRecords.RecordCount > 0 Then
        FillUp
            
            If rstRecords.RecordCount = 1 Then
               SingleMode
               
               Else
                    
                   If rstRecords.bookmark = 1 Then
                      FirstMode
                      ElseIf rstRecords.bookmark = rstRecords.RecordCount Then
                      LastMode
                      Else
                      BrowseMode
                   End If
               
            End If
        
        
       Else
           
             ZeroMode
             BlankForm
           
    End If
End Sub


Public Sub SetToolBarStatusBM()
    If rstBenchmarks.RecordCount > 0 Then
        FillUpBenchmarks
            
            If rstBenchmarks.RecordCount = 1 Then
               SingleModeBM
               
               Else
                    
                   If rstBenchmarks.bookmark = 1 Then
                      FirstModeBM
                      ElseIf rstBenchmarks.bookmark = rstBenchmarks.RecordCount Then
                      LastModeBM
                      Else
                      BrowseModeBM
                   End If
               
            End If
        
        
       Else
           
             ZeroModeBM
           
    End If
End Sub

Public Sub SetToolBarStatusGravity()
    
    If rstGravity.RecordCount > 0 Then
        FillUpGravity
            
            If rstGravity.RecordCount = 1 Then
               SingleModeGravity
               
               Else
                    
                   If rstGravity.bookmark = 1 Then
                      FirstModeGravity
                      ElseIf rstGravity.bookmark = rstGravity.RecordCount Then
                      LastModeGravity
                      Else
                      BrowseModeGravity
                   End If
               
            End If
        
        
       Else
           
             ZeroModeGravity
             BlankFormGravity
           
    End If
End Sub



Private Sub TxtS_long_Change()
End Sub

Private Sub txtSLat_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub txtSLong_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If

If (KeyAscii) = 46 Then
   If InStr(1, Me.ActiveControl, ".") = 0 And Trim(Me.ActiveControl) <> "" Then
           Exit Sub
      Else
      KeyAscii = 0
      Exit Sub
   End If
End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub




Private Sub InitializeRaves()
RaveAdd.Left = RaveSearch.Left + RaveSearch.Width
RaveEdit.Left = RaveAdd.Left + RaveSearch.Width
RaveDelete.Left = RaveEdit.Left + RaveSearch.Width
RaveSave.Left = RaveDelete.Left + RaveSearch.Width
RaveCancel.Left = RaveSave.Left + RaveSearch.Width
RaveBack.Left = RaveCancel.Left + RaveSearch.Width
RaveNext.Left = RaveBack.Left + RaveSearch.Width
RavePrint.Left = RaveNext.Left + RaveSearch.Width
RaveGIS.Left = RavePrint.Left + RaveSearch.Width
RaveInfo.Left = RaveGIS.Left + RaveSearch.Width
RaveImages.Left = RaveInfo.Left + RaveSearch.Width


RaveAdd.Top = RaveSearch.Top
RaveEdit.Top = RaveSearch.Top
RaveDelete.Top = RaveSearch.Top
RaveSave.Top = RaveSearch.Top
RaveCancel.Top = RaveSearch.Top
RaveBack.Top = RaveSearch.Top
RaveNext.Top = RaveSearch.Top
RavePrint.Top = RaveSearch.Top
RaveInfo.Top = RaveSearch.Top
RaveImages.Top = RaveSearch.Top

End Sub



Private Sub InitializeMap()
Dim DC As New DataConnection



Dim LayerMap As New MapLayer
Dim LayerMap2 As New MapLayer
Dim LayerRoads As New MapLayer
'Dim LayerPOI As New MapLayer

'Dim ptx As New MapObjects.Point
'Set ptx = MyMap.ToMapPoint(6450, 3945)
DC.Database = App.Path & "\maps"
DC.Connect


LayerMap.GeoDataset = DC.FindGeoDataset("PHIL_WGS84_LUZON11")
LayerMap2.GeoDataset = DC.FindGeoDataset("PHIL_WGS84_LUZON11")
LayerRoads.GeoDataset = DC.FindGeoDataset("philippines_highway.shp")


LayerMap.Symbol.Color = &H8000&     'moDarkGreen 'moLimeGreen



MyMap.Layers.Add LayerMap
MyMap.Layers.Add LayerMap2
MyMap.Layers.Add LayerRoads


LayerMap2.Symbol.Style = moTransparentFill
LayerRoads.Symbol.Color = RGB(0, 255, 0)


Set MyMap.Layers(1).Renderer = MyLabelRender



MyLabelRender.Symbol(0).Font = "Arial"
MyLabelRender.Symbol(0).Font.SIZE = 11
MyLabelRender.Symbol(0).Color = vbWhite
MyLabelRender.Symbol(0).HorizontalAlignment = moAlignCenter
MyLabelRender.Symbol(0).VerticalAlignment = moAlignCenter
MyLabelRender.AllowDuplicates = False
MyLabelRender.Field = "Province"


MyMap.TrackingLayer.SymbolCount = 6
End Sub


Public Sub InitializeMapBM()
Dim DC As New DataConnection

Dim LayerMap As New MapLayer
Dim LayerMap2 As New MapLayer
Dim LayerRoads As New MapLayer
Dim LayerPOI As New MapLayer

DC.Database = App.Path & "\maps"
DC.Connect

LayerMap.GeoDataset = DC.FindGeoDataset("PHIL_WGS84_LUZON11")
LayerRoads.GeoDataset = DC.FindGeoDataset("philippines_highway.shp")
LayerPOI.GeoDataset = DC.FindGeoDataset("philippines_poi.shp")
LayerMap2.GeoDataset = DC.FindGeoDataset("PHIL_WGS84_LUZON11")

LayerMap.Symbol.Color = &H8000&     'moDarkGreen 'moLimeGreen
LayerMap2.Symbol.Style = moTransparentFill
LayerMap2.Symbol.Outline = False





LayerPOI.Symbol.SymbolType = 0 'moPointSymbol
LayerPOI.Symbol.Style = moTrueTypeMarker
LayerPOI.Symbol.Font = "ESRI Cartography"
LayerPOI.Symbol.Color = vbGreen
LayerPOI.Symbol.CharacterIndex = 75
LayerPOI.Symbol.SIZE = 5

MyMap2.Layers.Add LayerMap
MyMap2.Layers.Add LayerRoads
MyMap2.Layers.Add LayerPOI
MyMap2.Layers.Add LayerMap2
LayerRoads.Symbol.Color = RGB(0, 255, 0)

End Sub

Private Sub PlotBM()
Dim i As Long
Dim rsttemp As New ADODB.Recordset
Dim longitude As Double
Dim latitude As Double



MyMap2.TrackingLayer.ClearEvents
MyMap2.TrackingLayer.SymbolCount = 1

   
    rsttemp.Open "Select stat_name,longitude,latitude from benchmarks where longitude is not null and latitude is not null", cnn, adOpenStatic
    
ReDim gEventBm(0)
    For i = 0 To rsttemp.RecordCount - 1
        
       
            DrawBM rsttemp!longitude, rsttemp!latitude, 1
            ReDim Preserve gEventBm(UBound(gEventBm) + 1)
            gEventBm(UBound(gEventBm)) = rsttemp!Stat_Name
            rsttemp.MoveNext
    Next
    
End Sub



Private Sub PlotGCP()
Dim i As Long
Dim rsttemp As New ADODB.Recordset
Dim longitude As Double
Dim latitude As Double


Me.MousePointer = vbHourglass
MyMap.TrackingLayer.ClearEvents

Dim txtsym As New MapObjects.TextSymbol



    fnt.name = "ESRI Cartography"
    fnt.Bold = False



'First Order
With MyMap.TrackingLayer.Symbol(1)
 .SymbolType = 0 'moPointSymbol
 .Style = moTrueTypeMarker
 .Font = fnt
 .CharacterIndex = 121
 .SIZE = 20
 .Color = vbCyan
End With

'2nd Order
With MyMap.TrackingLayer.Symbol(2)
 .SymbolType = 0 'moPointSymbol
 .Style = moTrueTypeMarker
 .Font = fnt
 .CharacterIndex = 121

.SIZE = 20
.Color = vbGreen
End With

'3rd Order
With MyMap.TrackingLayer.Symbol(3)
 .SymbolType = 0 'moPointSymbol
 .Style = moTrueTypeMarker
 .Font = fnt
 .CharacterIndex = 121

.SIZE = 20
.Color = RGB(249, 166, 238)
End With

'4rd Order
With MyMap.TrackingLayer.Symbol(4)
 .SymbolType = 0 'moPointSymbol
 .Style = moTrueTypeMarker
 .Font = fnt
 .CharacterIndex = 121

.SIZE = 20
.Color = RGB(102, 192, 253)
End With

'AGS
With MyMap.TrackingLayer.Symbol(5)
 .SymbolType = 0 'moPointSymbol
 .Style = moTrueTypeMarker
 .Font = fnt
 .CharacterIndex = 121

.SIZE = 40
.Color = RGB(255, 255, 0)
End With

DoEvents
rsttemp.Open rstRecords.Source, cnn, adOpenStatic, adLockBatchOptimistic
rsttemp.MoveFirst

ReDim gEventList(0)
ReDim gEventsTag(0)
ReDim gEventOrder(0)


    Me.ProgressBar1.Visible = True
    Me.ProgressBar1.Value = 0
    Me.LabelProgress.Visible = True
    Me.LabelProgress = "Loading GCPs"

For i = 0 To rsttemp.RecordCount - 1
   
   Me.ProgressBar1.Value = i * (100 / rsttemp.RecordCount)
   DoEvents
   
    If IsNumeric(rsttemp("wgs84ED")) And IsNumeric(rsttemp("wgs84EM")) And IsNumeric(rsttemp("wgs84ES")) And IsNumeric(rsttemp("wgs84ND")) And IsNumeric(rsttemp("wgs84NM")) And IsNumeric(rsttemp("wgs84NS")) Then
           latitude = rsttemp("wgs84ED") + (rsttemp("wgs84EM") / 60) + (rsttemp("wgs84ES") / 3600)
           longitude = rsttemp("wgs84ND") + (rsttemp("wgs84NM") / 60) + (rsttemp("wgs84NS") / 3600)
        
        If rsttemp("h_order") = 1 Then
        
        
            DrawPoints longitude, latitude, 1
            ElseIf rsttemp("h_order") = 2 Then
            DrawPoints longitude, latitude, 2
            ElseIf rsttemp("h_order") = 3 Then
            DrawPoints longitude, latitude, 3
            ElseIf rsttemp("h_order") = 4 Then
            DrawPoints longitude, latitude, 4
            ElseIf rsttemp("h_order") = 5 Then
            DrawPoints longitude, latitude, 5
            Else
            DrawPoints longitude, latitude, 5
            
        End If
        
       
        
        ReDim Preserve gEventList(UBound(gEventList) + 1)
        ReDim Preserve gEventsTag(UBound(gEventsTag) + 1)
        ReDim Preserve gEventOrder(UBound(gEventOrder) + 1)
        
        Set gEventList(UBound(gEventList)) = Me.MyMap.TrackingLayer.Event(UBound(gEventList) - 1)
        gEventsTag(UBound(gEventsTag)) = rsttemp("stat_name")
        gEventOrder(UBound(gEventOrder)) = IIf(IsNull(rsttemp("h_order")), 0, rsttemp("h_order"))
        
    End If
    
    rsttemp.MoveNext
Next


    Me.ProgressBar1.Visible = False
    Me.ProgressBar1.Value = 0
    Me.LabelProgress.Visible = False
    Me.LabelProgress = ""
    Me.MousePointer = vbArrow
    
    
End Sub


Public Sub DrawPoints(x As Double, y As Double, Index As Integer)

Dim gcp As New MapObjects.Point

   'Set gcp = FrmGCPDS.mymap.ToMapPoint(X, Y)
   gcp.x = x
   gcp.y = y
 
   MyMap.TrackingLayer.AddEvent gcp.x, gcp.y, Index
  

End Sub
Public Sub DrawBM(x As Double, y As Double, Index As Integer)

Dim txt As New MapObjects.TextSymbol
MyMap2.TrackingLayer.Symbol(0).SIZE = 15
MyMap2.TrackingLayer.Symbol(0).Font = "ESRI Cartography"
MyMap2.TrackingLayer.Symbol(0).CharacterIndex = 120
MyMap2.TrackingLayer.Symbol(0).SymbolType = moPointSymbol
MyMap2.TrackingLayer.Symbol(0).Style = moTrueTypeMarker
MyMap2.TrackingLayer.Symbol(0).Color = vbCyan


Dim bm As New MapObjects.Point
   bm.x = x
   bm.y = y
   MyMap2.TrackingLayer.AddEvent bm.x, bm.y, 0
   
End Sub

Private Sub CenterView()

Dim ptx As New MapObjects.Point
Dim latitude As Double
Dim longitude As Double
Dim R

Me.MyMap.Extent = Me.MyMap.FullExtent
Set R = MyMap.Extent
        R.ScaleRectangle 0.001
        
        MyMap.Extent = R

If IsNumeric(rstRecords("wgs84ED")) And IsNumeric(rstRecords("wgs84EM")) And IsNumeric(rstRecords("wgs84ES")) And IsNumeric(rstRecords("wgs84ND")) And IsNumeric(rstRecords("wgs84NM")) And IsNumeric(rstRecords("wgs84NS")) Then
           latitude = rstRecords("wgs84ED") + (rstRecords("wgs84EM") / 60) + (rstRecords("wgs84ES") / 3600)
           longitude = rstRecords("wgs84ND") + (rstRecords("wgs84NM") / 60) + (rstRecords("wgs84NS") / 3600)
            ptx.y = latitude
            ptx.x = longitude
            Me.MyMap.CenterAt ptx.x, ptx.y
            
DoEvents
End If

Me.StatusBar1.Panels(3).Text = MyMap.Extent.Height
End Sub

Private Sub CenterViewBM()

Dim ptx As New MapObjects.Point
Dim latitude As Double
Dim longitude As Double
Dim R

Me.MyMap2.Extent = Me.MyMap2.FullExtent
Set R = MyMap.Extent
        R.ScaleRectangle 0.0002
        MyMap2.Extent = R

If IsNumeric(Me.TextBoxLatitude) Then
           
            ptx.y = Me.TextBoxLatitude
            ptx.x = Me.TextBoxLongitude
            Me.MyMap2.CenterAt ptx.x, ptx.y
            
DoEvents
End If


End Sub
Private Sub TabVisibleFalse()
    Me.SSTab1.TabVisible(0) = False
    Me.SSTab1.TabVisible(1) = False
    Me.SSTab1.TabVisible(2) = False
    Me.SSTab1.TabVisible(3) = False
    Me.SSTab1.TabVisible(4) = False
    Me.SSTab1.TabVisible(5) = False
    Me.SSTab1.TabVisible(6) = False
    Me.SSTab1.TabVisible(7) = False
    Me.SSTab1.TabVisible(8) = False
End Sub


Public Sub ZeroModeBM()
    Me.RaveSearchBM.Enabled = 0
    Me.RaveAddBM.Enabled = 1
    Me.RaveEditBM.Enabled = 0
    Me.RaveDeleteBM.Enabled = 0
    Me.RaveSaveBM.Enabled = 0
    Me.RaveCancelBM.Enabled = 0
    Me.RaveNextBM.Enabled = 0
    Me.RaveBackBM.Enabled = 0
    Me.RavePrintBM.Enabled = 0
    Me.RaveMapBM.Enabled = 0
End Sub

Private Sub BrowseModeBM()
    Me.RaveSearchBM.Enabled = 1
    Me.RaveAddBM.Enabled = 1
    Me.RaveEditBM.Enabled = 1
    Me.RaveDeleteBM.Enabled = 1
    Me.RaveSaveBM.Enabled = 0
    Me.RaveCancelBM.Enabled = 0
    Me.RaveNextBM.Enabled = 1
    Me.RaveBackBM.Enabled = 1
    Me.RavePrintBM.Enabled = 1
    Me.RaveMapBM.Enabled = 1
End Sub

Private Sub AddEditModeBM()
    Me.RaveSearchBM.Enabled = 0
    Me.RaveAddBM.Enabled = 0
    Me.RaveEditBM.Enabled = 0
    Me.RaveDeleteBM.Enabled = 0
    Me.RaveSaveBM.Enabled = 1
    Me.RaveCancelBM.Enabled = 1
    Me.RaveNextBM.Enabled = 0
    Me.RaveBackBM.Enabled = 0
    Me.RavePrintBM.Enabled = 0
     Me.RaveMapBM.Enabled = 0
End Sub

Public Sub FirstModeBM()
    Me.RaveSearchBM.Enabled = 1
    Me.RaveAddBM.Enabled = 1
    Me.RaveEditBM.Enabled = 1
    Me.RaveDeleteBM.Enabled = 1
    Me.RaveSaveBM.Enabled = 0
    Me.RaveCancelBM.Enabled = 0
    Me.RaveNextBM.Enabled = 1
    Me.RaveBackBM.Enabled = 0
    Me.RavePrintBM.Enabled = 1
    Me.RaveMapBM.Enabled = 1
End Sub

Public Sub SingleModeBM()
    Me.RaveSearchBM.Enabled = 1
    Me.RaveAddBM.Enabled = 1
    Me.RaveEditBM.Enabled = 1
    Me.RaveDeleteBM.Enabled = 1
    Me.RaveSaveBM.Enabled = 0
    Me.RaveCancelBM.Enabled = 0
    Me.RaveNextBM.Enabled = 0
    Me.RaveBackBM.Enabled = 0
    Me.RavePrintBM.Enabled = 1
    Me.RaveMapBM.Enabled = 1
End Sub

Private Sub LastModeBM()
    Me.RaveSearchBM.Enabled = 1
    Me.RaveAddBM.Enabled = 1
    Me.RaveEditBM.Enabled = 1
    Me.RaveDeleteBM.Enabled = 1
    Me.RaveSaveBM.Enabled = 0
    Me.RaveCancelBM.Enabled = 0
    Me.RaveNextBM.Enabled = 0
    Me.RaveBackBM.Enabled = 1
    Me.RavePrintBM.Enabled = 1
    Me.RaveMapBM.Enabled = 1
End Sub

Private Sub TxtZone_KeyPress(KeyAscii As MSForms.ReturnInteger)
If (KeyAscii) = 8 Then
    Exit Sub
End If


    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub


'Private Sub InsertGCP()
'
'       Dim cmd As New ADODB.Command
'
'       cmd.ActiveConnection = cnn
'       cmd.CommandType = adCmdText
'
'                                                   '1      '2     '3        '4        '5       '6         '7              '8           '9        '10       '11     '12   '13   '14  '15    '16    '17    '18    '19       '20         '21        '22    '23        '24    '25     '26       '27        '28        '29      '30     '31     '32     '33     '34     '35      '36    '37        '38       '39      '40    '41      '42   '43
'       cmd.CommandText = "insert into geoprov (stat_name,region,province,municipal,barangay,date_est,date_est_month,date_est_year,date_est_day,date_las_r,island,d_lat,m_lat,s_lat,d_long,m_long,s_long,h_date_ety,h_ref,hor_authty,h_order,h_date_com,h_fix,ell_hgt,mark_stat,mark_type,mark_const,authority,wgs84ND,wgs84NM,wgs84NS,wgs84ED,wgs84EM,wgs84ES,ellipz,description,latitude,longitude,status,northing,easting,zone) " & _
'                         "values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
'
'
'       cmd.Parameters.Append cmd.CreateParameter("?P1", adVarChar, adParamInput, 255, IIf(Trim(Me.TxtName) = "", Trim(Me.TxtName), Null)) 'stat_name
'       cmd.Parameters.Append cmd.CreateParameter("?P2", adVarChar, adParamInput, 255, IIf(Trim(Me.TxtRegion) = "", Trim(Me.Region), Null)) 'region
'       cmd.Parameters.Append cmd.CreateParameter("?P3", adVarChar, adParamInput, 255, IIf(Trim(Me.TxtProvince) = "", Trim(Me.TxtProvince), Null)) 'province
'       cmd.Parameters.Append cmd.CreateParameter("?P4", adVarChar, adParamInput, 255, IIf(Trim(Me.TxtMunicipality) = "", Trim(Me.TxtMunicipality), Null)) 'municipality
'       cmd.Parameters.Append cmd.CreateParameter("?P5", adVarChar, adParamInput, 255, IIf(Trim(Me.TxtBarangay) = "", Trim(Me.TxtBarangay), Null)) 'barangay
'       cmd.Parameters.Append cmd.CreateParameter("P6", adDate, adParamInput, , IIf(IsDate(Me.txtEstablished), CDate(Me.txtEstablished), Null)) 'Date Established
'       cmd.Parameters.Append cmd.CreateParameter("P7", adInteger, adParamInput, , IIf(Me.ComboBoxMonth.Text <> "", Me.ComboBoxMonth.ListIndex, Null)) 'Date Month
'       cmd.Parameters.Append cmd.CreateParameter("P8", adInteger, adParamInput, , IIf(isyear(Me.txtYear), Val(Me.txtYear), Null)) 'Date Year
'       cmd.Parameters.Append cmd.CreateParameter("P9", adInteger, adParamInput, , IIf(Me.ComboBoxDay.Text <> "", Me.ComboBoxMonth.ListIndex, Null)) 'Date Day
'       cmd.Parameters.Append cmd.CreateParameter("P10", adDate, adParamInput, , IIf(IsDate(Me.txtLastRecover), CDate(Me.txtLastRecover), Null)) 'Date Last Recovered
'       cmd.Parameters.Append cmd.CreateParameter("P11", adVarChar, adParamInput, 50, IIf(Trim(Me.Island) = "", Trim(Me.TxtIsland), Null)) 'Island
'       cmd.Parameters.Append cmd.CreateParameter("P12", adInteger, adParamInput, , IIf(IsNumeric(Me.TxtDLat), Val(Me.TxtDLat), Null)) 'Latitude Degree
'       cmd.Parameters.Append cmd.CreateParameter("P13", adInteger, adParamInput, , IIf(IsNumeric(Me.TxtMLat), Val(Me.TxtMLat), Null)) 'Latitude Minutes
'       cmd.Parameters.Append cmd.CreateParameter("P14", adDouble, adParamInput, , IIf(IsNumeric(Me.TxtSLat), CDec(Me.TxtSLat), Null)) 'Latitude Seconds
'       cmd.Parameters.Append cmd.CreateParameter("P15", adInteger, adParamInput, , IIf(IsNumeric(Me.TxtDLong), Val(Me.TxtDLong), Null)) 'Longitude Degree
'       cmd.Parameters.Append cmd.CreateParameter("P16", adInteger, adParamInput, , IIf(IsNumeric(Me.TxtMLong), Val(Me.TxtMLong), Null)) 'Longitude Minutes
'       cmd.Parameters.Append cmd.CreateParameter("P17", adDouble, adParamInput, , IIf(IsNumeric(Me.TxtSLong), CDec(Me.TxtSLong), Null)) 'Longitude Seconds
'
'       cmd.Parameters.Append cmd.CreateParameter("P18", adDate, adParamInput, , IIf(IsDate(Me.TxtDateEntry), CDate(Me.TxtDateEntry), Null)) 'Date of Entry
'       cmd.Parameters.Append cmd.CreateParameter("P19", adVarChar, adParamInput, 50, "PRS92") 'Reference
'       cmd.Parameters.Append cmd.CreateParameter("P20", adVarChar, adParamInput, 50, IIf(Trim(Me.TxtEstablishedBy) = "", Trim(Me.TxtEstablishedBy), Null)) 'Established By
'
'       cmd.Parameters.Append cmd.CreateParameter("P21", adInteger, adParamInput, , IIf(Me.txtOrder.Text <> "", Me.txtOrder.ListIndex, Null)) 'Order
'       cmd.Parameters.Append cmd.CreateParameter("P22", adDate, adParamInput, , IIf(IsDate(buf(19)), buf(19), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P23", adInteger, adParamInput, , IIf(IsNumeric(buf(20)), buf(20), Null))
'
'       cmd.Parameters.Append cmd.CreateParameter("P24", adDouble, adParamInput, , IIf(IsNumeric(buf(21)), buf(21), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P25", adInteger, adParamInput, , IIf(IsNumeric(buf(22)), buf(22), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P26", adInteger, adParamInput, , IIf(IsNumeric(buf(23)), buf(23), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P27", adInteger, adParamInput, , IIf(IsNumeric(buf(24)), buf(24), Null))
'
'       cmd.Parameters.Append cmd.CreateParameter("P28", adVarChar, adParamInput, 50, buf(25))
'
'       cmd.Parameters.Append cmd.CreateParameter("P29", adInteger, adParamInput, , IIf(IsNumeric(buf(26)), buf(26), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P30", adInteger, adParamInput, , IIf(IsNumeric(buf(27)), buf(27), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P31", adDouble, adParamInput, , IIf(IsNumeric(buf(28)), buf(28), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P32", adInteger, adParamInput, , IIf(IsNumeric(buf(29)), buf(29), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P33", adInteger, adParamInput, , IIf(IsNumeric(buf(30)), buf(30), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P34", adDouble, adParamInput, , IIf(IsNumeric(buf(31)), buf(31), Null))
'
'       cmd.Parameters.Append cmd.CreateParameter("P35", adDouble, adParamInput, , IIf(IsNumeric(buf(32)), buf(32), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P36", adVarWChar, adParamInput, 1073741823, buf(33))
'       cmd.Parameters.Append cmd.CreateParameter("P37", adDouble, adParamInput, , IIf(IsNumeric(buf(34)), buf(34), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P38", adDouble, adParamInput, , IIf(IsNumeric(buf(35)), buf(35), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P39", adInteger, adParamInput, , IIf(IsNumeric(buf(36)), buf(36), Null))
'
'
'       cmd.Parameters.Append cmd.CreateParameter("P40", adDouble, adParamInput, , IIf(IsNumeric(buf(37)), buf(37), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P41", adDouble, adParamInput, , IIf(IsNumeric(buf(38)), buf(38), Null))
'       cmd.Parameters.Append cmd.CreateParameter("P42", adVarChar, adParamInput, 3, buf(39))
'
'       cmd.Execute , , adExecuteNoRecords
'
'End Sub


'Public Sub LoadAdopters()
   
      '  Dim i As Integer
      '  Dim rst As New ADODB.Recordset
        
       ' rst.Open "Select distinct AdoptedBy from geoprov where Adoptedby is not null order by AdoptedBy", cnn, adOpenStatic
       ' FrmGCPDS.TxtMSL.Clear
        
       
       ' For i = 1 To rst.RecordCount
       '     FrmGCPDS.TxtMSL.AddItem rst!AdoptedBy
      '      rst.MoveNext
       ' Next
    
            
'End Sub

Private Sub SaveGravity()
    
    Dim strSQl As String
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    
    If EditModegravity = False Then
        strSQl = "INSERT INTO gravity(stat_name,region,province,municipal,barangay,h_order,elevation,elevationUnit,latitude,longitude,observedValues,encoder,dateLastUpdated,description) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Else
        strSQl = "Update gravity set stat_name=?,region=?,province=?,municipal=?,barangay=?,h_order=?,elevation=?,elevationUnit=?,latitude=?,longitude=?,observedValues=?,encoder=?,dateLastUpdated=?,description=? WHERE stat_name='" & rstGravity!Stat_Name & "'"
    End If

    With cmd
        .ActiveConnection = cnn
        .CommandType = adCmdText
        .CommandText = strSQl
        .Prepared = True

        cmd.Parameters.Append cmd.CreateParameter("?P1", adVarChar, adParamInput, 255, IIf(Trim(Me.TextBoxGravityName.Text) <> "", Trim(Me.TextBoxGravityName.Text), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P2", adVarChar, adParamInput, 255, IIf(Trim(Me.TextBoxGravityRegion) <> "", Trim(Me.TextBoxGravityRegion), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P3", adVarChar, adParamInput, 255, IIf(Trim(Me.TextBoxGravityProvince) <> "", Trim(Me.TextBoxGravityProvince), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P4", adVarChar, adParamInput, 255, IIf(Trim(Me.TextBoxGravityMunicipality) <> "", Trim(Me.TextBoxGravityMunicipality), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P5", adVarChar, adParamInput, 255, IIf(Trim(Me.TextBoxGravityBarangay) <> "", Trim(Me.TextBoxGravityBarangay), Null))
               
        cmd.Parameters.Append cmd.CreateParameter("?P6", adInteger, adParamInput, , IIf(Trim(Me.ComboBoxOrderGravity.Text) <> "", OrderGravity(Me.ComboBoxOrderGravity.ListIndex + 1), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P7", adDouble, adParamInput, , IIf(Trim(Me.TextBoxGravityElevation.Text) <> "", Trim(Me.TextBoxGravityElevation.Text), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P8", adVarChar, adParamInput, 5, IIf(Trim(Me.RaveButtonsUnits.Caption) <> "", Trim(Me.RaveButtonsUnits.Caption), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P9", adDouble, adParamInput, , IIf(Trim(Me.TextBoxGravityLatitude) <> "", DMStoDD(Trim(Me.TextBoxGravityLatitude)), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P10", adDouble, adParamInput, , IIf(Trim(Me.TextBoxGravityLongitude) <> "", DMStoDD(Trim(Me.TextBoxGravityLongitude)), Null))
        cmd.Parameters.Append cmd.CreateParameter("?P11", adDouble, adParamInput, , IIf(Trim(Me.TextBoxObservedValues.Text) <> "", Trim(Me.TextBoxObservedValues.Text), Null))
       

        cmd.Parameters.Append cmd.CreateParameter("?P12", adVarChar, adParamInput, 255, Encoder)
        cmd.Parameters.Append cmd.CreateParameter("?P13", adDate, adParamInput, , Format(Now, "mm/dd/yyyy"))
        
        cmd.Parameters.Append cmd.CreateParameter("?P14", adVarWChar, adParamInput, 1073741823, IIf(Trim(Me.TextBoxGravityDescription.Text) <> "", Trim(Me.TextBoxGravityDescription.Text), Null))
       
        .Execute , , adCmdText + adExecuteNoRecords
    End With

End Sub

Public Function GravityFieldValidation() As Boolean

        

        If Trim(Me.TextBoxGravityName.Text) = "" Then
                MsgBox "Gravity name is required.", vbCritical, "Gravity"
                Exit Function
        End If
        
        If EditModegravity = True Then
            If IsDuplicateGravity(Trim(Me.TextBoxGravityName.Text)) = True And rstGravity!Stat_Name <> Trim(Me.TextBoxGravityName.Text) Then
                MsgBox "Gravity name already exist", vbCritical, "Gravity"
                Exit Function
            End If
            
            Else
            
             If IsDuplicateGravity(Trim(Me.TextBoxGravityName.Text)) = True Then
                MsgBox "Gravity name already exist", vbCritical, "Gravity"
                Exit Function
            End If
            
        End If
        
        If Trim(Me.TextBoxObservedValues.Text) <> "" And IsNumeric(Me.TextBoxObservedValues.Text) = False Then
                MsgBox "Observed values should be numeric", vbCritical, "Gravity"
                Exit Function
        End If
        
     
        
         If Trim(Me.TextBoxGravityElevation.Text) <> "" And IsNumeric(Me.TextBoxGravityElevation.Text) = False Then
                MsgBox "Elevation should be numeric", vbCritical, "Gravity"
                Exit Function
        End If
        
        If Trim(Me.TextBoxGravityLatitude.Text) <> "" Then
                If isValidCoordinate(Me.TextBoxGravityLatitude.Text) = False Then
                    MsgBox "Invalid latitude", vbCritical, "Gravity"
                Exit Function
                End If
        End If
        
         If Trim(Me.TextBoxGravityLongitude.Text) <> "" Then
                If isValidCoordinate(Me.TextBoxGravityLongitude.Text) = False Then
                    MsgBox "Invalid longitude", vbCritical, "Gravity"
                Exit Function
                End If
        End If
        
        GravityFieldValidation = True
End Function
