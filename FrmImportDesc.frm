VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form FrmImportDesc 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12945
   ScaleWidth      =   17355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4680
      TabIndex        =   4
      Top             =   9120
      Width           =   11295
      Begin CCRProgressBar6.ccrpProgressBar ccrpProgressBar1 
         Height          =   450
         Left            =   8640
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   794
         AutoCaption     =   1
         BackColor       =   0
         Caption         =   "0%"
         FillColor       =   49152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ReverseFill     =   -1  'True
         Smooth          =   -1  'True
      End
      Begin MSForms.OptionButton OptionButton1 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         BackColor       =   0
         ForeColor       =   16777215
         DisplayStyle    =   5
         Size            =   "3836;873"
         Value           =   "1"
         Caption         =   "Main Information"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.OptionButton OptionButton2 
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   2535
         BackColor       =   0
         ForeColor       =   16777215
         DisplayStyle    =   5
         Size            =   "4471;873"
         Value           =   "0"
         Caption         =   "WGS84 Coordinates"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.OptionButton OptionButton3 
         Height          =   495
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   3615
         BackColor       =   0
         ForeColor       =   16777215
         DisplayStyle    =   5
         Size            =   "6376;873"
         Value           =   "0"
         Caption         =   "PRS92 Coordinates"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.DirListBox DbaseDir 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5490
      Left            =   960
      TabIndex        =   1
      Top             =   3360
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   450
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   3495
   End
   Begin MSComctlLib.ListView LstExcel 
      Height          =   4845
      Left            =   4680
      TabIndex        =   2
      Top             =   4080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8546
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   16384
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   8819
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons ImportRave 
      Height          =   525
      Left            =   7800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   11640
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   926
      BTYPE           =   4
      TX              =   "Import"
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
      BCOL            =   0
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveSelect 
      Height          =   525
      Left            =   12600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   926
      BTYPE           =   4
      TX              =   "Select"
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
      BCOL            =   0
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveDeselect 
      Height          =   525
      Left            =   14280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   926
      BTYPE           =   4
      TX              =   "Deselect"
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
      BCOL            =   0
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveSelectNone 
      Height          =   525
      Left            =   10920
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   926
      BTYPE           =   4
      TX              =   "Select None"
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
      BCOL            =   0
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveSelectAll 
      Height          =   525
      Left            =   9240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   926
      BTYPE           =   4
      TX              =   "Select All"
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
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveClose 
      Height          =   525
      Left            =   16680
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   926
      BTYPE           =   11
      TX              =   "X"
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
      BCOL            =   128
      BCOLO           =   128
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveButtons1 
      Height          =   795
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   17355
      _ExtentX        =   30612
      _ExtentY        =   1402
      BTYPE           =   4
      TX              =   "Extraction Tool"
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
      BCOL            =   128
      BCOLO           =   128
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":00A8
      PICN            =   "FrmImportDesc.frx":00C4
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
      Height          =   675
      Left            =   4680
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2760
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   1191
      BTYPE           =   8
      TX              =   "Excel Files"
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
      BCOL            =   32768
      BCOLO           =   32768
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":1346
      PICN            =   "FrmImportDesc.frx":1362
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
      Height          =   675
      Left            =   4680
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3480
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   1191
      BTYPE           =   8
      TX              =   "Excel Files"
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
      BCOL            =   32768
      BCOLO           =   32768
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":1A5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveButtons4 
      Height          =   675
      Left            =   10440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   1191
      BTYPE           =   8
      TX              =   "Status"
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
      BCOL            =   32768
      BCOLO           =   32768
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmImportDesc.frx":1A78
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   1200
      Top             =   9360
      Width           =   2535
   End
   Begin MSForms.TextBox TxtEstablishedBy 
      Height          =   405
      Left            =   10920
      TabIndex        =   20
      Top             =   10373
      Width           =   5115
      VariousPropertyBits=   1820346387
      BackColor       =   10972206
      ForeColor       =   16777215
      BorderStyle     =   1
      Size            =   "9022;714"
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
      Caption         =   "Established By:"
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
      Left            =   9120
      TabIndex        =   19
      Top             =   10440
      Width           =   1635
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5520
      TabIndex        =   18
      Top             =   10440
      Width           =   690
   End
   Begin MSForms.ComboBox OrderComboBox 
      Height          =   390
      Left            =   6360
      TabIndex        =   17
      Top             =   10380
      Width           =   2295
      VariousPropertyBits=   746604563
      BackColor       =   0
      ForeColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4048;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   5775
      Left            =   840
      Top             =   3240
      Width           =   3735
   End
End
Attribute VB_Name = "FrmImportDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub DbaseDir_Click()
Dim sNextFile As String
Dim varlist

sNextFile = Dir$(Me.DbaseDir.List(Me.DbaseDir.ListIndex) & "\*.xls")

Me.LstExcel.ListItems.Clear
While sNextFile <> ""
    Set varlist = Me.LstExcel.ListItems.Add
    varlist.SubItems(1) = sNextFile
    varlist.SubItems(2) = "Selected"
    sNextFile = Dir$
Wend
End Sub

Private Sub Drive1_Change()
'On Error Resume Next
Me.DbaseDir.Path = Me.Drive1.Drive
DbaseDir_Click
End Sub

Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Load()
TranslucentForm Me, 230
Me.OrderComboBox.AddItem "1st Order"
Me.OrderComboBox.AddItem "2nd Order"
Me.OrderComboBox.AddItem "3rd Order"
Me.OrderComboBox.AddItem "4th Order"
Me.OrderComboBox.ListIndex = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
    
    If Button = 1 Then
    Call ReleaseCapture
    
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    
    End If
End Sub

Private Sub ImportRave_Click()

Dim xlApp As Excel.Application
Dim Wb As Excel.Workbook
Dim ws As Excel.Worksheet
'Dim myImage As Object
Dim n As Integer

Set xlApp = New Excel.Application

    Me.ccrpProgressBar1.Min = 0
    Me.ccrpProgressBar1.Max = 100
    
    For i = 1 To Me.LstExcel.ListItems.Count
    If Me.LstExcel.ListItems(i).SubItems(2) = "Selected" Then
    
        Me.ccrpProgressBar1.Value = (100 / Me.LstExcel.ListItems.Count) * i
        Set Wb = xlApp.Workbooks.Open(DbaseDir.List(DbaseDir.ListIndex) & "\" & Me.LstExcel.ListItems(i).SubItems(1))
        
        Set ws = Wb.Sheets(1)
        
    If Me.OptionButton1 = True Then
    
        If If_Already_Exist(Trim(ws.Cells(3, 13))) = False Then
           cnn.Execute "Insert into geoprov (stat_name) values('" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "')"
        End If
       
        cnn.Execute "update geoprov set description='" & Replace(ws.Cells(16, 2) & vbCrLf & vbCrLf & ws.Cells(20, 2), "'", "''") & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set region='" & GetRegion(Replace(ws.Cells(13, 4), "'", "''")) & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set province='" & Replace(ws.Cells(13, 4), "'", "''") & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set municipal='" & Replace(ws.Cells(12, 11), "'", "''") & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set barangay='" & Replace(ws.Cells(13, 11), "'", "''") & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set island='" & Replace(ws.Cells(12, 4), "'", "''") & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set date_est='" & Month(ws.Cells(27, 12)) & "-01-" & Year(ws.Cells(27, 12)) & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set h_order=" & Me.OrderComboBox.ListIndex + 1 & " where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set mark_stat=1 where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set mark_const=10 where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set Authority='NAMRIA' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set H_date_ety='" & Format(Date, "mm-dd-yyyy") & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set H_Fix=1 where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set hor_authty='" & Trim(Replace(Me.TxtEstablishedBy, "'", "''")) & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        cnn.Execute "update geoprov set encoder='" & Trim(Replace(Encoder, "'", "''")) & "' where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        
        
        If InStr(ws.Cells(16, 2), "copper nail") > 0 Then
            cnn.Execute "update geoprov set mark_type=3 where stat_name='" & Trim(Replace(ws.Cells(3, 13), "'", "''")) & "'"
        End If
        
       
        
        
    End If
    
    If Me.OptionButton2 = True Then
       Dim x As Integer
            
       For x = 1 To ws.UsedRange.Rows.Count
        If If_Already_Exist(Trim(ws.Cells(x, 1))) = False Then
           cnn.Execute "Insert into geoprov (stat_name) values('" & Trim(Replace(ws.Cells(x, 1), "'", "''")) & "')"
        End If
        
        cnn.Execute "update geoprov set ellipz=" & ws.Cells(x, 4) & " where stat_name='" & Trim(Replace(ws.Cells(x, 1), "'", "''")) & "'"
        cnn.Execute "update geoprov set h_order=" & Me.OrderComboBox.ListIndex + 1 & " where stat_name='" & Trim(Replace(ws.Cells(x, 1), "'", "''")) & "'"
        InsertWGS84 Trim(Replace(ws.Cells(x, 1), "'", "''")), ws.Cells(x, 3), ws.Cells(x, 2)
        
       Next x
        
    End If
    
    If Me.OptionButton3 = True Then
       Dim y As Integer
            
       For y = 1 To ws.UsedRange.Rows.Count
        If If_Already_Exist(Trim(ws.Cells(y, 1))) = False Then
           cnn.Execute "Insert into geoprov (stat_name) values('" & Trim(Replace(ws.Cells(y, 1), "'", "''")) & "')"
        End If
        
        cnn.Execute "update geoprov set ell_hgt=" & ws.Cells(y, 4) & " where stat_name='" & Trim(Replace(ws.Cells(y, 1), "'", "''")) & "'"
        cnn.Execute "update geoprov set h_ref='PRS92' where stat_name='" & Trim(Replace(ws.Cells(y, 1), "'", "''")) & "'"
        cnn.Execute "update geoprov set h_order=" & Me.OrderComboBox.ListIndex + 1 & " where stat_name='" & Trim(Replace(ws.Cells(y, 1), "'", "''")) & "'"
        InsertPRS92 Trim(Replace(ws.Cells(y, 1), "'", "''")), ws.Cells(y, 2), ws.Cells(y, 3)
        
       Next y
        
        
        
    End If
    
    End If
    Next i
    
  MsgBox "Extraction Finished", vbInformation, "Extraction Tool"
    
End Sub





Private Sub RaveClose_Click()
Unload Me
End Sub

Private Sub RaveDeselect_Click()
    Dim i As Integer
       For i = 1 To Me.LstExcel.ListItems.Count
            If Me.LstExcel.ListItems(i).Selected = True Then
              Me.LstExcel.ListItems(i).SubItems(2) = ""
            End If
       Next
End Sub

Private Sub RaveSelect_Click()
    Dim i As Integer
       For i = 1 To Me.LstExcel.ListItems.Count
            If Me.LstExcel.ListItems(i).Selected = True Then
              Me.LstExcel.ListItems(i).SubItems(2) = "Selected"
            End If
       Next
End Sub

Private Sub RaveSelectAll_Click()
    
       Dim i As Integer
       For i = 1 To Me.LstExcel.ListItems.Count
           Me.LstExcel.ListItems(i).SubItems(2) = "Selected"
       Next
   
End Sub

Private Sub RaveSelectNone_Click()

       Dim i As Integer
       For i = 1 To Me.LstExcel.ListItems.Count
           Me.LstExcel.ListItems(i).SubItems(2) = ""
       Next
   
End Sub


Public Sub InsertWGS84(Station As String, Northing As String, Easting As String)
    
    Dim buf
    Dim degreeN As Integer
    Dim minutesN As Integer
    Dim secondsN As Double
    Dim degreeE As Integer
    Dim minutesE As Integer
    Dim secondsE As Double
        
    buf = Split(Northing, "°")
    degreeN = buf(0)
    buf = Split(buf(1), "'")
    minutesN = buf(0)
    buf = Split(buf(1), """")
    secondsN = buf(0)
    
    buf = Split(Easting, "°")
    degreeE = buf(0)
    buf = Split(buf(1), "'")
    minutesE = buf(0)
    buf = Split(buf(1), """")
    secondsE = buf(0)
    
    cnn.Execute "update geoprov set wgs84ND=" & degreeN & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set wgs84NM=" & minutesN & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set wgs84NS=" & secondsN & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set wgs84ED=" & degreeE & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set wgs84EM=" & minutesE & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set wgs84ES=" & secondsE & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"


    
End Sub

Public Sub InsertPRS92(Station As String, Northing As String, Easting As String)
    
    Dim buf
    Dim degreeN As Integer
    Dim minutesN As Integer
    Dim secondsN As Double
    Dim degreeE As Integer
    Dim minutesE As Integer
    Dim secondsE As Double
        
    buf = Split(Northing, "°")
    degreeN = buf(0)
    buf = Split(buf(1), "'")
    minutesN = buf(0)
    buf = Split(buf(1), """")
    secondsN = buf(0)
    
    buf = Split(Easting, "°")
    degreeE = buf(0)
    buf = Split(buf(1), "'")
    minutesE = buf(0)
    buf = Split(buf(1), """")
    secondsE = buf(0)
    
    cnn.Execute "update geoprov set D_lat=" & degreeN & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set M_lat=" & minutesN & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set S_lat=" & secondsN & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set D_Long=" & degreeE & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set M_Long=" & minutesE & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"
    cnn.Execute "update geoprov set S_Long=" & secondsE & " where stat_name='" & Trim(Replace(Station, "'", "''")) & "'"


    
End Sub

