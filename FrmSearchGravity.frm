VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmSearchGravity 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Search"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSearchGravity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstResults 
      Height          =   2085
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5200
      _ExtentX        =   9181
      _ExtentY        =   3678
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   4210752
      BackColor       =   16777215
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Results"
         Object.Width           =   8467
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons RaveSearch 
      Height          =   405
      Left            =   4800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   ""
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
      MICON           =   "FrmSearchGravity.frx":0442
      PICN            =   "FrmSearchGravity.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type the station name here then press enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3240
   End
   Begin MSForms.TextBox TxtName 
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      VariousPropertyBits=   1820346395
      BackColor       =   16777215
      ForeColor       =   0
      BorderStyle     =   1
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
End
Attribute VB_Name = "FrmSearchGravity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim varlist
End Sub

Private Sub LstResults_DblClick()
          rstGravity.MoveFirst
          rstGravity.Find "stat_name = '" & Trim(Replace(Me.LstResults.SelectedItem.Text, "'", "''")) & "'", , adSearchForward
        
          If rstGravity.EOF = True Then
             MsgBox "The station you're searching for don't exist.", vbInformation, "Search Result"
             rstGravity.MoveFirst
          End If
          
          FrmGCPDS.SetToolBarStatusGravity
          Unload Me
End Sub

Private Sub RaveSearch_Click()
    If Trim(Me.TxtName) = "" Then
        Exit Sub
    End If
        
    Dim rst As New ADODB.Recordset
    Dim i As Integer
        rst.Open "Select stat_name from gravity where stat_name like '%" & Trim(Replace(Me.TxtName, "'", "''")) & "%'", cnn, adOpenStatic, adLockOptimistic
    

    
    If rst.RecordCount > 0 Then
       
       If rst.RecordCount = 1 Then
          rstGravity.MoveFirst
          rstGravity.Find "stat_name like '%" & Trim(Replace(Me.TxtName, "'", "''")) & "%'", , adSearchForward
          
          If rstGravity.EOF = True Then
             MsgBox "The station you're searching for don't exist.", vbInformation, "Gravity"
             rstGravity.MoveFirst
          End If
          
          
          FrmGCPDS.SetToolBarStatusGravity
          Unload Me
          Else
          
           Me.LstResults.ListItems.Clear
          
        
          For i = 1 To rst.RecordCount
            Set varlist = Me.LstResults.ListItems.Add
            varlist.Text = rst("Stat_name")
            rst.MoveNext
          Next
          
       End If
 
        Else
        
        MsgBox "No Records", vbInformation, "GNIS"
        
    End If
End Sub

Private Sub TxtName_Change()
    Me.TxtName = UCase(Me.TxtName)
    Me.LstResults.ListItems.Clear
End Sub

Private Sub TxtName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    RaveSearch_Click
End If

If KeyCode = 27 Then
    Unload Me
   End If
End Sub
