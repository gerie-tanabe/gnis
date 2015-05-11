VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStation 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Station"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   10950
   Icon            =   "FrmStation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      Picture         =   "FrmStation.frx":3452
      ScaleHeight     =   1455
      ScaleWidth      =   10815
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "National Mapping and Resource Information Authority"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3840
         TabIndex        =   2
         Top             =   720
         Width           =   3315
      End
   End
   Begin MSComctlLib.ListView LstStation 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "No."
         Object.Width           =   1329
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Station Name"
         Object.Width           =   7144
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Municipality"
         Object.Width           =   4678
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Barangay"
         Object.Width           =   3969
      EndProperty
   End
End
Attribute VB_Name = "FrmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varlist As Variant


Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()

    Station = Me.LstStation.SelectedItem.SubItems(2)
    rstRecords.MoveFirst
    rstRecords.Find "Stat_name='" & Replace(Station, "'", "''") & "'", , adSearchForward
    FrmGCPDS.FormCaption.Caption = Format(rstRecords.AbsolutePosition, "#,##0") & " of " & Format(rstRecords.RecordCount, "#,##0") & " Records"
    FillUp
    Unload Me

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

Set rstStation = New ADODB.Recordset
    rstStation.Open "Select stat_new,stat_name,Municipal,Barangay from geoprov where province=" & "'" & ProvinceCode & "' order by stat_name", cnn, adOpenStatic
      For x = 1 To rstStation.RecordCount
       
       Set varlist = Me.LstStation.ListItems.Add()
                 varlist.SubItems(1) = IIf(IsNull(rstStation.Fields(0).Value) = False, rstStation.Fields(0).Value, "")
                 varlist.SubItems(2) = IIf(IsNull(rstStation.Fields(1).Value) = False, rstStation.Fields(1).Value, "")
                 varlist.SubItems(3) = IIf(IsNull(rstStation.Fields(2).Value) = False, rstStation.Fields(2).Value, "")
                 varlist.SubItems(4) = IIf(IsNull(rstStation.Fields(3).Value) = False, rstStation.Fields(3).Value, "")
                 
          rstStation.MoveNext
      Next
  
      
End Sub

Private Sub LstStation_DblClick()
        Call CmdOk_Click
End Sub

Private Sub LstStation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CmdOk_Click
    End If
End Sub

