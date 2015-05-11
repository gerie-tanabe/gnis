VERSION 5.00
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "ravebuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLibrary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstMarkType 
      Height          =   5610
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   9895
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Mark Type"
         Object.Width           =   11818
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons RaveEditControl 
      Height          =   420
      Left            =   2925
      TabIndex        =   1
      Top             =   5940
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      BTYPE           =   3
      TX              =   "Edit"
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
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmLibrary.frx":0000
      PICN            =   "FrmLibrary.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveDeleteControl 
      Height          =   420
      Left            =   4680
      TabIndex        =   2
      Top             =   5940
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      BTYPE           =   3
      TX              =   "Delete"
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
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmLibrary.frx":02A6
      PICN            =   "FrmLibrary.frx":02C2
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveAddControl 
      Height          =   420
      Left            =   1170
      TabIndex        =   3
      Top             =   5940
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      BTYPE           =   3
      TX              =   "Add"
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
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmLibrary.frx":052A
      PICN            =   "FrmLibrary.frx":0546
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    RaveDeleteControl_Click
End If
End Sub

Private Sub Form_Load()
    LoadMarkType
    
End Sub

Private Sub RaveAddControl_Click()
    FrmmarkType.Show 1
End Sub


Public Sub LoadMarkType()
Dim i As Integer
    Dim rst As New ADODB.Recordset
    Dim varlist
    
    rst.Open "select * from marktype where MTDesc<>'' order by MTDesc", cnn, adOpenStatic
  Me.LstMarkType.ListItems.Clear
For i = 1 To rst.RecordCount
        Set varlist = Me.LstMarkType.ListItems.Add
            'varlist.Text = IIf(IsNull(rst("MTDesc")), "", rst("MTCode"))
            varlist.subitems(1) = IIf(IsNull(rst("MTDesc")), "", rst("MTDesc"))
            rst.MoveNext
Next
End Sub

Private Sub RaveDeleteControl_Click()
If MsgBox("Are you sure you want to delete this record.", vbYesNo, "Mark Type") = vbNo Then
   Exit Sub
End If


    
    cnn.Execute "delete from marktype where mtdesc='" & Trim(Replace(Me.LstMarkType.SelectedItem.subitems(1), "'", "''")) & "'"
    LoadMarkType
End Sub

Private Sub RaveEditControl_Click()
FrmMarkType2.Show 1
End Sub
