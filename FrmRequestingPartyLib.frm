VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmRequestingPartyLib 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requesting Party Library"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstRequestingParty 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   4210752
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
         Text            =   "Name"
         Object.Width           =   7056
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons CmdAdd 
      Height          =   450
      Left            =   840
      TabIndex        =   1
      Top             =   5880
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   794
      BTYPE           =   4
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
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRequestingPartyLib.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons CmdEdit 
      Height          =   450
      Left            =   2280
      TabIndex        =   2
      Top             =   5880
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   4
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
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRequestingPartyLib.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons CmdDelete 
      Height          =   450
      Left            =   3600
      TabIndex        =   3
      Top             =   5880
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   4
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
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRequestingPartyLib.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmRequestingPartyLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()
FrmAddParty.Show 1
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()
If Me.LstRequestingParty.ListItems.Count > 0 Then
    cnn.Execute "Delete from Requesting_Party_Lib Where Requesting_party='" & Replace(Trim(Me.LstRequestingParty.SelectedItem.Text), "'", "''") & "'"
    Me.LstRequestingParty.ListItems.Remove (Me.LstRequestingParty.SelectedItem.Index)
    If Me.LstRequestingParty.ListItems.Count > 0 Then
        Me.LstRequestingParty.ListItems(Me.LstRequestingParty.ListItems.Count).Selected = True
    End If
End If
End Sub

Private Sub CmdEdit_Click()
If Me.LstRequestingParty.ListItems.Count > 0 Then
    FrmEditParty.Show 1
End If
End Sub

Private Sub Form_Activate()
Dim rst As New ADODB.Recordset
rst.Open "Select * from Requesting_Party_Lib order by Requesting_Party", cnn, adOpenStatic, adLockOptimistic
Me.LstRequestingParty.ListItems.Clear

If rst.RecordCount > 0 Then
Dim i As Integer
Dim varlist
    For i = 1 To rst.RecordCount
        Set varlist = Me.LstRequestingParty.ListItems.Add
        varlist.Text = rst("Requesting_Party").Value
        rst.MoveNext
    Next
    rst.Close
End If
End Sub

Private Sub RaveButtons1_Click()

End Sub
