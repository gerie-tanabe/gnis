VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPurpose 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purpose of Request Library"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   975
   End
   Begin MSComctlLib.ListView LstPurpose 
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   5592405
      BackColor       =   16382457
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1170
      TabIndex        =   5
      Top             =   90
      Width           =   2025
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   240
      Top             =   0
      Width           =   4005
   End
End
Attribute VB_Name = "FrmPurpose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()
FrmAddPurpose.Show 1
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()
If Me.LstPurpose.ListItems.Count > 0 Then
    cnn.Execute "Delete from Purpose_Lib Where Purpose='" & Replace(Trim(Me.LstPurpose.SelectedItem.Text), "'", "''") & "'"
    Me.LstPurpose.ListItems.Remove (Me.LstPurpose.SelectedItem.Index)
    If Me.LstPurpose.ListItems.Count > 0 Then
        Me.LstPurpose.ListItems(Me.LstPurpose.ListItems.Count).Selected = True
    End If
End If
End Sub

Private Sub CmdEdit_Click()
If Me.LstPurpose.ListItems.Count > 0 Then
    FrmEditPurpose.Show 1
End If
End Sub

Private Sub Form_Activate()
Dim rst As New ADODB.Recordset
rst.Open "Select * from Purpose_Lib order by Purpose", cnn, adOpenStatic, adLockOptimistic
Me.LstPurpose.ListItems.Clear

If rst.RecordCount > 0 Then
Dim i As Integer
Dim varlist
    For i = 1 To rst.RecordCount
        Set varlist = Me.LstPurpose.ListItems.Add
        varlist.Text = StrConv(rst("Purpose").Value, vbProperCase)
        rst.MoveNext
    Next
    rst.Close
End If
End Sub
