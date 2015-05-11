VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "ravebuttons.ocx"
Begin VB.Form FrmPrintSelection 
   Caption         =   "List of station descriptions to be print"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Selected stations to be print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8445
      Left            =   5250
      TabIndex        =   3
      Top             =   210
      Width           =   5265
      Begin VB.ListBox PrintListBox 
         BackColor       =   &H007D3D1C&
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
         Height          =   7080
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   4965
      End
      Begin Rave_Buttons.RaveButtons cmdOK 
         Height          =   510
         Left            =   1860
         TabIndex        =   5
         ToolTipText     =   "Add Location"
         Top             =   7770
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   900
         BTYPE           =   3
         TX              =   "OK"
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
         BCOL            =   11292960
         BCOLO           =   11292960
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmPrintSelection.frx":0000
         PICN            =   "FrmPrintSelection.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Station Name:"
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
      Index           =   7
      Left            =   210
      TabIndex        =   2
      Top             =   60
      Width           =   2295
   End
   Begin MSForms.ListBox ResultListBox 
      Height          =   6675
      Left            =   180
      TabIndex        =   1
      Top             =   1140
      Width           =   4935
      BackColor       =   8207644
      ForeColor       =   16777215
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "8705;11774"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox SearchTextBox 
      Height          =   465
      Left            =   180
      TabIndex        =   0
      Top             =   450
      Width           =   4935
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "8705;820"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmPrintSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()

End Sub

Private Sub ListBox1_Click()

End Sub



Private Sub cmdOK_Click()
If Me.PrintListBox.ListCount < 1 Then
    Exit Sub
End If
FrmDescriptionList.Show 1
End Sub

Private Sub PrintListBox_DblClick()
 Me.PrintListBox.RemoveItem (Me.PrintListBox.ListIndex)
End Sub

Private Sub ResultListBox_DblClick(Cancel As MSForms.ReturnBoolean)
    Me.PrintListBox.AddItem Me.ResultListBox.List(Me.ResultListBox.ListIndex)
End Sub

Private Sub SearchTextBox_Change()

If Trim(Me.SearchTextBox) = "" Then
   Me.ResultListBox.Clear
   Exit Sub
End If


   Dim rst As New ADODB.Recordset
   Dim i As Integer
       rst.Open "select stat_name from geoprov where stat_name like '%" & Me.SearchTextBox & "%' order by stat_name", cnn, adOpenStatic
       
       If rst.RecordCount > 0 Then
       
           Me.ResultListBox.Clear
           
           For i = 1 To rst.RecordCount
                
                Me.Label(7).Caption = "Search Station Name: " & rst.RecordCount
                Me.ResultListBox.AddItem rst("stat_name")
                rst.MoveNext
           Next
           
           
           Else
       Me.ResultListBox.Clear
       End If
End Sub
