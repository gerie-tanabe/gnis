VERSION 5.00
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmMarkType2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mark Type"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons cmdOK 
      Height          =   510
      Left            =   675
      TabIndex        =   1
      ToolTipText     =   "Add Location"
      Top             =   495
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   900
      BTYPE           =   3
      TX              =   "Okay"
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
      MICON           =   "FrmmarkType2.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox TxtMarkType 
      Height          =   390
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   3015
      VariousPropertyBits=   1820346395
      BackColor       =   16777215
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "5318;688"
      BorderColor     =   0
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmMarkType2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()


    cnn.Execute "update marktype set mtdesc='" & Trim(StrConv(Replace(Me.txtMarkType, "'", "''"), vbProperCase)) & "' where mtdesc='" & Trim(StrConv(Replace(FrmLibrary.LstMarkType.SelectedItem.subitems(1), "'", "''"), vbProperCase)) & "'"
    FrmLibrary.LoadMarkType
    LoadMarkType
    Unload Me
End Sub

Public Function IfDuplicateMT(MTDesc As String) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "Select MTDesc from marktype where MTdesc='" & Trim(MTDesc) & "'", cnn, adOpenStatic
    If rst.RecordCount > 0 Then
        IfDuplicateMT = True
        Else
        IfDuplicateMT = False
    End If
End Function



Private Sub Form_Load()
    Me.txtMarkType = FrmLibrary.LstMarkType.SelectedItem.subitems(1)
End Sub

Private Sub TxtMarkType_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    cmdOK_Click
End If
End Sub

