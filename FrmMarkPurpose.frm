VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "Ravebuttons.ocx"
Begin VB.Form FrmMarkPurpose 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mark Purpose"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons cmdOK 
      Height          =   510
      Left            =   675
      TabIndex        =   1
      ToolTipText     =   "Add Location"
      Top             =   720
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
      MICON           =   "FrmMarkPurpose.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox TxtMarkPurpose 
      Height          =   390
      Left            =   165
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      VariousPropertyBits=   1820346395
      BackColor       =   16777215
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "4895;688"
      BorderColor     =   0
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmMarkPurpose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

If IfDuplicateMP(Trim(Me.TxtMarkPurpose)) Then
    MsgBox "Mark purpose already exist in the list.", vbInformation, "Mark Purpose"
    Exit Sub
End If

    Dim rst As New ADODB.Recordset
    rst.Open "select max(mCode) from Markpur", cnn, adOpenStatic
    cnn.Execute "insert into markpur (mcode,mdesc) values(" & IIf(IsNull(rst(0)), 0, rst(0)) + 1 & ",'" & Trim(StrConv(Replace(Me.TxtMarkPurpose, "'", "''"), vbProperCase)) & "')"
    
    FrmLibraryMarkPurpose.LoadMarkPurposeX
    LoadMarkPurpose
    
    Unload Me
End Sub

Public Function IfDuplicateMP(MTDesc As String) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "Select MDesc from markpur where Mdesc='" & Trim(MDesc) & "'", cnn, adOpenStatic
    If rst.RecordCount > 0 Then
        IfDuplicateMP = True
        Else
        IfDuplicateMP = False
    End If
End Function

Private Sub TxtMarkPurpose_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    cmdOK_Click
End If
End Sub

