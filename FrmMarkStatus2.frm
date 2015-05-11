VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmMarkStatus2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mark Status - Edit"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons cmdOK 
      Height          =   510
      Left            =   1350
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
      MICON           =   "FrmMarkStatus2.frx":0000
      PICN            =   "FrmMarkStatus2.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox TxtMarkStatus 
      Height          =   390
      Left            =   810
      TabIndex        =   0
      Top             =   120
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
Attribute VB_Name = "FrmMarkStatus2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()


    cnn.Execute "update markStatus set mSdesc='" & Trim(StrConv(Replace(Me.TxtMarkStatus, "'", "''"), vbProperCase)) & "' where msdesc='" & Trim(StrConv(Replace(FrmLibraryStatus.LstMarkStatus.SelectedItem.SubItems(1), "'", "''"), vbProperCase)) & "'"
    FrmLibraryStatus.LoadMarkStatusX
    LoadMarkStatus
    
    Unload Me
End Sub

Public Function IfDuplicateMS(MSDesc As String) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "Select MSDesc from markstatus where MSdesc='" & Trim(MSDesc) & "'", cnn, adOpenStatic
    If rst.RecordCount > 0 Then
        IfDuplicateMS = True
        Else
        IfDuplicateMS = False
    End If
End Function



Private Sub Form_Load()
    Me.TxtMarkStatus = FrmLibraryStatus.LstMarkStatus.SelectedItem.SubItems(1)
End Sub

Private Sub TxtMarkStatus_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    cmdOK_Click
End If
End Sub

