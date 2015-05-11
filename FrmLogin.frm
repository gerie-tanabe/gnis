VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmLogin 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5070
   Icon            =   "FrmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmLogin.frx":014A
   ScaleHeight     =   4140
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveClose 
      Height          =   585
      Left            =   4410
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   1032
      BTYPE           =   11
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   4194304
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmLogin.frx":AA07
      PICN            =   "FrmLogin.frx":AA23
      PICH            =   "FrmLogin.frx":B6FD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveLogin 
      Height          =   465
      Left            =   1665
      TabIndex        =   1
      Top             =   3285
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "Login"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmLogin.frx":C3D7
      PICN            =   "FrmLogin.frx":C3F3
      PICH            =   "FrmLogin.frx":C8B9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   0
      Left            =   450
      TabIndex        =   5
      Top             =   2025
      Width           =   945
   End
   Begin MSForms.TextBox TxtPassword 
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Top             =   2430
      Width           =   2700
      VariousPropertyBits=   1820346395
      BackColor       =   16777215
      ForeColor       =   32768
      BorderStyle     =   1
      Size            =   "4762;688"
      PasswordChar    =   61
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Webdings"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   2
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   1
      Left            =   450
      TabIndex        =   3
      Top             =   2475
      Width           =   1095
   End
   Begin MSForms.ComboBox ComboBoxUsernames 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1980
      Width           =   2715
      VariousPropertyBits=   746604571
      ForeColor       =   0
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4789;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image ImageBorder 
      Height          =   1485
      Left            =   30
      MousePointer    =   15  'Size All
      Top             =   0
      Width           =   5085
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdCancel_Click()
    End
End Sub

Private Sub CheckUser()

  
    Dim rst As New ADODB.Recordset
        
    rst.Open "select * from useraccounts where username='" & Trim(Replace(Me.ComboBoxUsernames, "'", "''")) & "' and password='" & Trim(Replace(Me.TxtPassword, "'", "''")) & "'", cnn, adOpenStatic
    
    If rst.RecordCount > 0 Then
  
       Encoder = Trim(Me.ComboBoxUsernames)
       Password = rst("password")
       AccessLevel = rst("access")
       Unload Me
       FrmGCPDS.Show
       FrmGCPDS.LabelUser.Caption = Encoder
       
       Else
       
      
       MsgBox "Invalid username or password. Please verify your username and password.", vbCritical, "Login"
       Me.ComboBoxUsernames = ""
       Me.TxtPassword = ""
       Me.ComboBoxUsernames.SetFocus
       Exit Sub
    End If
    
End Sub






Private Sub Form_Load()
     PopulateUserNames
End Sub

Private Sub ImageBorder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    TranslucentForm Me, 200
End Sub

Private Sub ImageBorder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngReturnValue As Long
    
    If Button = 1 Then
    Call ReleaseCapture
    
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    TranslucentForm Me, 255
    End If
End Sub

Private Sub ImageBorder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    TranslucentForm Me, 255

End Sub


Private Sub RaveClose_Click()

    End
    
End Sub


Private Sub RaveLogin_Click()
    
    CheckUser
    
End Sub

Private Sub TxtPassword_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode = 13 Then
        Call RaveLogin_Click
    End If
    
End Sub

Public Sub PopulateUserNames()
Dim rst As New ADODB.Recordset
    rst.Open "select * from useraccounts order by username", cnn, adOpenStatic

    For i = 1 To rst.RecordCount
        Me.ComboBoxUsernames.AddItem rst("username")
        rst.MoveNext
    Next
    
    If rst.RecordCount > 0 Then
        FrmLogin.ComboBoxUsernames.ListIndex = 0
    End If
End Sub

