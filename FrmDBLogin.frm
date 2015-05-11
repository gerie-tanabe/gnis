VERSION 5.00
Begin VB.Form FrmDBLogin 
   BackColor       =   &H00F9F9F9&
   Caption         =   "Database Connection"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   Icon            =   "FrmDBLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H80000009&
      Caption         =   "Cancel"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton CmdLogin 
      BackColor       =   &H80000009&
      Caption         =   "Open"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      Picture         =   "FrmDBLogin.frx":3452
      ScaleHeight     =   1095
      ScaleWidth      =   5775
      TabIndex        =   2
      Top             =   0
      Width           =   5775
   End
   Begin VB.TextBox TxtPassword 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "="
      TabIndex        =   1
      Text            =   "namria"
      Top             =   1815
      Width           =   2655
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C48546&
      Height          =   405
      Left            =   1440
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   2865
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000040C0&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   10
      Top             =   1100
      Width           =   5655
   End
End
Attribute VB_Name = "FrmDBLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdCancel_Click()
End
End Sub

Private Sub CmdLogin_Click()
On Error GoTo hell

    Set cnn = New ADODB.Connection
    'cnn.ConnectionString = "dsn=gcpds;pwd=" & Trim(Me.TxtPassword.Text)
    cnn.ConnectionString = "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\GCPDS.mdb;pwd=" & Me.TxtPassword.Text
    cnn.Open
    Unload Me
    FrmLogin.Show 1

hell:

Debug.Print Err.Description
   If Err.Description = "[Microsoft][ODBC Microsoft Access Driver] Could not find file '(unknown)'." Then
      MsgBox "GCPDS.mdb Access Database cannot be found. " & vbCrLf & "Please try to locate it ", vbInformation, "Connection Failed."
   End If
   
 
   If Err.Description = "[Microsoft][ODBC Microsoft Access Driver] Not a valid password." Then
      MsgBox "Invalid Database Password.", vbInformation, "Connection Failed"
      Me.TxtPassword.Text = ""
      Me.TxtPassword.SetFocus
   End If
   
   If Err.Description = "[Microsoft][ODBC Microsoft Access Driver] Disk or network error." Then
    MsgBox "Cannot connect to a network drive. ", vbInformation, "Connection Failed."
   End If
End Sub



Private Sub Form_Activate()
Me.TxtPassword.SetFocus
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdLogin_Click
End If
End Sub
