VERSION 5.00
Begin VB.Form FrmNewUser 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New User"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox cmbRights 
      Height          =   315
      ItemData        =   "FrmNewUser.frx":0000
      Left            =   1680
      List            =   "FrmNewUser.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox TxtPassword 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "="
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Access"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Password"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   750
   End
End
Attribute VB_Name = "FrmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
Me.Hide
End Sub

Private Sub CmdOK_Click()
Dim rstTmpUser As New ADODB.Recordset

If Trim(Me.txtUsername.Text) = "" Then
     MsgBox "Username should not be blank.", vbInformation, "Username"
     Exit Sub
End If

rstTmpUser.Open "Select * from tblusers where username='" & Trim(Me.txtUsername.Text) & "'", cnn, adOpenStatic, adLockOptimistic
If rstTmpUser.RecordCount = 0 Then
 cnn.Execute "Insert into tblusers (username,password,access) values('" & Me.txtUsername & "','" & Me.TxtPassword & "','" & Me.cmbRights.ListIndex + 1 & "')"
 ListUsers
 Unload Me
 Else
 MsgBox "Username already exist.", vbInformation, "Duplicate"
 Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.cmbRights.ListIndex = 0
End Sub
