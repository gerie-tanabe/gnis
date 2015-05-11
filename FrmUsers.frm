VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Profile"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   Icon            =   "FrmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Account"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton CmdPassword 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Password"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PathFinder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   825
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   2865
      End
   End
   Begin MSComctlLib.ImageList UsersImageList 
      Left            =   240
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   61
      ImageHeight     =   57
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsers.frx":3452
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsers.frx":6D7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUsers.frx":D017
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "FrmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPassword_Click()
    FrmPassword.Show 1
End Sub

Private Sub Command1_Click()
'    If FrmGCPDS.ListSecurity.ListItems.Count > 1 Then
'        cnn.Execute "delete * from tblusers where username='" & rstUserinfo("username").Value & "'"
'
'        ListUsers
'        Unload Me
'            Else
'            MsgBox "You cannot delete this user.", vbInformation, "Security Login"
'    End If
End Sub



Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
'Set rstUserinfo = New ADODB.Recordset
'rstUserinfo.Open "Select * from tblusers where username='" & FrmGCPDS.ListSecurity.SelectedItem.Text & "'", cnn, adOpenStatic, adLockOptimistic
'    Me.Image1.Picture = Me.UsersImageList.ListImages(rstUserinfo("access").Value).Picture
'    Me.Caption = FrmGCPDS.ListSecurity.SelectedItem.Text
'
'    If rstUserinfo("access").Value = 1 Then
'        Me.Label1.Caption = UCase(rstUserinfo("username").Value) & vbCrLf & "Administrator"
'        ElseIf rstUserinfo("access").Value = 2 Then
'        Me.Label1.Caption = UCase(rstUserinfo("username").Value) & vbCrLf & "Normal User"
'        Else
'        Me.Label1.Caption = UCase(rstUserinfo("username").Value) & vbCrLf & "Guest"
'    End If
End Sub


