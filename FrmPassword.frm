VERSION 5.00
Begin VB.Form FrmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   Icon            =   "FrmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton CmdPassword 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.TextBox TxtPassword 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "="
         TabIndex        =   0
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtNewPassword 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "="
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1065
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2205
      Left            =   120
      Picture         =   "FrmPassword.frx":3452
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdPassword_Click()
    If Trim(Me.TxtPassword) = Trim(rstUserinfo("password").Value) Then
       rstUserinfo("password").Value = IIf(Trim(Me.txtNewPassword.Text) = "", "", Me.txtNewPassword.Text)
       rstUserinfo.Update
       
       Unload Me
       Else
       MsgBox "Invalid Password", vbInformation
    End If
End Sub


