VERSION 5.00
Begin VB.Form FrmDialog 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   4530
   Icon            =   "FrmDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleMode       =   0  'User
   ScaleWidth      =   4631.903
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label LblSave 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   240
      Picture         =   "FrmDialog.frx":3452
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim prevX As Single, prevY As Single
'Dim xfader As Byte
'Dim fader As Integer
'
'Private Sub Form_dblClick()
'Me.Hide
'End Sub
'
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    prevX = X
'    prevY = Y
'
'End Sub
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If Button = 1 Then
'    Move Left - (prevX - X), Top - (prevY - Y)
'
'    fader = fader + 1
'
'
'If fader < 101 Then
'xfader = 200 - fader
'
'  Call Make_Transparent(Me.hWnd, xfader)
'
'End If
'End If
'
'End Sub
'
'
'
'Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim n As Byte
'For n = 1 To 100
'    Call Make_Transparent(Me.hWnd, 150 + n)
'
'Next
'fader = 0
'End Sub

Private Sub CmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    If EditMode = True Then
        Me.LblSave.Caption = "Record of " & FrmGCPDS.txtName & " was updated."
    End If
    If AddMode = True Then
        Me.LblSave.Caption = "Record of " & FrmGCPDS.txtName & " was created."
    End If
End Sub

