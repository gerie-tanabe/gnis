VERSION 5.00
Begin VB.Form FrmDialog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleMode       =   0  'User
   ScaleWidth      =   4003.068
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prevX As Single, prevY As Single
Dim xfader As Byte
Dim fader As Integer





Private Sub Form_dblClick()
Me.Hide
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    prevX = X
    prevY = Y
 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
    Move Left - (prevX - X), Top - (prevY - Y)
   
    fader = fader + 1
    

If fader < 101 Then
xfader = 200 - fader
 
  Call Make_Transparent(Me.hWnd, xfader)
  
End If
End If

End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Byte
For n = 1 To 100
    Call Make_Transparent(Me.hWnd, 150 + n)
    
Next
fader = 0
End Sub

