VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txt As TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents cbo As ComboBox
Attribute cbo.VB_VarHelpID = -1



 
Public Sub Init(ByVal c As Control)

    Select Case True
        Case TypeOf c Is TextBox
            Set txt = c
        Case TypeOf c Is ComboBox
            Set cbo = c
    End Select
   

End Sub

Private Sub cbo_GotFocus()

     If (txt.BorderStyle = fmBorderStyleSingle) Then
         txt.BackStyle = 1
      End If

End Sub



Private Sub txt_GotFocus()

      If (txt.BorderStyle = fmBorderStyleSingle) Then
         txt.BackStyle = 1
      End If
      
End Sub


