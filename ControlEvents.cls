VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1



 
Public Sub Init(ByVal c As Control)

    Select Case True
        Case TypeOf c Is MSForms.TextBox
            Set txt = c
        Case TypeOf c Is MSForms.ComboBox
            Set cbo = c
    End Select
   

End Sub

Private Sub cbo_GotFocus()
     If (cbo.BorderStyle = fmBorderStyleSingle) Then
          cbo.BackColor = &HC0FFFF
          cbo.BackStyle = 1
     End If

End Sub

Private Sub cbo_LostFocus()

     If (cbo.BorderStyle = fmBorderStyleSingle) Then
         cbo.BackStyle = 0
     End If

End Sub


Private Sub txt_GotFocus()

     If (txt.BorderStyle = fmBorderStyleSingle) Then
         txt.BackColor = &HC0FFFF
         txt.BackStyle = 1
     End If
      
End Sub

Private Sub txt_LostFocus()

      If (txt.BorderStyle = fmBorderStyleSingle) Then
         txt.BackStyle = 0
      End If
      
End Sub


