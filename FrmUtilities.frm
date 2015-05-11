VERSION 5.00
Begin VB.Form FrmUtilities 
   Caption         =   "Form1"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   " PRS92 -> WGS84 (Apply to all gcps)"
      Height          =   735
      Left            =   270
      TabIndex        =   4
      Top             =   180
      Width           =   4065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dbase 4 to Microsoft Access Convertion"
      Height          =   7335
      Left            =   285
      TabIndex        =   0
      Top             =   1215
      Width           =   4095
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.DirListBox DbaseDir 
         Height          =   5490
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton CmdConvert 
         Caption         =   "Convert"
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   6600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConvert_Click()

If MsgBox("Current Database will be deleted. Do you want to continue.", vbYesNo) = vbYes Then
    FrmDbase4.Show 1
End If

End Sub

Private Sub Command1_Click()
Dim rst As New ADODB.Recordset
rst.Open "select * from geoprov", cnn, adOpenStatic

Dim i As Integer

For i = 1 To rst.RecordCount
DoEvents
Me.Caption = i & "/" & rst.RecordCount
    If rst("h_ref") = "PRS92" And IsNumeric(rst("d_long")) = True And IsNumeric(rst("m_long")) = True And IsNumeric(rst("s_long")) = True And IsNumeric(rst("d_lat")) = True And IsNumeric(rst("m_lat")) = True And IsNumeric(rst("s_lat")) = True And IsNumeric(rst("ell_hgt")) = True And rst("ell_hgt") > 0 Then
      PRS92_TO_WGS84_X rst("d_long"), rst("m_long"), rst("s_long"), rst("d_lat"), rst("m_lat"), rst("s_lat"), rst("ell_hgt"), rst("stat_name")
    End If
    rst.MoveNext
Next

MsgBox "finish"
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Me.DbaseDir.Path = Me.Drive1.Drive
End Sub


