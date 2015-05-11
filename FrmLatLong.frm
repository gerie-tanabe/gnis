VERSION 5.00
Begin VB.Form FrmLatLong 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Latitude and Longitude"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "FrmLatLong.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cmbDLong 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "FrmLatLong.frx":0442
      Left            =   960
      List            =   "FrmLatLong.frx":0476
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cmbMLong 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "FrmLatLong.frx":04CA
      Left            =   2040
      List            =   "FrmLatLong.frx":04CC
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtSLong 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtSLat 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cmbMLat 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "FrmLatLong.frx":04CE
      Left            =   2040
      List            =   "FrmLatLong.frx":04D0
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.ComboBox cmbDLat 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "FrmLatLong.frx":04D2
      Left            =   960
      List            =   "FrmLatLong.frx":0518
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Longitude"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   840
      Width           =   150
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   2760
      TabIndex        =   11
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   3960
      TabIndex        =   10
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   3960
      TabIndex        =   6
      Top             =   360
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   150
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latitude"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   570
   End
End
Attribute VB_Name = "FrmLatLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbDLat_Click()
If IsNumeric(Trim(cmbDLat.Text)) = True Then
rstProvinceInfo("d_lat") = CDec(Trim(cmbDLat.Text))
End If
End Sub

Private Sub cmbDLong_Click()
If IsNumeric(Trim(cmbDLong.Text)) = True Then
rstProvinceInfo("d_long") = CDec(Trim(cmbDLong.Text))
End If
End Sub

Private Sub cmbMLat_Click()
If IsNumeric(Trim(cmbMLat.Text)) = True Then
    rstProvinceInfo("m_lat") = CDec(Trim(cmbMLat.Text))
End If
End Sub

Private Sub cmbMLong_Click()
If IsNumeric(Trim(cmbMLong.Text)) = True Then
    rstProvinceInfo("m_long") = CDec(Trim(cmbMLong.Text))
End If
End Sub

Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub CmdOk_Click()
If IsNumeric(Me.cmbDLat.Text) = False Then
   MsgBox ("Invalid Entry"), vbInformation, "Latitude - Degree"
   Exit Sub
End If
If IsNumeric(Me.cmbMLat.Text) = False Then
   MsgBox ("Invalid Entry"), vbInformation, "Latitude - Minutes"
   Exit Sub
End If
If IsNumeric(Me.txtSLat.Text) = False Then
   MsgBox ("Invalid Entry"), vbInformation, "Latitude - Seconds"
   Exit Sub
End If
If IsNumeric(Me.cmbDLong.Text) = False Then
   MsgBox ("Invalid Entry"), vbInformation, "Longitude - Degree"
   Exit Sub
End If
If IsNumeric(Me.cmbMLong.Text) = False Then
   MsgBox ("Invalid Entry"), vbInformation, "Longitude - Minutes"
   Exit Sub
End If
If IsNumeric(Me.txtSLong.Text) = False Then
   MsgBox ("Invalid Entry"), vbInformation, "Longitude - Seconds"
   Exit Sub
End If


FrmGCPDS.TxtLatitude = Me.cmbDLat & "° " & Me.cmbMLat & "' " & Me.txtSLat & """"
FrmGCPDS.TxtLongitude = Me.cmbDLong & "° " & Me.cmbMLong & "' " & Me.txtSLong & """"
rstProvinceInfo("Latitude").Value = Val(Me.cmbDLat.Text) + Val(Me.cmbMLat.Text) / 60 + CDec(Me.txtSLat) / 3600
rstProvinceInfo("Longitude").Value = Val(Me.cmbDLong.Text) + Val(Me.cmbMLong.Text) / 60 + CDec(Me.txtSLong) / 3600

Call Compute(Me.cmbDLat, Me.cmbMLat, Me.txtSLat, Me.cmbDLong, Me.cmbMLong, Me.txtSLong)
            FrmGCPDS.TxtEasting = East
            FrmGCPDS.TxtNorthing = North
            FrmGCPDS.TxtZone = Zone
            
Me.Hide
End Sub



Private Sub Form_Activate()


'If IsNull(rstProvinceInfo.Fields(3).Value) = False And IsEmpty(rstProvinceInfo.Fields(3).Value) = False And rstProvinceInfo.Fields(3).Value <> 0 Then
'   Me.cmbDLat.Text = rstProvinceInfo.Fields(3).Value
'Else
'   Me.cmbDLat.ListIndex = -1
'End If
'
'If IsNull(rstProvinceInfo.Fields(4).Value) = False And IsEmpty(rstProvinceInfo.Fields(4).Value) = False And rstProvinceInfo.Fields(4).Value <> 0 Then
'   Me.cmbMLat.Text = rstProvinceInfo.Fields(4).Value
'Else
'   Me.cmbMLat.ListIndex = -1
'End If
'
'If IsNull(rstProvinceInfo.Fields(5).Value) = False Then
'   Me.txtSLat.Text = rstProvinceInfo.Fields(5).Value
'Else
'   Me.txtSLat.Text = ""
'End If
'
'If IsNull(rstProvinceInfo.Fields(6).Value) = False And IsEmpty(rstProvinceInfo.Fields(6).Value) = False And rstProvinceInfo.Fields(6).Value <> 0 Then
'   Me.cmbDLong.Text = rstProvinceInfo.Fields(6).Value
'Else
'   Me.cmbDLong.ListIndex = -1
'End If
'
'If IsNull(rstProvinceInfo.Fields(7).Value) = False And IsEmpty(rstProvinceInfo.Fields(7).Value) = False And rstProvinceInfo.Fields(7).Value <> 0 Then
'   Me.cmbMLong.Text = rstProvinceInfo.Fields(7).Value
'Else
'   Me.cmbMLong.ListIndex = -1
'End If
'
'If IsNull(rstProvinceInfo.Fields(8).Value) = False Then
'   Me.txtSLong.Text = rstProvinceInfo.Fields(8).Value
'Else
'   Me.txtSLong.Text = ""
'End If
'

End Sub

Private Sub Form_Load()
Dim i As Integer
Me.cmbMLat.Clear
Me.cmbMLong.Clear
    For i = 1 To 59
        Me.cmbMLat.AddItem i
        Me.cmbMLong.AddItem i
    Next
End Sub

Private Sub txtSLat_Change()
If IsNumeric(Trim(txtSLat.Text)) = True Then
    rstProvinceInfo("s_lat") = CDec(Trim(txtSLat.Text))
End If
End Sub

Private Sub txtSLat_KeyPress(KeyAscii As Integer)
If (KeyAscii) = 8 Then
        Exit Sub
    End If
    
    If (KeyAscii) = 46 Then
       If InStr(1, Me.txtSLat, ".") = 0 And Trim(Me.txtSLat) <> "" Then
               Exit Sub
          Else
          KeyAscii = 0
          Exit Sub
       End If
    End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtSLong_Change()
If IsNumeric(Trim(txtSLong.Text)) = True Then
    rstProvinceInfo("S_long") = Trim(txtSLong.Text)
End If
End Sub

Private Sub txtSLong_KeyPress(KeyAscii As Integer)
If (KeyAscii) = 8 Then
        Exit Sub
    End If
    
    If (KeyAscii) = 46 Then
       If InStr(1, Me.txtSLong, ".") = 0 And Trim(Me.txtSLong) <> "" Then
               Exit Sub
          Else
          KeyAscii = 0
          Exit Sub
       End If
    End If
    
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub
