VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmBrgyLib 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Barangay"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveAddBrgy 
      Height          =   465
      Left            =   2880
      TabIndex        =   4
      Top             =   1710
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmBrgyLib.frx":0000
      PICN            =   "FrmBrgyLib.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveClose 
      Height          =   465
      Left            =   1215
      TabIndex        =   5
      Top             =   1710
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmBrgyLib.frx":10AE
      PICN            =   "FrmBrgyLib.frx":10CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay Name:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   7
      Left            =   240
      TabIndex        =   3
      Top             =   525
      Width           =   1725
   End
   Begin MSForms.TextBox TxtBrgyName 
      Height          =   390
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      VariousPropertyBits=   1820346395
      BackColor       =   16777215
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "5318;688"
      BorderColor     =   0
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TxtBrgyNumber 
      Height          =   390
      Left            =   2400
      TabIndex        =   1
      Top             =   945
      Width           =   3015
      VariousPropertyBits=   1821394971
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   3
      BorderStyle     =   1
      Size            =   "5318;688"
      BorderColor     =   0
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay No:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1395
   End
End
Attribute VB_Name = "FrmBrgyLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If BrgyEditMode = True Then
       Me.TxtBrgyNumber = FrmPSGCLibrary.LstBarangay.SelectedItem.Text
       Me.TxtBrgyName = FrmPSGCLibrary.LstBarangay.SelectedItem.SubItems(1)
    End If
End Sub






Private Sub RaveAddBrgy_Click()
'On Error GoTo hell

If Trim(Me.TxtBrgyName) = "" Then
    MsgBox "Barangay name is required.", vbCritical, "Barangay Name"
    Exit Sub
End If

If IsNumeric(Me.TxtBrgyNumber) = False Then
    MsgBox "Barangay Number should be numeric.", vbCritical, "Barangay Number"
    Exit Sub
End If



If BrgyEditMode = False Then
  
  If IfDuplicate(FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & FrmPSGCLibrary.LstMunicipality.SelectedItem.Text & Format(TxtBrgyNumber, "000")) = False Then
   cnn.Execute "insert into psgc (psgc_cd,name,reg,prov,mun,brgy) values('" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & FrmPSGCLibrary.LstMunicipality.SelectedItem.Text & Format(TxtBrgyNumber, "000") & "'," & "'" & Replace(TxtBrgyName, "'", "''") & "'" & "," & "'" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & "'" & "," & "'" & FrmPSGCLibrary.LstProvince.SelectedItem.Text & "'," & "'" & FrmPSGCLibrary.LstMunicipality.SelectedItem.Text & "','" & Format(TxtBrgyNumber, "000") & "')"
   Else
    MsgBox "Municipality number already exist."
   Exit Sub
  End If
   
   
   
   
   Else
   
   cnn.Execute "Update psgc set psgc_cd='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & FrmPSGCLibrary.LstMunicipality.SelectedItem.Text & Format(TxtBrgyNumber, "000") & "',name='" & Replace(Me.TxtBrgyName, "'", "''") & "',reg='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & "',prov='" & FrmPSGCLibrary.LstProvince.SelectedItem.Text & "',mun='" & FrmPSGCLibrary.LstMunicipality.SelectedItem.Text & "',brgy='" & Format(TxtBrgyNumber, "000") & "' where psgc_cd='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & FrmPSGCLibrary.LstMunicipality.SelectedItem.Text & FrmPSGCLibrary.LstBarangay.SelectedItem.Text & "'"
   
End If


   'FrmPsgclibrary.LoadProvince FrmPsgclibrary.LstRegion.SelectedItem.Text
   'FrmPsgclibrary.LoadMunicipality FrmPsgclibrary.LstRegion.SelectedItem.Text, FrmPsgclibrary.LstProvince.SelectedItem.Text
   If FrmPSGCLibrary.LstMunicipality.ListItems.Count > 0 Then
   FrmPSGCLibrary.LoadBarangay FrmPSGCLibrary.LstRegion.SelectedItem.Text, FrmPSGCLibrary.LstProvince.SelectedItem.Text, FrmPSGCLibrary.LstMunicipality.SelectedItem.Text
   Else
   FrmPSGCLibrary.LstBarangay.ListItems.Clear
   End If
Unload Me

Exit Sub
hell:


If Err.Number = -2147217900 Then
   MsgBox "Municipality Number already exist."
End If

End Sub

Private Sub RaveClose_Click()
Unload Me
End Sub
