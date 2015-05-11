VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "ravebuttons.ocx"
Begin VB.Form FrmMunicipalityLib 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Municipality"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveAddMunicipality 
      Height          =   405
      Left            =   2160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   714
      BTYPE           =   9
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   8421504
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmMunicipalityLib.frx":0000
      PICN            =   "FrmMunicipalityLib.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveClose 
      Height          =   405
      Left            =   2880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   714
      BTYPE           =   9
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   8421504
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "FrmMunicipalityLib.frx":01A0
      PICN            =   "FrmMunicipalityLib.frx":01BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Municipality Name:"
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
      Width           =   2040
   End
   Begin MSForms.TextBox TxtMunicipalityName 
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
   Begin MSForms.TextBox TxtMunicipalityNumber 
      Height          =   390
      Left            =   2400
      TabIndex        =   1
      Top             =   945
      Width           =   3015
      VariousPropertyBits=   1821394971
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   2
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
      Caption         =   "Municipality No:"
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
      Width           =   1710
   End
End
Attribute VB_Name = "FrmMunicipalityLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If MunicipalityEditMode = True Then
       Me.TxtMunicipalityNumber = FrmPSGCLibrary.LstMunicipality.SelectedItem.Text
       Me.TxtMunicipalityName = FrmPSGCLibrary.LstMunicipality.SelectedItem.SubItems(1)
    End If
End Sub






Private Sub RaveAddMunicipality_Click()
'On Error GoTo hell

If Trim(Me.TxtMunicipalityName) = "" Then
    MsgBox "Municipality name is required.", vbCritical, "Municipality Name"
    Exit Sub
End If

If IsNumeric(Me.TxtMunicipalityNumber) = False Then
    MsgBox "Municipality Number should be numeric.", vbCritical, "Municipality Code"
    Exit Sub
End If



If MunicipalityEditMode = False Then
  
  If IfDuplicate(FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & Format(TxtMunicipalityNumber, "00") & "000") = False Then
   cnn.Execute "insert into psgc (psgc_cd,name,reg,prov,mun,brgy) values('" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & Format(TxtMunicipalityNumber, "00") & "000'" & "," & "'" & Replace(TxtMunicipalityName, "'", "''") & "'" & "," & "'" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & "'" & "," & "'" & FrmPSGCLibrary.LstProvince.SelectedItem.Text & "','" & Format(TxtMunicipalityNumber, "00") & "','000'" & ")"
   Else
    MsgBox "Municipality number already exist."
   Exit Sub
  End If
   
   
   
   
   Else
   
   cnn.Execute "Update psgc set psgc_cd='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & Format(TxtMunicipalityNumber, "00") & "000',name='" & Replace(Me.TxtMunicipalityName, "'", "''") & "',reg='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & "',prov='" & FrmPSGCLibrary.LstProvince.SelectedItem.Text & "',mun='" & Format(TxtMunicipalityNumber, "00") & "',brgy='000' where psgc_cd='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & Format(FrmPSGCLibrary.LstMunicipality.SelectedItem.Text, "00") & "000'"
   
End If


   'FrmPsgclibrary.LoadProvince FrmPsgclibrary.LstRegion.SelectedItem.Text
   FrmPSGCLibrary.LoadMunicipality FrmPSGCLibrary.LstRegion.SelectedItem.Text, FrmPSGCLibrary.LstProvince.SelectedItem.Text
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
