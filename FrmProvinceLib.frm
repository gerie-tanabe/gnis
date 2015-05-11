VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "ravebuttons.ocx"
Begin VB.Form FrmProvinceLib 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Province"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveAddProvince 
      Height          =   435
      Left            =   1500
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   767
      BTYPE           =   8
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmProvinceLib.frx":0000
      PICN            =   "FrmProvinceLib.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveEditProvince 
      Height          =   435
      Left            =   3030
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   767
      BTYPE           =   8
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "FrmProvinceLib.frx":01A0
      PICN            =   "FrmProvinceLib.frx":01BC
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
      Caption         =   "Acronym:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   1455
      Width           =   1020
   End
   Begin MSForms.TextBox TxtAcronym 
      Height          =   390
      Left            =   2250
      TabIndex        =   6
      Top             =   1410
      Width           =   3015
      VariousPropertyBits=   747653147
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
      Caption         =   "Province Name:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   525
      Width           =   1695
   End
   Begin MSForms.TextBox TxtProvinceName 
      Height          =   390
      Left            =   2250
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      VariousPropertyBits=   746604571
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
   Begin MSForms.TextBox TxtProvinceNumber 
      Height          =   390
      Left            =   2250
      TabIndex        =   1
      Top             =   945
      Width           =   3015
      VariousPropertyBits=   747653147
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
      Caption         =   "PSGC Code:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   990
      Width           =   1230
   End
End
Attribute VB_Name = "FrmProvinceLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If ProvinceEditMode = True Then
       Me.TxtProvinceNumber = FrmPSGCLibrary.LstProvince.SelectedItem.Text
       Me.TxtProvinceName = FrmPSGCLibrary.LstProvince.SelectedItem.SubItems(1)
       Me.TxtAcronym = FrmPSGCLibrary.LstProvince.SelectedItem.SubItems(2)
    End If
End Sub

Private Sub RaveAddProvince_Click()
'Err.Clear
'On Error GoTo hell

If Trim(Me.TxtProvinceName) = "" Then
    MsgBox "Province name is required.", vbCritical, "Province Name"
    Exit Sub
End If

If IsNumeric(Me.TxtProvinceNumber) = False Then
    MsgBox "Province number should be numeric.", vbCritical, "Province Code"
    Exit Sub
End If



If ProvinceEditMode = False Then
  
  If IfDuplicate(FrmPSGCLibrary.LstRegion.SelectedItem.Text & Format(TxtProvinceNumber, "00") & "00000") = False Then
    cnn.Execute "insert into psgc (psgc_cd,name,reg,prov,mun,brgy,acronym) values('" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & Format(TxtProvinceNumber, "00") & "00000'" & "," & "'" & Replace(TxtProvinceName, "'", "''") & "'" & "," & "'" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & "'" & "," & "'" & Format(TxtProvinceNumber, "00") & "','00','000'" & "," & "'" & TxtAcronym & "'" & ")"
   Else
    MsgBox "Province number already exist."
   Exit Sub
  End If
   
   
   
   
   Else
   
   cnn.Execute "Update psgc set psgc_cd='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & Format(TxtProvinceNumber, "00") & "00000',name='" & Replace(Me.TxtProvinceName, "'", "''") & "',reg='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & "',prov='" & Format(Me.TxtProvinceNumber, "00") & "',mun='00',brgy='000',acronym='" & Trim(TxtAcronym) & "' where psgc_cd='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & Format(FrmPSGCLibrary.LstProvince.SelectedItem.Text, "00") & "00000'"
   
End If


FrmPSGCLibrary.LoadProvince FrmPSGCLibrary.LstRegion.SelectedItem.Text
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
   MsgBox "Province Number already exist."
End If

End Sub

Private Sub RaveEditProvince_Click()
    Unload Me
End Sub




