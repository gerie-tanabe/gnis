VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmIdentify 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Basic Information"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveDelete 
      Height          =   435
      Left            =   5310
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   45
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   767
      BTYPE           =   9
      TX              =   ""
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   4210752
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmIdentify.frx":0000
      PICN            =   "FrmIdentify.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox TextBoxOrder 
      Height          =   345
      Left            =   315
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5535
      Width           =   2430
      VariousPropertyBits=   2088781855
      BackColor       =   4210752
      ForeColor       =   16777215
      Size            =   "635;609"
      Value           =   "1"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Eurostile"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   7
      Left            =   465
      TabIndex        =   16
      Top             =   5850
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "ESRI Cartography"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   405
      TabIndex        =   15
      Top             =   135
      Width           =   780
   End
   Begin VB.Label TxtName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MMA-1"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1170
      TabIndex        =   14
      Top             =   180
      Width           =   3480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Station Name"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   6
      Left            =   2295
      TabIndex        =   13
      Top             =   495
      Width           =   1050
   End
   Begin VB.Label LongitudeLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latitude"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   450
      TabIndex        =   11
      Top             =   4905
      Width           =   5010
   End
   Begin VB.Label LatitudeLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latitude"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   450
      TabIndex        =   10
      Top             =   4185
      Width           =   5100
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Longitude:"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   465
      TabIndex        =   9
      Top             =   5175
      Width           =   780
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latitude:"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   465
      TabIndex        =   8
      Top             =   4455
      Width           =   660
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   3780
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Municipality"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   465
      TabIndex        =   6
      Top             =   3060
      Width           =   870
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Province"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   465
      TabIndex        =   5
      Top             =   2340
      Width           =   645
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      BeginProperty Font 
         Name            =   "Eurostile"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   465
      TabIndex        =   4
      Top             =   1575
      Width           =   510
   End
   Begin MSForms.TextBox TxtBarangay 
      Height          =   345
      Left            =   315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3465
      Width           =   5265
      VariousPropertyBits=   2088781855
      BackColor       =   4210752
      ForeColor       =   16777215
      Size            =   "2381;609"
      Value           =   "Barangay"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Eurostile"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TxtMunicipality 
      Height          =   345
      Left            =   315
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2745
      Width           =   5235
      VariousPropertyBits=   2088781855
      BackColor       =   4210752
      ForeColor       =   16777215
      Size            =   "3598;609"
      Value           =   "Parañaque City"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Eurostile"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TxtProvince 
      Height          =   345
      Left            =   315
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2025
      Width           =   5175
      VariousPropertyBits=   2088781855
      BackColor       =   4210752
      ForeColor       =   16777215
      Size            =   "3175;609"
      Value           =   "Metro Manila"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Eurostile"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TxtRegion 
      Height          =   345
      Left            =   315
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1305
      Width           =   5145
      VariousPropertyBits=   2088781855
      BackColor       =   4210752
      ForeColor       =   16777215
      Size            =   "1296;609"
      Value           =   "NCR"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Eurostile"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
TranslucentForm Me, 230

Dim rst As New ADODB.Recordset
rst.Open "Select stat_name,region,province,municipal,barangay,wgs84nd,wgs84nm,wgs84ns,wgs84ed,wgs84em,wgs84es,order_lib.description as h_order from geoprov left join order_lib on geoprov.h_order = order_lib.h_order where stat_name='" & CurrentStationToIdentify & "'", cnn, adOpenStatic, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.TxtName = rst!Stat_Name
        Me.TxtBarangay = IIf(IsNull(rst!Barangay), "", StrConv(rst!Barangay, vbProperCase))
        Me.TxtMunicipality = IIf(IsNull(rst!Barangay), "", StrConv(rst!Municipal, vbProperCase))
        Me.TxtRegion = IIf(IsNull(rst!Region), "", StrConv(rst!Region, vbUpperCase))
        Me.TxtProvince = IIf(IsNull(rst!Province), "", StrConv(rst!Province, vbProperCase))
        Me.LatitudeLabel = rst!wgs84ED & "° " & rst!wgs84EM & "' " & rst!wgs84ES & """"
        Me.LongitudeLabel = rst!wgs84ND & "° " & rst!wgs84NM & "' " & rst!wgs84NS & """"
        Me.TextBoxOrder = IIf(IsNull(rst!H_Order), "", rst!H_Order)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
    
    If Button = 1 Then
    Call ReleaseCapture
    
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   
    End If
End Sub

Private Sub RaveDelete_Click()
Unload Me
End Sub

