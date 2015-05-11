VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmServer 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Connection"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   FillColor       =   &H00008000&
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2100
      Left            =   0
      Picture         =   "FrmServer.frx":0000
      ScaleHeight     =   2100
      ScaleWidth      =   1830
      TabIndex        =   3
      Top             =   0
      Width           =   1830
   End
   Begin Rave_Buttons.RaveButtons ConnectRaveButtons 
      Height          =   510
      Left            =   2160
      TabIndex        =   4
      Top             =   1350
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   900
      BTYPE           =   3
      TX              =   "Connect"
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmServer.frx":0FD7
      PICN            =   "FrmServer.frx":0FF3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox TextBoxServer 
      Height          =   405
      Left            =   1980
      TabIndex        =   2
      Top             =   765
      Width           =   2280
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "4022;714"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1980
      TabIndex        =   1
      Top             =   450
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ex: 192.168.100.1"
      Height          =   240
      Left            =   2880
      TabIndex        =   0
      Top             =   2070
      Width           =   1500
   End
End
Attribute VB_Name = "FrmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ConnectRaveButtons_Click()


On Error GoTo Hell
   
   If Trim(Me.TextBoxServer) = "" Then
        Exit Sub
   End If
    
    Me.ConnectRaveButtons.Enabled = False
    cnn.Open "Provider=SQLOLEDB.1;User ID=gcpds;password=gcpds;Initial Catalog=gcpds;Data Source =" & Me.TextBoxServer & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
    SaveSetting App.EXEName, "Server", "ServerName", Trim(Me.TextBoxServer)
    Unload Me
    FrmLogin.Show 1
    
    
Exit Sub
Hell:
    
    Me.ConnectRaveButtons.Enabled = True
    MsgBox "Unable to connect to server" & vbCrLf & Err.Description, vbCritical, "Connection"
    
   
End Sub

Private Sub Form_Load()

   Me.TextBoxServer.Text = GetSetting(App.EXEName, "Server", "ServerName")
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
'End
End Sub



