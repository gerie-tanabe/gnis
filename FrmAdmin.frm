VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmAdmin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit / Delete as Adminitrator"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveOK 
      Height          =   465
      Left            =   2610
      TabIndex        =   3
      Top             =   1935
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
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmAdmin.frx":0000
      PICN            =   "FrmAdmin.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveCancel 
      Height          =   465
      Left            =   945
      TabIndex        =   4
      Top             =   1935
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
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmAdmin.frx":10AE
      PICN            =   "FrmAdmin.frx":10CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox TextBoxUsername 
      Height          =   390
      Left            =   1665
      TabIndex        =   0
      Top             =   675
      Width           =   2700
      VariousPropertyBits=   1820346395
      BackColor       =   12632256
      ForeColor       =   16777215
      BorderStyle     =   1
      Size            =   "4762;688"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Eurostile"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Only users with administrator account can edit/delete a record."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   450
      TabIndex        =   6
      Top             =   90
      Width           =   4560
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   450
      TabIndex        =   5
      Top             =   720
      Width           =   945
   End
   Begin MSForms.TextBox TxtPassword 
      Height          =   390
      Left            =   1665
      TabIndex        =   1
      Top             =   1170
      Width           =   2700
      VariousPropertyBits=   1820346395
      BackColor       =   12632256
      ForeColor       =   16777215
      BorderStyle     =   1
      Size            =   "4762;688"
      PasswordChar    =   61
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Webdings"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   2
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   450
      TabIndex        =   2
      Top             =   1215
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckUser()

  
    Dim rst As New ADODB.Recordset
    rst.Open "select * from useraccounts where username='" & Trim(Replace(Me.TextBoxUsername, "'", "''")) & "' and password='" & Trim(Replace(Me.TxtPassword, "'", "''")) & "' And access=1", cnn, adOpenStatic
    
    If rst.RecordCount > 0 Then
    
      TemporaryPass = True
       Unload Me
     
       
       Else
       
       MsgBox "Invalid login for an administrator", vbCritical, "Login"
       Me.TextBoxUsername = ""
       Me.TxtPassword = ""
       Me.TextBoxUsername.SetFocus
       Exit Sub
    End If
    
End Sub

Private Sub Form_Load()
TemporaryPass = False
End Sub

Private Sub RaveCancel_Click()
    Unload Me
End Sub

Private Sub RaveOK_Click()
    Call CheckUser
End Sub


Private Sub TxtPassword_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    Call CheckUser
End If
End Sub
