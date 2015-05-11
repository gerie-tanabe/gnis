VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "ravebuttons.ocx"
Begin VB.Form FrmAddParty 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Requesting Party"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons CmdOK 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   4
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAddParty.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons CmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   4
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAddParty.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox txtUserName 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4815
      VariousPropertyBits=   746604563
      ForeColor       =   16777215
      BorderStyle     =   1
      Size            =   "8493;873"
      BorderColor     =   6974058
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C2935F&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   405
   End
End
Attribute VB_Name = "FrmAddParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
If Trim(Me.txtUserName.Text) <> "" Then
On Error GoTo Hell
   cnn.Execute "insert into Requesting_Party_Lib (Requesting_Party) values('" & Replace(Trim(Me.txtUserName.Text), "'", "''") & "')"
   Me.Hide
End If

Exit Sub
Hell:
If Err.Number = -2147217900 Then
    MsgBox "Requesting Party already exist.", vbInformation, "Requesting Party"
End If
End Sub

Private Sub Form_Activate()
Me.txtUserName.Text = ""
Me.txtUserName.SetFocus
End Sub

