VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmDbase4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dbase IV"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   Icon            =   "FrmDbase4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4125
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1350
      Width           =   2265
      BackColor       =   14737632
      Size            =   "3995;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblProvince 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Province"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   765
   End
   Begin VB.Label LlbStation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Station"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "FrmDbase4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim prov As New ADODB.Recordset

Dim x, i As Integer
DbaseCount = 0


'cnn.Execute "Delete from geoprov"
'cnn.Execute "Delete from benchmarks"
'cnn.Execute "Delete from duplicategcps"

prov.Open "select prov_code,prov from prov", cnn, adOpenStatic, adLockOptimistic

FrmDbase4.ProgressBar1.Min = 1
FrmDbase4.ProgressBar1.Max = prov.RecordCount

For x = 1 To prov.RecordCount
   DoEvents
   FrmDbase4.LblProvince.Caption = prov("Prov").Value
   Call DBaseToAccess(FrmUtilities.DbaseDir.Path & "\" & prov("prov_code").Value, prov("prov_code").Value)
   FrmDbase4.ProgressBar1.Value = x
   
   prov.MoveNext
Next

Unload Me


End Sub

