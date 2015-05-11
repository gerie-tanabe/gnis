VERSION 5.00
Begin VB.Form FrmSearchProvince 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Search Province"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   5655
   Icon            =   "FrmSearchProvince.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
   End
   Begin VB.ComboBox CmbProvince 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   1200
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Text            =   "CmbProvince"
      Top             =   2280
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   0
      Picture         =   "FrmSearchProvince.frx":014A
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "National Mapping and Resource Information Authority"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   3315
      End
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Province Name"
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
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   1305
   End
End
Attribute VB_Name = "FrmSearchProvince"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbProvince_DblClick()
Call CmdOk_Click
End Sub

Private Sub CmdCancel_Click()
Me.Hide
End Sub

Private Sub CmdOk_Click()
Set rstProvinceCode = New ADODB.Recordset
    rstProvinceCode.Open "Select provincealpha from provmast where prov_name=" & "'" & Trim(Me.CmbProvince.Text) & "'", cnn, adOpenStatic
    ProvinceCode = rstProvinceCode.Fields(0).Value
Me.Hide
FrmStation.Show 1
End Sub

Private Sub Form_Activate()
Me.CmbProvince.SetFocus
End Sub

Private Sub Form_Load()
    
Set rst = New ADODB.Recordset
    rst.Open "Select prov_name,provincealpha from Provmast order by Prov_name", cnn, adOpenStatic
    
    For x = 1 To rst.RecordCount
        Me.CmbProvince.AddItem StrConv(rst.Fields(0).Value, vbProperCase)
        rst.MoveNext
    Next
    
    Me.CmbProvince.ListIndex = 0

End Sub



