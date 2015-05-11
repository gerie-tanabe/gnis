VERSION 5.00
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmDelete 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   Icon            =   "FrmDelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveOK 
      Height          =   465
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   1680
      _ExtentX        =   2963
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmDelete.frx":000C
      PICN            =   "FrmDelete.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   240
      Picture         =   "FrmDelete.frx":10BA
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label LblDelete 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "FrmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
    If DataType = "GCP" Then
        Me.Caption = "GCPs"
        Me.LblDelete.Caption = "Record of " & FrmGCPDS.TxtName & " was deleted."
        BlankForm
        ElseIf DataType = "Benchmarks" Then
        Me.Caption = "Benchmarks"
        Me.LblDelete.Caption = "Record of " & FrmGCPDS.TxtEName & " was deleted."
        BlankFormBM
        ElseIf DataType = "Gravity" Then
        Me.Caption = "Gravity"
        Me.LblDelete.Caption = "Record of " & FrmGCPDS.TextBoxGravityName & " was deleted."
        BlankFormGravity
        
    End If
    
End Sub

Private Sub RaveAddQuery_Click()

End Sub

Private Sub RaveOK_Click()
 Unload Me
End Sub
