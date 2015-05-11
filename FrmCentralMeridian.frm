VERSION 5.00
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmCentralMeridian 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local to Grid Conversion"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   303
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox PTMZoneCombobox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "FrmCentralMeridian.frx":0000
      Left            =   270
      List            =   "FrmCentralMeridian.frx":0016
      TabIndex        =   0
      Top             =   1215
      Width           =   4065
   End
   Begin Rave_Buttons.RaveButtons RaveConvert 
      Height          =   465
      Left            =   2250
      TabIndex        =   2
      Top             =   1890
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
      MICON           =   "FrmCentralMeridian.frx":00A1
      PICN            =   "FrmCentralMeridian.frx":00BD
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
      Left            =   585
      TabIndex        =   3
      Top             =   1890
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
      MICON           =   "FrmCentralMeridian.frx":114F
      PICN            =   "FrmCentralMeridian.frx":116B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LabelMessage 
      BackStyle       =   0  'Transparent
      Height          =   1230
      Left            =   225
      TabIndex        =   1
      Top             =   180
      Width           =   4020
   End
End
Attribute VB_Name = "FrmCentralMeridian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
    PTMZone = 0
    Me.LabelMessage.Caption = FrmGCPDS.TxtRegion & " " & FrmGCPDS.TxtProvince & " " & FrmGCPDS.TxtMunicipality & " doesn't exist in the PTM Zone library." & vbCrLf & vbCrLf & "You are required to manually select a zone."
End Sub

Private Sub RaveCancel_Click()
Unload Me
End Sub

Private Sub RaveConvert_Click()
    If Trim(Me.PTMZoneCombobox.Text) <> "" Then
        If Me.PTMZoneCombobox.ListIndex = 0 Then
           PTMZone = 117
           'PTMZone = 1
           Zone = "1"  'modified dec72009
        ElseIf Me.PTMZoneCombobox.ListIndex = 1 Then
           PTMZone = 119
           'PTMZone = 2
           Zone = "2"  'modified dec72009
        ElseIf Me.PTMZoneCombobox.ListIndex = 2 Then
           PTMZone = 121
           'PTMZone = 3
           Zone = "3"  'modified dec72009
        ElseIf Me.PTMZoneCombobox.ListIndex = 3 Then
           PTMZone = 123
           'PTMZone = 4
           Zone = "4"  'modified dec72009
        ElseIf Me.PTMZoneCombobox.ListIndex = 4 Then
           PTMZone = 125
           'PTMZone = 5
           Zone = "5"  'modified dec72009
        ElseIf Me.PTMZoneCombobox.ListIndex = 5 Then  'modified dec72009 for zone 1A
           PTMZone = 118.5
           'PTMZone = 1.75
           Zone = "1A"  'modified dec72009
        End If
        
        Unload Me
    End If
End Sub
