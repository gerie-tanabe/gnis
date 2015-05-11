VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "Ravebuttons.ocx"
Begin VB.Form FrmSpatialSearchX2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spatial Search"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   66
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox FieldCombo 
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
      ItemData        =   "FrmSpatialSearchX2.frx":0000
      Left            =   3420
      List            =   "FrmSpatialSearchX2.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2220
   End
   Begin Rave_Buttons.RaveButtons RaveFind 
      Height          =   495
      Left            =   5805
      TabIndex        =   1
      Top             =   315
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Okay"
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
      BCOL            =   32768
      BCOLO           =   32768
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmSpatialSearchX2.frx":0026
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox TextSearchBox 
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   315
      Width           =   3150
      VariousPropertyBits=   746604571
      BackColor       =   8207644
      ForeColor       =   16777215
      BorderStyle     =   1
      Size            =   "5556;767"
      BorderColor     =   8207644
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmSpatialSearchX2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.FieldCombo.ListIndex = 0
End Sub

Private Sub RaveFind_Click()

   CurrentSpatialQuery = Trim(Me.TextSearchBox)
   CurrentSpatialQueryField = Trim(Me.FieldCombo.Text)
   
   ZoomToSelected2 Trim(Me.TextSearchBox), Trim(Me.FieldCombo.Text)
   FrmGCPDS.MyMap.Refresh
   
   Unload Me
End Sub

Private Sub TextSearchBox_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    RaveFind_Click
End If
End Sub
