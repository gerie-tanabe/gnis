VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmStationDetails 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "station details"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   Icon            =   "FrmStationDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin MSForms.TextBox Description 
      Height          =   3270
      Left            =   1935
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3915
      Width           =   4140
      VariousPropertyBits=   -327137249
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   50
      BorderStyle     =   1
      ScrollBars      =   2
      Size            =   "7302;5768"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   6
      Left            =   270
      MouseIcon       =   "FrmStationDetails.frx":0706
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4050
      Width           =   465
   End
   Begin MSForms.TextBox Order 
      Height          =   390
      Left            =   1935
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   945
      Width           =   4140
      VariousPropertyBits=   1820346399
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "7302;688"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   5
      Left            =   270
      MouseIcon       =   "FrmStationDetails.frx":0A10
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   990
      Width           =   690
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   4
      Left            =   270
      MouseIcon       =   "FrmStationDetails.frx":0D1A
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3105
      Width           =   1050
   End
   Begin MSForms.TextBox Barangay 
      Height          =   390
      Left            =   1935
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3060
      Width           =   4140
      VariousPropertyBits=   1820346399
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "7302;688"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Municipality:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   3
      Left            =   270
      MouseIcon       =   "FrmStationDetails.frx":1024
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2655
      Width           =   1320
   End
   Begin MSForms.TextBox Municipality 
      Height          =   390
      Left            =   1935
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2610
      Width           =   4140
      VariousPropertyBits=   1820346399
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "7302;688"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Province:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   2
      Left            =   270
      MouseIcon       =   "FrmStationDetails.frx":132E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2205
      Width           =   1020
   End
   Begin MSForms.TextBox Province 
      Height          =   390
      Left            =   1935
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   4140
      VariousPropertyBits=   1820346399
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "7302;688"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   1
      Left            =   270
      MouseIcon       =   "FrmStationDetails.frx":1638
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1755
      Width           =   825
   End
   Begin MSForms.TextBox Region 
      Height          =   390
      Left            =   1935
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1710
      Width           =   4140
      VariousPropertyBits=   1820346399
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "7302;688"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TxtName 
      Height          =   390
      Left            =   1935
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   4140
      VariousPropertyBits=   1820346399
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "7302;688"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Station Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   0
      Left            =   270
      MouseIcon       =   "FrmStationDetails.frx":1942
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   315
      Width           =   1470
   End
End
Attribute VB_Name = "FrmStationDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub StationName_Click()

End Sub

Private Sub Form_Load()
Dim rst As New ADODB.Recordset
rst.Open "Select * from sent_table where stat_name='" & CurrentStationToView & "'", cnn, adOpenStatic

Me.TxtName = rst!Stat_Name
Me.Order = rst!H_Order
Me.Region = rst!Region
Me.Province = rst!Province
Me.Municipality = rst!Municipal
Me.Barangay = rst!Barangay
Me.Description = rst!Description
End Sub

