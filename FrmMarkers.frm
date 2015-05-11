VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmMarkers 
   BackColor       =   &H007D3D1C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Go To"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons OKRaveButtons 
      Height          =   420
      Left            =   225
      TabIndex        =   2
      Top             =   2340
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   741
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMarkers.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons CancelRaveButtons1 
      Height          =   420
      Left            =   1485
      TabIndex        =   4
      Top             =   2340
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   741
      BTYPE           =   3
      TX              =   "Close"
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
      MICON           =   "FrmMarkers.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latitude: (DD mm ss)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   2250
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Longitude: (DD mm ss)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1125
      Width           =   2415
   End
   Begin MSForms.TextBox LatitudeTextBox 
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   495
      Width           =   3735
      VariousPropertyBits=   1820346387
      BackColor       =   16777215
      ForeColor       =   16443356
      BorderStyle     =   1
      Size            =   "6588;688"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox LoangitudeTextBox 
      Height          =   390
      Left            =   180
      TabIndex        =   1
      Top             =   1530
      Width           =   3735
      VariousPropertyBits=   1820346387
      BackColor       =   16777215
      ForeColor       =   16443356
      BorderStyle     =   1
      Size            =   "6588;688"
      BorderColor     =   11110782
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmMarkers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelRaveButtons1_Click()
Unload Me
End Sub

Private Sub OKRaveButtons_Click()

Dim ptx As New MapObjects.Point
Dim latitude As Double
Dim longitude As Double
Dim buf
Dim buf2
Dim R
buf = Split(Me.LatitudeTextBox, " ")
buf2 = Split(Me.LoangitudeTextBox, " ")

If UBound(buf) < 2 Then
    MsgBox "Invalid Coordinates."
    Exit Sub
End If

If UBound(buf2) < 2 Then
    MsgBox "Invalid Coordinates."
    Exit Sub
End If

If IsNumeric(buf(0)) = False Or IsNumeric(buf(1)) = False Or IsNumeric(buf(2)) = False Or IsNumeric(buf2(0)) = False Or IsNumeric(buf2(1)) = False Or IsNumeric(buf2(2)) = False Then
    MsgBox "Invalid Coordinates. "
    Exit Sub
End If

If buf(0) < 4 Or buf(0) > 22 Then
    MsgBox "Latitude degree should be 4° to 22°"
    Exit Sub
End If

If buf(1) < 0 Or buf(1) > 60 Then
    MsgBox "Latitude minutes should be 0° to 59°"
    Exit Sub
End If

If buf(2) < 0 Or buf(2) > 60 Then
    MsgBox "Latitude seconds should be 0° to less than 60°"
    Exit Sub
End If

If buf2(0) < 112 Or buf2(0) > 127 Then
    MsgBox "Longitude Degree should be 112° to 127°"
    Exit Sub
End If

If buf2(1) < 0 Or buf2(1) > 60 Then
    MsgBox "Longitude minutes should be 0° to 59°"
    Exit Sub
End If
If buf2(2) < 0 Or buf2(2) > 60 Then
    MsgBox "Longitude seconds should be 0° to less than 60°"
    Exit Sub
End If


            latitude = buf(0) + (buf(1) / 60) + (buf(2) / 3600)
            longitude = buf2(0) + (buf2(1) / 60) + (buf2(2) / 3600)
            ptx.y = latitude
            ptx.x = longitude
            
         

 FrmGCPDS.MyMap.Extent = FrmGCPDS.MyMap.FullExtent
Set R = FrmGCPDS.MyMap.Extent
        R.ScaleRectangle 0.01
        
         FrmGCPDS.MyMap.Extent = R
         FrmGCPDS.MyMap.CenterAt ptx.x, ptx.y
           FlashX = ptx.x
           FlashY = ptx.y
            Unload Me

End Sub
