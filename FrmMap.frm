VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0460CA20-346F-11CF-8682-00805F7CED21}#1.1#0"; "Mo10.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmMap 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Map Objects"
   ClientHeight    =   9360
   ClientLeft      =   -1350
   ClientTop       =   -1620
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   Picture         =   "FrmMap.frx":0000
   ScaleHeight     =   9360
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin MapObjects.Map MyMap 
      Height          =   6465
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   9285
      _Version        =   65537
      _ExtentX        =   16378
      _ExtentY        =   11404
      _StockProps     =   225
      BackColor       =   4194304
      BorderStyle     =   1
      ScrollBars      =   0   'False
      BackColor       =   4194304
      Contents        =   "FrmMap.frx":2F0FA
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   9645
      TabIndex        =   11
      Top             =   2205
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Min             =   5
      Max             =   30
      SelStart        =   5
      TickStyle       =   3
      Value           =   5
   End
   Begin Rave_Buttons.RaveButtons RaveZoomIn 
      Height          =   945
      Left            =   210
      TabIndex        =   1
      Top             =   1710
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1667
      BTYPE           =   11
      TX              =   "Zoom In"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmMap.frx":2F114
      PICN            =   "FrmMap.frx":2F130
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveZoomOut 
      Height          =   945
      Left            =   1890
      TabIndex        =   2
      Top             =   1710
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1667
      BTYPE           =   11
      TX              =   "  Zoom Out      "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmMap.frx":2F44B
      PICN            =   "FrmMap.frx":2F467
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RavePan 
      Height          =   945
      Left            =   3570
      TabIndex        =   3
      Top             =   1680
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1667
      BTYPE           =   11
      TX              =   "      Pan              "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmMap.frx":2F75E
      PICN            =   "FrmMap.frx":2F77A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveExtent 
      Height          =   945
      Left            =   5250
      TabIndex        =   4
      Top             =   1710
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1667
      BTYPE           =   11
      TX              =   "   Full Extent     "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmMap.frx":2FBF3
      PICN            =   "FrmMap.frx":2FC0F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveIdentify 
      Height          =   945
      Left            =   6900
      TabIndex        =   5
      Top             =   1710
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1667
      BTYPE           =   11
      TX              =   "     Focus          "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmMap.frx":2FFA9
      PICN            =   "FrmMap.frx":2FFC5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveNearest 
      Height          =   945
      Left            =   8580
      TabIndex        =   6
      Top             =   1680
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1667
      BTYPE           =   11
      TX              =   "  Find Nearest "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmMap.frx":3025D
      PICN            =   "FrmMap.frx":30279
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   9645
      TabIndex        =   12
      Top             =   3315
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Min             =   5
      Max             =   30
      SelStart        =   5
      TickStyle       =   3
      Value           =   5
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   255
      Left            =   9690
      TabIndex        =   13
      Top             =   4290
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Min             =   5
      Max             =   30
      SelStart        =   5
      TickStyle       =   3
      Value           =   5
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   12660
      TabIndex        =   0
      Top             =   3300
      Width           =   2385
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   12690
      TabIndex        =   15
      Top             =   2310
      Width           =   2385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   12690
      TabIndex        =   14
      Top             =   1185
      Width           =   2385
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   330
      Left            =   9615
      TabIndex        =   10
      Top             =   4770
      Width           =   2055
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3625;582"
      Value           =   "0"
      Caption         =   "Proposed Controls"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox Chk3rd 
      Height          =   360
      Left            =   9585
      TabIndex        =   9
      Top             =   3960
      Width           =   1200
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2117;635"
      Value           =   "0"
      Caption         =   "3rd Order"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox Chk2nd 
      Height          =   360
      Left            =   9585
      TabIndex        =   8
      Top             =   2955
      Width           =   1200
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2117;635"
      Value           =   "0"
      Caption         =   "2nd Order"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox Chk1st 
      Height          =   360
      Left            =   9585
      TabIndex        =   7
      Top             =   1830
      Width           =   1155
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2037;635"
      Value           =   "0"
      Caption         =   "1st Order"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
MyMap.TrackingLayer.ClearEvents
GIS
End Sub

Private Sub Chk1st_Click()
MyMap.TrackingLayer.ClearEvents
GIS
End Sub

Private Sub Chk2nd_Click()
MyMap.TrackingLayer.ClearEvents
GIS
End Sub

Private Sub Chk3rd_Click()
MyMap.TrackingLayer.ClearEvents
GIS
End Sub

Private Sub First_Order()
Dim x As Long
    Dim tmprecs As New ADODB.Recordset
    tmprecs.Open "Select latitude,longitude from geoprov where h_order=1", cnn, adOpenStatic, adLockOptimistic
    Me.Label1.Caption = tmprecs.RecordCount & " controls"
    
For x = 0 To tmprecs.RecordCount - 1
    DrawPoints CDbl(tmprecs("longitude").Value), CDbl(tmprecs("latitude").Value), 1, moCyan
    ReDim Preserve gEventList(x)
    'Set gEventList(x) = MyMap.TrackingLayer.Event(x)
    tmprecs.MoveNext
Next
End Sub

Private Sub Second_Order()
Dim x As Long
    Dim tmprecs As New ADODB.Recordset
    tmprecs.Open "Select latitude,longitude from geoprov where h_order=2", cnn, adOpenStatic, adLockOptimistic
    Me.Label2.Caption = tmprecs.RecordCount & " controls"
For x = 0 To tmprecs.RecordCount - 1
    DrawPoints CDbl(tmprecs("longitude").Value), CDbl(tmprecs("latitude").Value), 2, moRed
    ReDim Preserve gEventList(x)
    'Set gEventList(x) = MyMap.TrackingLayer.Event(x)
    tmprecs.MoveNext
Next
End Sub

Private Sub Third_Order()
Dim x As Long
    Dim tmprecs As New ADODB.Recordset
    tmprecs.Open "Select latitude,longitude from geoprov where h_order=3", cnn, adOpenStatic, adLockOptimistic
    Me.Label3.Caption = tmprecs.RecordCount & " controls"
For x = 0 To tmprecs.RecordCount - 1
    DrawPoints CDbl(tmprecs("longitude").Value), CDbl(tmprecs("latitude").Value), 3, moPurple
    ReDim Preserve gEventList(x)
    'Set gEventList(x) = MyMap.TrackingLayer.Event(x)
    tmprecs.MoveNext
Next
End Sub

Private Sub GIS()




    If Me.Chk3rd.Value = True Then
        Third_Order
    End If
    
    If Me.Chk2nd.Value = True Then
        Second_Order
    End If
    
    If Me.Chk1st.Value = True Then
        First_Order
    End If
End Sub


Public Sub DrawPoints(x As Double, y As Double, Index As Integer, Color As MapObjects.ColorConstants)

Dim gcp As New MapObjects.Point

   'Set gcp = FrmGCPDS.mymap.ToMapPoint(X, Y)
   gcp.x = x
   gcp.y = y
   MyMap.TrackingLayer.AddEvent gcp.x, gcp.y, Index
  

End Sub



Private Sub RaveExtent_Click()
Me.MyMap.Extent = Me.MyMap.FullExtent
'Me.StatusBar1.Panels(4) = "Zoom: " & MyMap.Extent.Height
End Sub

Private Sub RaveZoomIn_Click()
If Me.RaveZoomIn.Value = True Then
    Me.MyMap.MousePointer = moZoomIn
    Me.RavePan.Value = False
    Me.RaveNearest.Value = False
    Else
    Me.MyMap.MousePointer = moArrow
 End If
End Sub


Sub DoZoom()
Dim R  ' get a rectangle from the user
  Set R = Me.MyMap.TrackRectangle
  ' zoom to the rectangle if its valid
  If Not R Is Nothing Then MyMap.Extent = R
  Me.MyMap.Refresh
  'Me.StatusBar1.Panels(4) = "Zoom: " & MyMap.Extent.Height
  
End Sub

Private Sub Slider1_Change()
Me.MyMap.TrackingLayer.Symbol(1).SIZE = Me.Slider1.Value
MyMap.Refresh
End Sub

Private Sub Slider2_Change()
Me.MyMap.TrackingLayer.Symbol(2).SIZE = Me.Slider2.Value
MyMap.Refresh
End Sub


Private Sub Slider3_Change()
Me.MyMap.TrackingLayer.Symbol(3).SIZE = Me.Slider3.Value
MyMap.Refresh
End Sub

