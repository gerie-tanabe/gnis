VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmLocation 
   BackColor       =   &H00A9897E&
   BorderStyle     =   0  'None
   Caption         =   "Location"
   ClientHeight    =   4740
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   5070
   Icon            =   "FrmLocation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmLocation.frx":3452
   ScaleHeight     =   4740
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveClose 
      Height          =   585
      Left            =   4380
      TabIndex        =   8
      Top             =   90
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1032
      BTYPE           =   11
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   4194304
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmLocation.frx":D270
      PICN            =   "FrmLocation.frx":D28C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveOK 
      Height          =   465
      Left            =   1890
      TabIndex        =   9
      Top             =   3960
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
      MICON           =   "FrmLocation.frx":DF66
      PICN            =   "FrmLocation.frx":DF82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image ImageBorder 
      Height          =   1575
      Left            =   30
      MousePointer    =   15  'Size All
      Top             =   0
      Width           =   5085
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   0
      Left            =   390
      TabIndex        =   7
      Top             =   1890
      Width           =   705
   End
   Begin MSForms.ComboBox txtRegion 
      Height          =   390
      Left            =   1800
      TabIndex        =   6
      Top             =   1890
      Width           =   2865
      VariousPropertyBits=   612390939
      BackColor       =   16777215
      ForeColor       =   4210752
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5054;688"
      ListWidth       =   7055
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox txtBarangay 
      Height          =   390
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Width           =   2865
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      ForeColor       =   4210752
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5054;688"
      ListWidth       =   7055
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox txtMunicipality 
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Top             =   2790
      Width           =   2865
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      ForeColor       =   4210752
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5054;688"
      ListWidth       =   7055
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox txtProvince 
      Height          =   390
      Left            =   1800
      TabIndex        =   3
      Top             =   2340
      Width           =   2865
      VariousPropertyBits=   612386843
      BackColor       =   16777215
      ForeColor       =   4210752
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5054;688"
      ListWidth       =   7055
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Trebuchet MS"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Province"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   1
      Left            =   390
      TabIndex        =   2
      Top             =   2340
      Width           =   930
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Municipality"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   2
      Left            =   390
      TabIndex        =   1
      Top             =   2820
      Width           =   1275
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   3
      Left            =   390
      TabIndex        =   0
      Top             =   3300
      Width           =   960
   End
End
Attribute VB_Name = "FrmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XRegion() As Integer
Dim XProvince() As Integer
Dim XMunicipality() As Integer
Dim XBarangay() As Integer





Private Sub RaveOK_Click()
 
 FrmGCPDS.Controls(regionTextbox).Text = Trim(StrConv(Me.TxtRegion.Text, vbUpperCase))
 FrmGCPDS.Controls(provinceTextbox).Text = Trim(StrConv(Me.TxtProvince.Text, vbUpperCase))
 FrmGCPDS.Controls(municipalityTextbox).Text = Trim(StrConv(Me.TxtMunicipality.Text, vbUpperCase))
 FrmGCPDS.Controls(barangayTextbox).Text = Trim(StrConv(Me.TxtBarangay.Text, vbUpperCase))

    
    
    Unload Me
End Sub

Private Sub Form_Activate()
    TxtRegion.SetFocus
End Sub

Private Sub Form_Load()
    
    Location_LoadRegions
    
End Sub

Public Sub Location_LoadRegions()
    
        Dim i As Integer
        Dim rstregion As ADODB.Recordset
        
            Set rstregion = New ADODB.Recordset
                rstregion.Open "Select name,reg from psgc where prov='00' and mun='00' and brgy='000' ORDER BY Case name " & _
                                                       " WHEN 'Region I' THEN 1 " & _
                                                       " WHEN 'Region II' THEN 2 " & _
                                                       " WHEN 'Region III' THEN 3 " & _
                                                       " WHEN 'Region IV-A' THEN 4 " & _
                                                       " WHEN 'Region IV-B' THEN 5 " & _
                                                       " WHEN 'Region V' THEN 6 " & _
                                                       " WHEN 'Region VI' THEN 7 " & _
                                                       " WHEN 'Region VII' THEN 8 " & _
                                                       " WHEN 'Region VIII' THEN 9 " & _
                                                       " WHEN 'Region IX' THEN 10 " & _
                                                       " WHEN 'Region X' THEN 11 " & _
                                                       " WHEN 'Region XI' THEN 12 " & _
                                                       " WHEN 'Region XII' THEN 13 " & _
                                                       " WHEN 'Region XIII' THEN 14" & _
                                                       " WHEN 'CAR' THEN 15 " & _
                                                       " WHEN 'NCR' THEN 16 " & _
                                                       " WHEN 'ARMM' THEN 17 " & _
                                                       "END", cnn, adOpenStatic, adLockOptimistic
            
        'ReDim XRegion(1 To rstcmbprovince.RecordCount)
        ReDim XRegion(1 To rstregion.RecordCount)
        
        For i = 1 To rstregion.RecordCount
            Me.TxtRegion.AddItem StrConv(rstregion("name").Value, vbUpperCase)
            XRegion(i) = rstregion("reg").Value
            rstregion.MoveNext
        Next
            
End Sub



Public Sub Location_LoadProvinces(Region As Integer)
    
        Dim i As Integer
        Dim rstcmbprovince As ADODB.Recordset
        
            Set rstcmbprovince = New ADODB.Recordset
                rstcmbprovince.Open "Select name,prov from psgc where reg='" & Format(Region, "00") & "' and prov<>'00' and mun='00' and brgy='000' order by name", cnn, adOpenStatic, adLockOptimistic
            
        'ReDim XRegion(1 To rstcmbprovince.RecordCount)
        ReDim XProvince(1 To rstcmbprovince.RecordCount)
        
        Me.TxtProvince.Clear
        For i = 1 To rstcmbprovince.RecordCount
            Me.TxtProvince.AddItem StrConv(rstcmbprovince("name").Value, vbUpperCase)
            
            'XRegion(i) = rstcmbprovince.Fields(1).Value
            XProvince(i) = rstcmbprovince("prov").Value
            rstcmbprovince.MoveNext
        Next
            
End Sub

Public Function Location_LoadMunicipality(Region As Integer, Province As Integer)
    
        Dim i As Integer
        Dim rstcmbMunicipality As ADODB.Recordset
        
            Set rstcmbMunicipality = New ADODB.Recordset
                rstcmbMunicipality.Open "Select name,mun from psgc where reg='" & Format(Region, "00") & "' and prov='" & Format(Province, "00") & "' and mun<>'00' and brgy='000' order by name", cnn, adOpenStatic, adLockOptimistic
                
                Me.TxtMunicipality.Clear
        
        ReDim XMunicipality(1 To rstcmbMunicipality.RecordCount)
        Me.TxtMunicipality.Clear
        
        For i = 1 To rstcmbMunicipality.RecordCount
            Me.TxtMunicipality.AddItem StrConv(rstcmbMunicipality("name").Value, vbUpperCase)
            XMunicipality(i) = rstcmbMunicipality("mun").Value
            rstcmbMunicipality.MoveNext
        Next
        
        'Me.txtMunicipality.ListIndex = 0
            
End Function

Public Function Location_LoadBarangay(Region As Integer, Province As Integer, Municipality As Integer)
    
        Dim i As Integer
        Dim rstcmbBarangay As ADODB.Recordset
        
            Set rstcmbBarangay = New ADODB.Recordset
                rstcmbBarangay.Open "Select name,brgy from psgc where reg='" & Format(Region, "00") & "' and prov='" & Format(Province, "00") & "' and mun='" & Format(Municipality, "00") & "' and brgy<>'000' order by name", cnn, adOpenStatic, adLockOptimistic
            If rstcmbBarangay.RecordCount > 0 Then
                ReDim XBarangay(1 To rstcmbBarangay.RecordCount)
            End If
               Me.TxtBarangay.Clear
        
        For i = 1 To rstcmbBarangay.RecordCount
            Me.TxtBarangay.AddItem StrConv(rstcmbBarangay("name"), vbUpperCase)
            XBarangay(i) = rstcmbBarangay("brgy").Value
            rstcmbBarangay.MoveNext
        Next
            
End Function

Private Sub Form_Unload(Cancel As Integer)
'Dim i As Byte
'    For i = 1 To 254
'        TranslucentForm Me, 255 - i
'        DoEvents
'    Next
End Sub

Private Sub ImageBorder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
TranslucentForm Me, 100
End Sub

Private Sub ImageBorder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long

If Button = 1 Then
Call ReleaseCapture

lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
TranslucentForm Me, 255
End If
End Sub

Private Sub ImageBorder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
TranslucentForm Me, 255
End Sub

Private Sub RaveClose_Click()
Unload Me
End Sub

Private Sub txtMunicipality_Click()
Location_LoadBarangay XRegion(TxtRegion.ListIndex + 1), XProvince(TxtProvince.ListIndex + 1), XMunicipality(TxtMunicipality.ListIndex + 1)
End Sub

Private Sub txtProvince_Click()
Location_LoadMunicipality XRegion(Me.TxtRegion.ListIndex + 1), XProvince(Me.TxtProvince.ListIndex + 1)
End Sub

Private Sub txtRegion_Click()
Location_LoadProvinces XRegion(Me.TxtRegion.ListIndex + 1)
End Sub
