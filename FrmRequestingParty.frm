VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmRequestingParty 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requesting Party"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   5415
   Icon            =   "FrmRequestingParty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons CmdOK 
      Height          =   450
      Left            =   840
      TabIndex        =   0
      Top             =   5520
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   794
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
      MICON           =   "FrmRequestingParty.frx":000C
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
      Height          =   450
      Left            =   2760
      TabIndex        =   1
      Top             =   5520
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   794
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
      MICON           =   "FrmRequestingParty.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RequestingPartyRaveButtons 
      Height          =   330
      Left            =   4800
      TabIndex        =   2
      Top             =   840
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      BTYPE           =   4
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "FrmRequestingParty.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons CmdPurpose 
      Height          =   330
      Left            =   4800
      TabIndex        =   3
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      BTYPE           =   4
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "FrmRequestingParty.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveSignatory 
      Height          =   330
      Left            =   4800
      TabIndex        =   4
      Top             =   4320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      BTYPE           =   4
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "FrmRequestingParty.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveGets 
      Height          =   330
      Left            =   4800
      TabIndex        =   16
      Top             =   3360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      BTYPE           =   4
      TX              =   "Get"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "FrmRequestingParty.frx":0098
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox TxtSignatory 
      Height          =   330
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   4485
      VariousPropertyBits=   746604571
      BackColor       =   4210752
      ForeColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7911;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      BorderColor     =   6974058
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TxtOR 
      Height          =   330
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   4485
      VariousPropertyBits=   746604563
      ForeColor       =   16777215
      BorderStyle     =   1
      Size            =   "7911;582"
      BorderColor     =   6974058
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Signatory"
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
      TabIndex        =   13
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label TxtDesignation 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   12
      Top             =   4680
      Width           =   3075
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Requesting Party"
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
      TabIndex        =   11
      Top             =   585
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose"
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
      TabIndex        =   10
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O.R"
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
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   285
   End
   Begin MSForms.ComboBox TxtRequestingParty 
      Height          =   330
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   4485
      VariousPropertyBits=   746604563
      BackColor       =   4210752
      ForeColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7911;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      BorderColor     =   6974058
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TxtTN 
      Height          =   330
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   4485
      VariousPropertyBits=   746604563
      ForeColor       =   16777215
      BorderStyle     =   1
      Size            =   "7911;582"
      BorderColor     =   6974058
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox TxtPurpose 
      Height          =   330
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   4485
      VariousPropertyBits=   746604563
      BackColor       =   4210752
      ForeColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7911;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      BorderColor     =   6974058
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction No."
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
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1140
   End
End
Attribute VB_Name = "FrmRequestingParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdCancel_Click()
  requestingParty = Me.TxtRequestingParty
   Purpose = Me.TxtPurpose
   O_R = Me.TxtOR
   signatory = Me.TxtSignatory
   Designation = Me.TxtDesignation
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting App.EXEName, "Requesting Party", "Requesting Party", Me.TxtRequestingParty
    SaveSetting App.EXEName, "Purpose", "Purpose", Me.TxtPurpose
    SaveSetting App.EXEName, "OR Number", "OR Number", Me.TxtOR
    SaveSetting App.EXEName, "Signatory", "Signatory", Me.TxtSignatory
    SaveSetting App.EXEName, "Designation", "Designation", Me.TxtDesignation
    SaveSetting App.EXEName, "TN", "TN", Me.TxtTN
    Unload Me
End Sub




Private Sub CmdPurpose_Click()
FrmPurpose.Show 1
End Sub



Private Sub Form_Activate()
'Me.TxtRequestingParty.SetFocus
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim i As Integer
rst.Open "Select * from Requesting_Party_Lib order by Requesting_Party", cnn, adOpenStatic, adLockOptimistic
Me.TxtRequestingParty.Clear
If rst.RecordCount > 0 Then

    For i = 1 To rst.RecordCount
        Me.TxtRequestingParty.AddItem rst("Requesting_party").Value
        rst.MoveNext
    Next
   
    rst.Close
End If

rst2.Open "Select * from Purpose_Lib order by Purpose", cnn, adOpenStatic, adLockOptimistic
Me.TxtPurpose.Clear
If rst2.RecordCount > 0 Then
    For i = 1 To rst2.RecordCount
        Me.TxtPurpose.AddItem rst2("Purpose").Value
        rst2.MoveNext
    Next
  
    rst2.Close
End If

rst3.Open "Select * from Signatory order by Signatory", cnn, adOpenStatic, adLockOptimistic
Me.TxtSignatory.Clear
Me.TxtDesignation = ""
If rst3.RecordCount > 0 Then

    For i = 1 To rst3.RecordCount
        Me.TxtSignatory.AddItem rst3("Signatory").Value
        rst3.MoveNext
    Next
    
    rst3.Close
End If

Me.TxtRequestingParty.Text = GetSetting(App.EXEName, "Requesting Party", "Requesting Party")
Me.TxtSignatory.Text = GetSetting(App.EXEName, "Signatory", "Signatory")
Me.TxtPurpose = GetSetting(App.EXEName, "Purpose", "Purpose")
Me.TxtOR = GetSetting(App.EXEName, "OR Number", "OR Number")
Me.TxtTN = GetSetting(App.EXEName, "TN", "TN")
End Sub


Private Sub Form_Load()
    TranslucentForm Me, 250
End Sub

Private Sub Form_Unload(Cancel As Integer)
    requestingParty = Me.TxtRequestingParty
   Purpose = Me.TxtPurpose
   O_R = Me.TxtOR
   signatory = Me.TxtSignatory
   Designation = Me.TxtDesignation
End Sub

Private Sub RaveButtons1_Click()
SaveSetting App.EXEName, "Requesting Party", "Requesting Party", Me.TxtRequestingParty
    SaveSetting App.EXEName, "Purpose", "Purpose", Me.TxtPurpose
    SaveSetting App.EXEName, "OR Number", "OR Number", Me.TxtOR
    SaveSetting App.EXEName, "Signatory", "Signatory", Me.TxtSignatory
    SaveSetting App.EXEName, "Designation", "Designation", Me.TxtDesignation
    
    Unload Me
End Sub

Private Sub RaveButtons2_Click()

End Sub



Private Sub RaveGets_Click()
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT TOP(1) * FROM print_inventory WHERE tn IS NOT NULL AND tn!=''  order by date desc", cnn, adOpenStatic, adLockOptimistic
   
   If (rst.RecordCount > 0) Then
        Me.TxtTN.Text = rst!tn
   End If
   
   
End Sub

Private Sub RaveSignatory_Click()
    FrmSignatoryLib.Show 1
End Sub

Private Sub RequestingPartyRaveButtons_Click()
FrmRequestingPartyLib.Show 1
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TxtSignatory_Click()
Dim rst As New ADODB.Recordset
rst.Open "Select Designation from Signatory where Signatory='" & Me.TxtSignatory & "'", cnn, adOpenStatic, adLockOptimistic
Me.TxtDesignation = ""
If rst.RecordCount > 0 Then
    Me.TxtDesignation = rst("Designation").Value
    rst.Close
End If

End Sub
