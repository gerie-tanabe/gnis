VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmQuery 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveAddQuery 
      Height          =   315
      Left            =   2310
      TabIndex        =   5
      Top             =   2610
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      BTYPE           =   8
      TX              =   ""
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
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmQuery.frx":0000
      PICN            =   "FrmQuery.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveButtons3 
      Height          =   315
      Left            =   2790
      TabIndex        =   6
      Top             =   2610
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      BTYPE           =   8
      TX              =   ""
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
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmQuery.frx":01A0
      PICN            =   "FrmQuery.frx":01BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox BooleanOp 
      Height          =   435
      Left            =   1050
      TabIndex        =   7
      Top             =   90
      Width           =   3675
      VariousPropertyBits=   612390939
      BackColor       =   9136220
      ForeColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "6482;767"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbValue 
      Height          =   435
      Left            =   1050
      TabIndex        =   4
      Top             =   1770
      Width           =   3675
      VariousPropertyBits=   612390939
      BackColor       =   9136220
      ForeColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6482;767"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbOperator 
      Height          =   435
      Left            =   1050
      TabIndex        =   3
      Top             =   1230
      Width           =   3675
      VariousPropertyBits=   612390939
      BackColor       =   9136220
      ForeColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "6482;767"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbFields 
      Height          =   435
      Left            =   1050
      TabIndex        =   2
      Top             =   660
      Width           =   3675
      VariousPropertyBits=   612390939
      BackColor       =   9136220
      ForeColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "6482;767"
      ListRows        =   10
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
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
      Caption         =   "Value:"
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
      Left            =   180
      TabIndex        =   1
      Top             =   1860
      Width           =   675
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field:"
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
      Index           =   7
      Left            =   210
      TabIndex        =   0
      Top             =   690
      Width           =   615
   End
End
Attribute VB_Name = "FrmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conditions(1 To 22) As String
Dim Alias(1 To 22) As String
Dim Code() As Integer




Private Sub CmbFields_Click()
Me.cmbValue.Clear
If Trim(cmbFields) = "Region" Then
    LoadRegions
ElseIf Trim(cmbFields) = "Province" Then
    LoadProvinces
ElseIf Trim(cmbFields) = "Mark Purpose" Then
    LoadMarkPurpose
ElseIf Trim(cmbFields) = "Reference Type" Then
    LoadReference
ElseIf Trim(cmbFields) = "Horizontal Fixing Method" Then
    LoadHorizontalFixingMethod
ElseIf Trim(cmbFields) = "Mark Type" Then
    LoadMarkType
ElseIf Trim(cmbFields) = "Mark Status" Then
    LoadStatus
ElseIf Trim(cmbFields) = "Order" Then
    LoadOrder
Else
    Me.cmbValue.Style = 0
End If


End Sub

Private Sub Form_Activate()

If AddQuery = False Then
    Me.cmbFields = FrmGCPDS.LstConditions.SelectedItem.Text
    Me.cmbOperator = FrmGCPDS.LstConditions.SelectedItem.SubItems(1)
    Me.cmbValue = FrmGCPDS.LstConditions.SelectedItem.SubItems(2)
    
    If FrmGCPDS.LstConditions.SelectedItem.Index <> 1 Then
       Me.BooleanOp = FrmGCPDS.LstConditions.ListItems(FrmGCPDS.LstConditions.SelectedItem.Index - 1).SubItems(3)
    End If
    
End If

End Sub

Private Sub Form_Load()
 Dim i As Integer
 
 
 
 If FrmGCPDS.LstConditions.ListItems.Count = 0 Then
    Me.BooleanOp.Visible = False
    Else
    Me.BooleanOp.Visible = True
 End If
 
 If FrmGCPDS.LstConditions.ListItems.Count > 0 Then
    If AddQuery = False And FrmGCPDS.LstConditions.SelectedItem.Index = 1 Then
            Me.BooleanOp.Visible = False
        Else
            Me.BooleanOp.Visible = True
    End If
 End If
 
 
    InitializeFields
    InitializeAliases
    
    
        For i = 1 To UBound(Conditions)
            Me.cmbFields.AddItem Conditions(i)
        Next

    
    Me.cmbOperator.AddItem "="
    Me.cmbOperator.AddItem "<>"
    Me.cmbOperator.AddItem "<"
    Me.cmbOperator.AddItem ">"
    Me.cmbOperator.AddItem " LIKE "
    Me.cmbOperator.AddItem " NOT LIKE "
    Me.cmbOperator.ListIndex = 0
    
    Me.BooleanOp.AddItem " And "
    Me.BooleanOp.AddItem " Or "
    Me.BooleanOp.ListIndex = 0
    Me.cmbFields.ListIndex = 0
End Sub

Public Sub LoadRegions()
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open "select name from psgc where prov='00' order by psgc_cd", cnn, adOpenStatic
    
    Do Until rst.EOF
        DoEvents
        Me.cmbValue.AddItem rst("name")
        rst.MoveNext
    Loop
End Sub

Public Sub LoadProvinces()
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open "select name from psgc where reg<>'00'and prov<>'00' and mun='00' and brgy='000'order by name", cnn, adOpenStatic
    
    Do Until rst.EOF
        DoEvents
        Me.cmbValue.AddItem rst("name")
        rst.MoveNext
    Loop
End Sub

Public Sub LoadReference()
    
        Me.cmbValue.Clear
        Me.cmbValue.Style = 2
        Me.cmbValue.AddItem "PRS92"
        Me.cmbValue.AddItem "OLD"
        Me.cmbValue.ListIndex = 0
       
End Sub

Public Sub LoadMarkPurpose()
    Dim rst As New ADODB.Recordset
    Dim i As Integer
    rst.CursorLocation = adUseClient
    rst.Open "Select Mdesc,Mcode from MarkPur order by Mdesc", cnn, adOpenStatic, adLockOptimistic
    
    Me.cmbValue.Clear
    Me.cmbValue.Style = 2
    ReDim Code(1 To rst.RecordCount)
    For i = 1 To rst.RecordCount
        DoEvents
        Me.cmbValue.AddItem IIf(IsNull(rst("MDesc")), "", rst("MDesc"))
        Code(i) = rst("MCode")
        rst.MoveNext
    Next
    Me.cmbValue.ListIndex = 0
End Sub

Public Sub LoadMarkType()
    Dim rst As New ADODB.Recordset
    Dim i As Integer
    rst.CursorLocation = adUseClient
    rst.Open "Select MTdesc,MTcode from MarkType order by MTdesc", cnn, adOpenStatic, adLockOptimistic
    
    Me.cmbValue.Clear
    Me.cmbValue.Style = 2
    ReDim Code(1 To rst.RecordCount)
    For i = 1 To rst.RecordCount
        DoEvents
        Me.cmbValue.AddItem IIf(IsNull(rst("MTDesc")), "", rst("MTDesc"))
        Code(i) = rst("MTCode")
        rst.MoveNext
    Next
    Me.cmbValue.ListIndex = 0
End Sub

Public Sub LoadStatus()
    Dim rst As New ADODB.Recordset
    Dim i As Integer
    rst.CursorLocation = adUseClient
    rst.Open "Select MSdesc,MScode from MarkStatus order by MSdesc", cnn, adOpenStatic, adLockOptimistic
    
    Me.cmbValue.Clear
    Me.cmbValue.Style = 2
    ReDim Code(1 To rst.RecordCount)
    For i = 1 To rst.RecordCount
        DoEvents
        Me.cmbValue.AddItem IIf(IsNull(rst("MSDesc")), "", rst("MSDesc"))
        Code(i) = rst("MSCode")
        rst.MoveNext
    Next
    Me.cmbValue.ListIndex = 0
End Sub

Public Sub LoadOrder()
    Dim rst As New ADODB.Recordset
    Dim i As Integer
    rst.CursorLocation = adUseClient
    rst.Open "Select * from Order_Lib order by H_Order", cnn, adOpenStatic, adLockOptimistic
    
    Me.cmbValue.Clear
    Me.cmbValue.Style = 2
    ReDim Code(1 To rst.RecordCount)
    For i = 1 To rst.RecordCount
        DoEvents
        Me.cmbValue.AddItem IIf(IsNull(rst("Description")), "", rst("Description"))
        Code(i) = rst("H_Order")
        rst.MoveNext
    Next
    Me.cmbValue.ListIndex = 0
End Sub

Public Sub LoadHorizontalFixingMethod()
    Dim rst As New ADODB.Recordset
    Dim i As Integer
    rst.CursorLocation = adUseClient
    rst.Open "Select Hdesc,Hcode from HorFixMe order by Hdesc", cnn, adOpenStatic, adLockOptimistic
    
    Me.cmbValue.Clear
    Me.cmbValue.Style = 2
    ReDim Code(1 To rst.RecordCount)
    For i = 1 To rst.RecordCount
        DoEvents
        Me.cmbValue.AddItem IIf(IsNull(rst("HDesc")), "", rst("HDesc"))
        Code(i) = rst("HCode")
        rst.MoveNext
    Next
    Me.cmbValue.ListIndex = 0
End Sub

Private Sub RaveAddQuery_Click()


If FrmGCPDS.optBMs.Value = False And Trim(cmbFields) = "Elevation" Then
    Exit Sub
End If


Dim varlist
Dim varlist2
Dim i As Integer
strcondition = ""

If FrmGCPDS.LstConditions.ListItems.Count > 0 Then
   If RstQuery.State = 1 Then
   RstQuery.Close
   End If
End If


'Validation
                If Trim(cmbFields) = "Station Number" Or Trim(cmbFields) = "Ellipsoidal Height" Or Trim(cmbFields) = "Elevation" Then
                    If IsNumeric(Me.cmbValue) = False Then
                        MsgBox "Numeric value required."
                        Exit Sub
                    End If
                End If
                
                If Trim(cmbFields) = "Date of Entry" Or Trim(cmbFields) = "Date Computed" Or Trim(cmbFields) = "Date Last Recovered" Or Trim(cmbFields) = "Date Established" Then
                    If IsDate(Me.cmbValue) = False Then
                        MsgBox "Invalid Date."
                        Exit Sub
                    End If
                End If

'End Validation

   
    
            If AddQuery = True Then 'ADD QUERY
            
             If FrmGCPDS.LstConditions.ListItems.Count > 0 Then
            FrmGCPDS.LstConditions.ListItems(FrmGCPDS.LstConditions.ListItems.Count).SubItems(3) = (Me.BooleanOp)
             End If
    
            
            Set varlist = FrmGCPDS.LstConditions.ListItems.Add
                        varlist.Text = Trim(Me.cmbFields)
                        varlist.SubItems(1) = (Me.cmbOperator)
                        varlist.SubItems(2) = Trim(Me.cmbValue)
                        varlist.SubItems(4) = Alias(Me.cmbFields.ListIndex + 1)
                        'Mark Purpose
                        If Trim(cmbFields) = "Mark Purpose" Or Trim(cmbFields) = "Mark Status" Or Trim(cmbFields) = "Horizontal Fixing Method" Or Trim(cmbFields) = "Order" Or Trim(cmbFields) = "Mark Type" Then
                            varlist.SubItems(5) = Code(Me.cmbValue.ListIndex + 1)
                        End If
                        'Convert to decimal, Station Number and Ellipsoidal height
                        If Trim(cmbFields) = "Station Number" Or Trim(cmbFields) = "Ellipsoidal Height" Or Trim(cmbFields) = "Elevation" Then
                            varlist.SubItems(5) = CDec(Me.cmbValue)
                        End If
                        
                        'Like Operator
                        If Trim(cmbFields) = "Station Name" Or Trim(cmbFields) = "Region" Or Trim(cmbFields) = "Province" Or Trim(cmbFields) = "Municipality" Or Trim(cmbFields) = "Barangay" Or Trim(cmbFields) = "Reference Type" Or Trim(cmbFields) = "Established By" Or Trim(cmbFields) = "Responsible Authority" Or Trim(cmbFields) = "Description" Or Trim(cmbFields) = "Island" Then
                           If Trim(Me.cmbOperator) = "LIKE" Or Trim(Me.cmbOperator) = "NOT LIKE" Then
                                   varlist.SubItems(5) = "'%" & Replace(Me.cmbValue.Text, "'", "''") & "%'"
                               Else
                                   varlist.SubItems(5) = "'" & Replace(Me.cmbValue.Text, "'", "''") & "'"
                           End If
                        End If
                        
                        If Trim(cmbFields) = "Date of Entry" Or Trim(cmbFields) = "Date Computed" Or Trim(cmbFields) = "Date Last Recovered" Or Trim(cmbFields) = "Date Established" Then
                            varlist.SubItems(5) = "'" & Format((Me.cmbValue), "mm-dd-yyyy") & "'"
                        End If
                        
             End If
             
              If AddQuery = False Then 'EDIT QUERY
                 FrmGCPDS.LstConditions.SelectedItem.Text = Trim(Me.cmbFields)
                 FrmGCPDS.LstConditions.SelectedItem.SubItems(1) = (Me.cmbOperator)
                 FrmGCPDS.LstConditions.SelectedItem.SubItems(2) = Trim(Me.cmbValue)
                 FrmGCPDS.LstConditions.SelectedItem.SubItems(4) = Alias(Me.cmbFields.ListIndex + 1)
                 
                 If FrmGCPDS.LstConditions.SelectedItem.Index <> 1 Then
                     FrmGCPDS.LstConditions.ListItems(FrmGCPDS.LstConditions.SelectedItem.Index - 1).SubItems(3) = (Me.BooleanOp)
                 End If
                
                'Mark Purpose
                        If Trim(cmbFields) = "Mark Purpose" Or Trim(cmbFields) = "Mark Status" Or Trim(cmbFields) = "Horizontal Fixing Method" Or Trim(cmbFields) = "Order" Or Trim(cmbFields) = "Mark Type" Then
                            FrmGCPDS.LstConditions.SelectedItem.SubItems(5) = Code(Me.cmbValue.ListIndex + 1)
                        End If
                        'Convert to decimal, Station Number and Ellipsoidal height
                        If Trim(cmbFields) = "Station Number" Or Trim(cmbFields) = "Ellipsoidal Height" Or Trim(cmbFields) = "Elevation" Then
                           FrmGCPDS.LstConditions.SelectedItem.SubItems(5) = CDec(Me.cmbValue)
                        End If
                        
                        'Like Operator
                        If Trim(cmbFields) = "Station Name" Or Trim(cmbFields) = "Region" Or Trim(cmbFields) = "Province" Or Trim(cmbFields) = "Municipality" Or Trim(cmbFields) = "Barangay" Or Trim(cmbFields) = "Reference Type" Or Trim(cmbFields) = "Established By" Or Trim(cmbFields) = "Responsible Authority" Or Trim(cmbFields) = "Description" Or Trim(cmbFields) = "Island" Then
                           If Trim(Me.cmbOperator) = "LIKE" Or Trim(Me.cmbOperator) = "NOT LIKE" Then
                                   FrmGCPDS.LstConditions.SelectedItem.SubItems(5) = "'%" & Replace(Me.cmbValue.Text, "'", "''") & "%'"
                               Else
                                  FrmGCPDS.LstConditions.SelectedItem.SubItems(5) = "'" & Replace(Me.cmbValue.Text, "'", "''") & "'"
                           End If
                        End If
                        
                        If Trim(cmbFields) = "Date of Entry" Or Trim(cmbFields) = "Date Computed" Or Trim(cmbFields) = "Date Last Recovered" Or Trim(cmbFields) = "Date Established" Then
                            FrmGCPDS.LstConditions.SelectedItem.SubItems(5) = "'" & Format((Me.cmbValue), "mm-dd-yyyy") & "'"
                        End If
              End If
             
     
      
       BuildExecuteQuery
      
    
    Unload Me
End Sub




Public Sub InitializeFields()
    Conditions(1) = "Station Name"
    Conditions(2) = "Station Number"
    Conditions(3) = "Region"
    Conditions(4) = "Province"
    Conditions(5) = "Municipality"
    Conditions(6) = "Barangay"
    Conditions(7) = "Order"
    Conditions(8) = "Reference Type"
    Conditions(9) = "Ellipsoidal Height"
    Conditions(10) = "Date of Entry"
    Conditions(11) = "Date Computed"
    Conditions(12) = "Date Last Recovered"
    Conditions(13) = "Date Established"
    Conditions(14) = "Established By"
    Conditions(15) = "Responsible Authority"
    Conditions(16) = "Mark Type"
    Conditions(17) = "Mark Purpose"
    Conditions(18) = "Mark Status"
    Conditions(19) = "Horizontal Fixing Method"
    Conditions(20) = "Description"
    Conditions(21) = "Elevation"
    Conditions(22) = "Island"
End Sub

Public Sub InitializeAliases()
    Alias(1) = "Stat_Name"
    Alias(2) = "Stat_New"
    Alias(3) = "Region"
    Alias(4) = "Province"
    Alias(5) = "Municipal"
    Alias(6) = "Barangay"
    Alias(7) = "Geoprov.H_Order"
    Alias(8) = "H_Ref"
    Alias(9) = "Ell_Hgt"
    
    If (FrmGCPDS.optBMs.Value = False) Then
        Alias(10) = "H_Date_Ety"
    Else
        Alias(10) = "E_Date_Ety"
    End If
    
    
    Alias(11) = "H_Date_Com"
    Alias(12) = "Date_Las_R"
    Alias(13) = "Date_Est"
    Alias(14) = "Hor_Authty"
    Alias(15) = "Authority"
    Alias(16) = "Mark_Type"
    Alias(17) = "Mark_Const"
    Alias(18) = "Mark_Stat"
    Alias(19) = "H_Fix"
   
     
     If (FrmGCPDS.optBMs.Value = False) Then
        Alias(20) = "Geoprov.Description"
    Else
         Alias(20) = "Benchmarks.Description"
    End If
    
    
    Alias(21) = "Elevation"
    Alias(22) = "Island"
End Sub

Private Sub RaveButtons3_Click()
Unload Me
End Sub
