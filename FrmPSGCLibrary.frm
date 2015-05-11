VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmPSGCLibrary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Location Library"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   14325
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstRegion 
      Height          =   7785
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   13732
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   8207644
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Region"
         Object.Width           =   5115
      EndProperty
   End
   Begin MSComctlLib.ListView LstProvince 
      Height          =   7785
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   13732
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   8207644
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Province"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Code"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView LstMunicipality 
      Height          =   7785
      Left            =   7230
      TabIndex        =   2
      Top             =   240
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   13732
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   8207644
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Municipality"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.ListView LstBarangay 
      Height          =   7785
      Left            =   10770
      TabIndex        =   3
      Top             =   240
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   13732
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   8207644
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Barangay"
         Object.Width           =   7056
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons RaveDeleteProvince 
      Height          =   405
      Left            =   5430
      TabIndex        =   4
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Delete"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":0000
      PICN            =   "FrmPSGCLibrary.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveAddProvince 
      Height          =   405
      Left            =   3210
      TabIndex        =   5
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Add"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":0284
      PICN            =   "FrmPSGCLibrary.frx":02A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveEditProvince 
      Height          =   405
      Left            =   4320
      TabIndex        =   6
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Edit"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":0424
      PICN            =   "FrmPSGCLibrary.frx":0440
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveDeletemunicipality 
      Height          =   405
      Left            =   9480
      TabIndex        =   7
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Delete"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":06CA
      PICN            =   "FrmPSGCLibrary.frx":06E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveAddMunicipality 
      Height          =   405
      Left            =   7260
      TabIndex        =   8
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Add"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":094E
      PICN            =   "FrmPSGCLibrary.frx":096A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveEditmunicipality 
      Height          =   405
      Left            =   8370
      TabIndex        =   9
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Edit"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":0AEE
      PICN            =   "FrmPSGCLibrary.frx":0B0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveDeleteBrgy 
      Height          =   405
      Left            =   13050
      TabIndex        =   10
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Delete"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":0D94
      PICN            =   "FrmPSGCLibrary.frx":0DB0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveAddBrgy 
      Height          =   405
      Left            =   10830
      TabIndex        =   11
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Add"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":1018
      PICN            =   "FrmPSGCLibrary.frx":1034
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveEditBrgy 
      Height          =   405
      Left            =   11940
      TabIndex        =   12
      Top             =   8205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   8
      TX              =   "Edit"
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
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPSGCLibrary.frx":11B8
      PICN            =   "FrmPSGCLibrary.frx":11D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmPSGCLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
     LoadRegion
     LoadProvince Me.LstRegion.SelectedItem.Text
     LoadMunicipality Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text
     If Me.LstMunicipality.ListItems.Count > 0 Then
     LoadBarangay Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text, Me.LstMunicipality.SelectedItem.Text
     Else
     Me.LstBarangay.ListItems.Clear
     End If
End Sub


Private Sub LoadRegion()
    Dim i As Integer
    Dim varlist
        Dim rstregion As ADODB.Recordset
        
            Set rstregion = New ADODB.Recordset
                rstregion.Open "Select name,reg from psgc where prov='00' and mun='00' and brgy='000' order by psgc_cd", cnn, adOpenStatic, adLockOptimistic
            
      
        For i = 1 To rstregion.RecordCount
            Set varlist = Me.LstRegion.ListItems.Add
            varlist.SubItems(1) = StrConv(rstregion("name").Value, vbUpperCase)
            varlist.Text = rstregion("reg").Value
            rstregion.MoveNext
        Next
End Sub

Public Sub LoadProvince(Region As String)
    Dim i As Integer
    Dim varlist
        Dim rstProvince As ADODB.Recordset
        
            Set rstProvince = New ADODB.Recordset
                rstProvince.Open "Select name,prov,ACRONYM from psgc where reg='" & Region & "' and prov<>'00' and mun='00' and brgy='000' order by name", cnn, adOpenStatic, adLockOptimistic
            
            Me.LstProvince.ListItems.Clear
        For i = 1 To rstProvince.RecordCount
            Set varlist = Me.LstProvince.ListItems.Add
            varlist.SubItems(1) = StrConv(rstProvince("name").Value, vbUpperCase)
            varlist.SubItems(2) = IIf(IsNull(rstProvince("ACRONYM")), "", StrConv(rstProvince("ACRONYM").Value, vbUpperCase))
            varlist.Text = rstProvince("prov").Value
            rstProvince.MoveNext
        Next
End Sub

Public Sub LoadMunicipality(Region As String, Province As String)
    Dim i As Integer
    Dim varlist
        Dim rst As ADODB.Recordset
        
            Set rst = New ADODB.Recordset
                rst.Open "Select name,mun from psgc where reg='" & Region & "' and prov='" & Province & "' and mun<>'00' and brgy='000' order by name", cnn, adOpenStatic, adLockOptimistic
            
            Me.LstMunicipality.ListItems.Clear
        For i = 1 To rst.RecordCount
            Set varlist = Me.LstMunicipality.ListItems.Add
            varlist.SubItems(1) = StrConv(rst("name").Value, vbUpperCase)
            varlist.Text = rst("mun").Value
            rst.MoveNext
        Next
End Sub

Public Sub LoadBarangay(Region As String, Province As String, Municipality As String)
    Dim i As Integer
    Dim varlist
        Dim rst As ADODB.Recordset
        
            Set rst = New ADODB.Recordset
                rst.Open "Select name,brgy from psgc where reg='" & Region & "' and prov='" & Province & "' and mun='" & Municipality & "' and brgy<>'000' order by name", cnn, adOpenStatic, adLockOptimistic
            
            Me.LstBarangay.ListItems.Clear
        For i = 1 To rst.RecordCount
            Set varlist = Me.LstBarangay.ListItems.Add
            varlist.SubItems(1) = StrConv(rst("name").Value, vbUpperCase)
            varlist.Text = rst("brgy").Value
            rst.MoveNext
        Next
End Sub

Private Sub LstRegion_ItemClick(ByVal Item As MSComctlLib.ListItem)
LoadProvince Me.LstRegion.SelectedItem.Text
LoadMunicipality Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text
     If Me.LstMunicipality.ListItems.Count > 0 Then
     LoadBarangay Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text, Me.LstMunicipality.SelectedItem.Text
     Else
     Me.LstBarangay.ListItems.Clear
     End If
End Sub

Private Sub LstMunicipality_ItemClick(ByVal Item As MSComctlLib.ListItem)
LoadBarangay Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text, Me.LstMunicipality.SelectedItem.Text
End Sub

Private Sub LstProvince_ItemClick(ByVal Item As MSComctlLib.ListItem)
 LoadMunicipality Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text
 
 If LstMunicipality.ListItems.Count > 0 Then
    LoadBarangay Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text, Me.LstMunicipality.SelectedItem.Text
    Else
    LstBarangay.ListItems.Clear
 End If
End Sub

Private Sub RaveAddProvince_Click()
    ProvinceEditMode = False
    FrmProvinceLib.Show 1
End Sub

Private Sub RaveEditProvince_Click()
    ProvinceEditMode = True
    FrmProvinceLib.Show 1
End Sub

Private Sub RaveDeleteProvince_Click()
If MsgBox("Are you sure you want to delete this province?" & vbCrLf & "Municipalities and Barangays under this province will also be deleted.", vbYesNo, "DELETE PROVINCE?") = vbNo Then
   Exit Sub
End If
    
    cnn.Execute "delete from psgc where substring(psgc_cd,1,4)='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & Format(Val(FrmPSGCLibrary.LstProvince.SelectedItem.Text), "00") & "'"
    LoadProvince Me.LstRegion.SelectedItem.Text
    LoadMunicipality Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text
    
    If Me.LstMunicipality.ListItems.Count > 0 Then
        LoadBarangay Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text, Me.LstMunicipality.SelectedItem.Text
    Else
        Me.LstBarangay.ListItems.Clear
    End If
End Sub

Private Sub RaveAddMunicipality_Click()
MunicipalityEditMode = False
FrmMunicipalityLib.Show 1
End Sub

Private Sub RaveEditmunicipality_Click()
MunicipalityEditMode = True
FrmMunicipalityLib.Show 1
End Sub

Private Sub RaveDeletemunicipality_Click()
If MsgBox("Are you sure you want to delete this municipality?" & vbCrLf & "Barangays under this Municipality will also be deleted.", vbYesNo, "DELETE MUNICIPALITY?") = vbNo Then
   Exit Sub
End If
    
    cnn.Execute "delete from psgc where substring(psgc_cd,1,6)='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & Format(Val(FrmPSGCLibrary.LstMunicipality.SelectedItem.Text), "00") & "'"
    'LoadProvince Me.LstRegion.SelectedItem.Text
    LoadMunicipality Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text
    
    If Me.LstMunicipality.ListItems.Count > 0 Then
        LoadBarangay Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text, Me.LstMunicipality.SelectedItem.Text
    Else
        Me.LstBarangay.ListItems.Clear
    End If
End Sub

Private Sub RaveAddBrgy_Click()
BrgyEditMode = False
FrmBrgyLib.Show 1
End Sub

Private Sub RaveEditBrgy_Click()
BrgyEditMode = True
FrmBrgyLib.Show 1
End Sub


Private Sub RaveDeleteBrgy_Click()
If MsgBox("Are you sure you want to delete this Barangay?", vbYesNo, "DELETE BARANGAY?") = vbNo Then
   Exit Sub
End If
    
    cnn.Execute "delete from psgc where substring(psgc_cd,1,9)='" & FrmPSGCLibrary.LstRegion.SelectedItem.Text & FrmPSGCLibrary.LstProvince.SelectedItem.Text & FrmPSGCLibrary.LstMunicipality.SelectedItem.Text & FrmPSGCLibrary.LstBarangay.SelectedItem.Text & "'"
    'LoadProvince Me.LstRegion.SelectedItem.Text
    'LoadMunicipality Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text
    
    If Me.LstMunicipality.ListItems.Count > 0 Then
        LoadBarangay Me.LstRegion.SelectedItem.Text, Me.LstProvince.SelectedItem.Text, Me.LstMunicipality.SelectedItem.Text
    Else
        Me.LstBarangay.ListItems.Clear
    End If
    
End Sub

