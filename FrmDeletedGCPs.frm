VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmDeletedGCPs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deleted GCP's Recycle Bin"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   Icon            =   "FrmDeletedGCPs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11775
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView LstDeleted 
      Height          =   6285
      Left            =   240
      TabIndex        =   0
      Top             =   420
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   11086
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Region"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Province"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Municipality"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Barangay"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date Deleted"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Deleted By"
         Object.Width           =   5292
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons RaveRestore 
      Height          =   510
      Left            =   5010
      TabIndex        =   1
      ToolTipText     =   "Add Location"
      Top             =   7020
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   900
      BTYPE           =   3
      TX              =   "Restore"
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
      BCOL            =   11292960
      BCOLO           =   11292960
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmDeletedGCPs.frx":0CCA
      PICN            =   "FrmDeletedGCPs.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmDeletedGCPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstDeleted As ADODB.Recordset
Public Sub FillDeletedGCPs()
    Dim i As Integer
    Dim varlist
    
    Me.LstDeleted.ListItems.Clear
    
    For i = 1 To rstDeleted.RecordCount
    Set varlist = Me.LstDeleted.ListItems.Add
        varlist.Text = IIf(IsNull(rstDeleted("Stat_name")), "", rstDeleted("Stat_Name"))
        varlist.SubItems(1) = IIf(IsNull(rstDeleted("Region")), "", rstDeleted("region"))
        varlist.SubItems(2) = IIf(IsNull(rstDeleted("Province")), "", rstDeleted("Province"))
        varlist.SubItems(3) = IIf(IsNull(rstDeleted("Municipal")), "", rstDeleted("Municipal"))
        varlist.SubItems(4) = IIf(IsNull(rstDeleted("Barangay")), "", rstDeleted("Barangay"))
        varlist.SubItems(5) = IIf(IsNull(rstDeleted("Date_Deleted")), "", rstDeleted("Date_Deleted"))
        varlist.SubItems(6) = IIf(IsNull(rstDeleted("Deleted_By")), "", StrConv(rstDeleted("Deleted_By"), vbProperCase))
        rstDeleted.MoveNext
    Next
    
    If rstDeleted.RecordCount > 0 Then
        rstDeleted.MoveFirst
    End If
    
End Sub

Public Sub LoadDeletedGCPs()
    
    Set rstDeleted = New ADODB.Recordset
    rstDeleted.Open "Select * from deleted order by date_deleted", cnn, adOpenStatic, adLockOptimistic
    FillDeletedGCPs
End Sub

Private Sub Form_Load()
    LoadDeletedGCPs
End Sub



Private Sub LstDeleted_Click()
If LstDeleted.ListItems.Count > 0 Then
rstDeleted.AbsolutePosition = Me.LstDeleted.SelectedItem.Index
End If
End Sub

Private Sub RaveRestore_Click()
If LstDeleted.ListItems.Count = 0 Then
    Exit Sub
End If

cnn.Execute "Insert into geoprov (stat_name,stat_new,island,region,province,municipal,barangay,h_order,d_lat,m_lat,s_lat,d_long,m_long,s_long,ell_hgt,wgs84ED,wgs84EM,wgs84ES,wgs84ND,wgs84NM,wgs84NS,ellipz,h_date_ety,h_date_com,date_las_r,date_est,mark_const,mark_type,mark_stat,h_fix,hor_authty,authority,h_ref)" _
             & " values('" & Replace(rstDeleted("stat_name"), "'", "''") & "'" & "," & IIf(IsNull(rstDeleted("stat_new")), "Null", rstDeleted("stat_new")) & "," & "'" & Replace(rstDeleted("island"), "'", "''") & "'" & "," & "'" & Replace(rstDeleted("region"), "'", "''") & "'" & "," & "'" & Replace(rstDeleted("province"), "'", "''") & "'" & "," & "'" & Replace(rstDeleted("municipal"), "'", "''") & "'" & "," & "'" & Replace(rstDeleted("barangay"), "'", "''") & "'" & "," & IIf(IsNull(rstDeleted("h_order")), "Null", rstDeleted("h_order")) & "," & IIf(IsNull(rstDeleted("d_lat")), "Null", rstDeleted("d_lat")) & "," & IIf(IsNull(rstDeleted("m_lat")), "Null", rstDeleted("m_lat")) & "," & IIf(IsNull(rstDeleted("s_lat")), "Null", rstDeleted("s_lat")) _
             & "," & IIf(IsNull(rstDeleted("d_long")), "Null", rstDeleted("d_long")) & "," & IIf(IsNull(rstDeleted("m_long")), "Null", rstDeleted("m_long")) & "," & IIf(IsNull(rstDeleted("s_long")), "Null", rstDeleted("s_long")) & "," & IIf(IsNull(rstDeleted("ell_hgt")), "Null", rstDeleted("ell_hgt")) & "," & IIf(IsNull(rstDeleted("wgs84ED")), "Null", rstDeleted("wgs84ED")) & "," & IIf(IsNull(rstDeleted("wgs84EM")), "Null", rstDeleted("wgs84EM")) & "," & IIf(IsNull(rstDeleted("wgs84ES")), "Null", rstDeleted("wgs84ES")) _
             & "," & IIf(IsNull(rstDeleted("wgs84ND")), "Null", rstDeleted("wgs84ND")) & "," & IIf(IsNull(rstDeleted("wgs84NM")), "Null", rstDeleted("wgs84NM")) & "," & IIf(IsNull(rstDeleted("wgs84NS")), "Null", rstDeleted("wgs84NS")) & "," & IIf(IsNull(rstDeleted("ellipz")), "Null", rstDeleted("ellipz")) & "," & IIf(IsNull(rstDeleted("h_date_ety")), "Null", "'" & rstDeleted("h_date_ety") & "'") & "," & IIf(IsNull(rstDeleted("h_date_com")), "Null", "'" & rstDeleted("h_date_com") & "'") & "," & IIf(IsNull(rstDeleted("date_las_r")), "Null", "'" & rstDeleted("date_las_r") & "'") & "," & IIf(IsNull(rstDeleted("date_est")), "Null", "'" & rstDeleted("date_est") & "'") & "," & IIf(IsNull(rstDeleted("mark_const")), "Null", rstDeleted("mark_const")) _
             & "," & IIf(IsNull(rstDeleted("mark_type")), "Null", rstDeleted("mark_type")) & "," & IIf(IsNull(rstDeleted("mark_stat")), "Null", rstDeleted("mark_stat")) & "," & IIf(IsNull(rstDeleted("h_fix")), "Null", rstDeleted("h_fix")) & "," & "'" & Replace(rstDeleted("hor_authty"), "'", "''") & "'" & "," & "'" & Replace(rstDeleted("authority"), "'", "''") & "'" & "," & "'" & Replace(rstDeleted("h_ref"), "'", "''") & "'" & ")"

rstDeleted.Delete
rstDeleted.Requery
FillDeletedGCPs


End Sub
