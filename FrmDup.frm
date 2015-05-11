VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Duplicate Records"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstDup 
      Height          =   7650
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   13494
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Region"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Province"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Municipality"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Barangay"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Elevation"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "FrmDup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rst As New ADODB.Recordset
rst.Open "SELECT * from benchmarks where stat_name='" & Replace(rstBenchmarks!Stat_Name, "'", "''") & "' AND ucode<>" & rstBenchmarks!ucode, cnn, adOpenStatic
Me.LstDup.ListItems.Clear
Dim varlist
Dim i As Integer
If rst.RecordCount > 0 Then
 
 For i = 1 To rst.RecordCount
 Set varlist = Me.LstDup.ListItems.Add
 varlist.Text = rst!Stat_Name
 varlist.SubItems(1) = rst!Region
 varlist.SubItems(2) = rst!Province
 varlist.SubItems(3) = rst!Municipal
 varlist.SubItems(4) = rst!Barangay
 varlist.SubItems(5) = IIf(IsNull(rst!Elevation), "", rst!Elevation)
 varlist.SubItems(6) = rst!Description
 rst.MoveNext
  Next
End If


End Sub

