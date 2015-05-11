VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmSummary 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Summary of GCPs"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   13080
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView LstSummary 
      Height          =   6285
      Left            =   840
      TabIndex        =   0
      Top             =   1560
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Province"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "1st Order"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "2nd Order"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "3rd Order"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "4th Order"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rst As New ADODB.Recordset
Dim varlist
Dim i As Long
rst.Open "SELECT province,(SELECT COUNT(*) AS Count FROM geoprov AS geoprov_5 WHERE (province = Parent.province) AND (h_order = 0) AND (h_ref = 'PRS92') ) AS [0 Order]," & _
                          "(SELECT     COUNT(*) AS Count FROM geoprov AS geoprov_4 WHERE (province = Parent.province) AND (h_order = 1)AND (h_ref = 'PRS92')) AS [1st Order]," & _
                          "(SELECT     COUNT(*) AS Count FROM geoprov AS geoprov_3 WHERE (province = Parent.province) AND (h_order = 2)AND (h_ref = 'PRS92')) AS [2nd Order]," & _
                          "(SELECT     COUNT(*) AS Count  FROM geoprov AS geoprov_2 WHERE (province = Parent.province) AND (h_order = 3)AND (h_ref = 'PRS92')) AS [3rd Order]," & _
                          "(SELECT     COUNT(*) AS Count  FROM geoprov AS geoprov_1 WHERE (province = Parent.province) AND (h_order = 4)AND (h_ref = 'PRS92')) AS [4th Order] " & _
                          " FROM geoprov AS Parent GROUP BY region, province ORDER BY  province", cnn, adOpenStatic

If rst.RecordCount > 0 Then
    For i = 1 To rst.RecordCount
        Set varlist = Me.LstSummary.ListItems.Add
            varlist.Text = IIf(IsNull(rst("Province")), "", rst("Province"))
            varlist.SubItems(1) = rst("1st Order")
            varlist.SubItems(2) = rst("2nd Order")
            varlist.SubItems(3) = rst("3rd Order")
            varlist.SubItems(4) = rst("4th Order")
            rst.MoveNext
    Next
End If


End Sub
