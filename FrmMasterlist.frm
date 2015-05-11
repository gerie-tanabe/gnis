VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmMasterlist 
   Caption         =   "Masterlist"
   ClientHeight    =   11385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   11385
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmMasterlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New Masterlist_fat

Private Sub Form_Load()
Report.Database.Tables(1).SetLogOnInfo "", "", "gcpds", "gcpds"

Screen.MousePointer = vbHourglass
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
Report.DiscardSavedData
Report.PaperSize = crPaperLetter
CRViewer1.Zoom 150

Dim rst As New ADODB.Recordset
rst.Open "SELECT region,province,(SELECT     COUNT(*) AS Count FROM geoprov AS geoprov_4 WHERE (province = Parent.province) AND (h_order = 1)AND (h_ref = 'PRS92') ) AS [1st Order]," & _
                          "(SELECT     COUNT(*) AS Count FROM geoprov AS geoprov_3 WHERE (province = Parent.province) AND (h_order = 2)AND (h_ref = 'PRS92')) AS [2nd Order]," & _
                          "(SELECT     COUNT(*) AS Count  FROM geoprov AS geoprov_2 WHERE (province = Parent.province) AND (h_order = 3)AND (h_ref = 'PRS92')) AS [3rd Order]," & _
                          "(SELECT     COUNT(*) AS Count  FROM geoprov AS geoprov_1 WHERE (province = Parent.province) AND (h_order = 4)AND (h_ref = 'PRS92')) AS [4th Order], " & _
                          "(SELECT     COUNT(*) AS Count FROM geoprov AS geoprov_5 WHERE (province = Parent.province) AND (h_order = 5)AND (h_ref = 'PRS92') ) AS [AGN]," & _
                          "(SELECT     COUNT(*) AS Count FROM geoprov AS geoprov_6 WHERE (province = Parent.province) AND (h_order = 0)AND (h_ref = 'PRS92') ) AS [0 Order] " & _
                          " FROM geoprov AS Parent   GROUP BY region, province", cnn, adOpenStatic
'rst.Open "SELECT region,province,(SELECT     COUNT(*) AS Count FROM geoprov AS geoprov_4 WHERE (province = Parent.province) AND (h_order = 1)AND (h_ref = 'PRS92') ) AS [1st Order]," & _
                          "(SELECT     COUNT(*) AS Count FROM geoprov AS geoprov_3 WHERE (province = Parent.province) AND (h_order = 2)AND (h_ref = 'PRS92')) AS [2nd Order]," & _
                          "(SELECT     COUNT(*) AS Count  FROM geoprov AS geoprov_2 WHERE (province = Parent.province) AND (h_order = 3)AND (h_ref = 'PRS92')) AS [3rd Order]," & _
                          "(SELECT     COUNT(*) AS Count  FROM geoprov AS geoprov_1 WHERE (province = Parent.province) AND (h_order = 4)AND (h_ref = 'PRS92')) AS [4th Order] " & _
                          " FROM geoprov AS Parent   GROUP BY region, province ORDER BY  region,province", cnn, adOpenStatic
'where (h_date_ety < '01/01/2007')
Report.Database.SetDataSource rst


End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

