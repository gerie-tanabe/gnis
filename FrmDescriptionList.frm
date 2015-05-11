VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmDescriptionList 
   Caption         =   "Description"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmDescriptionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport2

Private Sub Form_Load()
Report.Database.Tables(1).SetLogOnInfo "", "", "", "namria"
Screen.MousePointer = vbHourglass
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
Report.DiscardSavedData
CRViewer1.Zoom 100

Dim i As Integer
Dim condition As String
Dim rst As New ADODB.Recordset


For i = 0 To FrmPrintSelection.PrintListBox.ListCount - 1
    
        condition = condition & " stat_name='" & FrmPrintSelection.PrintListBox.List(i) & "'"
      
       
       If i <> FrmPrintSelection.PrintListBox.ListCount - 1 Then
          condition = condition & " or"
       End If
       
       
    
Next

rst.Open "select * from geoprov where " & condition, cnn, adOpenStatic
Report.Database.SetDataSource rst

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
