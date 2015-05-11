VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmTriangulationPrintOut 
   Caption         =   "Triangulation Station Information"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   13215
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton BtnEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
Attribute VB_Name = "FrmTriangulationPrintOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New TriangulationPrintOut

Private Sub BtnEdit_Click()
    FrmRequestingParty.Show 1

    Report.RequestingPartyTextBox.SetText (GetSetting(App.EXEName, "Requesting Party", "Requesting Party"))
    Report.ORNumberTextbox.SetText (GetSetting(App.EXEName, "OR Number", "OR Number"))
    Report.PurposeTextBox.SetText (GetSetting(App.EXEName, "Purpose", "Purpose"))
   ' Report.SignatoryTextBox.SetText (UCase(GetSetting(App.EXEName, "Signatory", "Signatory")))
   ' Report.DesignationTextBox.SetText (GetSetting(App.EXEName, "Designation", "Designation"))
    Report.TNTextBox.SetText (GetSetting(App.EXEName, "TN", "TN"))
    CRViewer1.Refresh
End Sub

Private Sub Form_Load()
Report.Database.Tables(1).SetLogOnInfo "", "gcpds", "gcpds", "gcpds"
Screen.MousePointer = vbHourglass
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
CRViewer1.Zoom 150
Report.PaperSize = crPaperA4

Screen.MousePointer = vbDefault
 Dim rst As New ADODB.Recordset
    
    rst.Open "select * from triangulation where stat_name='" & Replace(Trim(FrmGCPDS.TextBoxOldStationName.Text), "'", "''") & "'", cnn, adOpenStatic
    Report.Database.SetDataSource rst
    
    InitializeCertificate
Report.DiscardSavedData
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0

CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub InitializeCertificate()
    Report.RequestingPartyTextBox.SetText (GetSetting(App.EXEName, "Requesting Party", "Requesting Party"))
    Report.ORNumberTextbox.SetText (GetSetting(App.EXEName, "OR Number", "OR Number"))
    Report.PurposeTextBox.SetText (GetSetting(App.EXEName, "Purpose", "Purpose"))
    Report.TNTextBox.SetText (GetSetting(App.EXEName, "TN", "TN"))
    Report.SignatoryTextBox.SetText UCase((GetSetting(App.EXEName, "Signatory", "Signatory")))
    Report.DesignationTextBox.SetText (GetSetting(App.EXEName, "Designation", "Designation"))
End Sub
