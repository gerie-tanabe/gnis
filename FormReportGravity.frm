VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FormReportGravity 
   Caption         =   "Gravity"
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   14085
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Rave_Buttons.RaveButtons Rave 
      Height          =   285
      Left            =   8160
      TabIndex        =   2
      Top             =   0
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   503
      BTYPE           =   7
      TX              =   "Edit..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FormReportGravity.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RavePrint 
      Height          =   330
      Left            =   7800
      TabIndex        =   1
      Top             =   0
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      BTYPE           =   8
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MCOL            =   255
      MPTR            =   1
      MICON           =   "FormReportGravity.frx":001C
      PICN            =   "FormReportGravity.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   13905
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
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FormReportGravity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New Gravity
Dim stationName, orNumber, transactionNumber, requestingParty, signatory, barcodeID, Northing, Easting, Zone As String

Private Sub Form_Load()


    Report.Database.Tables(1).SetLogOnInfo "", "gcpds", "gcpds", "gcpds"
    Screen.MousePointer = vbHourglass

    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport

    Screen.MousePointer = vbDefault
    Report.PaperSize = crPaperA4
    Report.PaperOrientation = crPortrait
    Report.DiscardSavedData
    
    Call SetDataSource
    Call InitializeCertificate
    Call InitializeDetails
    
    
    
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
    CRViewer1.Zoom (100)
End Sub

Private Sub SetDataSource()

    Dim rst As New ADODB.Recordset
    
    rst.Open "select * from gravity where stat_name='" & Replace(Trim(FrmGCPDS.TextBoxGravityName.Text), "'", "''") & "'", cnn, adOpenStatic
    Report.Database.SetDataSource rst
    
       
End Sub

Private Sub InitializeCertificate()
    
      
    Report.RequestingPartyTextBox.SetText (GetSetting(App.EXEName, "Requesting Party", "Requesting Party"))
    Report.ORNumberTextbox.SetText (GetSetting(App.EXEName, "OR Number", "OR Number"))
    Report.PurposeTextBox.SetText (GetSetting(App.EXEName, "Purpose", "Purpose"))
    Report.SignatoryTextBox.SetText UCase((GetSetting(App.EXEName, "Signatory", "Signatory")))
    Report.DesignationTextBox.SetText (GetSetting(App.EXEName, "Designation", "Designation"))
    Report.TNTextBox.SetText (GetSetting(App.EXEName, "TN", "TN"))
    Report.BarCode.SetText "99" & Format(Now, "mmddyyyyhhmmss")
    
End Sub



Private Sub Rave_Click()
FrmRequestingParty.Show 1
    Report.RequestingPartyTextBox.SetText (GetSetting(App.EXEName, "Requesting Party", "Requesting Party"))
    Report.ORNumberTextbox.SetText (GetSetting(App.EXEName, "OR Number", "OR Number"))
    Report.PurposeTextBox.SetText (GetSetting(App.EXEName, "Purpose", "Purpose"))
    Report.SignatoryTextBox.SetText (UCase(GetSetting(App.EXEName, "Signatory", "Signatory")))
    Report.DesignationTextBox.SetText (GetSetting(App.EXEName, "Designation", "Designation"))
    Report.TNTextBox.SetText (GetSetting(App.EXEName, "TN", "TN"))
    CRViewer1.Refresh
End Sub

Private Sub RavePrint_Click()
Call InitializeDetails
    Report.PrintOut True
  
    If Report.PrintingStatus.Progress = crPrintingCompleted Then
    
    'On Error GoTo Hell
       cnn.Execute "Insert into print_inventory (id,[date],stat_name,print_by,orno,tn,requesting_party,signatory) values('" & barcodeID & "','" & Now & "','" & Replace(stationName, "'", "''") & "','" & Replace(Encoder, "'", "''") & "','" & Replace(orNumber, "'", "''") & "','" & Replace(transactionNumber, "'", "''") & "','" & Replace(requestingParty, "'", "''") & "','" & Replace(signatory, "'", "''") & "')"
        
    End If
    
    Exit Sub
Hell:
    MsgBox ("Error. Cannot save transaction")
End Sub

Private Sub InitializeDetails()
    
    barcodeID = "99" & Format(Now, "mmddyyyyhhmmss")
    orNumber = GetSetting(App.EXEName, "OR Number", "OR Number")
    transactionNumber = GetSetting(App.EXEName, "TN", "TN")
    requestingParty = GetSetting(App.EXEName, "Requesting Party", "Requesting Party")
    signatory = GetSetting(App.EXEName, "Signatory", "Signatory")
    
End Sub
