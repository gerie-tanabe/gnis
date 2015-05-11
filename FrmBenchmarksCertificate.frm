VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmBenchmarksCertificate 
   Caption         =   "Benchmarks"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   12570
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Rave_Buttons.RaveButtons RaveButtons4 
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      Top             =   30
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   " A4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmBenchmarksCertificate.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons Rave 
      Height          =   285
      Left            =   7920
      TabIndex        =   2
      Top             =   45
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
      MICON           =   "FrmBenchmarksCertificate.frx":001C
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
      Left            =   7515
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
      MICON           =   "FrmBenchmarksCertificate.frx":0038
      PICN            =   "FrmBenchmarksCertificate.frx":0054
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
      Height          =   8805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
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
Attribute VB_Name = "FrmBenchmarksCertificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New BenchmarkCertificate
Dim Report As New BMCertificate_fat
Dim stationName, orNumber, transactionNumber, requestingParty, signatory, barcodeID As String

Private Sub Form_Load()
    
    Call InitializeCRViewer
    Call SetDataSource
    Call InitializeCertificate
    Call InitializeDetails
    Call BarCode
    
End Sub

Private Sub InitializeCRViewer()

    Report.Database.Tables(1).SetLogOnInfo "", "gcpds", "gcpds", "gcpds"
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport

    Screen.MousePointer = vbDefault
    Report.PaperSize = crPaperA4
    Report.PaperOrientation = crPortrait
    Report.DiscardSavedData

End Sub

Private Sub SetDataSource()

    Dim rst As New ADODB.Recordset
    
    rst.Open "SELECT benchmarks.stat_name, benchmarks.region, benchmarks.province, benchmarks.municipal, benchmarks.barangay, benchmarks.island,benchmarks.elevation, benchmarks.bmplus, benchmarks.latitude, benchmarks.longitude,benchmarks.description, order_lib.description AS Xorder, vertdat.VDDesc FROM benchmarks LEFT OUTER JOIN order_lib ON benchmarks.e_order = order_lib.h_order LEFT OUTER JOIN vertdat ON benchmarks.e_datum = vertdat.VDCode where ucode=" & rstBenchmarks!ucode, cnn, adOpenStatic, adLockOptimistic
    Report.Database.SetDataSource rst
    
    stationName = IIf(IsNull(rst!Stat_Name), "", rst!Stat_Name)
        
End Sub

Private Sub InitializeCertificate()
    
    Report.RequestingPartyTextBox.SetText (GetSetting(App.EXEName, "Requesting Party", "Requesting Party"))
    Report.ORNumberTextBox.SetText (GetSetting(App.EXEName, "OR Number", "OR Number"))
    Report.PurposeTextBox.SetText (GetSetting(App.EXEName, "Purpose", "Purpose"))
    Report.SignatoryTextBox.SetText UCase((GetSetting(App.EXEName, "Signatory", "Signatory")))
    Report.DesignationTextBox.SetText (GetSetting(App.EXEName, "Designation", "Designation"))
    Report.TNTextBox.SetText (GetSetting(App.EXEName, "TN", "TN"))
    
End Sub

Private Sub Form_Resize()
    
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
    CRViewer1.Zoom (100)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub Rave_Click()
FrmRequestingParty.Show 1

    Report.RequestingPartyTextBox.SetText (GetSetting(App.EXEName, "Requesting Party", "Requesting Party"))
    Report.ORNumberTextBox.SetText (GetSetting(App.EXEName, "OR Number", "OR Number"))
    Report.PurposeTextBox.SetText (GetSetting(App.EXEName, "Purpose", "Purpose"))
    Report.SignatoryTextBox.SetText (UCase(GetSetting(App.EXEName, "Signatory", "Signatory")))
    Report.DesignationTextBox.SetText (GetSetting(App.EXEName, "Designation", "Designation"))
    Report.TNTextBox.SetText (GetSetting(App.EXEName, "TN", "TN"))
    CRViewer1.Refresh
End Sub

Private Sub BarCode()
    
    Report.BarCode.SetText barcodeID

End Sub

Private Sub RaveButtons4_Click()
    Report.PaperSize = crPaperA4
    CRViewer1.Refresh
End Sub

Private Sub RavePrint_Click()
 
    Call PrintCertificate
  
End Sub

Private Sub PrintCertificate()
    
    Call InitializeDetails
    Report.PrintOut True
  
    If Report.PrintingStatus.Progress = crPrintingCompleted Then
    
        On Error GoTo Hell
        cnn.Execute "Insert into print_inventory (id,[date],stat_name,print_by,orno,tn,requesting_party,signatory) values('" & barcodeID & "','" & Now & "','" & stationName & "','" & Encoder & "','" & orNumber & "','" & transactionNumber & "','" & requestingParty & "','" & signatory & "')"
        
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



