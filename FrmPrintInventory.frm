VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPrintInventory 
   Caption         =   "Print Inventory"
   ClientHeight    =   13035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16875
   LinkTopic       =   "Form2"
   ScaleHeight     =   13035
   ScaleWidth      =   16875
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CommandGo 
      Caption         =   "Go!"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   152371201
      CurrentDate     =   41102
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   152371201
      CurrentDate     =   41102
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   13365
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   16605
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   195
   End
   Begin VB.Label LlbStation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   345
   End
End
Attribute VB_Name = "FrmPrintInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New PrintInventory

Private Sub CommandGo_Click()
Dim rst As New ADODB.Recordset
rst.Open "Select * from print_inventory where [date] BETWEEN '" & Me.DTPicker1.Value & "' And '" & Me.DTPicker2.Value & "'  order by date desc", cnn, adOpenStatic, adLockOptimistic
Report.Database.SetDataSource rst
CRViewer1.Refresh
Report.Title.SetText ("From " & Format(Me.DTPicker1.Value, "MMM dd, yyyy") & " To " & Format(Me.DTPicker2.Value, "MMM dd, yyyy"))
End Sub

Private Sub DTPicker1_Change()
    
    DTPicker2.MinDate = DTPicker1.Value
End Sub

Private Sub Form_Load()

DTPicker1.MaxDate = Now
DTPicker2.MaxDate = Now
DTPicker1.Value = Now
DTPicker2.Value = Now
DTPicker2.MinDate = DTPicker1.Value

Dim rst As New ADODB.Recordset
rst.Open "Select * from print_inventory where [date]>='" & Me.DTPicker1.Value & "' And [date]<='" & Me.DTPicker2.Value + 1 & "'  order by date desc", cnn, adOpenStatic, adLockOptimistic

Report.Database.SetDataSource rst

Report.Database.Tables(1).SetLogOnInfo "", "gcpds", "gcpds", "gcpds"
Screen.MousePointer = vbHourglass
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault

Report.Title.SetText ("From " & Format(Me.DTPicker1.Value, "MMM dd, yyyy") & " To " & Format(Me.DTPicker2.Value, "MMM dd, yyyy"))

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 700
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
