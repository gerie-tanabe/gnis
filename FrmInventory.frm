VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInventory 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory of certificates"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   151846912
      CurrentDate     =   40941
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   151846912
      CurrentDate     =   40941
   End
   Begin MSComctlLib.ListView LstInventory 
      Height          =   5790
      Left            =   225
      TabIndex        =   0
      Top             =   720
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   10213
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date/Time"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Barcode ID"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Station Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Print by"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "OR No."
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Transaction No."
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Requesting Party"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Signatory"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Sequence"
         Object.Width           =   2540
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons RaveDelete 
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   6720
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      BTYPE           =   3
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmInventory.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LabelRecordCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "FrmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub LoadPrintInventory()

    Dim rst As New ADODB.Recordset
    rst.Open "Select * from print_inventory where [date] BETWEEN '" & Me.DTPicker1.Value & "' And '" & Me.DTPicker2.Value & "'  order by date desc", cnn, adOpenStatic, adLockOptimistic
    Me.LstInventory.ListItems.Clear
    'Debug.Print ("Select * from print_inventory where [date] BETWEEN '" & Me.DTPicker1.Value & "' And '" & Me.DTPicker2.Value & "'  order by date desc" + vbCrLf)
    
    
    If rst.RecordCount > 0 Then
        Dim i As Integer
        Dim varlist
        For i = 1 To rst.RecordCount
            Set varlist = Me.LstInventory.ListItems.Add
                varlist.Text = Format(rst("Date").Value, "dd-MMM-yyyy hh:mm:ss AM/PM")
                varlist.SubItems(1) = rst!ID
                varlist.SubItems(2) = rst!Stat_Name
                varlist.SubItems(3) = rst!print_by
                varlist.SubItems(4) = IIf(IsNull(rst!orno), "", rst!orno)
                varlist.SubItems(5) = IIf(IsNull(rst!tn), "", rst!tn)
                varlist.SubItems(6) = IIf(IsNull(rst!requesting_party), "", rst!requesting_party)
                varlist.SubItems(7) = IIf(IsNull(rst!signatory), "", rst!signatory)
                varlist.SubItems(8) = IIf(IsNull(rst!seq), "", rst!seq)
                
                rst.MoveNext
        Next
      Me.LabelRecordCount.Caption = "Records: " & rst.RecordCount
       rst.Close
    End If


    
End Sub


Private Sub DTPicker1_Change()
Me.LabelRecordCount.Caption = "No records"
Call LoadPrintInventory
End Sub

Private Sub DTPicker2_Change()
Me.LabelRecordCount.Caption = "No records"
Call LoadPrintInventory
End Sub

Private Sub Form_Load()
DTPicker1.Value = Now
DTPicker2.Value = Now
    Call LoadPrintInventory
End Sub

Private Sub RaveOK_Click()
    
    Unload Me
    
End Sub

Private Sub RaveDelete_Click()
If Me.LstInventory.ListItems.Count = 0 Then
    Exit Sub
End If

If AccessLevel = 2 Then
   FrmAdmin.Show 1
   If TemporaryPass = False Then
        Exit Sub
   End If
End If


cnn.Execute "Delete from print_inventory where id='" & Me.LstInventory.SelectedItem.SubItems(1) & "'"
Me.LstInventory.ListItems.Remove Me.LstInventory.SelectedItem.Index

End Sub
