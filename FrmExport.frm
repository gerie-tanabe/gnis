VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "ravebuttons.ocx"
Begin VB.Form FrmExport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialogExport 
      Left            =   5400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Export"
   End
   Begin MSComctlLib.ListView LstRegion 
      Height          =   6225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   10980
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Region"
         Object.Width           =   8819
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons RaveExport 
      Height          =   465
      Left            =   2745
      TabIndex        =   1
      Top             =   6525
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmExport.frx":0000
      PICN            =   "FrmExport.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveCancel 
      Height          =   465
      Left            =   1080
      TabIndex        =   2
      Top             =   6525
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmExport.frx":10AE
      PICN            =   "FrmExport.frx":10CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.LstRegion.ListItems.Add 1, , "Region I"
Me.LstRegion.ListItems.Add 2, , "Region II"
Me.LstRegion.ListItems.Add 3, , "Region III"
Me.LstRegion.ListItems.Add 4, , "Region IV-A"
Me.LstRegion.ListItems.Add 5, , "Region IV-B"
Me.LstRegion.ListItems.Add 6, , "Region V"
Me.LstRegion.ListItems.Add 7, , "Region VI"
Me.LstRegion.ListItems.Add 8, , "Region VII"
Me.LstRegion.ListItems.Add 9, , "Region VIII"
Me.LstRegion.ListItems.Add 10, , "Region IX"
Me.LstRegion.ListItems.Add 11, , "Region X"
Me.LstRegion.ListItems.Add 12, , "Region XI"
Me.LstRegion.ListItems.Add 13, , "Region XII"
Me.LstRegion.ListItems.Add 14, , "Region XIII"
Me.LstRegion.ListItems.Add 15, , "ARMM"
Me.LstRegion.ListItems.Add 16, , "CAR"
Me.LstRegion.ListItems.Add 17, , "NCR"
Me.LstRegion.ListItems.Add 18, , "ALL"

End Sub


Private Sub RaveAddControl_Click()

End Sub

Private Sub RaveCancel_Click()
Unload Me
End Sub

Private Sub RaveExport_Click()
On Error GoTo Hell
    Dim rst As New ADODB.Recordset
    
    If Me.LstRegion.SelectedItem.Text = "ALL" Then
        rst.Open "Select * from geoprov ", cnn, adOpenStatic
        Else
        rst.Open "Select * from geoprov where region='" & Me.LstRegion.SelectedItem.Text & "'", cnn, adOpenStatic
    End If
       
    Me.CommonDialogExport.InitDir = App.Path
    Me.CommonDialogExport.Filter = "XML File|*.XML"
    Me.CommonDialogExport.filename = Me.LstRegion.SelectedItem.Text & ".xml"
    Me.CommonDialogExport.ShowSave
    
    If Dir(Me.CommonDialogExport.filename) <> "" Then
        Kill Me.CommonDialogExport.filename
    End If
    
rst.Save Me.CommonDialogExport.filename, adPersistXML

MsgBox "Export Successful", vbInformation, "Export"

Exit Sub

Hell:
If Err.Number = 32755 Then
    Else
    MsgBox "Error."
    
End If

End Sub
