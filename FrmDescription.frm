VERSION 5.00
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmDescription 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Station Description"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   Icon            =   "FrmDescription.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin Rave_Buttons.RaveButtons RaveFontSize 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "A"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDescription.frx":3452
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   540
      Width           =   10680
   End
   Begin Rave_Buttons.RaveButtons RaveSave 
      Height          =   495
      Left            =   3150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Save"
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
      FOCUSR          =   0   'False
      BCOL            =   32768
      BCOLO           =   32768
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmDescription.frx":346E
      PICN            =   "FrmDescription.frx":348A
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RavePrint 
      Height          =   495
      Left            =   5535
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Print"
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
      FOCUSR          =   0   'False
      BCOL            =   32768
      BCOLO           =   32768
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmDescription.frx":39EC
      PICN            =   "FrmDescription.frx":3A08
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveFontSizeSmall 
      Height          =   375
      Left            =   540
      TabIndex        =   4
      Top             =   90
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "A"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDescription.frx":3CF5
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
Attribute VB_Name = "FrmDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

If rstRecords.RecordCount > 0 Then
Dim rst As New ADODB.Recordset
    rst.Open "select description from geoprov where stat_name='" & Replace(rstRecords("stat_name"), "'", "''") & "'", cnn, adOpenStatic, adLockOptimistic
    If rst.RecordCount > 0 Then
        If IsNull(rst!Description) Then
            
            Me.txtDescription = ""
            Else
            Me.txtDescription = rst!Description
        End If
        
    End If
End If
End Sub

Private Sub Form_Resize()
'On Error Resume Next
'            Me.txtDescription.Top = 100
'            Me.txtDescription.Left = 80
'            Me.txtDescription.Height = ScaleHeight - 1000
'            Me.txtDescription.Width = ScaleWidth - 260
'            Me.RaveSave.Top = Me.txtDescription.Height + 400
'            Me.RaveSave.Left = (Me.ScaleWidth / 2) - (Me.RaveSave.Width)
'            Me.RavePrint.Top = Me.txtDescription.Height + 400
'            Me.RavePrint.Left = Me.RaveSave.Left + Me.RaveSave.Width + 100
End Sub



Private Sub RaveFontSize_Click()
    Me.txtDescription.FontSize = Me.txtDescription.FontSize + 1
End Sub

Private Sub RaveFontSizeSmall_Click()
    Me.txtDescription.FontSize = Me.txtDescription.FontSize - 1
End Sub

Private Sub RavePrint_Click()

FrmPrintSelection.Show 1
End Sub

Private Sub RaveSave_Click()
cnn.Execute "update geoprov set description='" & Replace(Me.txtDescription, "'", "''") & "' where stat_name='" & Replace(rstRecords("stat_name"), "'", "''") & "'"
Me.RaveSave.Enabled = False
Unload Me
End Sub

Private Sub txtDescription_Change()
If Me.RaveSave.Enabled = False Then
    Me.RaveSave.Enabled = True
End If
End Sub
