VERSION 5.00
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form FrmWGS84 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Extract WGS84 Coordinates"
   ClientHeight    =   10290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin CCRProgressBar6.ccrpProgressBar ccrpProgressBar1 
      Height          =   810
      Left            =   3960
      Top             =   4320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1429
      Appearance      =   2
      AutoCaption     =   1
      BackColor       =   32768
      Caption         =   "0%"
      FillColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReverseFill     =   -1  'True
      Smooth          =   -1  'True
   End
End
Attribute VB_Name = "FrmWGS84"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim rst As New ADODB.Recordset
Dim i As Integer

rst.Open "select  stat_name,description from geoprov where h_ref='PRS92'", cnn, adOpenStatic


Me.ccrpProgressBar1.Min = 0
Me.ccrpProgressBar1.Max = 1000

For i = 1 To rst.RecordCount
    'Me.Caption = rst.RecordCount - i
    ExtractWGS84 rst("stat_name"), rst("description")
    Me.ccrpProgressBar1.Value = (1000 / rst.RecordCount) * i
    
    rst.MoveNext
    DoEvents
Next


Unload Me
End Sub

Private Sub Form_Load()
    TranslucentForm Me, 200
End Sub

Private Sub Form_Resize()
Me.ccrpProgressBar1.Width = Me.Width / 2
Me.ccrpProgressBar1.Left = Me.Width / 2 - Me.ccrpProgressBar1.Width / 2
'Me.ccrpProgressBar1.Width = Me.Width
Me.ccrpProgressBar1.Top = Me.Height / 2
End Sub
