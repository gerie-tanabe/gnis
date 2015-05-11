VERSION 5.00
Object = "{4F0E71DF-2B8E-4193-904B-C964443BD659}#8.0#0"; "NeoCalendarIII.ocx"
Begin VB.Form FrmCalendar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar"
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   435
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin NeoCalendarIII.MonthCalendar Calendar 
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   4763
      BackgroundImage =   "FrmCalendar.frx":0000
      Value           =   38693
      HeaderColor     =   12632256
      BackColor       =   16777215
      EmptyButtonCaption=   "None"
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Data            =   "FrmCalendar.frx":BEB6
   End
End
Attribute VB_Name = "FrmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()

End Sub

Private Sub Calendar_Change()
    FrmGCPDS.ActiveControl.Text = Format(Calendar.Value, "MMM d, yyyy")
    'MsgBox FrmGCPDS.ActiveControl.Name
End Sub

Private Sub Calendar_DateClicked()
    Me.Hide
End Sub

Private Sub Calendar_EmptyClicked()
    FrmGCPDS.ActiveControl.Text = ""
    Me.Hide
End Sub

Private Sub Form_Activate()
    If Trim(FrmGCPDS.ActiveControl.Text) <> "" Then
        Calendar.Value = CDate(FrmGCPDS.ActiveControl)
        Else
        Calendar.Value = Date
    End If
End Sub

