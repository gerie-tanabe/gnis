VERSION 5.00
Begin VB.Form FrmIdentity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attributes"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   Icon            =   "FrmIdentity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4335
      Begin VB.TextBox txtStationName 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtProvince 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtMunicipality 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtbarangay 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtLongitude 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Txtlatitude 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Station Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Municipality"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   585
      End
   End
   Begin VB.CommandButton CmdDetails 
      Caption         =   "Details"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "FrmIdentity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDetails_Click()
    Station = Trim(Me.txtStationName.Text)
    rstRecords.MoveFirst
    rstRecords.Find "Stat_name='" & Replace(Station, "'", "''") & "'", , adSearchForward
    FrmGCPDS.Caption = "Geodetic Control Points Databanking System - " & Format(rstRecords.AbsolutePosition, "#,##0") & " of " & Format(rstRecords.RecordCount, "#,##0") & " Records"
    FillUp
    Me.Hide
    FrmGCPDS.SSTab1.Tab = 0
End Sub

Private Sub Form_Activate()
For Each fld In recs.Fields  ' iterate over the fields
        MsgBox fld.name
      If fld.name = "STAT_NAME" Then
         FrmIdentity.txtStationName = fld
      End If
'      If fld.Name = "province" Then
'         FrmIdentity.txtProvince = GetProvinceName(CStr(fld))
'      End If
'      If fld.Name = "municipal" Then
'         FrmIdentity.txtMunicipality = fld
'      End If
'      If fld.Name = "barangay" Then
'         FrmIdentity.txtbarangay = fld
'      End If
      If fld.name = "LONGITUDE" Then
         FrmIdentity.txtLongitude = fld
      End If
      If fld.name = "LATITUDE" Then
         FrmIdentity.Txtlatitude = fld
      End If
    Next fld
End Sub

