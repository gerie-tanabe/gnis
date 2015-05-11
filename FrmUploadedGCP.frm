VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Begin VB.Form FrmUploadedGCP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "uploaded geodetic control points"
   ClientHeight    =   10950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10260
      Top             =   10305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUploadedGCP.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUploadedGCP.frx":0513
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUploadedGCP.frx":0A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUploadedGCP.frx":0F47
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUploadedGCP.frx":1464
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUploadedGCP.frx":1986
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   585
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   " Region I            "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":1F08
      PICN            =   "FrmUploadedGCP.frx":1F24
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   9420
      Left            =   2700
      TabIndex        =   0
      Top             =   765
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   16616
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date Sent"
         Object.Width           =   5292
      EndProperty
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   600
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   1170
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   1058
      BTYPE           =   3
      TX              =   "  Region II            "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":23A4
      PICN            =   "FrmUploadedGCP.frx":23C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   2
      Left            =   45
      TabIndex        =   3
      Top             =   1800
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region III          "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":2812
      PICN            =   "FrmUploadedGCP.frx":282E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   3
      Left            =   45
      TabIndex        =   4
      Top             =   2385
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region IV-A      "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":2D1C
      PICN            =   "FrmUploadedGCP.frx":2D38
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   4
      Left            =   45
      TabIndex        =   5
      Top             =   2970
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region IV-B      "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":324E
      PICN            =   "FrmUploadedGCP.frx":326A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   5
      Left            =   45
      TabIndex        =   6
      Top             =   3555
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region V           "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":34DE
      PICN            =   "FrmUploadedGCP.frx":34FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   6
      Left            =   45
      TabIndex        =   7
      Top             =   4140
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region VI          "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":3995
      PICN            =   "FrmUploadedGCP.frx":39B1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   7
      Left            =   45
      TabIndex        =   8
      Top             =   4725
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region VII         "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":3E6A
      PICN            =   "FrmUploadedGCP.frx":3E86
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   8
      Left            =   45
      TabIndex        =   9
      Top             =   5310
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region VIII        "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":4334
      PICN            =   "FrmUploadedGCP.frx":4350
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   9
      Left            =   45
      TabIndex        =   10
      Top             =   5895
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region IX          "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":480C
      PICN            =   "FrmUploadedGCP.frx":4828
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   10
      Left            =   45
      TabIndex        =   11
      Top             =   6480
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region X           "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":4D61
      PICN            =   "FrmUploadedGCP.frx":4D7D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   11
      Left            =   45
      TabIndex        =   12
      Top             =   7065
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region XI          "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":5304
      PICN            =   "FrmUploadedGCP.frx":5320
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   12
      Left            =   45
      TabIndex        =   13
      Top             =   7650
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Region XII         "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":5810
      PICN            =   "FrmUploadedGCP.frx":582C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   13
      Left            =   45
      TabIndex        =   14
      Top             =   9405
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "CARAGA             "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":5D4D
      PICN            =   "FrmUploadedGCP.frx":5D69
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   14
      Left            =   45
      TabIndex        =   15
      Top             =   8235
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   " ARMM                 "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":6232
      PICN            =   "FrmUploadedGCP.frx":624E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   15
      Left            =   45
      TabIndex        =   16
      Top             =   8820
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "CAR                     "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":64D2
      PICN            =   "FrmUploadedGCP.frx":64EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveRegion 
      Height          =   555
      Index           =   16
      Left            =   45
      TabIndex        =   17
      Top             =   9990
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "NCR                    "
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":6992
      PICN            =   "FrmUploadedGCP.frx":69AE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveView 
      Height          =   330
      Left            =   2700
      TabIndex        =   19
      Top             =   10305
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "View Record"
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":6E80
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Rave_Buttons.RaveButtons RaveCommit 
      Height          =   330
      Left            =   4545
      TabIndex        =   20
      Top             =   10305
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "Commit"
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmUploadedGCP.frx":6E9C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LabelRegion 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Region I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   2700
      TabIndex        =   18
      Top             =   315
      Width           =   8295
   End
End
Attribute VB_Name = "FrmUploadedGCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LoadRegion(Region As String)
    Dim rstUploaded As New ADODB.Recordset
    Dim i As Integer
    Dim varlist
    
    rstUploaded.Open "Select * from sent_table where region='" & Region & "' order by datesent desc", cnn, adOpenStatic, adLockOptimistic
    
    Me.ListView.ListItems.Clear
    
    For i = 1 To rstUploaded.RecordCount
    Set varlist = Me.ListView.ListItems.Add
    
        
        
       varlist.SmallIcon = rstUploaded("h_order").Value + 1
       
       
       varlist.Text = IIf(IsNull(rstUploaded("Stat_name")), "", rstUploaded("Stat_Name"))
       varlist.SubItems(1) = IIf(IsNull(rstUploaded("Datesent")), "", Format(rstUploaded("Datesent"), "mmm dd, yyyy - hh:mm am/pm"))
       
   
    
        rstUploaded.MoveNext
        
    Next
    
End Sub


Private Sub Form_Load()
LoadRegion "Region I"
End Sub

Private Sub ListView_DblClick()
CurrentStationToView = Me.ListView.SelectedItem.Text


FrmStationDetails.Show 1

            
End Sub

Private Sub RaveButtons1_Click()

End Sub

Private Sub RaveButtons2_Click()

End Sub

Private Sub RaveCommit_Click()
Dim i As Integer



For i = 1 To Me.ListView.ListItems.Count
    If Me.ListView.ListItems(i).Selected = True Then
        If If_Already_Exist(Replace(Me.ListView.ListItems(i).Text, "'", "''")) = True Then
                If MsgBox(Me.ListView.ListItems(i).Text & " station already exist do you want to overwrite?", vbYesNo, "gnis") = vbYes Then
                End If
            Else
                
         End If
    End If
Next

End Sub

Private Sub RaveRegion_Click(Index As Integer)
If Index = 0 Then
    Me.LabelRegion.Caption = "Region I"
    LoadRegion "Region I"
ElseIf Index = 1 Then
    Me.LabelRegion.Caption = "Region II"
    LoadRegion "Region II"
ElseIf Index = 2 Then
    Me.LabelRegion.Caption = "Region III"
    LoadRegion "Region III"
ElseIf Index = 3 Then
    Me.LabelRegion.Caption = "Region IV-A"
    LoadRegion "Region IV-A"
    ElseIf Index = 4 Then
    Me.LabelRegion.Caption = "Region IV-B"
    LoadRegion "Region IV-B"
    ElseIf Index = 5 Then
    Me.LabelRegion.Caption = "Region V"
    LoadRegion "Region V"
    ElseIf Index = 6 Then
    Me.LabelRegion.Caption = "Region VI"
    LoadRegion "Region VI"
    ElseIf Index = 7 Then
    Me.LabelRegion.Caption = "Region VII"
    LoadRegion "Region VII"
    ElseIf Index = 8 Then
    Me.LabelRegion.Caption = "Region VIII"
    LoadRegion "Region VIII"
    ElseIf Index = 9 Then
    Me.LabelRegion.Caption = "Region IX"
    LoadRegion "Region IX"
    ElseIf Index = 10 Then
    Me.LabelRegion.Caption = "Region X"
    LoadRegion "Region X"
    ElseIf Index = 11 Then
    Me.LabelRegion.Caption = "Region XI"
    LoadRegion "Region XI"
    ElseIf Index = 12 Then
    Me.LabelRegion.Caption = "Region XII"
    LoadRegion "Region XII"
    ElseIf Index = 13 Then
    Me.LabelRegion.Caption = "CARAGA"
    LoadRegion "Region XIII"
    ElseIf Index = 14 Then
    Me.LabelRegion.Caption = "ARMM"
    LoadRegion "ARMM"
    ElseIf Index = 15 Then
    Me.LabelRegion.Caption = "CAR"
    LoadRegion "CAR"
    ElseIf Index = 16 Then
    Me.LabelRegion.Caption = "NCR"
    LoadRegion "NCR"
    
End If
End Sub

Private Sub RaveRegiond_Click()

End Sub

Private Sub RaveView_Click()
Dim i As Integer
For i = 1 To Me.ListView.ListItems.Count
    If Me.ListView.ListItems(i).Selected = True Then
        CurrentStationToView = Me.ListView.ListItems(i).Text
        FrmStationDetails.Show 1
    End If
Next
End Sub
