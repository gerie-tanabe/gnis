VERSION 5.00
Object = "{6FD34FAD-BA34-4E74-BB92-B9F0BB900FB9}#5.0#0"; "RaveButtons.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form FrmDetectLocation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detect Location"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin CCRProgressBar6.ccrpProgressBar ccrpProgressBar1 
      Height          =   330
      Left            =   270
      Top             =   180
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   582
      Appearance      =   2
      AutoCaption     =   1
      Caption         =   "0%"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin Rave_Buttons.RaveButtons RaveButtons1 
      Height          =   375
      Left            =   1935
      TabIndex        =   0
      Top             =   675
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "RaveButtons1"
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetectLocation.frx":0000
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
Attribute VB_Name = "FrmDetectLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RaveButtons1_Click()
Dim rst As New ADODB.Recordset
Dim i As Long

rst.Open "select stat_name,region,province from geoprov", cnn, adOpenStatic

Me.ccrpProgressBar1.Min = 0
Me.ccrpProgressBar1.Max = rst.RecordCount

For i = 1 To rst.RecordCount
    DoEvents
    Me.ccrpProgressBar1.Value = i
    
    
 
  If IsNull(rst!Province) Or Trim(rst!Province) = "" Then
     cnn.Execute "update geoprov set province='" & getprovince(Left(rst!Stat_Name, 3)) & "' where stat_name='" & Replace(rst!Stat_Name, "'", "''") & "'"
  End If
  
  If IsNull(rst!Region) Or Trim(rst!Region) = "" Then
     cnn.Execute "update geoprov set region='" & GetRegion(Left(rst!Stat_Name, 3)) & "' where stat_name='" & Replace(rst!Stat_Name, "'", "''") & "'"
  End If
  
  

'
   rst.MoveNext
    
    
Next

End Sub


Public Function getprovince(Code As String) As String
    Dim rst As New ADODB.Recordset
    
    rst.Open "select name from psgc where acronym='" & Code & "'", cnn, adOpenStatic
    
    If rst.RecordCount > 0 Then
        getprovince = rst!name
       Else
        getprovince = ""
    End If
    
    
End Function

Public Function GetRegion(Code As String) As String
    Dim rst As New ADODB.Recordset
    
    rst.Open "SELECT REGION.name AS RegionName FROM psgc LEFT OUTER JOIN psgc AS REGION ON SUBSTRING(psgc.psgc_cd, 1, 2) + '0000000' = REGION.psgc_cd where psgc.ACRONYM='" & Code & "'", cnn, adOpenStatic
    
    If rst.RecordCount > 0 Then
        GetRegion = rst!RegionName
       Else
        GetRegion = ""
    End If
    
    
End Function


