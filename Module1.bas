Attribute VB_Name = "Module1"
 
Option Explicit



Global DataType As String
Global MapType As String

Global regionTextbox As String
Global provinceTextbox As String
Global municipalityTextbox As String
Global barangayTextbox As String

'Popup Menu
Public Enum BTN_STYLE
    MF_CHECKED = &H8&
    MF_APPEND = &H100&
    TPM_LEFTALIGN = &H0&
    MF_DISABLED = &H2&
    MF_Grayed = &H1&
    MF_SEPARATOR = &H800&
    MF_STRING = &H0&
    TPM_RETURNCMD = &H100&
    TPM_RIGHTBUTTON = &H2&
End Enum

Public Type PointAPI
    x As Long
    y As Long
End Type

Global TemporaryPass As Boolean
Global FlashX, FlashY As Double

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, _
                ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, _
                ByVal hWnd As Long, ByVal lptpm As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
                (ByVal hMenu As Long, _
                ByVal wFlags As BTN_STYLE, ByVal wIDNewItem As Long, _
                ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

'End popup Menu

'<BitMap>
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
   
Public Const COLORONCOLOR As Long = &H3
Public Const HALFTONE As Long = &H4
 '</BitMap>


Global AddQuery As Boolean

Global MyLabelRender As New MapObjects.LabelRenderer
Global MyLabelRender2 As New MapObjects.LabelRenderer

Global CurrentSpatialQueryField As String
Global CurrentSpatialQuery As String

Global fnt As New StdFont

Global Encoder As String
Global Password As String
Global AccessLevel As Integer
Global CurrentUserAccount As String
Global CurrentUserPassword As String

Global CurrentStationToIdentify As String
Global CurrentStationToView As String

Global ProvinceEditMode As Boolean
Global MunicipalityEditMode As Boolean
Global BrgyEditMode As Boolean

Global PageCount As Integer
Global PageCounter As Integer
Global RstQuery As New ADODB.Recordset

Global requestingParty As String
Global Purpose As String
Global O_R As String
Global signatory As String
Global Designation As String
Global Current_User As String
Global Access_Level As Integer

Global HorizontalFixingMethod() As Integer
Global VerticalFixingMethod() As Integer
Global VerticalDatum() As Integer
Global MarkPurpose() As Integer
Global MarkStatus() As Integer
Global MarkType() As Integer
Global Order() As Integer
Global OrderBM() As Integer
Global OrderGravity() As Integer
Global strcondition As String

Global BMQuery As Integer
Private Const ODBC_ADD_SYS_DSN = 4      'Constant for Adding the DSN


Private Const ODBC_REMOVE_SYS_DSN = 6   'Constant for Removing the DSN


Public Type Marker
    name As String
    latitude As String
    longitude As String
End Type




Global Markers() As Marker

Public Type NewRecord
    stationName As String
    Region As String
    Province As String
    Municipality As String
    Barangay As String
    Island As String
    Order As String
    SurveyedBy As String
End Type

Global PreviousRecord As NewRecord

Public Type Result
     name As String
     Province As String
     municipaliy As String
     Barangay As String
     longitude As Double
     latitude As Double
End Type

Global Result() As Result

Global w As MapObjects.Rectangle
Global gEventList() As MapObjects.GeoEvent
Global gEventOrder() As Integer
Global gEventsTag() As String
Global i, recs, fld
Global cnn As New ADODB.Connection
Global rst As ADODB.Recordset

Global gEventBm() As String


Global DbaseCount As Long
Global Station As String
Global RegionPntr As Integer
Global ProvincePntr As Integer
Global Region() As Integer
Global Province() As Integer
Global Municipality() As Integer

Global rstRecords As New ADODB.Recordset
Global rstBenchmarks As New ADODB.Recordset
Global rstGravity As New ADODB.Recordset
Global rstTriangulation As New ADODB.Recordset


Global rstcmbprovince As ADODB.Recordset
Global rstcmbMunicipality As ADODB.Recordset
Global rstcmbBarangay As ADODB.Recordset

Global rstStation As ADODB.Recordset
Global rstProvinceCode As ADODB.Recordset
Global rstProvince As ADODB.Recordset
Global rstProvinceInfo As ADODB.Recordset
Global rstFixMethod As ADODB.Recordset
Global rstVFixMethod As ADODB.Recordset
Global rstVDatum As ADODB.Recordset
Global rstMarkPurpose As ADODB.Recordset
Global rstMarkType As ADODB.Recordset
Global rstMarkStatus As ADODB.Recordset
Global ProvinceCode As String

Global cntr As Integer
Global x As Integer
Global retval
Global PTMZone As Double
Global North As Double
Global East As Double
Global Zone As String
Public Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C

Global EditMode As Boolean
Global EditModegravity As Boolean
Global AddMode As Boolean
Global AddModeBM As Boolean
Global Blank As Boolean
Global rstUsers As New ADODB.Recordset
Global rstUserinfo As New ADODB.Recordset
Global mrst As New MapObjects.Recordset


'**************************************
'Windows API/Global Declarations
'     + Transparent Forms +
'**************************************
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function MakeRgn Lib "region.dll" (ByVal filename As String, ByVal R As Integer, ByVal G As Integer, ByVal b As Integer) As Long
Public Declare Function DeleteRgn Lib "region.dll" (ByVal Region As Long)

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)





Public Type SIZE
    cx As Long
    cy As Long
End Type


Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
    End Type
    Public Const GWL_STYLE = (-16)
    Public Const GWL_EXSTYLE = (-20)
    Public Const WS_EX_LAYERED = &H80000
    Public Const ULW_COLORKEY = &H1
    Public Const ULW_ALPHA = &H2
    Public Const ULW_OPAQUE = &H4
    Public Const AC_SRC_OVER = &H0
    Public Const AC_SRC_ALPHA = &H1
    Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
    Public Const AC_SRC_NO_ALPHA = &H2
    Public Const AC_DST_NO_PREMULT_ALPHA = &H10
    Public Const AC_DST_NO_ALPHA = &H20
    Public Const LWA_COLORKEY = &H1
    Public Const LWA_ALPHA = &H2
    
    
    Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
    End Type
    
    Global OSx As OSVERSIONINFO


Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


Sub Main()
      
   ' FrmLogin.Show 1
      
      
    On Error GoTo Hell
    cnn.Open "Provider=SQLOLEDB.1;User ID=gcpds;password=gcpds;Initial Catalog=gcpds;Data Source =" & GetSetting(App.EXEName, "Server", "ServerName")
    FrmLogin.Show 1
    Exit Sub
    
Hell:
    FrmServer.Show 1
End Sub

Public Function GetCentralMeridian(Region As String, Province As String, Municipality As String) As Double
    Dim rst As New ADODB.Recordset
        rst.CursorLocation = adUseClient
    
    rst.Open "Select Zone from PTMZone Where Region='" & Region & "' and Province='" & Province & "' and Municipality='" & Replace(Municipality, "'", "''") & "'", cnn, adOpenStatic
    
    If rst.RecordCount > 0 Then
        Zone = rst!Zone 'modified dec72009
            If rst!Zone = "1" Then      'modified dec72009
               GetCentralMeridian = 117
            ElseIf rst!Zone = "2" Then
               GetCentralMeridian = 119
            ElseIf rst!Zone = "3" Then
               GetCentralMeridian = 121
            ElseIf rst!Zone = "4" Then
               GetCentralMeridian = 123
            ElseIf rst!Zone = "5" Then
               GetCentralMeridian = 125
            ElseIf rst!Zone = "1A" Then
               GetCentralMeridian = 118.5
            Else
               GetCentralMeridian = 0
            End If
       Else
            GetCentralMeridian = 0
    End If
    
    
    rst.Close
    Set rst = Nothing
End Function


'PTM

Public Sub Compute(Latitude_Degrees As Double, Latitude_Minutes As Double, Latitude_Seconds As Double, Longitude_Degrees As Double, Longitude_Minutes As Double, Longitude_Seconds As Double, PZone As Double)

Dim a, f, b, E, ESEC As Double
Dim latitude, longitude As Double
Const PI = 3.1415926535898
Dim Z2 As Double

Dim s, c, t As Double

Dim DEN As Double
Dim RM As Double
Dim RPV As Double
Dim RAT As Double

Dim ZA As Double
Dim ZB As Double
Dim Z As Double
Dim IZone As Double

Dim DIFF As Double
Dim DELL As Double

Dim t2 As Double
Dim WD As Double
Dim EA, EB, EC, Easting As Double

Dim MD, MD1, MD2, MD3, MD4 As Double
Dim XNA, XNB, XNC, Northing As Double



'Spheroid

a = 6378206.4
f = 294.9786982
'f = 294.9787
b = a * (f - 1) / f
f = 1 / f
E = 1 - (b / a) ^ 2
ESEC = E / (1 - E)


'Latitude and Longitude in Radians

latitude = (Latitude_Degrees + (Latitude_Minutes / 60) + (Latitude_Seconds / 3600)) * (PI / 180)
longitude = (Longitude_Degrees + (Longitude_Minutes / 60) + (Longitude_Seconds / 3600)) * (PI / 180)
Z2 = 0

'Compute parameters

s = Sin(latitude)
c = Cos(latitude)
t = Tan(latitude)


'Compute Radii of Curvature (RM,RPV)

DEN = Sqr(1 - E * s * s)
RM = (a * (1 - E)) / DEN ^ 3
RPV = a / DEN
RAT = RPV / RM


''Compute Zone and Overlap      ''modified dec72009
'
'ZA = Pi * 2 / 180
'ZB = (Longitude / ZA) - 57
''---commented by fat 09042009
''Z = Int(ZB)
''IZone = Z
''----
''---edited by fat 09042009
'Z = PZone
'IZone = Z
''----
''Z = Zone Number
'
'DIFF = ZB - Z
'    '---commented by fat 09042009
''    If DIFF < 0.0833 Then
''       Z2 = Z - 1
''    ElseIf DIFF > 0.9167 Then
''       Z2 = Z + 1
''    End If
    '---
    
'Compute Diff in Longitude between Point and CM

'DELL = Longitude - ((Z + 57) * ZA + ZA / 2)  'modified dec72009
DELL = longitude - (PZone * (PI / 180))

'Compute DIFF in Longitude Between

t2 = t * t
WD = DELL * DELL * c * c
EA = (((179 - t2) * t2 - 479) * t2 + 61) * WD / 42
EB = ((((1 - 6 * t2) * 4 * RAT + (1 + 8 * t2)) * RAT - 2 * t2) * RAT + t2 * t2 + EA) * WD / 20
EC = (RAT - t2 + EB) * WD / 6
Easting = 0.99995 * RPV * DELL * c * (1 + EC) + 500000


'Initialize Counter

MD1 = (((-5) * E / 256 - 3 / 64) * E - 0.25) * E + 1
MD2 = 3 / 8 * ((15 * E / 128 + 0.25) * E + 1) * E
MD3 = 15 / 256 * E * E * (1 + 0.75 * E)
MD4 = 35 * (E ^ 3) / 3072
MD = a * (MD1 * latitude - MD2 * Sin(2 * latitude) + MD3 * Sin(4 * latitude) - MD4 * Sin(6 * latitude))

'Compute Northing

XNA = (((543 - t2) * t2 - 3111) * t2 + 1385) * WD / 56
XNB = (((((11 - 24 * t2) * 8 * RAT - 28 * (1 - 6 * t2)) * RAT + (1 - 32 * t2)) * RAT - 2 * t2) * RAT + t2 + XNA) * WD / 30
XNC = (4 * RAT * RAT + RAT - t2 + XNB) * WD / 12
Northing = 0.99995 * (MD + RPV * s * DELL * DELL * c / 2 * (1 + XNC))

    If latitude < 0 Then
       Northing = Northing + 10000000
    End If

East = Round(Easting, 3)
North = Round(Northing, 3)
'Zone = Z  'modified dec72009
End Sub

Public Function TranslucentForm(frm As Form, TranslucenceLevel As Byte) As Boolean
    SetWindowLong frm.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hWnd, 0, TranslucenceLevel, LWA_ALPHA
    TranslucentForm = Err.LastDllError = 0
End Function
Public Sub Make_Transparent(hWnd As Long, Rate As Byte)
'Exported Codes "This codes is not mine."

    Dim WinInfo As Long
    WinInfo = GetWindowLong(hWnd, GWL_EXSTYLE)
    WinInfo = WinInfo Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, WinInfo
    SetLayeredWindowAttributes hWnd, 0, Rate, LWA_ALPHA
End Sub


Public Sub FillUp()

If rstRecords.RecordCount = 0 Then
    BlankForm
    Exit Sub
End If

     
    Dim rstdetails As New ADODB.Recordset
    
    
    rstdetails.Open "SELECT geoprov.*, order_lib.description as HOR_Order, horfixme.HDesc, marktype.MTDesc, markpur.MDesc, markstatus.MSDesc FROM ((((geoprov LEFT JOIN order_lib ON geoprov.h_order = order_lib.h_order) LEFT JOIN marktype ON geoprov.mark_type = marktype.MTCode) LEFT JOIN markstatus ON geoprov.mark_stat = markstatus.MSCode) LEFT JOIN horfixme ON geoprov.h_fix = horfixme.HCode) LEFT JOIN markpur ON geoprov.mark_const = markpur.MCode where geoprov.stat_name='" & Replace(rstRecords("stat_name"), "'", "''") & "' Order By Province", cnn, adOpenStatic, adLockOptimistic
   
    
    FrmGCPDS.TxtRegion = IIf(IsNull(rstdetails("Region")), "", rstdetails("Region"))
    FrmGCPDS.TxtProvince = IIf(IsNull(rstdetails("Province")), "", rstdetails("Province"))
    FrmGCPDS.TxtMunicipality = IIf(IsNull(rstdetails("Municipal")), "", rstdetails("Municipal"))
    FrmGCPDS.TxtBarangay = IIf(IsNull(rstdetails("Barangay")), "", rstdetails("Barangay"))
    FrmGCPDS.TxtName = IIf(IsNull(rstdetails("Stat_name")), "", rstdetails("Stat_Name"))
   
    FrmGCPDS.TxtIsland = IIf(IsNull(rstdetails("Island")), "", rstdetails("Island"))
    
    FrmGCPDS.TxtLatitude = IIf(IsNull(rstdetails("D_Lat")), "", rstdetails("D_Lat") & "º ") & IIf(IsNull(rstdetails("M_Lat")), "", rstdetails("M_Lat") & "' ") & IIf(IsNull(rstdetails("S_Lat")), "", rstdetails("S_Lat") & """ ")
    FrmGCPDS.TxtLongitude = IIf(IsNull(rstdetails("D_Long")), "", rstdetails("D_Long") & "º ") & IIf(IsNull(rstdetails("M_Long")), "", rstdetails("M_Long") & "' ") & IIf(IsNull(rstdetails("S_Long")), "", rstdetails("S_Long") & """ ")
    
    
           
    
    FrmGCPDS.TxtDLat = IIf(IsNull(rstdetails("wgs84ED")), "", rstdetails("wgs84ED"))
    FrmGCPDS.TxtMLat = IIf(IsNull(rstdetails("wgs84EM")), "", rstdetails("wgs84EM"))
    FrmGCPDS.TxtSLat = IIf(IsNull(rstdetails("wgs84ES")), "", rstdetails("wgs84ES"))
    FrmGCPDS.TxtDLong = IIf(IsNull(rstdetails("wgs84ND")), "", rstdetails("wgs84ND"))
    FrmGCPDS.TxtMLong = IIf(IsNull(rstdetails("wgs84NM")), "", rstdetails("wgs84NM"))
    FrmGCPDS.TxtSLong = IIf(IsNull(rstdetails("wgs84NS")), "", rstdetails("wgs84NS"))
    
    If IsNull(rstdetails("wgs84ED")) And IsNull(rstdetails("wgs84EM")) And IsNull(rstdetails("wgs84ES")) And IsNull(rstdetails("wgs84ND")) And IsNull(rstdetails("wgs84NM")) And IsNull(rstdetails("wgs84NS")) Then
        FrmGCPDS.Label(38).Visible = False
        FrmGCPDS.Label(37).Visible = False
        FrmGCPDS.Label(35).Visible = False
        FrmGCPDS.Label(28).Visible = False
        FrmGCPDS.Label(33).Visible = False
        FrmGCPDS.Label(34).Visible = False
        Else
        FrmGCPDS.Label(38).Visible = True
        FrmGCPDS.Label(37).Visible = True
        FrmGCPDS.Label(35).Visible = True
        FrmGCPDS.Label(28).Visible = True
        FrmGCPDS.Label(33).Visible = True
        FrmGCPDS.Label(34).Visible = True
    End If
    
    If (IsNull(rstdetails("D_lat")) And IsNull(rstdetails("m_lat")) And IsNull(rstdetails("s_lat")) And IsNull(rstdetails("D_long")) And IsNull(rstdetails("m_long")) And IsNull(rstdetails("s_long"))) Or _
       (rstdetails("D_lat") = 0 And rstdetails("m_lat") = 0 And rstdetails("s_lat") = 0 And rstdetails("D_long") = 0 And rstdetails("M_long") = 0 And rstdetails("s_long") = 0) Then
        
        FrmGCPDS.Label(20).Visible = False
        FrmGCPDS.Label(22).Visible = False
        FrmGCPDS.Label(23).Visible = False
        FrmGCPDS.Label(27).Visible = False
        FrmGCPDS.Label(25).Visible = False
        FrmGCPDS.Label(24).Visible = False
            FrmGCPDS.TxtLatD = ""
            FrmGCPDS.TxtLatM = ""
            FrmGCPDS.TxtLatS = ""
            FrmGCPDS.TxtLongD = ""
            FrmGCPDS.TxtLongM = ""
            FrmGCPDS.TxtLongS = ""
        
        Else
        FrmGCPDS.Label(20).Visible = True
        FrmGCPDS.Label(22).Visible = True
        FrmGCPDS.Label(23).Visible = True
        FrmGCPDS.Label(27).Visible = True
        FrmGCPDS.Label(25).Visible = True
        FrmGCPDS.Label(24).Visible = True
        
            FrmGCPDS.TxtLatD = IIf(IsNull(rstdetails("D_Lat")), "", rstdetails("D_Lat"))
            FrmGCPDS.TxtLatM = IIf(IsNull(rstdetails("M_Lat")), "", rstdetails("M_Lat"))
            FrmGCPDS.TxtLatS = IIf(IsNull(rstdetails("S_Lat")), "", rstdetails("S_Lat"))
            FrmGCPDS.TxtLongD = IIf(IsNull(rstdetails("D_Long")), "", rstdetails("D_Long"))
            FrmGCPDS.TxtLongM = IIf(IsNull(rstdetails("M_Long")), "", rstdetails("M_Long"))
            FrmGCPDS.TxtLongS = IIf(IsNull(rstdetails("S_Long")), "", rstdetails("S_Long"))
    End If
    
    
    
    
    FrmGCPDS.TxtEllipsoidalH2 = IIf(IsNull(rstdetails("ellipz")), "", rstdetails("ellipz"))
   
    
'    If IsNumeric(rstdetails("D_Lat")) = True And IsNumeric(rstdetails("M_Lat")) = True And IsNumeric(rstdetails("S_Lat")) = True And IsNumeric(rstdetails("D_Long")) = True And IsNumeric(rstdetails("M_Long")) = True And IsNumeric(rstdetails("S_Long")) = True And rstdetails("H_ref") = "PRS92" Then
'      Call Compute(rstdetails("D_Lat"), rstdetails("M_Lat"), rstdetails("S_Lat"), rstdetails("D_Long"), rstdetails("M_Long"), rstdetails("S_Long"))
'        FrmGCPDS.TxtNorthing = Format(North, "#,##0.####")
'        FrmGCPDS.TxtEasting = Format(East, "#,##0.####")
'        FrmGCPDS.TxtZone = Zone
'        Else
'        FrmGCPDS.TxtNorthing = ""
'        FrmGCPDS.TxtEasting = ""
'        FrmGCPDS.TxtZone = ""
'    End If
   FrmGCPDS.TxtNorthing = IIf(IsNull(rstdetails("northing")), "", Format(rstdetails("northing"), "#,##0.####"))
   FrmGCPDS.TxtEasting = IIf(IsNull(rstdetails("easting")), "", Format(rstdetails("easting"), "#,##0.####"))
   FrmGCPDS.TxtZone = IIf(IsNull(rstdetails("zone")), "", rstdetails("zone"))
    
   FrmGCPDS.TextBoxUTMNorthing.Text = IIf(IsNull(rstdetails("utmy")), "", Format(rstdetails("utmy"), "#,##0.###"))
   FrmGCPDS.TextBoxUTMEasting.Text = IIf(IsNull(rstdetails("utmx")), "", Format(rstdetails("utmx"), "#,##0.###"))
   FrmGCPDS.TextBoxUTMZone = IIf(IsNull(rstdetails("utmz")), "", rstdetails("utmz"))
    
    FrmGCPDS.TxtFixingMethod = IIf(IsNull(rstdetails("HDesc")), "", StrConv(rstdetails("HDesc"), vbUpperCase))
    
    
    If IsNull(rstdetails("HOR_Order")) Or Trim(rstdetails("HOR_Order")) = "" Then
        'FrmGCPDS.txtOrder.ListIndex = -1   'modified by fat 10/13/2009
        FrmGCPDS.txtOrder = ""
        Else
        FrmGCPDS.txtOrder = StrConv(rstdetails("HOR_ORDER"), vbUpperCase)
    End If
    
    FrmGCPDS.TxtRef = IIf(IsNull(rstdetails("H_Ref")), "", StrConv(rstdetails("H_Ref"), vbUpperCase))
    
    
    'FrmGCPDS.TxtDatum = IIf(IsNull(rstdetails("H_Datum")), "", IIf(rstdetails("H_datum") = 1, "Luzon", "0"))
    FrmGCPDS.txtEllipsoidalH = IIf(IsNull(rstdetails("Ell_Hgt")), "", rstdetails("Ell_Hgt"))
    FrmGCPDS.TxtEstablishedBy = IIf(IsNull(rstdetails("Hor_Authty")), "", rstdetails("Hor_Authty"))
    FrmGCPDS.TxtDateEntry = IIf(IsNull(rstdetails("H_Date_Ety")), "", Format(rstdetails("H_Date_Ety"), "Mm-dd-yyyy"))
    FrmGCPDS.TxtDateComputed = IIf(IsNull(rstdetails("H_Date_Com")), "", Format(rstdetails("H_Date_Com"), "Mm-dd-yyyy"))
    
    FrmGCPDS.txtMarkPurpose = IIf(IsNull(rstdetails("Mdesc")), "", StrConv(rstdetails("MDesc"), vbUpperCase))
    FrmGCPDS.txtMarkType = IIf(IsNull(rstdetails("MTdesc")), "", StrConv(rstdetails("MTDesc"), vbUpperCase))
    FrmGCPDS.txtMarkStatus = IIf(IsNull(rstdetails("MSdesc")), "", StrConv(rstdetails("MSDesc"), vbUpperCase))
    
    FrmGCPDS.txtAuthority = IIf(IsNull(rstdetails("Authority")), "", rstdetails("Authority"))
    FrmGCPDS.txtEstablished = IIf(IsNull(rstdetails("Date_Est")), "", Format(rstdetails("Date_Est"), "Mm-dd-yyyy"))
   
   If IsNull(rstdetails!date_est_month) = False Then
      FrmGCPDS.ComboBoxMonth.ListIndex = rstdetails!date_est_month
      Else
      FrmGCPDS.ComboBoxMonth.ListIndex = 0
   End If
   
   If IsNull(rstdetails!date_est_day) = False Then
      FrmGCPDS.ComboBoxDay.ListIndex = rstdetails!date_est_day
      Else
      FrmGCPDS.ComboBoxDay.ListIndex = -1
   End If
   
   If IsNull(rstdetails!date_est_year) = False Then
      FrmGCPDS.txtYear = rstdetails!date_est_year
      Else
      FrmGCPDS.txtYear = ""
   End If
   
   FrmGCPDS.TxtMSL = IIf(IsNull(rstdetails!AdoptedBy), "", rstdetails!AdoptedBy)
   
    FrmGCPDS.txtLastRecover = IIf(IsNull(rstdetails("Date_Las_R")), "", Format(rstdetails("Date_Las_R"), "Mm-dd-yyyy"))
   
    
    FrmGCPDS.FormCaption.Caption = Format(rstRecords.bookmark, "#,##0") & " of " & Format(rstRecords.RecordCount, "#,##0")
    FrmGCPDS.LabelEncoder.Caption = IIf(IsNull(rstdetails!Encoder), "No Data", rstdetails!Encoder)
    FrmGCPDS.LabelDateUpdated.Caption = IIf(IsNull(rstdetails!dateUpdated), "No Data", rstdetails!dateUpdated)
    
End Sub

Public Sub FillUpBenchmarks()

If rstBenchmarks.RecordCount = 0 Then
    BlankForm
    Exit Sub
End If

     
    Dim rstdetails As New ADODB.Recordset
    rstdetails.Open "SELECT benchmarks.*, bm_order_lib.description as BMOrder, marktype.MTDesc as BMMarkType, markstatus.MSDesc as BMMarkStatus, markpur.MDesc as BMMarkPurpose, vertdat.VDDesc as BMVerticalDatum, vertfix.vfdesc as BMVerticalFixing " & _
                    "FROM (((((benchmarks LEFT JOIN bm_order_lib ON benchmarks.e_order = bm_order_lib.e_order) LEFT JOIN marktype ON benchmarks.mark_type = marktype.MTCode) LEFT JOIN markpur ON benchmarks.mark_const = markpur.MCode) LEFT JOIN markstatus ON benchmarks.mark_stat = markstatus.MSCode) LEFT JOIN vertdat ON benchmarks.e_datum = vertdat.VDCode) LEFT JOIN vertfix ON benchmarks.e_fix = vertfix.vfcode where benchmarks.ucode=" & rstBenchmarks("ucode"), cnn, adOpenStatic


   
    FrmGCPDS.TxtEName = IIf(IsNull(rstdetails("Stat_name")), "", rstdetails("Stat_name"))
    FrmGCPDS.TxtERegion = IIf(IsNull(rstdetails("Region")), "", rstdetails("Region"))
    FrmGCPDS.TxtEProvince = IIf(IsNull(rstdetails("Province")), "", rstdetails("Province"))
    FrmGCPDS.TxtEMunicipality = IIf(IsNull(rstdetails("Municipal")), "", rstdetails("Municipal"))
    FrmGCPDS.TxtEBarangay = IIf(IsNull(rstdetails("Barangay")), "", rstdetails("Barangay"))
    FrmGCPDS.TxtEIsland = IIf(IsNull(rstdetails("Island")), "", rstdetails("Island"))
    FrmGCPDS.TxtEDatum = IIf(IsNull(rstdetails("BMVerticalDatum")), "", StrConv(rstdetails("BMVerticalDatum"), vbProperCase))
    FrmGCPDS.TxtElevation = IIf(IsNull(rstdetails("Elevation")), "", rstdetails("Elevation"))
    FrmGCPDS.TxtBMPlus = IIf(IsNull(rstdetails("BMPlus")), "", rstdetails("BMPlus"))
    
    If (IsNull(rstdetails("Latitude") = True)) Then
        FrmGCPDS.TextBoxLatitude = ""
    Else
        FrmGCPDS.TextBoxLatitude = DDtoDMS(rstdetails("Latitude"))
    End If
    
    If (IsNull(rstdetails("Longitude") = True)) Then
        FrmGCPDS.TextBoxLongitude = ""
    Else
        FrmGCPDS.TextBoxLongitude = DDtoDMS(rstdetails("Longitude"))
    End If
    
    
    FrmGCPDS.TxtBMOrder = IIf(IsNull(rstdetails("BMOrder")), "", rstdetails("BMOrder"))
    FrmGCPDS.TxtEFix = IIf(IsNull(rstdetails("BMVerticalFixing")), "", rstdetails("BMVerticalFixing"))
    
    FrmGCPDS.TxtElevationAuthority = IIf(IsNull(rstdetails("Elv_Authty")), "", rstdetails("Elv_Authty"))
    FrmGCPDS.TxtBMDateOfEntry = IIf(IsNull(rstdetails("E_Date_Ety")), "", Format(rstdetails("E_Date_Ety"), "Mm-dd-yyyy"))
    FrmGCPDS.TxtBMDateComputed = IIf(IsNull(rstdetails("E_Date_Com")), "", Format(rstdetails("E_Date_Com"), "Mm-dd-yyyy"))
    FrmGCPDS.TxtBMDateEstablished = IIf(IsNull(rstdetails("Date_est")), "", Format(rstdetails("Date_est"), "Mm-dd-yyyy"))
'
    FrmGCPDS.BMMarkPurpose = IIf(IsNull(rstdetails("BMMarkPurpose")), "", StrConv(rstdetails("BMMarkPurpose"), vbUpperCase))
    FrmGCPDS.BMMarkType = IIf(IsNull(rstdetails("BMMarkType")), "", StrConv(rstdetails("BMMarkType"), vbUpperCase))
    FrmGCPDS.BMMarkStatus = IIf(IsNull(rstdetails("BMMarkStatus")), "", StrConv(rstdetails("BMMarkStatus"), vbUpperCase))
    FrmGCPDS.TxtBMAuthority = IIf(IsNull(rstdetails("Authority")), "", rstdetails("Authority"))
'    FrmGCPDS.txtEstablished = IIf(IsNull(rstdetails("Date_Est")), "", Format(rstdetails("Date_Est"), "Mm-dd-yyyy"))
'    'FrmGCPDS.txtProjectNo = IIf(IsNull(rstdetails("Project_No")), "", rstdetails("Project_No"))
    FrmGCPDS.TxtBMDateLastRecovered = IIf(IsNull(rstdetails("Date_Las_R")), "", Format(rstdetails("Date_Las_R"), "Mm-dd-yyyy"))
'    'FrmGCPDS.txtFileNo = IIf(IsNull(rstdetails("File_ID")), "", rstdetails("File_ID"))
'
    FrmGCPDS.LabelDateUpdated2.Caption = IIf(IsNull(rstdetails("dateupdated")), "", rstdetails("dateupdated"))
    FrmGCPDS.LabelEncoder2.Caption = IIf(IsNull(rstdetails("encoder")), "", rstdetails("encoder"))
    
    FrmGCPDS.FormCaption2.Caption = Format(rstBenchmarks.bookmark, "#,##0") & " of " & Format(rstBenchmarks.RecordCount, "#,##0") & " Records"
    FrmGCPDS.TxtBMDescription = IIf(IsNull(rstdetails("Description")), "", Replace(rstdetails("Description"), "ì", ""))
End Sub


Public Sub FillUpGravity()

If rstGravity.RecordCount = 0 Then
    BlankFormGravity
    Exit Sub
End If

     
    Dim rstdetails As New ADODB.Recordset
    rstdetails.Open "Select gravity.*,bm_order_lib.description  as [G_ORDER] from gravity left join bm_order_lib on gravity.h_order=bm_order_lib.e_order where stat_name='" & Replace(rstGravity!Stat_Name, "'", "''") & "'", cnn, adOpenStatic


   
    FrmGCPDS.TextBoxGravityName.Text = IIf(IsNull(rstdetails("stat_name")), "", rstdetails("stat_name"))
    FrmGCPDS.TextBoxGravityRegion.Text = IIf(IsNull(rstdetails("region")), "", rstdetails("region"))
    FrmGCPDS.TextBoxGravityProvince.Text = IIf(IsNull(rstdetails("province")), "", rstdetails("province"))
    FrmGCPDS.TextBoxGravityMunicipality.Text = IIf(IsNull(rstdetails("municipal")), "", rstdetails("municipal"))
    FrmGCPDS.TextBoxGravityBarangay.Text = IIf(IsNull(rstdetails("barangay")), "", rstdetails("barangay"))
    FrmGCPDS.ComboBoxOrderGravity = IIf(IsNull(rstdetails("G_ORDER")), "", rstdetails("G_ORDER"))
    
    If IsNull(rstdetails("latitude")) = False Then
        FrmGCPDS.TextBoxGravityLatitude.Text = DDtoDMS(rstdetails("latitude"))
        Else
        FrmGCPDS.TextBoxGravityLatitude.Text = ""
    End If
    
    If IsNull(rstdetails("longitude")) = False Then
        FrmGCPDS.TextBoxGravityLongitude.Text = DDtoDMS(rstdetails("longitude"))
        Else
        FrmGCPDS.TextBoxGravityLongitude.Text = ""
    End If
    
    FrmGCPDS.TextBoxObservedValues.Text = IIf(IsNull(rstdetails("observedValues")), "", rstdetails("observedValues"))
   
    FrmGCPDS.TextBoxGravityElevation.Text = IIf(IsNull(rstdetails("elevation")), "", rstdetails("elevation") & " " & rstdetails("elevationunit"))
    
  
    FrmGCPDS.TextBoxGravityDescription.Text = IIf(IsNull(rstdetails("description")), "", rstdetails("description"))
    FrmGCPDS.LabelencoderGravity.Caption = IIf(IsNull(rstdetails("encoder")), "", rstdetails("encoder"))
    FrmGCPDS.LabelUpdatedGravity.Caption = IIf(IsNull(rstdetails("dateLastUpdated")), "", rstdetails("dateLastUpdated"))
    FrmGCPDS.LabelGravityRecordStatus.Caption = Format(rstGravity.bookmark, "#,##0") & " of " & Format(rstGravity.RecordCount, "#,##0")
    
End Sub


Public Sub FillUpTriangulation()

If rstTriangulation.RecordCount = 0 Then
    'BlankFormTriangulation
    Exit Sub
End If

     
    Dim rstdetails As New ADODB.Recordset
    rstdetails.Open "Select * FROM TRIANGULATION where stat_name='" & Replace(rstTriangulation!Stat_Name, "'", "''") & "'", cnn, adOpenStatic


   
    FrmGCPDS.TextBoxOldStationName.Text = IIf(IsNull(rstdetails("stat_name")), "", rstdetails("stat_name"))
    FrmGCPDS.TextBoxOldRegion.Text = IIf(IsNull(rstdetails("region")), "", rstdetails("region"))
    FrmGCPDS.TextBoxOldProvince.Text = IIf(IsNull(rstdetails("province")), "", rstdetails("province"))
    FrmGCPDS.TextBoxOldMunicipality.Text = IIf(IsNull(rstdetails("municipal")), "", rstdetails("municipal"))
    FrmGCPDS.TextBoxOldBarangay.Text = IIf(IsNull(rstdetails("barangay")), "", rstdetails("barangay"))
    FrmGCPDS.TextBoxOldOrder = IIf(IsNull(rstdetails("h_order")), "", rstdetails("h_order"))
    FrmGCPDS.TextBoxOldLatitude.Text = rstdetails("d_lat") & "º " & rstdetails("m_lat") & "' " & rstdetails("s_lat") & "''"
    FrmGCPDS.TextBoxOldLongitude.Text = rstdetails("d_long") & "º " & rstdetails("m_long") & "' " & rstdetails("s_long") & "''"
    FrmGCPDS.TextBoxOldDescription.Text = IIf(IsNull(rstdetails("description")), "", rstdetails("description"))
    
    
    FrmGCPDS.TextBoxOldDateEntry = IIf(IsNull(rstdetails("h_date_ety")), "", rstdetails("h_date_ety"))
    FrmGCPDS.TextBoxOldDateEstablished = IIf(IsNull(rstdetails("date_est")), "", rstdetails("date_est"))
    FrmGCPDS.TextBoxOldBookmark.Text = rstTriangulation.bookmark & " of " & rstTriangulation.RecordCount
'    If IsNull(rstdetails("latitude")) = False Then
'        FrmGCPDS.TextBoxGravityLatitude.Text = DDtoDMS(rstdetails("latitude"))
'        Else
'        FrmGCPDS.TextBoxGravityLatitude.Text = ""
'    End If
'
'    If IsNull(rstdetails("longitude")) = False Then
'        FrmGCPDS.TextBoxGravityLongitude.Text = DDtoDMS(rstdetails("longitude"))
'        Else
'        FrmGCPDS.TextBoxGravityLongitude.Text = ""
'    End If
'
'    FrmGCPDS.TextBoxObservedValues.Text = IIf(IsNull(rstdetails("observedValues")), "", rstdetails("observedValues"))
'
'    FrmGCPDS.TextBoxGravityElevation.Text = IIf(IsNull(rstdetails("elevation")), "", rstdetails("elevation") & " " & rstdetails("elevationunit"))
'
'

'    FrmGCPDS.LabelencoderGravity.Caption = IIf(IsNull(rstdetails("encoder")), "", rstdetails("encoder"))
'    FrmGCPDS.LabelUpdatedGravity.Caption = IIf(IsNull(rstdetails("dateLastUpdated")), "", rstdetails("dateLastUpdated"))
'    FrmGCPDS.LabelGravityRecordStatus.Caption = Format(rstGravity.bookmark, "#,##0") & " of " & Format(rstGravity.RecordCount, "#,##0")
    
End Sub

'Encryption

Public Function Encrypt(str As String) As String
Dim i As Integer
Dim ascval As Integer
    
    For i = 1 To Len(str)
    
        ascval = Asc(Mid(str, i, 1)) + Len(str) + 60
    
        If ascval > 255 Then
            ascval = ascval - 255
        End If
       
        Mid(str, i, 1) = Chr(ascval)
    Next
    
    Encrypt = str

End Function


'Decryption Routine

Public Function Decrypt(str As String) As String
Dim i As Integer
Dim ascval As Integer


    For i = 1 To Len(str)
        ascval = Asc(Mid(str, i, 1)) - Len(str) - 60
    
        If ascval < 1 Then
            ascval = ascval + 255
        End If
       
        Mid(str, i, 1) = Chr(ascval)
    Next
    
    Decrypt = str

End Function



Public Sub LoadProvinces()
    
        Dim i As Integer
        
            Set rstcmbprovince = New ADODB.Recordset
                rstcmbprovince.Open "Select prov_name,region,province from Provmast order by prov_name", cnn, adOpenStatic, adLockOptimistic
            
        ReDim Region(1 To rstcmbprovince.RecordCount)
        ReDim Province(1 To rstcmbprovince.RecordCount)
        
        For i = 1 To rstcmbprovince.RecordCount
            FrmGCPDS.TxtProvince.AddItem StrConv(rstcmbprovince.Fields(0).Value, vbProperCase)
            
            Region(i) = rstcmbprovince.Fields(1).Value
            Province(i) = rstcmbprovince.Fields(2).Value
            rstcmbprovince.MoveNext
        Next
            
End Sub



Public Function LoadMunicipality(Region As Integer, Province As Integer)
    
        Dim i As Integer
        
            Set rstcmbMunicipality = New ADODB.Recordset
                rstcmbMunicipality.Open "Select city_name,Mun_code from cmast where region=" & Region & " and province=" & Province, cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.TxtMunicipality.Clear
        
            ReDim Municipality(1 To rstcmbMunicipality.RecordCount)
        For i = 1 To rstcmbMunicipality.RecordCount
            FrmGCPDS.TxtMunicipality.AddItem StrConv(rstcmbMunicipality.Fields(0).Value, vbProperCase)
            Municipality(i) = rstcmbMunicipality.Fields(1).Value
            rstcmbMunicipality.MoveNext
        Next
        
        
            
End Function


Public Function LoadBarangay(Region As Integer, Province As Integer, Municipality As Integer)
    
        Dim i As Integer
        
            Set rstcmbBarangay = New ADODB.Recordset
                rstcmbBarangay.Open "Select bgy_name from rurban where region=" & Region & " and province=" & Province & " and city_mun=" & Municipality, cnn, adOpenStatic, adLockOptimistic
                
                FrmGCPDS.TxtBarangay.Clear
        
        For i = 1 To rstcmbBarangay.RecordCount
            FrmGCPDS.TxtBarangay.AddItem StrConv(rstcmbBarangay.Fields(0).Value, vbProperCase)
            rstcmbBarangay.MoveNext
        Next
        
        
            
End Function


Public Sub LoadHorizontalFixingMethod()
        Dim i As Integer
        Dim rstHorizontalFixingMethod As New ADODB.Recordset
        rstHorizontalFixingMethod.Open "Select Hdesc,Hcode from HorFixMe order by Hdesc", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.TxtFixingMethod.Clear
        
        ReDim HorizontalFixingMethod(0 To rstHorizontalFixingMethod.RecordCount)
        For i = 1 To rstHorizontalFixingMethod.RecordCount
            FrmGCPDS.TxtFixingMethod.AddItem IIf(IsNull(rstHorizontalFixingMethod("HDesc")), "", rstHorizontalFixingMethod("HDesc"))
            HorizontalFixingMethod(i) = rstHorizontalFixingMethod("Hcode").Value
            rstHorizontalFixingMethod.MoveNext
        Next
            
End Sub

Public Sub LoadVerticalFixingMethod()
        Dim i As Integer
        Dim rstVerticalFixingMethod As New ADODB.Recordset
        rstVerticalFixingMethod.Open "Select vfdesc,vfcode from vertfix order by vfdesc", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.TxtEFix.Clear
        
        ReDim VerticalFixingMethod(0 To rstVerticalFixingMethod.RecordCount)
        For i = 1 To rstVerticalFixingMethod.RecordCount
            FrmGCPDS.TxtEFix.AddItem IIf(IsNull(rstVerticalFixingMethod("vfDesc")), "", rstVerticalFixingMethod("vfDesc"))
            VerticalFixingMethod(i) = rstVerticalFixingMethod("vfcode").Value
            rstVerticalFixingMethod.MoveNext
        Next
            
End Sub

Public Sub LoadMarkPurpose()
    
        Dim i As Integer
        Dim rstMarkPurpose As New ADODB.Recordset
        rstMarkPurpose.Open "Select Mdesc,Mcode from MarkPur order by Mdesc", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.txtMarkPurpose.Clear
        ReDim MarkPurpose(0 To rstMarkPurpose.RecordCount)
        For i = 1 To rstMarkPurpose.RecordCount
            FrmGCPDS.txtMarkPurpose.AddItem IIf(IsNull(rstMarkPurpose("MDesc")), "", StrConv(rstMarkPurpose("MDesc"), vbProperCase))
            FrmGCPDS.BMMarkPurpose.AddItem IIf(IsNull(rstMarkPurpose("MDesc")), "", StrConv(rstMarkPurpose("MDesc"), vbProperCase))
            MarkPurpose(i) = rstMarkPurpose("Mcode").Value
            rstMarkPurpose.MoveNext
        Next
            
End Sub

Public Sub LoadOrder()
    
        Dim i As Integer
        Dim rstOrder As New ADODB.Recordset
        
        rstOrder.Open "Select * from Order_Lib order by H_Order", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.txtOrder.Clear
        
        ReDim Order(0 To rstOrder.RecordCount)
        For i = 1 To rstOrder.RecordCount
            FrmGCPDS.txtOrder.AddItem StrConv(rstOrder("description").Value, vbUpperCase)
            Order(i) = rstOrder("H_Order").Value
            rstOrder.MoveNext
        Next
            
End Sub

Public Sub LoadMonths()
    

        FrmGCPDS.ComboBoxMonth.Clear
        FrmGCPDS.ComboBoxMonth.AddItem ""
        FrmGCPDS.ComboBoxMonth.AddItem "Jan"
       FrmGCPDS.ComboBoxMonth.AddItem "Feb"
       FrmGCPDS.ComboBoxMonth.AddItem "Mar"
       FrmGCPDS.ComboBoxMonth.AddItem "Apr"
       FrmGCPDS.ComboBoxMonth.AddItem "May"
       FrmGCPDS.ComboBoxMonth.AddItem "June"
       FrmGCPDS.ComboBoxMonth.AddItem "July"
       FrmGCPDS.ComboBoxMonth.AddItem "Aug"
       FrmGCPDS.ComboBoxMonth.AddItem "Sep"
       FrmGCPDS.ComboBoxMonth.AddItem "Oct"
       FrmGCPDS.ComboBoxMonth.AddItem "Nov"
       FrmGCPDS.ComboBoxMonth.AddItem "Dec"
     

End Sub

Public Sub LoadDays()
Dim i As Integer

       FrmGCPDS.ComboBoxDay.Clear
       FrmGCPDS.ComboBoxDay.AddItem ""
  
  For i = 1 To 31
      FrmGCPDS.ComboBoxDay.AddItem i
  Next

End Sub

Public Sub LoadOrderBM()
    
        Dim i As Integer
        Dim rstOrder As New ADODB.Recordset
        
        rstOrder.Open "Select * from bm_order_lib order by e_order", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.TxtBMOrder.Clear
        
        ReDim OrderBM(0 To rstOrder.RecordCount)
        
        For i = 1 To rstOrder.RecordCount
            FrmGCPDS.TxtBMOrder.AddItem StrConv(rstOrder("description").Value, vbUpperCase)
            OrderBM(i) = rstOrder("e_Order").Value
            rstOrder.MoveNext
        Next
            
End Sub

Public Sub LoadOrderGravity()
    
        Dim i As Integer
        Dim rstOrder As New ADODB.Recordset
        
        rstOrder.Open "Select * from bm_order_lib order by e_order", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.ComboBoxOrderGravity.Clear
        
        ReDim OrderGravity(0 To rstOrder.RecordCount)
        
        For i = 1 To rstOrder.RecordCount
            FrmGCPDS.ComboBoxOrderGravity.AddItem StrConv(rstOrder("description").Value, vbUpperCase)
            OrderGravity(i) = rstOrder("e_Order").Value
            rstOrder.MoveNext
        Next
            
End Sub

Public Sub LoadMarkType()
    
        Dim i As Integer
        Dim rstMarkType As New ADODB.Recordset
        rstMarkType.Open "Select MTdesc,MTCode from MarkType order by MTDesc", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.txtMarkType.Clear
        ReDim MarkType(0 To rstMarkType.RecordCount)
        
        For i = 1 To rstMarkType.RecordCount
            FrmGCPDS.txtMarkType.AddItem IIf(IsNull(rstMarkType("MTDesc")), "", StrConv(rstMarkType("MTDesc"), vbProperCase))
            FrmGCPDS.BMMarkType.AddItem IIf(IsNull(rstMarkType("MTDesc")), "", StrConv(rstMarkType("MTDesc"), vbProperCase))
            MarkType(i) = rstMarkType("MTcode").Value
            rstMarkType.MoveNext
        Next
            
End Sub

Public Sub LoadMarkStatus()
    
        Dim i As Integer
        Dim rstMarkStatus As New ADODB.Recordset
        rstMarkStatus.Open "Select MSdesc,MSCode from MarkStatus order by MSDesc", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.txtMarkStatus.Clear
        ReDim MarkStatus(0 To rstMarkStatus.RecordCount)
        For i = 1 To rstMarkStatus.RecordCount
            FrmGCPDS.txtMarkStatus.AddItem IIf(IsNull(rstMarkStatus("MSDesc")), "", StrConv(rstMarkStatus("MSDesc"), vbProperCase))
            FrmGCPDS.BMMarkStatus.AddItem IIf(IsNull(rstMarkStatus("MSDesc")), "", StrConv(rstMarkStatus("MSDesc"), vbProperCase))
            MarkStatus(i) = rstMarkStatus("MSCode").Value
            rstMarkStatus.MoveNext
        Next
            
End Sub

Public Sub LoadVerticalDatum()
    
        Dim i As Integer
        Dim rst As New ADODB.Recordset
        rst.Open "Select vddesc,vdCode from vertdat", cnn, adOpenStatic, adLockOptimistic
        FrmGCPDS.TxtEDatum.Clear
        ReDim VerticalDatum(0 To rst.RecordCount)
        For i = 1 To rst.RecordCount
            FrmGCPDS.TxtEDatum.AddItem IIf(IsNull(rst("vdDesc")), "", StrConv(rst("vdDesc"), vbProperCase))
            VerticalDatum(i) = rst("vdCode").Value
            rst.MoveNext
        Next
            
End Sub

Public Sub BlankForm()

With FrmGCPDS
   
    .TxtName.Text = ""
    .TxtIsland.Text = ""
    .TxtRegion.Text = ""
    .TxtProvince.Text = ""
    .TxtMunicipality.Text = ""
    .TxtBarangay.Text = ""
    .TxtLatitude.Text = ""
    .TxtLongitude.Text = ""
    .txtOrder.ListIndex = -1

    .TxtEasting.Text = ""
    .TxtNorthing.Text = ""
    .TxtZone.Text = ""
    '.TxtDatum.Text = ""
    .TextBoxUTMNorthing.Text = ""
    .TextBoxUTMEasting.Text = ""
    .TextBoxUTMZone.Text = ""
    
    .TxtEstablishedBy.Text = ""
    .ComboBoxMonth.ListIndex = -1
    .ComboBoxDay.ListIndex = -1
    .txtYear.Text = ""

    .TxtRef = "PRS92"
    .TxtFixingMethod.ListIndex = -1
    .TxtDateEntry.Value = Empty
    .TxtDateComputed = Empty


    .txtEllipsoidalH.Text = ""
    .TxtEllipsoidalH2.Text = ""

    .txtMarkPurpose.ListIndex = -1
    .txtMarkType.ListIndex = -1
     .txtAuthority = ""
    .txtEstablished = ""
    .txtLastRecover = ""
    .txtMarkStatus.ListIndex = -1
   ' .txtProjectNo.Text = ""
    '.txtFileNo.Text = ""
    .TxtProvince.Locked = False
    .TxtMunicipality.Locked = False
    .TxtBarangay.Locked = False
    .TxtFixingMethod.Locked = False

    
    .txtMarkPurpose.Locked = False
    .txtMarkType.Locked = False
    .txtMarkStatus.Locked = False
    .TxtMSL.Locked = False
    
    .TxtLatD = ""
    .TxtLatM = ""
    .TxtLatS = ""
    .TxtLongD = ""
    .TxtLongM = ""
    .TxtLongS = ""
    
    'WGS84
    .TxtDLat = ""
    .TxtMLat = ""
    .TxtSLat = ""
    .TxtDLong = ""
    .TxtMLong = ""
    .TxtSLong = ""
    
End With
End Sub

Public Sub BlankFormGravity()


With FrmGCPDS
   
    .TextBoxGravityName = ""
    .TextBoxGravityRegion = ""
    .TextBoxGravityProvince.Text = ""
    .TextBoxGravityMunicipality.Text = ""
    .TextBoxGravityBarangay.Text = ""
    .ComboBoxOrderGravity.ListIndex = -1
    .TextBoxGravityElevation.Text = ""
    .TextBoxGravityLatitude.Text = ""
    .TextBoxGravityLongitude.Text = ""
    .TextBoxObservedValues.Text = ""
 
   .TextBoxGravityDescription.Text = ""
   

    
End With
End Sub

Public Sub BlankFormBM()

With FrmGCPDS
    
    .TxtEName = ""
    .TxtERegion = ""
    .TxtEProvince = ""
    .TxtEMunicipality = ""
    .TxtEBarangay = ""
    .TxtBMOrder.ListIndex = -1
    .TxtEDatum.ListIndex = -1
    .TxtEIsland = ""
    .TxtElevation = ""
    .TxtBMPlus = ""
    .TextBoxLatitude = ""
    .TextBoxLongitude = ""
    .TxtEFix = ""
    .BMMarkType.ListIndex = -1
    .BMMarkPurpose.ListIndex = -1
    .BMMarkStatus.ListIndex = -1
    .TxtBMDateComputed = ""
    .TxtBMDateLastRecovered = ""
    .TxtBMDescription = ""
    .TxtBMDateEstablished = ""
    .TxtBMDateOfEntry = ""
    .TxtBMAuthority = ""
    .TxtElevationAuthority = ""
    .LabelEncoder2.Caption = ""
    .LabelDateUpdated2 = ""
End With
End Sub

Public Sub EnableFieldsGravity()

With FrmGCPDS
    
     .TextBoxGravityName.BorderStyle = fmBorderStyleSingle
    .TextBoxGravityRegion.BorderStyle = fmBorderStyleSingle
    .TextBoxGravityProvince.BorderStyle = fmBorderStyleSingle
    .TextBoxGravityMunicipality.BorderStyle = fmBorderStyleSingle
    .TextBoxGravityBarangay.BorderStyle = fmBorderStyleSingle
    .ComboBoxOrderGravity.BorderStyle = fmBorderStyleSingle
    .TextBoxGravityElevation.BorderStyle = fmBorderStyleSingle
    .TextBoxGravityLatitude.BorderStyle = fmBorderStyleSingle
    .TextBoxGravityLongitude.BorderStyle = fmBorderStyleSingle
    .TextBoxObservedValues.BorderStyle = fmBorderStyleSingle
   
   .TextBoxGravityDescription.BorderStyle = fmBorderStyleSingle
   
   .TextBoxGravityName.Locked = False
 
    .ComboBoxOrderGravity.Locked = False
    .TextBoxGravityElevation.Locked = False
    .TextBoxGravityLatitude.Locked = False
    .TextBoxGravityLongitude.Locked = False
    .TextBoxObservedValues.Locked = False
   
   .TextBoxGravityDescription.Locked = False
   
   .RaveButtonsUnits.Visible = True
   .RaveLocation.Visible = True
    


End With
End Sub

Public Sub DisableFieldsGravity()

With FrmGCPDS
    
     .TextBoxGravityName.BorderStyle = fmBorderStyleNone
    .TextBoxGravityRegion.BorderStyle = fmBorderStyleNone
    .TextBoxGravityProvince.BorderStyle = fmBorderStyleNone
    .TextBoxGravityMunicipality.BorderStyle = fmBorderStyleNone
    .TextBoxGravityBarangay.BorderStyle = fmBorderStyleNone
    .ComboBoxOrderGravity.BorderStyle = fmBorderStyleNone
    .TextBoxGravityElevation.BorderStyle = fmBorderStyleNone
    .TextBoxGravityLatitude.BorderStyle = fmBorderStyleNone
    .TextBoxGravityLongitude.BorderStyle = fmBorderStyleNone
    .TextBoxObservedValues.BorderStyle = fmBorderStyleNone
   
  ' .TextBoxGravityDescription.BorderStyle = fmBorderStyleNone
    
   .TextBoxGravityName.BackStyle = fmBackStyleTransparent
    .TextBoxGravityRegion.BackStyle = fmBackStyleTransparent
    .TextBoxGravityProvince.BackStyle = fmBackStyleTransparent
    .TextBoxGravityMunicipality.BackStyle = fmBackStyleTransparent
    .TextBoxGravityBarangay.BackStyle = fmBackStyleTransparent
    .ComboBoxOrderGravity.BackStyle = fmBackStyleTransparent
    .TextBoxGravityElevation.BackStyle = fmBackStyleTransparent
    .TextBoxGravityLatitude.BackStyle = fmBackStyleTransparent
    .TextBoxGravityLongitude.BackStyle = fmBackStyleTransparent
    .TextBoxObservedValues.BackStyle = fmBackStyleTransparent
    
  ' .TextBoxGravityDescription.BackStyle = fmBackStyleTransparent
   
      .TextBoxGravityName.Locked = True
    .TextBoxGravityRegion.Locked = True
    .TextBoxGravityProvince.Locked = True
    .TextBoxGravityMunicipality.Locked = True
    .TextBoxGravityBarangay.Locked = True
    .ComboBoxOrderGravity.Locked = True
    .TextBoxGravityElevation.Locked = True
    .TextBoxGravityLatitude.Locked = True
    .TextBoxGravityLongitude.Locked = True
    .TextBoxObservedValues.Locked = True
   
   .TextBoxGravityDescription.Locked = True
   .RaveButtonsUnits.Visible = False
   .RaveLocation.Visible = False
End With
End Sub




Public Sub EnableFields()

 With FrmGCPDS
   
    .TxtName.Locked = False
    .TxtIsland.Locked = False
    .TxtRegion.Locked = False
    .TxtProvince.Locked = False
    .TxtMunicipality.Locked = False
    .TxtBarangay.Locked = False
 
    .TxtLatD.Locked = False
    .TxtLatM.Locked = False
    .TxtLatS.Locked = False
    .TxtLongD.Locked = False
    .TxtLongM.Locked = False
    .TxtLongS.Locked = False
    
    'WGS84
    .TxtDLat.Locked = False
    .TxtMLat.Locked = False
    .TxtSLat.Locked = False
    .TxtDLong.Locked = False
    .TxtMLong.Locked = False
    .TxtSLong.Locked = False
    
    
    'PTM
    .TxtNorthing.Locked = False
    .TxtEasting.Locked = False
    .TxtZone.Locked = False
    'UTM
    .TextBoxUTMNorthing.Locked = False
    .TextBoxUTMEasting.Locked = False
    .TextBoxUTMZone.Locked = False
    
    .txtOrder.Locked = False
    '.TxtDatum.Locked = False
    .TxtEstablishedBy.Locked = False
    
    '.TxtRef.Locked = False
    
    .TxtFixingMethod.Locked = False
    .TxtDateEntry.Locked = False
    .TxtDateComputed.Locked = False
    
  
    .txtEllipsoidalH.Locked = False
    .TxtEllipsoidalH2.Locked = False
   
   
    
    .txtMarkPurpose.Locked = False
    .txtMarkType.Locked = False
    .txtAuthority.Locked = False
    .txtEstablished.Locked = False
    .txtYear.Locked = False
    .ComboBoxMonth.Locked = False
    .ComboBoxDay.Locked = False
    .txtLastRecover.Locked = False
    .txtMarkStatus.Locked = False
    .TxtMSL.Locked = False

    .cmdLocation.Enabled = True
    
    'Border
    
    .TxtName.BorderStyle = fmBorderStyleSingle
  
    .TxtIsland.BorderStyle = fmBorderStyleSingle
    .TxtRegion.BorderStyle = fmBorderStyleSingle
    .TxtProvince.BorderStyle = fmBorderStyleSingle
    .TxtMunicipality.BorderStyle = fmBorderStyleSingle
    .TxtBarangay.BorderStyle = fmBorderStyleSingle
    .txtOrder.BorderStyle = fmBorderStyleSingle
    
    .TxtLatD.BorderStyle = fmBorderStyleSingle
    .TxtLatM.BorderStyle = fmBorderStyleSingle
    .TxtLatS.BorderStyle = fmBorderStyleSingle
    .TxtLongD.BorderStyle = fmBorderStyleSingle
    .TxtLongM.BorderStyle = fmBorderStyleSingle
    .TxtLongS.BorderStyle = fmBorderStyleSingle
    
    .txtEllipsoidalH.BorderStyle = fmBorderStyleSingle
    
    .TxtDLat.BorderStyle = fmBorderStyleSingle
    .TxtMLat.BorderStyle = fmBorderStyleSingle
    .TxtSLat.BorderStyle = fmBorderStyleSingle
    .TxtDLong.BorderStyle = fmBorderStyleSingle
    .TxtMLong.BorderStyle = fmBorderStyleSingle
    .TxtSLong.BorderStyle = fmBorderStyleSingle
    
    .TxtEllipsoidalH2.BorderStyle = fmBorderStyleSingle
    
    
    .TxtNorthing.BorderStyle = fmBorderStyleSingle
    .TxtEasting.BorderStyle = fmBorderStyleSingle
    .TxtZone.BorderStyle = fmBorderStyleSingle
    
    'UTM
    .TextBoxUTMNorthing.BorderStyle = fmBorderStyleSingle
    .TextBoxUTMEasting.BorderStyle = fmBorderStyleSingle
    .TextBoxUTMZone.BorderStyle = fmBorderStyleSingle
    
    .TxtDateEntry.BorderStyle = fmBorderStyleSingle
    .TxtDateComputed.BorderStyle = fmBorderStyleSingle
    .txtLastRecover.BorderStyle = fmBorderStyleSingle
    .TxtFixingMethod.BorderStyle = fmBorderStyleSingle
    
    .txtMarkPurpose.BorderStyle = fmBorderStyleSingle
    .txtMarkType.BorderStyle = fmBorderStyleSingle
    .txtMarkStatus.BorderStyle = fmBorderStyleSingle
    .TxtMSL.BorderStyle = fmBorderStyleSingle
    
    .TxtEstablishedBy.BorderStyle = fmBorderStyleSingle
    .txtEstablished.BorderStyle = fmBorderStyleSingle
    .ComboBoxMonth.BorderStyle = fmBorderStyleSingle
    .ComboBoxDay.BorderStyle = fmBorderStyleSingle
    .txtYear.BorderStyle = fmBorderStyleSingle
    .txtAuthority.BorderStyle = fmBorderStyleSingle

    .RaveTranform2WGS84.Visible = True
    .RaveTranform2PRS92.Visible = True
    .RaveTranformToPTM.Visible = True
    .RaveTranform2UTM.Visible = True
    
    .cmdLocation.Visible = True
 End With
End Sub
Public Sub DisableFields()

 With FrmGCPDS
    .SSTab1.TabEnabled(1) = True
    .SSTab1.TabEnabled(2) = True
   
    .TxtName.Locked = True
    .TxtIsland.Locked = False
    .TxtRegion.Locked = False
    .TxtProvince.Locked = True
    .TxtMunicipality.Locked = True
    .TxtBarangay.Locked = True
    .TxtLatitude.Locked = True
    .TxtLongitude.Locked = True
    .txtOrder.Locked = True
    .TxtEasting.Locked = True
    .TxtNorthing.Locked = True
    .TxtZone.Locked = True
    'UTM
    .TextBoxUTMNorthing.Locked = True
    .TextBoxUTMEasting.Locked = True
    .TextBoxUTMZone.Locked = True
    
    '.TxtDatum.Locked = True
    .TxtEstablishedBy.Locked = True
   
    .TxtRef.Locked = True
    
    .TxtFixingMethod.Locked = True
    .TxtDateEntry.Locked = True
    .TxtDateComputed.Locked = True
    
    .txtEllipsoidalH.Locked = True
    .TxtEllipsoidalH2.Locked = True
    
    .txtMarkPurpose.Locked = True
    .txtMarkType.Locked = True
    .TxtMSL.Locked = True
    .txtAuthority.Locked = True
    .txtEstablished.Locked = True
    .txtYear.Locked = True
    .ComboBoxMonth.Locked = True
    .ComboBoxDay.Locked = True
    .txtLastRecover.Locked = True
    .txtMarkStatus.Locked = True
    '.txtProjectNo.Locked = True
    '.txtFileNo.Locked = True
    
    'Border
    .TxtName.BorderStyle = fmBorderStyleNone
    
    .TxtIsland.BorderStyle = fmBorderStyleNone
    .TxtRegion.BorderStyle = fmBorderStyleNone
    .TxtProvince.BorderStyle = fmBorderStyleNone
    .TxtMunicipality.BorderStyle = fmBorderStyleNone
    .TxtBarangay.BorderStyle = fmBorderStyleNone
    .txtOrder.BorderStyle = fmBorderStyleNone
    
    .TxtLatD.BorderStyle = fmBorderStyleNone
    .TxtLatM.BorderStyle = fmBorderStyleNone
    .TxtLatS.BorderStyle = fmBorderStyleNone
    .TxtLongD.BorderStyle = fmBorderStyleNone
    .TxtLongM.BorderStyle = fmBorderStyleNone
    .TxtLongS.BorderStyle = fmBorderStyleNone
    
    .txtEllipsoidalH.BorderStyle = fmBorderStyleNone
    
    .TxtDLat.BorderStyle = fmBorderStyleNone
    .TxtMLat.BorderStyle = fmBorderStyleNone
    .TxtSLat.BorderStyle = fmBorderStyleNone
    .TxtDLong.BorderStyle = fmBorderStyleNone
    .TxtMLong.BorderStyle = fmBorderStyleNone
    .TxtSLong.BorderStyle = fmBorderStyleNone
    
    .TxtEllipsoidalH2.BorderStyle = fmBorderStyleNone
    
    .TxtDateEntry.BorderStyle = fmBorderStyleNone
    .TxtDateComputed.BorderStyle = fmBorderStyleNone
    .txtLastRecover.BorderStyle = fmBorderStyleNone
    .TxtFixingMethod.BorderStyle = fmBorderStyleNone
    
    .txtMarkPurpose.BorderStyle = fmBorderStyleNone
    .txtMarkType.BorderStyle = fmBorderStyleNone
    .txtMarkStatus.BorderStyle = fmBorderStyleNone
    .TxtMSL.BorderStyle = fmBorderStyleNone
    
    .TxtEstablishedBy.BorderStyle = fmBorderStyleNone
    .txtEstablished.BorderStyle = fmBorderStyleNone
    .ComboBoxMonth.BorderStyle = fmBorderStyleNone
    .ComboBoxDay.BorderStyle = fmBorderStyleNone
    .txtYear.BorderStyle = fmBorderStyleNone
    .txtAuthority.BorderStyle = fmBorderStyleNone
    
    '.TxtDatum.BorderStyle = fmBorderStyleNone
    .TxtNorthing.BorderStyle = fmBorderStyleNone
    .TxtEasting.BorderStyle = fmBorderStyleNone
    .TxtZone.BorderStyle = fmBorderStyleNone
    
    'UTM
    .TextBoxUTMNorthing.BorderStyle = fmBorderStyleNone
    .TextBoxUTMEasting.BorderStyle = fmBorderStyleNone
    .TextBoxUTMZone.BorderStyle = fmBorderStyleNone
    
    .RaveTranform2WGS84.Visible = False
    .RaveTranform2PRS92.Visible = False
    .RaveTranformToPTM.Visible = False
    .RaveTranform2UTM.Visible = False
    
    .cmdLocation.Visible = False
 End With
End Sub





Public Function GetProvinceCode(Province As String) As String
Dim TempRecordSet As New ADODB.Recordset
    TempRecordSet.Open "Select ProvinceAlpha from Provmast where Prov_name='" & Replace(Province, "'", "''") & "'", cnn, adOpenStatic, adLockOptimistic
    If TempRecordSet.RecordCount > 0 Then
       GetProvinceCode = TempRecordSet("ProvinceAlpha").Value
       Else
       GetProvinceCode = ""
    End If
End Function


Public Sub DBaseToAccess(Directory As String, filename As String)
Dim rstTrueProvince As ADODB.Recordset
Dim rstregion As ADODB.Recordset
    Dim w As Integer
    Dim dbase4_cnn As New ADODB.Connection
    Dim dbase4_rst As New ADODB.Recordset
    Dim x As Integer
    Dim Stat_Name As String
    Dim Region As String
    Dim Province As String
    Dim Municipal As String
    Dim Barangay As String
    Dim Stat_New As String
    Dim Island As String
    Dim DLat As String
    Dim MLat As String
    Dim SLat As String
    Dim DLong As String
    Dim MLong As String
    Dim SLong As String
    Dim H_Order As String
    Dim H_Datum As String
    Dim Hor_Authty As String
    Dim E_H_Datum As String
    Dim H_Ref As String
    Dim H_Fix As String
    Dim H_Date_Ety As String
    Dim H_Date_Com As String
    Dim Elevation As Double
    Dim E_Order As String
    Dim E_Fix As Integer
    Dim E_datum As Integer
    Dim E_E_Datum As String
    Dim E_date_com As String
    Dim Ell_Hgt As String
    Dim Elv_Authty As String
    Dim E_Date_Ety As String
    Dim Mark_Const As String
    Dim Mark_Type As String
    Dim Mark_Stat As Integer
    Dim Date_Est As String
    Dim Date_Las_R As String
    'Dim File_ID As String
    'Dim Project_No As String
    Dim Authority As String
    Dim Description As String
    'Dim Longitude As Double
    'Dim Latitude As Double
    
    
    
    dbase4_cnn.ConnectionString = "Driver=Microsoft dBase Driver (*.dbf);DBQ=" & Directory
   
    'On Error GoTo hell
    
    dbase4_cnn.Open
    
    dbase4_rst.Open "select * from " & Trim(Directory) & "\" & Trim(filename) & ".mdb Where h_ref='OLD'", dbase4_cnn, adOpenStatic, adLockOptimistic

    
    If dbase4_rst.RecordCount > 0 Then 'IF no records, dont set the progressbar
        FrmDbase4.ProgressBar2.Min = 0
        FrmDbase4.ProgressBar2.Max = dbase4_rst.RecordCount
    End If
    
    For x = 1 To dbase4_rst.RecordCount
        DoEvents
        DbaseCount = DbaseCount + 1
        FrmDbase4.ProgressBar2.Value = x
        FrmDbase4.Caption = "Dbase IV - " & Format(DbaseCount, "#,##0")
        FrmDbase4.LlbStation.Caption = "Adding: " & StrConv(dbase4_rst("Stat_name").Value, vbProperCase)
        
        
       ' Stat_Name = IIf(IsNull(dbase4_rst("Stat_name").Value) = False, trim(dbase4_rst("Stat_name").Value, "")
        'Station Name
        
        
        If IsNull(dbase4_rst("Stat_name").Value) = False Then
           Stat_Name = Replace(Replace(dbase4_rst("Stat_name").Value, "'", "''"), ",", " ")
           Else
           Stat_Name = ""
        End If
        
        
        'Province
        Set rstTrueProvince = New ADODB.Recordset
            rstTrueProvince.Open "Select name,psgc_cd from psgc where ACRONYM='" & filename & "'", cnn, adOpenStatic
        Set rstregion = New ADODB.Recordset
      
            rstregion.Open "Select name from psgc where psgc_cd='" & Mid(rstTrueProvince("psgc_cd"), 1, 2) & "0000000" & "'", cnn, adOpenStatic
           
           Province = Replace(rstTrueProvince("name").Value, "'", "''")
           Region = Replace(rstregion("name").Value, "'", "''")
          
        
        
        
        'Municipality
        If IsNull(dbase4_rst("Municipal").Value) = False Then
           Municipal = Replace(dbase4_rst("Municipal").Value, "'", "''")
           Else
           Municipal = ""
        End If
        'Barangay
        If IsNull(dbase4_rst("Barangay").Value) = False Then
           Barangay = Replace(dbase4_rst("barangay").Value, "'", "''")
           Else
           Barangay = ""
        End If
        'Station Number
        If IsNull(dbase4_rst("Stat_new").Value) = False Then
           Stat_New = Replace(dbase4_rst("stat_new").Value, "'", "''")
           Else
           Stat_New = 0
        End If
        'Island
        If IsNull(dbase4_rst("Island").Value) = False Then
           Island = Replace(dbase4_rst("island").Value, "'", "''")
           Else
           Island = ""
        End If
        
        'Degree Latitude
        If IsNull(dbase4_rst("d_lat").Value) = False Then
           DLat = Replace(dbase4_rst("d_lat").Value, "'", "''")
           Else
           DLat = 0
           End If
        
        'Minutes Latitude
        If IsNull(dbase4_rst("m_lat").Value) = False Then
           MLat = Replace(dbase4_rst("m_lat").Value, "'", "''")
           Else
           MLat = 0
        End If
        
        'Seconds Latitude
        If IsNull(dbase4_rst("s_lat").Value) = False Then
           SLat = Replace(dbase4_rst("s_lat").Value, "'", "''")
           Else
           SLat = 0
        End If
        
        'Degree Longitude
        If IsNull(dbase4_rst("d_long").Value) = False Then
           DLong = Replace(dbase4_rst("d_long").Value, "'", "''")
           Else
           DLong = 0
        End If
        
        'Minutes Longitude
        If IsNull(dbase4_rst("m_long").Value) = False Then
           MLong = Replace(dbase4_rst("m_long").Value, "'", "''")
           Else
           MLong = 0
        End If
        
        'Seconds Longitude
        If IsNull(dbase4_rst("s_long").Value) = False Then
           SLong = Replace(dbase4_rst("s_long").Value, "'", "''")
           Else
           SLong = 0
        End If
        
         'Horizontal Order
        If IsNull(dbase4_rst("h_order").Value) = False Then
           H_Order = Replace(dbase4_rst("h_order").Value, "'", "''")
           Else
           H_Order = 0
        End If
        
        
         'Horizontal Datum
        If IsNull(dbase4_rst("h_datum").Value) = False Then
           H_Datum = Replace(dbase4_rst("h_datum").Value, "'", "''")
           Else
           H_Datum = 0
        End If
        
         'Horizontal Authority
        If IsNull(dbase4_rst("hor_authty").Value) = False Then
           Hor_Authty = Replace(dbase4_rst("hor_authty").Value, "'", "''")
           Else
           Hor_Authty = ""
        End If
        
         'Epoch of Datum
        If IsNull(dbase4_rst("e_h_datum").Value) = False Then
           E_H_Datum = Replace(dbase4_rst("e_h_datum").Value, "'", "''")
           Else
           E_H_Datum = ""
        End If
        
         'H Ref
        If IsNull(dbase4_rst("h_ref").Value) = False Then
           H_Ref = Replace(dbase4_rst("h_ref").Value, "'", "''")
           Else
           H_Ref = ""
        End If
        
         'H Fixing Method
        If IsNull(dbase4_rst("h_fix").Value) = False Then
           H_Fix = Replace(dbase4_rst("h_fix").Value, "'", "''")
           Else
          H_Fix = 0
        End If
        
        'H Date of Entry
        If IsNull(dbase4_rst("h_date_ety").Value) = False Then
           H_Date_Ety = "'" & Replace(dbase4_rst("h_date_ety").Value, "'", "''") & "'"
           Else
           H_Date_Ety = "Null"
        End If
        
        'H Date Computed
        If IsNull(dbase4_rst("h_date_com").Value) = False Then
           H_Date_Com = "'" & Replace(dbase4_rst("h_date_com").Value, "'", "''") & "'"
           Else
           H_Date_Com = "Null"
        End If
        
        'Elevation
        If IsNull(dbase4_rst("elevation").Value) = False Then
           Elevation = CDec(dbase4_rst("elevation"))
           Else
           Elevation = 0
        End If

        'E Order
        If IsNull(dbase4_rst("e_order").Value) = False Then
           E_Order = CDec(dbase4_rst("e_order").Value)
           Else
           E_Order = 0
        End If

        'E Fixing Method
        If IsNull(dbase4_rst("E_fix").Value) = False Then
           E_Fix = CDec(dbase4_rst("E_fix").Value)
           Else
           E_Fix = 0
        End If

        'E Datum
        If IsNull(dbase4_rst("e_datum").Value) = False Then
           E_datum = Replace(dbase4_rst("e_datum").Value, "'", "''")
           Else
           E_datum = 0
        End If

'        'Epoch of E datum
'        If IsNull(dbase4_rst("e_e_datum").Value) = False Then
'           E_E_Datum = Replace(dbase4_rst("e_e_datum").Value, "'", "''")
'           Else
'           E_E_Datum = ""
'        End If
        
        'E Date Computed
        If IsNull(dbase4_rst("e_date_com").Value) = False And IsDate(dbase4_rst("e_date_com")) = True Then
           E_date_com = "'" & Replace(dbase4_rst("e_date_com").Value, "'", "''") & "'"
           Else
           E_date_com = "Null"
        End If

         'E Date Entry
        If IsNull(dbase4_rst("e_date_ety").Value) = False And IsDate(dbase4_rst("e_date_ety")) = True Then
           E_Date_Ety = "'" & Replace(dbase4_rst("e_date_ety").Value, "'", "''") & "'"
           Else
           E_Date_Ety = "Null"
        End If

        'Elevation Authority
        If IsNull(dbase4_rst("elv_authty").Value) = False Then
           Elv_Authty = Replace(dbase4_rst("elv_authty").Value, "'", "''")
           Else
           Elv_Authty = ""
        End If
        
        'Ellipsoidal Height
        If IsNull(dbase4_rst("ell_hgt").Value) = False Then
           Ell_Hgt = Replace(dbase4_rst("ell_hgt").Value, "'", "''")
           Else
           Ell_Hgt = 0
        End If
        
         'Mark Constant
        If IsNull(dbase4_rst("mark_const").Value) = False Then
           Mark_Type = Replace(dbase4_rst("mark_const").Value, "'", "''")
           Else
           Mark_Type = 0
        End If
        
        'Mark Type
        If IsNull(dbase4_rst("mark_type").Value) = False Then
           Mark_Const = Replace(dbase4_rst("mark_type").Value, "'", "''")
           Else
           Mark_Const = 0
        End If
        
        'Mark Status
        If dbase4_rst("mark_stat").Value = True Then
           Mark_Stat = 1
           Else
           Mark_Stat = 0
          
        End If
        
        'Date Established
        If IsNull(dbase4_rst("Date_Est").Value) = False And IsDate(dbase4_rst("Date_Est")) = True Then
           Date_Est = "'" & Replace(dbase4_rst("Date_Est").Value, "'", "''") & "'"
           Else
           Date_Est = "Null"
        End If
        
        'Date Last Recover
        If IsNull(dbase4_rst("date_las_r").Value) = False Then
           Date_Las_R = "'" & Replace(dbase4_rst("date_las_r").Value, "'", "''") & "'"
           Else
           Date_Las_R = "Null"
        End If
        

        
        'Authority
        If IsNull(dbase4_rst("Authority").Value) = False Then
           Authority = Replace(dbase4_rst("Authority").Value, "'", "''")
           Else
           Authority = ""
        End If
        
        'Description
        If IsNull(dbase4_rst("Descriptio").Value) = False Then
           Description = Replace(dbase4_rst("Descriptio").Value, "'", "''")
           Description = Replace(Description, Chr(0), "")
           
          For w = 10 To 12
                Description = Replace(Description, Chr(w), "")
          Next
           Description = Replace(Description, Chr(13), vbCrLf)
           Else
           Description = ""
        End If
        
        
        'Long
'        Longitude = (DLong) + (MLong / 60) + (SLong / 3600)
'        Latitude = (DLat) + (MLat / 60) + (SLat / 3600)
        
        'DrawPoints CDbl(Longitude), CDbl(Latitude)

        
       
        
       
If Elevation <> 0 Then
    
    InsertRecords "benchmarks", "stat_name", "region", "province", "municipal", "barangay", "stat_new", "Island", "Elevation", "E_order", "E_Fix", "E_Datum", "E_Date_Ety", "E_Date_Com", "Elv_Authty", "Mark_Type", "Mark_Const", "Mark_Stat", "Date_Est", "Date_las_R", "Authority", "Description", _
        "'" & Stat_Name & "'", "'" & Region & "'", "'" & Province & "'", "'" & Municipal & "'", "'" & Barangay & "'", Stat_New, "'" & Island & "'", Elevation, E_Order, E_Fix, E_datum, E_Date_Ety, E_date_com, "'" & Elv_Authty & "'", Mark_Type, Mark_Const, Mark_Stat, Date_Est, Date_Las_R, "'" & Authority & "'", "'" & Description & "'"
    
Else
    
     InsertRecords "triangulation", "stat_name", "region", "province", "municipal", "barangay", "stat_new", "Island", "D_lat", "M_Lat", "S_Lat", "D_Long", "M_Long", "S_Long", "h_order", "h_datum", "hor_authty", "h_ref", "H_fix", "H_Date_Ety", "H_Date_Com", "Ell_hgt", "Mark_Type", "Mark_Const", "mark_stat", "Date_Est", "Date_las_R", "Authority", "Description", _
        "'" & Stat_Name & "'", "'" & Region & "'", "'" & Province & "'", "'" & Municipal & "'", "'" & Barangay & "'", Stat_New, "'" & Island & "'", DLat, MLat, SLat, DLong, MLong, SLong, H_Order, H_Datum, "'" & Hor_Authty & "'", "'" & H_Ref & "'", H_Fix, H_Date_Ety, H_Date_Com, Ell_Hgt, Mark_Type, Mark_Const, Mark_Stat, Date_Est, Date_Las_R, "'" & Authority & "'", "'" & Description & "'"
        
End If


       
  
        
'        rstProvinceInfo.AddNew
'        rstProvinceInfo("stat_name").Value = dbase4_rst("Stat_name").Value
'        rstProvinceInfo("province") = Filename
'        rstProvinceInfo("municipal") = dbase4_rst("municipal")
'        rstProvinceInfo("barangay") = dbase4_rst("barangay")
'        rstProvinceInfo("stat_new") = dbase4_rst("stat_new")
'        rstProvinceInfo("island") = dbase4_rst("island")
'        rstProvinceInfo("d_lat") = dbase4_rst("d_lat")
'        rstProvinceInfo("m_lat") = dbase4_rst("m_lat")
'        rstProvinceInfo("s_lat") = dbase4_rst("s_lat")
'        rstProvinceInfo("d_long") = dbase4_rst("d_long")
'        rstProvinceInfo("m_long") = dbase4_rst("m_long")
'        rstProvinceInfo("s_long") = dbase4_rst("s_long")
'        rstProvinceInfo("h_order") = dbase4_rst("h_order")
'        rstProvinceInfo("h_datum") = dbase4_rst("h_datum")
'        rstProvinceInfo("hor_authty") = dbase4_rst("hor_authty")
'        rstProvinceInfo("e_h_datum") = dbase4_rst("e_h_datum")
'        rstProvinceInfo("h_ref") = dbase4_rst("h_ref")
'        rstProvinceInfo("h_fix") = dbase4_rst("h_fix")
'        rstProvinceInfo("h_date_ety") = dbase4_rst("h_date_ety")
'        rstProvinceInfo("h_date_com") = dbase4_rst("h_date_com")
'        rstProvinceInfo("elevation") = dbase4_rst("elevation")
'        rstProvinceInfo("e_order") = dbase4_rst("e_order")
'        rstProvinceInfo("e_fix") = dbase4_rst("e_fix")
'        rstProvinceInfo("e_datum") = dbase4_rst("e_datum")
'        rstProvinceInfo("e_e_datum") = dbase4_rst("e_e_datum")
'        rstProvinceInfo("e_date_com") = dbase4_rst("e_date_com")
'        rstProvinceInfo("ell_hgt") = dbase4_rst("ell_hgt")
'        rstProvinceInfo("elv_authty") = dbase4_rst("elv_authty")
'        rstProvinceInfo("e_date_ety") = dbase4_rst("e_date_ety")
'        rstProvinceInfo("mark_const") = dbase4_rst("mark_type")
'        rstProvinceInfo("date_est") = dbase4_rst("date_est")
'        rstProvinceInfo("project_no") = dbase4_rst("project_no")
'        rstProvinceInfo("mark_type") = dbase4_rst("mark_const")
'        rstProvinceInfo("date_las_r") = dbase4_rst("date_las_r")
'        rstProvinceInfo("file_id") = dbase4_rst("file_id")
'        rstProvinceInfo("authority") = dbase4_rst("authority")
'        rstProvinceInfo("mark_stat") = dbase4_rst("mark_stat")
'        rstProvinceInfo("description") = dbase4_rst("descriptio")
'
'
'        rstProvinceInfo.Update
        dbase4_rst.MoveNext
    Next
       
    Exit Sub
Hell:

    MsgBox "Invalid Path for " & filename & " Database.", vbInformation, Directory
    
      
End Sub

Public Function GetProvinceName(Province_Code As String) As String
Dim ProvinceRst As New ADODB.Recordset
    ProvinceRst.Open "Select prov_name from ProvMast where Provincealpha='" & Province_Code & "'", cnn, adOpenStatic, adLockBatchOptimistic
               
           If ProvinceRst.RecordCount > 0 Then
            If IsNull(ProvinceRst.Fields(0).Value) = False Then
                GetProvinceName = ProvinceRst.Fields(0).Value
                Else
                GetProvinceName = ""
            End If
           End If
End Function

Public Sub InsertRecords(Xtable As String, ParamArray TempFields())
 Dim Fields()
 Dim Values()
 Dim cntr As Integer

On Error GoTo Hell
 
 ReDim Fields(1 To (UBound(TempFields) + 1) / 2)
 ReDim Values(1 To (UBound(TempFields) + 1) / 2)
 
 For cntr = 1 To UBound(Fields)
     Fields(cntr) = TempFields(cntr - 1)
 Next cntr
 
  For cntr = 1 To UBound(Values)
     Values(cntr) = TempFields(UBound(Fields) + cntr - 1)
  Next cntr
cnn.Execute "INSERT INTO " & Xtable & " (" & Join(Fields, ",") & ")" & _
                  " VALUES" & " (" & Join(Values, ",") & ")"


Exit Sub
Hell:

If Err.Number = -2147217900 Or Err.Number = -2147217873 Then
    
   
      cnn.Execute "INSERT INTO duplicategcps  (" & Join(Fields, ",") & ")" & _
                  " VALUES" & " (" & Join(Values, ",") & ")"
                  
ElseIf Err.Number = -2147467259 Then
    FrmDbase4.Label1 = "Blank Station Name"
    
    
Else
    MsgBox Err.Description & " " & Err.Number
End If

End Sub




Public Function IsDuplicateStation(stationName As String) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "Select stat_name from geoprov where stat_name='" & Trim(Replace(stationName, "'", "''")) & "'", cnn, adOpenStatic
    
    If rst.RecordCount > 0 Then
        IsDuplicateStation = True
        Else
        IsDuplicateStation = False
    End If
    
End Function








Public Sub SaveRegion()
SaveSetting App.EXEName, "Region", "Region", "01"
End Sub




Public Sub PRS92_TO_WGS84()
Const PI = 3.14159265358979
Const O = PI / 180
Dim Ellipsoidal_Ht As Double

Dim dx, DY, DZ, s, RX, RY, RZ, a, A2, E2, E22, N1, X1, Y1, P1, L1, Z1, X2, Y2, Z2, R, H2, P2, P, N2, K, A1, DG, MN, MN1, SC, B2, L2 As Double

P1 = (FrmGCPDS.TxtLatD + (FrmGCPDS.TxtLatM / 60) + (FrmGCPDS.TxtLatS / 3600)) * PI / 180
L1 = (FrmGCPDS.TxtLongD + (FrmGCPDS.TxtLongM / 60) + (FrmGCPDS.TxtLongS / 3600)) * PI / 180
Ellipsoidal_Ht = FrmGCPDS.txtEllipsoidalH

'7 parameters
'DX = -127.62195 'Translation X
'DY = -67.24478  'Translation Y
'DZ = -47.04305  'Translation Z
'S = -1.06002    'Scale Factor
'RX = 3.06762 / 3600 'Rotation X
'RY = -4.90291 / 3600 'Rotation Y
'RZ = -1.5779 / 3600 'Rotation Z

dx = -127.62153 'Translation X
DY = -67.2434   'Translation Y
DZ = -47.04738  'Translation Z
s = -1.06002112    'Scale Factor
RX = 3.068038 / 3600 'Rotation X
RY = -4.902977 / 3600 'Rotation Y
RZ = -1.578073 / 3600 'Rotation Z


a = 6378206.4
E2 = 0.006768658
A2 = 6378137#
E22 = 0.00669438

N1 = a / Sqr(1 - E2 * (Sin(P1)) ^ 2)
X1 = (N1 + Ellipsoidal_Ht) * Cos(P1) * Cos(L1)
Y1 = (N1 + Ellipsoidal_Ht) * Cos(P1) * Sin(L1)
Z1 = (N1 * (1 - E2) + Ellipsoidal_Ht) * Sin(P1)
X2 = dx + (1 + s * 10 ^ -6) * (X1 + RZ * O * Y1 - RY * O * Z1)
Y2 = DY + (1 + s * 10 ^ -6) * (-RZ * O * X1 + Y1 + RX * O * Z1)
Z2 = DZ + (1 + s * 10 ^ -6) * (RY * O * X1 - RX * O * Y1 + Z1)
R = Sqr(X2 ^ 2 + Y2 ^ 2)
H2 = 0
P2 = 0
P = P2
N2 = A2 / Sqr(1 - E22 * (Sin(P2)) ^ 2)
P2 = Atn((N2 + H2) * Z2 / ((N2 * (1 - E22) + H2) * R))
H2 = R / Cos(P2) - N2
K = Abs(P - P2)

Do Until K < 0.000000009999999
    P = P2
    N2 = A2 / Sqr(1 - E22 * (Sin(P2)) ^ 2)
    P2 = Atn((N2 + H2) * Z2 / ((N2 * (1 - E22) + H2) * R))
    H2 = R / Cos(P2) - N2
    K = Abs(P - P2)
Loop

A1 = P2 * 180 / PI
DG = Int(A1)
MN1 = (A1 - DG) * 60
MN = Int(MN1)
SC = (MN1 - MN) * 60


FrmGCPDS.TxtDLat = DG
FrmGCPDS.TxtMLat = MN
FrmGCPDS.TxtSLat = Round(SC, 5)

B2 = (Atn(Y2 / X2)) * 180 / PI
L2 = B2 + 180
A1 = L2


DG = Int(A1)
MN1 = (A1 - DG) * 60
MN = Int(MN1)
SC = (MN1 - MN) * 60

FrmGCPDS.TxtDLong = DG
FrmGCPDS.TxtMLong = MN
FrmGCPDS.TxtSLong = Round(SC, 5)

FrmGCPDS.TxtEllipsoidalH2 = Round(H2, 5)
End Sub



Public Sub WGS84_TO_PRS92()

Const PI = 3.14159265358979
Const O = PI / 180
Dim Ellipsoidal_Ht As Double 'Ellipsoidal Height

Dim dx, DY, DZ, s, RX, RY, RZ, a, A2, E2, E22, N1, X1, Y1, P1, L1, Z1, X2, Y2, Z2, R, H2, P2, P, N2, K, A1, DG, MN, MN1, SC, B2, L2




P1 = (FrmGCPDS.TxtDLat + (FrmGCPDS.TxtMLat / 60) + (FrmGCPDS.TxtSLat / 3600)) * PI / 180
L1 = (FrmGCPDS.TxtDLong + (FrmGCPDS.TxtMLong / 60) + (FrmGCPDS.TxtSLong / 3600)) * PI / 180
Ellipsoidal_Ht = FrmGCPDS.TxtEllipsoidalH2

'7 parameters
dx = 127.621531 'Translation X
DY = 67.243395  'Translation Y
DZ = 47.047384  'Translation Z
s = 1.06002112    'Scale Factor
RX = -3.06803751 / 3600 'Rotation X
RY = 4.90297653 / 3600 'Rotation Y
RZ = 1.57807293 / 3600 'Rotation Z



        
        a = 6378137
        E2 = 0.006694380003
        A2 = 6378206.4
        E22 = 0.006768657997
        
        N1 = a / Sqr(1 - E2 * (Sin(P1)) ^ 2)
        X1 = (N1 + Ellipsoidal_Ht) * Cos(P1) * Cos(L1)
        Y1 = (N1 + Ellipsoidal_Ht) * Cos(P1) * Sin(L1)
        Z1 = (N1 * (1 - E2) + Ellipsoidal_Ht) * Sin(P1)
        X2 = dx + (1 + s * 10 ^ -6) * (X1 + RZ * O * Y1 - RY * O * Z1)
        Y2 = DY + (1 + s * 10 ^ -6) * (-RZ * O * X1 + Y1 + RX * O * Z1)
        Z2 = DZ + (1 + s * 10 ^ -6) * (RY * O * X1 - RX * O * Y1 + Z1)
        R = Sqr(X2 ^ 2 + Y2 ^ 2)
        H2 = 0
        P2 = 0
        P = P2
        N2 = A2 / Sqr(1 - E22 * (Sin(P2)) ^ 2)
        P2 = Atn((N2 + H2) * Z2 / ((N2 * (1 - E22) + H2) * R))
        H2 = R / Cos(P2) - N2
        K = Abs(P - P2)
        
        Do Until K < 0.000000009999999
            P = P2
            N2 = A2 / Sqr(1 - E22 * (Sin(P2)) ^ 2)
            P2 = Atn((N2 + H2) * Z2 / ((N2 * (1 - E22) + H2) * R))
            H2 = R / Cos(P2) - N2
            K = Abs(P - P2)
        Loop
        
        A1 = P2 * 180 / PI
        DG = Int(A1)
        MN1 = (A1 - DG) * 60
        MN = Int(MN1)
        SC = (MN1 - MN) * 60
        
        FrmGCPDS.TxtLatD = DG
        FrmGCPDS.TxtLatM = MN
        FrmGCPDS.TxtLatS = Round(SC, 5)

        B2 = (Atn(Y2 / X2)) * 180 / PI
        L2 = B2 + 180
        A1 = L2
        
        
        DG = Int(A1)
        MN1 = (A1 - DG) * 60
        MN = Int(MN1)
        SC = (MN1 - MN) * 60
        
        FrmGCPDS.TxtLongD = DG
        FrmGCPDS.TxtLongM = MN
        FrmGCPDS.TxtLongS = Round(SC, 5)

        FrmGCPDS.txtEllipsoidalH = Round(H2, 5)
End Sub


Public Function IfDuplicate(PSGC As String) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "Select psgc_cd from psgc where psgc_cd='" & PSGC & "'", cnn, adOpenStatic
    If rst.RecordCount > 0 Then
        IfDuplicate = True
        Else
        IfDuplicate = False
        End If
End Function

Public Function If_Already_Exist(Station_Name As String) As Boolean
    Dim rst As New ADODB.Recordset
   
    rst.Open "Select stat_name from geoprov where stat_name='" & Station_Name & "'", cnn, adOpenStatic
    If rst.RecordCount > 0 Then
        If_Already_Exist = True
        Else
        If_Already_Exist = False
        End If
End Function


Public Function GetRegion(Province As String) As String
    Dim rst As New ADODB.Recordset
    Dim rst2 As New ADODB.Recordset
   
    rst.Open "Select psgc_cd from psgc where name='" & Trim(Province) & "'", cnn, adOpenStatic
    
    
    If rst.RecordCount > 0 Then
         rst2.Open "Select name from psgc where psgc_cd='" & Mid(rst("psgc_cd"), 1, 2) & "0000000" & "'", cnn, adOpenStatic
         GetRegion = rst2("name")
         Else
         GetRegion = ""
    End If
    
    
End Function

Public Sub PRS92_TO_WGS84_X(ND, NM, NS, ED, EM, es, E, ST)
Const PI = 3.14159265358979
Const O = PI / 180
Dim Ellipsoidal_Ht As Double

Dim dx, DY, DZ, s, RX, RY, RZ, a, A2, E2, E22, N1, X1, Y1, P1, L1, Z1, X2, Y2, Z2, R, H2, P2, P, N2, K, A1, DG, MN, MN1, SC, B2, L2 As Double

P1 = (ED + (EM / 60) + (es / 3600)) * PI / 180
L1 = (ND + (NM / 60) + (NS / 3600)) * PI / 180
Ellipsoidal_Ht = E

'7 parameters
dx = -127.62195 'Translation X
DY = -67.24478  'Translation Y
DZ = -47.04305  'Translation Z
s = -1.06002    'Scale Factor
RX = 3.06762 / 3600 'Rotation X
RY = -4.90291 / 3600 'Rotation Y
RZ = -1.5779 / 3600 'Rotation Z

a = 6378206.4
E2 = 0.006768658
A2 = 6378137#
E22 = 0.00669438

N1 = a / Sqr(1 - E2 * (Sin(P1)) ^ 2)
X1 = (N1 + Ellipsoidal_Ht) * Cos(P1) * Cos(L1)
Y1 = (N1 + Ellipsoidal_Ht) * Cos(P1) * Sin(L1)
Z1 = (N1 * (1 - E2) + Ellipsoidal_Ht) * Sin(P1)
X2 = dx + (1 + s * 10 ^ -6) * (X1 + RZ * O * Y1 - RY * O * Z1)
Y2 = DY + (1 + s * 10 ^ -6) * (-RZ * O * X1 + Y1 + RX * O * Z1)
Z2 = DZ + (1 + s * 10 ^ -6) * (RY * O * X1 - RX * O * Y1 + Z1)
R = Sqr(X2 ^ 2 + Y2 ^ 2)
H2 = 0
P2 = 0
P = P2
N2 = A2 / Sqr(1 - E22 * (Sin(P2)) ^ 2)
P2 = Atn((N2 + H2) * Z2 / ((N2 * (1 - E22) + H2) * R))
H2 = R / Cos(P2) - N2
K = Abs(P - P2)

Do Until K < 0.000000009999999
    P = P2
    N2 = A2 / Sqr(1 - E22 * (Sin(P2)) ^ 2)
    P2 = Atn((N2 + H2) * Z2 / ((N2 * (1 - E22) + H2) * R))
    H2 = R / Cos(P2) - N2
    K = Abs(P - P2)
Loop

A1 = P2 * 180 / PI
DG = Int(A1)
MN1 = (A1 - DG) * 60
MN = Int(MN1)
SC = (MN1 - MN) * 60


'FrmGCPDS.TxtDLat = DG
'FrmGCPDS.TxtMLat = MN
'FrmGCPDS.txtSLat = Round(SC, 5)
'
cnn.Execute "Update geoprov set wgs84ED=" & DG & " where stat_name='" & Replace(ST, "'", "''") & "'"
cnn.Execute "Update geoprov set wgs84EM=" & MN & " where stat_name='" & Replace(ST, "'", "''") & "'"
cnn.Execute "Update geoprov set wgs84ES=" & Round(SC, 5) & " where stat_name='" & Replace(ST, "'", "''") & "'"

'cnn.Execute "Update geoprov set wgs84ED=NULL where stat_name='" & Replace(ST, "'", "''") & "'"
'cnn.Execute "Update geoprov set wgs84EM=NULL where stat_name='" & Replace(ST, "'", "''") & "'"
'cnn.Execute "Update geoprov set wgs84ES= Null where stat_name='" & Replace(ST, "'", "''") & "'"

B2 = (Atn(Y2 / X2)) * 180 / PI
L2 = B2 + 180
A1 = L2


DG = Int(A1)
MN1 = (A1 - DG) * 60
MN = Int(MN1)
SC = (MN1 - MN) * 60

'FrmGCPDS.TxtDLong = DG
'FrmGCPDS.TxtMLong = MN
'FrmGCPDS.txtSLong = Round(SC, 5)
cnn.Execute "Update geoprov set wgs84nD=" & DG & " where stat_name='" & Replace(ST, "'", "''") & "'"
cnn.Execute "Update geoprov set wgs84nM=" & MN & " where stat_name='" & Replace(ST, "'", "''") & "'"
cnn.Execute "Update geoprov set wgs84nS=" & Round(SC, 5) & " where stat_name='" & Replace(ST, "'", "''") & "'"
cnn.Execute "Update geoprov set ellipz=" & Round(H2, 5) & " where stat_name='" & Replace(ST, "'", "''") & "'"

'cnn.Execute "Update geoprov set wgs84nD=NULL where stat_name='" & Replace(ST, "'", "''") & "'"
'cnn.Execute "Update geoprov set wgs84nM=NULL where stat_name='" & Replace(ST, "'", "''") & "'"
'cnn.Execute "Update geoprov set wgs84nS=NULL where stat_name='" & Replace(ST, "'", "''") & "'"
'cnn.Execute "Update geoprov set ellipz=NULL where stat_name='" & Replace(ST, "'", "''") & "'"
End Sub


Public Function DuplicateBenchmark(stationName As String, Province As String) As Boolean
    Dim rst As New ADODB.Recordset
     rst.Open "select * from benchmarks where stat_name='" & stationName & "' and Province='" & Province & "'", cnn, adOpenStatic
    
    If rst.RecordCount > 0 Then
       DuplicateBenchmark = True
       Else
       DuplicateBenchmark = False
    End If
    
End Function

Public Function IsDuplicateGravity(stationName As String)
    Dim rst As New ADODB.Recordset
     rst.Open "select * from gravity where stat_name='" & Replace(stationName, "'", "''") & "'", cnn, adOpenStatic
    
    If rst.RecordCount > 0 Then
       IsDuplicateGravity = True
       Else
       IsDuplicateGravity = False
    End If
    
End Function

Public Sub ExtractWGS84(Station As String, str As String)
    
    Dim WGS84, EllipsoidalH As String
    
    Dim Start, Start2, End1, End2 As Long
    Dim Values() As Double
    Dim EValue As Double
    Dim buf, buf2
    Dim i, x As Integer
    
    
    On Error GoTo Hell
   ' Debug.Print STR
    ReDim Values(0)
    
    Start = InStr(1, str, "WGS84", vbTextCompare)
    Start2 = InStr(1, str, "ELLIPSOIDAL", vbTextCompare)
    
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLPSOIDAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLIPSOPIDAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "EELIPSOIDAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLIPSOIDA", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLISPOIDAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLIPSOIADAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLIPSODIAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLIPSOODAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLISOIDAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLIPSOODAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    If Start2 = 0 Then
        Start2 = InStr(1, str, "ELLIPSOOIDAL", vbTextCompare) 'to compensate to the wrong spelling
    End If
    
    Start2 = Start2 - 1
    End1 = Start2 - Start
    
        If Start = 0 Then
            Exit Sub
        End If
        
        If Start2 <= 0 Then
            Exit Sub
        End If
        'MsgBox STR
    
   
            
            WGS84 = Mid(str, Start, End1)
            WGS84 = Replace(Replace(WGS84, "  ", " "), vbCrLf, " ")
            
             EllipsoidalH = Mid(str, Start2, Len(str) - Start2)
            'End2 = InStr(1, EllipsoidalH, "m", vbTextCompare)
            
            
            buf = Split(WGS84, " ")
            buf2 = Split(EllipsoidalH, " ")
    
                For i = 0 To UBound(buf)
                
                    If IsNumeric(buf(i)) Then
                        
                        If UBound(Values) > 0 Then
                            ReDim Preserve Values(1 To UBound(Values) + 1)
                                ElseIf UBound(Values) = 0 Then
                            ReDim Values(1 To 1)
                        End If
                        
                        Values(UBound(Values)) = buf(i)
                    End If
                Next
                
                For i = 0 To UBound(buf2)
                    If IsNumeric(buf2(i)) Then
                        EValue = buf2(i)
                        Exit For
                    End If
                Next
    
       
                   
                   cnn.Execute "update geoprov set wgs84ED=" & Values(1) & "," & "wgs84EM=" & Values(2) & "," & "wgs84ES=" & Values(3) & "," & "wgs84ND=" & Values(4) & "," & "wgs84NM=" & Values(5) & "," & "wgs84NS=" & Values(6) & "," & "Ellipz=" & EValue & " Where stat_name='" & Replace(Station, "'", "''") & "'"
                    
                  'FrmWGS84.List1.AddItem WGS84
                  'FrmWGS84.List1.ListIndex = FrmWGS84.List1.ListCount
                   
                    
                    'MsgBox Mid(EllipsoidalH, 19, End2 - 20)
                    Exit Sub
Hell:
End Sub

'Public Function LocateAPoint(ClickedPt As MapObjects.Point) As Long 'by point
'    Dim i As Long
'
'    Dim pnt As New MapObjects.Point
'
'
'    Dim dist() As Double
'    ReDim dist(0)
'    Dim evt As New MapObjects.GeoEvent
'
'
'
'
'
'    For i = 1 To FrmGCPDS.MyMap.TrackingLayer.EventCount
'        pnt.X = FrmGCPDS.MyMap.TrackingLayer.Event(i - 1).X
'        pnt.Y = FrmGCPDS.MyMap.TrackingLayer.Event(i - 1).Y
'
'        If UBound(dist) = 0 Then
'            ReDim dist(1 To UBound(dist) + 1)
'
'            Else
'        ReDim Preserve dist(1 To UBound(dist) + 1)
'         End If
'
'        dist(UBound(dist)) = ClickedPt.DistanceTo(pnt)
'     ' MsgBox gEventsTag(i) & " " & dist(UBound(dist))
'    Next
'
'    FrmGCPDS.StatusBar1.Panels(3).Text = gEventsTag((GetNearestPoint(dist)))
'
'   CurrentStationToIdentify = gEventsTag((GetNearestPoint(dist)))
'   FrmIdentify.Show
'
'
'
'
'
'End Function

Public Function FindGeoEvent(pt As MapObjects.Point) As Long

'Dim ptm As MapObjects
Dim tol As Double
Dim ellCircle As New MapObjects.Ellipse
Dim i As Long
Dim ptEvent As MapObjects.Point
Dim idx As Long

idx = -1


tol = FrmGCPDS.MyMap.ToMapDistance(10 * Screen.TwipsPerPixelX)

With ellCircle
  .Left = pt.x - tol
  .Right = pt.x + tol
  .Bottom = pt.y - tol
  .Top = pt.y + tol
End With

For i = 0 To FrmGCPDS.MyMap.TrackingLayer.EventCount - 1
  Set ptEvent = New MapObjects.Point
   ptEvent.x = FrmGCPDS.MyMap.TrackingLayer.Event(i).x
   ptEvent.y = FrmGCPDS.MyMap.TrackingLayer.Event(i).y
  
  If ellCircle.IsPointIn(ptEvent) Then
    idx = i
    Exit For
  End If
Next i


FindGeoEvent = idx

End Function


Public Sub BubbleSortNumbers(iarray As Variant)
  Dim lLoop1 As Long
   Dim lLoop2 As Long
  Dim lTemp As Double
 
 
  For lLoop1 = UBound(iarray) To LBound(iarray) Step -1
    For lLoop2 = LBound(iarray) + 1 To lLoop1
      If iarray(lLoop2 - 1) > iarray(lLoop2) Then
        lTemp = iarray(lLoop2 - 1)
       
        
        iarray(lLoop2 - 1) = iarray(lLoop2)
        
        
        iarray(lLoop2) = lTemp
       
      End If
    Next lLoop2
  Next lLoop1
  
End Sub


Public Function GetNearestPoint(iarray As Variant) As Long
    Dim i As Long
    Dim minvalue As Double
    Dim myindex As Long
        
        minvalue = iarray(1)
        For i = 1 To UBound(iarray)
                
            If iarray(i) <= minvalue Then
                minvalue = iarray(i)
                myindex = i
            End If
            
        Next i
        
        GetNearestPoint = myindex
End Function

Public Sub ZoomToSelected2(SearchString As String, FieldName As String)
    
    Dim selrecs As New MapObjects.Recordset
    Dim newrect As New MapObjects.Rectangle
    Dim curreccount As Long
    
    If FieldName = "Province" Then
        Set selrecs = FrmGCPDS.MyMap.Layers(0).SearchExpression("PROVINCE like '%" & UCase(SearchString) & "%'")
         MyLabelRender.Field = "PROVINCE"
    ElseIf FieldName = "Municipality" Then
        Set selrecs = FrmGCPDS.MyMap.Layers(0).SearchExpression("TOWN like '%" & StrConv(SearchString, vbProperCase) & "%'")
        MyLabelRender.Field = "TOWN"
    End If

 
        If Not selrecs.EOF Then
            
           
        
            selrecs.MoveFirst
            Set newrect = selrecs.Fields("shape").Value.Extent
        
                While Not selrecs.EOF
                    newrect.Union selrecs.Fields("shape").Value.Extent
                    selrecs.MoveNext
                Wend
        
        
            FrmGCPDS.MyMap.Extent = newrect
        
        End If
        
      


End Sub


Public Sub SearchMunicipality(SearchString As String, FieldName As String)
    Dim selrecs As MapObjects.Recordset
    Dim newrect As MapObjects.Rectangle
    Dim sym As New MapObjects.Symbol
 
    If UCase(Trim(FieldName)) = "PROVINCE" Then
        Set selrecs = FrmGCPDS.MyMap.Layers(0).SearchExpression("PROVINCE like '%" & UCase(SearchString) & "%'")
    ElseIf UCase(FieldName) = "MUNICIPALITY" Then
        Set selrecs = FrmGCPDS.MyMap.Layers(0).SearchExpression("TOWN like '%" & StrConv(SearchString, vbProperCase) & "%'")
    End If
    
    If Not selrecs.EOF Then
    
    sym.SymbolType = moFillSymbol
    sym.Style = moSolidFill
   
    sym.Color = &HB400&
    FrmGCPDS.MyMap.DrawShape selrecs, sym
    End If
End Sub

Public Function DMStoDD(DMS As String) As Double
        Dim buf
        buf = Split(DMS, " ")
        If UBound(buf) = 2 Then
            If IsNumeric(buf(0)) And IsNumeric(buf(1)) And IsNumeric(buf(2)) Then
                DMStoDD = buf(0) + (buf(1) / 60 + buf(2) / 3600)
            Else
                DMStoDD = -1
            End If
        Else
                DMStoDD = -1
        End If
        
End Function


Public Function DDtoDMS(dd As Double) As String
      'On Error GoTo hell
        Dim dx As Double
        Dim mm As Double
        Dim ss As Double
        
        dx = Int(dd)
        mm = (dd - dx) * 60
        ss = Format(((mm - Int(mm)) * 60), "0.####")
        
        DDtoDMS = dx & "° " & Int(mm) & "' " & ss & Chr(34)
        Exit Function
Hell:
        DDtoDMS = ""
End Function

Public Function DDtoDMS2(dd As Double) As String
        Dim dx As Double
        Dim mm As Double
        Dim ss As Double
        
        dx = Int(dd)
        
   
        mm = (dd - dx) * 60
        
        ss = Format(((mm - Int(mm)) * 60), "0.####")
        
        DDtoDMS2 = dx & " " & Int(mm) & " " & ss
        
End Function


Public Function isValidCoordinate(coordinate As String) As Boolean
    Dim buf
    buf = Split(Trim(coordinate), " ")
    
    If UBound(buf) = 2 Then
    
                If IsNumeric(buf(0)) And IsNumeric(buf(1)) And IsNumeric(buf(2)) Then
                        
                        If (buf(0) > 0 And buf(0) < 359) And (buf(1) < 60) And (buf(2) < 60) Then
                                
                                isValidCoordinate = True
                                
                                Else
                                
                                isValidCoordinate = False
                                
                        End If
                        
                Else
                     
                     isValidCoordinate = False
                     
                End If
    
    Else
        
        isValidCoordinate = False
        
    End If
End Function

Public Sub BuildExecuteQuery()
    Dim i As Integer
    Dim varlist2
    
    If RstQuery.State = 1 Then
    RstQuery.Close
   End If
    
    For i = 1 To FrmGCPDS.LstConditions.ListItems.Count
        strcondition = strcondition & FrmGCPDS.LstConditions.ListItems(i).SubItems(4) & FrmGCPDS.LstConditions.ListItems(i).SubItems(1) & FrmGCPDS.LstConditions.ListItems(i).SubItems(5) & FrmGCPDS.LstConditions.ListItems(i).SubItems(3)
     Next

    RstQuery.PageSize = 10
    RstQuery.CacheSize = 10
    
    RstQuery.CursorLocation = adUseClient
   
   '** edited by fat - jul14 2009
    If BMQuery = 0 Then
    
    'MsgBox "SELECT geoprov.*, order_lib.description as HOR_Order, horfixme.HDesc, marktype.MTDesc, markpur.MDesc, markstatus.MSDesc FROM ((((geoprov LEFT JOIN order_lib ON geoprov.h_order = order_lib.h_order) LEFT JOIN marktype ON geoprov.mark_type = marktype.MTCode) LEFT JOIN markstatus ON geoprov.mark_stat = markstatus.MSCode) LEFT JOIN horfixme ON geoprov.h_fix = horfixme.HCode) LEFT JOIN markpur ON geoprov.mark_const = markpur.MCode where " & strcondition & " AND h_ref='PRS92' order by stat_name"
        RstQuery.Open "SELECT geoprov.*, order_lib.description as HOR_Order, horfixme.HDesc, marktype.MTDesc, markpur.MDesc, markstatus.MSDesc FROM ((((geoprov LEFT JOIN order_lib ON geoprov.h_order = order_lib.h_order) LEFT JOIN marktype ON geoprov.mark_type = marktype.MTCode) LEFT JOIN markstatus ON geoprov.mark_stat = markstatus.MSCode) LEFT JOIN horfixme ON geoprov.h_fix = horfixme.HCode) LEFT JOIN markpur ON geoprov.mark_const = markpur.MCode where " & strcondition & " AND h_ref='PRS92' order by stat_name", cnn, adOpenStatic, adLockOptimistic
    Else
   ' MsgBox "SELECT benchmarks.*, order_lib.description as HOR_Order, horfixme.HDesc, marktype.MTDesc, markpur.MDesc, markstatus.MSDesc FROM ((((benchmarks LEFT JOIN order_lib ON benchmarks.h_order = order_lib.h_order) LEFT JOIN marktype ON benchmarks.mark_type = marktype.MTCode) LEFT JOIN markstatus ON benchmarks.mark_stat = markstatus.MSCode) LEFT JOIN horfixme ON benchmarks.h_fix = horfixme.HCode) LEFT JOIN markpur ON benchmarks.mark_const = markpur.MCode where " & strcondition & " order by stat_name"
        RstQuery.Open "SELECT benchmarks.*, order_lib.description as HOR_Order, horfixme.HDesc, marktype.MTDesc, markpur.MDesc, markstatus.MSDesc FROM ((((benchmarks LEFT JOIN order_lib ON benchmarks.h_order = order_lib.h_order) LEFT JOIN marktype ON benchmarks.mark_type = marktype.MTCode) LEFT JOIN markstatus ON benchmarks.mark_stat = markstatus.MSCode) LEFT JOIN horfixme ON benchmarks.h_fix = horfixme.HCode) LEFT JOIN markpur ON benchmarks.mark_const = markpur.MCode where " & strcondition & " order by stat_name", cnn, adOpenStatic, adLockOptimistic
        
    End If
    FrmGCPDS.ResultLabel = "Your query returns " & RstQuery.RecordCount & " result" & IIf(RstQuery.RecordCount > 1, "s.", ".")
   
    
    
    PageCounter = 1
    
    If RstQuery.PageCount > 1 Then
        FrmGCPDS.PageCounterLabel = "Page " & RstQuery.AbsolutePage & " of " & RstQuery.PageCount
        FrmGCPDS.RaveNextPage.Enabled = True
        FrmGCPDS.RavePreviousPage.Enabled = False
        FrmGCPDS.LstResult.ListItems.Clear
        For i = 1 To 10
            Set varlist2 = FrmGCPDS.LstResult.ListItems.Add
            varlist2.Text = RstQuery("Stat_name")
            If BMQuery <> 0 Then
                varlist2.SubItems(1) = RstQuery!ucode
            End If
            RstQuery.MoveNext
        Next
    ElseIf RstQuery.PageCount = 1 Then
        FrmGCPDS.PageCounterLabel = ""
        FrmGCPDS.RaveNextPage.Enabled = False
        FrmGCPDS.RavePreviousPage.Enabled = False
        FrmGCPDS.LstResult.ListItems.Clear
        
        For i = 1 To 10
            Set varlist2 = FrmGCPDS.LstResult.ListItems.Add
            varlist2.Text = RstQuery("Stat_name")
            If BMQuery <> 0 Then
                varlist2.SubItems(1) = RstQuery!ucode
            End If
            RstQuery.MoveNext
            
            If RstQuery.EOF Then
                Exit For
            End If
        Next
    ElseIf RstQuery.PageCount = 0 Then
        FrmGCPDS.PageCounterLabel = ""
        FrmGCPDS.RaveNextPage.Enabled = False
        FrmGCPDS.RavePreviousPage.Enabled = False
        FrmGCPDS.LstResult.ListItems.Clear
    End If
    
    'RstQuery.Close
End Sub


Public Function DoesColumnExist(tableName As String, columnName As String) As Boolean
        Dim rsttemp As New ADODB.Recordset
        rsttemp.Open "select column_name from INFORMATION_SCHEMA.columns where table_name = '" & tableName & "' and column_name='" & columnName & "'", cnn, adOpenStatic
        If rsttemp.RecordCount > 0 Then
            DoesColumnExist = True
        Else
            DoesColumnExist = False
        End If
        rsttemp.Close
End Function


Public Sub UpdateTables()
    
    If DoesColumnExist("benchmarks", "dateupdated") = False Then cnn.Execute "ALTER TABLE benchmarks ADD dateupdated datetime"
    
End Sub
