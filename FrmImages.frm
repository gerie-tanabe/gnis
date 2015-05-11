VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmImages 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Images"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   7350
   Icon            =   "FrmImages.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   586
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   8265
      Left            =   90
      ScaleHeight     =   549
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   0
      Top             =   90
      Width           =   7170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6750
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "File"
      Begin VB.Menu MnuLoad 
         Caption         =   "Load Picture"
      End
      Begin VB.Menu MnuClear 
         Caption         =   "Clear Picture"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentPicture As String
Dim HasPicture As Boolean


Private Sub Form_Load()

 CurrentPicture = ""
Dim rst As New ADODB.Recordset
rst.Open "select image from geoprov where stat_name='" & Replace(rstRecords("stat_name"), "'", "''") & "'", cnn, adOpenStatic, adLockOptimistic

If IsNull(rst!Image) = False Then
    LoadImages "image", Me.PictureBox, rst
    HasPicture = True
    Else
    Me.PictureBox.Cls
    HasPicture = False
End If

      
SetMenu

End Sub


Public Sub LoadImages(Field As String, MyImage As PictureBox, rst As ADODB.Recordset)


Dim mystream As New ADODB.Stream
            mystream.Type = adTypeBinary
            mystream.Open
            mystream.Write rst(Field)
            mystream.SaveToFile App.Path & "\temp.bin", adSaveCreateOverWrite
            
                FitPictureToBox App.Path & "\temp.bin", PictureBox
                CurrentPicture = App.Path & "\temp.bin"
            
            'Cleaning
            
            mystream.Close
            Set mystream = Nothing
            
End Sub

Private Sub Form_Resize()

On Error GoTo hell
    DoEvents
    Me.PictureBox.Top = 5
    Me.PictureBox.Left = 5
    Me.PictureBox.Height = ScaleHeight - 10
    Me.PictureBox.Width = ScaleWidth - 10
    Me.PictureBox.Picture = LoadPicture()
    FitPictureToBox CurrentPicture, Me.PictureBox
    Exit Sub
hell:

End Sub









Private Sub MnuClear_Click()
ClearPicture
HasPicture = False
SetMenu
End Sub

Private Sub MnuCopy_Click()
Clipboard.Clear
Clipboard.SetData PictureBox.Image, vbCFBitmap
SetMenu
End Sub

Private Sub MnuExit_Click()
Unload Me
End Sub

Private Sub MnuLoad_Click()
LoadPix
SetMenu
End Sub

Private Sub MnuPaste_Click()
            Me.PictureBox.Picture = LoadPicture()
          If Clipboard.GetFormat(2) Then
            Me.PictureBox.Picture = Clipboard.GetData(vbCFBitmap)
          End If
          
          If Clipboard.GetFormat(3) Then
            Me.PictureBox.Picture = Clipboard.GetData(vbCFMetafile)
          End If
          LoadPixFromPicturebox
          
          HasPicture = True
          SetMenu
End Sub

Private Sub PictureBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim pt As PointAPI
    Dim ret As BTN_STYLE

    'DoEvents
    If Button = vbRightButton Then
    
        hMenu = CreatePopupMenu()
      
      AppendMenu hMenu, MF_STRING, 1, "Load Picture"
      
      
      
      
     If HasPicture = False Then
     AppendMenu hMenu, MF_Grayed, 2, "Clear Picture"
      AppendMenu hMenu, MF_Grayed, 3, "Copy"
     Else
      AppendMenu hMenu, MF_STRING, 2, "Clear Picture"
      AppendMenu hMenu, MF_STRING, 3, "Copy"
     End If
      
      
     
      
      
         If Clipboard.GetFormat(2) Or Clipboard.GetFormat(3) Then
            AppendMenu hMenu, MF_STRING, 4, "Paste"
            Else
            AppendMenu hMenu, MF_Grayed, 4, "Paste"
         End If
        
        
        

        GetCursorPos pt

        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                pt.x, pt.y, Me.hWnd, ByVal 0&)
        DestroyMenu hMenu

        If ret = 1 Then
            
           LoadPix
        ElseIf ret = 2 Then
         ClearPicture
         HasPicture = False
        ElseIf ret = 3 Then
        Clipboard.Clear
         Clipboard.SetData PictureBox.Image, vbCFBitmap
        ElseIf ret = 4 Then
          
          Me.PictureBox.Picture = LoadPicture()
          If Clipboard.GetFormat(2) Then
            Me.PictureBox.Picture = Clipboard.GetData(vbCFBitmap)
          End If
          
          If Clipboard.GetFormat(3) Then
            Me.PictureBox.Picture = Clipboard.GetData(vbCFMetafile)
          End If
          HasPicture = True
          LoadPixFromPicturebox
        
        
        End If
    End If
End Sub

Private Sub ClearPicture()
    Me.PictureBox.Picture = LoadPicture()
        cnn.Execute "Update geoprov set Image=NULL where stat_name='" & rstRecords("stat_name") & "'"
    CurrentPicture = ""
    
End Sub

Private Sub RaveIns_Click()

End Sub

Private Sub LoadPixFromPicturebox()
    
        
     Dim rst As New ADODB.Recordset
        rst.CursorLocation = adUseClient
    rst.Open "select image from geoprov where stat_name='" & Replace(rstRecords("stat_name"), "'", "''") & "'", cnn, adOpenStatic, adLockOptimistic
  
        
    Dim mystream As New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    
   SavePicture Me.PictureBox.Image, App.Path & "\temp.bin"
   mystream.LoadFromFile App.Path & "\temp.bin"
   
   
   
   
    
   
    rst("image").Value = mystream.Read
    rst.Update
CurrentPicture = App.Path & "\temp.bin"
    
    
   
  
    
    
    'Cleaning
    mystream.Close
    rst.Close
    Set mystream = Nothing
    Set rst = Nothing
    
    
    


End Sub





Private Sub LoadPix()
    'On Error GoTo hell
    
    Dim rst As New ADODB.Recordset
    rst.Open "select image from geoprov where stat_name='" & Replace(rstRecords("stat_name"), "'", "''") & "'", cnn, adOpenStatic, adLockOptimistic
    
     Me.CommonDialog1.Filter = "JPEG *.jpg| *.jpg|GIF *.gif|*.gif|Bitmap *.bmp|*.bmp"
        Me.CommonDialog1.InitDir = App.Path
        Me.CommonDialog1.ShowOpen
        
    Dim mystream As New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile Me.CommonDialog1.filename
    
    
    rst("image").Value = mystream.Read
    
    If mystream.SIZE > 0 Then
        rst.Update
    End If

    
    
    FitPictureToBox Me.CommonDialog1.filename, Me.PictureBox
    HasPicture = True
    CurrentPicture = Me.CommonDialog1.filename
    
    
    'Cleaning
    mystream.Close
    rst.Close
    Set mystream = Nothing
    Set rst = Nothing
    
    
    
    Exit Sub
hell:
 
If Err.Number = 32755 Then
    
   Else
   
   MsgBox "Error Loading Image", vbCritical, "GNIS"
End If

End Sub


Private Sub FitPictureToBox(SourcePicture As String, MyPictureBox As PictureBox)

'<Aspect Ratio>
Dim aspect_src As Single
Dim wid As Single
Dim hgt As Single
'</Aspect Ratio>


Dim imgWidth As Long
Dim imgHeight As Long

Dim stdPic As StdPicture
Dim stdHDC As Long
Dim OrigHandle As Long
Dim DT_hwnd As Long
Dim DC As Long



Set stdPic = LoadPicture(SourcePicture)

imgWidth = Round(Me.ScaleX(stdPic.Width, vbHimetric, vbPixels))
imgHeight = Round(Me.ScaleY(stdPic.Height, vbHimetric, vbPixels))
   

DT_hwnd = GetDesktopWindow()
DC = GetDC(DT_hwnd)
stdHDC = CreateCompatibleDC(DC)
ReleaseDC DT_hwnd, DC
OrigHandle = SelectObject(stdHDC, stdPic.Handle)


    
    aspect_src = imgWidth / imgHeight
    
    wid = Me.PictureBox.ScaleWidth
    hgt = Me.PictureBox.ScaleHeight
    
    If wid / hgt > aspect_src Then
        wid = aspect_src * hgt
    Else
        hgt = wid / aspect_src
    End If
    
    MyPictureBox.Cls
    
SetStretchBltMode Me.PictureBox.hDC, HALFTONE
StretchBlt Me.PictureBox.hDC, (Me.PictureBox.ScaleWidth - wid) / 2, (Me.PictureBox.ScaleHeight - hgt) / 2, wid, hgt, stdHDC, 0, 0, imgWidth, imgHeight, vbSrcCopy
SelectObject stdHDC, OrigHandle
DeleteDC stdHDC


Set stdPic = Nothing
   
End Sub

Public Sub SetMenu()
     If HasPicture = False Then
        Me.MnuClear.Enabled = False
        Me.MnuCopy.Enabled = False
     Else
      Me.MnuClear.Enabled = True
        Me.MnuCopy.Enabled = True
     End If
      
      
     
      
      
         If Clipboard.GetFormat(2) Or Clipboard.GetFormat(3) Then
            Me.MnuPaste.Enabled = True
       
            Else
            Me.MnuPaste.Enabled = False
         End If
End Sub

