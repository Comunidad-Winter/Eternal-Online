Attribute VB_Name = "modScreenCapture"
Option Explicit
Private Enum IJLERR
  IJL_OK = 0
  IJL_INTERRUPT_OK = 1
  IJL_ROI_OK = 2
  IJL_EXCEPTION_DETECTED = -1
  IJL_INVALID_ENCODER = -2
  IJL_UNSUPPORTED_SUBSAMPLING = -3
  IJL_UNSUPPORTED_BYTES_PER_PIXEL = -4
  IJL_MEMORY_ERROR = -5
  IJL_BAD_HUFFMAN_TABLE = -6
  IJL_BAD_QUANT_TABLE = -7
  IJL_INVALID_JPEG_PROPERTIES = -8
  IJL_ERR_FILECLOSE = -9
  IJL_INVALID_FILENAME = -10
  IJL_ERROR_EOF = -11
  IJL_PROG_NOT_SUPPORTED = -12
  IJL_ERR_NOT_JPEG = -13
  IJL_ERR_COMP = -14
  IJL_ERR_SOF = -15
  IJL_ERR_DNL = -16
  IJL_ERR_NO_HUF = -17
  IJL_ERR_NO_QUAN = -18
  IJL_ERR_NO_FRAME = -19
  IJL_ERR_MULT_FRAME = -20
  IJL_ERR_DATA = -21
  IJL_ERR_NO_IMAGE = -22
  IJL_FILE_ERROR = -23
  IJL_INTERNAL_ERROR = -24
  IJL_BAD_RST_MARKER = -25
  IJL_THUMBNAIL_DIB_TOO_SMALL = -26
  IJL_THUMBNAIL_DIB_WRONG_COLOR = -27
  IJL_RESERVED = -99
End Enum

Private Enum IJLIOTYPE
  IJL_SETUP = -1&
  IJL_JFILE_READPARAMS = 0&
  IJL_JBUFF_READPARAMS = 1&
  IJL_JFILE_READWHOLEIMAGE = 2&
  IJL_JBUFF_READWHOLEIMAGE = 3&
  IJL_JFILE_READHEADER = 4&
  IJL_JBUFF_READHEADER = 5&
  IJL_JFILE_READENTROPY = 6&
  IJL_JBUFF_READENTROPY = 7&
  IJL_JFILE_WRITEWHOLEIMAGE = 8&
  IJL_JBUFF_WRITEWHOLEIMAGE = 9&
  IJL_JFILE_WRITEHEADER = 10&
  IJL_JBUFF_WRITEHEADER = 11&
  IJL_JFILE_WRITEENTROPY = 12&
  IJL_JBUFF_WRITEENTROPY = 13&
  IJL_JFILE_READONEHALF = 14&
  IJL_JBUFF_READONEHALF = 15&
  IJL_JFILE_READONEQUARTER = 16&
  IJL_JBUFF_READONEQUARTER = 17&
  IJL_JFILE_READONEEIGHTH = 18&
  IJL_JBUFF_READONEEIGHTH = 19&
  IJL_JFILE_READTHUMBNAIL = 20&
  IJL_JBUFF_READTHUMBNAIL = 21&
End Enum

Private Type JPEG_CORE_PROPERTIES_VB
  UseJPEGPROPERTIES As Long
  DIBBytes As Long
  DIBWidth As Long
  DIBHeight As Long
  DIBPadBytes As Long
  DIBChannels As Long
  DIBColor As Long
  DIBSubsampling As Long
  JPGFile As Long
  JPGBytes As Long
  JPGSizeBytes As Long
  JPGWidth As Long
  JPGHeight As Long
  JPGChannels As Long
  JPGColor As Long
  JPGSubsampling As Long
  JPGThumbWidth As Long
  JPGThumbHeight As Long
  cconversion_reqd As Long
  upsampling_reqd As Long
  jquality As Long
  jprops(0 To 19999) As Byte
End Type

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)
Private Declare Function ijlInit Lib "ijl11.dll" (jcprops As Any) As Long
Private Declare Function ijlFree Lib "ijl11.dll" (jcprops As Any) As Long
Private Declare Function ijlRead Lib "ijl11.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Declare Function ijlWrite Lib "ijl11.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Const MAX_PATH = 260

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Const OF_WRITE = &H1
Private Const OF_SHARE_DENY_WRITE = &H20
Private Const INVALID_HANDLE As Long = -1
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Function LoadJPG( _
      ByRef cDib As cDIBSection, _
      ByVal sFile As String _
   ) As Boolean
Dim tJ As JPEG_CORE_PROPERTIES_VB
Dim bFile() As Byte
Dim lR As Long
Dim lPtr As Long
Dim lJPGWidth As Long, lJPGHeight As Long

   lR = ijlInit(tJ)
   If lR = IJL_OK Then
      
      ' Write the filename to the jcprops.JPGFile member:
      bFile = StrConv(sFile, vbFromUnicode)
      ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
      bFile(UBound(bFile)) = 0
      lPtr = VarPtr(bFile(0))
      CopyMemory tJ.JPGFile, lPtr, 4
      
      ' Read the JPEG file parameters:
      lR = ijlRead(tJ, IJL_JFILE_READPARAMS)
      If lR <> IJL_OK Then
         ' Throw error
         MsgBox "Failed to read JPG", vbExclamation
      Else
        ' set JPG color
         If tJ.JPGChannels = 1 Then
            tJ.JPGColor = 4& ' IJL_G
         Else
            tJ.JPGColor = 3& ' IJL_YCBCR
         End If
            
         ' Get the JPGWidth ...
         lJPGWidth = tJ.JPGWidth
         ' .. & JPGHeight member values:
         lJPGHeight = tJ.JPGHeight
      
         ' Create a buffer of sufficient size to hold the image:
         If cDib.Create(lJPGWidth, lJPGHeight) Then
            ' Store DIBWidth:
            tJ.DIBWidth = lJPGWidth
            ' Very important: tell IJL how many bytes extra there
            ' are on each DIB scan line to pad to 32 bit boundaries:
            tJ.DIBPadBytes = cDib.BytesPerScanLine - lJPGWidth * 3
            ' Store DIBHeight:
            tJ.DIBHeight = -lJPGHeight
            ' Store Channels:
            tJ.DIBChannels = 3&
            ' Store DIBBytes (pointer to uncompressed JPG data):
            tJ.DIBBytes = cDib.DIBSectionBitsPtr
            
            ' Now decompress the JPG into the DIBSection:
            lR = ijlRead(tJ, IJL_JFILE_READWHOLEIMAGE)
            If lR = IJL_OK Then
               ' That's it!  cDib now contains the uncompressed JPG.
               LoadJPG = True
            Else
               ' Throw error:
               MsgBox "Cannot read Image Data from file.", vbExclamation
            End If
         Else
            ' failed to create the DIB...
         End If
      End If
                        
      ' Ensure we have freed memory:
      ijlFree tJ
   Else
      ' Throw error:
      MsgBox "Failed to initialise the IJL library: " & lR, vbExclamation
   End If
   
   
End Function
Public Function LoadJPGFromPtr( _
      ByRef cDib As cDIBSection, _
      ByVal lPtr As Long, _
      ByVal lSize As Long _
   ) As Boolean
Dim tJ As JPEG_CORE_PROPERTIES_VB
Dim lR As Long
Dim lJPGWidth As Long, lJPGHeight As Long

   lR = ijlInit(tJ)
   If lR = IJL_OK Then
            
      ' set JPEG buffer
      tJ.JPGBytes = lPtr
      tJ.JPGSizeBytes = lSize
            
      ' Read the JPEG parameters:
      lR = ijlRead(tJ, IJL_JBUFF_READPARAMS)
      If lR <> IJL_OK Then
         ' Throw error
         MsgBox "Failed to read JPG", vbExclamation
      Else
        ' set JPG color
         If tJ.JPGChannels = 1 Then
            tJ.JPGColor = 4& ' IJL_G
         Else
            tJ.JPGColor = 3& ' IJL_YCBCR
         End If
      
         ' Get the JPGWidth ...
         lJPGWidth = tJ.JPGWidth
         ' .. & JPGHeight member values:
         lJPGHeight = tJ.JPGHeight
      
         ' Create a buffer of sufficient size to hold the image:
         If cDib.Create(lJPGWidth, lJPGHeight) Then
            ' Store DIBWidth:
            tJ.DIBWidth = lJPGWidth
            ' Very important: tell IJL how many bytes extra there
            ' are on each DIB scan line to pad to 32 bit boundaries:
            tJ.DIBPadBytes = cDib.BytesPerScanLine - lJPGWidth * 3
            ' Store DIBHeight:
            tJ.DIBHeight = -lJPGHeight
            ' Store Channels:
            tJ.DIBChannels = 3&
            ' Store DIBBytes (pointer to uncompressed JPG data):
            tJ.DIBBytes = cDib.DIBSectionBitsPtr
            
            ' Now decompress the JPG into the DIBSection:
            lR = ijlRead(tJ, IJL_JBUFF_READWHOLEIMAGE)
            If lR = IJL_OK Then
               ' That's it!  cDib now contains the uncompressed JPG.
               LoadJPGFromPtr = True
            Else
               ' Throw error:
               MsgBox "Cannot read Image Data from file.", vbExclamation
            End If
         Else
            ' failed to create the DIB...
         End If
      End If
                        
      ' Ensure we have freed memory:
      ijlFree tJ
   Else
      ' Throw error:
      MsgBox "Failed to initialise the IJL library: " & lR, vbExclamation
   End If
   
End Function

Public Function SaveJPG( _
      ByRef cDib As cDIBSection, _
      ByVal sFile As String, _
      Optional ByVal lQuality As Long = 90 _
   ) As Boolean
Dim tJ As JPEG_CORE_PROPERTIES_VB
Dim bFile() As Byte
Dim lPtr As Long
Dim lR As Long
Dim tFnd As WIN32_FIND_DATA
Dim hFile As Long
Dim bFileExisted As Boolean
Dim lFileSize As Long
   
   hFile = -1
   
   lR = ijlInit(tJ)
   If lR = IJL_OK Then
      
      ' Check if we're attempting to overwrite an existing file.
      ' If so hFile <> INVALID_FILE_HANDLE:
      bFileExisted = (FindFirstFile(sFile, tFnd) <> -1)
      If bFileExisted Then
         Kill sFile
      End If
      
      ' Set up the DIB information:
      ' Store DIBWidth:
      tJ.DIBWidth = cDib.Width
      ' Store DIBHeight:
      tJ.DIBHeight = -cDib.Height
      ' Store DIBBytes (pointer to uncompressed JPG data):
      tJ.DIBBytes = cDib.DIBSectionBitsPtr
      ' Very important: tell IJL how many bytes extra there
      ' are on each DIB scan line to pad to 32 bit boundaries:
      tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3
      
      ' Set up the JPEG information:
      
      ' Store JPGFile:
      bFile = StrConv(sFile, vbFromUnicode)
      ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
      bFile(UBound(bFile)) = 0
      lPtr = VarPtr(bFile(0))
      CopyMemory tJ.JPGFile, lPtr, 4
      ' Store JPGWidth:
      tJ.JPGWidth = cDib.Width
      ' .. & JPGHeight member values:
      tJ.JPGHeight = cDib.Height
      ' Set the quality/compression to save:
      tJ.jquality = lQuality
            
      ' Write the image:
      lR = ijlWrite(tJ, IJL_JFILE_WRITEWHOLEIMAGE)
      
      ' Check for success:
      If lR = IJL_OK Then
      
         ' Now if we are replacing an existing file, then we want to
         ' put the file creation and archive information back again:
         If bFileExisted Then
            
            hFile = lopen(sFile, OF_WRITE Or OF_SHARE_DENY_WRITE)
            If hFile = 0 Then
               ' problem
            Else
               SetFileTime hFile, tFnd.ftCreationTime, tFnd.ftLastAccessTime, tFnd.ftLastWriteTime
               lclose hFile
               SetFileAttributes sFile, tFnd.dwFileAttributes
            End If
            
         End If
         
         lFileSize = tJ.JPGSizeBytes - tJ.JPGBytes
         
         ' Success:
         SaveJPG = True
         
      Else
         ' Throw error
         Err.Raise 26001, "No se pudo Guarrdar el JPG" & lR, vbExclamation
      End If
      
      ' Ensure we have freed memory:
      ijlFree tJ
   Else
      ' Throw error:
      Err.Raise 26001, App.EXEName & ".mIntelJPEGLibrary", "No se pudo inicializar la Libreria " & lR
   End If
   

End Function

Public Function SaveJPGToPtr( _
      ByRef cDib As cDIBSection, _
      ByVal lPtr As Long, _
      ByRef lBufSize As Long, _
      Optional ByVal lQuality As Long = 90 _
   ) As Boolean
Dim tJ As JPEG_CORE_PROPERTIES_VB
Dim lR As Long
Dim hFile As Long
   
   hFile = -1
   
   lR = ijlInit(tJ)
   If lR = IJL_OK Then
      
      ' Set up the DIB information:
      ' Store DIBWidth:
      tJ.DIBWidth = cDib.Width
      ' Store DIBHeight:
      tJ.DIBHeight = -cDib.Height
      ' Store DIBBytes (pointer to uncompressed JPG data):
      tJ.DIBBytes = cDib.DIBSectionBitsPtr
      ' Very important: tell IJL how many bytes extra there
      ' are on each DIB scan line to pad to 32 bit boundaries:
      tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3
      
      ' Set up the JPEG information:
      ' Store JPGWidth:
      tJ.JPGWidth = cDib.Width
      ' .. & JPGHeight member values:
      tJ.JPGHeight = cDib.Height
      ' Set the quality/compression to save:
      tJ.jquality = lQuality
      ' set JPEG buffer
      tJ.JPGBytes = lPtr
      tJ.JPGSizeBytes = lBufSize
            
      ' Write the image:
      lR = ijlWrite(tJ, IJL_JBUFF_WRITEWHOLEIMAGE)
            
      ' Check for success:
      If lR = IJL_OK Then
         
         lBufSize = tJ.JPGSizeBytes
         
         ' Success:
         SaveJPGToPtr = True
         
      Else
         ' Throw error
         Err.Raise 26001, App.EXEName & ".mIntelJPEGLibrary", "Failed to save to JPG " & lR, vbExclamation
      End If
      
      ' Ensure we have freed memory:
      ijlFree tJ
   Else
      ' Throw error:
      Err.Raise 26001, App.EXEName & ".mIntelJPEGLibrary", "Failed to initialise the IJL library: " & lR
   End If
   

End Function

Public Sub ScreenCapture(Optional ByVal Autofragshooter As Boolean = False)
'Medio desprolijo donde pongo la pic, pero es lo que hay por ahora
On Error GoTo Err:
    Dim hwnd As Long
    Dim file As String
    Dim sI As String
    Dim c As New cDIBSection
    Dim i As Long
    Dim hdcc As Long
    
    Dim dirFile As String
    
    hdcc = GetDC(frmMain.hwnd)
    
    frmScreenshots.Picture1.AutoRedraw = True
    frmScreenshots.Picture1.Width = frmScaleWidth
    frmScreenshots.Picture1.Height = frmScaleHeight

    Call BitBlt(frmScreenshots.Picture1.hdc, 0, 0, Config_Inicio.ResolutionX, Config_Inicio.ResolutionY, hdcc, 0, 0, SRCCOPY)
    Call ReleaseDC(frmMain.hwnd, hdcc)
    
    hdcc = INVALID_HANDLE
    
    dirFile = IIf(Autofragshooter, "\Screenshots\FragShooter", "\Screenshots")
    
    If Not FileExist(App.Path & dirFile, vbDirectory) Then MkDir (App.Path & dirFile)
    
    file = App.Path & dirFile & "\" & Format(Now, "DD-MM-YYYY hh-mm-ss") & ".jpg"
    
    frmScreenshots.Picture1.Refresh
    frmScreenshots.Picture1.Picture = frmScreenshots.Picture1.Image
    
    c.CreateFromPicture frmScreenshots.Picture1.Picture
    
    SaveJPG c, file
    
    AddtoRichTextBox frmMain.RecTxt, "Screen Capturada!", 200, 200, 200, False, False, True
Exit Sub

Err:
    Call AddtoRichTextBox(frmMain.RecTxt, Err.number & "-" & Err.Description, 200, 200, 200, False, False, True)
    
    If hdcc <> INVALID_HANDLE Then _
        Call ReleaseDC(frmMain.hwnd, hdcc)
End Sub

Public Function FullScreenCapture(ByVal file As String) As Boolean
'Medio desprolijo donde pongo la pic, pero es lo que hay por ahora
    Dim c As New cDIBSection
    Dim hdcc As Long
    Dim Handle As Long
    
    hdcc = GetDC(Handle)
    
    frmScreenshots.Picture1.AutoRedraw = True
    
    If NoRes Then
        frmScreenshots.Picture1.Width = Screen.Width
        frmScreenshots.Picture1.Height = Screen.Height
        
        Call BitBlt(frmScreenshots.Picture1.hdc, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, hdcc, 0, 0, SRCCOPY)
    Else
        frmScreenshots.Picture1.Width = frmScaleWidth
        frmScreenshots.Picture1.Height = frmScaleHeight
        
        Call BitBlt(frmScreenshots.Picture1.hdc, 0, 0, Config_Inicio.ResolutionX, Config_Inicio.ResolutionY, hdcc, 0, 0, SRCCOPY)
    End If
    
    Call ReleaseDC(Handle, hdcc)
    
    hdcc = INVALID_HANDLE
    
    If Not FileExist(App.Path & "\TEMP", vbDirectory) Then MkDir (App.Path & "\TEMP")
    
    frmScreenshots.Picture1.Refresh
    frmScreenshots.Picture1.Picture = frmScreenshots.Picture1.Image
    
    c.CreateFromPicture frmScreenshots.Picture1.Picture
    
    SaveJPG c, file
    
    FullScreenCapture = True
End Function
