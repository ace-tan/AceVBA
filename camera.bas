Attribute VB_Name = "camera"
'camera tested 2/1/16
'folder : Camera
'form : frmCamera

Option Compare Database
Option Explicit

Const WM_CAP As Long = &H400
Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP + 10
Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Const WM_CAP_EDIT_COPY As Long = WM_CAP + 30
Const WS_CHILD As Long = &H40000000
Const WS_VISIBLE As Long = &H10000000
Const WM_CAP_SET_PREVIEW As Long = WM_CAP + 50
Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Const WM_CAP_SET_SCALE As Long = WM_CAP + 53
Const WM_CAP_FILE_SAVEDIB = WM_CAP + 25
Const WM_DESTROY As Long = &H2
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Long = 1
Const SWP_NOZORDER As Long = &H4
Const SWP_SHOWWINDOW As Long = &H40
Const HWND_BOTTOM As Long = 1
Const HWND_TOPMOST As Long = -1

Dim hwnd As Long
Dim iDevice As Long
Dim click1 As Boolean

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
        
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Boolean
    
Declare Function capCreateCaptureWindowA Lib "avicap32.dll" _
    (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Integer, ByVal hWndParent As Long, ByVal nID As Long) As Long

Declare Function capGetDriverDescriptionA Lib "avicap32.dll" _
    (ByVal wDriver As Integer, ByVal lpszName As String, ByVal cbName As Long, _
    ByVal lpszVer As String, ByVal cbVer As Long) As Boolean

Declare Function GetDesktopWindow Lib "user32" () As Long


'The API format types we're interested in
Const CF_BITMAP = 2
Const CF_PALETTE = 9
Const CF_ENHMETAFILE = 14
Const IMAGE_BITMAP = 0
Const LR_COPYRETURNORG = &H4
' Addded by SL Apr/2000
Const xlPicture = CF_BITMAP
Const xlBitmap = CF_BITMAP

Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) _
   As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function GetClipboardData Lib "user32" (ByVal wFormat As _
   Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal _
   dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
   ByVal lpString2 As Any) As Long

'Does the clipboard contain a bitmap/metafile?
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long

'Create our own copy of the metafile, so it doesn't get wiped out bysubsequent clipboard updates.
Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long

'Create our own copy of the bitmap, so it doesn't get wiped out bysubsequent clipboard updates.
Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
 
Type picImage
    Height As Long
    width As Long
End Type


Sub OffCamera()
'1. stop the camera device
Call StopCamera
End Sub

Sub StopCamera()
'1. stop the camera device
click1 = False
iDevice = 0
SendMessage hwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0
DestroyWindow hwnd
End Sub

Sub Save(fileName As String)
'1. save the file

If click1 <> False Then
    MsgBox "Picture saved", vbInformation, "Information"
'Call ScreenGrabToBMP
    SendMessage hwnd, WM_CAP_FILE_SAVEDIB, 0, ByVal CStr(fileName)
End If
End Sub

Function formOnLoad() As Collection
'1. set the device name
'resource from http://www.glengilchrist.co.uk/inserting-a-webcam-picture-into-excel/
click1 = False
Dim strName As String
Dim strVer As String
Dim bReturn As Boolean
Dim data As Collection
Set data = New Collection
iDevice = 0
strName = Space(100)
strVer = Space(100)

Do
    bReturn = capGetDriverDescriptionA(iDevice, strName, 100, strVer, 100)
    If bReturn = 1 Then
        data.Add Trim(strName)
        'Debug.Print Trim(strName)
    End If
    iDevice = iDevice + 1
            
Loop Until bReturn = False

Set formOnLoad = data

End Function

Sub OnCamera()
'1. set the camera screen position
Dim cameraWidth As Long
Dim cameraHeight As Long
Dim cameraPosX As Long
Dim cameraPosY As Long

cameraWidth = 300
cameraHeight = 300
cameraPosX = 30
cameraPosY = 440

Dim pic As picImage
pic.Height = 96
pic.width = 96

If click1 <> True Then
    iDevice = 0
        
    hwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640, _
            480, GetDesktopWindow(), 0)

    'hwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640, _
   '         480, pic, 0)
    SendMessage hwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0
    SendMessage hwnd, WM_CAP_SET_SCALE, True, 0.15
    SendMessage hwnd, WM_CAP_SET_PREVIEWRATE, 66, 0
    SendMessage hwnd, WM_CAP_SET_PREVIEW, True, 0
    SetWindowPos hwnd, HWND_TOPMOST, cameraPosX, cameraPosY, cameraWidth, cameraHeight, SWP_SHOWWINDOW
End If

click1 = True
End Sub


 
 Function GetClipBoard() As Long
 
   SendMessage hwnd, WM_CAP_EDIT_COPY, 0, 0
' Adapted from original Source Code by:
'* MODULE NAME:     Paste Picture
'* AUTHOR & DATE:   STEPHEN BULLEN, Business Modelling Solutions Ltd.
'*                  15 November 1998
'*
'* CONTACT:         Ste...@BMSLtd.co.uk
'* WEB SITE:        http://www.BMSLtd.co.uk

' Handles for graphic Objects
Dim hClipBoard As Long
Dim hBitmap As Long
Dim hBitmap2 As Long

'Check if the clipboard contains the required format
'hPicAvail = IsClipboardFormatAvailable(lPicType)

 ' Open the ClipBoard
 hClipBoard = OpenClipboard(0&)

 If hClipBoard <> 0 Then
    ' Get a handle to the Bitmap
    hBitmap = GetClipboardData(CF_BITMAP)

    If hBitmap = 0 Then GoTo exit_error
    ' Create our own copy of the image on the clipboard, in theappropriate format.
    'If lPicType = CF_BITMAP Then
        hBitmap2 = CopyImage(hBitmap, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
     '   Else
      '  hBitmap2 = CopyEnhMetaFile(hBitmap, vbNullString)
       ' End If

        'Release the clipboard to other programs
        hClipBoard = CloseClipboard

 GetClipBoard = hBitmap2
 Exit Function

 End If


exit_error:
' Return False
GetClipBoard = -1
End Function


Public Sub ScreenGrabToBMP()
Dim sDir As String
sDir = CurrentProject.path & "\Logo"
Const sFilename As String = "screen.bmp"


Dim lngRet As Long
Dim lngBytes As Long
Dim hPix As IPicture
Dim hBitmap As Long



'If bFullScreen Then
'  PrintScreen
'Else
 ' AltPrintScreen
'End If

hBitmap = GetClipBoard
Debug.Print hBitmap
'Set hPix = BitmapToPicture(hBitmap)
'SavePicture hPix, sDir & sFilename



Set hPix = Nothing
End Sub



