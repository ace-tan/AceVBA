Attribute VB_Name = "SystemTool"
Option Compare Database

Private Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" ( _
        ByVal lpBuffer As String, _
        ByRef nSize As Long) As Long

Public Property Get ComputerName() As String

  Dim stBuff As String * 255, lAPIResult As Long
  Dim lBuffLen As Long
  
  lBuffLen = 255
  
  lAPIResult = GetComputerName(stBuff, lBuffLen)
  
  If lBuffLen > 0 Then ComputerName = Left(stBuff, lBuffLen)

End Property

Public Function CompName() As String
  CompName = ComputerName
End Function

Sub Main()
'Website to refer : https://awrcorp.com/download/faq/english/scripts/basic_file_opperation.aspx
    Dim fso As Scripting.FileSystemObject
    Dim dr As Scripting.Drive

    Set fso = New Scripting.FileSystemObject

    ' Needed for not ready drives.
    On Error Resume Next

    'Debug.Clear
    For Each dr In fso.Drives
        Debug.Print dr.DriveLetter & ": File System " & dr.FileSystem
        Debug.Print dr.DriveLetter & ": Free Space " & CInt(dr.AvailableSpace / 10 ^ 9) & " GB, " & CInt(dr.AvailableSpace / dr.TotalSize * 100) & "%"
        Debug.Print dr.DriveLetter & ": Used Space " & CInt((dr.TotalSize - dr.AvailableSpace) / 10 ^ 9) & " GB, " & CInt((dr.TotalSize - dr.AvailableSpace) / dr.TotalSize * 100) & "%"
        Debug.Print dr.DriveLetter & ": Total Size " & CInt((dr.TotalSize / 10 ^ 9)) & " GB"
        
    Next dr
End Sub

Function GetIPAddress()
    Const strComputer As String = "."   ' Computer name. Dot means local computer
    Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
    Dim strIPAddress As String

    ' Connect to the WMI service
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    ' Get all TCP/IP-enabled network adapters
    Set IPConfigSet = objWMIService.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    ' Get all IP addresses associated with these adapters
    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        
        If Not IsNull(IPAddress) Then
        
            strIPAddress = strIPAddress & Join(IPAddress, ", ")
            
        End If
    Next

    GetIPAddress = strIPAddress
End Function

Function GetMACAddress()
' get a list of enabled adaptor names and MAC addresses
' from msdn.microsoft.com/en-us/library/windows/desktop/aa394217(v=vs.85).aspx
'note cmd getmac
Dim objVMI As Object
Dim vAdptr As Variant
Dim objAdptr As Object
Dim macLiscense As String
Dim macOriginal As String
'set the macLisence
macLiscense = ""

Set objVMI = GetObject("winmgmts:\\" & "." & "\root\cimv2")
Set vAdptr = objVMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each objAdptr In vAdptr
  '  Debug.Print objAdptr.MACAddress
    macOriginal = objAdptr.MACAddress
Next objAdptr

If macLiscense <> macOriginal Then
    MsgBox "Please contact us for technical support to get a new copy!", vbInformation, "Invalid Copy"
    DoCmd.Quit acQuitSaveNone
End If
End Function
