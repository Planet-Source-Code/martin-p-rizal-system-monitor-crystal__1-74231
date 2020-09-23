Attribute VB_Name = "Module1"
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
' Windows version
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260
Global drv As String


Function SystemDrive() As String
    SystemDrive = Left$(WindowsDirectory(), 1)
End Function

' Return the Windows directory.
Function WindowsDirectory() As String
Dim windows_dir As String
Dim length As Long

    ' Get the Windows directory.
    windows_dir = Space$(MAX_PATH)
    length = GetWindowsDirectory(windows_dir, Len(windows_dir))
    WindowsDirectory = Left$(windows_dir, length)
End Function

'return True is the OS is WindowsNT3.5(1), NT4.0, 2000 or XP
Public Function IsWinNT() As Boolean
  Dim OSInfo As OSVERSIONINFO
  OSInfo.dwOSVersionInfoSize = Len(OSInfo)
  'retrieve OS version info
  GetVersionEx OSInfo
  'if we're on NT, return True
  IsWinNT = (OSInfo.dwPlatformId = 2)
End Function


