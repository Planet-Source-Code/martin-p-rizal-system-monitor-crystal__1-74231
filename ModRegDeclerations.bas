Attribute VB_Name = "ModRegDeclerations"
Option Explicit

'Registery Stuff
Public Declare Function RegCloseKey Lib "advapi32" (ByVal HKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32" (ByVal HKey As Long) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal HKey As Long, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
    
    Public Const READ_CONTROL = &H20000
    Public Const SYNCHRONIZE = &H100000
    Public Const STANDARD_RIGHTS_ALL = &H1F0000
    Public Const STANDARD_RIGHTS_READ = READ_CONTROL
    Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    Public Const KEY_CREATE_SUB_KEY = &H4
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_NOTIFY = &H10
    Public Const KEY_CREATE_LINK = &H20
    Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    
    Public Const ERROR_SUCCESS = 0&
    Public Const ERROR_BADDB = 1009&
    Public Const ERROR_BADKEY = 1010&
    Public Const ERROR_CANTOPEN = 1011&
    Public Const ERROR_CANTREAD = 1012&
    Public Const ERROR_CANTWRITE = 1013&
    Public Const ERROR_OUTOFMEMORY = 14&
    Public Const ERROR_INVALID_PARAMETER = 87&
    Public Const ERROR_ACCESS_DENIED = 5&
    
    Public Const REG_NONE = 0
    Public Const REG_SZ = 1
    Public Const REG_EXPAND_SZ = 2
    Public Const REG_BINARY = 3
    Public Const REG_DWORD = 4
    Public Const REG_DWORD_LITTLE_ENDIAN = 4
    Public Const REG_DWORD_BIG_ENDIAN = 5

Function GetCPUDescription() As String

    Dim HKey As Long
    Dim C As Long
    Dim R As Long
    Dim S As String
    Dim T As Long
    
    R = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", 0, KEY_READ, HKey)
    C = 255
    S = String(C, Chr(0))
    R = RegQueryValueEx(HKey, "ProcessorNameString", 0, T, S, C)
    
    GetCPUDescription = Trim(Left(S, C - 1))

End Function
