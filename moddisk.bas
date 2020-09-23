Attribute VB_Name = "moddisk"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
   Alias "GetLogicalDriveStringsA" _
   (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32" _
   Alias "GetDiskFreeSpaceA" _
  (ByVal lpRootPathName As String, _
   lpSectorsPerCluster As Long, _
   lpBytesPerSector As Long, _
   lpNumberOfFreeClusters As Long, _
   lpTtoalNumberOfClusters As Long) As Long
   
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
   Alias "GetDiskFreeSpaceExA" _
   (ByVal lpRootPathName As String, _
   lpFreeBytesAvailableToCaller As Currency, _
   lpTotalNumberOfBytes As Currency, _
   lpTotalNumberOfFreeBytes As Currency) As Long
      
Private Declare Function GetModuleHandle Lib "kernel32" _
   Alias "GetModuleHandleA" _
  (ByVal lpModuleName As String) As Long
  
Private Declare Function GetProcAddress Lib "kernel32" _
  (ByVal hModule As Long, _
   ByVal lpProcName As String) As Long
   

   
      
Function GetDriveString() As String

  'returns string of available
  'drives each separated by a null
   Dim sBuffer As String
   
  'possible 26 drives, three characters each
  'plus a null, with a final trailing null
   sBuffer = Space$((26 * 4) + 1)
  
  If GetLogicalDriveStrings(Len(sBuffer), sBuffer) Then

     'do not strip off trailing null
      GetDriveString = Trim$(sBuffer)
      
   End If

End Function


Private Sub LoadAvailableDrives(cmbo As ComboBox)

   Dim lpBuffer As String

  'get list of available drives
   lpBuffer = GetDriveString()

  'Separate the drive strings
  'and add to the combo. StripNulls
  'will continually shorten the
  'string. Loop until a single
  'remaining terminating null is
  'encountered.
   Do Until lpBuffer = Chr(0)
  
    'strip off one drive item
    'and add to the combo
     cmbo.AddItem StripNulls(lpBuffer)
    
   Loop
  
End Sub


Function StripNulls(startstr As String) As String

  'Take a string separated by chr$(0)
  'and split off 1 item, shortening the
  'string so next item is ready for removal.
   Dim pos As Long

   pos = InStr(startstr$, Chr$(0))
  
   If pos Then
      
      StripNulls = Mid$(startstr, 1, pos - 1)
      startstr = Mid$(startstr, pos + 1, Len(startstr))
      Exit Function
    
   End If

End Function


Function GetDiskSpaceUsed(sDrive As String) As Currency

  'for GetDiskFreeSpaceEx
   Dim BytesFree As Currency
   Dim TotalBytes As Currency
   Dim TotalBytesFree As Currency
   Dim TotalBytesUsed As Currency
   
  'for GetDiskFreeSpace
   Dim nSectors As Long
   Dim nBytesPerSector As Long
   Dim nFreeClusters As Long
   Dim nTotalClusters As Long
   Dim DrvSpaceTotal As Long
   Dim DrvSpaceFree As Long
     
  'for GetProcAddress
   Dim ptr As Long
  
  'attempt to obtain a pointer to
  'the GetDiskFreeSpaceExA API in kernel32
   ptr = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA")
  
   If ptr Then

     'get drive info using GetDiskFreeSpaceEx
      If GetDiskFreeSpaceEx(sDrive, _
                            BytesFree, _
                            TotalBytes, _
                            TotalBytesFree) <> 0 Then
      
        'adjust the by multiplying the returned
        'value by 10000 accommodate for the decimal
        'places the currency data type returns.
         GetDiskSpaceUsed = (TotalBytes - BytesFree) * 10000
      
      End If  'if GetDiskFreeSpaceEx
   
   Else
   
     'get drive info using GetDiskFreeSpace
      If GetDiskFreeSpace(sDrive, nSectors, _
                          nBytesPerSector, _
                          nFreeClusters, _
                          nTotalClusters) <> 0 Then
   
        'perform math to get the data
         On Local Error Resume Next
         DrvSpaceTotal = (nSectors * nBytesPerSector * nTotalClusters)
         DrvSpaceFree = (nSectors * nBytesPerSector * nFreeClusters)
         GetDiskSpaceUsed = (DrvSpaceTotal - DrvSpaceFree)
         On Local Error GoTo 0
     
     End If  'if GetDiskFreeSpace
   End If  'If ptr

End Function


Function GetDiskSpace(sDrive As String) As Currency

  'for GetDiskFreeSpaceEx
   Dim BytesFree As Currency
   Dim TotalBytes As Currency
   Dim TotalBytesFree As Currency
   Dim TotalBytesUsed As Currency
   
  'for GetDiskFreeSpace
   Dim nSectors As Long
   Dim nBytesPerSector As Long
   Dim nFreeClusters As Long
   Dim nTotalClusters As Long
     
  'for GetProcAddress
   Dim ptr As Long
  
  'attempt to obtain a pointer to
  'the GetDiskFreeSpaceExA API in kernel32
   ptr = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA")
   
   If ptr Then

     'get drive info using GetDiskFreeSpaceEx
      If GetDiskFreeSpaceEx(sDrive, _
                            BytesFree, _
                            TotalBytes, _
                            TotalBytesFree) <> 0 Then
      
        'adjust the by multiplying the returned
        'value by 10000 accommodate for the decimal
        'places the currency data type returns.
         GetDiskSpace = TotalBytes * 10000
      
      End If  'if GetDiskFreeSpaceEx
   
   Else
   
     'get drive info using GetDiskFreeSpace
      If GetDiskFreeSpace(sDrive, nSectors, _
                          nBytesPerSector, _
                          nFreeClusters, _
                          nTotalClusters) <> 0 Then
   
        'perform math to get the data
         On Local Error Resume Next
         GetDiskSpace = (nSectors * nBytesPerSector * nTotalClusters)
         On Local Error GoTo 0
     
     End If  'if GetDiskFreeSpace
   End If  'If ptr
   
End Function


Function GetDiskSpaceFree(sDrive As String) As Currency

  'for GetDiskFreeSpaceEx
   Dim BytesFree As Currency
   Dim TotalBytes As Currency
   Dim TotalBytesFree As Currency
   Dim TotalBytesUsed As Currency
   
  'for GetDiskFreeSpace
   Dim nSectors As Long
   Dim nBytesPerSector As Long
   Dim nFreeClusters As Long
   Dim nTotalClusters As Long
     
  'for GetProcAddress
   Dim ptr As Long
  
  'attempt to obtain a pointer to
  'the GetDiskFreeSpaceExA API in kernel32
   ptr = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA")
  
   If ptr Then

     'get drive info using GetDiskFreeSpaceEx
      If GetDiskFreeSpaceEx(sDrive, _
                            BytesFree, _
                            TotalBytes, _
                            TotalBytesFree) <> 0 Then
      
        'adjust the by multiplying the returned
        'value by 10000 accommodate for the decimal
        'places the currency data type returns.
         GetDiskSpaceFree = TotalBytesFree * 10000
      
      End If  'if GetDiskFreeSpaceEx
   
   Else
   
     'get drive info using GetDiskFreeSpace
      If GetDiskFreeSpace(sDrive, _
                          nSectors, _
                          nBytesPerSector, _
                          nFreeClusters, _
                          nTotalClusters) <> 0 Then
   
        'perform math to get the data
         On Local Error Resume Next
         GetDiskSpaceFree = (nSectors * nBytesPerSector * nFreeClusters)
         On Local Error GoTo 0
     
     End If  'if GetDiskFreeSpace
   End If  'If ptr

End Function


Function GetDiskBytesAvailable(sDrive As String) As Currency

  'for GetDiskFreeSpaceEx
   Dim BytesFree As Currency
   Dim TotalBytes As Currency
   Dim TotalBytesFree As Currency
   Dim TotalBytesUsed As Currency
   
  'for GetDiskFreeSpace
   Dim nSectors As Long
   Dim nBytesPerSector As Long
   Dim nFreeClusters As Long
   Dim nTotalClusters As Long
     
  'for GetProcAddress
   Dim ptr As Long
  
  'attempt to obtain a pointer to
  'the GetDiskFreeSpaceExA API in kernel32
   ptr = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA")
  
   If ptr Then

     'get drive info using GetDiskFreeSpaceEx
      If GetDiskFreeSpaceEx(sDrive, _
                            BytesFree, _
                            TotalBytes, _
                            TotalBytesFree) <> 0 Then
      
        'adjust the by multiplying the returned
        'value by 10000 accommodate for the decimal
        'places the currency data type returns.
         GetDiskBytesAvailable = BytesFree * 10000
      
      End If  'if GetDiskFreeSpaceEx
   
   Else
   
     'get drive info using GetDiskFreeSpace
      If GetDiskFreeSpace(sDrive, _
                          nSectors, _
                          nBytesPerSector, _
                          nFreeClusters, _
                          nTotalClusters) <> 0 Then
   
        'bytes available is not returned,
        'so return the free space instead.
         On Local Error Resume Next
         GetDiskBytesAvailable = (nSectors * nBytesPerSector * nFreeClusters)
         On Local Error GoTo 0
     
     End If  'if GetDiskFreeSpace
   End If  'If ptr
   
End Function

