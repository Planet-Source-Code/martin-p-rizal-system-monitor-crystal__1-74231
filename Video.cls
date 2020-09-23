VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim ssdxs As DirectX7 'direct x 7 main object
Dim ssdds As DirectDraw7 'direct draw 7 module
Dim ssddsds As DDSURFACEDESC2 'Surface Description type
Dim ssddis As DirectDrawIdentifier 'direct draw identifier
Dim ssddEnums As DirectDrawEnum 'direct draw enumeration
Dim ssddCapss As DDSCAPS2 'direct draw capabilities structure

Dim rect As rect
Dim n As Long
Dim cname As String
Dim cres As String
Dim cdepth As String
Dim cmem As String

Sub QueryVideoInfo()
On Error GoTo xc:

Set ssdxs = New DirectX7  'create direct x object

Set ssdds = ssdxs.DirectDrawCreate("")
Set ssddis = ssdds.GetDeviceIdentifier(DDGDI_DEFAULT)
Set ssddEnums = ssdxs.GetDDEnum
ssdds.GetDisplayMode ssddsds

cmem = FormatFileSize(ssdds.GetFreeMem(ssddCapss))
cname = ssddis.GetDescription
cres = ssddsds.lWidth & "x" & ssddsds.lHeight
cdepth = ssddsds.ddpfPixelFormat.lRGBBitCount

Set ssdds = Nothing
Set ssddis = Nothing
Set ssddEnums = Nothing

Exit Sub

xc:

cmem = "Unknown"
cname = "Unknown"
cres = "Unknown"
cdepth = "Unknown"

End Sub

Property Get VideoCardName() As String
VideoCardName = cname
End Property

Property Get VideoResolution() As String
VideoResolution = cres
End Property

Property Get VideoMemory() As String
VideoMemory = cmem
End Property

Property Get ColorDepth() As String
ColorDepth = cdepth & "-bit"
End Property
