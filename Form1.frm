VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmstatus 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "System Monitor"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pblevel 
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbdisk 
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbcpu 
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbpage 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbvirtual 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbmload 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbram 
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4200
      Top             =   2760
   End
   Begin VB.Label lbldep 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   47
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      Caption         =   "Display Color Depth:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblres 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   45
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   "Video Resolution:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lblvcname 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   43
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "Video Card Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblgram 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   41
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Caption         =   "Video Memory:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblcpuname 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   39
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "Processor Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "Power Life:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lbllife 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   36
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lblstatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   35
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Power Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lbllevel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Power Level:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblpower 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   30
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Power Source:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblhd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   28
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Disk Capacity:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lbldisk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Disk Space:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblx 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      ToolTipText     =   "Close"
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblpage2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Total Page File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblvram 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Total Virtual Memory:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblram2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Total Physical Memory:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblcpu 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "CPU Usage:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Processor Speed:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblcpuspeed 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Page File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblpage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblmload 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblvirtual 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Virtual memory:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Memory Load:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblram 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Physical Memory:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private QueryObject As Object

Private Declare Sub GlobalMemoryStatusEx Lib "KERNEL32" (lpBuffer As MEMORYSTATUSEX)

Private Type INT64
   LoPart As Long
   HiPart As Long
End Type

Private Type MEMORYSTATUSEX
   dwLength As Long
   dwMemoryLoad As Long
   ulTotalPhys As INT64
   ulAvailPhys As INT64
   ulTotalPageFile As INT64
   ulAvailPageFile As INT64
   ulTotalVirtual As INT64
   ulAvailVirtual As INT64
   ulAvailExtendedVirtual As INT64
End Type

Private Sub Form_Load()

If App.PrevInstance = True Then
End
End If


Dim vg As New VideoCard

If IsWinNT Then
    Set QueryObject = New clsWinNT
    Call SetFormOpacity(Me, 190)
  Else
    Set QueryObject = New clsWin9x
  End If
  
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - (Screen.Height - (Me.Height) + 5100)
  
QueryObject.Initialize

vg.QueryVideoInfo

PrintRamInformation
lblcpuname.Caption = GetCPUDescription
lblgram.Caption = vg.VideoMemory
lblvcname.Caption = vg.VideoCardName

drv = LCase(SystemDrive) & ":\"


End Sub


Private Sub lblx_Click()
End
End Sub

Private Sub Timer1_Timer()

Dim ram As Double
   
Dim udtMemStatEx As MEMORYSTATUSEX
   
udtMemStatEx.dwLength = Len(udtMemStatEx)

Call GlobalMemoryStatusEx(udtMemStatEx)

On Error Resume Next
pbcpu = CStr(QueryObject.Query)
lblcpu.Caption = Format(pbcpu.Value, "0.000") & " %"
   
If pbcpu.Value > 94 Then
Call alert(lblcpu, True)
Else
Call alert(lblcpu, False)
End If


ram = Round(CLargeInt(udtMemStatEx.ulAvailPhys.LoPart, udtMemStatEx.ulAvailPhys.HiPart) / (CLargeInt(udtMemStatEx.ulTotalPhys.LoPart, udtMemStatEx.ulTotalPhys.HiPart)) * 100)



pbram.Value = ram
lblram = Format(ram, "0.000") & " %"

If pbram.Value < 10 Then
Call alert(lblram, True)
Else
Call alert(lblram, False)
End If


pbmload.Value = CStr(udtMemStatEx.dwMemoryLoad)
lblmload = Format(CStr(udtMemStatEx.dwMemoryLoad), "0.000") & " %"

If pbmload.Value > 94 Then
Call alert(lblmload, True)
Else
Call alert(lblmload, False)
End If



pbpage.Value = (CLargeInt(udtMemStatEx.ulAvailPageFile.LoPart, udtMemStatEx.ulAvailPageFile.HiPart) / CLargeInt(udtMemStatEx.ulTotalPageFile.LoPart, udtMemStatEx.ulTotalPageFile.HiPart)) * 100
lblpage = Format(Round((CLargeInt(udtMemStatEx.ulAvailPageFile.LoPart, udtMemStatEx.ulAvailPageFile.HiPart) / CLargeInt(udtMemStatEx.ulTotalPageFile.LoPart, udtMemStatEx.ulTotalPageFile.HiPart)) * 100, 3), "0.000") & " %"

If pbpage.Value < 10 Then
Call alert(lblpage, True)
Else
Call alert(lblpage, False)
End If



pbvirtual = (CLargeInt(udtMemStatEx.ulAvailVirtual.LoPart, udtMemStatEx.ulAvailVirtual.HiPart) / CLargeInt(udtMemStatEx.ulTotalVirtual.LoPart, udtMemStatEx.ulTotalVirtual.HiPart)) * 100
lblvirtual = Format(Round(CLargeInt(udtMemStatEx.ulAvailVirtual.LoPart, udtMemStatEx.ulAvailVirtual.HiPart) / CLargeInt(udtMemStatEx.ulTotalVirtual.LoPart, udtMemStatEx.ulTotalVirtual.HiPart) * 100, 3), "0.000") & " %"

If pbvirtual.Value < 10 Then
Call alert(lblvirtual, True)
Else
Call alert(lblvirtual, False)
End If


pbdisk.Value = (GetDiskSpaceFree(drv) / GetDiskSpace(drv)) * 100

lbldisk = Format(((GetDiskSpaceFree(drv) / GetDiskSpace(drv)) * 100), "0.000") & " %"

If pbdisk.Value < 10 Then
Call alert(lbldisk, True)
Else
Call alert(lbldisk, False)
End If

On Error Resume Next
modCPUInfo.GetSysInfo

On Error Resume Next
lblcpuspeed = modCPUInfo.CPU_Speed

PowerCheck


Dim hn As New VideoCard

hn.QueryVideoInfo

lblres.Caption = hn.VideoResolution
lbldep.Caption = hn.ColorDepth

Set hn = Nothing

End Sub


Private Sub PrintRamInformation()

Dim udtMemStatEx As MEMORYSTATUSEX

udtMemStatEx.dwLength = Len(udtMemStatEx)

Call GlobalMemoryStatusEx(udtMemStatEx)

lblram2 = NumberInKB(CLargeInt(udtMemStatEx.ulTotalPhys.LoPart, udtMemStatEx.ulTotalPhys.HiPart))
lblvram = NumberInKB(CLargeInt(udtMemStatEx.ulTotalVirtual.LoPart, udtMemStatEx.ulTotalVirtual.HiPart))
lblpage2 = NumberInKB(CLargeInt(udtMemStatEx.ulTotalPageFile.LoPart, udtMemStatEx.ulTotalPageFile.HiPart))
lblhd = NumberInKB(GetDiskSpace(drv))


End Sub

'This function converts the LARGE_INTEGER data type to a double
Private Function CLargeInt(Lo As Long, Hi As Long) As Double
   Dim dblLo As Double
   Dim dblHi As Double

   If Lo < 0 Then
      dblLo = 2 ^ 32 + Lo
   Else
      dblLo = Lo
   End If

   If Hi < 0 Then
      dblHi = 2 ^ 32 + Hi
   Else
      dblHi = Hi
   End If

   CLargeInt = dblLo + dblHi * 2 ^ 32

End Function

Public Function NumberInKB(ByVal vNumber As Currency) As String
   Dim strReturn As String

   Select Case vNumber
      Case Is < 1024 ^ 1
         strReturn = CStr(vNumber) & " bytes"

      Case Is < 1024 ^ 2
         strReturn = Format(CStr(Round(vNumber / 1024, 1)), "0.000") & " KB"

      Case Is < 1024 ^ 3
         strReturn = Format(CStr(Round(vNumber / 1024 ^ 2, 2)), "0.000") & " MB"

      Case Is < 1024 ^ 4
         strReturn = Format(CStr(Round(vNumber / 1024 ^ 3, 2)), "0.000") & " GB"
      Case Is < 1024 ^ 5
         strReturn = Format(CStr(Round(vNumber / 1024 ^ 3, 2)), "0.000") & " TB"
      Case Is < 1024 ^ 6
         strReturn = Format(CStr(Round(vNumber / 1024 ^ 3, 2)), "0.000") & " PB"
      Case Is < 1024 ^ 6
         strReturn = Format(CStr(Round(vNumber / 1024 ^ 3, 2)), "0.000") & " EB"
   End Select

   NumberInKB = strReturn

End Function

Private Sub Timer2_Timer()
pbcpu = CStr(QueryObject.Query)
lblcpu.Caption = Format(pbcpu.Value, "0.000") & " %"
End Sub


Function alert(HostLabel As Label, AlertMode As Boolean)

If AlertMode = True Then
HostLabel.BackColor = vbRed
Else
HostLabel.BackColor = vbBlack
End If

End Function

Function PowerCheck()
Dim power_status As SystemPowerStatus
Dim power As SystemPowerStatus
Dim txt As String

Call GetSystemPowerStatus(power_status)


        If power_status.BatteryFlag = 1 Then
        txt = "High"
        ElseIf power_status.BatteryFlag = 2 Then
        txt = "Low"
        ElseIf power_status.BatteryFlag = 4 Then
        txt = "Critical"
        ElseIf power_status.BatteryFlag = 8 Or power_status.BatteryFlag = 9 Then
        txt = "Charging"
        ElseIf power_status.BatteryFlag = 128 Then
        txt = "No system battery"
        Else
        txt = "Unknown"
        End If
        lblstatus.Caption = txt
        
        
        
    If GetSystemPowerStatus(power) = 0 Then
        lblpower.Caption = "Error"
    Else
        Select Case power.ACLineStatus
            Case 0
                lblpower.Caption = "Battery"
            Case 1
                lblpower.Caption = "On Line"
            Case 255
                lblpower.Caption = "Unknown"
        End Select

    End If
    'If power_status.BatteryFullLifeTime = -1 Then
     '       lblFullLifetime.Caption = "Unknown"
      '  Else
      '      lblFullLifetime.Caption = _
      '          power_status.BatteryFullLifeTime & " " & _
      '          "seconds"
      '  End If

        If power_status.BatteryLifeTime = -1 Then
            lbllife.Caption = "Unknown"
        Else
            lbllife.Caption = ToTime(power_status.BatteryLifeTime)
        End If
        
        
        If power_status.BatteryLifePercent = 255 And power.ACLineStatus = 1 Then
        pblevel.Value = 100
        lbllevel = Format(100, "0.000") & " %"
        ElseIf power_status.BatteryLifePercent = 255 Then
        pblevel.Value = 0
        lbllevel = Format(0, "0.000") & " %"
        Else
        pblevel.Value = power_status.BatteryLifePercent
        lbllevel = Format(power_status.BatteryLifePercent, "0.000") & " %"
            If pblevel.Value <= 10 Then
            Call alert(lbllevel, True)
            Else
            Call alert(lbllevel, False)
            End If
        End If


End Function


Function ToTime(Value As Long) As String
Dim pref As Double

If Value > 59 Then
ToTime = Value & " seconds"
End If

If Value >= 60 Or Value <= 3599 Then
pref = Value / 60
ToTime = Round(pref, 2) & " minutes"
End If

If Value >= 3600 Then
pref = Value / 3600
ToTime = Round(pref, 2) & " hours"
End If

End Function

