Attribute VB_Name = "modpower"
Public Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SystemPowerStatus) As Long
Public Type SystemPowerStatus
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    Reserved1 As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type





