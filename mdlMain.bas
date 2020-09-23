Attribute VB_Name = "mdlMain"
Public Const DAWN_LIGHT As Long = &HA0A0A0
Public Const MIDDAY_LIGHT As Long = &HFFFFFF
Public Const DUSK_LIGHT As Long = &HC0C0C0
Public Const EVENING_LIGHT As Long = &H606060
Public Const NIGHT_LIGHT As Long = &H101010
Public Const EARLYMORNING_LIGHT As Long = &H303030

Public lngFillMode As Long
Public lngLightType As Long

Public searchLightIntensity As Single
Public torchLightIntensity As Single

Sub Main()
    searchLightIntensity = 0.2
    torchLightIntensity = 0.2
    Load frmMain
    frmMain.Show
End Sub
