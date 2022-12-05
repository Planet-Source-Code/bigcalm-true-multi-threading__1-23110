Attribute VB_Name = "Declarations"
Option Explicit

' BitBlt included here so that the time it takes to put the map on the screen is reduced
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062
' Function to time A*
Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Type MapNode2
    MoveCost As Long
    NodeType As Long
End Type
Public Const MaxX As Long = 17 ' These constants would normally be part of Map - i.e. Map.Width, Map.Height
Public Const MaxY As Long = 19
Public MapArr(MaxX, MaxY) As MapNode
