Attribute VB_Name = "modMinimap"
Option Explicit
Private lMapColor() As Long
Private lMapColorCount As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Sub LoadMapColor()
  
    'If FileExist(App.Path & "\GrhIndex\MapColor.bin", vbArchive) = False Then
    '    MsgBox "Falta el archivo 'MapColor.bin'", vbCritical
    '    End
    'End If
    
    Dim iT As Long
    Dim handle As Long
    handle = FreeFile
    
    Open IniPath & "small_maps.bin" For Binary Access Read As handle
    Seek handle, 1
    Get handle, , lMapColorCount
    
    ReDim lMapColor(1 To lMapColorCount) As Long
    
    For iT = 1 To lMapColorCount
        Get handle, , lMapColor(iT)
    Next
    
    Close handle
End Sub
Public Sub DrawMiniMap()
Dim bMapX As Byte, bMapY As Byte, bCapa As Byte, tGrh As Long
frmMain.Minimap.Cls
    If lMapColorCount = 0 Then Exit Sub
    For bMapX = 1 To 100
        For bMapY = 1 To 100
            For bCapa = 1 To 3 ' Dibujamos las 3 primeras capas
                tGrh = MapData(bMapX, bMapY).Layer(bCapa).GrhIndex
                If tGrh > 0 And tGrh <= lMapColorCount Then
                    SetPixel frmMain.Minimap.hdc, bMapX - 1, bMapY - 1, lMapColor(tGrh)
                End If
            Next bCapa
        Next bMapY
    Next bMapX
frmMain.Minimap.Refresh
End Sub

Public Sub SmallMap_UserPOS()
    frmMain.UserP.Left = UserPos.X
    frmMain.UserP.Top = UserPos.Y
    frmMain.Minimap.Refresh
End Sub

