Attribute VB_Name = "modLightEngine"
Option Explicit

Public Type LightVertex
    type As Byte
    affected As Byte
End Type

Public Type DayLightType
    R As Byte
    G As Byte
    B As Byte
End Type

Public DayLightByte As DayLightType

Public TwinkLightByteHandle As Long

Public LightMap(1 To 100, 1 To 100) As LightVertex

Public Function LightValue(value As Integer) As Long

If value > 255 Then value = 255
value = value - TwinkLightByteHandle
LightValue = RGB(value, value, value)

End Function

Public Function DayLight() As Long

With DayLightByte

    .R = 200
    .G = 150
    .B = 150
    
    DayLight = RGB(.R, .G, .B)
    
End With

End Function

Public Function GetLightValue(ByVal X As Byte, ByVal Y As Byte, vertice As Byte) As Long

Select Case vertice

        Case 0: 'DN LT VERTEX
        If Y > 99 Then Exit Function
        With LightMap(X, Y + 1)
            If .affected Then
                GetLightValue = LightValue(150 + .affected * (255 - 150) / 4)
            Else
                GetLightValue = DayLight
            End If
        End With
        
    Case 1: 'UP LT VERTEX
        With LightMap(X, Y)
            If .affected Then
                GetLightValue = LightValue(150 + .affected * (255 - 150) / 4)
            Else
                GetLightValue = DayLight
            End If
        End With
        
    Case 2: 'DN RT VERTEX
        If X > 99 Or Y > 99 Then Exit Function
        With LightMap(X + 1, Y + 1)
            If .affected Then
                GetLightValue = LightValue(150 + .affected * (255 - 150) / 4)
            Else
                GetLightValue = DayLight
            End If
        End With
        
    Case 3: 'UP RT VERTEX
        If X > 99 Then Exit Function
        With LightMap(X + 1, Y)
            If .affected Then
                GetLightValue = LightValue(150 + .affected * (255 - 150) / 4)
            Else
                GetLightValue = DayLight
            End If
        End With

End Select

End Function

Public Sub SetLight(X As Byte, Y As Byte)

With LightMap(X, Y)
    .affected = 4
    .type = 1
End With

AffectVertex X + 1, Y, 2
AffectVertex X - 1, Y, 2
AffectVertex X, Y - 1, 2
AffectVertex X, Y + 1, 2

AffectVertex X - 1, Y - 1, 1
AffectVertex X + 1, Y - 1, 1
AffectVertex X - 1, Y + 1, 1
AffectVertex X + 1, Y + 1, 1

End Sub

Public Sub AffectVertex(X As Byte, Y As Byte, value As Byte)

With LightMap(X, Y)
    If .affected < value Then .affected = value
    .type = 1
End With

End Sub

