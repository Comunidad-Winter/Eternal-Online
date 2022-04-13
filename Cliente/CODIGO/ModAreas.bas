Attribute VB_Name = "ModAreas"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer

Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)
    Dim loopX As Long, loopY As Long
    MinLimiteX = (X \ 9 - 1) * 9
    MaxLimiteX = MinLimiteX + 26
    MinLimiteY = (Y \ 9 - 1) * 9
    MaxLimiteY = MinLimiteY + 26
    
    For loopX = MinMapSize To MaxMapSize
        For loopY = MinMapSize To MaxMapSize
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                If MapData(loopX, loopY).CharIndex > 0 Then
                    If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                            charlist(MapData(loopX, loopY).CharIndex).StatusAlpha = True
                            'Call EraseChar(MapData(loopX, loopY).CharIndex)
                    End If
                End If
                
                'Erase OBJs
                If EsObjMapeado(MapData(loopX, loopY).ObjGrh.GrhIndex) Then 'si objetos del mapa es al pedo, ahora si son del user que los borre.
                    MapData(loopX, loopY).ObjGrh.GrhIndex = 0
                End If
                
                'Erase Huellas // no lo veo necesario que se queden las huellas en el mapa.
                'MapData(loopX, loopY).Huella.GrhIndex = 0
            End If
        Next loopY
    Next loopX
    
    Call RefreshAllChars
End Sub

Public Function EsObjMapeado(ByVal grh_index As Integer) As Boolean
' tengo que ver como hago para hacer que lea el objtype, por que esto de hacerlo por grhindex es un asco.
'pd: faltan algunas puertas.

Dim i As Byte '// para el for.
    For i = 0 To 23
        If grh_index = 11121 + i Then
            EsObjMapeado = False
            Exit Function
        End If
    Next i
    For i = 0 To 43
        If grh_index = 11199 + i Then
            EsObjMapeado = False
            Exit Function
        End If
    Next i
    
    If grh_index = 11456 Or 11457 Or 11464 Or 11465 Or 11468 Or 11469 Or 11489 Or 11490 Or 11491 Or 11492 Or 11493 Or 11494 Then
        EsObjMapeado = False
        Exit Function
    End If

EsObjMapeado = True

End Function
