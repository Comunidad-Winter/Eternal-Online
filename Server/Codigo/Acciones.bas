Attribute VB_Name = "Acciones"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim tempIndex As Integer
    
On Error Resume Next
    '�Rango Visi�n? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    '�Posicion valida?
    If InMapBounds(map, X, Y) Then
        With UserList(UserIndex)
            If MapData(map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(map, X, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then
                    '�Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    '�Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'A depositar de una
                    Call IniciarDeposito(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
                        Call RevivirUsuario(UserIndex)
                    End If
                    
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                        If .flags.Envenenado Then
                            .flags.Envenenado = 0
                            Call WriteCharVeneno(UserIndex, .Char.CharIndex, .flags.Envenenado)
                            Call WriteConsoleMsg(UserIndex, "��El sacerdote curo tu envenenamiento!!", FontTypeNames.FONTTYPE_VENENO)
                        End If
                    
                        'curamos totalmente
                        .Stats.MinHp = .Stats.MaxHp
                        Call WriteUpdateUserStats(UserIndex)
                    End If
                End If
                
            '�Es un obj?
            ElseIf MapData(map, X, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, X, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X, Y, UserIndex)
                    Case eOBJType.otCarteles 'Es un cartel
                        Call AccionParaCartel(map, X, Y, UserIndex)
                    Case eOBJType.otForos 'Foro
                        Call AccionParaForo(map, X, Y, UserIndex)
                    Case eOBJType.otLe�a    'Le�a
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(map, X, Y, UserIndex)
                        End If
                End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X + 1, Y, UserIndex)
                    
                End Select
            
            ElseIf MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X + 1, Y + 1, UserIndex)
                End Select
            
            ElseIf MapData(map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X, Y + 1, UserIndex)
                End Select
            End If
        End With
    End If
End Sub

Public Sub AccionParaForo(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Agrego foros faccionarios
'***************************************************

On Error Resume Next

    Dim Pos As WorldPos
    
    Pos.map = map
    Pos.X = X
    Pos.Y = Y
    
    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If SendPosts(UserIndex, ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).ForoID) Then
        Call WriteShowForumForm(UserIndex)
    End If
    
End Sub

Sub AccionParaPuerta(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                    
                    'Desbloquea
                    MapData(map, X, Y).Blocked = 0
                    MapData(map, X - 1, Y).Blocked = 0
        
                    MapData(map, X, Y + 1).Blocked = 0
                    MapData(map, X - 1, Y + 1).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, map, X, Y, 0)
                    Call Bloquear(True, map, X - 1, Y, 0)
                    Call Bloquear(True, map, X, Y + 1, 0)
                    Call Bloquear(True, map, X - 1, Y + 1, 0)
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                     Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                                
                MapData(map, X, Y).Blocked = 15
                MapData(map, X - 1, Y).Blocked = 15
                MapData(map, X, Y + 1).Blocked = 14
                MapData(map, X - 1, Y + 1).Blocked = 14
                
                
                Call Bloquear(False, UserIndex, X, Y, 15)
                Call Bloquear(False, UserIndex, X - 1, Y, 15)
                Call Bloquear(False, UserIndex, X, Y + 1, 14)
                Call Bloquear(False, UserIndex, X - 1, Y + 1, 14)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(map, X, Y).ObjInfo.ObjIndex
    Else
        Call WriteConsoleMsg(UserIndex, "La puerta est� cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
    Call WriteShowSignal(UserIndex, MapData(map, X, Y).ObjInfo.ObjIndex)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj

Dim Pos As WorldPos
Pos.map = map
Pos.X = X
Pos.Y = Y

With UserList(UserIndex)
    If Distancia(Pos, .Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If MapData(map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(map).Pk = False Then
        Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If .Stats.UserSkills(Supervivencia) > 1 And .Stats.UserSkills(Supervivencia) < 6 Then
        Suerte = 3
    ElseIf .Stats.UserSkills(Supervivencia) >= 6 And .Stats.UserSkills(Supervivencia) <= 10 Then
        Suerte = 2
    ElseIf .Stats.UserSkills(Supervivencia) >= 10 And .Stats.UserSkills(Supervivencia) Then
        Suerte = 1
    End If
    
    exito = RandomNumber(1, Suerte)
    
    If exito = 1 Then
        If MapInfo(.Pos.map).Zona <> Ciudad Then
            Obj.ObjIndex = FOGATA
            Obj.Amount = 1
            
            Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
            
            Call MakeObj(Obj, map, X, Y)
            
            'Las fogatas prendidas se deben eliminar
            Dim Fogatita As New cGarbage
            Fogatita.map = map
            Fogatita.X = X
            Fogatita.Y = Y
            Call TrashCollector.Add(Fogatita)
            
            Call SubirSkill(UserIndex, eSkill.Supervivencia, True)
        Else
            Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
        Call SubirSkill(UserIndex, eSkill.Supervivencia, False)
    End If

End With

End Sub
