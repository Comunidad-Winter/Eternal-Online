Attribute VB_Name = "modMacros"
Public Enum eMacros
    aComando = 1
    aLanzar
    aEquipar
    aUsar
End Enum

Public Type tMacros
    mTipe As Byte
    Grh As Integer
    Nombre As String
    slot As Byte
    ObjIndex As Integer
    SpellSlot As Byte
End Type
Public MacroIndex As Integer

Public MacroList(1 To 6) As tMacros
Public Sub LoadMacros(ByVal Nombre As String)
'***************************************************
'Author:Bateman
'***************************************************
    Dim MacroPatch As String
    Dim i As Integer
    MacroPatch = App.Path & "\resources\Init\MACROS\" & Nombre & ".mac"
    If FileExist(MacroPatch, vbNormal) Then
        For i = 1 To 6
            With MacroList(i)
                .Nombre = GetVar(MacroPatch, "Macro" & i, "Nombre")
                .Grh = Val(GetVar(MacroPatch, "Macro" & i, "Grh"))
                .mTipe = Val(GetVar(MacroPatch, "Macro" & i, "Tipo"))
                .slot = Val(GetVar(MacroPatch, "Macro" & i, "Slot"))
                .SpellSlot = Val(GetVar(MacroPatch, "Macro" & i, "SlotSpell"))
                .ObjIndex = Val(GetVar(MacroPatch, "Macro" & i, "ObjIndex"))
            End With
        Next i
    Else
        For i = 1 To 6
            With MacroList(i)
                .Nombre = vbNullString
                .Grh = 0
                .mTipe = 0
                .slot = 0
                .SpellSlot = 0
                .ObjIndex = 0
            End With
         Next i
            Call SaveMacros(Nombre)
    End If
End Sub
Public Sub SaveMacros(ByVal Nombre As String)
'***************************************************
'Author:Bateman
'***************************************************
    Dim MacroPatch As String
    Dim i As Integer
    MacroPatch = App.Path & "\resources\Init\MACROS\" & Nombre & ".mac"

        For i = 1 To 6
            With MacroList(i)
                Call WriteVar(MacroPatch, "Macro" & i, "Nombre", .Nombre)
                Call WriteVar(MacroPatch, "Macro" & i, "Grh", .Grh)
                Call WriteVar(MacroPatch, "Macro" & i, "Tipo", .mTipe)
                Call WriteVar(MacroPatch, "Macro" & i, "Slot", .slot)
                Call WriteVar(MacroPatch, "Macro" & i, "SlotSpell", .SpellSlot)
                Call WriteVar(MacroPatch, "Macro" & i, "ObjIndex", .ObjIndex)
            End With
        Next i
End Sub
Public Function CheckMacrosSpells(ByVal SlotSpells As Byte, ByVal NameSpell As String, ByVal MacroIndex As Byte) As Byte
'***************************************************
'Author:Bateman
'***************************************************
    Dim i As Integer
    If SlotSpells < 0 Or SlotSpells > MAXHECHI - 1 Or _
       NameSpell = "" Then Exit Function

    If frmMain.hlst.List(SlotSpells) = NameSpell Then
        CheckMacrosSpells = SlotSpells
        Exit Function
    Else
        'Cambio el Slot del spells :P,entonces lo buscamos
        For i = 0 To MAXHECHI - 1
            If frmMain.hlst.List(i) = NameSpell Then
                Exit For
            End If
        Next i

      
        CheckMacrosSpells = i
        MacroList(MacroIndex).SpellSlot = i
        Call SaveMacros(UserName)
        Exit Function
    End If
    'ERROR!!
    CheckMacrosSpells = -1
    MacroList(MacroIndex).mTipe = 0

End Function
Public Function UsarYequiparObjValido(ByVal TIPO As Integer, ByVal Usable As Boolean) As Boolean
'***************************************************
'Author:Bateman
'***************************************************
    If Usable Then
        UsarYequiparObjValido = _
        TIPO = eObjType.otBarcos Or _
                                TIPO = eObjType.otBebidas Or _
                                TIPO = eObjType.otBotellaLlena Or _
                                TIPO = eObjType.otBotellaVacia Or _
                                TIPO = eObjType.otGuita Or _
                                TIPO = eObjType.otInstrumentos Or _
                                TIPO = eObjType.otLlaves Or _
                                TIPO = eObjType.otMinerales Or _
                                TIPO = eObjType.otPergaminos Or _
                                TIPO = eObjType.otPociones Or _
                                TIPO = eObjType.otWeapon
    Else
        UsarYequiparObjValido = _
        TIPO = eObjType.otAnillo Or _
                                TIPO = eObjType.otArmadura Or _
                                TIPO = eObjType.otcasco Or _
                                TIPO = eObjType.otescudo Or _
                                TIPO = eObjType.otFlechas Or _
                                TIPO = eObjType.otWeapon
    End If
End Function
Public Function CheckMacrosUsarItem(ByVal slot As Byte, ByVal ObjIndex As Integer, ByVal MacroIndex As Byte) As Byte
'***************************************************
'Author:Bateman
'***************************************************
    Dim i As Byte

    If slot = 0 Or slot > MAX_INVENTORY_SLOTS Then Exit Function

    If Inventario.ObjIndex(slot) = ObjIndex Then
        CheckMacrosUsarItem = slot
        Exit Function
    Else
        For i = 1 To MAX_INVENTORY_SLOTS - 1
            If Inventario.ObjIndex(i) = ObjIndex Then
                Exit For
            End If
        Next i

        If Inventario.ObjIndex(i) = ObjIndex Then
            CheckMacrosUsarItem = i
            MacroList(MacroIndex).slot = i
            Call SaveMacros(UserName)
            Exit Function
        Else
            CheckMacrosUsarItem = 0
        End If


        Exit Function
    End If
End Function
Public Sub UsarMacro(ByVal Index As Byte)
'***************************************************
'Author:Bateman
'***************************************************
    Dim slot As Byte

    Select Case MacroList(Index).mTipe

    Case eMacros.aLanzar
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        slot = CheckMacrosSpells(MacroList(Index).SpellSlot, MacroList(Index).Nombre, Index)
        If slot < 0 Then
            Exit Sub
        End If
        Call WriteCastSpell(slot + 1)
        Call WriteWork(eSkill.Magia)

    Case eMacros.aUsar
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        slot = CheckMacrosUsarItem(MacroList(Index).slot, MacroList(Index).ObjIndex, Index)
        If slot = 0 Then Exit Sub
        If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
        If MainTimer.Check(TimersIndex.UseItemWithU) Then _
           Call WriteUseItem(slot)
        If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo

    Case eMacros.aEquipar
        slot = CheckMacrosUsarItem(MacroList(Index).slot, MacroList(Index).ObjIndex, Index)
        If slot = 0 Then Exit Sub

        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            If Comerciando Then Exit Sub
            Call WriteEquipItem(slot)
        End If
    Case eMacros.aComando
    If LenB(MacroList(Index).Nombre) > 0 Then _
    Call ParseUserCommand(MacroList(Index).Nombre)
    End Select

End Sub

Public Function TotalItemAmountGet(ByVal ObjIndex As Integer) As Long
    Dim i As Byte
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ObjIndex(i) = ObjIndex Then
            TotalItemAmountGet = TotalItemAmountGet + Inventario.Amount(i)
        End If
    Next i
End Function

Public Function ObjIndexEquipped(ByVal ObjIndex As Integer) As Boolean
    Dim i As Byte
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.Equipped(i) And (Inventario.ObjIndex(i) = ObjIndex) Then
            ObjIndexEquipped = True
            Exit Function
        End If
    Next i
End Function
