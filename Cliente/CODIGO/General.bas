Attribute VB_Name = "Mod_General"
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

'================================================================================================
' Saque una linea de 3 desde la resolucion mas baja para un frm que es 14x14 que permite VB6
' y cualquier resolucion multiplicado por 15 saco el scale real que toma VB6.
'
' 14Pixeles ----- 210 Valor de Scale
' 1920Pixeles ----- 1920x210 / 14 = 28800 ---- VALOR JUSTO QUE ME TOMABA VB6 AL ESTIRAR LOS FRM.
' 210/14 = 15
' 15 = VALOR REAL DE LOS FORMULARIOS.
'
'================================================================================================
' Resolucion de formularios - ===================================================================
Public frmScaleWidth As Long
Public frmScaleHeight As Long
Public Const M_SCALE_FRM As Byte = 15
'================================================================================================

Public bFogata As Boolean
Private lFrameTimer As Long
Private keysMovementPressedQueue As clsArrayList
Sub Main()
'===============================================
'Inizializate Eternal Online...
'===============================================
    Set frmMain.Client = New clsSocket
    Call LoadConfigINI
    Call ChangeScaleAllForms

    'Executed one more client?
    #If Testeo = 0 Then '// Evit in mode DEBUG!
        If FindPreviousInstance Then
            Call MsgBox("Eternal Online ya esta corriendo! No esta permitido ejecutar otra instancia del juego. Haga click en aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Eternal Online")
            End
        End If
    #End If
    
    'Full Screen?
    If Config_Inicio.FullScreen = True Then
        Call ChangeResolution.SetResolution
    End If
    
    Call Init_Names
    Call Protocol.Init_Fonts
    
    Set keysMovementPressedQueue = New clsArrayList
    Call keysMovementPressedQueue.Initialize(1, 4)
    
    UserMap = 1
    engine.DirectX_Init
    Call LoadGrhData
    Call LoadMapColor
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call CargarParticulas
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call engine.Font_Initialize
    Call Load_Climas
    Call Inventario.Initialize(frmMain.PicInv, MAX_INVENTORY_SLOTS)
    
    'Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hWnd, DirSound, DirMidi)
    Audio.MusicActivated = Config_Inicio.Music
    Audio.SoundActivated = Config_Inicio.Sounds
    Audio.SoundVolume = Config_Inicio.SoundVolume
    Audio.MusicVolume = Config_Inicio.MusicVolume
    Audio.SoundEffectsActivated = False 'MUY FEO MANITO XD
    Call Audio.MusicMP3Play(App.Path & "\Resources\Sounds\MP3\" & MP3_Inicio & ".mp3")
    
    INTRO = True
    frmConnect.Visible = True
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    Call Timers_Init
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)
    engine.Engine_Start
End Sub
Public Function DirInterface() As String
    DirInterface = App.Path & "\Resources\Interface\"
End Function
Public Function DirGraficos() As String
    DirGraficos = App.Path & "\Resources\"
End Function
Public Function DirSound() As String
    DirSound = App.Path & "\Resources\Sounds\WAV\"
End Function
Public Function DirMidi() As String
    DirMidi = App.Path & "\Resources\Sounds\MIDI\"
End Function
Public Function DirMapas() As String
    DirMapas = App.Path & "\Resources\Maps\"
End Function

Public Function DirExtras() As String
    DirExtras = App.Path & "\EXTRAS\"
End Function
Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim(Left(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

End Function
Sub CargarAnimArmas()
On Error Resume Next
    Dim N As Integer
    Dim i As Long
    Dim NumWeaponAnims As Integer

    Dim MisArmas() As tIndiceArmas
    
    N = FreeFile()
    Open IniPath & "Weapons.ind" For Binary Access Read As #N
    
    Get #N, , MiCabecera
    Get #N, , NumWeaponAnims
    
    'Resize array
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim MisArmas(1 To NumWeaponAnims) As tIndiceArmas
    
    For i = 1 To NumWeaponAnims
        Get #N, , MisArmas(i)
    
        If MisArmas(i).Dir(1) Then
            Call InitGrh(WeaponAnimData(i).WeaponWalk(1), MisArmas(i).Dir(1), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(2), MisArmas(i).Dir(2), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(3), MisArmas(i).Dir(3), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(4), MisArmas(i).Dir(4), 0)
        End If
    Next i
    
    Close #N
End Sub
Sub CargarAnimEscudos()
On Error Resume Next
    Dim N As Integer
    Dim i As Long
    Dim NumWeaponAnims As Integer

    Dim MisEscudos() As tIndiceEscudos
    
    N = FreeFile()
    Open IniPath & "Shields.ind" For Binary Access Read As #N
    
    Get #N, , MiCabecera
    Get #N, , NumEscudosAnims
    
    'Resize array
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    ReDim MisEscudos(1 To NumEscudosAnims) As tIndiceEscudos
    
    For i = 1 To NumEscudosAnims
        Get #N, , MisEscudos(i)
    
        If MisEscudos(i).Dir(1) Then
            Call InitGrh(ShieldAnimData(i).ShieldWalk(1), MisEscudos(i).Dir(1), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(2), MisEscudos(i).Dir(2), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(3), MisEscudos(i).Dir(3), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(4), MisEscudos(i).Dir(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
Dim i As Byte '// for console render.

    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
    
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
    
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
    
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
    
        RichTextBox.Refresh
    
    End With

    If RichTextBox = frmMain.RecTxt Then
        For i = 2 To MaxLineas
            Con(i - 1).T = Con(i).T
            Con(i - 1).b = Con(i).b
            Con(i - 1).g = Con(i).g
            Con(i - 1).r = Con(i).r
            Con(i - 1).a = 30 + (i * 30)
        Next i
    
        Con(MaxLineas).T = Text
        Con(MaxLineas).b = blue
        Con(MaxLineas).g = green
        Con(MaxLineas).r = red
        Con(MaxLineas).a = 255

        OffSetConsola = 16
        UltimaLineavisible = False
    
    End If

End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopC As Long
    
    For loopC = 1 To LastChar
        If charlist(loopC).active = 1 Then
            MapData(charlist(loopC).Pos.X, charlist(loopC).Pos.Y).CharIndex = loopC
        End If
    Next loopC
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(Mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopC As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserPassword)
        CharAscii = Asc(Mid$(UserPassword, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopC
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserName)
        CharAscii = Asc(Mid$(UserName, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopC
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    
    'Vaciamos la cola de movimiento
    keysMovementPressedQueue.Clear
    
    'frmMain.lblName.Caption = UserName
    'Load main form
    Call SmallMap_UserPOS
    frmMain.Visible = True
    
    Call LoadMacros(UserName)
    
    FPSFLAG = True

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    'frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            If Not MapData(UserPos.X, UserPos.Y).Blocked = 14 And _
               Not MapData(UserPos.X, UserPos.Y).Blocked = 7 And _
               Not MapData(UserPos.X, UserPos.Y).Blocked = 11 Then
                    LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
            End If
        Case E_Heading.EAST
            If Not MapData(UserPos.X, UserPos.Y).Blocked = 8 And _
               Not MapData(UserPos.X, UserPos.Y).Blocked = 10 And _
               Not MapData(UserPos.X, UserPos.Y).Blocked = 11 Then
                    LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
            End If
        Case E_Heading.SOUTH
            If Not MapData(UserPos.X, UserPos.Y).Blocked = 6 And _
               Not MapData(UserPos.X, UserPos.Y).Blocked = 10 And _
               Not MapData(UserPos.X, UserPos.Y).Blocked = 15 Then
                    LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
            End If
        Case E_Heading.WEST
            If Not MapData(UserPos.X, UserPos.Y).Blocked = 12 And _
               Not MapData(UserPos.X, UserPos.Y).Blocked = 6 And _
               Not MapData(UserPos.X, UserPos.Y).Blocked = 7 Then
                    LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
            End If
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
            Call SmallMap_UserPOS
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
            Call SmallMap_UserPOS
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub
Private Sub AddMovementToKeysMovementPressedQueue()
    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Remueve la tecla que teniamos presionada
    End If
End Sub

Public Function checkDIRPAD() As Boolean

checkDIRPAD = False

If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then checkDIRPAD = True
If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then checkDIRPAD = True
If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then checkDIRPAD = True
If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then checkDIRPAD = True


End Function


Public Sub Check_Keys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static lastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            Call AddMovementToKeysMovementPressedQueue
            
            'Move Up
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyUp) Then
                Call MoveTo(NORTH)
                Exit Sub
            End If
            
            'Move Right
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyRight) Then
                Call MoveTo(EAST)
                Exit Sub
            End If
        
            'Move down
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyDown) Then
                Call MoveTo(SOUTH)
                Exit Sub
            End If
        
            'Move left
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyLeft) Then
                Call MoveTo(WEST)
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            End If

        End If
    End If
End Sub

Sub Load_Map(ByVal Map As Integer)
    Audio.StopWave

    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    
    handle = FreeFile()
    
    Open DirMapas & "Mapa" & Map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MapInfo.Zone 'Zone
    Get handle, , MapInfo.Terrain 'Terrain
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For Y = MinMapSize To MaxMapSize
        For X = MinMapSize To MaxMapSize
            Get handle, , ByFlags
            
            'MapData(x, y).blocked = (ByFlags And 1)
            
            If ByFlags And 1 Then
                Get handle, , MapData(X, Y).Blocked
            Else
                MapData(X, Y).Blocked = 0
            End If
            
            Get handle, , MapData(X, Y).Layer(1).GrhIndex
            InitGrh MapData(X, Y).Layer(1), MapData(X, Y).Layer(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Layer(2).GrhIndex
                InitGrh MapData(X, Y).Layer(2), MapData(X, Y).Layer(2).GrhIndex
            Else
                MapData(X, Y).Layer(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Layer(3).GrhIndex
                InitGrh MapData(X, Y).Layer(3), MapData(X, Y).Layer(3).GrhIndex
            Else
                MapData(X, Y).Layer(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Layer(4).GrhIndex
                InitGrh MapData(X, Y).Layer(4), MapData(X, Y).Layer(4).GrhIndex
            Else
                MapData(X, Y).Layer(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            'Erase Particles
            'If MapData(X, Y).particle_group_index > 0 Then
                'engine.Particle_Group_Remove (MapData(X, Y).particle_group_index)
                'MapData(X, Y).particle_group_index = 0
            'End If
            
            If ByFlags And 32 Then
                Get handle, , MapData(X, Y).particle_group_index
                General_Particle_Create MapData(X, Y).particle_group_index, X, Y, -1
            Else
                engine.Particle_Group_Remove (MapData(X, Y).particle_group_index)
                MapData(X, Y).particle_group_index = 0
            End If
            
            'Erase NPCs
            If MapData(X, Y).CharIndex > 0 Then
                Call EraseChar(MapData(X, Y).CharIndex)
            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
            
            'Erase Huellas
            MapData(X, Y).Huella.GrhIndex = 0
        Next X
    Next Y
    
    Close handle
    
    MapInfo.name = ""
    MapInfo.Music = ""
    CurMap = Map
    
    If bRain Then
        engine.Particle_Group_Remove (particle_group_index_render(1))
        If Not MapInfo.Zone = 1 Then
            Select Case MapInfo.Terrain
                Case 1
                    General_Particle_Create_Render 57, 1, -1
                    frmMain.tTrueno.Enabled = False
        
                Case 2
                    General_Particle_Create_Render 59, 1, -1
                    frmMain.tTrueno.Enabled = False
        
                Case Else
                    General_Particle_Create_Render 58, 1, -1
                   frmMain.tTrueno.Enabled = True
                   SoundRainIndex = Audio.PlayWave(SND_LLUVIA, , , Enabled)
            End Select
        End If
    End If
    
    If MapInfo.Zone = 2 Then '¿la zona es una ciudad?
        Select Case Clima(Hour(Time)).WhatIsClime
            Case "MAÑANA"
                Call Audio.PlayWave(SND_DAYCITY, , , Enabled)
            Case "TARDE"
                Call Audio.PlayWave(SND_DAYCITY, , , Enabled)
            Case "NOCHE"
                Call Audio.PlayWave(SND_EVENINGCITY, , , Enabled)
        End Select
    End If
    
    Call DrawMiniMap

End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = Mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = Mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, Value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(Mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Layer(1).GrhIndex >= 21781 And MapData(X, Y).Layer(1).GrhIndex <= 21796) Or _
                (MapData(X, Y).Layer(1).GrhIndex >= 21797 And MapData(X, Y).Layer(1).GrhIndex <= 21812)) And _
                    MapData(X, Y).Layer(2).GrhIndex = 0
                
End Function
Private Sub Init_Names()
    Ciudades(eCiudad.cOnirem) = "Onirem (Imperial)"
    Ciudades(eCiudad.cOsurac) = "Osurac (Neutral)"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    
    Call RemoveAllDialogs
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    
    '// HIJO DE MIL PUTA atte: WINDOWS 7 PARA ARRIBA
    'Call Resolution.ResetResolution
    
    'Stop tile engine
    Call engine.Engine_DeInit
    
    'Destruimos los objetos públicos creados
    Set FormParser = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    End
End Sub
Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function
Public Function getStrenghtColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getStrenghtColor = RGB(255 - (M * UserFuerza), (M * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getDexterityColor = RGB(255, M * UserAgilidad, 0)
End Function
Public Function getCharIndexByName(ByVal name As String) As Integer
Dim i As Long
For i = 1 To LastChar
    If charlist(i).Nombre = name Then
        getCharIndexByName = i
        Exit Function
    End If
Next i
End Function
Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function
Sub CargarParticles()
    Dim N As Integer '// abrir archivo.
    Dim StreamFile As String
    Dim loopC As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    
    StreamFile = IniPath & "Particles.ind"
    N = FreeFile()
    Open StreamFile For Binary Access Read As #N
    
        Get #N, , MiCabecera
        Get #N, , TotalStreams
    
        'resize StreamData array
        ReDim StreamData(1 To TotalStreams) As Stream
    
        'fill StreamData array with info from Particles.ini
        For loopC = 1 To TotalStreams
            Get #N, , StreamData(loopC).name
            Get #N, , StreamData(loopC).NumOfParticles
            Get #N, , StreamData(loopC).x1
            Get #N, , StreamData(loopC).y1
            Get #N, , StreamData(loopC).x2
            Get #N, , StreamData(loopC).y2
            Get #N, , StreamData(loopC).angle
            Get #N, , StreamData(loopC).vecx1
            Get #N, , StreamData(loopC).vecx2
            Get #N, , StreamData(loopC).vecy1
            Get #N, , StreamData(loopC).vecy2
            Get #N, , StreamData(loopC).life1
            Get #N, , StreamData(loopC).life2
            Get #N, , StreamData(loopC).friction
            Get #N, , StreamData(loopC).spin
            Get #N, , StreamData(loopC).spin_speedL
            Get #N, , StreamData(loopC).spin_speedH
            Get #N, , StreamData(loopC).AlphaBlend
            Get #N, , StreamData(loopC).gravity
            Get #N, , StreamData(loopC).grav_strength
            Get #N, , StreamData(loopC).bounce_strength
            Get #N, , StreamData(loopC).XMove
            Get #N, , StreamData(loopC).YMove
            Get #N, , StreamData(loopC).move_x1
            Get #N, , StreamData(loopC).move_x2
            Get #N, , StreamData(loopC).move_y1
            Get #N, , StreamData(loopC).move_y2
            Get #N, , StreamData(loopC).life_counter
            Get #N, , StreamData(loopC).Speed
            'Get #N, , StreamData(loopc).grh_resize
            'Get #N, , StreamData(loopc).grh_resizex
            'Get #N, , StreamData(loopc).grh_resizey
            Get #N, , StreamData(loopC).NumGrhs
        
            ReDim StreamData(loopC).grh_list(1 To StreamData(loopC).NumGrhs)
            Get #N, , GrhListing
       
            For i = 1 To StreamData(loopC).NumGrhs
                Get #N, , StreamData(loopC).grh_list(i)
            Next i
            
            StreamData(loopC).grh_list(i - 1) = StreamData(loopC).grh_list(i - 1)
            
            For ColorSet = 1 To 4
                Get #N, , TempSet
                'Get #N, , StreamData(loopC).colortintR
                'Get #N, , StreamData(loopC).colortintG
                ''Get #N, , StreamData(loopC).colortintB
            Next ColorSet
            
        Next loopC
    Close #N
End Sub
Sub CargarParticulas()
Dim StreamFile As String
Dim loopC As Long
Dim i As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
   
StreamFile = IniPath & "Particles.ini"
TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
 
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For loopC = 1 To TotalStreams
        StreamData(loopC).name = General_Var_Get(StreamFile, Val(loopC), "Name")
        StreamData(loopC).NumOfParticles = General_Var_Get(StreamFile, Val(loopC), "NumOfParticles")
        StreamData(loopC).x1 = General_Var_Get(StreamFile, Val(loopC), "X1")
        StreamData(loopC).y1 = General_Var_Get(StreamFile, Val(loopC), "Y1")
        StreamData(loopC).x2 = General_Var_Get(StreamFile, Val(loopC), "X2")
        StreamData(loopC).y2 = General_Var_Get(StreamFile, Val(loopC), "Y2")
        StreamData(loopC).angle = General_Var_Get(StreamFile, Val(loopC), "Angle")
        StreamData(loopC).vecx1 = General_Var_Get(StreamFile, Val(loopC), "VecX1")
        StreamData(loopC).vecx2 = General_Var_Get(StreamFile, Val(loopC), "VecX2")
        StreamData(loopC).vecy1 = General_Var_Get(StreamFile, Val(loopC), "VecY1")
        StreamData(loopC).vecy2 = General_Var_Get(StreamFile, Val(loopC), "VecY2")
        StreamData(loopC).life1 = General_Var_Get(StreamFile, Val(loopC), "Life1")
        StreamData(loopC).life2 = General_Var_Get(StreamFile, Val(loopC), "Life2")
        StreamData(loopC).friction = General_Var_Get(StreamFile, Val(loopC), "Friction")
        StreamData(loopC).spin = General_Var_Get(StreamFile, Val(loopC), "Spin")
        StreamData(loopC).spin_speedL = General_Var_Get(StreamFile, Val(loopC), "Spin_SpeedL")
        StreamData(loopC).spin_speedH = General_Var_Get(StreamFile, Val(loopC), "Spin_SpeedH")
        StreamData(loopC).AlphaBlend = General_Var_Get(StreamFile, Val(loopC), "AlphaBlend")
        StreamData(loopC).gravity = General_Var_Get(StreamFile, Val(loopC), "Gravity")
        StreamData(loopC).grav_strength = General_Var_Get(StreamFile, Val(loopC), "Grav_Strength")
        StreamData(loopC).bounce_strength = General_Var_Get(StreamFile, Val(loopC), "Bounce_Strength")
        StreamData(loopC).XMove = General_Var_Get(StreamFile, Val(loopC), "XMove")
        StreamData(loopC).YMove = General_Var_Get(StreamFile, Val(loopC), "YMove")
        StreamData(loopC).move_x1 = General_Var_Get(StreamFile, Val(loopC), "move_x1")
        StreamData(loopC).move_x2 = General_Var_Get(StreamFile, Val(loopC), "move_x2")
        StreamData(loopC).move_y1 = General_Var_Get(StreamFile, Val(loopC), "move_y1")
        StreamData(loopC).move_y2 = General_Var_Get(StreamFile, Val(loopC), "move_y2")
        StreamData(loopC).life_counter = General_Var_Get(StreamFile, Val(loopC), "life_counter")
        StreamData(loopC).Speed = Val(General_Var_Get(StreamFile, Val(loopC), "Speed"))
        StreamData(loopC).grh_resize = Val(General_Var_Get(StreamFile, Val(loopC), "resize"))
        StreamData(loopC).grh_resizex = Val(General_Var_Get(StreamFile, Val(loopC), "rx"))
        StreamData(loopC).grh_resizey = Val(General_Var_Get(StreamFile, Val(loopC), "ry"))
        StreamData(loopC).NumGrhs = General_Var_Get(StreamFile, Val(loopC), "NumGrhs")
       
        ReDim StreamData(loopC).grh_list(1 To StreamData(loopC).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopC), "Grh_List")
       
        For i = 1 To StreamData(loopC).NumGrhs
            StreamData(loopC).grh_list(i) = General_Field_Read(str(i), GrhListing, 44)
        Next i
        StreamData(loopC).grh_list(i - 1) = StreamData(loopC).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = GetVar(StreamFile, Val(loopC), "ColorSet" & ColorSet)
            StreamData(loopC).colortint(ColorSet - 1).r = ReadField(1, TempSet, 44)
            StreamData(loopC).colortint(ColorSet - 1).g = ReadField(2, TempSet, 44)
            StreamData(loopC).colortint(ColorSet - 1).b = ReadField(3, TempSet, 44)
        Next ColorSet
    Next loopC
 
End Sub
 
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long
   
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)
 
General_Particle_Create = engine.Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)
 
End Function

Public Function General_Particle_Create_Render(ByVal ParticulaInd As Long, ByVal AutomaticScaleResolution As Byte, Optional ByVal particle_life As Long = -1) As Long
   
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)
 
General_Particle_Create_Render = engine.Particle_Group_Create_Render(StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey, AutomaticScaleResolution)
 
End Function
 
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long
 
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)
 
General_Char_Particle_Create = engine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)
 
End Function
Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
   
    szReturn = ""
   
    sSpaces = Space$(5000)
   
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
   
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
   
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(Mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = Mid$(Text, LastPos + 1)
    End If
End Function
Public Sub RemoveDialog(ByVal CharIndex As Integer)
If charlist(CharIndex).dialog_life > 0 Then charlist(CharIndex).dialog = ""
charlist(CharIndex).dialog_life = 0
charlist(CharIndex).dialog_offset_counter_y = 0
End Sub

Public Sub RemoveAllDialogs()
Dim i As Long
For i = 1 To LastChar
    If charlist(i).dialog <> "" Then
        engine.Char_Dialog_Set i, "", 0, 0
    End If
Next i
End Sub

Public Sub RemoveDialogsNPCArea()
'El valor X 8 es el minXBorder y el 6 es el minYBorder
Dim PosX As Byte, PosY As Byte
For PosX = charlist(UserCharIndex).Pos.X - 8 To charlist(UserCharIndex).Pos.X + 8
    For PosY = charlist(UserCharIndex).Pos.Y - 6 To charlist(UserCharIndex).Pos.Y + 6
        If MapData(PosX, PosY).CharIndex > 0 Then _
            If Len(charlist(MapData(PosX, PosY).CharIndex).Nombre) <= 1 Then _
                Call RemoveDialog(MapData(PosX, PosY).CharIndex)
    Next PosY
Next PosX
End Sub

Private Sub Timers_Init()

    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.HUD_CLICK, INT_HUD)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.HUD_CLICK)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
End Sub
Public Sub LogError(ByVal Text As String)
Dim file As Integer
Dim str  As String

    '// Fecha, Tiempo  y Error
    str = "[" & Date$ & "/" & Time$ & "] " & Text
    
    file = FreeFile
    Open App.Path & "\errores.log" For Append Shared As #file
        Print #file, str
    Close #file
End Sub
Public Sub ChangeScaleAllForms()
    frmScaleWidth = Config_Inicio.ResolutionX * M_SCALE_FRM
    frmScaleHeight = Config_Inicio.ResolutionY * M_SCALE_FRM
End Sub
