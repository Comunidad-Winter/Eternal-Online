Attribute VB_Name = "TileEngine"
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

Public engine As New clsEngine

Public DirectX As New DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8
Public DispMode  As D3DDISPLAYMODE
Public D3DWindow As D3DPRESENT_PARAMETERS

Public AmbientColor As D3DCOLORVALUE

Public SurfaceDB As clsTextureManager
Public SpriteBatch As clsBatch
Public Projection As D3DMATRIX
Public View As D3DMATRIX


Public Const PI As Single = 3.14159265358979

''
'Sets a Grh animation to loop indefinitely.
Public Const INFINITE_LOOPS As Integer = -1

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

'System of console in render
Public OffSetConsola As Byte
Public UltimaLineavisible As Boolean
Public Const MaxLineas As Byte = 7

Public Type TConsola
    T As String
    a As Byte
    r As Byte
    g As Byte
    b As Byte
End Type

Public Con(1 To MaxLineas) As TConsola
'================================================================================================
Public UserWritting As Boolean 'esta escribiendo?
Public ChatBuffer As String 'texto q escribe
'================================================================================================
'System of HUD
Public PosHUDX As Integer
Public PosHUDY As Integer

Dim font_count As Long
Dim font_last As Long

Dim char_list() As Char

'Screen positioning
Public minY As Integer          'Start Y pos on current screen + tilebuffer
Public maxY As Integer          'End Y pos on current screen
Public minX As Integer          'Start X pos on current screen
Public maxX As Integer          'End X pos on current screen

'Map sizes in tiles
Public Const MaxMapSize As Byte = 100
Public Const MinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
    active As Boolean
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    Scroll_Pixels_Per_Frame As Byte
    AlphaX As Integer
    StatusAlpha As Boolean
    last_tick As Long
    active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Atacable As Boolean
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    Envenenado As Byte
    priv As Byte
    
    dialog As String
    dialog_color As Long
    dialog_life As Byte
    dialog_font_index As Integer
    dialog_offset_counter_y As Single
    dialog_scroll As Boolean
    
    particle_count As Integer
    particle_group() As Long
End Type

'Info de un objeto
Public Type OBJ
    ObjIndex As Integer
    Amount As Integer
    ObjType As Byte
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    particle_group_index As Integer
    Layer(1 To 5) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    Huella As Grh
    
    NPCIndex As Integer
    OBJInfo As OBJ
    TileExit As WorldPos
    Blocked As Byte
    
    light_value(3) As Long
    
    Color(3) As Long
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
    Zone As Integer
    Terrain As Integer
End Type

Public IniPath As String
Public MapPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public FPS As Long
Public FramesPerSecCounter As Long
Public FPSLastCheck As Long

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Tamaño de los tiles en pixels
Public Tile_Pixel_Size As Integer

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está lloviendo?
Public bTecho       As Boolean 'hay techo?

Public charlist(1 To 10000) As Char

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open IniPath & "Heads.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open IniPath & "Helmets.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open IniPath & "Bodys.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open IniPath & "FXs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub
Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef TX As Byte, ByRef TY As Byte)
On Error Resume Next
Dim TilePosClickX As Byte, TilePosClickY As Byte

    TilePosClickX = UserPos.X + viewPortX \ 32
    TilePosClickY = UserPos.Y + viewPortY \ 32

    'If TilePosClickX < HalfWindowTileWidth Then Exit Sub
    'If TilePosClickY < HalfWindowTileHeight Then Exit Sub

    TX = TilePosClickX - HalfWindowTileWidth
    TY = TilePosClickY - HalfWindowTileHeight
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .active = 1
        .AlphaX = 0
        .StatusAlpha = False
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .AlphaX = 0
        .StatusAlpha = False
        .active = 0
        .Criminal = 0
        .Atacable = False
        .FxIndex = 0
        .invisible = False
        .Envenenado = 0
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
        MapData(.Pos.X, .Pos.Y).CharIndex = 0
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
With charlist(CharIndex)
    .active = 0

    'Update lastchar
    If CharIndex = LastChar Then
        Do Until .active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(.Pos.X, .Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Call Char_Particle_Group_Remove_All(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End With
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        .Heading = nHeading
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
        If nX < 1 Or nX > 100 Or nY < 1 Or nY > 100 Then Exit Sub
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (Tile_Pixel_Size * addx)
        .MoveOffsetY = -1 * (Tile_Pixel_Size * addy)
        
        .Moving = 1
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function
Sub DoPasosFx(ByVal CharIndex As Integer)
Dim Terrain As String
    If Not UserNavegando Then
        With charlist(CharIndex)
            If .muerto = False And EstaPCarea(CharIndex) = True Then
                .pie = Not .pie
                Terrain = Map_GetTerrenoDePaso(GrhData(MapData(.Pos.X, .Pos.Y).Layer(1).GrhIndex).FileNum)
          
                Select Case Terrain
                    Case "CUALQUIERA"
                        If .pie Then
                            Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                        Else
                            Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                        End If
                        
                    Case "PASTO"
                        If .pie Then
                            Call Audio.PlayWave(SND_PASOS3, .Pos.X, .Pos.Y)
                        Else
                            Call Audio.PlayWave(SND_PASOS4, .Pos.X, .Pos.Y)
                        End If
                        
                    Case "ARENA"
                        If .pie Then
                            Call Audio.PlayWave(SND_PASOS5, .Pos.X, .Pos.Y)
                            If .Heading = NORTH Then MapData(.Pos.X, .Pos.Y + 1).Huella.GrhIndex = 22008
                            If .Heading = SOUTH Then MapData(.Pos.X, .Pos.Y - 1).Huella.GrhIndex = 22011
                            If .Heading = EAST Then MapData(.Pos.X - 1, .Pos.Y).Huella.GrhIndex = 22015
                            If .Heading = WEST Then MapData(.Pos.X + 1, .Pos.Y).Huella.GrhIndex = 22013
                        Else
                            Call Audio.PlayWave(SND_PASOS6, .Pos.X, .Pos.Y)
                            If .Heading = NORTH Then MapData(.Pos.X, .Pos.Y + 1).Huella.GrhIndex = 22009
                            If .Heading = SOUTH Then MapData(.Pos.X, .Pos.Y - 1).Huella.GrhIndex = 22010
                            If .Heading = EAST Then MapData(.Pos.X - 1, .Pos.Y).Huella.GrhIndex = 22014
                            If .Heading = WEST Then MapData(.Pos.X + 1, .Pos.Y).Huella.GrhIndex = 22012
                        End If
                    
                    Case "NIEVE"
                        If .pie Then
                            Call Audio.PlayWave(SND_PASOS7, .Pos.X, .Pos.Y)
                            If .Heading = NORTH Then MapData(.Pos.X, .Pos.Y + 1).Huella.GrhIndex = 22008
                            If .Heading = SOUTH Then MapData(.Pos.X, .Pos.Y - 1).Huella.GrhIndex = 22011
                            If .Heading = EAST Then MapData(.Pos.X - 1, .Pos.Y).Huella.GrhIndex = 22015
                            If .Heading = WEST Then MapData(.Pos.X + 1, .Pos.Y).Huella.GrhIndex = 22013
                        Else
                            Call Audio.PlayWave(SND_PASOS8, .Pos.X, .Pos.Y)
                            If .Heading = NORTH Then MapData(.Pos.X, .Pos.Y + 1).Huella.GrhIndex = 22009
                            If .Heading = SOUTH Then MapData(.Pos.X, .Pos.Y - 1).Huella.GrhIndex = 22010
                            If .Heading = EAST Then MapData(.Pos.X - 1, .Pos.Y).Huella.GrhIndex = 22014
                            If .Heading = WEST Then MapData(.Pos.X + 1, .Pos.Y).Huella.GrhIndex = 22012
                        End If
                End Select
                
            End If
        End With
    End If
End Sub
Private Function Map_GetTerrenoDePaso(ByVal TerrainFileNum As Integer) As String
    If (TerrainFileNum >= 4 And TerrainFileNum <= 16) Or (TerrainFileNum >= 20 And TerrainFileNum <= 21) Or (TerrainFileNum >= 37 And TerrainFileNum <= 50) Then
        Map_GetTerrenoDePaso = "PASTO"
        Exit Function
    ElseIf (TerrainFileNum >= 26 And TerrainFileNum <= 35) Then
        Map_GetTerrenoDePaso = "NIEVE"
        Exit Function
    ElseIf (TerrainFileNum >= 17 And TerrainFileNum <= 19) Or (TerrainFileNum >= 22 And TerrainFileNum <= 25) Then
        Map_GetTerrenoDePaso = "ARENA"
        Exit Function
    Else
        Map_GetTerrenoDePaso = "CUALQUIERA"
    End If
End Function

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (Tile_Pixel_Size * addx)
        .MoveOffsetY = -1 * (Tile_Pixel_Size * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim TX As Integer
    Dim TY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        Case E_Heading.EAST
            X = 1
        Case E_Heading.SOUTH
            Y = 1
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    TX = UserPos.X + X
    TY = UserPos.Y + Y
    
    'Check to see if its out of bounds
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = TX
        AddtoUserPos.Y = Y
        UserPos.Y = TY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim J As Long
    Dim k As Long
    
    For J = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = J
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next J
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopC As Long
    Dim Dale As Boolean
    
    loopC = 1
    Do While charlist(loopC).active And Dale
        loopC = loopC + 1
        Dale = (loopC <= UBound(charlist))
    Loop
    
    NextOpenChar = loopC
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Public Function LoadGrhData() As Boolean
On Local Error Resume Next
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    'Open files
    handle = FreeFile()
    Open IniPath & "Graphics.ind" For Binary Access Read As handle
    Get handle, , fileVersion
    
    Get handle, , grhCount
    
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        With GrhData(Grh)
           ' GrhData(Grh).active = True
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then Resume Next
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        Resume Next
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                'GrhData(Grh).speed = GrhData(Grh).speed * 2
                
                If .Speed <= 0 Then Resume Next
                
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then Resume Next
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then Resume Next
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then Resume Next
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then Resume Next
            Else
                Get handle, , .FileNum
                If .FileNum <= 0 Then Resume Next
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then Resume Next
                
                Get handle, , .sY
                If .sY < 0 Then Resume Next
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then Resume Next
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then Resume Next
                
                .TileWidth = .pixelWidth / 32
                .TileHeight = .pixelHeight / 32
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close handle

    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function
Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
        If Not charlist(CharIndex).StatusAlpha = True Then ' si ya se fue puedo pasar.
    
        If MapData(UserPos.X, UserPos.Y).Blocked > 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < MinMapSize Or X > MaxMapSize Or Y < MinMapSize Or Y > MaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal srchdc As Long, ByRef SourceRect As RECT, ByRef destRect As RECT, ByVal TransparentColor)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/22/2009
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************
    Dim Color As Long
    Dim X As Long
    Dim Y As Long
    
    For X = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.bottom
            Color = GetPixel(srchdc, X, Y)
            
            If Color <> TransparentColor Then
                Call SetPixel(dsthdc, destRect.Left + (X - SourceRect.Left), destRect.Top + (Y - SourceRect.Top), Color)
            End If
        Next Y
    Next X
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal x1 As Single, ByVal y1 As Single, Optional Width1, Optional Height1, Optional x2, Optional y2, Optional Width2, Optional Height2)
    Call PictureBox.PaintPicture(Picture, x1, y1, Width1, Height1, x2, y2, Width2, Height2)
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

Public Sub Chat_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
    ChatBuffer = ChatBuffer + ChrW$(KeyAscii)
End Sub
Public Sub Chat_DestroyAll()
    ChatBuffer = vbNullString
    UserWritting = False
End Sub
Public Sub Chat_Change(CharAscii As Integer)
Dim i As Long
Dim tempstr As String

    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    For i = 1 To Len(ChatBuffer)
        CharAscii = Asc(Mid$(ChatBuffer, i, 1))
        If CharAscii >= vbKeySpace And CharAscii <= 250 Then
            tempstr = tempstr & ChrW$(CharAscii) '// chrW$ Es mas eficiente, pero no siempre es recomendable su uso
        End If
        
        If CharAscii = vbKeyBack And Len(tempstr) > 0 Then
            tempstr = Left$(tempstr, Len(tempstr) - 1)
        End If
    Next i
    
    If tempstr <> ChatBuffer Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        ChatBuffer = tempstr
    End If
        
    stxtbuffer = ChatBuffer
End Sub
Public Sub Chat_KeyUP(KeyCode As Integer)
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        stxtbuffer = vbNullString
        ChatBuffer = vbNullString
        KeyCode = 0
        UserWritting = False
    End If
End Sub

Public Sub LoadOBJFormConnect()
    frmConnect.imgConectarse.Visible = True
    frmConnect.imgCrearPj.Visible = True
    frmConnect.imgSalir.Visible = True
    frmConnect.txtNombre.Visible = True
    frmConnect.txtPasswd.Visible = True
    frmConnect.lst_servers.Visible = True
End Sub

Public Sub UnloadOBJFormConnect()
    frmConnect.imgConectarse.Visible = False
    frmConnect.imgCrearPj.Visible = False
    frmConnect.imgSalir.Visible = False
    frmConnect.txtNombre.Visible = False
    frmConnect.txtPasswd.Visible = False
    frmConnect.lst_servers.Visible = False
    
End Sub
