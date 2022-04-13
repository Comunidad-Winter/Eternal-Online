Attribute VB_Name = "modMapIO"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modMapIO
'
' @remarks Funciones Especificas al trabajo con Archivos de Mapas
' @author gshaxor@gmail.com
' @version 0.1.15
' @date 20060602

Option Explicit
Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Obtener el tama�o de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tama�o

Public Function FileSize(ByVal FileName As String) As Long
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    On Error GoTo FalloFile
    Dim nFileNum As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open FileName For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1
End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
If LenB(Dir(file, FileType)) = 0 Then
    FileExist = False
Else
    FileExist = True
End If

End Function

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByRef path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************

Call MapaV2_Cargar(path)

End Sub

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************

frmMain.Dialog.CancelError = True
On Error GoTo ErrHandler

If LenB(path) = 0 Then
    frmMain.ObtenerNombreArchivo True
    path = frmMain.Dialog.FileName
    If LenB(path) = 0 Then Exit Sub
End If

Call MapaV2_Guardar(path)

ErrHandler:
End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If MapInfo.Changed = 1 Then
    If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
        GuardarMapa path
    End If
End If
End Sub


''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 21/05/06
'*************************************************

On Error Resume Next

Dim loopc As Integer
Dim Y As Integer
Dim X As Integer

bAutoGuardarMapaCount = 0

'frmMain.mnuUtirialNuevoFormato.Checked = True
frmMain.mnuReAbrirMapa.Enabled = False
frmMain.TimAutoGuardarMapa.Enabled = False
frmMain.lblMapVersion.Caption = 0

MapaCargado = False

For loopc = 0 To frmMain.MapPest.Count - 1
    frmMain.MapPest(loopc).Enabled = False
Next

frmMain.MousePointer = 11



    Dim valorMin As Integer
    Dim valorMax As Integer
    valorMin = 1
    valorMax = 100

For Y = valorMin To valorMax
    For X = valorMin To valorMax
    
        ' Capa 1
        MapData(X, Y).Graphic(1).grhindex = 0
        
        ' Bloqueos
        MapData(X, Y).Blocked = 0

        ' Capas 2, 3 y 4
        MapData(X, Y).Graphic(2).grhindex = 0
        MapData(X, Y).Graphic(3).grhindex = 0
        MapData(X, Y).Graphic(4).grhindex = 0

        ' NPCs
        If MapData(X, Y).NPCIndex > 0 Then
            EraseChar MapData(X, Y).CharIndex
            MapData(X, Y).NPCIndex = 0
        End If

        ' OBJs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.grhindex = 0

        ' Translados
        MapData(X, Y).TileExit.map = 0
        MapData(X, Y).TileExit.X = 0
        MapData(X, Y).TileExit.Y = 0
        
        ' Triggers
        MapData(X, Y).Trigger = 0
        
        InitGrh MapData(X, Y).Graphic(1), 0
    Next X
Next Y

MapInfo.MapVersion = 0
MapInfo.name = "Nuevo Mapa"
MapInfo.Music = 0
MapInfo.PK = True
MapInfo.MagiaSinEfecto = 0
MapInfo.InviSinEfecto = 0
MapInfo.ResuSinEfecto = 0
MapInfo.Terreno = "BOSQUE"
MapInfo.Zona = "CAMPO"
MapInfo.Restringir = "No"
MapInfo.NoEncriptarMP = 0
MapInfo.ELVMIN = 0
MapInfo.ELVMAX = 50

Call MapInfo_Actualizar

bRefreshRadar = True ' Radar

'Set changed flag
MapInfo.Changed = 0
frmMain.MousePointer = 0

' Vacio deshacer
modEdicion.Deshacer_Clear

MapaCargado = True
EngineRun = True

frmMain.SetFocus

End Sub

''
' Guardar Mapa con el formato V2
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV2_Guardar(ByVal SaveAs As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo ErrorSave
Dim FreeFileMap As Long
Dim FreeFileInf As Long
Dim loopc As Long
Dim TempInt As Integer
Dim Y As Long
Dim X As Long
Dim ByFlags As Byte
Dim ZoneInt As Integer
Dim TerrainInt As Integer

Select Case MapInfo.Zona
    Case "DUNGEON"
        ZoneInt = 1
    Case "CIUDAD"
        ZoneInt = 2
    Case Else
        ZoneInt = 0
End Select

Select Case MapInfo.Terreno
    Case "NIEVE"
        TerrainInt = 1
    Case "DESIERTO"
        TerrainInt = 2
    Case "BOSQUE"
        TerrainInt = 3
    Case Else
        TerrainInt = 0
End Select


If FileExist(SaveAs, vbNormal) = True Then
    If MsgBox("�Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    Else
        Kill SaveAs
    End If
End If

frmMain.MousePointer = 11

' y borramos el .inf tambien
If FileExist(Left(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
    Kill Left(SaveAs, Len(SaveAs) - 4) & ".inf"
End If

'Open .map file
FreeFileMap = FreeFile
Open SaveAs For Binary As FreeFileMap
Seek FreeFileMap, 1

SaveAs = Left(SaveAs, Len(SaveAs) - 4)
SaveAs = SaveAs & ".inf"

'Open .inf file
FreeFileInf = FreeFile
Open SaveAs For Binary As FreeFileInf
Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If
    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    
    Put FreeFileMap, , ZoneInt
    Put FreeFileMap, , TerrainInt
    
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(X, Y).Blocked Then ByFlags = ByFlags Or 1
                If MapData(X, Y).Graphic(2).grhindex Then ByFlags = ByFlags Or 2
                If MapData(X, Y).Graphic(3).grhindex Then ByFlags = ByFlags Or 4
                If MapData(X, Y).Graphic(4).grhindex Then ByFlags = ByFlags Or 8
                If MapData(X, Y).Trigger Then ByFlags = ByFlags Or 16
                If MapData(X, Y).particle_group_index Then ByFlags = ByFlags Or 32
                
                Put FreeFileMap, , ByFlags
                
                If MapData(X, Y).Blocked Then _
                    Put FreeFileMap, , MapData(X, Y).Blocked
                
                Put FreeFileMap, , MapData(X, Y).Graphic(1).grhindex
                
                For loopc = 2 To 4
                    If MapData(X, Y).Graphic(loopc).grhindex Then _
                        Put FreeFileMap, , MapData(X, Y).Graphic(loopc).grhindex
                Next loopc
                
                If MapData(X, Y).Trigger Then _
                    Put FreeFileMap, , MapData(X, Y).Trigger
                    
                If MapData(X, Y).particle_group_index Then _
                    Put FreeFileMap, , MapData(X, Y).parti_index
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(X, Y).TileExit.map Then ByFlags = ByFlags Or 1
                If MapData(X, Y).NPCIndex Then ByFlags = ByFlags Or 2
                If MapData(X, Y).OBJInfo.OBJIndex Then ByFlags = ByFlags Or 4
                
                
                Put FreeFileInf, , ByFlags
                
                If MapData(X, Y).TileExit.map Then
                    Put FreeFileInf, , MapData(X, Y).TileExit.map
                    Put FreeFileInf, , MapData(X, Y).TileExit.X
                    Put FreeFileInf, , MapData(X, Y).TileExit.Y
                End If
                
                If MapData(X, Y).NPCIndex Then
                
                    Put FreeFileInf, , CInt(MapData(X, Y).NPCIndex)
                End If
                
                If MapData(X, Y).OBJInfo.OBJIndex Then
                    Put FreeFileInf, , MapData(X, Y).OBJInfo.OBJIndex
                    Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount
                End If
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf


Call Pesta�as(SaveAs)

'write .dat file
SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
MapInfo_Guardar SaveAs

'Change mouse icon
frmMain.MousePointer = 0
MapInfo.Changed = 0

Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description
End Sub

''
' Guardar Mapa con el formato V1
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV1_Guardar(SaveAs As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc As Long
    Dim TempInt As Integer
    Dim T As String
    Dim Y As Long
    Dim X As Long
    
    If FileExist(SaveAs, vbNormal) = True Then
        If MsgBox("�Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs
        End If
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 11
    T = SaveAs
    If FileExist(Left(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill Left(SaveAs, Len(SaveAs) - 4) & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    
    SaveAs = Left(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"
    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If
    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.map file
            
            ' Bloqueos
            Put FreeFileMap, , MapData(X, Y).Blocked
            
            ' Capas
            For loopc = 1 To 4
                If loopc = 2 Then Call FixCoasts(MapData(X, Y).Graphic(loopc).grhindex, X, Y)
                Put FreeFileMap, , MapData(X, Y).Graphic(loopc).grhindex
            Next loopc
            
            ' Triggers
            Put FreeFileMap, , MapData(X, Y).Trigger
            Put FreeFileMap, , TempInt
            
            '.inf file
            'Tile exit
            Put FreeFileInf, , MapData(X, Y).TileExit.map
            Put FreeFileInf, , MapData(X, Y).TileExit.X
            Put FreeFileInf, , MapData(X, Y).TileExit.Y
            
            'NPC
            Put FreeFileInf, , MapData(X, Y).NPCIndex
            
            'Object
            Put FreeFileInf, , MapData(X, Y).OBJInfo.OBJIndex
            Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount
            
            'Empty place holders for future expansion
            Put FreeFileInf, , TempInt
            Put FreeFileInf, , TempInt
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    'Close .inf file
    Close FreeFileInf
    FreeFileMap = FreeFile
    Open T & "2" For Binary Access Write As FreeFileMap
        Put FreeFileMap, , MapData
    Close FreeFileMap
    Call Pesta�as(SaveAs)
    
    'write .dat file
    SaveAs = Left(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description
End Sub

''
' Abrir Mapa con el formato V2
'
' @param Map Especifica el Path del mapa

Public Sub MapaV2_Cargar(ByVal map As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

Engine.Particle_Group_Remove_All
Engine.Char_Clean


On Error Resume Next
    Dim loopc As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim Y As Integer
    Dim X As Integer
    Dim ByFlags As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim ZoneInt As Integer
    Dim TerrainInt As Integer
    Dim i As Byte
    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Open files
    FreeFileMap = FreeFile
    Open map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    map = Left(map, Len(map) - 4)
    map = map & ".inf"
    
    FreeFileInf = FreeFile
    Open map For Binary As FreeFileInf
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    
    Get FreeFileMap, , TempInt 'Zone
    Get FreeFileMap, , TempInt 'Terrain
    
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    
    Dim valorMin As Integer
    Dim valorMax As Integer

    valorMin = 1
    valorMax = 100
    
    
    'Load arrays
    For Y = valorMin To valorMax
        For X = valorMin To valorMax
    
            Get FreeFileMap, , ByFlags
            
            'Varios tipos de bloqueos ?
            If ByFlags And 1 Then
                Get FreeFileMap, , MapData(X, Y).Blocked
            Else
                MapData(X, Y).Blocked = 0
            End If
            
            Get FreeFileMap, , MapData(X, Y).Graphic(1).grhindex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).grhindex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(2).grhindex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).grhindex
            Else
                MapData(X, Y).Graphic(2).grhindex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(3).grhindex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).grhindex
            Else
                MapData(X, Y).Graphic(3).grhindex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(4).grhindex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).grhindex
            Else
                MapData(X, Y).Graphic(4).grhindex = 0
            End If
            
             
            'Trigger used?
            If ByFlags And 16 Then
                Get FreeFileMap, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            If ByFlags And 32 Then
                Get FreeFileMap, , TempInt
                MapData(X, Y).particle_group_index = General_Particle_Create(TempInt, X, Y, -1)
                MapData(X, Y).parti_index = TempInt
            Else
                MapData(X, Y).particle_group_index = 0
                MapData(X, Y).parti_index = 0
            End If
            
            '.inf file
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(X, Y).TileExit.map
                Get FreeFileInf, , MapData(X, Y).TileExit.X
                Get FreeFileInf, , MapData(X, Y).TileExit.Y
            End If
    
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(X, Y).NPCIndex
    
                If MapData(X, Y).NPCIndex < 0 Then
                    MapData(X, Y).NPCIndex = 0
                Else
                    Body = NpcData(MapData(X, Y).NPCIndex).Body
                    Head = NpcData(MapData(X, Y).NPCIndex).Head
                    Heading = NpcData(MapData(X, Y).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
                End If
            End If
    
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(X, Y).OBJInfo.OBJIndex
                Get FreeFileInf, , MapData(X, Y).OBJInfo.Amount
                If MapData(X, Y).OBJInfo.OBJIndex > 0 Then
                    InitGrh MapData(X, Y).ObjGrh, ObjData(MapData(X, Y).OBJInfo.OBJIndex).grhindex
                End If
            End If
    
        Next X
    Next Y
    
    Call DrawMiniMap
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
    
    Call Pesta�as(map)
    
    bRefreshRadar = True ' Radar
    
    map = Left$(map, Len(map) - 4) & ".dat"
    
    MapInfo_Cargar map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True
    
End Sub

''
' Abrir Mapa con el formato V1
'
' @param Map Especifica el Path del mapa

Public Sub MapaV1_Cargar(ByVal map As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    On Error Resume Next
    Dim TBlock As Byte
    Dim loopc As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim Y As Integer
    Dim X As Integer
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim T As String
    DoEvents
    'Change mouse icon
    frmMain.MousePointer = 11
    
    'Open files
    FreeFileMap = FreeFile
    Open map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    map = Left(map, Len(map) - 4)
    map = map & ".inf"
    FreeFileInf = FreeFile
    Open map For Binary As #2
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            '.map file
            Get FreeFileMap, , MapData(X, Y).Blocked
            
            For loopc = 1 To 4
                Get FreeFileMap, , MapData(X, Y).Graphic(loopc).grhindex
                'Set up GRH
                If MapData(X, Y).Graphic(loopc).grhindex > 0 Then
                    InitGrh MapData(X, Y).Graphic(loopc), MapData(X, Y).Graphic(loopc).grhindex
                End If
            Next loopc
            'Trigger
            Get FreeFileMap, , MapData(X, Y).Trigger
            
            Get FreeFileMap, , TempInt
            '.inf file
            
            'Tile exit
            Get FreeFileInf, , MapData(X, Y).TileExit.map
            Get FreeFileInf, , MapData(X, Y).TileExit.X
            Get FreeFileInf, , MapData(X, Y).TileExit.Y
                          
            'make NPC
            Get FreeFileInf, , MapData(X, Y).NPCIndex
            If MapData(X, Y).NPCIndex > 0 Then
                Body = NpcData(MapData(X, Y).NPCIndex).Body
                Head = NpcData(MapData(X, Y).NPCIndex).Head
                Heading = NpcData(MapData(X, Y).NPCIndex).Heading
                Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
            End If
            
            'Make obj
            Get FreeFileInf, , MapData(X, Y).OBJInfo.OBJIndex
            Get FreeFileInf, , MapData(X, Y).OBJInfo.Amount
            If MapData(X, Y).OBJInfo.OBJIndex > 0 Then
                InitGrh MapData(X, Y).ObjGrh, ObjData(MapData(X, Y).OBJInfo.OBJIndex).grhindex
            End If
            
            'Empty place holders for future expansion
            Get FreeFileInf, , TempInt
            Get FreeFileInf, , TempInt
                 
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
     
    Call Pesta�as(map)
    
    bRefreshRadar = True ' Radar
    
    map = Left(map, Len(map) - 4) & ".dat"
        
    MapInfo_Cargar map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True

End Sub


Public Sub MapaV3_Cargar(ByVal map As String)
'*************************************************
'Author: Loopzer
'Last modified: 22/11/07
'*************************************************

    On Error Resume Next
    Dim TBlock As Byte
    Dim loopc As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim Y As Integer
    Dim X As Integer
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim T As String
    DoEvents
    'Change mouse icon
    frmMain.MousePointer = 11
    
     FreeFileMap = FreeFile
    Open map For Binary Access Read As FreeFileMap
        Get FreeFileMap, , MapData
    Close FreeFileMap
    
    Call Pesta�as(map)
    
    bRefreshRadar = True ' Radar
    
    map = Left(map, Len(map) - 4) & ".dat"
        
    MapInfo_Cargar map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True

End Sub
Public Sub MapaV3_Guardar(Mapa As String)
'*************************************************
'Author: Loopzer
'Last modified: 22/11/07
'*************************************************
'copy&paste RLZ
On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc As Long
    Dim TempInt As Integer
    Dim T As String
    Dim Y As Long
    Dim X As Long
    
    If FileExist(Mapa, vbNormal) = True Then
        If MsgBox("�Desea sobrescribir " & Mapa & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill Mapa
        End If
    End If
    
    frmMain.MousePointer = 11
    
    FreeFileMap = FreeFile
    Open Mapa For Binary Access Write As FreeFileMap
        Put FreeFileMap, , MapData
    Close FreeFileMap
    Call Pesta�as(Mapa)
    
    
    Mapa = Left(Mapa, Len(Mapa) - 4) & ".dat"
    MapInfo_Guardar Mapa
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description
End Sub




' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save
    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", Str(MapInfo.BackUp))
    
    Call WriteVar(Archivo, MapTitulo, "ELVMIN", frmMapInfo.Text1.text)
    Call WriteVar(Archivo, MapTitulo, "ELVMAX", frmMapInfo.Text2.text)

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")
    End If
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    Dim Leer As New clsIniReader
    Dim loopc As Integer
    Dim path As String
    MapTitulo = Empty
    Leer.Initialize Archivo

    For loopc = Len(Archivo) To 1 Step -1
        If mid(Archivo, loopc, 1) = "\" Then
            path = Left(Archivo, loopc)
            Exit For
        End If
    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(path)))
    MapTitulo = UCase(Left(Archivo, Len(Archivo) - 4))

    MapInfo.name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
    MapInfo.ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    MapInfo.ELVMIN = Val(Leer.GetValue(MapTitulo, "ELVMIN"))
    MapInfo.ELVMAX = Val(Leer.GetValue(MapTitulo, "ELVMAX"))
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    ' Mostrar en Formularios
    frmMapInfo.txtMapNombre.text = MapInfo.name
    frmMapInfo.txtMapMusica.text = MapInfo.Music
    frmMapInfo.txtMapTerreno.text = MapInfo.Terreno
    frmMapInfo.txtMapZona.text = MapInfo.Zona
    frmMapInfo.txtMapRestringir.text = MapInfo.Restringir
    frmMapInfo.chkMapBackup.value = MapInfo.BackUp
    frmMapInfo.chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
    frmMapInfo.chkMapInviSinEfecto.value = MapInfo.InviSinEfecto
    frmMapInfo.chkMapResuSinEfecto.value = MapInfo.ResuSinEfecto
    frmMapInfo.chkMapNoEncriptarMP.value = MapInfo.NoEncriptarMP
    frmMapInfo.chkMapPK.value = IIf(MapInfo.PK = True, 1, 0)
    frmMapInfo.txtMapVersion = MapInfo.MapVersion
    frmMapInfo.Text1 = MapInfo.ELVMIN
    frmMapInfo.Text2 = MapInfo.ELVMAX
    frmMain.lblMapNombre = MapInfo.name
    frmMain.lblMapMusica = MapInfo.Music

End Sub

''
' Calcula la orden de Pesta�as
'
' @param Map Especifica path del mapa

Public Sub Pesta�as(ByVal map As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer

For loopc = Len(map) To 1 Step -1
    If mid(map, loopc, 1) = "\" Then
        PATH_Save = Left(map, loopc)
        Exit For
    End If
Next
map = Right(map, Len(map) - (Len(PATH_Save)))
For loopc = Len(Left(map, Len(map) - 4)) To 1 Step -1
    If IsNumeric(mid(Left(map, Len(map) - 4), loopc, 1)) = False Then
        NumMap_Save = Right(Left(map, Len(map) - 4), Len(Left(map, Len(map) - 4)) - loopc)
        NameMap_Save = Left(map, loopc)
        Exit For
    End If
Next
For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)
        If FileExist(PATH_Save & NameMap_Save & loopc & ".map", vbArchive) = True Then
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
        Else
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False
        End If
Next
End Sub




