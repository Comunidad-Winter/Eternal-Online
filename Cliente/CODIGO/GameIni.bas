Attribute VB_Name = "GameIni"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Type tGameIni
    ResolutionX As Long
    ResolutionY As Long
    AccountName As String '// por seguridad guardamos unicamente el nombre.
    FullScreen As Boolean '¿Estoy en pantalla completa?
    Sounds As Boolean
    Music As Boolean
    SoundVolume As Byte
    MusicVolume As Byte
    CursorGraphic As Boolean
    VSYNC As Byte
End Type

Public Config_Inicio As tGameIni
Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub
Private Sub Load_EngineColors()
    LongWhite(0) = D3DColorXRGB(255, 255, 255)
    LongWhite(1) = LongWhite(0)
    LongWhite(2) = LongWhite(0)
    LongWhite(3) = LongWhite(0)
    
    LongYellow(0) = D3DColorXRGB(255, 255, 0)
    LongYellow(1) = LongYellow(0)
    LongYellow(2) = LongYellow(0)
    LongYellow(3) = LongYellow(0)
    
    LongBlue(0) = D3DColorXRGB(0, 0, 255)
    LongBlue(1) = LongBlue(0)
    LongBlue(2) = LongBlue(0)
    LongBlue(3) = LongBlue(0)
    
    LongGreen(0) = D3DColorXRGB(0, 255, 0)
    LongGreen(1) = LongGreen(0)
    LongGreen(2) = LongGreen(0)
    LongGreen(3) = LongGreen(0)
    
    LongBlack(0) = D3DColorARGB(200, 0, 0, 0)
    LongBlack(1) = LongBlack(0)
    LongBlack(2) = LongBlack(0)
    LongBlack(3) = LongBlack(0)
End Sub

Public Sub LoadConfigINI()
'===================================================
' Load principal config.
' Autor: ZenitraM
' Last Modification: 25/06/2020
'===================================================
'Dim N As Integer

    InitCommonControls '// Apariencia de Windows.
    Load_EngineColors

    'Principal Variables.
    IniPath = App.Path & "\resources\Init\"
    
    'N = FreeFile
    'Open IniPath & "Config.ind" For Binary As #N
    '    Get #N, , MiCabecera
    '    Get #N, , Config_Inicio
    'Close #N
    With Config_Inicio
        .AccountName = GetVar(IniPath & "Config.ini", "GameCFG", "AccountName")
        .CursorGraphic = GetVar(IniPath & "Config.ini", "GameCFG", "CursorGraphic")
        .ResolutionX = GetVar(IniPath & "Config.ini", "GameCFG", "ResolutionX")
        .ResolutionY = GetVar(IniPath & "Config.ini", "GameCFG", "ResolutionY")
        .FullScreen = GetVar(IniPath & "Config.ini", "GameCFG", "FullScreen")
        .Sounds = GetVar(IniPath & "Config.ini", "GameCFG", "Sounds")
        .Music = GetVar(IniPath & "Config.ini", "GameCFG", "Music")
        .SoundVolume = GetVar(IniPath & "Config.ini", "GameCFG", "SoundVolume")
        .MusicVolume = GetVar(IniPath & "Config.ini", "GameCFG", "MusicVolume")
        .VSYNC = GetVar(IniPath & "Config.ini", "GameCFG", "VSYNC")
    End With
    
    Set FormParser = New clsCursor '// Cursor Grafico.
    
    If Config_Inicio.CursorGraphic Then
        Call FormParser.Init
    End If
End Sub
