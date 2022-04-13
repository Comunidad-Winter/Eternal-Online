Attribute VB_Name = "modScaleResolution"
'================================================================================================
'Eternal Online v1.0 - Based in Argentum Online v13.0 of No-Land Studios.
'Autor: ZenitraM
'Aplicate and chance resolution.
'Comments: Unifique todo aca, es decir, sistema de HUD, Cambios de resolucion, renders, etc.
'================================================================================================

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

'Resolucion de formularios ======================================================================
Public frmScaleWidth As Long
Public frmScaleHeight As Long
Public Const M_SCALE_FRM As Byte = 15
'================================================================================================
'System of console in render
Public OffSetConsola As Byte
Public UltimaLineavisible As Boolean
Public Const MaxLineas As Byte = 7

Type TConsola
    t As String
    R As Byte
    G As Byte
    B As Byte
End Type

Public Con(1 To MaxLineas) As TConsola
'================================================================================================
Public UserWritting As Boolean 'esta escribiendo?
Public ChatBuffer As String 'texto q escribe
'================================================================================================
'System of HUD
Public PosHUDX As Integer
Public PosHUDY As Integer
'Graficos (ASI IDENTIFICO EN CADA LLAMADA CUAL ES CUAL)
Private Const GRH_HUD As Long = 21632
Private Const GRH_BARRA_EXP As Long = 21633
Private Const GRH_CONNECT As Long = 21813
Private Const GRH_MINIMAP As Long = 21815
Private Const GRH_INVENTARIOS As Long = 21816
Private Const GRH_STATS As Long = 21817
Private Const GRH_E_HP As Long = 21822
Private Const GRH_E_MP As Long = 21823
Private Const GRH_BARRA_HAMBRE As Long = 21819
Private Const GRH_BARRA_SED As Long = 21820
Private Const GRH_BARRA_ENERGIA As Long = 21821
'================================================================================================

Public Sub ChangeScaleAllForms()
    frmScaleWidth = Config_Inicio.ResolutionX * M_SCALE_FRM
    frmScaleHeight = Config_Inicio.ResolutionY * M_SCALE_FRM
End Sub
Private Sub Render_UserStats()
Dim line As String
Dim lineExp As String
Dim linePJ As String

    'Dibujamos la esfera vacia de vida.
    Call Draw_GrhIndex(GRH_E_HP, PosHUDX - 220, PosHUDY, 1, 85, (86 - (((UserMinHP / 100) / (UserMaxHP / 100)) * 86)))
    
    'Dibujamos la esfera vacia de mana.
    If Not UserMaxMAN = 0 Then '// Si sos guerre que dibuje el "0/0" y no me bugee todo el engine.
        Call Draw_GrhIndex(GRH_E_MP, PosHUDX + 217, PosHUDY - 1, 1, 85, (86 - (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 86)))
    Else
        Call Draw_GrhIndex(GRH_E_MP, PosHUDX + 217, PosHUDY - 1, 1)
    End If
    
    '// HP TEXT
    line = UserMinHP & "/" & UserMaxHP
    Engine_Text_Render line, PosHUDX - 205 - Engine_Text_Width(line) / 2, PosHUDY - 20, LongWhite

    '// MP TEXT
    line = UserMinMAN & "/" & UserMaxMAN
    If Not UserMaxMAN = 0 Then
        Engine_Text_Render line, PosHUDX + 233 - Engine_Text_Width(line) / 2, PosHUDY - 20, LongWhite
    End If

    '// BARRA DE EXPERIENCIA
    If Not UserExp = 0 And Not UserPasarNivel = 0 Then
        lineExp = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel)) & "%"
        Call Draw_GrhIndex(GRH_BARRA_EXP, PosHUDX - 1, PosHUDY - 9, 1, Round(((UserExp / 100) / (UserPasarNivel / 100)) * 341), 9)
        Engine_Text_Render lineExp, PosHUDX - Engine_Text_Width(lineExp) / 2, PosHUDY + 11, LongWhite
    Else
        lineExp = "¡Nivel maximo!"
        Call Draw_GrhIndex(GRH_BARRA_EXP, PosHUDX - 1, PosHUDY - 9, 1)
        Engine_Text_Render lineExp, PosHUDX - Engine_Text_Width(lineExp) / 2, PosHUDY + 11, LongWhite
    End If

    Call Draw_GrhIndex(GRH_STATS, 105, 65, 1)
    
    Call Draw_GrhIndex(GRH_BARRA_HAMBRE, 135, -11, 1, (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 171), 11)
    Call Draw_GrhIndex(GRH_BARRA_SED, 141, 7, 1, (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 160), 11)
    Call Draw_GrhIndex(GRH_BARRA_ENERGIA, 136, 34, 1, (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 169), 11)
    
    linePJ = UserMinHAM & "/" & UserMaxHAM
    Engine_Text_Render linePJ, 153 - Engine_Text_Width(linePJ) / 2, 8, LongWhite
    
    linePJ = UserMinAGU & "/" & UserMaxAGU
    Engine_Text_Render linePJ, 153 - Engine_Text_Width(linePJ) / 2, 26, LongWhite
    
    linePJ = UserMinSTA & "/" & UserMaxSTA
    Engine_Text_Render linePJ, 153 - Engine_Text_Width(linePJ) / 2, 53, LongWhite
    
    linePJ = UserGLD
    Engine_Text_Render linePJ, 160 - Engine_Text_Width(linePJ) / 2, 74, LongYellow
    
    linePJ = UserFuerza
    Engine_Text_Render linePJ, 28 - Engine_Text_Width(linePJ) / 2, 76, LongGreen
    
    linePJ = UserAgilidad
    Engine_Text_Render linePJ, 51 - Engine_Text_Width(linePJ) / 2, 76, LongYellow


End Sub
Public Sub Render_HUD()
Dim i As Byte, LongRed(3) As Long

    LongRed(0) = D3DColorXRGB(255, 0, 0)
    LongRed(1) = LongRed(0)
    LongRed(2) = LongRed(0)
    LongRed(3) = LongRed(0)
    
    Render_Console
    Render_Chat
    
    Call Draw_GrhIndex(GRH_HUD, PosHUDX, PosHUDY, 1)
    Call Draw_GrhIndex(GRH_MINIMAP, Config_Inicio.ResolutionX - 70, 73, 1)
    Call Draw_GrhIndex(GRH_INVENTARIOS, Config_Inicio.ResolutionX - 99, Config_Inicio.ResolutionY - 34, 1)
    
    Render_UserStats
    
    'Call Draw_GrhIndex(23655, Config_Inicio.ResolutionX - 102, Config_Inicio.ResolutionY - 32, 1, lighthandle) '//Inv
    Call Draw_GrhIndex(Clima(Hour(Time)).GRH_CLIMA, PosHUDX + 110, PosHUDY - 25, 1)
    
    For i = 1 To 6
        If MacroList(i).mTipe = 0 Or MacroList(i).Grh <= 0 Then
            'Call Draw_GrhIndex(1, frmMain.Macros(i).Left, PosHUDY - 25, 1, lighthandle)
        Else
            Call Draw_GrhIndex(MacroList(i).Grh, frmMain.Macros(i).Left, PosHUDY - 25, 1)
        End If
        
        If MacroList(i).mTipe = eMacros.aUsar Or MacroList(i).mTipe = eMacros.aEquipar Then
            Engine_Text_Render CStr(TotalItemAmountGet(MacroList(i).ObjIndex)), frmMain.Macros(i).Left + 1, frmMain.Macros(i).Top + 20, LongWhite
        End If
        
        If ObjIndexEquipped(MacroList(i).ObjIndex) Then
           Engine_Text_Render "+", frmMain.Macros(i).Left + 22, frmMain.Macros(i).Top, LongRed
        End If
    Next i
    
End Sub
Sub Render_Console()
Dim i As Byte
Dim temp_array_console(3) As Long
If OffSetConsola > 0 Then OffSetConsola = OffSetConsola - 1
If OffSetConsola = 0 Then UltimaLineavisible = True
 
For i = 1 To MaxLineas - 1
    temp_array_console(0) = D3DColorXRGB(Con(i).R, Con(i).G, Con(i).B)
    temp_array_console(1) = temp_array_console(0)
    temp_array_console(2) = temp_array_console(0)
    temp_array_console(3) = temp_array_console(0)
    Engine_Text_Render Con(i).t, 0, (Config_Inicio.ResolutionY - 235) + (i * 15) + OffSetConsola, temp_array_console
Next i

If UltimaLineavisible = True Then
    temp_array_console(0) = D3DColorXRGB(Con(MaxLineas).R, Con(MaxLineas).G, Con(MaxLineas).B)
    temp_array_console(1) = temp_array_console(0)
    temp_array_console(2) = temp_array_console(0)
    temp_array_console(3) = temp_array_console(0)
    Engine_Text_Render Con(MaxLineas).t, 0, (Config_Inicio.ResolutionY - 235) + (MaxLineas * 15) + OffSetConsola, temp_array_console
End If

End Sub
Public Sub Render_Chat()
    If UserWritting Then
        Engine_Text_Render "Chat: " & ChatBuffer, 0, Config_Inicio.ResolutionY - 115, LongWhite
    End If
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
Public Sub Chat_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
    ChatBuffer = ChatBuffer + ChrW$(KeyAscii)
End Sub
Public Sub Chat_DestroyAll()
    ChatBuffer = vbNullString
    UserWritting = False
End Sub
Public Sub Render_Connect()
'===============================
'Conectar renderizado
'===============================
Dim View_X As Long
Dim View_Y As Long
Dim AmbientClima(3) As Long
Dim MaxViewTilesX As Byte
Dim MaxViewTilesY As Byte
Dim TotalLayers As Byte
Dim GRH_CONNECT_POS_X As Long
Dim GRH_CONNECT_POS_Y As Long

GRH_CONNECT_POS_X = Config_Inicio.ResolutionX / 2
GRH_CONNECT_POS_Y = Config_Inicio.ResolutionY / 2 + 158

'Max view tiles in render.
MaxViewTilesX = (Config_Inicio.ResolutionX / 32) + 1
MaxViewTilesY = (Config_Inicio.ResolutionY / 32) + 1
'Day
AmbientClima(0) = D3DColorXRGB(255, 255, 255)
AmbientClima(1) = AmbientClima(0)
AmbientClima(2) = AmbientClima(0)
AmbientClima(3) = AmbientClima(0)

    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene

    'Draw all layers
    For TotalLayers = 1 To 4
        For View_X = 1 To MaxViewTilesX
            For View_Y = 1 To MaxViewTilesY
                With MapData(View_X + 10, View_Y + 10)
                    If .Layer(TotalLayers).GrhIndex <> 0 Then
                        Call Draw_Grh(.Layer(TotalLayers), _
                            (View_X - 1) * 32, (View_Y - 1) * 32, _
                            1, 1, AmbientClima)
                    End If
                End With
            Next View_Y
        Next View_X
    Next TotalLayers

    
    Call Draw_GrhIndex(GRH_CONNECT, GRH_CONNECT_POS_X, GRH_CONNECT_POS_Y, 1)
    Call Draw_GrhIndex(21818, Config_Inicio.ResolutionX / 2, GRH_CONNECT_POS_Y - 400, 1) 'Eternal Online

    Engine_Text_Render Fps & " FPS", Config_Inicio.ResolutionX - 50, 0, LongWhite
    
    DirectDevice.EndScene
    DirectDevice.Present ByVal 0, ByVal 0, frmConnect.picRender.hwnd, ByVal 0

End Sub

