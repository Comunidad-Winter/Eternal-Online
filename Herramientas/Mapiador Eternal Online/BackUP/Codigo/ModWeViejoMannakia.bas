Attribute VB_Name = "ModWeViejoMannakia"
'COMO EL GIL DE LUCAS BORRO FUNCIONES QUE NADA QUE VER PONGO ACA LAS QUE SIRVEN
Public Sub BlockearSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    Dim Vacio As MapBlock
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             If MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 1 Then
                MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 0
             Else
                MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 1
            End If
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CortarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    CopiarSeleccion
    Dim X As Integer
    Dim Y As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SeleccionIX, Y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CopiarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim Y As Integer
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            SeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
End Sub
Public Sub DePegar()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
             MapData(X + DeSeleccionOX, Y + DeSeleccionOY) = DeSeleccionMap(X, Y)
        Next
    Next
End Sub
Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim Y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SobreX, Y + SobreY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SobreX, Y + SobreY) = SeleccionMap(X, Y)
        Next
    Next
    Seleccionando = False
End Sub
Public Sub AccionSeleccion()
On Error Resume Next
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + Y
        Next
    Next
    Seleccionando = False
End Sub
Sub DrawGrhtoHdc(picX As PictureBox, Grh As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim destRect As RECT
    
    destRect.Bottom = picX.ScaleHeight
    destRect.Right = picX.ScaleWidth
    destRect.Left = 0
    destRect.Top = 0
    
    D3DDevice.BeginScene
    'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        Engine.Draw_GrhIndex Grh, X, Y
    D3DDevice.EndScene
    D3DDevice.Present destRect, ByVal 0, picX.hwnd, ByVal 0
End Sub



Function InMapBounds(X As Integer, Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

'?ESTO ES MIO
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = Engine.Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).X1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).X2, _
    StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, False, 0, 0)
    

End Function

Public Sub GenerarVista()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
   ' hacer una llamada a un seter o geter , es mas lento q una variable
   ' con esto hacemos q no este preguntando a el objeto cadavez
   ' q dibuja , Render mas rapido ;)
    VerBlockeados = frmMain.cVerBloqueos.value
    VerTriggers = frmMain.cVerTriggers.value
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerParticle = frmMain.mnuParticle.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    
End Sub

