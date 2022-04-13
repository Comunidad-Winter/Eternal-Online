Attribute VB_Name = "modAmbient"
Option Explicit

Public Type ValueDay
    WhatIsClime As String
    GRH_CLIMA As Long
End Type

Public DayR As Byte
Public DayG As Byte
Public DayB As Byte

Public AmbientClima(3) As Long

Public HayTrueno As Byte
Public TypeTrueno As Byte

Public Reproducir As Byte

Public Clima(0 To 24) As ValueDay

Public Sub Load_Climas()
Dim i As Byte ' // para el for.

    '// GRH CLIMA
    For i = 0 To 24
        Clima(i).GRH_CLIMA = 21608 + i
    Next i

    For i = 0 To 5
        Clima(i).WhatIsClime = "NOCHE"
    Next i
    
    For i = 6 To 11
        Clima(i).WhatIsClime = "MAÑANA"
    Next i
    
    For i = 12 To 19
        Clima(i).WhatIsClime = "TARDE"
    Next i
        
    For i = 20 To 24
        Clima(i).WhatIsClime = "NOCHE"
    Next i
    
    Select Case Clima(Hour(Time)).WhatIsClime
        Case "MAÑANA"
            If Not bRain Then
                DayR = 255
                DayG = 255
                DayB = 128
            Else
                DayR = 127
                DayG = 127
                DayB = 127
            End If
            
        Case "TARDE"
            If Not bRain Then
                DayR = 255
                DayG = 255
                DayB = 255
            Else
                DayR = 127
                DayG = 127
                DayB = 127
            End If
    
        Case "NOCHE"
            If Not bRain Then
                DayR = 40
                DayG = 40
                DayB = 40
            Else
                DayR = 30
                DayG = 30
                DayB = 30
            End If
    End Select
        
End Sub

Public Sub EffectDay()

If MapInfo.Zone = 1 Then '¿estoy en una dungeon?
    DayR = 255
    DayG = 255
    DayB = 255
End If

If bRain Then
    If HayTrueno > 0 Then HayTrueno = HayTrueno - 1 '// Tiempo de trueno
    If HayTrueno > 0 Then
        If Not MapInfo.Zone = 1 Then 'Si no estoy en un mapa dungeon no genero ningun efecto.
            If TypeTrueno = 1 Then
                DayR = 150
                DayG = 150
                DayB = 255
            Else
                DayR = 255
                DayG = 255
                DayB = 255
            End If
        End If
    End If
End If

Select Case Clima(Hour(Time)).WhatIsClime
    Case "MAÑANA"
        If Not bRain Then
            If DayR < 255 Then
                If Not DayR = 255 Then DayR = DayR + 1
            Else
                If Not DayB = 255 Then DayB = DayB - 1
            End If
            
            If DayG < 255 Then
                If Not DayG = 255 Then DayG = DayG + 1
            Else
                If Not DayG = 255 Then DayG = DayG - 1
            End If
            
            If DayB < 128 Then
                If Not DayB = 128 Then DayB = DayB + 1
            Else
                If Not DayB = 128 Then DayB = DayB - 1
            End If
        Else
            If DayR < 127 Then
                If Not DayR = 127 Then DayR = DayR + 1
            Else
                If Not DayR = 127 Then DayR = DayR - 1
            End If
            
            If DayG < 127 Then
                If Not DayG = 127 Then DayG = DayG + 1
            Else
                If Not DayG = 127 Then DayG = DayG - 1
            End If
            
            If DayB < 127 Then
                If Not DayB = 127 Then DayB = DayB + 1
            Else
                If Not DayB = 127 Then DayB = DayB - 1
            End If
        End If
            
    Case "TARDE"
        If Not bRain Then
            If DayR < 255 Then
                If Not DayR = 255 Then DayR = DayR + 1
            Else
                If Not DayR = 255 Then DayR = DayR - 1
            End If
            
            If DayG < 255 Then
                If Not DayG = 255 Then DayG = DayG + 1
            Else
                If Not DayG = 255 Then DayG = DayG - 1
            End If
            
            If DayB < 255 Then
                If Not DayB = 255 Then DayB = DayB + 1
            Else
                If Not DayB = 255 Then DayB = DayB - 1
            End If
        Else
            If DayR < 127 Then
                If Not DayR = 127 Then DayR = DayR + 1
            Else
                If Not DayR = 127 Then DayR = DayR - 1
            End If
            
            If DayG < 127 Then
                If Not DayG = 127 Then DayG = DayG + 1
            Else
                If Not DayG = 127 Then DayG = DayG - 1
            End If
            
            If DayB < 127 Then
                If Not DayB = 127 Then DayB = DayB + 1
            Else
                If Not DayB = 127 Then DayB = DayB - 1
            End If
        End If
    
    Case "NOCHE"
        If Not bRain Then
            If DayR > 40 Then
                If Not DayR = 40 Then DayR = DayR - 1
            Else
                If Not DayR = 40 Then DayR = DayR + 1
            End If
            
            If DayG > 40 Then
                If Not DayG = 40 Then DayG = DayG - 1
            Else
                If Not DayG = 40 Then DayG = DayG + 1
            End If
            
            If DayB > 40 Then
                If Not DayB = 40 Then DayB = DayB - 1
            Else
                If Not DayB = 40 Then DayB = DayB + 1
            End If
        Else
            If DayR > 30 Then
                If Not DayR = 30 Then DayR = DayR - 1
            Else
                If Not DayR = 30 Then DayR = DayR + 1
            End If
            
            If DayG > 30 Then
                If Not DayG = 30 Then DayG = DayG - 1
            Else
                If Not DayG = 30 Then DayG = DayG + 1
            End If
                
            If DayB > 30 Then
                If Not DayB = 30 Then DayB = DayB - 1
            Else
                If Not DayB = 30 Then DayB = DayB + 1
            End If
        End If
            
End Select

AmbientClima(0) = D3DColorXRGB(DayR, DayG, DayB)
AmbientClima(1) = AmbientClima(0)
AmbientClima(2) = AmbientClima(0)
AmbientClima(3) = AmbientClima(0)

Call SoundClime

AmbientColor.r = DayR
AmbientColor.g = DayG
AmbientColor.b = DayB
AmbientColor.a = 255
End Sub

Public Sub SoundClime()
    
    Select Case Hour(Time)
        Case 6
            If Reproducir = 1 Then Exit Sub '// no me reproduzcas infinitas veces.
            Call Audio.PlayWave(SND_MORNING)
            Call AddtoRichTextBox(frmMain.RecTxt, "Esta amaneciendo en las tierras de Eternal, buenos dias!", 169, 170, 108, True, True, False)
            Reproducir = 1
        Case 12
            If Reproducir = 2 Then Exit Sub
            Call Audio.PlayWave(SND_AFTERNOON)
            Call AddtoRichTextBox(frmMain.RecTxt, "Ya es de día en las tierras de Eternal, buenas tardes!", 255, 242, 0, True, True, False)
            Reproducir = 2
        Case 20
            If Reproducir = 3 Then Exit Sub
            Call Audio.PlayWave(SND_EVENING)
            Call AddtoRichTextBox(frmMain.RecTxt, "El sol ha caido y se hizo de noche, buenas noches a todos!", 30, 30, 30, True, True, False)
            Reproducir = 3
        Case Else
            Reproducir = 0 'Default
    End Select
End Sub

