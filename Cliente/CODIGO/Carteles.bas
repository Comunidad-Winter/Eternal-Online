Attribute VB_Name = "Carteles"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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

Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer


Sub InitCartel(Ley As String, Grh As Integer)
If Not Cartel Then
    Leyenda = Ley
    textura = Grh
    Cartel = True
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim i As Integer, k As Integer, anti As Integer
    anti = 1
    k = 0
    i = 0
    Call DarFormato(Leyenda, i, k, anti)
    i = 0
    Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
        
       i = i + 1
    Loop
    ReDim Preserve LeyendaFormateada(0 To i)
Else
    Exit Sub
End If
End Sub


Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And Mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        LeyendaFormateada(k) = Mid(s, anti, i + 1)
        k = k + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, k, anti)
End If
End Function


Sub DibujarCartel()
Dim x As Integer, y As Integer, XPosCartel As Integer, YPosCartel As Integer, j As Integer, desp As Integer
If Not Cartel Then Exit Sub


XPosCartel = PosHUDX - 160
YPosCartel = PosHUDY - 270
x = XPosCartel + 20
y = YPosCartel + 60


Call Engine.Draw_GrhIndex(textura, XPosCartel, YPosCartel, 0)
For j = 0 To UBound(LeyendaFormateada)
    Engine.Engine_Text_Render LeyendaFormateada(j), x, y + desp, LongWhite
    desp = desp + (frmMain.Font.size) + 5
Next
End Sub

