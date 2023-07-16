VERSION 5.00
Begin VB.Form frm_u0_frac 
   Caption         =   "Fractal"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "z0_frac.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8475
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frm_u0_frac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------
' Shareware desarrollado por Manuel de la Herrán Gascón
' mherran@usa.net (Junio 1997 - Diciembre 1998) Madrid (Spain).
' http://www.geocities.com/SiliconValley/Vista/7491/
' -----------------------------------------------------------------------
' Este programa y sus ficheros fuente son grátis y de libre distribución.
' El código fuente está disponible y puede ser modificado, distribuido,
' o utilizado en otros programas con entera libertad.
' -----------------------------------------------------------------------
' Para mantenerse informado de las sucesivas versiones del programa
' y dónde conseguirlas, escriba un mail a mherran@usa.net
' Para sugerir posibles ampliaciones, enviar comentarios de cualquier tipo
' si se detectara algún error en la programación o en la instalación,
' o si se va a ampliar o utilizar una parte o todo este
' programa, no dude en ponerse en contacto con el autor.
'-----------------------------------------------------------------------

Dim detener As Boolean
Private Sub Form_Activate()

'------------------------------------------------------------------------
' Código adaptado de una adaptación de
' Charles Dumont de uno de los
' famosos algoritmos de fractales
' dumonfc1@SPACEMSG.JHUAPL.edu
' Gracias y un saludete!
'------------------------------------------------------------------------
    
    Dim X, Y, S As Single
    Dim NewX, NewY As Single
    Dim Iterations As Single
    Dim Transforms(1 To 4, 1 To 7) As Single
    Dim p(0 To 1024, 0 To 721) As Integer
    Dim i, j As Long
    Dim px, py, RandNum As Double
    
    ScaleMode = vbPixels
    detener = False
    BackColor = RGB(0, 0, 0)

    ' Initialize pixel color matrix
    For i = 0 To 1024
        For j = 0 To 721
            p(i, j) = 75
        Next
    Next

    ' Set scale factor
    S = 65

    'Initialize transforms
    Transforms(1, 1) = 0
    Transforms(1, 2) = 0
    Transforms(1, 3) = 0
    Transforms(1, 4) = 0.16
    Transforms(1, 5) = 0
    Transforms(1, 6) = 0
    Transforms(1, 7) = 0.01


    Transforms(2, 1) = 0.85
    Transforms(2, 2) = 0.04
    Transforms(2, 3) = -0.04
    Transforms(2, 4) = 0.85
    Transforms(2, 5) = 0
    Transforms(2, 6) = 1.6
    Transforms(2, 7) = 0.85

    Transforms(3, 1) = 0.2
    Transforms(3, 2) = -0.26
    Transforms(3, 3) = 0.23
    Transforms(3, 4) = 0.22
    Transforms(3, 5) = 0
    Transforms(3, 6) = 1.6
    Transforms(3, 7) = 0.07

    Transforms(4, 1) = -0.15
    Transforms(4, 2) = 0.28
    Transforms(4, 3) = 0.26
    Transforms(4, 4) = 0.24
    Transforms(4, 5) = 0
    Transforms(4, 6) = 0.44
    Transforms(4, 7) = 0.07

    ' Seed point
    X = 1
    Y = 1

    ' Number of points
    Iterations = 100000

    Randomize
    For i = 1 To Iterations
        ' Readjust coordinates to center of screen and with
        ' larger scale

        px = Int(S * X + ScaleWidth / 2)
        py = Int(ScaleHeight - S * Y)

        ' Increase color value of pixel
        If (px < ScaleWidth) And (px > 0) And (py > 0) And (py < ScaleHeight) Then
            p(px, py) = p(px, py) + 1
            If p(px, py) > 255 Then p(px, py) = 255
        End If

        ' Color current point
        If px >= 0 And py >= 0 Then
            If px <= ScaleWidth And py <= ScaleHeight Then
                PSet (px, py), RGB(0, p(px, py), 0)
                Me.FillColor = RGB(0, p(px, py), 0)
            End If
        End If

        Select Case Rnd
            Case 0 To Transforms(1, 7)
                RandNum = 1
            Case Transforms(1, 7) To Transforms(1, 7) + Transforms(2, 7)
                RandNum = 2
            Case Transforms(2, 7) To Transforms(2, 7) + Transforms(3, 7)
                RandNum = 3
            Case Transforms(3, 7) To 1
                RandNum = 4
        End Select

         ' Calculate next point
        NewX = Transforms(RandNum, 1) * X + Transforms(RandNum, 2) * Y + Transforms(RandNum, 5)
        NewY = Transforms(RandNum, 3) * X + Transforms(RandNum, 4) * Y + Transforms(RandNum, 6)
        ' Update current point
        X = NewX
        Y = NewY
        
        DoEvents
        If detener Then
            'Detenido por el usuario
            i = Iterations
        End If

    Next i
    
    
End Sub

Private Sub Form_Click()
    detener = True
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    detener = True
    Unload Me
End Sub
       
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    detener = True
    Unload Me
End Sub

