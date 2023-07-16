Attribute VB_Name = "bas_z0_img"
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


'With es ancho real en twips,
'scalewidth lo asignas a lo que quieras maximo para tener una escala de referencia
'drawmode,drawstyle,drawwidth como propiedades de form
'fillcolor,fillstyle tb

'cls borra
'pset pone un pixel de un color
'point retorna el color del pixel
'line linea o rectangulo
'step relativo al ultimo punto
'circle circulo, elipse
'paintpicture: grafico en una posicion
'scalemode
'0 usuario
'1 twips
'2 puntos
'3 pixels: es la minima, segun resolucion del monitor

'autoredraw memoria


Sub s_pintar_objeto_ejv(tipo_soporte As Integer, formulario As Object, Z As Double, Y As Double, X As Double, objeto As String, ByVal color_borde As Long, ByVal color_relleno As Long, direccion As Variant, p_zoom As Integer, grosor_linea As Long)

    Dim CZ As Double
    Dim CY As Double
    Dim CX As Double
    
    Dim sumar1 As Integer
    Dim sumar2 As Integer
    Dim sumar3 As Integer
    Dim sumar3_5 As Integer
    Dim sumar4 As Integer
    Dim sumar4_5 As Integer
    Dim sumar5 As Integer
    Dim sumar6 As Integer
    
    'Fijo el ancho de línea
    formulario.DrawWidth = grosor_linea
    
    'Fijo el tipo, el color del borde y el color de relleno
    Select Case color_borde
        Case CTE_DEGRADADOCOLOR
            color_borde = f_degradado_color(Z, Y, X, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0)
        Case CTE_DEGRADADOGRIS
            color_borde = f_degradado_gris(Z, Y, X, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0)
        Case CTE_TRANSPARENTE
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select
    Select Case color_relleno
        Case CTE_DEGRADADOCOLOR
            color_relleno = f_degradado_color(Z, Y, X, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0)
            formulario.FillStyle = vbFSSolid 'solid
            formulario.FillColor = color_relleno
        Case CTE_DEGRADADOGRIS
            color_relleno = f_degradado_gris(Z, Y, X, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0)
            formulario.FillStyle = vbFSSolid 'solid
            formulario.FillColor = color_relleno
        Case CTE_TRANSPARENTE
            formulario.FillStyle = vbFSTransparent 'transparent
        Case Else
            formulario.FillStyle = vbFSSolid 'solid
            formulario.FillColor = color_relleno
    End Select
    
    'Fijo la escala
    Select Case tipo_soporte
        Case CTE_FORMULARIO
            formulario.ScaleMode = vbPixels   ' Set scale to pixels.
            Select Case formulario.Name
                'Esto es para objetos gordos, todo menos puntos
                Case "frm_a0_va"
                    CZ = CTE_VA0_INI_Z + (separacion_mapa_va0 * Z)
                    CY = CTE_VA0_INI_Y + (separacion_mapa_va0 * Y)
                    CX = CTE_VA0_INI_X + (separacion_mapa_va0 * X)
                Case "frm_a0_mapa"
                    CZ = CTE_MAPA_INI_Z + (separacion_mapa_ma0 * Z)
                    CY = CTE_MAPA_INI_Y + (separacion_mapa_ma0 * Y)
                    CX = CTE_MAPA_INI_X + (separacion_mapa_ma0 * X)
                Case "frm_z0_graf"
                    CZ = Z
                    CY = Y
                    CX = X
                Case "frm_u0_font"
                    CZ = CTE_VA0_INI_Z + (separacion_mapa_va0 * Z)
                    CY = CTE_VA0_INI_Y + (separacion_mapa_va0 * Y)
                    CX = CTE_VA0_INI_X + (separacion_mapa_va0 * X)
                Case Else
                    MsgBox "Error: no existe ese nombre de formulario", vbCritical
            End Select
        Case CTE_IMPRESORA
        Case Else
            MsgBox "Error: no existe ese tipo de soporte", vbCritical
    End Select
            
    
    If objeto = CTE_HORMIGA Or objeto = CTE_PRISIONERO Then
        sumar1 = 1
        sumar2 = 2
        sumar3 = 3
        sumar3_5 = 3.5
        sumar4 = 4
        sumar4_5 = 4.5
        sumar5 = 5
        sumar6 = 6
        If p_zoom = CTE_ZOOM_SUPER3D Then
            sumar1 = 1 * 5
            sumar2 = 2 * 5
            sumar3_5 = 3.5 * 5
            sumar4 = 4 * 5
            sumar6 = 6 * 5
        End If
        
        ReDim HX(1 To 4) As Double
        ReDim HY(1 To 4) As Double
        direccion = CInt(direccion)
        Select Case direccion
            Case 0
                'es la inicial - arriba
                HX(1) = CX
                HY(1) = CY + sumar2
                HX(2) = CX
                HY(2) = CY - sumar2
                HX(3) = CX + sumar1
                HY(3) = CY - sumar6
                HX(4) = CX - sumar1
                HY(4) = CY - sumar6
            Case CTE_8_N
                'arriba
                HX(1) = CX
                HY(1) = CY + sumar2
                HX(2) = CX
                HY(2) = CY - sumar2
                HX(3) = CX + sumar1
                HY(3) = CY - sumar6
                HX(4) = CX - sumar1
                HY(4) = CY - sumar6
            Case CTE_8_NE
                'arriba-derecha
                HX(1) = CX - (sumar2 * CTE_1ENTRERAIZDE2)
                HY(1) = CY + (sumar2 * CTE_1ENTRERAIZDE2)
                HX(2) = CX + (sumar2 * CTE_1ENTRERAIZDE2)
                HY(2) = CY - (sumar2 * CTE_1ENTRERAIZDE2)
                HX(3) = CX + sumar3_5
                HY(3) = CY - sumar5
                HX(4) = CX + sumar4_5
                HY(4) = CY - sumar3
            Case CTE_8_E
                'derecha
                HX(1) = CX - sumar2
                HY(1) = CY
                HX(2) = CX + sumar2
                HY(2) = CY
                HX(3) = CX + sumar6
                HY(3) = CY + sumar1
                HX(4) = CX + sumar6
                HY(4) = CY - sumar1
            Case CTE_8_SE
                'abajo-derecha
                HX(1) = CX - (sumar2 * CTE_1ENTRERAIZDE2)
                HY(1) = CY - (sumar2 * CTE_1ENTRERAIZDE2)
                HX(2) = CX + (sumar2 * CTE_1ENTRERAIZDE2)
                HY(2) = CY + (sumar2 * CTE_1ENTRERAIZDE2)
                HX(3) = CX + sumar3_5
                HY(3) = CY + sumar5
                HX(4) = CX + sumar4_5
                HY(4) = CY + sumar3
            Case CTE_8_S
                'abajo
                HX(1) = CX
                HY(1) = CY - sumar2
                HX(2) = CX
                HY(2) = CY + sumar2
                HX(3) = CX + sumar1
                HY(3) = CY + sumar6
                HX(4) = CX - sumar1
                HY(4) = CY + sumar6
            Case CTE_8_SO
                'abajo-izquierda
                HX(1) = CX + (sumar2 * CTE_1ENTRERAIZDE2)
                HY(1) = CY - (sumar2 * CTE_1ENTRERAIZDE2)
                HX(2) = CX - (sumar2 * CTE_1ENTRERAIZDE2)
                HY(2) = CY + (sumar2 * CTE_1ENTRERAIZDE2)
                HX(3) = CX - sumar3_5
                HY(3) = CY + sumar4
                HX(4) = CX - sumar4_5
                HY(4) = CY + sumar3
            Case CTE_8_O
                'izquierda
                HX(1) = CX + sumar2
                HY(1) = CY
                HX(2) = CX - sumar2
                HY(2) = CY
                HX(3) = CX - sumar6
                HY(3) = CY + sumar1
                HX(4) = CX - sumar6
                HY(4) = CY - sumar1
            Case CTE_8_NO
                'arriba-izquierda
                HX(1) = CX + (sumar2 * CTE_1ENTRERAIZDE2)
                HY(1) = CY + (sumar2 * CTE_1ENTRERAIZDE2)
                HX(2) = CX - (sumar2 * CTE_1ENTRERAIZDE2)
                HY(2) = CY - (sumar2 * CTE_1ENTRERAIZDE2)
                HX(3) = CX - sumar3_5
                HY(3) = CY - sumar5
                HX(4) = CX - sumar4_5
                HY(4) = CY - sumar3
            Case Else
                MsgBox "Error: no existe esa dirección", vbCritical
        End Select
    End If
    
    
    Select Case objeto
        Case CTE_PUNTO
            'Los puntos y linea con anterior van aparte ya que usan X y no CX
            Select Case formulario.Name
                Case "frm_a0_va"
                    CX = separacion_mapa_va0 + X
                Case "frm_a0_mapa"
                    CX = separacion_mapa_ma0 + X
                Case "frm_z0_graf"
                    CX = separacion_grafico_gra + X
                Case Else
                    MsgBox "Error: no existe ese nombre de formulario", vbCritical
            End Select
            CY = Y
            formulario.PSet (CX, CY), color_borde
        Case "linea_con_anterior"
            'Los puntos y linea con anterior van aparte ya que usan X y no CX
            Select Case formulario.Name
                Case "frm_a0_va"
                    CX = separacion_mapa_va0 + X
                Case "frm_a0_mapa"
                    CX = separacion_mapa_ma0 + X
                Case "frm_z0_graf"
                    CX = separacion_grafico_gra + X
                Case Else
                    MsgBox "Error: no existe ese nombre de formulario", vbCritical
            End Select
            CY = Y
            formulario.Line (ant_CX_gra, ant_CY_gra)-(CX, CY), color_borde
            ant_CX_gra = CX
            ant_CY_gra = CY
            formulario.Line (CX, CY)-(CX, CY), color_borde, BF
        Case CTE_CUBO
            Select Case p_zoom
                Case CTE_ZOOM_DETALLE
                    If color_relleno = CTE_TRANSPARENTE Then
                        formulario.Line (CX - 4, CY - 4)-(CX + 4, CY + 4), color_borde, B
                    Else
                        formulario.Line (CX - 4, CY - 4)-(CX + 4, CY + 4), color_borde, BF
                    End If
                Case CTE_ZOOM_PANORAMICA
                    If color_relleno = CTE_TRANSPARENTE Then
                        formulario.Line (CX - 1, CY - 1)-(CX + 1, CY + 1), color_borde, B
                    Else
                        formulario.Line (CX - 1, CY - 1)-(CX + 1, CY + 1), color_borde, BF
                    End If
                Case CTE_ZOOM_PIXELS
                    formulario.PSet (CX, CY), color_borde
                Case CTE_ZOOM_3D
                    s_pintar_cubo3D_ejv tipo_soporte, formulario, CZ, CY, CX, color_borde, color_relleno, CTE_DIRECC_NINGUNA, 6, grosor_linea, False
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
            End Select
        Case CTE_CUADRADOCURSOR
            formulario.FillStyle = vbFSTransparent 'transparent
            Select Case p_zoom
                Case CTE_ZOOM_DETALLE
                    formulario.Line (CX - 5, CY - 5)-(CX + 5, CY + 5), color_borde, B
                Case CTE_ZOOM_PANORAMICA
                    formulario.Line (CX - 2, CY - 2)-(CX + 2, CY + 2), color_borde, B
                Case CTE_ZOOM_PIXELS
                    formulario.PSet (CX, CY), color_borde
                Case CTE_ZOOM_3D
                    s_pintar_cubo3D_ejv tipo_soporte, formulario, CZ, CY, CX, color_borde, color_relleno, CTE_DIRECC_NINGUNA, 7, grosor_linea, False
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
            End Select
            formulario.FillStyle = vbFSSolid 'solid
        Case CTE_ESFERA
            Select Case p_zoom
                Case CTE_ZOOM_DETALLE
                    formulario.Circle (CX, CY), 7, color_borde
                Case CTE_ZOOM_PANORAMICA
                    formulario.Circle (CX, CY), 2, color_borde
                Case CTE_ZOOM_PIXELS
                    formulario.PSet (CX, CY), color_borde
                Case CTE_ZOOM_3D
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, CY, CX, color_borde, color_relleno, CTE_DIRECC_NINGUNA, 8, grosor_linea
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
            End Select
        Case CTE_PLANTA
            formulario.FillColor = cct_ejv(CTE_VERDECLARO)
            Select Case p_zoom
                Case CTE_ZOOM_DETALLE
                    formulario.Circle (CX, CY), 5, cct_ejv(CTE_NEGRO)
                Case CTE_ZOOM_PANORAMICA
                    formulario.FillColor = cct_ejv(CTE_VERDECLARO)
                    formulario.Line (CX - 1, CY - 1)-(CX + 1, CY + 1), cct_ejv(CTE_VERDECLARO), BF
                Case CTE_ZOOM_PIXELS
                    formulario.PSet (CX, CY), cct_ejv(CTE_VERDECLARO)
                Case CTE_ZOOM_3D
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, CY, CX, cct_ejv(CTE_NEGRO), cct_ejv(CTE_VERDECLARO), CTE_DIRECC_NINGUNA, 5, grosor_linea
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
            End Select
        Case CTE_PLANTALLENA
            Select Case p_zoom
                Case CTE_ZOOM_DETALLE
                    formulario.FillColor = color_borde
                    formulario.Circle (CX, CY), 5, cct_ejv(CTE_NEGRO)
                    formulario.FillColor = cct_ejv(CTE_VERDECLARO)
                    formulario.Circle (CX, CY), 2, cct_ejv(CTE_VERDECLARO)
                Case CTE_ZOOM_PANORAMICA
                    formulario.FillColor = cct_ejv(CTE_VERDECLARO)
                    formulario.Line (CX - 1, CY - 1)-(CX + 1, CY + 1), cct_ejv(CTE_VERDECLARO), BF
                Case CTE_ZOOM_PIXELS
                    formulario.PSet (CX, CY), color_borde
                Case CTE_ZOOM_3D
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, CY, CX, cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 5, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, CY, CX, cct_ejv(CTE_VERDECLARO), cct_ejv(CTE_VERDECLARO), CTE_DIRECC_NINGUNA, 2, grosor_linea
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
            End Select
        Case CTE_HORMIGA
            Select Case p_zoom
                Case CTE_ZOOM_DETALLE
                    formulario.FillColor = color_relleno
                    formulario.Circle (HX(1), HY(1)), 3, color_borde
                    formulario.Circle (HX(2), HY(2)), 3, color_borde
                    formulario.Circle (HX(3), HY(3)), 1, color_borde
                    formulario.Circle (HX(4), HY(4)), 1, color_borde
                Case CTE_ZOOM_PANORAMICA
                    'En panoramica para hormigas pongo solo una bola
                    formulario.FillColor = color_relleno
                    formulario.Line (CX - 1, CY - 1)-(CX + 1, CY + 1), color_relleno, BF
                Case CTE_ZOOM_PIXELS
                    formulario.PSet (CX, CY), color_relleno
                Case CTE_ZOOM_3D
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(1), HX(1), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 3, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(2), HX(2), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 3, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(3), HX(3), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 1, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(4), HX(4), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 1, grosor_linea
                Case CTE_ZOOM_SUPER3D
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(1), HX(1), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 3 * 4, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(2), HX(2), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 3 * 4, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ + 4, HY(3), HX(3), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 1 * 4, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ + 4, HY(4), HX(4), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 1 * 4, grosor_linea
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
            End Select
        Case CTE_HORMIMUERTA
            'Pongo el color de relleno gris porque lo de transparente no va bien
            formulario.FillColor = color_relleno
            'formulario.FillStyle = vbFSTransparent 'transparent
            If p_zoom = CTE_ZOOM_DETALLE Then
                formulario.Line (CX - 2, CY)-(CX + 3, CY), cct_ejv(CTE_BLANCO)
                formulario.Line (CX, CY - 2)-(CX, CY + 3), cct_ejv(CTE_BLANCO)
                'Si es panoramica no pinto nada que solo molestaria
            End If
            'formulario.FillStyle = vbFSSolid 'solid
        Case CTE_PRISIONERO
            Select Case p_zoom
                Case CTE_ZOOM_DETALLE
                    formulario.FillColor = color_relleno
                    formulario.Circle (HX(1), HY(1)), 2, color_borde
                    formulario.Circle (HX(2), HY(2)), 2, color_borde
                    formulario.Circle (HX(3), HY(3)), 1, color_borde
                    formulario.Circle (HX(4), HY(4)), 1, color_borde
                Case CTE_ZOOM_PANORAMICA
                    'En panoramica pongo solo una bola
                    formulario.FillColor = color_relleno
                    formulario.Line (CX - 1, CY - 1)-(CX + 1, CY + 1), color_relleno, BF
                Case CTE_ZOOM_PIXELS
                    formulario.PSet (CX, CY), color_relleno
                Case CTE_ZOOM_3D
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(1), HX(1), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 2, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(2), HX(2), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 2, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(3), HX(3), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 1, grosor_linea
                    s_pintar_esfera3D_ejv tipo_soporte, formulario, CZ, HY(4), HX(4), cct_ejv(CTE_NEGRO), color_relleno, CTE_DIRECC_NINGUNA, 1, grosor_linea
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
            End Select
        Case Else
            MsgBox "Error: no existe ese objeto", vbCritical
    End Select
    
    
    
End Sub

Sub s_pintar_esfera3D_ejv(tipo_soporte As Integer, formulario As Object, Z As Double, Y As Double, X As Double, color_borde As Long, color_relleno As Long, direccion As Variant, radio As Double, grosor_linea As Long)

    'Pinta un objeto 3D en perspectiva Axonométrica

    Dim CY As Double
    Dim CX As Double
    Dim cont As Double
    Dim salto As Integer
    
    Dim px1 As Double
    Dim py1 As Double
    Dim px2 As Double
    Dim py2 As Double
    
    'Fijo el ancho de línea
    formulario.DrawWidth = grosor_linea
    
    'Fijo el tipo y el color de relleno
    Select Case color_relleno
        Case CTE_DEGRADADOCOLOR
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
        Case CTE_DEGRADADOGRIS
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
        Case CTE_TRANSPARENTE
            formulario.FillStyle = vbFSTransparent 'transparent
        Case Else
            formulario.FillStyle = vbFSSolid 'solid
            formulario.FillColor = color_relleno
    End Select
    
    'Fijo la escala
    Select Case tipo_soporte
        Case CTE_FORMULARIO
            formulario.ScaleMode = vbPixels   ' Set scale to pixels.
        Case CTE_IMPRESORA
        Case Else
            MsgBox "Error: no existe ese tipo de soporte", vbCritical
    End Select
            
    salto = radio * 2
            
    'Calculo el centro X e Y en el plano
    s_zyx2yx tipo_soporte, formulario, Z, Y, X, CY, CX

    'Pinto el objeto
    formulario.Circle (CX, CY), radio, color_borde '10


End Sub

Sub s_pintar_ejes3D(tipo_soporte As Integer, formulario As Object, grosor_linea As Long)

    'Pinta los ejes 3D en perspectiva Axonométrica

    Dim centroy As Double
    Dim centrox As Double

    Dim dist_z As Double
    Dim dist_y As Double
    Dim dist_x As Double
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    
    formulario.DrawWidth = grosor_linea
    
    Select Case formulario.Name
        Case "frm_a0_va"
            dist_z = mapa_pisos_va0 * salto
            dist_y = mapa_filas_va0 * salto
            dist_x = mapa_columnas_va0 * salto
        Case "frm_a0_mapa"
            dist_z = mapa_pisos_ma0 * salto
            dist_y = mapa_filas_ma0 * salto
            dist_x = mapa_columnas_ma0 * salto
        Case "frm_z0_graf"
            dist_z = 120
            dist_y = 120
            dist_x = 120
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: "
    End Select
    
    formulario.ScaleMode = vbPixels   ' Set scale to pixels.

    'Calculo el punto central
    s_centro_ventana_ejv formulario, centroy, centrox

    s_pintar_linea3D_ejv CTE_FORMULARIO, formulario, 0, 0, 0, dist_z, 0, 0, cct_ejv(CTE_NEGRO), 2
    s_pintar_linea3D_ejv CTE_FORMULARIO, formulario, 0, 0, 0, 0, dist_y, 0, cct_ejv(CTE_NEGRO), 2
    s_pintar_linea3D_ejv CTE_FORMULARIO, formulario, 0, 0, 0, 0, 0, dist_x, cct_ejv(CTE_NEGRO), 2

End Sub


Sub s_zyx2yx(tipo_soporte As Integer, formulario As Object, Z As Double, Y As Double, X As Double, CY As Double, CX As Double)

    Select Case tipo_soporte
        Case CTE_FORMULARIO
        Case CTE_IMPRESORA
        Case Else
            MsgBox "Error: no existe ese tipo de soporte", vbCritical
    End Select

    'Me situo en el punto central en 2D
    s_centro_ventana_ejv formulario, CY, CX
    
    'Avanzo en el eje X
    CX = CX + (CTE_RAIZDE2ENTRERAIZDE3 * X)
    CY = CY - (CTE_1ENTRERAIZDE3 * X) + X
    'Avanzo en el eje Y
    CX = CX - (CTE_RAIZDE2ENTRERAIZDE3 * Y)
    CY = CY - (CTE_1ENTRERAIZDE3 * Y) + Y
    'Avanzo en el eje Z
    CY = CY - (Z * CTE_2ENTRERAIZDE5)
    
End Sub

Sub s_pintar_cubo3D_ejv(tipo_soporte As Integer, formulario As Object, Z As Double, Y As Double, X As Double, color_borde As Long, color_relleno As Long, direccion As Variant, radio As Double, grosor_linea As Long, pintar_lineas_ocultas As Boolean)

    'Pinta un objeto 3D en perspectiva Axonométrica

    Dim CY As Double
    Dim CX As Double
    Dim cont As Double
    Dim salto As Integer
    
    Dim px1 As Double
    Dim py1 As Double
    Dim px2 As Double
    Dim py2 As Double
    
    'Fijo el ancho de línea
    formulario.DrawWidth = grosor_linea
    
    'Fijo el tipo y el color de relleno
    Select Case color_relleno
        Case CTE_DEGRADADOCOLOR
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
        Case CTE_DEGRADADOGRIS
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
        Case CTE_TRANSPARENTE
            formulario.FillStyle = vbFSTransparent 'transparent
        Case Else
            formulario.FillStyle = vbFSSolid 'solid
            formulario.FillColor = color_relleno
    End Select
    
    'Fijo la escala
    Select Case tipo_soporte
        Case CTE_FORMULARIO
            formulario.ScaleMode = vbPixels   ' Set scale to pixels.
        Case CTE_IMPRESORA
        Case Else
            MsgBox "Error: no existe ese tipo de soporte", vbCritical
    End Select
            
    salto = radio * 2
            
    'Calculo el centro X e Y en el plano
    s_zyx2yx tipo_soporte, formulario, Z, Y, X, CY, CX
    
    'Pinto el objeto
    'solo se rellenan en VB circulos y rectangulos
    'hacemos un heptagono (6 caras) de puntos
    Dim ptox(1 To 7) As Integer
    Dim ptoy(1 To 7) As Integer
    ptox(1) = CX 'centro
    ptoy(1) = CY 'centro
    ptox(2) = CX  'arriba
    ptoy(2) = CY - (salto * CTE_2ENTRERAIZDE5) 'arriba
    ptox(3) = CX + (CTE_RAIZDE2ENTRERAIZDE3 * salto)  'la de la derecha
    ptoy(3) = CY - (CTE_1ENTRERAIZDE5 * salto) 'la de la derecha
    ptox(4) = CX + (CTE_RAIZDE2ENTRERAIZDE3 * salto)
    ptoy(4) = CY - (CTE_1ENTRERAIZDE3 * salto) + salto
    ptox(5) = CX  '
    ptoy(5) = CY + (salto * CTE_2ENTRERAIZDE5) '
    ptox(6) = CX - (CTE_RAIZDE2ENTRERAIZDE3 * salto)  '
    ptoy(6) = CY - (CTE_1ENTRERAIZDE3 * salto) + salto '
    ptox(7) = CX - (CTE_RAIZDE2ENTRERAIZDE3 * salto)  '
    ptoy(7) = CY - (CTE_1ENTRERAIZDE5 * salto) '
    'eje Z
    'Pinto los bordes
    formulario.Line (ptox(1), ptoy(1))-(ptox(5), ptoy(5)), color_borde
    formulario.Line (ptox(3), ptoy(3))-(ptox(4), ptoy(4)), color_borde
    formulario.Line (ptox(7), ptoy(7))-(ptox(6), ptoy(6)), color_borde
    
    formulario.Line (ptox(1), ptoy(1))-(ptox(3), ptoy(3)), color_borde
    formulario.Line (ptox(7), ptoy(7))-(ptox(2), ptoy(2)), color_borde
    formulario.Line (ptox(5), ptoy(5))-(ptox(4), ptoy(4)), color_borde
    
    formulario.Line (ptox(1), ptoy(1))-(ptox(7), ptoy(7)), color_borde
    formulario.Line (ptox(5), ptoy(5))-(ptox(6), ptoy(6)), color_borde
    formulario.Line (ptox(3), ptoy(3))-(ptox(2), ptoy(2)), color_borde
    
    'Pinto las lineas ocultas
    If pintar_lineas_ocultas Then
        formulario.Line (ptox(1), ptoy(1))-(ptox(2), ptoy(2)), color_borde
        formulario.Line (ptox(1), ptoy(1))-(ptox(4), ptoy(4)), color_borde
        formulario.Line (ptox(1), ptoy(1))-(ptox(6), ptoy(6)), color_borde
    End If
    
    'Pinto el relleno si es distinto de transparente
    If color_relleno >= 0 Then
        'Familia
        For cont = 0 To salto
            'Avanzo en esa direccion un pixel
            px1 = ptox(7) + (CTE_RAIZDE2ENTRERAIZDE3 * cont)
            py1 = ptoy(7) + (CTE_1ENTRERAIZDE3 * cont)
            px2 = ptox(2) + (CTE_RAIZDE2ENTRERAIZDE3 * cont)
            py2 = ptoy(2) + (CTE_1ENTRERAIZDE3 * cont)
            formulario.Line (px1, py1)-(px2, py2), color_relleno
        Next cont
        'Familia
        For cont = 0 To salto
            'Avanzo en esa direccion un pixel
            px1 = ptox(6) + (CTE_RAIZDE2ENTRERAIZDE3 * cont)
            py1 = ptoy(6) + (CTE_1ENTRERAIZDE3 * cont)
            px2 = ptox(7) + (CTE_RAIZDE2ENTRERAIZDE3 * cont)
            py2 = ptoy(7) + (CTE_1ENTRERAIZDE3 * cont)
            formulario.Line (px1, py1)-(px2, py2), color_relleno
        Next cont
        For cont = 0 To salto
            'Avanzo en esa direccion un pixel
            px1 = ptox(1)
            py1 = ptoy(1) + (cont)
            px2 = ptox(3)
            py2 = ptoy(3) + (cont)
            formulario.Line (px1, py1)-(px2, py2), color_relleno
        Next cont
    End If
    
End Sub

Sub s_pintar_linea3D_ejv(tipo_soporte As Integer, formulario As Object, Z1 As Double, Y1 As Double, X1 As Double, Z2 As Double, Y2 As Double, X2 As Double, color_borde As Long, grosor_linea As Long)

    Dim CY1 As Double
    Dim CX1 As Double
    
    Dim CY2 As Double
    Dim CX2 As Double

    'Fijo el ancho de línea
    formulario.DrawWidth = grosor_linea

    'Fijo la escala
    Select Case tipo_soporte
        Case CTE_FORMULARIO
            formulario.ScaleMode = vbPixels   ' Set scale to pixels.
        Case CTE_IMPRESORA
        Case Else
            MsgBox "Error: no existe ese tipo de soporte", vbCritical
    End Select
            
    'Calculo el centro X e Y en el plano
    s_zyx2yx tipo_soporte, formulario, Z1, Y1, X1, CY1, CX1
    s_zyx2yx tipo_soporte, formulario, Z2, Y2, X2, CY2, CX2

    formulario.Line (CX1, CY1)-(CX2, CY2), color_borde

End Sub


Function hipotenusa(cateto_uno As Double, cateto_dos As Double) As Double

    hipotenusa = Sqr((cateto_uno ^ 2) + (cateto_dos ^ 2))
    
End Function

Function cateto(hipotenusa As Double, cateto_uno As Double) As Double

    If cateto_uno > hipotenusa Then
        MsgBox "Error: el cateto no puede ser mayor que la hipotenusa", vbCritical
    End If

    cateto = Sqr((hipotenusa ^ 2) - (cateto_uno ^ 2))
    
End Function
Sub s_centro_ventana_ejv(formulario As Form, Y As Double, X As Double)

    formulario.ScaleMode = vbPixels   ' Set scale to pixels.

    X = formulario.Height / (2 * 10) 'ancho
    Y = formulario.Width / (2 * 25) 'alto

End Sub

Sub s_genera_plano_cubo_esfera(frutero As Boolean, ult_z As Integer, ult_y As Integer, ult_x As Integer)
    
    Dim CZ As Integer
    Dim CY As Integer
    Dim CX As Integer
    
    Dim num_esf As Integer
    Dim desplaz As Integer
    
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    
    
    For CZ = 1 To ult_z
        desplaz = 0
        If frutero And CZ Mod 2 = 0 Then
            desplaz = 10
        End If
        For CY = 1 To ult_y
            For CX = 1 To ult_x
            num_esf = CZ
            
If CX = 1 Or CY = 1 Or CZ = 1 Then
    s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, (CY * salto) - desplaz, (CX * salto) + desplaz, ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 2, False
Else
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * (CTE_RAIZDE2ENTRERAIZDE3 * 10), (CY * salto) - desplaz, (CX * salto) + desplaz, ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
End If
            Next CX
        Next CY
    Next CZ


End Sub

Sub s_genera_cubos()

    Dim CZ As Integer
    Dim CY As Integer
    Dim CX As Integer
    
    Dim num_esf As Integer
    Dim desplaz As Integer
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    
    For CZ = 1 To 5
        For CY = 1 To 5
            For CX = 1 To 5
            num_esf = CZ
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 2, False
s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
            Next CX
        Next CY
    Next CZ

End Sub


Sub s_genera_cubos2()

    Dim CZ As Integer
    Dim CY As Integer
    Dim CX As Integer
    
    Dim num_esf As Integer
    Dim desplaz As Integer
    Dim hay_cubo As Integer
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    
    For CZ = 1 To 5
        For CY = 1 To 5
            For CX = 1 To 5
            num_esf = CZ + 10
            hay_cubo = (CZ - 1) * (CY - 1) * (CX - 1)
            If hay_cubo = 0 Then
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, RGB(CZ * 255 / 5, CY * 255 / 5, CX * 255 / 5), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1, True
            Else
's_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), f_SumCirc(ncs_i_ejv, num_esf, 45), CTE_DIRECC_NINGUNA, radio, 1
            End If
            Next CX
        Next CY
    Next CZ

End Sub

Sub s_genera_cubos3_1()

    Dim CZ As Integer
    Dim CY As Integer
    Dim CX As Integer
    
    Dim num_esf As Integer
    Dim desplaz As Integer
    Dim hay_cubo As Integer
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    

    For CZ = 1 To 1
        For CY = 1 To 1
            For CX = 2 To 5
            num_esf = CZ + 10
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 3, True
            Next CX
        Next CY
    Next CZ

    For CZ = 2 To 2
        For CY = 1 To 1
            For CX = 2 To 5
            num_esf = CZ + 10
's_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), f_SumCirc(ncs_i_ejv, num_esf, 45), CTE_DIRECC_NINGUNA, radio, 1
s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), RGB(CZ * 255 / 5, CY * 255 / 5, CX * 255 / 5), CTE_DIRECC_NINGUNA, radio, 3
            Next CX
        Next CY
    Next CZ

s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, 0, 0, 0, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, 60, 1, True
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, 0, 0, 2 * 60, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, 60, 1, True
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, 0, 2 * 60, 0, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, 60, 1, True
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, 0, 2 * 60, 2 * 60, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, 60, 1, True


s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, 0, 0, 0, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, 2 * 60, 1


End Sub

Sub s_genera_cubos3_2()

    Dim CZ As Integer
    Dim CY As Integer
    Dim CX As Integer
    
    Dim num_esf As Integer
    Dim desplaz As Integer
    Dim hay_cubo As Integer
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    
    For CZ = 2 To 5
        For CY = 1 To 1
            For CX = 1 To 1
            num_esf = CZ + 10
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 3, True
            Next CX
        Next CY
    Next CZ

    For CZ = 1 To 1
        For CY = 1 To 1
            For CX = 2 To 5
            num_esf = CZ + 10
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 3, True
            Next CX
        Next CY
    Next CZ

    For CZ = 2 To 5
        For CY = 1 To 1
            For CX = 2 To 5
            num_esf = CZ + 10
s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), RGB(CZ * 255 / 5, CY * 255 / 5, CX * 255 / 5), CTE_DIRECC_NINGUNA, radio, 3
            Next CX
        Next CY
    Next CZ

End Sub

Sub s_genera_cubos3_3()

    Dim CZ As Integer
    Dim CY As Integer
    Dim CX As Integer
    
    Dim num_esf As Integer
    Dim desplaz As Integer
    Dim hay_cubo As Integer
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    
    For CZ = 2 To 5
        For CY = 1 To 1
            For CX = 1 To 1
            num_esf = CZ + 10
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 3, True
            Next CX
        Next CY
    Next CZ

    For CZ = 1 To 1
        For CY = 2 To 5
            For CX = 1 To 1
            num_esf = CZ + 10
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 3, True
            Next CX
        Next CY
    Next CZ

    For CZ = 1 To 1
        For CY = 1 To 1
            For CX = 2 To 5
            num_esf = CZ + 10
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 3, True
            Next CX
        Next CY
    Next CZ


    For CZ = 2 To 5
        For CY = 2 To 5
            For CX = 2 To 5
            num_esf = CZ + 10
s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * salto, CY * salto, CX * salto, cct_ejv(CTE_NEGRO), RGB(CZ * 255 / 5, CY * 255 / 5, CX * 255 / 5), CTE_DIRECC_NINGUNA, radio, 3
            Next CX
        Next CY
    Next CZ

End Sub



Sub s_genera_cosas()
    
    Dim CZ As Integer
    Dim CY As Integer
    Dim CX As Integer
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    

    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, 0, 10, 0, cct_ejv(CTE_VERDEBRILLANTE), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, 0, 0, 10, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, 5, 0, 0, cct_ejv(CTE_AMARILLO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    
    s_genera_mundo_esferas
    
    MsgBox ""
    s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, 10, 10, 10, cct_ejv(CTE_NEGRO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 2, False
    s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, 10, 10, 0, cct_ejv(CTE_ROSA), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 2, False
    s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, 0, 10, 10, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 2, False
    s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, 10, 0, 10, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 2, False
    MsgBox ""

    For CZ = 1 To 2
        For CY = 1 To 3
            For CX = 1 To 3
s_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, CZ * 10, CY * 10, CX * 10, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 2, False
            Next CX
        Next CY
    Next CZ
    MsgBox ""

End Sub


Sub s_genera_mundo_esferas()

    Dim dime As Integer

    '10 esferas, 3 coordenadas zyx
    ReDim esfera(1 To 10, 1 To 3) As Double
    Dim num_esf As Integer
    Const Z = 1
    Const Y = 2
    Const X = 3
    
    ReDim Plano(1 To 3, 1 To 3) As Double
    ReDim cuarta_uva(1 To 3) As Double
    
    Dim radio As Double
    Dim salto As Double
    radio = 10
    salto = radio * 2
    
    num_esf = 0
    
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = 0
    esfera(num_esf, Y) = 3 * radio
    esfera(num_esf, X) = 0
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    'Debug.Print "Esfera " & num_esf & ":" & esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X)
    
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = 0
    esfera(num_esf, Y) = 2 * radio
    esfera(num_esf, X) = 0
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    'Debug.Print "Esfera " & num_esf & ":" & esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X)
    
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = 0
    esfera(num_esf, Y) = (3 * radio + 2 * radio) / 2
    esfera(num_esf, X) = (CTE_1ENTRERAIZDE3 * radio)
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    'Debug.Print "Esfera " & num_esf & ":" & esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X)
    
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = CTE_RAIZDE3porRAIZDE2_entre4 * radio
    esfera(num_esf, Y) = (3 * radio + 2 * radio) / 2
    esfera(num_esf, X) = (CTE_1ENTRERAIZDE3 * radio) / 2
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    'Debug.Print "Esfera " & num_esf & ":" & esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X)
    
    'Con 3 bolas hago un plano y calculo la cuarta, que sale del plano como un champiñon
    'Puntos 1-2-3 valores z-y-x
    For dime = 1 To 3 'z,y,x
        Plano(1, dime) = esfera(1, dime) 'pto1
        Plano(2, dime) = esfera(2, dime) 'pto2
        Plano(3, dime) = esfera(3, dime) 'pto3
    Next dime
    s_calcular_cuarta_uva Plano(), cuarta_uva(), radio, False
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = cuarta_uva(Z)
    esfera(num_esf, Y) = cuarta_uva(Y)
    esfera(num_esf, X) = cuarta_uva(X)
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    
    
    'Con 3 bolas hago un plano y calculo la cuarta, que sale del plano como un champiñon
    'Puntos 1-2-3 valores z-y-x
    For dime = 1 To 3 'z,y,x
        Plano(1, dime) = esfera(2, dime) 'pto1
        Plano(2, dime) = esfera(3, dime) 'pto2
        Plano(3, dime) = esfera(4, dime) 'pto3
    Next dime
    s_calcular_cuarta_uva Plano(), cuarta_uva(), radio, True
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = cuarta_uva(Z)
    esfera(num_esf, Y) = cuarta_uva(Y)
    esfera(num_esf, X) = cuarta_uva(X)
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
 
 
 
 
 '==========================================
 
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = 0
    esfera(num_esf, Y) = 0
    esfera(num_esf, X) = 3 * radio
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    'Debug.Print "Esfera " & num_esf & ":" & esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X)
 
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = 0
    esfera(num_esf, Y) = 0
    esfera(num_esf, X) = 2 * radio
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    'Debug.Print "Esfera " & num_esf & ":" & esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X)
 
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = 0
    esfera(num_esf, Y) = (CTE_1ENTRERAIZDE3 * radio)
    esfera(num_esf, X) = (3 * radio + 2 * radio) / 2
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
    'Debug.Print "Esfera " & num_esf & ":" & esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X)
 
    'Con 3 bolas hago un plano y calculo la cuarta, que sale del plano como un champiñon
    'Puntos 1-2-3 valores z-y-x
    For dime = 1 To 3 'z,y,x
        Plano(1, dime) = esfera(num_esf - 2, dime) 'pto1
        Plano(2, dime) = esfera(num_esf - 1, dime) 'pto2
        Plano(3, dime) = esfera(num_esf, dime) 'pto3
    Next dime
    s_calcular_cuarta_uva Plano(), cuarta_uva(), radio, True
    'Añado la esfera
    num_esf = num_esf + 1
    esfera(num_esf, Z) = cuarta_uva(Z)
    esfera(num_esf, Y) = cuarta_uva(Y)
    esfera(num_esf, X) = cuarta_uva(X)
    s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, esfera(num_esf, Z), esfera(num_esf, Y), esfera(num_esf, X), ccs_ejv(f_SumCirc(ncs_i_ejv, num_esf, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1
 
 
 
End Sub
Sub s_calcular_cuarta_uva(Plano() As Double, PtoDestino() As Double, radio As Double, invertir As Boolean)

    ReDim PtoCtro(1 To 3) As Double
    ReDim PtoDestino(1 To 3) As Double
    ReDim vector(1 To 3) As Double

    'Calculo el centro de ese plano (de esos tres puntos)
    s_calcula_centro_medios Plano(), PtoCtro()
    's_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_mapa, PtoCtro(Z), PtoCtro(Y), PtoCtro(X), cct_ejv(CTE_NEGRO),cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio,1
    
    
    'Convierto el plano del formato "tres puntos" al formato "vector"
    s_TresPtos2Vector Plano(), vector()
    's_pintar_cubo3D_ejv CTE_FORMULARIO, frm_a0_mapa, Vector(Z) * radio * 10, Vector(Y) * radio * 10, Vector(X) * radio * 10,  cct_ejv(CTE_AZUL),cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio, 1, false
    
    'Invierto el vector
    If invertir Then
        s_invertir_vector vector()
    End If
    
    'Avanzo desde el centro en esa direccion (vector) una distancia CTE_RAIZDE3 * (radio / 2)
    s_avanzar PtoCtro(), vector(), CTE_RAIZDE3porRAIZDE2_entre4 * radio, PtoDestino()

End Sub

Sub s_invertir_vector(vector() As Double)
    
    Dim dime As Integer
    
    For dime = 1 To 3 'z,y,x
        vector(dime) = -vector(dime)
    Next dime

End Sub
Sub s_avanzar(PtoOrigen() As Double, vector() As Double, distancia As Double, PtoDestino() As Double)

    ReDim PtoActual(1 To 3) As Double
    Dim Salir As Boolean
    Dim d_recorrer As Double
    Dim d_recorrida As Double
    Dim d_falta As Double
    Dim dime As Integer
    Dim mayor As Integer
    

    For dime = 1 To 3
        PtoActual(dime) = PtoOrigen(dime)
    Next dime

    'Avanzo cada vez la mitad de lo que me falta, como la tortuga de zenon y aquiles
    Salir = False
    d_recorrida = 0
    d_recorrer = distancia
    d_falta = distancia
    While Not Salir
        'Avanzo la mitad, pero solo en un eje, en el que mayor es el vector
        'mayor = f_mayor(vector(1), vector(2), vector(3))
        'PtoActual(mayor) = PtoActual(mayor) + ((d_falta / 2) * vector(mayor))
        'Avanzo en los tres ejes a la vez
        For dime = 1 To 3
            PtoActual(dime) = PtoActual(dime) + ((d_falta / 2) * vector(dime))
        Next dime
        d_recorrida = dist2ptos3d(PtoOrigen(), PtoActual())
        d_falta = d_recorrer - d_recorrida
        If Int(d_falta * 1000000) < 1 Then
            Salir = True
        End If
    Wend

    For dime = 1 To 3
        PtoDestino(dime) = PtoActual(dime)
    Next dime



End Sub

Function dist2ptos3d(Ori() As Double, Des() As Double) As Double

    dist2ptos3d = Sqr((Ori(1) - Des(1)) ^ 2 + (Ori(2) - Des(2)) ^ 2 + (Ori(3) - Des(3)) ^ 2)

End Function


Sub s_TresPtos2Vector(TresPtos() As Double, vector() As Double)

    Dim dime As Integer
    Dim suma_vector As Double
    
    ReDim vecA(1 To 3) As Double
    ReDim vecB(1 To 3) As Double

    'Obtengo el vector champiñon

    'Calculo los vectores
    For dime = 1 To 3
        vecA(dime) = TresPtos(1, dime) - TresPtos(2, dime)
        vecB(dime) = TresPtos(1, dime) - TresPtos(3, dime)
    Next dime

    'Hago el producto vectorial
    vector(1) = (vecA(3) * vecB(2)) - (vecA(2) * vecB(3))
    vector(2) = (vecA(3) * vecB(1)) - (vecA(1) * vecB(3))
    vector(3) = (vecA(1) * vecB(2)) - (vecA(2) * vecB(1))

    'Simplifico el vector a tanto por uno
    suma_vector = 0
    For dime = 1 To 3
        suma_vector = suma_vector + Abs(vector(dime))
    Next dime

    'Ahora la suma de los 3 valores del vector es 1
    For dime = 1 To 3
        vector(dime) = vector(dime) / suma_vector
    Next dime

End Sub
Sub s_calcula_centro_medios(TresPtos() As Double, PtoA() As Double)

    'Este centro se calcula haciendo la media de los 3 ptos
    'y coincide con el orto y el bari para triangulos equilateros

    Dim dime As Integer

    For dime = 1 To 3
        PtoA(dime) = (TresPtos(1, dime) + TresPtos(2, dime) + TresPtos(3, dime)) / 3
    Next dime

End Sub

Sub s_calcula_centro_ortocentro(TresPtos() As Double, PtoA() As Double)

    'Ortocentro es el centro del triangulo que se define
    'creando perpendiculares desde el punto medio de
    'cada segmento de recta que forma cada lado del triangulo


    Dim dime As Integer

    For dime = 1 To 3
        PtoA(dime) = (TresPtos(1, dime) + TresPtos(2, dime) + TresPtos(3, dime)) / 3
    Next dime
    
End Sub

Sub s_calcula_centro_baricentro(TresPtos() As Double, PtoA() As Double)

    'Baricentro es el centro del triangulo que se define
    'uniendo cada vertice (pto) del triangulo con el punto
    'medio del segmento opuesto

    Dim pto As Integer
    Dim dime As Integer
    Dim Salir As Boolean
    
    
    ReDim medios(1 To 3, 1 To 3) As Double '3ptos-zyx
    ReDim tmp(1 To 3, 1 To 3) As Double '3ptos-zyx
    'Calculo los ptos medios 2 a 2 recursivamente
    'hasta que el error sea indetectable por ser un ordenador discreto
    'osea, calculo los puntos medios de los segmentos y supongo que este
    'es el nuevo triangulo y asi hasta el infinito
    
    For dime = 1 To 3
        For pto = 1 To 3
            tmp(pto, dime) = TresPtos(pto, dime)
        Next pto
    Next dime
    
    Salir = False
    While Not Salir
        For dime = 1 To 3
            medios(1, dime) = (tmp(1, dime) + tmp(2, dime)) / 2
            medios(2, dime) = (tmp(2, dime) + tmp(3, dime)) / 2
            medios(3, dime) = (tmp(3, dime) + tmp(1, dime)) / 2
            
            For pto = 1 To 3
                tmp(pto, dime) = medios(pto, dime)
            Next pto
        Next dime
        'Control fin
        If tmp(1, 1) = tmp(2, 1) And tmp(2, 1) = tmp(3, 1) Then
        If tmp(1, 2) = tmp(2, 2) And tmp(2, 2) = tmp(3, 2) Then
        If tmp(1, 3) = tmp(2, 3) And tmp(2, 3) = tmp(3, 3) Then
            Salir = True
        End If
        End If
        End If
    Wend

    'Pongo la solucion
    For dime = 1 To 3
        PtoA(dime) = tmp(1, dime)
    Next dime


End Sub

Function f_degradado_gris(Z As Double, Y As Double, X As Double, max_z As Double, max_y As Double, max_x As Double) As Long

    Dim media As Integer
    
    media = Int(Z * 255 / max_z + Y * 255 / max_y + X * 255 / max_x)

    f_degradado_gris = RGB(media, media, media)
    If f_degradado_gris = cct_ejv(cfondo_ejv) Then
        If media < max_z Then
            media = media + 1
        Else
            media = media - 1
        End If
        f_degradado_gris = RGB(media, media, media)
    End If

End Function
Function f_degradado_color(Z As Double, Y As Double, X As Double, max_z As Double, max_y As Double, max_x As Double) As Long

    f_degradado_color = RGB(Z * 255 / max_z, Y * 255 / max_y, X * 255 / max_x)
    If f_degradado_color = cct_ejv(cfondo_ejv) Then
        If Z < max_z Then
            f_degradado_color = RGB((Z + 1) * 255 / max_z, Y * 255 / max_y, X * 255 / max_x)
        Else
            f_degradado_color = RGB((Z - 1) * 255 / max_z, Y * 255 / max_y, X * 255 / max_x)
        End If
    End If
    
End Function
Sub s_colorVB2colorRGB(valor_VB As Long, dec_R As Integer, dec_G As Integer, dec_B As Integer)

    dec_R = Int(valor_VB / 1) Mod 256
    dec_G = Int(valor_VB / 256) Mod 256
    dec_B = Int(valor_VB / 65536) Mod 256

End Sub

Sub s_girar_sobre_eje(eje_de_giro As String, angulo As Double, Z As Double, Y As Double, X As Double)

    Dim radio As Double
    Dim arco_actual As Double
    Dim cociente As Double
    Dim nuevaY As Double
    Dim nuevaX As Double

    Select Case UCase(Trim(eje_de_giro))
        Case "Z", "P", "PISOS"
            radio = hipotenusa(positivo(Y), X)
            cociente = Y / radio
            Select Case cociente
                Case 0
                    arco_actual = -CTE_MEDIAVUELTA
                Case 1
                    arco_actual = 0
                Case -1
                    arco_actual = CTE_MEDIAVUELTA
                Case Else
                    If cociente > 0 And cociente < 1 Then
                        arco_actual = ArcSin(Y / radio)
                    ElseIf cociente < 0 And cociente > -1 Then
                        arco_actual = ArcSin(Y / radio)
                    ElseIf cociente > 1 Then
                        arco_actual = ArcSin(Y / radio)
                    ElseIf cociente < -1 Then
                        arco_actual = ArcSin(Y / radio)
                    End If
            End Select
            If X < 0 And Y < 0 Then
                arco_actual = arco_actual + CTE_MEDIAVUELTA
                nuevaX = radio * Sin(angulo - arco_actual)
                nuevaY = radio * Cos(angulo - arco_actual)
            Else
                If X < 0 Then
                    arco_actual = arco_actual + CTE_CUARTODEVUELTA
                    nuevaX = -radio * Sin(angulo - arco_actual)
                    nuevaY = radio * Cos(angulo - arco_actual)
                Else
                    nuevaX = radio * Cos(angulo - arco_actual)
                    nuevaY = -radio * Sin(angulo - arco_actual)
                End If
            End If
            'arco_actual = ArcCos(Y / radio)
            'nuevaX = radio * Sin(arco_actual - angulo)
            'nuevaY = -radio * Cos(arco_actual - angulo)
            'nuevaX = radio * Sin(angulo + arco_actual)
            'nuevaY = -radio * Cos(angulo + arco_actual)
Debug.Print cociente & "  " & arco_actual & "  " & angulo - arco_actual & "  " & Cos(angulo - arco_actual) & "  " & Sin(angulo - arco_actual)
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: "
    End Select

    If nuevaY = Y And nuevaX = X Then
        Beep
    End If

    Y = nuevaY
    X = nuevaX
End Sub
Function ArcCos(X As Double) As Double
    
    ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)

End Function

Function ArcSin(X As Double) As Double
    
    ArcSin = Atn(X / Sqr(-X * X + 1))

End Function


