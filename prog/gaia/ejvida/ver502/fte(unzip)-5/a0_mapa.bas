Attribute VB_Name = "bas_a0_mapa"
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

'Mapas
Global ver_zoom_ma0 As Integer
Global separacion_mapa_ma0 As Integer
Global mapa_actual_ma0 As String 'Path completo
Global mapa_ma0() As Boolean
Global mapa_pisos_ma0 As Double
Global mapa_filas_ma0 As Double
Global mapa_columnas_ma0 As Double
Global copiar_mapa_a_va0_ma0 As Boolean
Global cursor_visible_mapa_ma0 As Boolean
Global mapa_sin_obstaculos_ma0 As Boolean

Function f_mapa_contar_obstaculos_va0() As Integer
    
    Dim p As Integer
    Dim f As Integer
    Dim c As Integer
    Dim total As Integer
    
    total = 0
    
    For p = 1 To mapa_pisos_ma0
    For f = 1 To mapa_filas_ma0
    For c = 1 To mapa_columnas_ma0
        If mapa_ma0(p, f, c) Then
            total = total + 1
        End If
    Next c
    Next f
    Next p

    f_mapa_contar_obstaculos_va0 = total

End Function

Sub s_mapa_inicializar_va0(mi_mapa() As Boolean, max_pisos As Double, max_filas As Double, max_col As Double, relleno As Boolean)
     
    'Ejemplo llamada
    's_mapa_inicializar_va0 mapa_ma0, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0, False
    
    Dim p As Double
    Dim f As Double
    Dim c As Double

    ReDim mi_mapa(1 To max_pisos, 1 To max_filas, 1 To max_col) As Boolean
    
    'Los arrays por defecto se rellenan de false asi que solo hace falta
    'en el caso de ser true
    If relleno Then
        For p = 1 To max_pisos
        For f = 1 To max_filas
        For c = 1 To max_col
            mi_mapa(p, f, c) = relleno
        Next c
        Next f
        Next p
    End If

End Sub
Sub s_ajustar_lugar_mapa(tipo_mapa As Integer, piso As Double, fila As Double, col As Double, max_pisos As Double, max_filas As Double, max_col As Double)

    Select Case tipo_mapa
        Case CTE_MAPA_ESFERICO
            If piso <= 0 Then piso = max_pisos + piso
            If piso > max_pisos Then piso = piso - max_pisos
            
            If fila <= 0 Then fila = max_filas + fila
            If fila > max_filas Then fila = fila - max_filas
            
            If col <= 0 Then col = max_col + col
            If col > max_col Then col = col - max_col
        Case CTE_MAPA_LIMITADO
            If piso <= 0 Then piso = 1
            If piso > max_pisos Then piso = max_pisos
            
            If fila <= 0 Then fila = 1
            If fila > max_filas Then fila = max_filas
            
            If col <= 0 Then col = 1
            If col > max_col Then col = max_col
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: "
    End Select
    
    

End Sub

Sub s_mapa_pintar_figura_va0(tipo As String, figura As String, longitud As Integer, Z As Double, Y As Double, X As Double)

Dim cont As Integer

Dim pos_z As Double
Dim pos_y As Double
Dim pos_x As Double

Dim p As Double
Dim f As Double
Dim c As Double

Screen.MousePointer = CTE_ARENA


Select Case tipo
    Case "Relleno"
        mapa_sin_obstaculos_ma0 = False
        Select Case figura
            Case "Horizontal"
                'Pintamos
                If longitud > mapa_columnas_ma0 Then longitud = mapa_columnas_ma0
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y
                    pos_x = X + cont
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = True
                Next cont
            Case "Vertical"
                If longitud > mapa_filas_ma0 Then longitud = mapa_filas_ma0
                'Pintamos
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y + cont
                    pos_x = X
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = True
                Next cont
            Case "Cuadrado"
                If longitud > mapa_filas_ma0 Then longitud = mapa_filas_ma0
                If longitud > mapa_columnas_ma0 Then longitud = mapa_columnas_ma0
                'Pintamos
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y
                    pos_x = X + cont
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = True
                Next cont
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y + longitud - 1
                    pos_x = X + cont
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(CTE_GRISOSCURO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = True
                Next cont
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y + cont
                    pos_x = X
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(CTE_GRISOSCURO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = True
                Next cont
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y + cont
                    pos_x = X + longitud - 1
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(CTE_GRISOSCURO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = True
                Next cont
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: figura a pintar no existe"
        End Select
    Case "Vacío"
        Select Case figura
            Case "Horizontal"
                If longitud > mapa_columnas_ma0 Then longitud = mapa_columnas_ma0
                'Pintamos
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y
                    pos_x = X + cont
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = False
                Next cont
            Case "Vertical"
                If longitud > mapa_filas_ma0 Then longitud = mapa_filas_ma0
                'Pintamos
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y + cont
                    pos_x = X
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = False
                Next cont
            Case "Cuadrado"
                If longitud > mapa_filas_ma0 Then longitud = mapa_filas_ma0
                If longitud > mapa_columnas_ma0 Then longitud = mapa_columnas_ma0
                'Pintamos
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y
                    pos_x = X + cont
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = False
                Next cont
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y + longitud - 1
                    pos_x = X + cont
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = False
                Next cont
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y + cont
                    pos_x = X
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = False
                Next cont
                For cont = 0 To longitud - 1
                    pos_z = Z
                    pos_y = Y + cont
                    pos_x = X + longitud - 1
                    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, pos_z, pos_y, pos_x, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                    mapa_ma0(pos_z, pos_y, pos_x) = False
                Next cont
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: figura a pintar no existe"
        End Select
    Case "Rellenar Todo"
        mapa_sin_obstaculos_ma0 = False
        For p = 1 To mapa_pisos_ma0
        For f = 1 To mapa_filas_ma0
        For c = 1 To mapa_columnas_ma0
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            mapa_ma0(p, f, c) = True
            'Como esta opcion es tan radical supongo que quiere volver a empezar el test
            nodo_visitado_va0(p, f, c) = 0
        Next c
        Next f
        Next p
    Case "Vaciar Todo"
        mapa_sin_obstaculos_ma0 = True
        For p = 1 To mapa_pisos_ma0
        For f = 1 To mapa_filas_ma0
        For c = 1 To mapa_columnas_ma0
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            mapa_ma0(p, f, c) = False
            'Como esta opcion es tan radical supongo que quiere volver a empezar el test
            nodo_visitado_va0(p, f, c) = 0
        Next c
        Next f
        Next p
    Case Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: tipo de figura a pintar no existe"
End Select
            
Screen.MousePointer = CTE_DEFECTO
            

End Sub
Sub copia_dim_va0_2_ma0_va0()

    mapa_pisos_ma0 = mapa_pisos_va0
    mapa_filas_ma0 = mapa_filas_va0
    mapa_columnas_ma0 = mapa_columnas_va0

End Sub

Sub copia_dim_ma0_2_va0_va0()
    
    mapa_pisos_va0 = mapa_pisos_ma0
    mapa_filas_va0 = mapa_filas_ma0
    mapa_columnas_va0 = mapa_columnas_ma0

End Sub

Sub copia_dim_va0_2_viejo_va0()

    viejo_mapa_pisos_va0 = mapa_pisos_va0
    viejo_mapa_filas_va0 = mapa_filas_va0
    viejo_mapa_columnas_va0 = mapa_columnas_va0

End Sub

Sub s_copiar_mapa_ma0_sobre_va0_va0()
    
    'Copio todo el array
    Dim p As Double
    Dim f As Double
    Dim c As Double

    Screen.MousePointer = CTE_ARENA
    
    'Copio el modo de ver el mapa
    ver_zoom_va0 = ver_zoom_ma0
    s_cargar_tipo_zoom_va0
    s_fijar_separacion_mapa_va0
    
    'Copio las dimensiones del mapa sobre el actual y el viejo
    copia_dim_ma0_2_va0_va0
    copia_dim_va0_2_viejo_va0
    
    'Borro el mapa si existia porque puede tener bichos y lo cargo
    ReDim mapa_va0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Integer
    
    'Copio el mapa
    mapa_sin_obstaculos_va0 = mapa_sin_obstaculos_ma0
    If Not mapa_sin_obstaculos_ma0 Then
        Screen.MousePointer = CTE_ARENA
        For p = 1 To mapa_pisos_ma0
        For f = 1 To mapa_filas_ma0
        For c = 1 To mapa_columnas_ma0
            If mapa_ma0(p, f, c) Then
                mapa_va0(p, f, c) = CTE_MAPA_OBSTACULO
            Else
                mapa_va0(p, f, c) = CTE_MAPA_VACIO
            End If
        Next c
        Next f
        Next p
    End If
    
    Screen.MousePointer = CTE_DEFECTO


End Sub
Sub s_copiar_mapa_va0_sobre_ma0_va0()
    
    'Copio todo el array
    Dim p As Double
    Dim f As Double
    Dim c As Double

    Screen.MousePointer = CTE_ARENA
    
    'Copio el modo de ver el mapa
    ver_zoom_ma0 = ver_zoom_va0
    s_cargar_tipo_zoom_ma0
    s_fijar_separacion_mapa_ma0
    
    'Copio las dimensiones del mapa
    copia_dim_va0_2_ma0_va0
    ReDim mapa_ma0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Boolean
    
    'Copio el mapa
    mapa_sin_obstaculos_ma0 = mapa_sin_obstaculos_va0
    If Not mapa_sin_obstaculos_va0 Then
        Screen.MousePointer = CTE_ARENA
        For p = 1 To mapa_pisos_ma0
        For f = 1 To mapa_filas_ma0
        For c = 1 To mapa_columnas_ma0
                If mapa_va0(p, f, c) = CTE_MAPA_OBSTACULO Then
                    mapa_ma0(p, f, c) = True
                Else
                    mapa_ma0(p, f, c) = False
                End If
        Next c
        Next f
        Next p
    End If
    Screen.MousePointer = CTE_DEFECTO


End Sub
Sub s_mapa_pintar_cursor_va0()

    direccion_old_test_va0 = direccion_test_va0
    cursor_old_z_va0 = cursor_z_va0
    cursor_old_y_va0 = cursor_y_va0
    cursor_old_x_va0 = cursor_x_va0
    num_direcc_old_algoritmo_va0 = num_direcc_algoritmo_va0

    cursor_z_va0 = CInt("0" & frm_a0_mapa.Op_MapaEjeZ.Text)
    cursor_y_va0 = CInt("0" & frm_a0_mapa.Op_MapaEjeY.Text)
    cursor_x_va0 = CInt("0" & frm_a0_mapa.Op_MapaEjeX.Text)
    
    s_mapa_cursor_pintado_va0 cursor_old_z_va0, cursor_old_y_va0, cursor_old_x_va0, direccion_old_test_va0, num_direcc_old_algoritmo_va0, cct_ejv(cfondo_ejv)
    s_mapa_cursor_pintado_va0 cursor_z_va0, cursor_y_va0, cursor_x_va0, direccion_test_va0, num_direcc_algoritmo_va0, cct_ejv(CTE_NEGRO)
    

End Sub

Sub s_mapa_parpadeo_cursor_va0()

    If cursor_visible_mapa_ma0 Then
        s_mapa_cursor_pintado_va0 cursor_z_va0, cursor_y_va0, cursor_x_va0, direccion_test_va0, num_direcc_algoritmo_va0, cct_ejv(cfondo_ejv)
        cursor_visible_mapa_ma0 = False
    Else
        s_mapa_cursor_pintado_va0 cursor_z_va0, cursor_y_va0, cursor_x_va0, direccion_test_va0, num_direcc_algoritmo_va0, cct_ejv(CTE_NEGRO)
        cursor_visible_mapa_ma0 = True
    End If

End Sub

Sub s_mapa_cursor_pintado_va0(pos_z As Double, pos_y As Double, pos_x As Double, p_direcc As Integer, p_num_direcc As Integer, p_color As Long)

    Dim pintar_z As Integer
    Dim pintar_y As Integer
    Dim pintar_x As Integer
    Dim aparta As Integer
    
    'Solo pinto el pirindolo de la direccion en modo detalle
    If ver_zoom_ma0 = CTE_ZOOM_DETALLE Then
    
        aparta = 7
        
        pintar_z = CTE_MAPA_INI_Z + (separacion_mapa_ma0 * pos_y)
        pintar_y = CTE_MAPA_INI_Y + (separacion_mapa_ma0 * pos_y)
        pintar_x = CTE_MAPA_INI_X + (separacion_mapa_ma0 * pos_x)
    
        If p_direcc <> 0 Then
            frm_a0_mapa.ScaleMode = vbPixels   ' Set scale to pixels.
            frm_a0_mapa.FillColor = cct_ejv(cfondo_ejv)
            Select Case p_num_direcc
                Case 4
                    Select Case p_direcc
                        Case CTE_NORTE
                            pintar_y = pintar_y - aparta
                            pintar_x = pintar_x
                        Case CTE_ESTE
                            pintar_y = pintar_y
                            pintar_x = pintar_x + aparta
                        Case CTE_SUR
                            pintar_y = pintar_y + aparta
                            pintar_x = pintar_x
                        Case CTE_OESTE
                            pintar_y = pintar_y
                            pintar_x = pintar_x - aparta
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
                    End Select
                Case 8
                    Select Case p_direcc
                        Case CTE_8_DEF
                            pintar_y = pintar_y - aparta
                            pintar_x = pintar_x
                        Case CTE_8_DEF_DER
                            pintar_y = pintar_y - aparta
                            pintar_x = pintar_x + aparta
                        Case CTE_8_DER
                            pintar_y = pintar_y
                            pintar_x = pintar_x + aparta
                        Case CTE_8_ATR_DER
                            pintar_y = pintar_y + aparta
                            pintar_x = pintar_x + aparta
                        Case CTE_8_ATR
                            pintar_y = pintar_y + aparta
                            pintar_x = pintar_x
                        Case CTE_8_ATR_IZQ
                            pintar_y = pintar_y + aparta
                            pintar_x = pintar_x - aparta
                        Case CTE_8_IZQ
                            pintar_y = pintar_y
                            pintar_x = pintar_x - aparta
                        Case CTE_8_DEF_IZQ
                            pintar_y = pintar_y - aparta
                            pintar_x = pintar_x - aparta
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
                    End Select
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
            End Select
        End If
    End If
    
    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, 1, pos_y, pos_x, CTE_CUADRADOCURSOR, p_color, p_color, CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
    If p_direcc <> 0 And ver_zoom_ma0 = CTE_ZOOM_DETALLE Then
        frm_a0_mapa.Circle (pintar_x, pintar_y), 1, p_color
    End If


End Sub
Sub s_fijar_separacion_mapa_ma0()
    
    Select Case ver_zoom_ma0
        Case CTE_ZOOM_DETALLE
            separacion_mapa_ma0 = 14
        Case CTE_ZOOM_PANORAMICA
            separacion_mapa_ma0 = 4
        Case CTE_ZOOM_PIXELS
            separacion_mapa_ma0 = 1
        Case CTE_ZOOM_3D
            separacion_mapa_ma0 = 20
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
    End Select

End Sub

Sub s_cargar_tipo_zoom_va0()

    frm_z0_mdi.Cb_Zoom.ListIndex = ver_zoom_va0

End Sub

Sub s_cargar_tipo_zoom_ma0()

    frm_z0_mdi.Cb_Zoom.ListIndex = ver_zoom_ma0
    If ver_zoom_ma0 = CTE_ZOOM_3D Then
        cursor_visible_mapa_ma0 = False
        'frm_a0_mapa.Ch_Parpadeo.Value = 0 no porque obliga a cargar el form en VA
    End If


End Sub
Sub s_test_crecimiento_ma0()

    Dim cont_ciclos As Integer
    Dim p As Double
    Dim f As Double
    Dim c As Double
    ReDim mapa_temp(1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Boolean
    
    Screen.MousePointer = CTE_ARENA
    'Pasamos a modo test
    estado_test_movimiento_va0 = True
    
    p = 1
    
    For cont_ciclos = 1 To CInt(frm_a0_testm.num_ciclos.Text)
        'Copio el array en el temporal
        For f = 1 To mapa_filas_ma0
        For c = 1 To mapa_columnas_ma0
            mapa_temp(f, c) = mapa_ma0(p, f, c)
        Next c
        Next f
        'Por cada obstaculo en el array inicial
        For f = 1 To mapa_filas_ma0
        For c = 1 To mapa_columnas_ma0
            If mapa_ma0(p, f, c) Then
                'Si hay espacio para crecer, crece
                expandir_ma0 f, c, mapa_temp()
            End If
        Next c
        Next f
        'Actualizo en el array el nuevo estado del automata
        For f = 1 To mapa_filas_ma0
        For c = 1 To mapa_columnas_ma0
            mapa_ma0(p, f, c) = mapa_temp(f, c)
            'Lo pinto
            If mapa_ma0(p, f, c) Then
                If frm_a0_testm.Op_CreandoObstaculosF = 1 Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_NEGRO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                Else
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_ROSA), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                End If
            End If
        Next c
        Next f
    Next cont_ciclos
    
    'Salimos de modo test
    estado_test_movimiento_va0 = False
    Screen.MousePointer = CTE_DEFECTO

End Sub

Sub expandir_ma0(fil As Double, col As Double, mapa_temp() As Boolean)
 
    Dim i As Integer
    Dim num_vacios As Integer
    Dim cont_vacios As Integer
    Dim num_vacios_necesarios As Integer
    Dim direccion As Integer
    Dim hay_suficientes_huecos As Boolean
    Dim anterior As Integer
    ReDim vecinos(1 To CTE_8_DIR) As Boolean
    
    Dim vecino_piso As Double
    Dim vecino_fila As Double
    Dim vecino_col As Double
    
    Dim empezar As Integer
    Dim actual As Integer
    Dim medio As Integer
        
    
    Select Case frm_a0_testm.Cb_Metodo.ListIndex
        Case 0
            'Método 1
            num_vacios_necesarios = 3
        Case 1
            'Método 2
            num_vacios_necesarios = 4
        Case 2
            'Método 3
            num_vacios_necesarios = 5
        Case 3
            'Método 4: hoja
            'Se ejecuta por otro lado
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Método inexistente"
            Exit Sub
    End Select


    'Analizo los vecinos
    For i = 1 To CTE_8_DIR
        vecinos(i) = False
    Next i
    num_vacios = 0
    For i = 1 To CTE_8_DIR
        'Me coloco en la posicion inicial
        vecino_fila = fil
        vecino_col = col
        'Avanzo
        f_avanzo_8_dir_va0 vecino_piso, vecino_fila, vecino_col, i, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
        'Guardo que hay ahí
        vecinos(i) = mapa_temp(vecino_col, vecino_fila)
        If vecinos(i) = False Then
            num_vacios = num_vacios + 1
        End If
    Next i
        
    
    hay_suficientes_huecos = True
    While hay_suficientes_huecos
        'Busco uno por el que empiezo a analizar
        'Si no hay los huecos necesarios en total, seguro que no los hay juntos
        If num_vacios >= num_vacios_necesarios Then
            'Si todos vacios, empiezo por cualquiera
            If num_vacios = 8 Then
                empezar = 1
            Else
                'Empiezo por uno ocupado para coger trozos vacios completos
                empezar = 1
                While Not vecinos(empezar)
                    empezar = empezar + 1
                Wend
            End If
            cont_vacios = 0
            actual = 0
            For i = empezar To empezar + 7
                actual = f_SumCirc(8, i, 0)
                If Not vecinos(actual) Then
                    cont_vacios = cont_vacios + 1
                End If
                If cont_vacios >= num_vacios_necesarios Then
                    'Pongo el del medio del hueco ocupado
                    Select Case frm_a0_testm.Cb_Metodo.ListIndex
                        Case 0
                            'Método 1
                            medio = f_SumCirc(8, actual, 8 - 1) 'uno anterior, son 3
                        Case 1
                            'Método 2
                            medio = f_SumCirc(8, actual, 8 - 1) 'uno anterior, son 4
                        Case 2
                            'Método 3
                            medio = f_SumCirc(8, actual, 8 - 2) 'dos anterior, son 5
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Método inexistente"
                            Exit Sub
                    End Select
                    'Lo pinto
                    vecinos(medio) = True
                    'Me coloco en la posicion inicial
                    vecino_fila = fil
                    vecino_col = col
                    'Calculo la xy del vecino vacio
                    f_avanzo_8_dir_va0 vecino_piso, vecino_fila, vecino_col, medio, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
                    mapa_temp(vecino_col, vecino_fila) = True
                    num_vacios = num_vacios - 1
                    'supongo que puede haber mas series vacias y vuelvo a empezar
                    hay_suficientes_huecos = True
                    Exit For
                Else
                    'no hay como crecer
                    hay_suficientes_huecos = False
                End If
            Next i
        Else
            hay_suficientes_huecos = False
        End If
    Wend
        
End Sub
Sub s_insertar_hoja_ma0()

    'Pido confirmacion
    'Borro el mapa
    'Pinto una hoja
'------------------------------------------------------------------------
' Código adaptado de una adaptación de
' Charles Dumont de uno de los
' famosos algoritmos de fractales
' dumonfc1@SPACEMSG.JHUAPL.edu
' Gracias y un saludete!
'------------------------------------------------------------------------
    
    Dim X, Y As Single
    Dim S As Single
    Dim NewX, NewY As Single
    Dim Iterations As Single
    Dim Transforms(1 To 4, 1 To 7) As Single
    Dim p(0 To 1024, 0 To 721) As Integer
    Dim i, j As Long
    Dim px, py, RandNum As Double
    
    Screen.MousePointer = CTE_ARENA
    
    'Initialize pixel color matrix
    For i = 0 To 1024
        For j = 0 To 721
            p(i, j) = 75
        Next
    Next
    'Set scale factor
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

    'Number of points
    Iterations = CDbl(frm_a0_testm.num_ciclos.Text)


    Randomize
    For i = 1 To Iterations
        ' Readjust coordinates to center of screen and with
        ' larger scale
        py = Int(S * X + mapa_filas_ma0 / 2)
        px = Int(mapa_columnas_ma0 - S * Y)
        ' Increase color value of pixel
        If (py < mapa_filas_ma0) And (py > 0) And (px > 0) And (px < mapa_columnas_ma0) Then
            p(py, px) = p(py, px) + 1
            If p(py, px) > 255 Then p(py, px) = 255
        End If
        ' Color current point
        If py > 0 And px > 0 Then
            If py <= mapa_filas_ma0 And px <= mapa_columnas_ma0 Then
                mapa_ma0(py, px) = True
                If frm_a0_testm.Op_CreandoObstaculosF = 1 Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, 1, CInt(px), CInt(py), CTE_CUBO, cct_ejv(CTE_NEGRO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                Else
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, 1, CInt(px), CInt(py), CTE_CUBO, cct_ejv(CTE_ROSA), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
                End If
                'PSet (px, py), RGB(0, p(px, py), 0)
                'Me.FillColor = RGB(0, p(px, py), 0)
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
    Next i
    Screen.MousePointer = CTE_DEFECTO
    

End Sub

Sub s_mover_cursor(direccion4 As Integer)

    Select Case direccion4
        Case CTE_DEFRENTE
            frm_a0_mapa.Op_MapaEjeY = frm_a0_mapa.Op_MapaEjeY - 1
            If frm_a0_mapa.Op_MapaEjeY <= 0 Then frm_a0_mapa.Op_MapaEjeY = mapa_filas_ma0 + frm_a0_mapa.Op_MapaEjeY
            If frm_a0_mapa.Op_MapaEjeY >= mapa_filas_ma0 + 1 Then frm_a0_mapa.Op_MapaEjeY = frm_a0_mapa.Op_MapaEjeY - mapa_filas_ma0
        Case CTE_DERECHA
            frm_a0_mapa.Op_MapaEjeX = frm_a0_mapa.Op_MapaEjeX + 1
            If frm_a0_mapa.Op_MapaEjeX <= 0 Then frm_a0_mapa.Op_MapaEjeX = mapa_columnas_ma0 + frm_a0_mapa.Op_MapaEjeX
            If frm_a0_mapa.Op_MapaEjeX >= mapa_columnas_ma0 + 1 Then frm_a0_mapa.Op_MapaEjeX = frm_a0_mapa.Op_MapaEjeX - mapa_columnas_ma0
        Case CTE_ATRAS
            frm_a0_mapa.Op_MapaEjeY = frm_a0_mapa.Op_MapaEjeY + 1
            If frm_a0_mapa.Op_MapaEjeY <= 0 Then frm_a0_mapa.Op_MapaEjeY = mapa_filas_ma0 + frm_a0_mapa.Op_MapaEjeY
            If frm_a0_mapa.Op_MapaEjeY >= mapa_filas_ma0 + 1 Then frm_a0_mapa.Op_MapaEjeY = frm_a0_mapa.Op_MapaEjeY - mapa_filas_ma0
        Case CTE_IZQUIERDA
            frm_a0_mapa.Op_MapaEjeX = frm_a0_mapa.Op_MapaEjeX - 1
            If frm_a0_mapa.Op_MapaEjeX <= 0 Then frm_a0_mapa.Op_MapaEjeX = mapa_columnas_ma0 + frm_a0_mapa.Op_MapaEjeX
            If frm_a0_mapa.Op_MapaEjeX >= mapa_columnas_ma0 + 1 Then frm_a0_mapa.Op_MapaEjeX = frm_a0_mapa.Op_MapaEjeX - mapa_columnas_ma0
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Dirección inexistente"
    End Select
    
End Sub


Sub s_cambio_zoom_ma0()
    
    If habilitar_change_zoom_va0 Then
        ver_zoom_ma0 = frm_z0_mdi.Cb_Zoom.ListIndex
        'Muestro el nuevo mapa
        s_fijar_separacion_mapa_ma0
        s_refrescar_mapa_actual_ma0
    End If

End Sub

Sub s_mapa_pintar_bordes_ma0()

    Dim p As Double
    Dim f As Double
    Dim c As Double

    p = 1
    'Pintamos bordes
    Screen.MousePointer = CTE_ARENA
    For f = 0 To mapa_filas_ma0 + 1
        If ver_zoom_ma0 = CTE_ZOOM_3D Then
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, 0, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, CTE_ZOOM_DETALLE, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, mapa_columnas_ma0 + 1, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, CTE_ZOOM_DETALLE, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, 0, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, mapa_columnas_ma0 + 1, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        Else
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, 0, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, mapa_columnas_ma0 + 1, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        End If
    Next f
    For c = 0 To mapa_columnas_ma0 + 1
        If ver_zoom_ma0 = CTE_ZOOM_3D Then
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, 0, c, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, CTE_ZOOM_DETALLE, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, mapa_filas_ma0 + 1, c, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, CTE_ZOOM_DETALLE, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, 0, c, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, mapa_filas_ma0 + 1, c, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        Else
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, 0, c, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, mapa_filas_ma0 + 1, c, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        End If
    Next c
    Screen.MousePointer = CTE_DEFECTO

End Sub

Sub s_refrescar_mapa_actual_ma0()

    frm_a0_mapa.Cls
    s_mapa_pintar_bordes_ma0
    s_mostrar_mapa_actual_ma0 False, ver_zoom_ma0

End Sub

Sub s_mostrar_mapa_actual_ma0(pintar_todo As Boolean, p_zoom As Integer)
    
    Dim p As Double
    Dim f As Double
    Dim c As Double
    Dim cont_obj As Integer
    Dim total_obj As Integer
    ReDim obj_pintar_p(1 To 1) As Double
    ReDim obj_pintar_f(1 To 1) As Double
    ReDim obj_pintar_c(1 To 1) As Double
    ReDim obj_pintar_o(1 To 1) As Integer
   
     
    If mapa_sin_obstaculos_ma0 Then Exit Sub
   
    'pintar_todo dice si se pintan tambien los huecos
    'no hay doevents porque va mucho mas lento
    Screen.MousePointer = CTE_ARENA
    
    If ver_zoom_ma0 = CTE_ZOOM_3D Then
        'Guardo los objetos a pintar en un array
        cont_obj = 0
        If UBound(mapa_ma0, 1) > 0 Then
            For p = 1 To mapa_pisos_ma0
            For f = 1 To mapa_filas_ma0
            For c = 1 To mapa_columnas_ma0
                If mapa_ma0(p, f, c) = True Then
                    cont_obj = cont_obj + 1
                    ReDim Preserve obj_pintar_p(1 To cont_obj) As Double
                    ReDim Preserve obj_pintar_f(1 To cont_obj) As Double
                    ReDim Preserve obj_pintar_c(1 To cont_obj) As Double
                    ReDim Preserve obj_pintar_o(1 To cont_obj) As Integer
                    obj_pintar_p(cont_obj) = p
                    obj_pintar_f(cont_obj) = f
                    obj_pintar_c(cont_obj) = c
                    obj_pintar_o(cont_obj) = CTE_MAPA_OBSTACULO
                End If
            Next c
            Next f
            Next p
        End If
        total_obj = cont_obj
        
        'Ordeno el array por lejania hasta el observador
        S_OrdenarArray3DIntMinMax obj_pintar_p(), obj_pintar_f(), obj_pintar_c(), obj_pintar_o()
        
        'Pinto el array en ese orden
        For cont_obj = 1 To total_obj
            If obj_pintar_o(cont_obj) = CTE_MAPA_OBSTACULO Then
                If p_zoom = CTE_ZOOM_3D Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, obj_pintar_p(cont_obj), obj_pintar_f(cont_obj), obj_pintar_c(cont_obj), CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, CTE_ZOOM_DETALLE, 1
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, obj_pintar_p(cont_obj), obj_pintar_f(cont_obj), obj_pintar_c(cont_obj), CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, CTE_ZOOM_3D, 1
                Else
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, obj_pintar_p(cont_obj), obj_pintar_f(cont_obj), obj_pintar_c(cont_obj), CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, p_zoom, 1
                End If
            Else
                If pintar_todo Then
                    If p_zoom = CTE_ZOOM_3D Then
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, obj_pintar_p(cont_obj), obj_pintar_f(cont_obj), obj_pintar_c(cont_obj), CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, CTE_ZOOM_DETALLE, 1
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, obj_pintar_p(cont_obj), obj_pintar_f(cont_obj), obj_pintar_c(cont_obj), CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, p_zoom, 1
                    Else
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, obj_pintar_p(cont_obj), obj_pintar_f(cont_obj), obj_pintar_c(cont_obj), CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, p_zoom, 1
                    End If
                End If
            End If
        Next cont_obj
    Else
        'No es en 3D
        If UBound(mapa_ma0, 1) > 0 Then
            For p = 1 To mapa_pisos_ma0
            For f = 1 To mapa_filas_ma0
            For c = 1 To mapa_columnas_ma0
                If mapa_ma0(p, f, c) = True Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, p_zoom, 1
                Else
                    If pintar_todo Then
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, p_zoom, 1
                    End If
                End If
            Next c
            Next f
            Next p
        End If
    End If
    Screen.MousePointer = CTE_DEFECTO


End Sub

