Attribute VB_Name = "bas_a0_mov"
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
Function f_alg1_calcular_siguiente_nodo_va0(nodo_z As Double, nodo_y As Double, nodo_x As Double, direccion As Integer, max_pisos As Double, max_filas As Double, max_col As Double) As Integer
    'Los parametros son de entrada/salida
    
    'busqueda_4_espiral_con_obstaculos
    'Hay 4 direcciones posibles
    'Tiene tendencia a hacer giros a la derecha
    'Derecha-Frente-Izquierda-Atrás

    Dim orden_exploracion(1 To CTE_4_DIR) As Integer
    Dim estoy_rodeado As Boolean

    'elijo un orden de exploracion
    orden_exploracion(1) = CTE_DERECHA
    orden_exploracion(2) = CTE_DEFRENTE
    orden_exploracion(3) = CTE_IZQUIERDA
    orden_exploracion(4) = CTE_ATRAS

    estoy_rodeado = f_alg_general_4_direcciones_va0(nodo_z, nodo_y, nodo_x, direccion, orden_exploracion(), max_pisos, max_filas, max_col)
    f_alg1_calcular_siguiente_nodo_va0 = estoy_rodeado

End Function
Function f_alg2_calcular_siguiente_nodo_va0(nodo_z As Double, nodo_y As Double, nodo_x As Double, direccion As Integer, max_pisos As Double, max_filas As Double, max_col As Double) As Integer
    'Los parametros son de entrada/salida
    
    'busqueda_4_defrente_con_obstaculos
    'Hay 4 direcciones posibles
    'Tiene tendencia a hacer rectas
    'Frente-Derecha-Izquierda-Atrás

    Dim orden_exploracion(1 To CTE_4_DIR) As Integer
    Dim estoy_rodeado As Boolean

    'elijo un orden de exploracion
    orden_exploracion(1) = CTE_DEFRENTE
    orden_exploracion(2) = CTE_DERECHA
    orden_exploracion(3) = CTE_IZQUIERDA
    orden_exploracion(4) = CTE_ATRAS

    estoy_rodeado = f_alg_general_4_direcciones_va0(nodo_z, nodo_y, nodo_x, direccion, orden_exploracion(), max_pisos, max_filas, max_col)
    f_alg2_calcular_siguiente_nodo_va0 = estoy_rodeado

End Function
Function f_alg3_calcular_siguiente_nodo_va0(nodo_z As Double, nodo_y As Double, nodo_x As Double, direccion As Integer, max_pisos As Double, max_filas As Double, max_col As Double) As Integer
    'Los parametros son de entrada/salida
    
    'busqueda_8_espiral_con_obstaculos
    'Hay 8 direcciones posibles
    'Tiene tendencia a hacer giros a la derecha

    Dim orden_exploracion(1 To CTE_8_DIR) As Integer
    Dim estoy_rodeado As Boolean

    'elijo un orden de exploracion
    orden_exploracion(1) = CTE_8_DEF_DER
    orden_exploracion(2) = CTE_8_DER
    orden_exploracion(3) = CTE_8_DEF
    orden_exploracion(4) = CTE_8_DEF_IZQ
    orden_exploracion(5) = CTE_8_IZQ
    orden_exploracion(6) = CTE_8_ATR_DER ' estos 3 ultimos estan a boleo
    orden_exploracion(7) = CTE_8_ATR_IZQ
    orden_exploracion(8) = CTE_8_ATR

    estoy_rodeado = f_alg_general_8_direcciones_va0(nodo_z, nodo_y, nodo_x, direccion, orden_exploracion(), max_pisos, max_filas, max_col)
    f_alg3_calcular_siguiente_nodo_va0 = estoy_rodeado

End Function

Function f_alg_general_4_direcciones_va0(nodo_z As Double, nodo_y As Double, nodo_x As Double, direccion As Integer, orden_exploracion() As Integer, max_pisos As Double, max_filas As Double, max_col As Double) As Boolean
    'Los parametros son de entrada/salida
    
    'Hay 4 direcciones posibles
    
    'Los nodos-lugares visitados los marco en el array nodo_visitado_va0(Y, X)
    Dim estoy_rodeado As Boolean

    Dim viejo_y As Integer
    Dim viejo_x As Integer
    Dim viejo_direccion As Integer
    Dim esta_vacio(1 To CTE_4_DIR) As Boolean

    'CTE_DEFRENTE = 1
    'CTE_DERECHA = 2
    'CTE_ATRAS = 3
    'CTE_IZQUIERDA = 4

    ReDim n_vec_visit(1 To CTE_4_DIR) As Integer
    ReDim lista_sin_obstaculos(1 To 1) As Integer

    Dim este_es_el_menor As Boolean
    Dim cont1 As Integer
    Dim cont2 As Integer
    Dim ultimo As Integer
    Dim i As Integer
    
    estoy_rodeado = False
   
    'primero busco nodos sin obstaculos y no visitados
    'si no hay, elijo de los visitados, el menos visitado,
    'lo marco como visitado una vez mas,
    'y una vez alli vuelvo a comprobar todo desde el principio
    
    'Inicializo los arrays
    For i = 1 To CTE_4_DIR
        esta_vacio(i) = True
    Next i
    'Guardo la posicion actual
    viejo_y = nodo_y
    viejo_x = nodo_x
    viejo_direccion = direccion
    
    '=============Intento 1=============
    'Intento ir en la direccion orden_exploracion(1)
    'giro
    direccion = f_giro_4_general_va0(direccion, orden_exploracion(1))
    'avanzo
    f_avanzo_4_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
    'compruebo la nueva posicion
    esta_vacio(orden_exploracion(1)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
    n_vec_visit(orden_exploracion(1)) = nodo_visitado_va0(1, nodo_y, nodo_x)
    If esta_vacio(orden_exploracion(1)) And n_vec_visit(orden_exploracion(1)) = 0 Then
        'Ya lo he encontrado, me salgo
        nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
    Else
        '=============Intento 2=============
        'No se ha podido girar, vuelvo al anterior
        'primero vuelvo a la direccion inicial
        nodo_y = viejo_y
        nodo_x = viejo_x
        direccion = viejo_direccion
        'Intento ir en la direccion orden_exploracion(2)
        'giro
        direccion = f_giro_4_general_va0(direccion, orden_exploracion(2))
        'avanzo
        f_avanzo_4_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
        'compruebo la nueva posicion
        esta_vacio(orden_exploracion(2)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
        n_vec_visit(orden_exploracion(2)) = nodo_visitado_va0(1, nodo_y, nodo_x)
        If esta_vacio(orden_exploracion(2)) And n_vec_visit(orden_exploracion(2)) = 0 Then
            'Ya lo he encontrado, me salgo
             nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
        Else
            '=============Intento 3=============
            'primero vuelvo a la direccion inicial
            nodo_y = viejo_y
            nodo_x = viejo_x
            direccion = viejo_direccion
            'Intento ir en la direccion orden_exploracion(3)
            'giro
            direccion = f_giro_4_general_va0(direccion, orden_exploracion(3))
            'avanzo
            f_avanzo_4_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
            'compruebo la nueva posicion
            esta_vacio(orden_exploracion(3)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
            n_vec_visit(orden_exploracion(3)) = nodo_visitado_va0(1, nodo_y, nodo_x)
            If esta_vacio(orden_exploracion(3)) And n_vec_visit(orden_exploracion(3)) = 0 Then
                'Ya lo he encontrado, me salgo
                 nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
            Else
                '=============Intento 4=============
                'primero vuelvo a la direccion inicial
                nodo_y = viejo_y
                nodo_x = viejo_x
                direccion = viejo_direccion
                'Intento ir en la direccion orden_exploracion(4)
                'giro
                direccion = f_giro_4_general_va0(direccion, orden_exploracion(4))
                'avanzo
                f_avanzo_4_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
                'compruebo la nueva posicion
                esta_vacio(orden_exploracion(4)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
                n_vec_visit(orden_exploracion(4)) = nodo_visitado_va0(1, nodo_y, nodo_x)
                If esta_vacio(orden_exploracion(4)) And n_vec_visit(orden_exploracion(4)) = 0 Then
                    'Ya lo he encontrado, me salgo
                     nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                Else
                    '=============Fin Intentos=============
                    'Si no se puede en ninguna de las direcciones,
                    'escojo el menos visitado de los que no sean abstaculo
                    'primero vuelvo a la direccion inicial
                    nodo_y = viejo_y
                    nodo_x = viejo_x
                    direccion = viejo_direccion
                    ultimo = 0
                    If esta_vacio(orden_exploracion(1)) Then
                        ultimo = ultimo + 1
                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                        lista_sin_obstaculos(ultimo) = orden_exploracion(1)
                    End If
                    If esta_vacio(orden_exploracion(2)) Then
                        ultimo = ultimo + 1
                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                        lista_sin_obstaculos(ultimo) = orden_exploracion(2)
                    End If
                    If esta_vacio(orden_exploracion(3)) Then
                        ultimo = ultimo + 1
                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                        lista_sin_obstaculos(ultimo) = orden_exploracion(3)
                    End If
                    If esta_vacio(orden_exploracion(4)) Then
                        ultimo = ultimo + 1
                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                        lista_sin_obstaculos(ultimo) = orden_exploracion(4)
                    End If
                    If ultimo = 0 Then
                        'No hay ningun sitio vacio alrededor, estoy rodeado
                        'Me salgo, pero aviso que no puedo moverme
                         estoy_rodeado = True
                         nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                    Else
                        'Hay alguno libre, aunque ya visitado
                        'Cojo el menos visitado
                        'Analizo cada uno
                        For cont1 = 1 To ultimo
                            'Lo comparo con el resto
                            este_es_el_menor = True
                            For cont2 = 1 To ultimo
                                If n_vec_visit(lista_sin_obstaculos(cont1)) > n_vec_visit(lista_sin_obstaculos(cont2)) Then
                                    'Este no es el menor
                                    este_es_el_menor = False
                                    Exit For
                                End If
                            Next cont2
                            If este_es_el_menor Then
                                'lista_sin_obstaculos(cont1) es el menor
                                'Elijo esa direccion, es decir, lista_sin_obstaculos(cont1)
                                direccion = f_giro_4_general_va0(direccion, lista_sin_obstaculos(cont1))
                                'avanzo
                                f_avanzo_4_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
                                'control errores de programacion
                                If control_errores_de_programacion_ejv Then
                                    If Not f_esta_vacio_va0(1, nodo_y, nodo_x) Then
                                        s_error_ejv CON_OPCION_FINALIZAR, "Error: no está vacío"
                                    End If
                                End If
                                nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                                Exit For
                            End If
                        Next cont1
                        If Not este_es_el_menor Then
                            'alguno tiene que ser menor o igual!
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: no es el menor"
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    
    f_alg_general_4_direcciones_va0 = estoy_rodeado
    

End Function

Function f_alg_general_8_direcciones_va0(nodo_z As Double, nodo_y As Double, nodo_x As Double, direccion As Integer, orden_exploracion() As Integer, max_pisos As Double, max_filas As Double, max_col As Double) As Boolean
    'Los parametros son de entrada/salida
    
    'Hay 8 direcciones posibles
    
    'Los nodos-lugares visitados los marco en el array nodo_visitado_va0(Y, X)
    Dim estoy_rodeado As Boolean

    Dim viejo_y As Integer
    Dim viejo_x As Integer
    Dim viejo_direccion As Integer
    Dim esta_vacio(1 To CTE_8_DIR) As Boolean

    'CTE_8_DEF = 1
    'CTE_8_DEF_DER = 2
    'CTE_8_DER = 3
    'CTE_8_ATR_DER = 4
    'CTE_8_ATR = 5
    'CTE_8_ATR_IZQ = 6
    'CTE_8_IZQ = 7
    'CTE_8_DEF_IZQ = 8

    ReDim n_vec_visit(1 To CTE_8_DIR) As Integer
    ReDim lista_sin_obstaculos(1 To 1) As Integer

    Dim este_es_el_menor As Boolean
    Dim cont1 As Integer
    Dim cont2 As Integer
    Dim ultimo As Integer
    Dim i As Integer
    
    estoy_rodeado = False
   
    'primero busco nodos sin obstaculos y no visitados
    'si no hay, elijo de los visitados, el menos visitado,
    'lo marco como visitado una vez mas,
    'y una vez alli vuelvo a comprobar todo desde el principio
    
    'Inicializo los arrays
    For i = 1 To CTE_8_DIR
        esta_vacio(i) = True
    Next i
    'Guardo la posicion actual
    viejo_y = nodo_y
    viejo_x = nodo_x
    viejo_direccion = direccion
    
    '=============Intento 1=============
    'Intento ir en la direccion orden_exploracion(1)
    'giro
    direccion = f_giro_8_general_va0(direccion, orden_exploracion(1))
    'avanzo
    f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
    'compruebo la nueva posicion
    esta_vacio(orden_exploracion(1)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
    n_vec_visit(orden_exploracion(1)) = nodo_visitado_va0(1, nodo_y, nodo_x)
    If esta_vacio(orden_exploracion(1)) And n_vec_visit(orden_exploracion(1)) = 0 Then
        'Ya lo he encontrado, me salgo
        nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
    Else
        '=============Intento 2=============
        'No se ha podido girar, vuelvo al anterior
        'primero vuelvo a la direccion inicial
        nodo_y = viejo_y
        nodo_x = viejo_x
        direccion = viejo_direccion
        'Intento ir en la direccion orden_exploracion(2)
        'giro
        direccion = f_giro_8_general_va0(direccion, orden_exploracion(2))
        'avanzo
        f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
        'compruebo la nueva posicion
        esta_vacio(orden_exploracion(2)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
        n_vec_visit(orden_exploracion(2)) = nodo_visitado_va0(1, nodo_y, nodo_x)
        If esta_vacio(orden_exploracion(2)) And n_vec_visit(orden_exploracion(2)) = 0 Then
            'Ya lo he encontrado, me salgo
             nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
        Else
            '=============Intento 3=============
            'primero vuelvo a la direccion inicial
            nodo_y = viejo_y
            nodo_x = viejo_x
            direccion = viejo_direccion
            'Intento ir en la direccion orden_exploracion(3)
            'giro
            direccion = f_giro_8_general_va0(direccion, orden_exploracion(3))
            'avanzo
            f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
            'compruebo la nueva posicion
            esta_vacio(orden_exploracion(3)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
            n_vec_visit(orden_exploracion(3)) = nodo_visitado_va0(1, nodo_y, nodo_x)
            If esta_vacio(orden_exploracion(3)) And n_vec_visit(orden_exploracion(3)) = 0 Then
                'Ya lo he encontrado, me salgo
                 nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
            Else
                '=============Intento 4=============
                'primero vuelvo a la direccion inicial
                nodo_y = viejo_y
                nodo_x = viejo_x
                direccion = viejo_direccion
                'Intento ir en la direccion orden_exploracion(4)
                'giro
                direccion = f_giro_8_general_va0(direccion, orden_exploracion(4))
                'avanzo
                f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
                'compruebo la nueva posicion
                esta_vacio(orden_exploracion(4)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
                n_vec_visit(orden_exploracion(4)) = nodo_visitado_va0(1, nodo_y, nodo_x)
                If esta_vacio(orden_exploracion(4)) And n_vec_visit(orden_exploracion(4)) = 0 Then
                    'Ya lo he encontrado, me salgo
                     nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                Else
                    '=============Intento 5=============
                    'primero vuelvo a la direccion inicial
                    nodo_y = viejo_y
                    nodo_x = viejo_x
                    direccion = viejo_direccion
                    'Intento ir en la direccion orden_exploracion(5)
                    'giro
                    direccion = f_giro_8_general_va0(direccion, orden_exploracion(5))
                    'avanzo
                    f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
                    'compruebo la nueva posicion
                    esta_vacio(orden_exploracion(5)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
                    n_vec_visit(orden_exploracion(5)) = nodo_visitado_va0(1, nodo_y, nodo_x)
                    If esta_vacio(orden_exploracion(5)) And n_vec_visit(orden_exploracion(5)) = 0 Then
                        'Ya lo he encontrado, me salgo
                         nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                    Else
                        '=============Intento 6=============
                        'primero vuelvo a la direccion inicial
                        nodo_y = viejo_y
                        nodo_x = viejo_x
                        direccion = viejo_direccion
                        'Intento ir en la direccion orden_exploracion(6)
                        'giro
                        direccion = f_giro_8_general_va0(direccion, orden_exploracion(6))
                        'avanzo
                        f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
                        'compruebo la nueva posicion
                        esta_vacio(orden_exploracion(6)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
                        n_vec_visit(orden_exploracion(6)) = nodo_visitado_va0(1, nodo_y, nodo_x)
                        If esta_vacio(orden_exploracion(6)) And n_vec_visit(orden_exploracion(6)) = 0 Then
                            'Ya lo he encontrado, me salgo
                             nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                        Else
                            '=============Intento 7=============
                            'primero vuelvo a la direccion inicial
                            nodo_y = viejo_y
                            nodo_x = viejo_x
                            direccion = viejo_direccion
                            'Intento ir en la direccion orden_exploracion(7)
                            'giro
                            direccion = f_giro_8_general_va0(direccion, orden_exploracion(7))
                            'avanzo
                            f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
                            'compruebo la nueva posicion
                            esta_vacio(orden_exploracion(7)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
                            n_vec_visit(orden_exploracion(7)) = nodo_visitado_va0(1, nodo_y, nodo_x)
                            If esta_vacio(orden_exploracion(7)) And n_vec_visit(orden_exploracion(7)) = 0 Then
                                'Ya lo he encontrado, me salgo
                                 nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                            Else
                                '=============Intento 8=============
                                'primero vuelvo a la direccion inicial
                                nodo_y = viejo_y
                                nodo_x = viejo_x
                                direccion = viejo_direccion
                                'Intento ir en la direccion orden_exploracion(8)
                                'giro
                                direccion = f_giro_8_general_va0(direccion, orden_exploracion(8))
                                'avanzo
                                f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
                                'compruebo la nueva posicion
                                esta_vacio(orden_exploracion(8)) = f_esta_vacio_va0(1, nodo_y, nodo_x)
                                n_vec_visit(orden_exploracion(8)) = nodo_visitado_va0(1, nodo_y, nodo_x)
                                If esta_vacio(orden_exploracion(8)) And n_vec_visit(orden_exploracion(8)) = 0 Then
                                    'Ya lo he encontrado, me salgo
                                     nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                                Else
                                    '=============Fin Intentos=============
                                    'Si no se puede en ninguna de las direcciones,
                                    'escojo el menos visitado de los que no sean abstaculo
                                    'primero vuelvo a la direccion inicial
                                    nodo_y = viejo_y
                                    nodo_x = viejo_x
                                    direccion = viejo_direccion
                                    ultimo = 0
                                    If esta_vacio(orden_exploracion(1)) Then
                                        ultimo = ultimo + 1
                                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                                        lista_sin_obstaculos(ultimo) = orden_exploracion(1)
                                    End If
                                    If esta_vacio(orden_exploracion(2)) Then
                                        ultimo = ultimo + 1
                                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                                        lista_sin_obstaculos(ultimo) = orden_exploracion(2)
                                    End If
                                    If esta_vacio(orden_exploracion(3)) Then
                                        ultimo = ultimo + 1
                                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                                        lista_sin_obstaculos(ultimo) = orden_exploracion(3)
                                    End If
                                    If esta_vacio(orden_exploracion(4)) Then
                                        ultimo = ultimo + 1
                                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                                        lista_sin_obstaculos(ultimo) = orden_exploracion(4)
                                    End If
                                    If esta_vacio(orden_exploracion(5)) Then
                                        ultimo = ultimo + 1
                                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                                        lista_sin_obstaculos(ultimo) = orden_exploracion(5)
                                    End If
                                    If esta_vacio(orden_exploracion(6)) Then
                                        ultimo = ultimo + 1
                                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                                        lista_sin_obstaculos(ultimo) = orden_exploracion(6)
                                    End If
                                    If esta_vacio(orden_exploracion(7)) Then
                                        ultimo = ultimo + 1
                                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                                        lista_sin_obstaculos(ultimo) = orden_exploracion(7)
                                    End If
                                    If esta_vacio(orden_exploracion(8)) Then
                                        ultimo = ultimo + 1
                                        ReDim Preserve lista_sin_obstaculos(1 To ultimo) As Integer
                                        lista_sin_obstaculos(ultimo) = orden_exploracion(8)
                                    End If
                                    If ultimo = 0 Then
                                        'No hay ningun sitio vacio alrededor, estoy rodeado
                                        'Me salgo, pero aviso que no puedo moverme
                                         estoy_rodeado = True
                                         nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                                    Else
                                        'Hay alguno libre, aunque ya visitado
                                        'Cojo el menos visitado
                                        'Analizo cada uno
                                        For cont1 = 1 To ultimo
                                            'Lo comparo con el resto
                                            este_es_el_menor = True
                                            For cont2 = 1 To ultimo
                                                If n_vec_visit(lista_sin_obstaculos(cont1)) > n_vec_visit(lista_sin_obstaculos(cont2)) Then
                                                    'Este no es el menor
                                                    este_es_el_menor = False
                                                    Exit For
                                                End If
                                            Next cont2
                                            If este_es_el_menor Then
                                                'lista_sin_obstaculos(cont1) es el menor
                                                'Elijo esa direccion, es decir, lista_sin_obstaculos(cont1)
                                                direccion = f_giro_8_general_va0(direccion, lista_sin_obstaculos(cont1))
                                                'avanzo
                                                f_avanzo_8_dir_va0 nodo_z, nodo_y, nodo_x, direccion, max_pisos, max_filas, max_col
                                                'control errores de programacion
                                                If control_errores_de_programacion_ejv Then
                                                    If Not f_esta_vacio_va0(1, nodo_y, nodo_x) Then
                                                        s_error_ejv CON_OPCION_FINALIZAR, "Error: no está vacío"
                                                    End If
                                                End If
                                                nodo_visitado_va0(1, nodo_y, nodo_x) = nodo_visitado_va0(1, nodo_y, nodo_x) + 1
                                                Exit For
                                            End If
                                        Next cont1
                                        If Not este_es_el_menor Then
                                            'alguno tiene que ser menor o igual!
                                            s_error_ejv CON_OPCION_FINALIZAR, "Error: no es el menor"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    
    f_alg_general_8_direcciones_va0 = estoy_rodeado

End Function

Function f_avanzo_4_dir_va0(nodo_z As Double, nodo_y As Double, nodo_x As Double, direccion As Integer, max_pisos As Double, max_filas As Double, max_col As Double)
'El resultado de la funcion lo devuelvo en los mismos parametros nodo_y, nodo_x

    'CTE_ESTE = 1
    'CTE_NORTE = 2
    'CTE_OESTE = 3
    'CTE_SUR = 4

    Dim nueva_z As Double
    Dim nueva_y As Double
    Dim nueva_x As Double
    
    nueva_z = 1
    
    Select Case direccion
        Case CTE_NORTE
            nueva_y = nodo_y - 1
            nueva_x = nodo_x
        Case CTE_ESTE
            nueva_y = nodo_y
            nueva_x = nodo_x + 1
        Case CTE_SUR
            nueva_y = nodo_y + 1
            nueva_x = nodo_x
        Case CTE_OESTE
            nueva_y = nodo_y
            nueva_x = nodo_x - 1
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
    End Select

    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, nueva_z, nueva_y, nueva_x, max_pisos, max_filas, max_col

    'devuelvo los nuevos valores
    nodo_z = nueva_z
    nodo_y = nueva_y
    nodo_x = nueva_x
    
End Function

Function f_avanzo_8_dir_va0(nodo_z As Double, nodo_y As Double, nodo_x As Double, direccion As Integer, max_pisos As Double, max_filas As Double, max_col As Double)
'El resultado de la funcion lo devuelvo en los mismos parametros nodo_y, nodo_x

    'CTE_8_DEF = 1
    'CTE_8_DEF_DER = 2
    'CTE_8_DER = 3
    'CTE_8_ATR_DER = 4
    'CTE_8_ATR = 5
    'CTE_8_ATR_IZQ = 6
    'CTE_8_IZQ = 7
    'CTE_8_DEF_IZQ = 8

    Dim nueva_z As Double
    Dim nueva_y As Double
    Dim nueva_x As Double
    
    nueva_z = 1
    
    Select Case direccion
        Case CTE_8_DEF
            nueva_y = nodo_y - 1
            nueva_x = nodo_x
        Case CTE_8_DEF_DER
            nueva_y = nodo_y - 1
            nueva_x = nodo_x + 1
        Case CTE_8_DER
            nueva_y = nodo_y
            nueva_x = nodo_x + 1
        Case CTE_8_ATR_DER
            nueva_y = nodo_y + 1
            nueva_x = nodo_x + 1
        Case CTE_8_ATR
            nueva_y = nodo_y + 1
            nueva_x = nodo_x
        Case CTE_8_ATR_IZQ
            nueva_y = nodo_y + 1
            nueva_x = nodo_x - 1
        Case CTE_8_IZQ
            nueva_y = nodo_y
            nueva_x = nodo_x - 1
        Case CTE_8_DEF_IZQ
            nueva_y = nodo_y - 1
            nueva_x = nodo_x - 1
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
    End Select

    s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, nueva_z, nueva_y, nueva_x, max_pisos, max_filas, max_col
    
    'devuelvo los nuevos valores
    nodo_z = nueva_z
    nodo_y = nueva_y
    nodo_x = nueva_x
    
End Function
Function f_giro_4_general_va0(direccion_actual As Integer, direccion_giro As Integer)

    'CTE_ESTE = 1
    'CTE_NORTE = 2
    'CTE_OESTE = 3
    'CTE_SUR = 4
    f_giro_4_general_va0 = f_SumCirc(4, direccion_actual, direccion_giro - 1)

End Function

Function f_giro_8_general_va0(direccion_actual As Integer, direccion_giro As Integer) As Integer

    'CTE_8_DEF = 1
    'CTE_8_DEF_DER = 2
    'CTE_8_DER = 3
    'CTE_8_ATR_DER = 4
    'CTE_8_ATR = 5
    'CTE_8_ATR_IZQ = 6
    'CTE_8_IZQ = 7
    'CTE_8_DEF_IZQ = 8
    f_giro_8_general_va0 = f_SumCirc(8, direccion_actual, direccion_giro - 1)

End Function


