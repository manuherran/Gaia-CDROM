Attribute VB_Name = "bas_a4_pri"
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

'Presentacion de datos de tipos de jugadores
Global nombre_fichero_jugadores_pri As String 'nombre del fichero
Global fichero_tipos_jugadores_pri() As String 'contenido del fichero
Global tipos_jugadores_mostrar_pri() As String
Global mostrar_por_orden_de_pesos_pri As Boolean

Global habilitar_change_pri As Boolean

'Presentacion de datos a excel automatico
Global cont_mensajes_pri() As Long

'Opciones
Global todos_contra_todos_pri As Boolean
Global numero_de_reglas_del_agente_de_mas_reglas_pri As Integer
Global num_part_pri As Integer
Global grabar_resumen_pri As Boolean
Global probabilidad_de_error_pri As Integer
'Puntos ganados en el juego
Global ambos_cooperan_pri As Integer
Global ambos_defraudan_pri  As Integer
Global el_que_coopera_pri  As Integer
Global el_que_defrauda_pri  As Integer



'Control de opciones
Global fich_jug_modificado_pri As Boolean


Sub s_analisis_sintactico_tipos_jugadores_pri()

    Dim n_lin_tipos As Integer
    Dim n_lin_tipos_mostrar As Integer
    Dim estado As Integer

    Dim tipo_jugador_actual As Integer
    Dim regla_actual As Integer
    
    
    Dim s_tmp As String
    Dim i As Long
    
    tipo_jugador_actual = 0

    'Hago una primera pasada de los CTE_BEGIN_JUGADOR
    'para saber cuantos jugadores hay, redimensiono la primera dim
    'de los arrays de forma fija, y el numero de reglas por jugador
    'sera la segunda dimensión variable
    num_tipos_agentes_va0 = 0
    For n_lin_tipos = 1 To UBound(fichero_tipos_jugadores_pri)
        If InStr(fichero_tipos_jugadores_pri(n_lin_tipos), CTE_BEGIN_JUGADOR) <> 0 Then
            num_tipos_agentes_va0 = num_tipos_agentes_va0 + 1
        End If
    Next n_lin_tipos
    
    'Redimensiono todo
    ReDim nombre_tipo_jugador_pri(1 To num_tipos_agentes_va0) As String
    ReDim num_agentes_tipo_va0(1 To num_tipos_agentes_va0) As Integer
    ReDim tendencia_rel_inicial_mov_tipo_agente_va0(1 To CTE_8_DIR, 1 To num_tipos_agentes_va0) As Long
    ReDim tendencia_abs_inicial_mov_tipo_agente_va0(1 To CTE_8_DIR, 1 To num_tipos_agentes_va0) As Long
    ReDim prioridad_regla_tipo_jugador_pri(1 To num_tipos_agentes_va0, 1 To 1) As Integer
    ReDim condicion_regla_tipo_jugador_pri(1 To num_tipos_agentes_va0, 1 To 1) As String
    ReDim accion_regla_tipo_jugador_pri(1 To num_tipos_agentes_va0, 1 To 1) As String
    If Not esta_modificado_num_agen_tipo_pri Then
        ReDim numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1 To num_tipos_agentes_va0) As Integer
    Else
        ReDim Preserve numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1 To num_tipos_agentes_va0) As Integer
    End If
    If grabar_resumen_pri Then
        ReDim resumen_todo(1 To num_tipos_agentes_va0, 1 To 1) As Long 'Tambien el resumen
    End If



    estado = CTE_INICIO
    n_lin_tipos_mostrar = 0
    numero_de_reglas_del_agente_de_mas_reglas_pri = 1
    For n_lin_tipos = 1 To UBound(fichero_tipos_jugadores_pri)
    
        'comentario
        If Left(fichero_tipos_jugadores_pri(n_lin_tipos), 1) = "'" Then
            'No hacemos nada, la ignoramos
    
        'BEGIN JUGADOR
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_BEGIN_JUGADOR)) = CTE_BEGIN_JUGADOR Then
            If estado <> CTE_NINGUNO Then End
            'Añado una linea vacia previa
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = ""
            'Añado la linea
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = fichero_tipos_jugadores_pri(n_lin_tipos)
            estado = CTE_B_JUGADOR_LEIDO
            tipo_jugador_actual = tipo_jugador_actual + 1
            regla_actual = 0
            'Inicializo con un valor erroneo
            'para detectar que faltan los parametros opcionales
            If Not esta_modificado_num_agen_tipo_pri Then
                numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(tipo_jugador_actual) = -1
            End If
            tendencia_rel_inicial_mov_tipo_agente_va0(1, tipo_jugador_actual) = -1
            
        'NOMBRE JUGADOR
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_NOMBRE_JUGADOR)) = CTE_NOMBRE_JUGADOR Then
            If estado <> CTE_B_JUGADOR_LEIDO Then End
            'Añado la linea con 2 espacios
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "  " & fichero_tipos_jugadores_pri(n_lin_tipos)
            estado = CTE_N_JUGADOR_LEIDO
            'Añado
            ReDim Preserve nombre_tipo_jugador_pri(1 To tipo_jugador_actual) As String
            nombre_tipo_jugador_pri(tipo_jugador_actual) = Trim(Right(fichero_tipos_jugadores_pri(n_lin_tipos), Len(fichero_tipos_jugadores_pri(n_lin_tipos)) - Len(CTE_NOMBRE_JUGADOR) - 1))
            
            
        'NUMERO: numero de agentes de ese tipo
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_NUMERO_AGENTES)) = CTE_NUMERO_AGENTES Then
            'Añado la linea con 2 espacios
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            If Not esta_modificado_num_agen_tipo_pri Then
                tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "  " & fichero_tipos_jugadores_pri(n_lin_tipos)
                'Añado
                numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(tipo_jugador_actual) = CInt(Right(fichero_tipos_jugadores_pri(n_lin_tipos), Len(fichero_tipos_jugadores_pri(n_lin_tipos)) - Len(CTE_NUMERO_AGENTES) - 1))
            Else
                tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "  " & numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(tipo_jugador_actual)
            End If
            
        'PARAMETROS MOVIMIENTO: forma en la que se mueve este tipo de agente
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_PARAMETROS_MOVIMIENTO)) = CTE_PARAMETROS_MOVIMIENTO Then
            'Añado la linea con 2 espacios
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "  " & fichero_tipos_jugadores_pri(n_lin_tipos)
            'Añado
            s_tmp = Right(fichero_tipos_jugadores_pri(n_lin_tipos), Len(fichero_tipos_jugadores_pri(n_lin_tipos)) - Len(CTE_PARAMETROS_MOVIMIENTO) - 1)
            For i = 1 To CTE_8_DIR
                tendencia_rel_inicial_mov_tipo_agente_va0(i, tipo_jugador_actual) = f_elemento_listacomas(s_tmp, i)
            Next i
            
        'BEGIN REGLA
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_BEGIN_REGLA)) = CTE_BEGIN_REGLA Then
            
            'Antes de añadir lo propio de la regla
            'Añado los parametros opcionales:
            'el numero de agentes si no ha salido ya
            'y los parametros del movimiento
            'Por defecto fijo 1 jugador de ese tipo si no ha salido ya
            If numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(tipo_jugador_actual) = -1 Then
                'Pongo el dato por defecto
                numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(tipo_jugador_actual) = 1
                'Añado la linea que no existia con 2 espacios
                n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
                ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
                tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "  " & CTE_NUMERO_AGENTES & ":1"
            End If
            
            'Por defecto fijo un movimiento normalillo
            If tendencia_rel_inicial_mov_tipo_agente_va0(1, tipo_jugador_actual) = -1 Then
                'Pongo el dato por defecto
                tendencia_rel_inicial_mov_tipo_agente_va0(1, tipo_jugador_actual) = 40
                tendencia_rel_inicial_mov_tipo_agente_va0(2, tipo_jugador_actual) = 20
                tendencia_rel_inicial_mov_tipo_agente_va0(2, tipo_jugador_actual) = 10
                tendencia_rel_inicial_mov_tipo_agente_va0(4, tipo_jugador_actual) = 5
                tendencia_rel_inicial_mov_tipo_agente_va0(5, tipo_jugador_actual) = 1
                tendencia_rel_inicial_mov_tipo_agente_va0(6, tipo_jugador_actual) = 5
                tendencia_rel_inicial_mov_tipo_agente_va0(7, tipo_jugador_actual) = 10
                tendencia_rel_inicial_mov_tipo_agente_va0(8, tipo_jugador_actual) = 20
                'Añado la linea que no existia con 2 espacios
                n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
                ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
                tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "  " & CTE_PARAMETROS_MOVIMIENTO & ":40,20,10,5,1,5,10,20"
            End If
            
            'Añado la linea con 2 espacios
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "  " & fichero_tipos_jugadores_pri(n_lin_tipos)
            regla_actual = regla_actual + 1
            
            If regla_actual > numero_de_reglas_del_agente_de_mas_reglas_pri Then
                numero_de_reglas_del_agente_de_mas_reglas_pri = regla_actual
            End If
            
            'inicializo
            ReDim Preserve prioridad_regla_tipo_jugador_pri(1 To num_tipos_agentes_va0, 1 To numero_de_reglas_del_agente_de_mas_reglas_pri) As Integer
            ReDim Preserve condicion_regla_tipo_jugador_pri(1 To num_tipos_agentes_va0, 1 To numero_de_reglas_del_agente_de_mas_reglas_pri) As String
            ReDim Preserve accion_regla_tipo_jugador_pri(1 To num_tipos_agentes_va0, 1 To numero_de_reglas_del_agente_de_mas_reglas_pri) As String
            prioridad_regla_tipo_jugador_pri(tipo_jugador_actual, regla_actual) = 0
            condicion_regla_tipo_jugador_pri(tipo_jugador_actual, regla_actual) = ""
            accion_regla_tipo_jugador_pri(tipo_jugador_actual, regla_actual) = "kk"
            If accion_regla_tipo_jugador_pri(1, 1) = "" Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error: regla vacia"
            End If
        
        'PRIORIDAD
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_PRIORIDAD)) = CTE_PRIORIDAD Then
            'Añado la linea con 4 espacios
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "    " & fichero_tipos_jugadores_pri(n_lin_tipos)
            'Añado
            prioridad_regla_tipo_jugador_pri(tipo_jugador_actual, regla_actual) = CInt(Trim(Right(fichero_tipos_jugadores_pri(n_lin_tipos), Len(fichero_tipos_jugadores_pri(n_lin_tipos)) - Len(CTE_PRIORIDAD) - 1)))
            
        
        'CONDICION
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_CONDICION)) = CTE_CONDICION Then
            'Añado la linea con 4 espacios
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "    " & fichero_tipos_jugadores_pri(n_lin_tipos)
            'Añado
            condicion_regla_tipo_jugador_pri(tipo_jugador_actual, regla_actual) = Trim(Right(fichero_tipos_jugadores_pri(n_lin_tipos), Len(fichero_tipos_jugadores_pri(n_lin_tipos)) - Len(CTE_CONDICION) - 1))
        
        'ACCION
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_ACCION)) = CTE_ACCION Then
            'Añado la linea con 4 espacios
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "    " & fichero_tipos_jugadores_pri(n_lin_tipos)
            'Añado
            accion_regla_tipo_jugador_pri(tipo_jugador_actual, regla_actual) = Trim(Right(fichero_tipos_jugadores_pri(n_lin_tipos), Len(fichero_tipos_jugadores_pri(n_lin_tipos)) - Len(CTE_ACCION) - 1))
            If accion_regla_tipo_jugador_pri(tipo_jugador_actual, regla_actual) = "" Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error: acción regla vacia"
            End If
        
        'END REGLA
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_END_REGLA)) = CTE_END_REGLA Then
            'Añado la linea con 2 espacios
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = "  " & fichero_tipos_jugadores_pri(n_lin_tipos)
        
        'END JUGADOR
        ElseIf Left(fichero_tipos_jugadores_pri(n_lin_tipos), Len(CTE_END_JUGADOR)) = CTE_END_JUGADOR Then
            'Añado la linea
            n_lin_tipos_mostrar = n_lin_tipos_mostrar + 1
            ReDim Preserve tipos_jugadores_mostrar_pri(1 To n_lin_tipos_mostrar) As String
            tipos_jugadores_mostrar_pri(n_lin_tipos_mostrar) = fichero_tipos_jugadores_pri(n_lin_tipos)
            estado = CTE_INICIO
        Else
            'Error
            s_error_ejv CON_OPCION_FINALIZAR, "Error: El fichero de tipos de jugadores del prisionero es incorrecto. La linea " & n_lin_tipos & " es erronea: " & fichero_tipos_jugadores_pri(n_lin_tipos)
        End If
    Next n_lin_tipos


End Sub

Sub s_mostrar_fichero_tipos_jugadores_pri()

    Dim n_lin As Integer
    Dim texto As String
    
    texto = ""
    
    Screen.MousePointer = CTE_ARENA
    For n_lin = 1 To UBound(tipos_jugadores_mostrar_pri)
        texto = texto & tipos_jugadores_mostrar_pri(n_lin) & vbCrLf
    Next n_lin
    habilitar_change_pri = False
    frm_a4_tipospri.txt_tipos.Text = texto
    habilitar_change_pri = True
    
    'frm_a4_tipospri.txt_tipos.Refresh
    Screen.MousePointer = CTE_DEFECTO

End Sub
Sub s_inicializar_ejemplo_elegido_pri()

    Dim exito_al_abrir As Boolean
    Dim p As Integer
    Dim f As Integer
    Dim c As Integer
    
    'Los tipos de agentes se cargan al leer el fichero de tipos de agentes
    
    'OPCIONES I
    'GENERALES DE EJEMPLOS DE VIDA(DISTINTAS A LAS DE POR DEFECTO)
    '2 Grabar Resumen
    un_ej_grabar_gra_ejv = True
    un_ej_fichero_gra_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_GRA), "r_pri_" & num_ej_activo_ejv & ".gra")
    un_ej_grabar_resumen_txt_ejv = False
    un_ej_fichero_resumen_txt_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_TXT), "r_pri_" & num_ej_activo_ejv & ".txt")
    un_ej_grabar_resumen_xls_ejv = False
    un_ej_fichero_resumen_xls_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS), "r_pri_" & num_ej_activo_ejv & ".xls")
    max_guardado_ejv = 1000000
    autoguardado_ejv = 100

    'OPCIONES II
    'GENERALES DE VIDA ARTIFICIAL
    '1 Modo de Ejecución
    ver_agentes_va0 = True
    '2 agentes inmortales
    agentes_inmortales_va0 = True
    muerte1_va0 = 35
    muerte2_va0 = 40
    '3 tasas de mutación
    probb_mutacion_tipo_inicial_va0 = 0
    probb_mutacion_mov_inicial_va0 = 0
    probb_mutacion_pm_inicial_va0 = 0
    PMPMCte_va0 = True
    '4 Lugar de nacimiento
    nacimiento_cerca_va0 = True
    '5 Búsqueda de Cadena binaria
    busqueda_cadena_binaria_va0 = False
    cadena_binaria_buscada_va0 = "000000000100000000010000000001"
    long_cadena_buscada_va0 = Len(cadena_binaria_buscada_va0)
    '6 Limite Muerte
    limite_muerte_va0 = -1

    
    
    
    Select Case num_ej_activo_ejv
        Case 1
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "pris1.map"
        'ESPECIFICAS DE EL DILEMA DEL PRISIONERO
        '2 Fichero de definición de jugadores
        nombre_fichero_jugadores_pri = "default.pri"
        '3 numero de partidas al prisionero que juegan 2 agentes cada vez que se encuentran
        num_part_pri = 10
        '4 energia_consumida_al_mover_va0 una posicion
        energia_consumida_al_mover_va0 = 0
        '6 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '7 Tipo de juego
        todos_contra_todos_pri = False 'Op_TodosContraTodos
        '9 energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 20
        '12 energia inicial de cada agente
        energia_inicial_agente_va0 = 0
        '16 Probabilidad de Error en la decision
        probabilidad_de_error_pri = 0
        'Puntos ganados en el juego
        ambos_cooperan_pri = 3
        ambos_defraudan_pri = 0
        el_que_coopera_pri = 0
        el_que_defrauda_pri = 5
        

        Case 2
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "pris2.map"
        'ESPECIFICAS DE EL DILEMA DEL PRISIONERO
        '2 Fichero de definición de jugadores
        nombre_fichero_jugadores_pri = "todos10.pri"
        '3 numero de partidas al prisionero que juegan 2 agentes cada vez que se encuentran
        num_part_pri = 10
        '4 energia_consumida_al_mover_va0 una posicion
        energia_consumida_al_mover_va0 = 0
        '6 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '7 Tipo de juego
        todos_contra_todos_pri = False
        '9 energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 20
        '12 energia inicial de cada agente
        energia_inicial_agente_va0 = 0
        '15 resumen
        grabar_resumen_pri = False
        '16 Probabilidad de Error en la decision
        probabilidad_de_error_pri = 0
        'Puntos ganados en el juego
        ambos_cooperan_pri = 3
        ambos_defraudan_pri = 0
        el_que_coopera_pri = 0
        el_que_defrauda_pri = 5


        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe ese ejemplo"
    End Select

    'Cargo el mapa
    mapa_actual_ma0 = f_nombre_completo_existente(path_largo_ejv(CTE_C_PRG_MAP), nombre_fichero_mapa_va0)
    s_aut_leer_mapa_ma0
    s_copiar_mapa_ma0_sobre_va0_va0
    
    
    'Elijo fichero de jugadores
    'nombre_fichero_ejv = f_nombre_completo(path_largo_ejv(CTE_C_PRG_PRI), nombre_fichero_jugadores_pri)
    exito_al_abrir = s_aut_abrir_jugadores_pri(f_nombre_completo(path_largo_ejv(CTE_C_PRG_PRI), nombre_fichero_jugadores_pri))

    If exito_al_abrir Then
        s_analisis_sintactico_tipos_jugadores_pri
    End If


End Sub
Sub s_inicializar_pri()
    
    Dim cont_tipo As Integer
    Dim i As Integer
    
    hay_que_detener_ejv = False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
    
    frm_a0_va.Refresh
    
    se_ha_empezado_a_crear_agentes_va0 = False
    'Creo la lista de agentes reales, en la que
    'existen N de cada tipo
    
    'Calculo el total de jugadores que voy a tener que crear
    numero_total_de_agentes_ejv = 0
    For cont_tipo = 1 To num_tipos_agentes_va0
        numero_total_de_agentes_ejv = numero_total_de_agentes_ejv + numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(cont_tipo)
    Next cont_tipo
    
    s_inicializar_arrays_va0
    
    For i = 1 To numero_total_de_agentes_ejv
        cont_mensajes_pri(i) = 0
    Next i

    s_mapa_pintar_bordes_va0 frm_a0_va
    s_mostrar_mapa_actual_va0 False
    
    se_ha_empezado_a_crear_agentes_va0 = True
    s_crear_agentes_iniciales_va0
    
    If un_ej_grabar_resumen_xls_ejv Then
        For i = 1 To numero_total_de_agentes_ejv
            cont_mensajes_pri(i) = cont_mensajes_pri(i) + 1
            s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, i, 3 + cont_mensajes_pri(i), CLng(i)
            
            cont_mensajes_pri(i) = cont_mensajes_pri(i) + 1
            s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, nombre_tipo_jugador_pri(agente_tipo_va0(i)), 3 + cont_mensajes_pri(i), CLng(i)
        Next i
    End If
    If un_ej_grabar_resumen_txt_ejv Then
        For i = 1 To numero_total_de_agentes_ejv
        Next i
    End If
    
    


End Sub

Sub s_cargar_opciones_pri()

    'Check:0,1
    'Option:true,false


    '1 Fichero de Mapa
    'nombre_fichero_mapa_va0 = "prisionero.map"
    '2 Fichero de definición de jugadores
    'nombre_fichero_jugadores_pri = "default.pri"
    '3 numero de partidas al prisionero que juegan 2 agentes cada vez que se encuentran
    frm_a4_oppri.Op_num_part = num_part_pri
    '4 energia_consumida_al_mover_va0 una posicion
    frm_a4_oppri.Op_EnergiaConsumUnaPos = energia_consumida_al_mover_va0
    '6 Algoritmo de busqueda de un espacio libre cercano
    'algoritmo_busqueda_va0 = 3
    '7 Tipo de juego
    frm_a4_oppri.Op_TodosContraTodos = todos_contra_todos_pri
    frm_a4_oppri.Op_nTodosContraTodos = Not todos_contra_todos_pri
    '9 energia_consumida_al_reproducirse
    frm_a4_oppri.Op_energiaConsumidaReproducirse = energia_consumida_al_reproducirse_va0
    '12 energia inicial de cada agente
    frm_a4_oppri.Op_EnergiaInicialAgente = energia_inicial_agente_va0
    '16 Probabilidad de Error en la decision
    frm_a4_oppri.Op_ProbbError.Text = probabilidad_de_error_pri
    'Puntos ganados en el juego
    frm_a4_oppri.ambos_cooperan = ambos_cooperan_pri
    frm_a4_oppri.ambos_defraudan = ambos_defraudan_pri
    frm_a4_oppri.el_que_coopera = el_que_coopera_pri
    frm_a4_oppri.el_que_defrauda = el_que_defrauda_pri



End Sub
Sub s_cargar_tipos_agentes_pri()

    Dim i As Integer
    Dim s_i As String
    Dim s_nombre As String
    Dim s_num As String
        
    frm_a4_oppri.Cb_TipoAgente.Clear
    frm_a4_oppri.tipos.Clear
    For i = 1 To num_tipos_agentes_va0
        s_i = CStr(i)
        s_nombre = Trim(nombre_tipo_jugador_pri(i))
        s_num = CStr(numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(i))
        While Len(s_i) < 4
            s_i = " " & s_i
        Wend
        While Len(s_nombre) < 40
            s_nombre = s_nombre & " "
        Wend
        If Len(s_nombre) > 40 Then
            s_nombre = Left(s_nombre, 40)
        End If
        While Len(s_num) < 4
            s_num = " " & s_num
        Wend
        frm_a4_oppri.Cb_TipoAgente.AddItem s_i & ": " & nombre_tipo_jugador_pri(i)
        frm_a4_oppri.tipos.AddItem s_i & ": " & s_nombre & " :" & s_num
    Next i
    frm_a4_oppri.Cb_TipoAgente.ListIndex = 0

End Sub


Sub s_accion_jugar_pri(ag As Integer, mis_vecinos() As Integer)
    
    'hormiga actual
    Dim X As Integer
    Dim Y As Integer
    
    Dim cont As Integer
    Dim vecino_elegido As Integer
    Dim vecinos_desordenados(1 To CTE_8_DIR) As Integer

    Dim he_jugado As Boolean

    'Calculo el vecino con el que jueaga
    'El agente siempre otro dispuesto a jugar
    
    he_jugado = False
    'Array ordenado
    For cont = 1 To CTE_8_DIR
        vecinos_desordenados(cont) = cont
    Next cont

    'hormiga actual
    X = agente_x_va0(ag)
    Y = agente_y_va0(ag)

    f_desordenar_array_i vecinos_desordenados()
    
    'Ahora el array está desordenado
    For cont = 1 To CTE_8_DIR
        vecino_elegido = vecinos_desordenados(cont)
        If mis_vecinos(vecino_elegido) = CTE_VEC_AGENTE Then
            'Es "agente"
            s_jugar_agente_pri ag, vecino_elegido
            he_jugado = True
            Exit For
        End If
    Next cont

    If Not he_jugado Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error: no he jugado"
    End If



End Sub

Sub s_jugar_agente_pri(ByVal age As Integer, direccion As Integer)

    
    Dim contrario_z As Double
    Dim contrario_y As Double
    Dim contrario_x As Double


    contrario_z = 1
    Dim age_contrario As Integer

    Select Case direccion
        Case CTE_8_N
            contrario_x = agente_x_va0(age)
            contrario_y = agente_y_va0(age) - 1
        Case CTE_8_NE
            contrario_x = agente_x_va0(age) + 1
            contrario_y = agente_y_va0(age) - 1
        Case CTE_8_E
            contrario_x = agente_x_va0(age) + 1
            contrario_y = agente_y_va0(age)
        Case CTE_8_SE
            contrario_x = agente_x_va0(age) + 1
            contrario_y = agente_y_va0(age) + 1
        Case CTE_8_S
            contrario_x = agente_x_va0(age)
            contrario_y = agente_y_va0(age) + 1
        Case CTE_8_SO
            contrario_x = agente_x_va0(age) - 1
            contrario_y = agente_y_va0(age) + 1
        Case CTE_8_O
            contrario_x = agente_x_va0(age) - 1
            contrario_y = agente_y_va0(age)
        Case CTE_8_NO
            contrario_x = agente_x_va0(age) - 1
            contrario_y = agente_y_va0(age) - 1
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
    End Select
    If contrario_y = 0 Then contrario_y = mapa_filas_va0
    If contrario_y = mapa_filas_va0 + 1 Then contrario_y = 1
    
    If contrario_x = 0 Then contrario_x = mapa_columnas_va0
    If contrario_x = mapa_columnas_va0 + 1 Then contrario_x = 1
    
    age_contrario = fi_indice_agente_va0(contrario_z, contrario_y, contrario_x)
    
    s_jugar_partidas age, age_contrario
    

End Sub
Sub s_jugar_partidas(ByVal age As Integer, ByVal age_contrario As Integer)

    Dim i As Integer
    
    Dim decision_age As Integer
    Dim decision_enemigo As Integer

    Dim suma_age As Integer
    Dim suma_contrario As Integer


    'borro las historias
    ReDim historia_agente_yo_pri(1 To numero_total_de_agentes_ejv, 1 To 1) As Integer
    ReDim historia_agente_el_pri(1 To numero_total_de_agentes_ejv, 1 To 1) As Integer
    
    'Juego n partidas: num_part_pri
    For i = 1 To num_part_pri

        PA_agente_pri(age) = i
        PA_agente_pri(age_contrario) = i
        
        'Leo la acccion de cada uno
        decision_age = f_tomar_decision_pri(age, age_contrario)
        decision_enemigo = f_tomar_decision_pri(age_contrario, age)
        
        'Control de errores
        If control_errores_de_programacion_ejv Then
            If i > 1 And decision_age = CTE_C And nombre_tipo_jugador_pri(agente_tipo_va0(age)) = "DONDE_LAS_DAN_LAS_TOMANJAT" And nombre_tipo_jugador_pri(agente_tipo_va0(age_contrario)) = "DON_CABRON" Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End If
            If i > 1 And decision_enemigo = CTE_C And nombre_tipo_jugador_pri(agente_tipo_va0(age_contrario)) = "DONDE_LAS_DAN_LAS_TOMANJAT" And nombre_tipo_jugador_pri(agente_tipo_va0(age)) = "DON_CABRON" Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End If
        End If
        
        
        ReDim Preserve historia_agente_yo_pri(1 To numero_total_de_agentes_ejv, 1 To i) As Integer
        ReDim Preserve historia_agente_el_pri(1 To numero_total_de_agentes_ejv, 1 To i) As Integer
        historia_agente_yo_pri(age, i) = decision_age
        historia_agente_el_pri(age, i) = decision_enemigo
        historia_agente_yo_pri(age_contrario, i) = decision_enemigo
        historia_agente_el_pri(age_contrario, i) = decision_age
        
        'Modifico pesos
        Select Case decision_age
            Case CTE_C
                Select Case decision_enemigo
                    Case CTE_C
                        'Ambos Cooperan
                        suma_age = ambos_cooperan_pri
                        suma_contrario = ambos_cooperan_pri
                    Case CTE_D
                        'Age coopera y enemigo defrauda
                        suma_age = el_que_coopera_pri
                        suma_contrario = el_que_defrauda_pri
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa acción"
                    End Select
            Case CTE_D
                Select Case decision_enemigo
                    Case CTE_C
                        'Age defrauda y enemigo coopera
                        suma_age = el_que_defrauda_pri
                        suma_contrario = el_que_coopera_pri
                    Case CTE_D
                        'Ambos Defraudan
                        suma_age = ambos_defraudan_pri
                        suma_contrario = ambos_defraudan_pri
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa acción"
                    End Select
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa acción"
        End Select
        
        peso_agente_va0(age) = peso_agente_va0(age) + suma_age
        peso_agente_va0(age_contrario) = peso_agente_va0(age_contrario) + suma_contrario
    
        If un_ej_grabar_resumen_xls_ejv Then
            'Agente
            cont_mensajes_pri(age) = cont_mensajes_pri(age) + 1
            If decision_age = CTE_C Then
                s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, "Cooperar jugando contra " & nombre_tipo_jugador_pri(agente_tipo_va0(age_contrario)), 3 + cont_mensajes_pri(age), CLng(age)
            Else
                s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, "Defraudar jugando contra " & nombre_tipo_jugador_pri(agente_tipo_va0(age_contrario)), 3 + cont_mensajes_pri(age), CLng(age)
            End If
            cont_mensajes_pri(age) = cont_mensajes_pri(age) + 1
            s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, peso_agente_va0(age), 3 + cont_mensajes_pri(age), CLng(age)
            'Contrario
            cont_mensajes_pri(age_contrario) = cont_mensajes_pri(age_contrario) + 1
            If decision_enemigo = CTE_C Then
                s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, "Cooperar jugando contra " & nombre_tipo_jugador_pri(agente_tipo_va0(age)), 3 + cont_mensajes_pri(age_contrario), CLng(age_contrario)
            Else
                s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, "Defraudar jugando contra " & nombre_tipo_jugador_pri(agente_tipo_va0(age)), 3 + cont_mensajes_pri(age_contrario), CLng(age_contrario)
            End If
            cont_mensajes_pri(age_contrario) = cont_mensajes_pri(age_contrario) + 1
            s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, peso_agente_va0(age_contrario), 3 + cont_mensajes_pri(age_contrario), CLng(age_contrario)
        End If
        If un_ej_grabar_resumen_txt_ejv Then
        End If
        
    Next i

End Sub

Function f_tomar_decision_pri(ByVal age As Integer, ByVal age_contrario As Integer) As Integer

    Dim cont_reglas_este_agente As Integer
    Dim cont_reglas_disparan As Integer
    Dim ultima_regla As Integer
    Dim regla_a_analizar As Integer
    Dim regla_elegida  As Integer
    
    Dim prioridad_maxima As Integer
    Dim condicion As String
    
    Dim lista_reglas() As Integer
    
    Dim tipo_jugador As Integer
    tipo_jugador = agente_tipo_va0(age)
    
    'Por cada regla veo si se cumple
    'Hago una lista de las reglas que se disparan
    'Todas las reglas tienen al menos condicion
    cont_reglas_disparan = 0
    cont_reglas_este_agente = 1
    condicion = condicion_regla_tipo_jugador_pri(tipo_jugador, 1)
    While condicion <> ""
        If f_evaluar_condicion_pri(age, age_contrario, condicion) Then
            cont_reglas_disparan = cont_reglas_disparan + 1
            ReDim Preserve lista_reglas(1 To cont_reglas_disparan) As Integer
            lista_reglas(cont_reglas_disparan) = cont_reglas_este_agente
        End If
        'voy a por el siguiente a ver si hay algo
        If cont_reglas_este_agente < numero_de_reglas_del_agente_de_mas_reglas_pri Then
            cont_reglas_este_agente = cont_reglas_este_agente + 1
            condicion = condicion_regla_tipo_jugador_pri(tipo_jugador, cont_reglas_este_agente)
        Else
            condicion = ""
        End If
    Wend
    ultima_regla = cont_reglas_disparan
    
    'si no se dispara ninguna, juego al azar ambos con igual posibilidad
    If ultima_regla = 0 Then
        If fi_azar1(2) = 1 Then
            f_tomar_decision_pri = CTE_C
        Else
            f_tomar_decision_pri = CTE_D
        End If
    Else
        'Primero busco la prioridad maxima
        prioridad_maxima = -1
        For cont_reglas_disparan = 1 To ultima_regla
            regla_a_analizar = lista_reglas(cont_reglas_disparan)
            If prioridad_regla_tipo_jugador_pri(tipo_jugador, regla_a_analizar) > prioridad_maxima Then
                prioridad_maxima = CInt(prioridad_regla_tipo_jugador_pri(tipo_jugador, lista_reglas(cont_reglas_disparan)))
            End If
        Next cont_reglas_disparan
        
        'Ahora Borro las que no tengan prioridad maxima
        cont_reglas_disparan = 0
        While cont_reglas_disparan < ultima_regla
            cont_reglas_disparan = cont_reglas_disparan + 1
            regla_a_analizar = lista_reglas(cont_reglas_disparan)
            If prioridad_regla_tipo_jugador_pri(tipo_jugador, regla_a_analizar) < prioridad_maxima Then
                If cont_reglas_disparan <> ultima_regla Then
                    'lo sustituyo por el ultimo
                    lista_reglas(cont_reglas_disparan) = lista_reglas(ultima_regla)
                End If
                'En cualquier caso borro ese elemento ultimo
                ultima_regla = ultima_regla - 1
            End If
        Wend
        
        'Si hay mas de uno con la maxima prioridad, elijo uno al azar
        If ultima_regla > 1 Then
            regla_elegida = lista_reglas(fi_azar1(ultima_regla))
        Else
            regla_elegida = lista_reglas(1)
        End If
        
        'Ejecuto la accion
        f_tomar_decision_pri = f_ejecutar_accion_regla_pri(accion_regla_tipo_jugador_pri(tipo_jugador, regla_elegida))
    
        'Cuando ya tengo decicida la accion, la invierto con cierta probabilidad
        'para simular posibles errores (factor de error, sugerencia de Juan Antonio Tubio)
        If probabilidad_de_error_pri > 0 Then
            If fi_azar1(100) <= probabilidad_de_error_pri Then
                If f_tomar_decision_pri = CTE_D Then
                    f_tomar_decision_pri = CTE_C
                Else
                    f_tomar_decision_pri = CTE_D
                End If
            End If
        End If

    End If

End Function
Function f_evaluar_condicion_pri(ByVal age As Integer, ByVal age_contrario As Integer, condicion As String) As Boolean

    Dim arr_con() As String
    Dim cont_cond As Integer
    Dim numero_cond As Integer
    Dim se_cumple As Boolean
    Dim pto As Integer
    
    Dim me_quedo As Integer
    
    Dim tipo_jugador As Integer
    tipo_jugador = agente_tipo_va0(age)
    
    se_cumple = True
    
    'Separo en trozos AND
    ReDim arr_con(1 To 1) As String
    numero_cond = 1
    pto = InStr(condicion, " AND ")
    If pto > 0 Then
        While pto > 0
            'pongo en el array
            arr_con(numero_cond) = Left(condicion, pto - 1)
            'lo quito de la variable
            me_quedo = Len(condicion) - pto + 1 - Len(" AND ")
            condicion = Right(condicion, me_quedo)
            'pongo el resto de la condicion
            numero_cond = numero_cond + 1
            ReDim Preserve arr_con(1 To numero_cond) As String
            arr_con(numero_cond) = condicion
            'veo si voy a trocear mas
            pto = InStr(arr_con(numero_cond), " AND ")
        Wend
    Else
        'No hay AND
        numero_cond = 1
        ReDim arr_con(1 To 1) As String
        arr_con(1) = condicion
    End If

    cont_cond = 0
    While se_cumple And cont_cond < numero_cond
        cont_cond = cont_cond + 1
        arr_con(cont_cond) = Trim(arr_con(cont_cond))
        If arr_con(cont_cond) = "SIEMPRE" Then
            'condicion: SIEMPRE
            se_cumple = True
        ElseIf InStr(arr_con(cont_cond), "%") > 0 Then
            'condicion: %
            se_cumple = f_trat_tanto_por_ciento_pri(arr_con(cont_cond))
        ElseIf Left(arr_con(cont_cond), Len("EL")) = "EL" Then
            'condicion: EL
            se_cumple = f_trat_EL_YO_pri("EL", age, age_contrario, arr_con(cont_cond))
        ElseIf Left(arr_con(cont_cond), Len("YO")) = "YO" Then
            'condicion: YO
            se_cumple = f_trat_EL_YO_pri("YO", age, age_contrario, arr_con(cont_cond))
        ElseIf Left(arr_con(cont_cond), Len("NP")) = "NP" Then
            'condicion: NP
            se_cumple = f_trat_NP_pri(age, arr_con(cont_cond))
        Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa condición"
        End If
    Wend
    
    f_evaluar_condicion_pri = se_cumple
    
End Function

Function f_trat_EL_YO_pri(EL_YO As String, ByVal age As Integer, ByVal age_contrario As Integer, subcond As String) As Boolean

    Dim i_jugada As Integer
    Dim s_partida As String
    Dim l_partida As Long
    Dim l_partida_pasado As Long
    Dim se_cumple As Boolean
    
    Dim tipo_jugador As Integer
    tipo_jugador = agente_tipo_va0(age)
    
    se_cumple = True

    'condicion: EL o YO
    Select Case EL_YO
        Case "EL"
            subcond = f_quitar_segmento_pri(subcond, "EL")
        Case "YO"
            subcond = f_quitar_segmento_pri(subcond, "YO")
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existen mas opciones"
    End Select
    
    subcond = f_quitar_segmento_pri(subcond, "=")
    
    i_jugada = f_quitar_trozo_CD_pri(subcond)
    
    subcond = f_quitar_segmento_pri(subcond, "EN")
    subcond = f_quitar_segmento_pri(subcond, "NP")
    subcond = f_quitar_segmento_pri(subcond, "=")
    
    s_partida = subcond
    If IsNumeric(s_partida) Then
        l_partida = CLng(s_partida)
        If l_partida < 1 Then
            se_cumple = False
        Else
            Select Case EL_YO
                Case "EL"
                    If historia_agente_el_pri(age, l_partida) = i_jugada Then
                        se_cumple = True
                    Else
                        se_cumple = False
                    End If
                Case "YO"
                    If historia_agente_yo_pri(age, l_partida) = i_jugada Then
                        se_cumple = True
                    Else
                        se_cumple = False
                    End If
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: no existen mas opciones"
            End Select
        End If
    Else
        subcond = f_quitar_segmento_pri(subcond, "PA")
        subcond = f_quitar_segmento_pri(subcond, "-")
        s_partida = subcond
        If IsNumeric(s_partida) Then
            l_partida = CLng(s_partida)
            l_partida_pasado = PA_agente_pri(age) - l_partida
            If l_partida_pasado < 1 Then
                se_cumple = False
            Else
                Select Case EL_YO
                    Case "EL"
                        If historia_agente_el_pri(age, l_partida_pasado) = i_jugada Then
                            se_cumple = True
                        Else
                            se_cumple = False
                        End If
                    Case "YO"
                        If historia_agente_yo_pri(age, l_partida_pasado) = i_jugada Then
                            se_cumple = True
                        Else
                            se_cumple = False
                        End If
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: no existen mas opciones"
                End Select
            End If
        Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: partida no numerica"
        End If
    End If

    f_trat_EL_YO_pri = se_cumple

End Function
Function f_trat_NP_pri(ByVal age As Integer, subcond As String) As Boolean

    Dim se_cumple As Boolean
    Dim l_partida As Long
    
    se_cumple = True

    subcond = f_quitar_segmento_pri(subcond, "NP")
    subcond = f_quitar_segmento_pri(subcond, "=")
    If IsNumeric(subcond) Then
        l_partida = CLng(subcond)
        If PA_agente_pri(age) = l_partida Then
            se_cumple = True
        Else
            se_cumple = False
        End If
    Else
        subcond = f_quitar_segmento_pri(subcond, "MULTIPLO")
        subcond = f_quitar_segmento_pri(subcond, "DE")
        If IsNumeric(subcond) Then
            l_partida = CLng(subcond)
            If PA_agente_pri(age) Mod l_partida = 0 Then
                se_cumple = True
            Else
                se_cumple = False
            End If
        Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: partida no numerica"
        End If
    End If


    f_trat_NP_pri = se_cumple

End Function
Function f_trat_tanto_por_ciento_pri(subcond As String) As Boolean
    
    Dim se_cumple As Boolean
    Dim tanto As Integer
    
    se_cumple = True
    subcond = Trim(subcond)
    'Quito el %
    subcond = Left(subcond, Len(subcond) - 1)
    
    If IsNumeric(subcond) Then
        tanto = CInt(subcond)
        If tanto < 0 Or tanto > 100 Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: debe estar entre 0 y 100"
        Else
            If fi_azar1(100) <= tanto Then
                se_cumple = True
            Else
                se_cumple = False
            End If
        End If
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: valor no numerico"
    End If
    
    
    f_trat_tanto_por_ciento_pri = se_cumple

End Function

Function f_ejecutar_accion_regla_pri(ByVal accion_regla As String) As Integer

    'byval para que no desaparezca la accion de la regla
    'aunque la vay borrando
    
    Dim jugada As Integer
    Dim tanto As Integer
    Dim se_cumple As Boolean

    se_cumple = True
    
    jugada = f_quitar_trozo_CD_pri(accion_regla)
    
    If Len(accion_regla) > 1 Then
        accion_regla = f_quitar_segmento_pri(accion_regla, "(")
        accion_regla = Left(accion_regla, Len(accion_regla) - Len(")"))
        accion_regla = Left(accion_regla, Len(accion_regla) - Len("%"))
        If IsNumeric(accion_regla) Then
            tanto = CInt(accion_regla)
            If tanto < 0 Or tanto > 100 Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error: debe estar entre 0 y 100"
            Else
                If fi_azar1(100) <= tanto Then
                    se_cumple = True
                Else
                    se_cumple = False
                End If
            End If
        Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: valor no numerico"
        End If
    End If

    
    If Not se_cumple Then
        If jugada = CTE_C Then
            jugada = CTE_D
        Else
            jugada = CTE_C
        End If
    End If

    f_ejecutar_accion_regla_pri = jugada

End Function


Function f_quitar_segmento_pri(subregla As String, segmento As String) As String

    Dim tmp As String

    subregla = Trim(subregla)
    tmp = Left(subregla, Len(segmento))
    If tmp <> segmento Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error: segmento incorrecto"
    Else
        subregla = Right(subregla, Len(subregla) - Len(segmento))
    End If

    f_quitar_segmento_pri = subregla
    
End Function

Function f_convertir_jugada_pri(s_jugada As String) As Integer
    
    Dim i_jugada As Integer
    
    If s_jugada = "COOPERAR" Or s_jugada = "C" Then
        i_jugada = CTE_C
    ElseIf s_jugada = "DEFRAUDAR" Or s_jugada = "D" Then
        i_jugada = CTE_D
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa acción"
    End If
    
     f_convertir_jugada_pri = i_jugada
End Function


Function f_quitar_trozo_CD_pri(trozo As String) As Integer
    
    'La variable es de entrada salida
    'se quita el trozo de cooperar y tal
    'y se devuelve el valor de la accion
    
    trozo = Trim(trozo)
    
    If Left(trozo, Len("COOPERAR")) = "COOPERAR" Then
        f_quitar_trozo_CD_pri = CTE_C
        trozo = f_quitar_segmento_pri(trozo, "COOPERAR")
    ElseIf Left(trozo, Len("C")) = "C" Then
        f_quitar_trozo_CD_pri = CTE_C
        trozo = f_quitar_segmento_pri(trozo, "C")
    ElseIf Left(trozo, Len("DEFRAUDAR")) = "DEFRAUDAR" Then
        f_quitar_trozo_CD_pri = CTE_D
        trozo = f_quitar_segmento_pri(trozo, "DEFRAUDAR")
    ElseIf Left(trozo, Len("D")) = "D" Then
        f_quitar_trozo_CD_pri = CTE_D
        trozo = f_quitar_segmento_pri(trozo, "D")
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa acción"
    End If

End Function
Sub s_mostrar_info_pri()

    'Muestro el resumen de pesos por tipos: la suma de pesos por cada tipo
    s_mostrar_resumen_pesos_pri
    
    'Muestro todos los pesos de todos los agentes
    'frm_a4_inpri.tipos.Clear
    'For i = 1 To numero_total_de_agentes_ejv
    '    frm_a4_inpri.tipos.AddItem agente_tipo_va0(i) & ":" & peso_agente_va0(i)
    'Next i
    
    'El ciclo actual
    frm_a4_inpri.txt_ciclo.Caption = ciclo_ejv
    'suma de todas la que nacen en ese ciclo
    frm_a4_inpri.txt_nacen.Caption = suma_nacen_va0
    'suma de todas la que mueren en ese ciclo
    frm_a4_inpri.txt_mueren.Caption = suma_mueren_va0
    'suma de todas la que mueren por vejez en ese ciclo
    frm_a4_inpri.txt_MuerenVejez.Caption = suma_mueren_vejez_va0
    
    'Actualizo
    frm_a4_inpri.Refresh


End Sub

Sub s_grabar_resumen_pri()

    Dim i As Integer
    Dim tipo As Integer
    Dim peso As Long
    Dim linea As String
        
    'Preparamos datos
    ReDim resumen_actual(1 To num_tipos_agentes_va0) As Long
    'Inicializo
    For i = 1 To num_tipos_agentes_va0
    resumen_actual(i) = 0
    Next i
    'Sumo los pesos por tipo
    For i = 1 To numero_total_de_agentes_ejv
    tipo = agente_tipo_va0(i)
    peso = peso_agente_va0(i)
    resumen_actual(tipo) = resumen_actual(tipo) + peso
    Next i
    'Grabamos datos
    linea = ""
    linea = linea & f_comillas(CStr(ciclo_ejv)) ' el ciclo actual
    For i = 1 To num_tipos_agentes_va0
    linea = linea & ";" & f_comillas(CStr(resumen_actual(i)))
    Next i
    s_grabar_dato_fichero_salida_ejv CTE_FIC_23W_1EJGRA, linea
        
        

End Sub

Sub s_mostrar_resumen_pesos_pri()

    Dim i As Integer
    Dim sumae As Long
    Dim ciclo_a_mostrar As Long
    
    ReDim temp_pesos(1 To num_tipos_agentes_va0) As Long
    ReDim temp_tipos(1 To num_tipos_agentes_va0) As Long

    sumae = 0

    '''Muestro siempre el ultimo ciclo del que tenemos información
    ''ciclo_a_mostrar = ciclo_ejv - 1
    ''If ciclo_a_mostrar = 0 Then Exit Sub
    
    'Copio lo que voy a mostrar en un temporal
    For i = 1 To num_tipos_agentes_va0
        temp_pesos(i) = peso_agente_va0(i)
        temp_tipos(i) = i
    Next i

    'Muestro el resumen
    frm_a4_inpri.tipos.Clear
    
    If Not mostrar_por_orden_de_pesos_pri Then
        'Muestro por orden de tipos
        For i = 1 To num_tipos_agentes_va0
            frm_a4_inpri.tipos.AddItem "Tipo " & f_espacios_izquierda(CStr(i), 3) & ":" & temp_pesos(i) & " " & nombre_tipo_jugador_pri(i)
            sumae = sumae + temp_pesos(i)
        Next i
    Else
        'Los muestro ordenados por peso
        'Ordeno de > a < el array temp_tipos que funciona como indice, en funcion de los valores de temp_pesos
        S_OrdenarEspecialLng temp_pesos(), temp_tipos()
        'Muestro por orden de pesos
        For i = 1 To num_tipos_agentes_va0
            frm_a4_inpri.tipos.AddItem "Tipo " & f_espacios_izquierda(CStr(temp_tipos(i)), 3) & ":" & temp_pesos(temp_tipos(i)) & " " & nombre_tipo_jugador_pri(temp_tipos(i))
            sumae = sumae + temp_pesos(i)
        Next i
    End If

    'Muestro la energia total
    frm_a4_inpri.txt_senergía = Format(sumae, "0.0000")

    'Actualizo
    frm_a4_inpri.Refresh

End Sub

Sub s_ver_jugar_contra_ordenador_pri()
    
    frm_c3_juego3r.Show CTE_MODAL

End Sub

Sub s_grabar_opciones_pri()

    
    '1 Fichero de Mapa
    '2 Fichero de definición de jugadores
    '3 numero de partidas al prisionero que juegan 2 agentes cada vez que se encuentran
    num_part_pri = CInt(frm_a4_oppri.Op_num_part)
    '4 energia_consumida_al_mover_va0 una posicion
    energia_consumida_al_mover_va0 = CDbl(frm_a4_oppri.Op_EnergiaConsumUnaPos.Text)
    '6 Algoritmo de busqueda de un espacio libre cercano
    'algoritmo_busqueda_va0
    '7 Tipo de juego
    If frm_a4_oppri.Op_TodosContraTodos Then
        todos_contra_todos_pri = True
    Else
        todos_contra_todos_pri = False
    End If
    '9 energia_consumida_al_reproducirse_va0
    energia_consumida_al_reproducirse_va0 = CDbl(frm_a4_oppri.Op_energiaConsumidaReproducirse.Text)
    '12 energia inicial de cada agente
    energia_inicial_agente_va0 = CDbl(frm_a4_oppri.Op_EnergiaInicialAgente.Text)
    '16 Probabilidad de Error en la decision
    probabilidad_de_error_pri = CInt(frm_a4_oppri.Op_ProbbError.Text)
    'Puntos ganados en el juego
    ambos_cooperan_pri = CInt(frm_a4_oppri.ambos_cooperan)
    ambos_defraudan_pri = CInt(frm_a4_oppri.ambos_defraudan)
    el_que_coopera_pri = CInt(frm_a4_oppri.el_que_coopera)
    el_que_defrauda_pri = CInt(frm_a4_oppri.el_que_defrauda)
    
    
End Sub

