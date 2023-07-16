Attribute VB_Name = "bas_c0_ce"
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
   
'===========================================
'Opciones Generales de Computacion Evolutiva
'===========================================
    '1 Modo de Ejecución
Global ver_agentes_ce0 As Boolean



Sub s_comenzar_ce0()

    ciclo_ejv = 0
    es_la_primera_vez_ejv = True
    
    s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, "Procesando...", 1, 4
    
    frm_c3_in3r.Show
    frm_c3_in3r.Caption = "Información Jugadores"
    
    frm_c3_in3r.entidad = ""
    frm_c3_in3r.accion = ""
        
    hay_que_detener_ejv = False
    hay_que_terminar_ejv = False
    esta_detenido_ejv = False
    esta_terminado_ejv = False
    
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, False
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, False
    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_EVALUACION, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_SELECCION, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION, False
            s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES, False
            s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO, False
    
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, False
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_MEJORES, False
    
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
    
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, True
    
    s_cambiar_estado_enabled_menus_ejv CTE_VER_JUGAR_CONTRA_ORDENADOR, False
    
    
    
    s_borrar_tiempo_comienzo
        
    s_grabar_opciones_3r
    'Si el numero de reglas es variable
    s_modificar_tamanio_agentes_3r
    s_crear_agentes_iniciales
    se_han_creado_los_agentes_3r = True
    s_bucle_general_ce0
    

End Sub

Sub s_bucle_general_ce0()

    'Ojo que la primera vez se pasa por "s_comenzar" se cargan
    'las opciones, pero el resto de las veces no se ejecuta y los cambios
    'se ignoran

    frm_c0_ce.BackColor = cct_ejv(cfondo_ejv)

    esta_detenido_ejv = False
    esta_terminado_ejv = False
    hay_que_detener_ejv = False
    hay_que_terminar_ejv = False

    'Pongo inhabilitado todos los programas
    s_cambiar_estado_enabled_programas_todos_ejv False

    s_botones_activos_3r False
    
    estado_3r = CTE_FUNCIONANDO
    s_mostrar_estado_semaforo frm_c3_in3r, estado_3r
    
    If es_la_primera_vez_ejv Then
        'Primera Vez: tomamos el tiempo
        s_leer_tiempo_inicial_ejv
        s_mostrar_fecha_hora_comienzo_ejv
        s_leer_tiempo_final_ejv
        s_mostrar_fecha_hora_actual_ejv
        
        frm_c3_in3r.txt_ciclo = ciclo_ejv
        frm_c3_in3r.maximo = Format(peso_agente_ce0(1), "0.00000000")
        frm_c3_in3r.minimo = Format(peso_agente_ce0(numero_total_de_agentes_ejv), "0.00000000")
        
        frm_c3_in3r.coco.Visible = False
        frm_c3_in3r.coco2.Visible = False
        frm_c3_in3r.cc.Visible = False
        
        'Pintamos los 20 primeros agentes
        If ver_agentes_3r Then
            s_pintar_mejores_agentes_3r
        End If
        ciclo_ejv = 0
        
    End If
    
    While hay_que_detener_ejv = False
        DoEvents
        ciclo_ejv = ciclo_ejv + 1
        
        'Modo automatico
        If automatico_ejv Then
            s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, ciclo_ejv, 2, ciclo_ejv
            If ciclo_ejv >= 5 Then
                'hay_que_terminar_ejv = True
                finalizacion_usuario_ejv = False
                s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
            End If
        End If

        'El ciclo actual
        frm_c3_in3r.txt_ciclo = ciclo_ejv
        'Si el numero de reglas es variable
        s_modificar_tamanio_agentes_3r
        DoEvents
        'Se evaluan las entidades - soluciones  existentes
        DoEvents
        s_evaluar_agentes_3r
        'Se clasifican las entidades en función de sus pesos de > a <
        DoEvents
        frm_c3_in3r.accion = "Ordenando"
        s_ordenar_agentes_3r
        DoEvents
        'Una vez calculado el peso en la forma normal -independiente-
        'Si esta la opcion de calcularlo relativo
        DoEvents
        If pesos_relativos_3r Then
            'Se ajustan los pesos de las entidades es como volver a evaluar
            DoEvents
            frm_c3_in3r.accion = "Ajustando"
            s_ajustar_pesos_agentes_3r
            DoEvents
            'Se clasifican de nuevo las entidades en función de sus nuevos pesos
            DoEvents
            frm_c3_in3r.accion = "Ordenando"
            s_ordenar_agentes_3r
            DoEvents
        End If
        'Pintamos los 20 primeros agentes y almacenamos las graficas y
        'miramos si se ha llegado al objetivo
        DoEvents
        
        frm_c3_in3r.maximo = Format(peso_agente_ce0(1), "0.00000000")
        frm_c3_in3r.minimo = Format(peso_agente_ce0(numero_total_de_agentes_ejv), "0.00000000")
        
        If ver_agentes_3r Then
            s_pintar_mejores_agentes_3r
        End If
        'Grabamos datos para graficas y mostramos estado actual
        s_grabar_resumen_ejv
        s_mostrar_info_ejv
        DoEvents
        'Reproducción y mutaciones
        frm_c3_in3r.accion = "Reproduciendo"
        s_reproducir_agentes_3r
        DoEvents
        'quitamos las reglas repetidas en cada agente
        If quitar_reglas_repetidas_3r Then
            frm_c3_in3r.accion = "Quitando R. Repetidas"
            s_quitar_reglas_repetidas
        End If
        'quitamos las reglas con peso menor que el indicado
        frm_c3_in3r.accion = "Quitando R. de P. bajo"
        s_quitar_reglas_peso_bajo
        'Muestro el numero de segundos trascurrido y la fecha-hora actual y la media de la duración de los ciclos
        s_mostrar_tiempo_transcurrido_ejv
        s_condiciones_parada_ejv CTE_PARADA_POR_IGUAL
    Wend
    
    If hay_que_terminar_ejv And ver_agentes_3r Then
        Unload frm_c3_juego3r
    End If
    
    'Se clasifican las entidades en función de sus pesos de > a <
    DoEvents
    frm_c3_in3r.accion = "Ordenando"
    s_ordenar_agentes_3r
   
    frm_c0_ce.fr_Todas.Visible = False
    frm_c0_ce.Fr_Ejecucion.Visible = True
    
    estado_3r = CTE_DETENIDO
    
    s_botones_activos_3r True
    
    s_fin_bucle_general_ejv
    
End Sub


Sub s_ver_opciones_ce0()

    frm_c0_ce.fr_Todas.Visible = False
    frm_c0_ce.Fr_Ejecucion.Visible = False
    
    s_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, True

    If ciclo_ejv > 0 Then
        s_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_MEJORES, True
    End If
    
    'Con esto borro el grafico
    frm_c0_ce.Refresh

End Sub

Function f_generar_cadena_al_azar_ce0(longitud As Long, alfabeto() As String) As String

    Dim i As Long
    Dim tamaño_alfabeto As Long
    
    f_generar_cadena_al_azar_ce0 = ""
    tamaño_alfabeto = UBound(alfabeto)
    For i = 1 To longitud
        f_generar_cadena_al_azar_ce0 = f_generar_cadena_al_azar_ce0 & alfabeto(fl_azar1(tamaño_alfabeto))
    Next i

End Function
