Attribute VB_Name = "bas_a0_va"
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

'=====================================
'Opciones Generales de Vida Artificial
'=====================================
    '1 Modo de Ejecución
Global ver_agentes_va0 As Boolean
    '2 Lugar de nacimiento
Global nacimiento_cerca_va0 As Boolean
    '3 tasas de mutacion
Global probb_mutacion_tipo_inicial_va0 As Double
Global probb_mutacion_mov_inicial_va0 As Double
Global probb_mutacion_pm_inicial_va0 As Double
Global PMPMCte_va0 As Boolean
    '4 agentes inmortales
Global agentes_inmortales_va0 As Boolean
Global muerte1_va0 As Integer
Global muerte2_va0 As Integer
    '5 Búsqueda de Cadena binaria
Global busqueda_cadena_binaria_va0 As Boolean
Global cadena_binaria_buscada_va0 As String
Global long_cadena_buscada_va0 As Long
    '6 Limite Muerte
Global limite_muerte_va0 As Long

Global energia_inicial_agente_va0 As Double
Global nombre_fichero_mapa_va0 As String
Global energia_consumida_al_mover_va0 As Double
Global energia_consumida_al_reproducirse_va0 As Double
Global numero_de_posiciones_alejar_reproducirse_va0 As Integer
Global numero_de_plantas_nacen_ciclo_va0 As Integer
Global algoritmo_busqueda_va0 As Integer




'Numero inicial de agentes de cada tipo actual (suma de todos los tipos)
'Este solo sirve para redimensionar 2 dimensiones
Global numero_inicial_de_agentes_va0 As Integer


'====================================================
'Datos del agente (hormiga o prisionero)
'Agente: Comunes
Global agente_z_va0() As Double 'Comun 0
Global agente_x_va0() As Double 'Comun 1
Global agente_y_va0() As Double 'Comun 2
Global agente_tipo_va0() As Integer 'Comun 3
Global peso_agente_va0() As Double 'Comun 4
Global agente_direccion_anterior_va0() As Integer 'Comun 5
Global agente_accion_anterior_va0() As Integer 'Comun 6
Global apellidos_agente_va0() As String 'Comun 7
Global muerte_agente_va0() As Integer 'Comun 8
Global ciclo_nacimiento_agente_va0() As Long 'Comun 9
Global agente_probb_mutacion_tipo_va0() As Double 'Comun 10
Global agente_probb_mutacion_mov_va0() As Double 'Comun 11
Global agente_probb_mutacion_pm_va0() As Double 'Comun 12
Global agente_tendencia_rel_mov_va0() As Long 'Comun 13 1-8,ag
Global agente_tendencia_abs_mov_va0() As Long 'Comun 14 1-8,ag
Global cadena_binaria_va0() As String 'Comun 15
Global sexo_va0() As Integer 'Comun 16
'Agente: hyp
Global hormiga_probabilidad_ganar_hyp() As Double 'hyp 1
'Agente: pri
Global historia_agente_yo_pri() As Integer 'pri 1
Global historia_agente_el_pri() As Integer 'pri 2
Global PA_agente_pri() As Integer ' pri 3 (PA es Partida Actual)
'Agente: explorando
Global mapa_exp() As Integer 'exp l
'Agente: uva
Global agente_mas_cercano_uva() As Integer 'uva 1
Global dist_al_mas_cercano_uva() As Double 'uva 2
'====================================================



'====================================================
'Datos de Tipos de agentes
'Tipos de Agentes: Comunes
Global num_tipos_agentes_va0 As Integer
Global numero_agentes_que_se_deben_crear_inicio_total_va0 As Integer
Global esta_modificado_num_agen_tipo_pri As Boolean
Global numero_agentes_que_se_deben_crear_inicio_de_tipo_va0() As Integer
Global num_agentes_tipo_va0() As Integer
Global tendencia_rel_inicial_mov_tipo_agente_va0() As Long '1-8,tipo
Global tendencia_abs_inicial_mov_tipo_agente_va0() As Long '1-8,tipo
'Tipos de Agentes: hyp
'Tipos de Agentes: pri
Global nombre_tipo_jugador_pri() As String
Global prioridad_regla_tipo_jugador_pri() As Integer
Global condicion_regla_tipo_jugador_pri() As String
Global accion_regla_tipo_jugador_pri() As String
'====================================================



Global se_ha_empezado_a_crear_agentes_va0 As Boolean

'Tendencias de movimiento
Global tipo_tendencia_en_modificacion_va0 As Integer
Global lista_tendencias_en_modificacion_va0() As Long
Global ha_habido_cambio_lista_tendencias_va0 As Boolean
Global tipo_agente_cambiar_tendencias_va0 As Integer



Global apellidos_posibles(1 To CTE_numero_maximo_apellidos) As String
Global c_d_ape As Integer
Global cont_apellidos_usados_va0 As Integer
Global ha_cambiado_el_diccionaro_pal As Integer

'Mapa
Global ver_zoom_va0 As Integer
Global separacion_mapa_va0 As Integer

Global mapa_pisos_va0 As Double
Global mapa_filas_va0 As Double
Global mapa_columnas_va0 As Double

Global viejo_mapa_pisos_va0 As Integer
Global viejo_mapa_filas_va0 As Integer
Global viejo_mapa_columnas_va0 As Integer

Global mapa_va0() As Integer
Global nodo_visitado_va0() As Integer
Global habilitar_change_zoom_va0 As Boolean
Global mapa_sin_obstaculos_va0 As Boolean

'El mapa podria ser booleano ya que para saber si hay o no agente en
'una determinada posicion se pueden recorrer todos los agentes analizando
'sus x y sus y, pero estas funciones se usan mucho, y es seguro que es
'mas rentable mantener en el mapa las posiciones ocupadas por agentes
'Lo que puede haber en un solo punto de la pantalla
'0:vacio CTE_MAPA_VACIO = 0
'1:obstaculo CTE_MAPA_OBSTACULO = 1
'2:agente(hormiga o prisionero) CTE_MAPA_AGENTE = 2
'3:planta CTE_MAPA_PLANTA = 3


'Test
Global estado_test_movimiento_va0 As Boolean
Global direccion_test_va0 As Integer
Global direccion_old_test_va0 As Integer
Global cursor_z_va0 As Double
Global cursor_y_va0 As Double
Global cursor_x_va0 As Double
Global cursor_old_z_va0 As Double
Global cursor_old_y_va0 As Double
Global cursor_old_x_va0 As Double
Global num_direcc_algoritmo_va0 As Integer
Global num_direcc_old_algoritmo_va0 As Integer

Sub s_bucle_general_va0()

    frm_a0_va.BackColor = cct_ejv(cfondo_ejv)
    
    esta_detenido_ejv = False
    esta_terminado_ejv = False
    hay_que_detener_ejv = False
    hay_que_terminar_ejv = False
    
    'Pongo inhabilitado todos los programas
    s_cambiar_estado_enabled_programas_todos_ejv False
    
    s_operacion_ver_ejv CTE_VER_ESTADO_EJECUCION
    
    If es_la_primera_vez_ejv Then
        'Primera Vez: tomamos el tiempo
        s_leer_tiempo_inicial_ejv
        s_mostrar_fecha_hora_comienzo_ejv
        s_leer_tiempo_final_ejv
        s_mostrar_fecha_hora_actual_ejv
        
        agente_contrario_pri = 1
        agente_actual_ejv = 1
        
        Select Case num_prg_activo_ejv
            Case CTE_HYP '1
                s_inicializar_hyp 'aqui se crean los agentes iniciales
            Case CTE_PRI '4
                s_inicializar_pri 'aqui se crean los agentes iniciales
            Case CTE_CEL '5
                s_inicializar_cel 'aqui se crean los agentes iniciales
            Case CTE_GAI '6
                s_inicializar_gai 'aqui se crean los agentes iniciales
            Case CTE_EXP '7
                s_inicializar_exp 'aqui se crean los agentes iniciales
            Case CTE_PEZ '9
                s_inicializar_pez 'aqui se crean los agentes iniciales
            Case CTE_UVA '10
                s_inicializar_uva 'aqui se crean los agentes iniciales
            Case CTE_YXY '11
                s_inicializar_yxy 'aqui se crean los agentes iniciales
            Case Else
                s_error_num_prog num_prg_activo_ejv
        End Select
        
        ciclo_ejv = 1
    End If
    
    While hay_que_detener_ejv = False
        DoEvents
        's_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, ciclo_ejv, 2, ciclo_ejv
        'El ciclo actual
        Select Case num_prg_activo_ejv
            Case CTE_HYP '1
                frm_a1_inhyp.txt_ciclo.Caption = ciclo_ejv
            Case CTE_PRI '4
                frm_a4_inpri.txt_ciclo.Caption = ciclo_ejv
            Case CTE_CEL '5
                frm_a5_incel.txt_ciclo.Caption = ciclo_ejv
            Case CTE_GAI '6
                frm_a6_ingaia.txt_ciclo.Caption = ciclo_ejv
            Case CTE_EXP '7
                frm_a7_inexp.txt_ciclo.Caption = ciclo_ejv
            Case CTE_PEZ '9
                frm_a9_inpez.txt_ciclo.Caption = ciclo_ejv
            Case CTE_UVA '10
                frm_aA_inuva.txt_ciclo.Caption = ciclo_ejv
            Case CTE_YXY '11
            Case Else
                s_error_num_prog num_prg_activo_ejv
        End Select
        suma_riegan_hyp = 0
        suma_nacen_va0 = 0
        suma_mueren_va0 = 0
        suma_mueren_vejez_va0 = 0
        suma_pelean_hyp = 0
        
        If num_prg_activo_ejv = CTE_PRI Then
            frm_a4_inpri.txt_total_agentes.Caption = numero_total_de_agentes_ejv
        End If
        
        If num_prg_activo_ejv = CTE_PRI And todos_contra_todos_pri Then
            '======================================================================'
            '      COMIENZO DEL CASO ESPECIAL PRISIONERO - TODOS CONTRA TODOS      '
            '======================================================================'
            'Ejecuto partidas de dos en dos todos contra todos
                'Pinto el agente
                If ver_agentes_va0 Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(agente_actual_ejv), agente_y_va0(agente_actual_ejv), agente_x_va0(agente_actual_ejv), CTE_PRISIONERO, ccs_ejv(f_SumCirc(ncs_i_ejv, agente_tipo_va0(agente_actual_ejv), 0)), cct_ejv(cfondo_ejv), agente_direccion_anterior_va0(agente_actual_ejv), ver_zoom_va0, 1
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(agente_contrario_pri), agente_y_va0(agente_contrario_pri), agente_x_va0(agente_contrario_pri), CTE_PRISIONERO, ccs_ejv(f_SumCirc(ncs_i_ejv, agente_tipo_va0(agente_contrario_pri), 0)), cct_ejv(cfondo_ejv), agente_direccion_anterior_va0(agente_contrario_pri), ver_zoom_va0, 1
                End If
                s_jugar_partidas agente_actual_ejv, agente_contrario_pri
                DoEvents
                frm_a4_inpri.txt_agente.Caption = CStr(agente_actual_ejv) & "-" & CStr(agente_contrario_pri)
                'Borro el agente
                If ver_agentes_va0 Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(agente_contrario_pri), agente_y_va0(agente_contrario_pri), agente_x_va0(agente_contrario_pri), CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                End If
                'Calculo los nuevos agentes
                agente_contrario_pri = agente_contrario_pri + 1
                If agente_contrario_pri = numero_total_de_agentes_ejv + 1 Then
                    'Fin de ciclo
                    s_grabar_resumen_ejv
                    s_mostrar_info_ejv
                    ciclo_ejv = ciclo_ejv + 1
                    agente_contrario_pri = 1
                    'Borro el agente
                    If ver_agentes_va0 Then
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(agente_actual_ejv), agente_y_va0(agente_actual_ejv), agente_x_va0(agente_actual_ejv), CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                    End If
                    'Paso el indice al siguiente
                    agente_actual_ejv = agente_actual_ejv + 1
                End If
            '======================================================================'
            '           FIN DEL CASO ESPECIAL PRISIONERO - TODOS CONTRA TODOS      '
            '======================================================================'
        Else
            '===================================='
            '      COMIENZO DEL CASO NORMAL      '
            '===================================='
            'Ejecuto una vez cada agente
            While agente_actual_ejv <= numero_total_de_agentes_ejv And hay_que_detener_ejv = False
                'El agente actual
                Select Case num_prg_activo_ejv
                    Case CTE_HYP '1
                        frm_a1_inhyp.txt_agente.Caption = agente_actual_ejv
                    Case CTE_PRI '4
                        frm_a4_inpri.txt_agente.Caption = agente_actual_ejv
                    Case CTE_CEL '5
                        frm_a5_incel.txt_agente.Caption = agente_actual_ejv
                    Case CTE_GAI '6
                        frm_a6_ingaia.txt_agente.Caption = agente_actual_ejv
                    Case CTE_EXP '7
                        frm_a7_inexp.txt_agente.Caption = agente_actual_ejv
                    Case CTE_PEZ '9
                        frm_a9_inpez.txt_agente.Caption = agente_actual_ejv
                    Case CTE_UVA '10
                        frm_aA_inuva.txt_agente.Caption = agente_actual_ejv
                    Case CTE_YXY '11
                    Case Else
                        s_error_num_prog num_prg_activo_ejv
                End Select
                s_ejecutar_agente_va0
                DoEvents
                'Paso el indice al siguiente
                agente_actual_ejv = agente_actual_ejv + 1
                'Muestro el numero de segundos trascurrido y la fecha-hora actual
                s_mostrar_tiempo_transcurrido_ejv
            Wend
            'Si he parado por fin de ciclo, me coloco en el primero para el siguiente ciclo
            'Si he parado por accion del usuario, no
            If agente_actual_ejv > numero_total_de_agentes_ejv Then 'Podria ser mayor que +1 si la ultima accion fue eliminar agente
                'Fin de ciclo
                s_grabar_resumen_ejv
                s_mostrar_info_ejv
                ciclo_ejv = ciclo_ejv + 1
                agente_actual_ejv = 1
            End If
            '===================================='
            '           FIN DEL CASO NORMAL      '
            '===================================='
        End If
        s_condiciones_parada_ejv CTE_PARADA_POR_MAYOR 'el ciclo actual es en realidad uno mas que el verdadero
        DoEvents
        If num_prg_activo_ejv = CTE_HYP Then
            If hay_que_detener_ejv = False Then
                s_crear_mas_plantas_hyp
            End If
        End If
        DoEvents
        'Muestro el numero de segundos trascurrido y la fecha-hora actual y la media de la duración de los ciclos
        s_mostrar_tiempo_transcurrido_ejv
    Wend
    
    s_fin_bucle_general_ejv
    
End Sub

Sub s_mostrar_info_ejv()

        Select Case num_prg_activo_ejv
            Case CTE_HYP '1
                s_mostrar_info_hyp
            Case CTE_PAL '2
                s_mostrar_info_pal
            Case CTE_3R '3
                s_mostrar_info_3r
            Case CTE_PRI '4
                s_mostrar_info_pri
            Case CTE_CEL '5
                s_mostrar_info_cel
            Case CTE_GAI '6
                s_mostrar_info_gai
            Case CTE_EXP '7
                s_mostrar_info_exp
            Case CTE_CAD '8
                s_mostrar_info_cad
            Case CTE_PEZ '9
                s_mostrar_info_pez
            Case CTE_UVA '10
                s_mostrar_info_uva
            Case CTE_YXY '11
            Case Else
                s_error_num_prog num_prg_activo_ejv
        End Select

End Sub


Sub s_grabar_resumen_ejv()

    '========================================================================
    'Grabar Resumen
    If ciclo_ejv <= max_guardado_ejv Then
        If un_ej_grabar_gra_ejv Then
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_grabar_resumen_hyp
                Case CTE_PAL '2
                    s_grabar_resumen_pal
                Case CTE_3R '3
                    s_grabar_resumen_3r
                Case CTE_PRI '4
                    s_grabar_resumen_pri
                Case CTE_CEL '5
                    s_grabar_resumen_cel
                Case CTE_GAI '6
                    s_grabar_resumen_gai
                Case CTE_EXP '7
                    s_grabar_resumen_exp
                Case CTE_CAD '8
                    s_grabar_resumen_cad
                Case CTE_PEZ '9
                    s_grabar_resumen_pez
                Case CTE_UVA '10
                    s_grabar_resumen_uva
                Case CTE_YXY '11
                Case Else
                    s_error_num_prog num_prg_activo_ejv
            End Select
        End If
        'Las anteriores acciones escriben en los ficheros, pero no los guardan
        'Si han transcurrido ya autoguardado_ejv ciclos desde la ultima vez que grabe el fichero, lo grabo
        If ciclo_ejv Mod autoguardado_ejv = 0 Then
            'Grabo los ficheros sin cerrarlos, por si hay un corte de luz y esas cosas
            s_grabar_ficheros_un_ejemplo_ejv
        End If
    End If
    '========================================================================
            

End Sub

Sub s_ejecutar_agente_va0()
    
    Dim vecinos As String
    Dim cont As Integer
    Dim direccion As Integer
    Dim mis_vecinos(1 To CTE_8_DIR) As Integer
    Dim accion As String
    Dim i As Integer
    Dim he_comido_planta As Boolean
    Dim muerte_por_vejez As Boolean
        
    
    'control errores de programacion
    If control_errores_de_programacion_ejv Then
        If ciclo_nacimiento_agente_va0(1) > ciclo_ejv Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: nacimiento"
        End If
    End If
        
    'Si es viejo y no son inmortales, se muere
    muerte_por_vejez = False
    If agentes_inmortales_va0 = False And ciclo_ejv - ciclo_nacimiento_agente_va0(agente_actual_ejv) >= muerte_agente_va0(agente_actual_ejv) Then
        muerte_por_vejez = True
        suma_mueren_vejez_va0 = suma_mueren_vejez_va0 + 1
    End If
    'Si no tiene energía, está muerto y paso al siguiente
    'solo para hyp, los prisioneros nunca mueren por falta de peso
    If (peso_agente_va0(agente_actual_ejv) <= limite_muerte_va0 And ciclo_ejv > 1 And num_prg_activo_ejv = CTE_HYP) Or muerte_por_vejez Then
        'Elimino el agente
        s_eliminar_agente_va0 agente_actual_ejv
        suma_mueren_va0 = suma_mueren_va0 + 1
    Else
         
        'Decidir accion
        'El array mis_vecinos() es un parametro de entrada/salida
        If num_prg_activo_ejv <> CTE_UVA Then
            vecinos = f_ver_vecinos_va0(agente_actual_ejv, mis_vecinos())
        End If
        If num_prg_activo_ejv = CTE_HYP Then
            'Antes de realizar la acción en sí, se hace
            '1.- Si hay una o mas plantas al lado comestible, se comen (siempre) todas
            he_comido_planta = False
            If vecinos = CTE_VEC_PLANTA Or vecinos = CTE_VEC_AGENTEYPLANTA Then
               'Por cada planta a mi alrededor (siempre se miran en el mismo orden)
               For cont = 1 To CTE_8_DIR
                   If mis_vecinos(cont) = CTE_VEC_PLANTA Then
                        'Me como la planta si es comestible
                        If f_intentar_comer_planta_hyp(agente_actual_ejv, cont) Then
                            he_comido_planta = True
                        End If
                   End If
               Next cont
            End If
            
            'Se vuelven a comprobar los vecinos, ya que nos hemos podido comer una planta
            If he_comido_planta Then
                vecinos = f_ver_vecinos_va0(agente_actual_ejv, mis_vecinos())
            End If
        End If
        
        
        Select Case num_prg_activo_ejv
            Case CTE_HYP '1
                'En las hyp la accion depende tambien del
                'tipo de agente
                Select Case agente_tipo_va0(agente_actual_ejv)
                   Case 1
                       Select Case vecinos
                           Case CTE_VEC_NADA
                               accion = f_calcular_accion_hyp(1)
                           Case CTE_VEC_AGENTE
                               accion = f_calcular_accion_hyp(2)
                           Case CTE_VEC_PLANTA
                               accion = f_calcular_accion_hyp(3)
                           Case CTE_VEC_AGENTEYPLANTA
                               accion = f_calcular_accion_hyp(4)
                           Case Else
                               s_error_ejv CON_OPCION_FINALIZAR, "Error: vecino incorrecto"
                       End Select
                   Case 2
                       Select Case vecinos
                           Case CTE_VEC_NADA
                               accion = f_calcular_accion_hyp(5)
                           Case CTE_VEC_AGENTE
                               accion = f_calcular_accion_hyp(6)
                           Case CTE_VEC_PLANTA
                               accion = f_calcular_accion_hyp(7)
                           Case CTE_VEC_AGENTEYPLANTA
                               accion = f_calcular_accion_hyp(8)
                           Case Else
                               s_error_ejv CON_OPCION_FINALIZAR, "Error: vecino incorrecto"
                       End Select
                   Case 3
                       Select Case vecinos
                           Case CTE_VEC_NADA
                               accion = f_calcular_accion_hyp(9)
                           Case CTE_VEC_AGENTE
                               accion = f_calcular_accion_hyp(10)
                           Case CTE_VEC_PLANTA
                               accion = f_calcular_accion_hyp(11)
                           Case CTE_VEC_AGENTEYPLANTA
                               accion = f_calcular_accion_hyp(12)
                           Case Else
                               s_error_ejv CON_OPCION_FINALIZAR, "Error: vecino incorrecto"
                       End Select
                   Case 4
                       Select Case vecinos
                           Case CTE_VEC_NADA
                               accion = f_calcular_accion_hyp(13)
                           Case CTE_VEC_AGENTE
                               accion = f_calcular_accion_hyp(14)
                           Case CTE_VEC_PLANTA
                               accion = f_calcular_accion_hyp(15)
                           Case CTE_VEC_AGENTEYPLANTA
                               accion = f_calcular_accion_hyp(16)
                           Case Else
                               s_error_ejv CON_OPCION_FINALIZAR, "Error: vecino incorrecto"
                       End Select
                   Case 5
                       Select Case vecinos
                           Case CTE_VEC_NADA
                               accion = f_calcular_accion_hyp(17)
                           Case CTE_VEC_AGENTE
                               accion = f_calcular_accion_hyp(18)
                           Case CTE_VEC_PLANTA
                               accion = f_calcular_accion_hyp(19)
                           Case CTE_VEC_AGENTEYPLANTA
                               accion = f_calcular_accion_hyp(20)
                           Case Else
                               s_error_ejv CON_OPCION_FINALIZAR, "Error: vecino incorrecto"
                       End Select
                   Case Else
                       s_error_ejv CON_OPCION_FINALIZAR, "Error: tipo agente inexistente"
                End Select
            Case CTE_PRI '4
                'Si hay vecino, siempre se juega una partida
                'para todos los agentes
                Select Case vecinos
                    Case CTE_VEC_NADA
                        accion = CTE_ACC_MOVER
                    Case CTE_VEC_AGENTE
                        accion = CTE_ACC_JUGAR
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: vecino incorrecto"
                End Select
            Case CTE_CEL '5
            Case CTE_GAI '6
            Case CTE_EXP '7
                'La unica accion posible es moverse
                accion = CTE_ACC_MOVER
            Case CTE_PEZ '9
                'La unica accion posible es moverse
                accion = CTE_ACC_MOVER
            Case CTE_UVA '10
                'La unica accion posible es moverse
                accion = CTE_ACC_MOVER
            Case CTE_YXY '11
            Case Else
                s_error_num_prog num_prg_activo_ejv
        End Select
        
    
        'Ahora se la accion que debo ejecutar
        'siempre se debe poder ejecutar la accion excepto
        'en el caso de que la acción sea mover y la hormiga
        'esté completamente rodeada de otras hormigas y plantas
        'en cuyo caso no se mueve, pero consume energía como si
        'se hubiera movido
    
    
        'Ejecutar accion
         Select Case accion
            'Solo los prisioneros juegan
            Case CTE_ACC_JUGAR
                'No puede pelear dos veces seguidas
                If agente_accion_anterior_va0(agente_actual_ejv) <> CTE_ACC_JUGAR Then
                    s_accion_jugar_pri agente_actual_ejv, mis_vecinos()
                    'Alejo el agente el numero de veces indicado
                     If numero_de_posiciones_alejar_pelear_hyp > 0 Then
                         s_alejar_agente_va0 agente_actual_ejv, numero_de_posiciones_alejar_pelear_hyp
                     End If
                Else
                    s_accion_mover_va0 mis_vecinos()
                End If
            Case CTE_ACC_MOVER
                s_accion_mover_va0 mis_vecinos()
            Case CTE_ACC_REGAR
                s_accion_regar_hyp agente_actual_ejv, mis_vecinos()
                'Alejo la hormiga el numero de veces indicado
                If numero_de_posiciones_alejar_pelear_hyp > 0 Then
                    s_alejar_agente_va0 agente_actual_ejv, numero_de_posiciones_alejar_regar_hyp
                End If
            Case CTE_ACC_PELEAR
                'Debe tener al menos la energía necesaria para pelearse
                If peso_agente_va0(agente_actual_ejv) > energia_consumida_al_pelearse_hyp Then
                    'No puede pelear dos veces seguidas
                     If agente_accion_anterior_va0(agente_actual_ejv) <> CTE_ACC_PELEAR Then
                         s_accion_pelear_hyp agente_actual_ejv, mis_vecinos()
                        'Alejo la hormiga el numero de veces indicado
                         If numero_de_posiciones_alejar_pelear_hyp > 0 Then
                             s_alejar_agente_va0 agente_actual_ejv, numero_de_posiciones_alejar_pelear_hyp
                         End If
                     Else
                        s_accion_mover_va0 mis_vecinos()
                     End If
                 Else
                    s_accion_mover_va0 mis_vecinos()
                 End If
            Case CTE_ACC_REPRODUCIRSE
                'No puede reproducirse dos veces seguidas
                'Si la hormiga no tiene energía suficiente para reproducirse, se mueve
                If peso_agente_va0(agente_actual_ejv) > energia_consumida_al_reproducirse_va0 Then
                    If agente_accion_anterior_va0(agente_actual_ejv) <> CTE_ACC_REPRODUCIRSE Then
                         s_accion_reproducirse_va0 agente_actual_ejv, mis_vecinos()
                        'Alejo la hormiga el numero de veces indicado
                         If numero_de_posiciones_alejar_pelear_hyp > 0 Then
                            s_alejar_agente_va0 agente_actual_ejv, numero_de_posiciones_alejar_reproducirse_va0
                         End If
                    Else
                       s_accion_mover_va0 mis_vecinos()
                    End If
                Else
                    s_accion_mover_va0 mis_vecinos()
                End If
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa acción"
         End Select
        
        'Grabo la acción realizada para que sea la acción anterior
         agente_accion_anterior_va0(agente_actual_ejv) = accion
    
    End If
    
    
    'control errores de programacion
    If control_errores_de_programacion_ejv Then
        For i = 1 To num_tipos_agentes_va0
            If num_agentes_tipo_va0(i) < 0 Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error: no hay agentes"
            End If
        Next i
    End If

End Sub

Sub s_botones_enabled_va0(activo As Boolean)

    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, activo
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, activo
    
    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPOS_AGENTES, activo
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, activo
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, activo
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES3, activo
    s_cambiar_estado_enabled_menus_ejv CTE_VER_MAPA, activo
    s_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, activo
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, activo
    s_cambiar_estado_enabled_menus_ejv CTE_VER_APELLIDOS, activo
    
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, Not activo
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, Not activo


End Sub

Function fi_indice_agente_va0(Z As Double, Y As Double, X As Double) As Integer

    'Devuelve el número de agente sabidos su posicion X-Y
    'si no existe, devuelve 0
    'Cuando se llama a eta funcion, se supone que ya se sabe
    'que existe un agente en esa posicion
    Dim encontrado As Boolean
    Dim cont As Integer
    Dim devolver As Integer

    encontrado = False
    devolver = 0
    For cont = 1 To numero_total_de_agentes_ejv
        If agente_z_va0(cont) = Z And agente_y_va0(cont) = Y And agente_x_va0(cont) = X Then
            encontrado = True
            devolver = cont
            Exit For
        End If
    Next cont

    'Control errores de programacion
    If control_errores_de_programacion_ejv Then
        If Not encontrado Then
             s_error_ejv CON_OPCION_FINALIZAR, "Error: ese agente no existe"
        End If
        If devolver > 0 And mapa_va0(Z, Y, X) <> CTE_MAPA_AGENTE Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Agente en una zona no vacía"
        End If
    End If

    fi_indice_agente_va0 = devolver

End Function

Function f_esta_vacio_va0(Z As Double, Y As Double, X As Double) As Boolean

    'Z pisos
    'Y filas
    'X columnas

    Dim dev As Boolean
    
    If num_prg_activo_ejv = CTE_UVA Then
        'el agente que analizo es el que todavia no he creado y f_esta_vacio_uva usa el agente_actual_ejv
        f_esta_vacio_va0 = f_esta_vacio_uva(numero_total_de_agentes_ejv + 1, Z, Y, X)
        Exit Function
    End If
    
    'en estado test
    'analizo el mapa temporal en vez del real
    'y ademas no hay hormigas ni plantas, solo puede haber obstaculos
    'no he creado una funcion f_esta_vacio_va0 porque
    'el test y la ejecucion usan las mismas funciones y asi
    'es mas sencillo

    If estado_test_movimiento_va0 Then
        'Estoy en el editor de mapas
        s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, Z, Y, X, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
        If mapa_ma0(Z, Y, X) = CTE_MAPA_VACIO Then
            dev = True
        Else
            dev = False
        End If
    Else
        'Estoy ejecutando vida artificial
        s_ajustar_lugar_mapa CTE_MAPA_ESFERICO, Z, Y, X, mapa_pisos_va0, mapa_filas_va0, mapa_columnas_va0
        If mapa_va0(Z, Y, X) = CTE_MAPA_VACIO Then
            dev = True
        Else
            dev = False
        End If
        
        'Control errores de programacion
        If control_errores_de_programacion_ejv Then
            If se_ha_empezado_a_crear_agentes_va0 Then
                'Puede haber agentes
                Select Case num_prg_activo_ejv
                    Case CTE_HYP
                        If Not f_hay_agente_va0(Z, Y, X) And Not f_hay_planta_hyp(Z, Y, X) And Not f_hay_obstaculo_va0(Z, Y, X) And Not dev Then
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Hay algo en zona no vacía"
                        End If
                    Case CTE_PRI
                        If Not f_hay_agente_va0(Z, Y, X) And Not f_hay_obstaculo_va0(Z, Y, X) And Not dev Then
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Hay algo en zona no vacía"
                        End If
                End Select
            Else
                'No puede haber agentes
                Select Case num_prg_activo_ejv
                    Case CTE_HYP
                        If Not f_hay_planta_hyp(Z, Y, X) And Not f_hay_obstaculo_va0(Z, Y, X) And Not dev Then
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Hay algo en zona no vacía"
                        End If
                    Case CTE_PRI
                        If Not f_hay_obstaculo_va0(Z, Y, X) And Not dev Then
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Hay algo en zona no vacía"
                        End If
                End Select
            End If
        End If
    End If

    f_esta_vacio_va0 = dev


End Function


Function f_hay_agente_va0(Z As Double, Y As Double, X As Double) As Boolean

    Dim i As Integer
    Dim dev As Boolean
    Dim existe As Boolean
    
    Dim nueva_z As Double
    Dim nueva_y As Double
    Dim nueva_x As Double
    
    'Tal vez se llama a esta funcion cuando aun no se han creado los agentes
    If Not se_ha_empezado_a_crear_agentes_va0 Then
        f_hay_agente_va0 = False
        'Control errores de programacion
        If control_errores_de_programacion_ejv Then
            s_error_ejv CON_OPCION_FINALIZAR, "Aviso: se llama a f_hay_agente_va0 cuando aun no se han creado los agentes"
        End If
        Exit Function
    End If
    
    nueva_z = Z
    nueva_y = Y
    nueva_x = X
    
    'Si nos toca inspecionar una celda en la barrera
    'esto significa inspecionar en realidad la que
    'está al otro lado
    If nueva_y = 0 Then nueva_y = mapa_filas_va0
    If nueva_y = mapa_filas_va0 + 1 Then nueva_y = 1
    If nueva_x = 0 Then nueva_x = mapa_columnas_va0
    If nueva_x = mapa_columnas_va0 + 1 Then nueva_x = 1
    
    If mapa_va0(nueva_z, nueva_y, nueva_x) = CTE_MAPA_AGENTE Then
        dev = True
    Else
        dev = False
    End If
    
    'Control errores de programacion
    If control_errores_de_programacion_ejv Then
        existe = False
        For i = 1 To numero_total_de_agentes_ejv
            If agente_z_va0(i) = nueva_z And agente_y_va0(i) = nueva_y And agente_x_va0(i) = nueva_x Then
                existe = True
                Exit For
            End If
            DoEvents
        Next i
        If existe <> dev Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Agente existe en una zona vacía"
        End If
    End If

    f_hay_agente_va0 = dev

End Function
Function f_hay_obstaculo_va0(Z As Double, Y As Double, X As Double) As Boolean
    
    Dim i As Integer
    
    Dim nueva_z As Double
    Dim nueva_y As Double
    Dim nueva_x As Double
    
    nueva_z = Z
    nueva_y = Y
    nueva_x = X

    'Si nos toca inspecionar una celda en la barrera
    'esto significa inspecionar en realidad la que
    'está al otro lado
    If nueva_z = 0 Then nueva_z = mapa_pisos_va0
    If nueva_z = mapa_pisos_va0 + 1 Then nueva_z = 1
    If nueva_y = 0 Then nueva_y = mapa_filas_va0
    If nueva_y = mapa_filas_va0 + 1 Then nueva_y = 1
    If nueva_x = 0 Then nueva_x = mapa_columnas_va0
    If nueva_x = mapa_columnas_va0 + 1 Then nueva_x = 1
    
    If mapa_va0(nueva_z, nueva_y, nueva_x) = CTE_MAPA_OBSTACULO Then
        f_hay_obstaculo_va0 = True
    Else
        f_hay_obstaculo_va0 = False
    End If

End Function
Sub s_fijar_separacion_mapa_va0()
    
    Select Case ver_zoom_va0
        Case CTE_ZOOM_DETALLE
            separacion_mapa_va0 = 14
        Case CTE_ZOOM_PANORAMICA
            separacion_mapa_va0 = 4
        Case CTE_ZOOM_PIXELS
            separacion_mapa_va0 = 1
        Case CTE_ZOOM_3D
            separacion_mapa_va0 = 20
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
    End Select

End Sub
Sub s_crear_un_agente_va0(ag_z As Double, ag_y As Double, ag_x As Double, tipo As Integer, p_energia_inicial As Double, p_apellidos As String, pm_tipo As Double, pm_mov As Double, pm_pm As Double, m_abs() As Long, m_rel() As Long, m_cadena As String)

    Dim cont As Integer
    Dim p As Double 'pisos    Z
    Dim f As Double 'filas    Y
    Dim c As Double 'columnas X
    
    num_agentes_tipo_va0(tipo) = num_agentes_tipo_va0(tipo) + 1
    numero_total_de_agentes_ejv = numero_total_de_agentes_ejv + 1
    'Datos del agente
    s_redim_preserve_agente_va0
    
    agente_z_va0(numero_total_de_agentes_ejv) = ag_z 'Comun 0
    agente_y_va0(numero_total_de_agentes_ejv) = ag_y 'Comun 1
    agente_x_va0(numero_total_de_agentes_ejv) = ag_x 'Comun 2
    agente_tipo_va0(numero_total_de_agentes_ejv) = tipo 'Comun 3
    peso_agente_va0(numero_total_de_agentes_ejv) = p_energia_inicial 'Comun 4
    agente_direccion_anterior_va0(numero_total_de_agentes_ejv) = fi_azar1(CTE_8_DIR) 'Comun 5 (empieza mirando al norte)
    agente_accion_anterior_va0(numero_total_de_agentes_ejv) = CTE_ACC_NADA 'Comun 6
    apellidos_agente_va0(numero_total_de_agentes_ejv) = p_apellidos 'Comun 7
    muerte_agente_va0(numero_total_de_agentes_ejv) = CInt(f_gauss_m1(CLng(muerte1_va0), CLng(muerte2_va0))) 'Comun 8
    ciclo_nacimiento_agente_va0(numero_total_de_agentes_ejv) = ciclo_ejv 'Comun 9
    agente_probb_mutacion_tipo_va0(numero_total_de_agentes_ejv) = pm_tipo 'Comun 10
    agente_probb_mutacion_mov_va0(numero_total_de_agentes_ejv) = pm_mov 'Comun 11
    agente_probb_mutacion_pm_va0(numero_total_de_agentes_ejv) = pm_pm 'Comun 12
    For cont = 1 To CTE_8_DIR
        agente_tendencia_rel_mov_va0(cont, numero_total_de_agentes_ejv) = m_rel(cont) 'Comun 13
        agente_tendencia_abs_mov_va0(cont, numero_total_de_agentes_ejv) = m_abs(cont) 'Comun 14
    Next cont
    If busqueda_cadena_binaria_va0 Then
        cadena_binaria_va0(numero_total_de_agentes_ejv) = m_cadena 'Comun 15
    End If
    sexo_va0(numero_total_de_agentes_ejv) = fi_azar1(2) 'Comun 16
    
    
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            hormiga_probabilidad_ganar_hyp(numero_total_de_agentes_ejv) = 50 'hyp 1
        Case CTE_PRI '4
            historia_agente_yo_pri(numero_total_de_agentes_ejv, 1) = 0 'pri 1
            historia_agente_el_pri(numero_total_de_agentes_ejv, 1) = 0 'pri 2
            PA_agente_pri(numero_total_de_agentes_ejv) = 0 'pri 3
        Case CTE_CEL '5
        Case CTE_GAI '6
        Case CTE_EXP '7
            For p = 1 To mapa_pisos_va0
            For f = 1 To mapa_filas_va0
            For c = 1 To mapa_columnas_va0
            mapa_exp(numero_total_de_agentes_ejv, p, f, c) = CTE_MAPA_VACIO
            Next c
            Next f
            Next p
        Case CTE_PEZ '9
        Case CTE_UVA '10
            agente_z_va0(numero_total_de_agentes_ejv) = ag_z 'uva 1
            agente_mas_cercano_uva(numero_total_de_agentes_ejv) = 0 'uva 2
            dist_al_mas_cercano_uva(numero_total_de_agentes_ejv) = 0 'uva 3
        Case CTE_YXY '11
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
   
    
    'Control errores de programacion
    If control_errores_de_programacion_ejv Then
        If mapa_va0(ag_z, ag_y, ag_x) <> CTE_MAPA_VACIO Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Agente creado en una zona no vacía"
        End If
    End If
    mapa_va0(ag_z, ag_y, ag_x) = CTE_MAPA_AGENTE
    'Pintamos el agente
    If ver_agentes_va0 Then
        Select Case num_prg_activo_ejv
            Case CTE_HYP, CTE_CEL, CTE_GAI, CTE_EXP, CTE_PEZ, CTE_UVA, CTE_YXY
                If ver_zoom_va0 = CTE_ZOOM_DETALLE Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, ag_z, ag_y, ag_x, CTE_ESFERA, cct_ejv(CTE_BLANCO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                End If
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, ag_z, ag_y, ag_x, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(tipo), agente_direccion_anterior_va0(numero_total_de_agentes_ejv), ver_zoom_va0, 1
            Case CTE_PRI '4
                If ver_zoom_va0 = CTE_ZOOM_DETALLE Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, ag_z, ag_y, ag_x, CTE_ESFERA, cct_ejv(CTE_BLANCO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                End If
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, ag_z, ag_y, ag_x, CTE_PRISIONERO, cct_ejv(CTE_NEGRO), ccs_ejv(f_SumCirc(ncs_i_ejv, tipo, 0)), agente_direccion_anterior_va0(numero_total_de_agentes_ejv), ver_zoom_va0, 1
            Case Else
                s_error_num_prog num_prg_activo_ejv
        End Select
    End If

End Sub

Sub s_eliminar_agente_va0(ag_borrar As Integer)

    Dim ag_z As Double
    Dim ag_y As Double
    Dim ag_x As Double
    
    Dim cont As Integer
    
    Dim tipo_del_agente_que_borro As Integer
    
    ag_z = agente_z_va0(ag_borrar)
    ag_y = agente_y_va0(ag_borrar)
    ag_x = agente_x_va0(ag_borrar)
    
    'control errores de programacion
    If control_errores_de_programacion_ejv Then
        If ag_borrar <> fi_indice_agente_va0(ag_z, ag_y, ag_x) Or ag_borrar = 0 Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error al eliminar agente"
        End If
        If mapa_va0(ag_z, ag_y, ag_x) <> CTE_MAPA_AGENTE Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ese agente no existe en esa posición"
        End If
    End If
    
    mapa_va0(ag_z, ag_y, ag_x) = CTE_MAPA_VACIO
    
    'Borro el agente
     If ver_agentes_va0 Then
        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, ag_z, ag_y, ag_x, CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
        If ver_zoom_va0 = CTE_ZOOM_DETALLE Then
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, ag_z, ag_y, ag_x, CTE_HORMIMUERTA, cct_ejv(CTE_GRISCLARO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
        End If
     End If

    'Guardo esto, ya que luego lo machaco
    tipo_del_agente_que_borro = agente_tipo_va0(ag_borrar)
    
    'Si la hormiga a eliminar no es la ultima de todas
    'sustituyo sus valores con los de la ultima. Si es la ultima
    'puedo sustituirlos y luego borrarla, que da lo mismo
    agente_z_va0(ag_borrar) = agente_z_va0(numero_total_de_agentes_ejv) 'Comun 0
    agente_y_va0(ag_borrar) = agente_y_va0(numero_total_de_agentes_ejv) 'Comun 1
    agente_x_va0(ag_borrar) = agente_x_va0(numero_total_de_agentes_ejv) 'Comun 2
    agente_tipo_va0(ag_borrar) = agente_tipo_va0(numero_total_de_agentes_ejv) 'Comun 3
    peso_agente_va0(ag_borrar) = peso_agente_va0(numero_total_de_agentes_ejv) 'Comun 4
    agente_direccion_anterior_va0(ag_borrar) = agente_direccion_anterior_va0(numero_total_de_agentes_ejv) 'Comun 5
    agente_accion_anterior_va0(ag_borrar) = agente_accion_anterior_va0(numero_total_de_agentes_ejv) 'Comun 6
    apellidos_agente_va0(ag_borrar) = apellidos_agente_va0(numero_total_de_agentes_ejv) 'Comun 7
    muerte_agente_va0(ag_borrar) = muerte_agente_va0(numero_total_de_agentes_ejv) 'Comun 8
    ciclo_nacimiento_agente_va0(ag_borrar) = ciclo_nacimiento_agente_va0(numero_total_de_agentes_ejv) 'Comun 9
    agente_probb_mutacion_tipo_va0(ag_borrar) = agente_probb_mutacion_tipo_va0(numero_total_de_agentes_ejv) 'Comun 10
    agente_probb_mutacion_mov_va0(ag_borrar) = agente_probb_mutacion_mov_va0(numero_total_de_agentes_ejv) 'Comun 11
    agente_probb_mutacion_pm_va0(ag_borrar) = agente_probb_mutacion_pm_va0(numero_total_de_agentes_ejv) 'Comun 12
    For cont = 1 To CTE_8_DIR
        agente_tendencia_rel_mov_va0(cont, ag_borrar) = agente_tendencia_rel_mov_va0(cont, numero_total_de_agentes_ejv) 'Comun 13
        agente_tendencia_abs_mov_va0(cont, ag_borrar) = agente_tendencia_abs_mov_va0(cont, numero_total_de_agentes_ejv) 'Comun 14
    Next cont
    If busqueda_cadena_binaria_va0 Then
        cadena_binaria_va0(ag_borrar) = cadena_binaria_va0(numero_total_de_agentes_ejv) 'Comun 15
    End If
    sexo_va0(ag_borrar) = sexo_va0(numero_total_de_agentes_ejv) 'Comun 16
    
    
    
    Select Case num_prg_activo_ejv
        Case CTE_HYP
            hormiga_probabilidad_ganar_hyp(ag_borrar) = hormiga_probabilidad_ganar_hyp(numero_total_de_agentes_ejv) 'hyp 1
        Case CTE_PRI
            For cont = 1 To num_part_pri
                historia_agente_yo_pri(ag_borrar, cont) = historia_agente_yo_pri(numero_total_de_agentes_ejv, cont) 'pri 1
                historia_agente_el_pri(ag_borrar, cont) = historia_agente_el_pri(numero_total_de_agentes_ejv, cont) 'pri 2
            Next cont
            PA_agente_pri(ag_borrar) = PA_agente_pri(numero_total_de_agentes_ejv) 'pri 3
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select

    'Actualizo la cuenta de los agentes de cada tipo
    num_agentes_tipo_va0(tipo_del_agente_que_borro) = num_agentes_tipo_va0(tipo_del_agente_que_borro) - 1
    
    'Actualizo la cuenta de los agentes totales
    numero_total_de_agentes_ejv = numero_total_de_agentes_ejv - 1
    'Redimensiono los arrays con su nuevo tamaño, uno menos
    If numero_total_de_agentes_ejv > 0 Then
        s_redim_preserve_agente_va0
    End If

End Sub
Sub s_redim_preserve_agente_va0()
    'OJO!!!!!!!!! Esta funcion va pareja con s_inicializar_arrays_va0

    ReDim Preserve agente_z_va0(1 To numero_total_de_agentes_ejv) As Double 'Comun 0
    ReDim Preserve agente_y_va0(1 To numero_total_de_agentes_ejv) As Double 'Comun 1
    ReDim Preserve agente_x_va0(1 To numero_total_de_agentes_ejv) As Double 'Comun 2
    ReDim Preserve agente_tipo_va0(1 To numero_total_de_agentes_ejv) As Integer 'Comun 3
    ReDim Preserve peso_agente_va0(1 To numero_total_de_agentes_ejv) As Double 'Comun 4
    ReDim Preserve agente_direccion_anterior_va0(1 To numero_total_de_agentes_ejv) As Integer 'Comun 5
    ReDim Preserve agente_accion_anterior_va0(1 To numero_total_de_agentes_ejv) As Integer 'Comun 6
    ReDim Preserve apellidos_agente_va0(1 To numero_total_de_agentes_ejv) As String 'Comun 7
    ReDim Preserve muerte_agente_va0(1 To numero_total_de_agentes_ejv) As Integer 'Comun 8
    ReDim Preserve ciclo_nacimiento_agente_va0(1 To numero_total_de_agentes_ejv) As Long 'Comun 9
    ReDim Preserve agente_probb_mutacion_tipo_va0(1 To numero_total_de_agentes_ejv) As Double 'Comun 10
    ReDim Preserve agente_probb_mutacion_mov_va0(1 To numero_total_de_agentes_ejv) As Double 'Comun 11
    ReDim Preserve agente_probb_mutacion_pm_va0(1 To numero_total_de_agentes_ejv) As Double 'Comun 12
    ReDim Preserve agente_tendencia_rel_mov_va0(1 To CTE_8_DIR, 1 To numero_total_de_agentes_ejv) As Long 'Comun 13
    ReDim Preserve agente_tendencia_abs_mov_va0(1 To CTE_8_DIR, 1 To numero_total_de_agentes_ejv) As Long 'Comun 14
    ReDim Preserve cadena_binaria_va0(1 To numero_total_de_agentes_ejv) As String 'Comun 15
    ReDim Preserve sexo_va0(1 To numero_total_de_agentes_ejv) As Integer 'Comun 16
    
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            ReDim Preserve hormiga_probabilidad_ganar_hyp(1 To numero_total_de_agentes_ejv) As Double 'hyp 1
        Case CTE_PRI '4
            ReDim Preserve historia_agente_yo_pri(1 To numero_agentes_que_se_deben_crear_inicio_total_va0, 1 To num_part_pri) As Integer 'pri 1
            ReDim Preserve historia_agente_el_pri(1 To numero_agentes_que_se_deben_crear_inicio_total_va0, 1 To num_part_pri) As Integer 'pri 2
            ReDim Preserve PA_agente_pri(1 To numero_agentes_que_se_deben_crear_inicio_total_va0) As Integer 'pri 3
        Case CTE_CEL '5
        Case CTE_GAI '6
        Case CTE_EXP '7
        Case CTE_PEZ '9
        Case CTE_UVA '10
            ReDim Preserve agente_mas_cercano_uva(1 To numero_total_de_agentes_ejv) As Integer 'uva 1
            ReDim Preserve dist_al_mas_cercano_uva(1 To numero_total_de_agentes_ejv) As Double 'uva 2
        Case CTE_YXY '11
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select

End Sub

Sub s_mover_agente_va0(age As Integer, direccion As Integer)

    'X=columnas
    'Y=filas

    'Borro el agente
    If ver_agentes_va0 Then
        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(age), agente_y_va0(age), agente_x_va0(age), CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
    End If
    mapa_va0(agente_z_va0(age), agente_y_va0(age), agente_x_va0(age)) = CTE_MAPA_VACIO
    
    Select Case direccion
        Case CTE_8_N
            agente_x_va0(age) = agente_x_va0(age)
            agente_y_va0(age) = agente_y_va0(age) - 1
        Case CTE_8_NE
            agente_x_va0(age) = agente_x_va0(age) + 1
            agente_y_va0(age) = agente_y_va0(age) - 1
        Case CTE_8_E
            agente_x_va0(age) = agente_x_va0(age) + 1
            agente_y_va0(age) = agente_y_va0(age)
        Case CTE_8_SE
            agente_x_va0(age) = agente_x_va0(age) + 1
            agente_y_va0(age) = agente_y_va0(age) + 1
        Case CTE_8_S
            agente_x_va0(age) = agente_x_va0(age)
            agente_y_va0(age) = agente_y_va0(age) + 1
        Case CTE_8_SO
            agente_x_va0(age) = agente_x_va0(age) - 1
            agente_y_va0(age) = agente_y_va0(age) + 1
        Case CTE_8_O
            agente_x_va0(age) = agente_x_va0(age) - 1
            agente_y_va0(age) = agente_y_va0(age)
        Case CTE_8_NO
            agente_x_va0(age) = agente_x_va0(age) - 1
            agente_y_va0(age) = agente_y_va0(age) - 1
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
    End Select
    If agente_y_va0(age) = 0 Then agente_y_va0(age) = mapa_filas_va0
    If agente_y_va0(age) = mapa_filas_va0 + 1 Then agente_y_va0(age) = 1
    
    If agente_x_va0(age) = 0 Then agente_x_va0(age) = mapa_columnas_va0
    If agente_x_va0(age) = mapa_columnas_va0 + 1 Then agente_x_va0(age) = 1
    
    
    'Control errores de programacion
    If control_errores_de_programacion_ejv Then
        If mapa_va0(agente_z_va0(age), agente_y_va0(age), agente_x_va0(age)) <> CTE_MAPA_VACIO Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Agente creado en una zona no vacía"
        End If
    End If
    mapa_va0(agente_z_va0(age), agente_y_va0(age), agente_x_va0(age)) = CTE_MAPA_AGENTE
    'Pintamos el agente en su nueva posicion
    If ver_agentes_va0 Then
        Select Case num_prg_activo_ejv
            Case CTE_HYP '1
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(age), agente_y_va0(age), agente_x_va0(age), CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(agente_tipo_va0(age)), direccion, ver_zoom_va0, 1
            Case CTE_PAL '2
            Case CTE_3R '3
            Case CTE_PRI '4
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(age), agente_y_va0(age), agente_x_va0(age), CTE_PRISIONERO, cct_ejv(CTE_NEGRO), ccs_ejv(f_SumCirc(ncs_i_ejv, agente_tipo_va0(age), 0)), direccion, ver_zoom_va0, 1
            Case CTE_CEL '5
            Case CTE_GAI '6
            Case CTE_EXP '7
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(age), agente_y_va0(age), agente_x_va0(age), CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(agente_tipo_va0(age)), direccion, ver_zoom_va0, 1
            Case CTE_CAD '8
            Case CTE_PEZ '9
            Case CTE_UVA '10
            Case CTE_YXY '11
            Case Else
                s_error_num_prog num_prg_activo_ejv
        End Select
    End If
    

End Sub
Sub s_accion_mover_va0(mis_vecinos() As Integer)

    Dim cont As Integer
    Dim direccion_elegida As Integer
    Dim es_mejor As Boolean
    Dim esta_vacio As Boolean
    
    Dim old_Z As Double
    Dim old_Y As Double
    Dim old_X As Double
    
    Dim num_direcciones_libres As Integer
    Dim direcciones_libres() As Integer
    
    If num_prg_activo_ejv = CTE_UVA Then
        'elijo una nueva posición
        old_Z = agente_z_va0(agente_actual_ejv)
        old_Y = agente_y_va0(agente_actual_ejv)
        old_X = agente_x_va0(agente_actual_ejv)
        agente_z_va0(agente_actual_ejv) = old_Z + CDbl(fi_azar1(CInt(radio_grande_uva))) / CInt(radio_grande_uva * 2) - fi_azar1(CInt(radio_grande_uva)) / CInt(radio_grande_uva * 2)
        agente_y_va0(agente_actual_ejv) = old_Y + CDbl(fi_azar1(CInt(radio_grande_uva))) / CInt(radio_grande_uva * 2) - fi_azar1(CInt(radio_grande_uva)) / CInt(radio_grande_uva * 2)
        agente_x_va0(agente_actual_ejv) = old_X + CDbl(fi_azar1(CInt(radio_grande_uva))) / CInt(radio_grande_uva * 2) - fi_azar1(CInt(radio_grande_uva)) / CInt(radio_grande_uva * 2)
        cont = 1
        ReDim Ori_old(1 To 3) As Double
        Ori_old(1) = old_Z
        Ori_old(2) = old_Y
        Ori_old(3) = old_X
        'compruebo si es mejor
        es_mejor = f_es_mejor_posicion(agente_actual_ejv, Ori_old())
        'compruebo si esta vacio
        esta_vacio = f_esta_vacio_uva(agente_actual_ejv, agente_z_va0(agente_actual_ejv), agente_y_va0(agente_actual_ejv), agente_x_va0(agente_actual_ejv))
        actualizar_agente_mas_cercano = False
        While (Not esta_vacio Or Not es_mejor) And cont < 10
            agente_z_va0(agente_actual_ejv) = old_Z + CDbl(fi_azar1(CInt(radio_grande_uva))) / CInt(radio_grande_uva * 2) - fi_azar1(CInt(radio_grande_uva)) / CInt(radio_grande_uva * 2)
            agente_y_va0(agente_actual_ejv) = old_Y + CDbl(fi_azar1(CInt(radio_grande_uva))) / CInt(radio_grande_uva * 2) - fi_azar1(CInt(radio_grande_uva)) / CInt(radio_grande_uva * 2)
            agente_x_va0(agente_actual_ejv) = old_X + CDbl(fi_azar1(CInt(radio_grande_uva))) / CInt(radio_grande_uva * 2) - fi_azar1(CInt(radio_grande_uva)) / CInt(radio_grande_uva * 2)
            cont = cont + 1
            es_mejor = f_es_mejor_posicion(agente_actual_ejv, Ori_old())
            esta_vacio = f_esta_vacio_uva(agente_actual_ejv, agente_z_va0(agente_actual_ejv), agente_y_va0(agente_actual_ejv), agente_x_va0(agente_actual_ejv))
        Wend
        'si no he conseguido nada mejor, vuelvo a la posicion original
        If cont = 10 Then
            agente_z_va0(agente_actual_ejv) = old_Z
            agente_y_va0(agente_actual_ejv) = old_Y
            agente_x_va0(agente_actual_ejv) = old_X
        Else
            'ha habido cambios
            'Actualizo el valor de cercano llamando a f_esta_vacio_uva y a la vez hago control de errores
            actualizar_agente_mas_cercano = True
            If Not f_esta_vacio_uva(agente_actual_ejv, agente_z_va0(agente_actual_ejv), agente_y_va0(agente_actual_ejv), agente_x_va0(agente_actual_ejv)) Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End If
            
            'control de errores
            'If control_errores_de_programacion_ejv Then
            '    ReDim Ori_nuevo(1 To 3) As Double
            '    ReDim Des(1 To 3) As Double
            '    Ori_nuevo(1) = agente_z_va0(agente_actual_ejv)
            '    Ori_nuevo(2) = agente_y_va0(agente_actual_ejv)
            '    Ori_nuevo(3) = agente_x_va0(agente_actual_ejv)
            '    Des(1) = agente_z_va0(agente_mas_cercano_uva(agente_actual_ejv))
            '    Des(2) = agente_y_va0(agente_mas_cercano_uva(agente_actual_ejv))
            '    Des(3) = agente_x_va0(agente_mas_cercano_uva(agente_actual_ejv))
            '    If dist2ptos3d(Ori_old(), Des()) < dist2ptos3d(Ori_nuevo(), Des()) Then
            '        s_error_ejv  CON_OPCION_FINALIZAR, "Error"
            '    End If
            'End If
        End If
    Else
    
        'Copio en el array vecinos_vacios() solo los vecinos vacios
        num_direcciones_libres = 0
        For cont = 1 To CTE_8_DIR
            If mis_vecinos(cont) = CTE_VEC_NADA Then
                num_direcciones_libres = num_direcciones_libres + 1
                ReDim Preserve direcciones_libres(1 To num_direcciones_libres) As Integer
                direcciones_libres(num_direcciones_libres) = cont
            End If
        Next cont
        'Se mueve si puede, si no puede (esta rodeada) no se mueve, pero si
        'consume energia
        If num_direcciones_libres > 0 Then
            'Ahora elijo la posicion de estre esas libres
            direccion_elegida = f_elegir_direccion_mover_va0(agente_actual_ejv, num_direcciones_libres, direcciones_libres())
            'A veces a pesar de haber una libre no puedo moveme a ella,
            'ya que tengo asignada una probabilidad 0 para esa direccion
            If direccion_elegida > 0 Then
                'muevo a esa posicion
                s_mover_agente_va0 agente_actual_ejv, direccion_elegida
                'solo tengo que actualizar el valor si me muevo
                agente_direccion_anterior_va0(agente_actual_ejv) = direccion_elegida
            End If
        End If
        'En cualquier caso siempre gasto energia
        peso_agente_va0(agente_actual_ejv) = peso_agente_va0(agente_actual_ejv) - energia_consumida_al_mover_va0
    
    End If

End Sub

Function f_elegir_direccion_mover_va0(num_agente As Integer, num_direcciones_libres As Integer, direcciones_libres() As Integer) As Integer

    'La probabilidad de ir en una o otra direccion es distinta
    'El chorro de numeros de tendencias es relativo a la direccion actual
    'es decir, la vieja antes de moverse
    Dim cont As Integer
    Dim suma_tendencias As Double
    Dim probabilidad_direccion() As Long
    Dim suma_probabilidad_direccion As Long
    Dim tipo As Integer
    Dim ultima_direccion As Integer
    ReDim tendencia_absoluta(1 To CTE_8_DIR) As Long
    Dim indice As Integer
    Dim tmp As Double
    Dim i_tmp As Integer
    Dim direccion_privilegiada As Integer
    
    
        
    tipo = agente_tipo_va0(num_agente)
    ultima_direccion = agente_direccion_anterior_va0(num_agente)
    
  
    'control de errrores
    For cont = 1 To CTE_8_DIR
        tendencia_absoluta(cont) = -1
    Next cont
  
    
    'Todas las direcciones son absolutas excepto el array de tendencias para cada tipo de agente
    'Tengo la tendencias de cada direccion, pero se trata de direcciones relativas a
    'la posicion actual. Calculo el array de tendencias de cada direccion absolutas,
    'en funcion de cual ha sido la direccion anterior
    'Como las direcciones tienen un orden circular, tengo que sumar en forma circular
    'la direccion vieja. Es decir, en el array de tendencias para cada tipo de agente
    'donde pone norte, en realidad no se refiere al norte, sino a la ultima direccion
    'usada, osea, a un norte relativo
    'Por ejemplo
    'Si estaba mirando al sureste, entonces la absoluta del sureste es igual a la relativa del norte
    For cont = 1 To CTE_8_DIR
        'Si parece un lio esto, mirar en la funcion f_SumCirc que lo explica
        i_tmp = f_SumCirc(8, cont - 1, ultima_direccion)
        tendencia_absoluta(i_tmp) = agente_tendencia_rel_mov_va0(cont, num_agente) + agente_tendencia_abs_mov_va0(i_tmp, num_agente)
    Next cont
    
    'control de errrores
    For cont = 1 To CTE_8_DIR
        If tendencia_absoluta(cont) < 0 Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: tendencia errónea"
        End If
    Next cont
    
    
    'Calculo la suma de las tendencias absolutas
    'pero solo de las libres
    suma_tendencias = 0
    For cont = 1 To num_direcciones_libres
        suma_tendencias = suma_tendencias + CDbl(tendencia_absoluta(direcciones_libres(cont)))
    Next cont
    'si no tengo tendencia hacia ninguna libre, no me muevo
    If suma_tendencias = 0 Then
        f_elegir_direccion_mover_va0 = -1
        Exit Function
    End If
    
    
    
    'Paso a un array las libres, las otras ya no importan
    'ademas tengo que hacerlo asi porque suma_tendencias es la suma
    'pero solo de las libres
    ReDim probabilidad_direccion(1 To num_direcciones_libres) As Long
    For cont = 1 To num_direcciones_libres
        i_tmp = direcciones_libres(cont)
        probabilidad_direccion(cont) = tendencia_absoluta(i_tmp)
        suma_probabilidad_direccion = suma_probabilidad_direccion + probabilidad_direccion(cont)
    Next cont
    
    'For cont = 1 To num_direcciones_libres
    '    i_tmp = direcciones_libres(cont)
    '    If tendencia_absoluta(i_tmp) = 0 Then
    '        tmp = 0
    '    Else
    '        tmp = (CDbl(tendencia_absoluta(i_tmp)) * 100) / CDbl(suma_tendencias)
    '        probabilidad_direccion(cont) = redondear_d(tmp)
    '        'Si sale cero pongo 1, para que al menos siempre salga algo de probabilidad
    '        'excepto que se haya indicado expresamente el cero
    '        If probabilidad_direccion(cont) < 1 Then
    '            probabilidad_direccion(cont) = 1
    '        End If
    '    End If
    '    suma_probabilidad_direccion = suma_probabilidad_direccion + probabilidad_direccion(cont)
    'Next cont
    'Ajusto uno cualquiera para que la suma sea 100
    'Por tanto esa direccion siempre tiene mayor probabilidad
    'pero como cada vez es uno distinto, supongo que no importa
    '...aunque a veces suma_probabilidad_direccion es mayor que 100 y
    'entonces se penaliza esa direccion
    
    'direccion_privilegiada = fi_azar1(num_direcciones_libres)
    'If suma_probabilidad_direccion <> 100 Then
    '    probabilidad_direccion(direccion_privilegiada) = probabilidad_direccion(direccion_privilegiada) + 100 - suma_probabilidad_direccion
    'End If
    
    
    indice = fl_AzarRangos(num_direcciones_libres, probabilidad_direccion(), suma_probabilidad_direccion)
    f_elegir_direccion_mover_va0 = direcciones_libres(indice)
    


End Function


Sub s_inicializar_arrays_va0()
    'OJO!!!!!!!!! Esta funcion va pareja con s_redim_preserve_agente_va0
 
 
    Dim i As Integer
    
    Dim p As Double
    Dim f As Double
    Dim c As Double
        
    Screen.MousePointer = CTE_ARENA
        
    'Inicializamos los tipos de Agentes
    ReDim num_agentes_tipo_va0(1 To num_tipos_agentes_va0) As Integer
    If num_prg_activo_ejv = CTE_PRI Then
        ReDim cont_mensajes_pri(1 To numero_total_de_agentes_ejv) As Long
    End If
    
    'Inicializamos Agentes Individuales: comunes
    ReDim agente_z_va0(1 To 1) As Double 'Comun 0
    ReDim agente_y_va0(1 To 1) As Double 'Comun 1
    ReDim agente_x_va0(1 To 1) As Double 'Comun 2
    ReDim agente_tipo_va0(1 To 1) As Integer 'Comun 3
    ReDim peso_agente_va0(1 To 1) As Double 'Comun 4
    ReDim agente_direccion_anterior_va0(1 To 1) As Integer 'Comun 5
    ReDim agente_accion_anterior_va0(1 To 1) As Integer 'Comun 6
    ReDim apellidos_agente_va0(1 To 1) As String 'Comun 7
    ReDim muerte_agente_va0(1 To 1) As Integer 'Comun 8
    ReDim ciclo_nacimiento_agente_va0(1 To 1) As Long 'Comun 9
    ReDim agente_probb_mutacion_tipo_va0(1 To 1) As Double 'Comun 10
    ReDim agente_probb_mutacion_mov_va0(1 To 1) As Double 'Comun 11
    ReDim agente_probb_mutacion_pm_va0(1 To 1) As Double 'Comun 12
    ReDim agente_tendencia_rel_mov_va0(1 To CTE_8_DIR, 1 To 1) As Long 'Comun 13
    ReDim agente_tendencia_abs_mov_va0(1 To CTE_8_DIR, 1 To 1) As Long 'Comun 14
    ReDim cadena_binaria_va0(1 To 1) As String 'Comun 15
    ReDim sexo_va0(1 To 1) As Integer 'Comun 16


    'Calculo el numero total que debo crear
    numero_agentes_que_se_deben_crear_inicio_total_va0 = 0
    For i = 1 To num_tipos_agentes_va0
        numero_agentes_que_se_deben_crear_inicio_total_va0 = numero_agentes_que_se_deben_crear_inicio_total_va0 + numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(i)
    Next i

    'Inicializamos Agentes Individuales: específicos
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            ReDim hormiga_probabilidad_ganar_hyp(1 To 1) As Double 'hyp 1
            ReDim planta_z(1 To 1) As Double
            ReDim planta_y(1 To 1) As Double
            ReDim planta_x(1 To 1) As Double
            ReDim planta_agua(1 To 1) As Integer
        Case CTE_PRI '4
            'cuando es mas de una dimension, solo se puede modificar la ultima, asi que todos los redim deben tener el valor maximo de las primeras
            ReDim historia_agente_yo_pri(1 To numero_agentes_que_se_deben_crear_inicio_total_va0, 1 To 1) As Integer 'pri 1
            ReDim historia_agente_el_pri(1 To numero_agentes_que_se_deben_crear_inicio_total_va0, 1 To 1) As Integer 'pri 2
            ReDim PA_agente_pri(1 To 1) As Integer 'pri 3
        Case CTE_CEL '5
        Case CTE_GAI '6
        Case CTE_EXP '7
            ReDim mapa_exp(1 To numero_agentes_que_se_deben_crear_inicio_total_va0, 1 To mapa_pisos_va0, 1 To mapa_filas_va0, 1 To mapa_columnas_va0) As Integer 'exp l
        Case CTE_PEZ '9
        Case CTE_UVA '10
            ReDim agente_mas_cercano_uva(1 To 1) As Integer 'uva 1
            ReDim dist_al_mas_cercano_uva(1 To 1) As Double 'uva 2
        Case CTE_YXY '11
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
    
    Screen.MousePointer = CTE_DEFECTO


End Sub


Sub s_crear_agentes_iniciales_va0()

    Dim cont As Integer
    
    Dim p As Double 'pisos    Z
    Dim f As Double 'filas    Y
    Dim c As Double 'columnas X
    
    Dim indice As Integer
    Dim Prueba As Integer
    Dim cont_tipos As Integer
    Dim exito_secuencial As Boolean
    Dim Salir As Boolean
    Dim suma_agentes_cada_tipo As Integer
    Dim cadena As String
        
    ReDim m_rel(1 To CTE_8_DIR) As Long
    ReDim m_abs(1 To CTE_8_DIR) As Long
        
    For cont_tipos = 1 To num_tipos_agentes_va0
        num_agentes_tipo_va0(cont_tipos) = 0
    Next cont_tipos
    numero_total_de_agentes_ejv = 0
    
    
    'Creamos los agentes
    Salir = False
    For cont_tipos = 1 To num_tipos_agentes_va0
        For cont = 1 To CTE_8_DIR
            m_rel(cont) = tendencia_rel_inicial_mov_tipo_agente_va0(cont, cont_tipos)
            m_abs(cont) = tendencia_abs_inicial_mov_tipo_agente_va0(cont, cont_tipos)
        Next cont
        For indice = 1 To numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(cont_tipos)
            DoEvents
            p = fi_azar1(CInt(mapa_pisos_va0))
            f = fi_azar1(CInt(mapa_filas_va0))
            c = fi_azar1(CInt(mapa_columnas_va0))
            If f_esta_vacio_va0(p, f, c) Then
                cont_apellidos_usados_va0 = f_SumCirc(CTE_numero_maximo_apellidos, cont_apellidos_usados_va0, 1)
                cadena = f_generar_cadena_al_azar_ce0(long_cadena_buscada_va0, alfabeto_binario_ejv)
                s_crear_un_agente_va0 p, f, c, cont_tipos, energia_inicial_agente_va0, apellidos_posibles(cont_apellidos_usados_va0) & "_" & numero_total_de_agentes_ejv + 1, probb_mutacion_tipo_inicial_va0, probb_mutacion_mov_inicial_va0, probb_mutacion_pm_inicial_va0, m_abs(), m_rel(), cadena
            Else
                'La celda está ocupada, probamos a ponerla en otro lugar
                Prueba = 0
                While Prueba < 10
                    p = fi_azar1(CInt(mapa_pisos_va0))
                    f = fi_azar1(CInt(mapa_filas_va0))
                    c = fi_azar1(CInt(mapa_columnas_va0))
                    If f_esta_vacio_va0(p, f, c) Then
                        cont_apellidos_usados_va0 = f_SumCirc(CTE_numero_maximo_apellidos, cont_apellidos_usados_va0, 1)
                        cadena = f_generar_cadena_al_azar_ce0(long_cadena_buscada_va0, alfabeto_binario_ejv)
                        s_crear_un_agente_va0 p, f, c, cont_tipos, energia_inicial_agente_va0, apellidos_posibles(cont_apellidos_usados_va0) & "_" & numero_total_de_agentes_ejv + 1, probb_mutacion_tipo_inicial_va0, probb_mutacion_mov_inicial_va0, probb_mutacion_pm_inicial_va0, m_abs(), m_rel(), cadena
                        Prueba = 11
                    Else
                        Prueba = Prueba + 1
                    End If
                Wend
                If Prueba = 10 Then
                    'No ha encontrado ninguna vacia en 10 pruebas
                    'La buscamos secuencialmente
                    exito_secuencial = False
                    For f = 1 To mapa_filas_va0
                    If Salir = True Then Exit For
                    For c = 1 To mapa_columnas_va0
                        If Salir = True Then Exit For
                        If f_esta_vacio_va0(p, f, c) Then
                            cont_apellidos_usados_va0 = f_SumCirc(CTE_numero_maximo_apellidos, cont_apellidos_usados_va0, 1)
                            cadena = f_generar_cadena_al_azar_ce0(long_cadena_buscada_va0, alfabeto_binario_ejv)
                            s_crear_un_agente_va0 p, f, c, cont_tipos, energia_inicial_agente_va0, apellidos_posibles(cont_apellidos_usados_va0) & "_" & numero_total_de_agentes_ejv + 1, probb_mutacion_tipo_inicial_va0, probb_mutacion_mov_inicial_va0, probb_mutacion_pm_inicial_va0, m_abs(), m_rel(), cadena
                            exito_secuencial = True
                            Exit For
                         End If
                    Next c
                    If Salir Then Exit For
                    If exito_secuencial Then Exit For
                    Next f
                    If Not exito_secuencial And Salir = False Then
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: No es posible crear tantos agentes"
                        Exit For
                    End If
                End If
            End If
        Next indice
    Next cont_tipos
    
    'control errores de programacion
    If control_errores_de_programacion_ejv Then
        'Cada tipo
        For indice = 1 To num_tipos_agentes_va0
            If num_agentes_tipo_va0(indice) <> numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(indice) Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error en la creación de agentes iniciales"
                hay_que_detener_ejv = True
                Exit For
            End If
        Next indice
        'Suma de todos los tipos
        suma_agentes_cada_tipo = 0
        For indice = 1 To num_tipos_agentes_va0
            suma_agentes_cada_tipo = suma_agentes_cada_tipo + num_agentes_tipo_va0(indice)
        Next indice
        If suma_agentes_cada_tipo <> numero_total_de_agentes_ejv Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error en la creación de hormigas"
        End If
    End If


    'En el caso de universo, una vez creados, calculo el mas cercano de cada uno
    If num_prg_activo_ejv = CTE_UVA Then
        actualizar_agente_mas_cercano = True
        For agente_actual_ejv = 1 To numero_total_de_agentes_ejv
            'Actualizo el valor de cercano
            If Not f_esta_vacio_uva(agente_actual_ejv, agente_z_va0(agente_actual_ejv), agente_y_va0(agente_actual_ejv), agente_x_va0(agente_actual_ejv)) Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End If
        Next agente_actual_ejv
    End If

End Sub




Function f_ver_vecinos_va0(ind_ag As Integer, mis_vecinos() As Integer) As String

    'El array mis_vecinos() es un parametro de entrada/salida

    Dim vecino_a_inspeccionar As Integer
    Dim hay_agente As Boolean
    Dim hay_planta As Boolean
    
    Dim vecino_z(1 To CTE_8_DIR) As Double
    Dim vecino_y(1 To CTE_8_DIR) As Double
    Dim vecino_x(1 To CTE_8_DIR) As Double
    
    Dim Z As Integer
    Dim Y As Integer
    Dim X As Integer
    
    Z = agente_z_va0(ind_ag)
    Y = agente_y_va0(ind_ag)
    X = agente_x_va0(ind_ag)
    
    vecino_x(CTE_8_N) = X
    vecino_y(CTE_8_N) = Y - 1

    vecino_x(CTE_8_NE) = X + 1
    vecino_y(CTE_8_NE) = Y - 1

    vecino_x(CTE_8_E) = X + 1
    vecino_y(CTE_8_E) = Y

    vecino_x(CTE_8_SE) = X + 1
    vecino_y(CTE_8_SE) = Y + 1
    
    vecino_x(CTE_8_S) = X
    vecino_y(CTE_8_S) = Y + 1
    
    vecino_x(CTE_8_SO) = X - 1
    vecino_y(CTE_8_SO) = Y + 1
    
    vecino_x(CTE_8_O) = X - 1
    vecino_y(CTE_8_O) = Y

    vecino_x(CTE_8_NO) = X - 1
    vecino_y(CTE_8_NO) = Y - 1

   
    'Para cada uno de los 8 vecinos
    'detecto que es
    hay_agente = False
    hay_planta = False
    For vecino_a_inspeccionar = 1 To CTE_8_DIR
        If f_hay_agente_va0(1, vecino_y(vecino_a_inspeccionar), vecino_x(vecino_a_inspeccionar)) Then
            'Si hay agente
            hay_agente = True
            mis_vecinos(vecino_a_inspeccionar) = CTE_VEC_AGENTE
            'Control de Errores
            If num_prg_activo_ejv = CTE_HYP Then
                If f_hay_planta_hyp(1, vecino_y(vecino_a_inspeccionar), vecino_x(vecino_a_inspeccionar)) Then
                    'Si hay agente, Si hay planta
                    'no es posible que haya agente y planta
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: detecta agente y planta a la vez"
                End If
            End If
            If f_hay_obstaculo_va0(1, vecino_y(vecino_a_inspeccionar), vecino_x(vecino_a_inspeccionar)) Then
                'no es posible que haya agente y obstaculo
                s_error_ejv CON_OPCION_FINALIZAR, "Error: detecta agente y obstaculo a la vez"
            End If
        Else
            'No hay agente
            If f_hay_obstaculo_va0(1, vecino_y(vecino_a_inspeccionar), vecino_x(vecino_a_inspeccionar)) Then
                'No hay agente, Si hay obstaculo
                mis_vecinos(vecino_a_inspeccionar) = CTE_VEC_OBSTACULO
                'Control de Errores
                If num_prg_activo_ejv = CTE_HYP Then
                    If f_hay_planta_hyp(1, vecino_y(vecino_a_inspeccionar), vecino_x(vecino_a_inspeccionar)) Then
                        'no es posible que haya planta y obstaculo
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: detecta planta y obstaculo a la vez"
                    End If
                End If
            Else
                'No hay agente, No hay obstaculo
                If num_prg_activo_ejv = CTE_HYP Then
                    If f_hay_planta_hyp(1, vecino_y(vecino_a_inspeccionar), vecino_x(vecino_a_inspeccionar)) Then
                        'Correcto
                        hay_planta = True
                        mis_vecinos(vecino_a_inspeccionar) = CTE_VEC_PLANTA
                    Else
                        'No hay nada
                        mis_vecinos(vecino_a_inspeccionar) = CTE_VEC_NADA
                    End If
                Else
                    'No hay nada
                    mis_vecinos(vecino_a_inspeccionar) = CTE_VEC_NADA
                End If
            End If
        End If
    Next vecino_a_inspeccionar
    
    'Detecto si entre los 8 vecinos hay al menos una hormiga o una planta
    If hay_agente = False Then
        If hay_planta = False Then
            f_ver_vecinos_va0 = CTE_VEC_NADA
        Else
            f_ver_vecinos_va0 = CTE_VEC_PLANTA
        End If
    Else
        If hay_planta = False Then
            f_ver_vecinos_va0 = CTE_VEC_AGENTE
        Else
            f_ver_vecinos_va0 = CTE_VEC_AGENTEYPLANTA
        End If
    End If

End Function

Sub s_comenzar_va0()

    Dim ha_habido_error As Boolean
    Dim exito_al_abrir As Boolean
    
    Dim c1 As Integer
    Dim c2 As Long
    
    ciclo_ejv = 0
    es_la_primera_vez_ejv = True
    
    s_grabar_dato_fichero_salida_ejv CTE_FIC_25_1EJXLS, "Procesando...", 1, 4
   
    s_operacion_ver_ejv CTE_VER_ESTADO_EJECUCION
   
    s_borrar_tiempo_comienzo
    
    hay_que_detener_ejv = False
    esta_detenido_ejv = False
    
    cont_apellidos_usados_va0 = 0

    
    s_botones_enabled_va0 (False)
    
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            's_grabar_opciones_hyp
            s_grabar_tipos_hyp
        Case CTE_PRI '4
            s_cargar_tipo_zoom_va0
            s_fijar_separacion_mapa_va0
            copia_dim_va0_2_viejo_va0
        Case CTE_CEL '5
            s_cargar_tipo_zoom_va0
            s_fijar_separacion_mapa_va0
            copia_dim_va0_2_viejo_va0
        Case CTE_GAI '6
            s_cargar_tipo_zoom_va0
            s_fijar_separacion_mapa_va0
            copia_dim_va0_2_viejo_va0
        Case CTE_EXP '7
            s_cargar_tipo_zoom_va0
            s_fijar_separacion_mapa_va0
            copia_dim_va0_2_viejo_va0
        Case CTE_PEZ '9
            s_cargar_tipo_zoom_va0
            s_fijar_separacion_mapa_va0
            copia_dim_va0_2_viejo_va0
        Case CTE_UVA '10
            s_cargar_tipo_zoom_va0
            s_fijar_separacion_mapa_va0
            copia_dim_va0_2_viejo_va0
        Case CTE_YXY '11
            s_cargar_tipo_zoom_va0
            s_fijar_separacion_mapa_va0
            copia_dim_va0_2_viejo_va0
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
    
    
    'Control de errores de usuario
    If mapa_columnas_va0 < 3 Or mapa_filas_va0 < 3 Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error: el mundo debe ser de un tamaño de al menos 3x3 celdas"
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPOS_AGENTES, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES3, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_MAPA, True
    Else
        s_bucle_general_va0
    End If

End Sub

Sub s_cargar_apellidos_va0()



c_d_ape = 1


apellidos_posibles(c_d_ape) = "Adams"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Armendariz"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Arriaga"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Asenjo"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Asimov"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Avila"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Bilbao"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Cases"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Colorado"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Corpas"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Cuñado"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Dawkins"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Dominguez"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Drescher"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Estevas"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Fernández"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "García"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Gascón"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Gould"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Heinlein"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Heitkötter"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Hernández"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Herrán"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Hofstader"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Holland"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Hoyle"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Isasi"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Jáuregui"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Juste"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Kropotkin"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Langton"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Larrañaga"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Larrea"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "López"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Lotina"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Lozano"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Luengo"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Matorras"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Molina"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Morgan"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Olmeda"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Paniagua"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Pozas"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Prata"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Prieto"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Redfield"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Revenga"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Rossignoli"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Sandín"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Simón"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Udaondo"
c_d_ape = c_d_ape + 1
apellidos_posibles(c_d_ape) = "Watson"
c_d_ape = c_d_ape + 1

End Sub
Sub s_mostrar_apellidos_va0()

    Dim i As Integer
    Dim txt As String
    Dim pal As String
    Dim ultimo As Integer
    
    Screen.MousePointer = CTE_ARENA
    s_centrar_ventana_ejv frm_z0_lista
    frm_z0_lista.Show CTE_AMODAL
    
    frm_z0_lista.Caption = "Lista de Apellidos Posibles"
    ultimo = CTE_numero_maximo_apellidos
    txt = ""

    For i = 1 To ultimo
        pal = ""
        pal = pal & f_espacios_izquierda(CStr(i), 5)
        pal = pal & "  "
        pal = pal & apellidos_posibles(i)
        txt = txt & pal & vbCrLf
    Next i
    
    If Len(txt) > MAX_LISTA Then
        txt = Left(txt, MAX_LISTA)
    End If
    frm_z0_lista.txt_lista.Text = txt
        
    Screen.MousePointer = CTE_DEFECTO

End Sub

Sub s_mostrar_agentes_vivos_va0()

    Dim i As Integer
    Dim j As Integer
    Dim txt As String
    Dim pal As String
    Dim ultimo As Integer
    ReDim tmp(1 To CTE_8_DIR) As Long
       
    If ciclo_ejv = 0 Then Exit Sub
    
    Screen.MousePointer = CTE_ARENA
    s_centrar_ventana_ejv frm_z0_lista
    frm_z0_lista.Show CTE_AMODAL
    
    frm_z0_lista.Caption = "Lista de Agentes Vivos"
    ultimo = numero_total_de_agentes_ejv
    txt = "NumAg  Nacim      Peso Tipo Muerte     P.Ganar     Movimiento Relativo      Movimiento Absoluto   PM.Tipo   PM.Mov.     PM.PM  Cadena Binaria                        Apellidos" & vbCrLf
    txt = txt & "========================================================================================================================================================================" & "======" & vbCrLf
    
    For i = 1 To ultimo
        pal = ""
        pal = pal & f_espacios_izquierda(CStr(i), 5)
        pal = pal & f_espacios_izquierda(ciclo_nacimiento_agente_va0(i), 7)
        pal = pal & f_espacios_izquierda(CStr(CInt(peso_agente_va0(i))), 10)
        pal = pal & f_espacios_izquierda(agente_tipo_va0(i), 5)
        pal = pal & f_espacios_izquierda(muerte_agente_va0(i), 7)
        pal = pal & f_espacios_izquierda_td(hormiga_probabilidad_ganar_hyp(i), 9, 11)
        
        For j = 1 To CTE_8_DIR
            tmp(j) = agente_tendencia_rel_mov_va0(j, i)
        Next j
        pal = pal & f_espacios_izquierda(f_array_l_a_listacomas(tmp()), 25)
        For j = 1 To CTE_8_DIR
            tmp(j) = agente_tendencia_abs_mov_va0(j, i)
        Next j
        pal = pal & f_espacios_izquierda(f_array_l_a_listacomas(tmp()), 25)
        
        pal = pal & f_espacios_izquierda_td(agente_probb_mutacion_tipo_va0(i), 8, 10)
        pal = pal & f_espacios_izquierda_td(agente_probb_mutacion_mov_va0(i), 8, 10)
        pal = pal & f_espacios_izquierda_td(agente_probb_mutacion_pm_va0(i), 8, 10)
        pal = pal & "  "
        pal = pal & f_espacios_izquierda_td(cadena_binaria_va0(i), 8, 36)
        pal = pal & "  "
        pal = pal & apellidos_agente_va0(i)
        txt = txt & pal & vbCrLf
    Next i
    
    If Len(txt) > MAX_LISTA Then
        txt = Left(txt, MAX_LISTA)
    End If
    frm_z0_lista.txt_lista.Text = txt
        
    Screen.MousePointer = CTE_DEFECTO

End Sub

Function f_combinar_apellidos_va0(padre As Integer, madre As Integer) As String

    Dim result As String
    Dim apellidos_p As String
    Dim apellidos_m As String
    Dim primer_espacio As Integer
        
    result = ""
    apellidos_p = apellidos_agente_va0(padre)
    apellidos_m = apellidos_agente_va0(madre)
    
    While Len(apellidos_p) > 0 And Len(apellidos_m) > 0
        'Cojo del Padre
        If Len(apellidos_p) > 0 Then
            primer_espacio = InStr(apellidos_p, " ")
            If primer_espacio = 0 Then
                result = result & apellidos_p & " "
                apellidos_p = ""
            Else
                result = result & Left(apellidos_p, primer_espacio)
                apellidos_p = Right(apellidos_p, Len(apellidos_p) - primer_espacio)
            End If
        End If
        'Cojo de la Madre
        If Len(apellidos_m) > 0 Then
            primer_espacio = InStr(apellidos_m, " ")
            If primer_espacio = 0 Then
                result = result & apellidos_m & " "
                apellidos_m = ""
            Else
                result = result & Left(apellidos_m, primer_espacio)
                apellidos_m = Right(apellidos_m, Len(apellidos_m) - primer_espacio)
            End If
        End If
    Wend
    
    If Len(result) > CTE_LONG_MAX_APELLIDOS Then result = Left(result, CTE_LONG_MAX_APELLIDOS)
    
    f_combinar_apellidos_va0 = Trim(result)
    
End Function

Function f_buscar_lugar_nacimiento_cerca_va0(progenitor_y As Double, progenitor_x As Double, hijo_y As Double, hijo_x As Double)
'El resultado de la funcion lo devuelvo en los mismos parametros i j


    'El hijo debe nacer pegado a un progenitor obligatoriamente, asi que
    'si no hay huecos al lado de alguno de los dos, no nace

    Dim se_ha_encontrado As Boolean
    Dim cont As Integer
    Dim mover As Integer
    Dim direccion_elegida As Integer
    Dim direccion(1 To CTE_8_DIR) As Long

    se_ha_encontrado = False

    'Array ordenado
    For cont = 1 To CTE_8_DIR
        direccion(cont) = cont
    Next cont

    f_desordenar_array_l direccion()
    'Ahora el array está desordenado
    
    For cont = 1 To CTE_8_DIR
        direccion_elegida = direccion(cont)
        Select Case direccion_elegida
            Case CTE_8_N
                hijo_x = progenitor_x
                hijo_y = progenitor_y - 1
            Case CTE_8_NE
                hijo_x = progenitor_x + 1
                hijo_y = progenitor_y - 1
            Case CTE_8_E
                hijo_x = progenitor_x + 1
                hijo_y = progenitor_y
            Case CTE_8_SE
                hijo_x = progenitor_x + 1
                hijo_y = progenitor_y + 1
            Case CTE_8_S
                hijo_x = progenitor_x
                hijo_y = progenitor_y + 1
            Case CTE_8_SO
                hijo_x = progenitor_x - 1
                hijo_y = progenitor_y + 1
            Case CTE_8_O
                hijo_x = progenitor_x - 1
                hijo_y = progenitor_y
            Case CTE_8_NO
                hijo_x = progenitor_x - 1
                hijo_y = progenitor_y - 1
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
        End Select
        If hijo_y = 0 Then hijo_y = mapa_filas_va0
        If hijo_y = mapa_filas_va0 + 1 Then hijo_y = 1
        
        If hijo_x = 0 Then hijo_x = mapa_columnas_va0
        If hijo_x = mapa_columnas_va0 + 1 Then hijo_x = 1
        'Compruebo si está libre
        If f_esta_vacio_va0(1, hijo_y, hijo_x) Then
            'Este es el lugar de nacimiento
            se_ha_encontrado = True
            Exit For
        End If
    Next cont

    f_buscar_lugar_nacimiento_cerca_va0 = se_ha_encontrado




End Function

Sub s_alejar_agente_va0(age As Integer, numero_de_celdas As Integer)


    Dim se_ha_encontrado As Boolean
    Dim estoy_rodeado As Boolean
   
    Dim vieja_z As Double
    Dim vieja_y As Double
    Dim vieja_x As Double
    
    Dim nueva_z As Double
    Dim nueva_x As Double
    Dim nueva_y As Double
    
    Dim direccion As Integer
    Dim cont As Integer
   
    vieja_z = 1
    nueva_z = 1
   
    'CTE_ESTE = 1
    'CTE_NORTE = 2
    'CTE_OESTE = 3
    'CTE_SUR = 4
    
    'Establezco la direccion inicial
    direccion = agente_direccion_anterior_va0(age)
    
    'Establezco el nodo inicial
    vieja_x = agente_x_va0(age)
    vieja_y = agente_y_va0(age)
    
    nueva_x = vieja_x
    nueva_y = vieja_y
    
    'Inicializo los nodos visitados
    'como hay otros obstaculos moviles (otras hormigas)
    'entonces antes de cada accion se inicializan los nodos visitados
    s_inicializar_nodos_visitados_va0 mapa_pisos_va0, mapa_filas_va0, mapa_columnas_va0
    
    
    'El nodo actual no lo tengo en cuenta
    'como nodo a visitar, y empiezo buscando uno
    'que no sea el actual
    
    For cont = 1 To numero_de_celdas
        Select Case algoritmo_busqueda_va0
            Case 1
                'Algoritmo 1
                s_error_ejv CON_OPCION_FINALIZAR, "Error: las hormigas tienen 8 direcciones y este algoritmo es de 4"
                estoy_rodeado = f_alg1_calcular_siguiente_nodo_va0(nueva_z, nueva_y, nueva_x, direccion, mapa_pisos_va0, mapa_filas_va0, mapa_columnas_va0)
            Case 2
                'Algoritmo 2
                s_error_ejv CON_OPCION_FINALIZAR, "Error: las hormigas tienen 8 direcciones y este algoritmo es de 4"
                estoy_rodeado = f_alg2_calcular_siguiente_nodo_va0(nueva_z, nueva_y, nueva_x, direccion, mapa_pisos_va0, mapa_filas_va0, mapa_columnas_va0)
            Case 3
                'Algoritmo 3
                estoy_rodeado = f_alg3_calcular_siguiente_nodo_va0(nueva_z, nueva_y, nueva_x, direccion, mapa_pisos_va0, mapa_filas_va0, mapa_columnas_va0)
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: Algoritmo inexistente"
        End Select
        If estoy_rodeado Then
            'no me puedo alejar
            Exit For
        Else
            'control errores de programacion
            If control_errores_de_programacion_ejv Then
                If Not f_esta_vacio_va0(1, nueva_y, nueva_x) Then
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: no está vacío"
                End If
            End If
            'Muevo el bicho:
            'Borro el agente viejo
            mapa_va0(vieja_z, vieja_y, vieja_x) = CTE_MAPA_VACIO
            If ver_agentes_va0 Then
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, vieja_z, vieja_y, vieja_x, CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
            End If
            'Cambio al agente de posición
            agente_x_va0(age) = nueva_x
            agente_y_va0(age) = nueva_y
            'Control errores de programacion
            If control_errores_de_programacion_ejv Then
                If mapa_va0(nueva_z, nueva_y, nueva_x) <> CTE_MAPA_VACIO Then
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Agente movido a una zona no vacía"
                End If
            End If
            mapa_va0(nueva_z, nueva_y, nueva_x) = CTE_MAPA_AGENTE
            'Pintamos el agente en su nueva posicion
            If ver_agentes_va0 Then
                Select Case num_prg_activo_ejv
                    Case CTE_HYP
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, nueva_z, nueva_y, nueva_x, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(agente_tipo_va0(age)), direccion, ver_zoom_va0, 1
                    Case CTE_PRI
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, nueva_z, nueva_y, nueva_x, CTE_PRISIONERO, cct_ejv(CTE_NEGRO), ccs_ejv(f_SumCirc(ncs_i_ejv, agente_tipo_va0(age), 0)), direccion, ver_zoom_va0, 1
                    Case Else
                        s_error_num_prog num_prg_activo_ejv
                End Select
            End If
            vieja_x = nueva_x
            vieja_y = nueva_y
        End If
    Next cont


End Sub

Sub s_inicializar_nodos_visitados_va0(max_pisos As Double, max_filas As Double, max_col As Double)
     
    'En los nodos visitados no se distingue entre el _va0 y el _map
    'Ejemplo llamadas
    's_inicializar_nodos_visitados_va0 mapa_pisos_va0, mapa_filas_va0, mapa_columnas_va0
    's_inicializar_nodos_visitados_va0 mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
     
    Dim p As Double
    Dim f As Double
    Dim c As Double

    ReDim nodo_visitado_va0(1 To max_pisos, 1 To max_filas, 1 To max_col) As Integer
    
    For p = 1 To max_pisos
    For f = 1 To max_filas
    For c = 1 To max_col
        nodo_visitado_va0(p, f, c) = 0
    Next c
    Next f
    Next p

End Sub

Sub s_ver_opciones_va0()
    
    'aqui habria que actualizar la suma de todos los tipos en hyp
    s_centrar_ventana_ejv frm_a0_opva
    frm_a0_opva.Caption = "Opciones Generales de Vida Artificial para " & nombre_programa_ejv(num_prg_activo_ejv)
    frm_a0_opva.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    
    If ciclo_ejv > 0 Then
        s_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, True
    End If

End Sub

Sub s_ver_tipos_va0()

    Select Case num_prg_activo_ejv
        Case CTE_HYP
            frm_a1_tiposhyp.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
            s_visualizar_tipos_hormigas_hyp
        Case CTE_PRI
            frm_a4_tipospri.Show CTE_AMODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case CTE_EXP
            frm_a7_tiposexp.Show CTE_AMODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
    

End Sub

Sub s_ver_mapa_va0()

    'Copio las opciones del mapa de va0 sobre ma0
    copiar_mapa_a_va0_ma0 = True
    s_copiar_mapa_va0_sobre_ma0_va0
    Unload frm_a0_mapa
    frm_a0_mapa.Show CTE_AMODAL

End Sub

Sub s_ver_refrescar_va0()

    Dim i As Integer
    
    
    'Vuelvo a pintar el borde, el mapa, las plantas y las hormigas
    If numero_total_de_agentes_ejv > 0 Or numero_plantas_hyp > 0 Then
        If ciclo_ejv > 0 Then
            s_mapa_pintar_bordes_va0 frm_a0_va
            s_mostrar_mapa_actual_va0 False
        
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    For i = 1 To numero_plantas_hyp
                        'Pintamos la planta
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, planta_z(i), planta_y(i), planta_x(i), CTE_PLANTA, cct_ejv(CTE_VERDEBRILLANTE), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                    Next i
                    For i = 1 To numero_total_de_agentes_ejv
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(i), agente_y_va0(i), agente_x_va0(i), CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(agente_tipo_va0(i)), agente_direccion_anterior_va0(i), ver_zoom_va0, 1
                    Next i
                Case CTE_PAL '2
                Case CTE_3R '3
                Case CTE_PRI '4
                    For i = 1 To numero_total_de_agentes_ejv
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(i), agente_y_va0(i), agente_x_va0(i), CTE_PRISIONERO, cct_ejv(CTE_NEGRO), ccs_ejv(f_SumCirc(ncs_i_ejv, agente_tipo_va0(i), 0)), agente_direccion_anterior_va0(i), ver_zoom_va0, 1
                    Next i
                Case CTE_CEL '5
                Case CTE_GAI '6
                Case CTE_EXP '7
                    For i = 1 To numero_total_de_agentes_ejv
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(i), agente_y_va0(i), agente_x_va0(i), CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(agente_tipo_va0(i)), agente_direccion_anterior_va0(i), ver_zoom_va0, 1
                    Next i
                Case CTE_CAD '8
                Case CTE_PEZ '9
                Case CTE_UVA '10
                Case CTE_YXY '11
                Case Else
                    s_error_num_prog num_prg_activo_ejv
            End Select
        End If
    End If
    
End Sub


Sub s_cambiar_tendencias_mov_va0(tipo_tendencia As Integer, mi_index As Integer)

    Dim cont As Integer
    
    tipo_tendencia_en_modificacion_va0 = tipo_tendencia
    ReDim lista_tendencias_en_modificacion_va0(1 To CTE_8_DIR) As Long
    tipo_agente_cambiar_tendencias_va0 = mi_index + 1
    
    'Copio en lista_tendencias_va0 la lista de este tipo
    For cont = 1 To CTE_8_DIR
        If tipo_tendencia = CTE_RELATIVAS Then
            lista_tendencias_en_modificacion_va0(cont) = tendencia_rel_inicial_mov_tipo_agente_va0(cont, tipo_agente_cambiar_tendencias_va0)
        Else
            lista_tendencias_en_modificacion_va0(cont) = tendencia_abs_inicial_mov_tipo_agente_va0(cont, tipo_agente_cambiar_tendencias_va0)
        End If
    Next cont
    'Llamo a la pantalla
    ha_habido_cambio_lista_tendencias_va0 = False
    frm_a0_mov.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    'Modifico los cambios
    If ha_habido_cambio_lista_tendencias_va0 Then
    'Modifico solo el texto, el verdadero dato
    'se modifica al dar aceptar a toda la ventana
        If tipo_tendencia = CTE_RELATIVAS Then
            frm_a1_tiposhyp.Tendencias_r(mi_index).Text = f_array_l_a_listacomas(lista_tendencias_en_modificacion_va0())
        Else
            frm_a1_tiposhyp.Tendencias_a(mi_index).Text = f_array_l_a_listacomas(lista_tendencias_en_modificacion_va0())
        End If
    End If
End Sub
Sub s_grabar_opciones_generales_va0()

    '1 Modo de Ejecución
    ver_agentes_va0 = frm_a0_opva.Op_VerAgentes
    '2 Lugar de nacimiento
    nacimiento_cerca_va0 = frm_a0_opva.Op_NacimientoCerca
    '3 tasas de mutacion
    probb_mutacion_tipo_inicial_va0 = CDbl(frm_a0_opva.Op_PMColor.Text)
    probb_mutacion_mov_inicial_va0 = CDbl(frm_a0_opva.Op_PMMov.Text)
    probb_mutacion_pm_inicial_va0 = CDbl(frm_a0_opva.Op_PMPM.Text)
    PMPMCte_va0 = frm_a0_opva.Op_PMPMCte
    '4 Agentes inmortales
    agentes_inmortales_va0 = frm_a0_opva.Op_Inmortales
    muerte1_va0 = CInt(frm_a0_opva.Op_Muerte1.Text)
    muerte2_va0 = CInt(frm_a0_opva.Op_Muerte2.Text)
    '5 Búsqueda de Cadena binaria
    If frm_a0_opva.Op_BusquedaCadena = 1 Then
        busqueda_cadena_binaria_va0 = True
    Else
        busqueda_cadena_binaria_va0 = False
    End If
    cadena_binaria_buscada_va0 = CStr(frm_a0_opva.CadenaBinaria)
    long_cadena_buscada_va0 = Len(cadena_binaria_buscada_va0)
    '6 Limite Muerte
    limite_muerte_va0 = CLng(frm_a0_opva.LimiteMuerte)


End Sub

Sub s_cargar_opciones_generales_va0()
    
    'Check:0,1
    'Option:true,false

    
    '1 Modo de Ejecución
    frm_a0_opva.Op_VerAgentes = ver_agentes_va0
    frm_a0_opva.Op_nVerAgentes = Not ver_agentes_va0
    '2 Lugar de nacimiento
    frm_a0_opva.Op_NacimientoCerca = nacimiento_cerca_va0
    frm_a0_opva.Op_nNacimientoCerca = Not nacimiento_cerca_va0
    '3 tasas de mutacion
    frm_a0_opva.Op_PMColor.Text = CStr(probb_mutacion_tipo_inicial_va0)
    frm_a0_opva.Op_PMMov.Text = CStr(probb_mutacion_mov_inicial_va0)
    frm_a0_opva.Op_PMPM.Text = CStr(probb_mutacion_pm_inicial_va0)
    frm_a0_opva.Op_PMPMCte = PMPMCte_va0
    frm_a0_opva.Op_nPMPMCte = Not PMPMCte_va0
    '4 agentes inmortales
    frm_a0_opva.Op_Inmortales = agentes_inmortales_va0
    frm_a0_opva.Op_nInmortales = Not (agentes_inmortales_va0)
    frm_a0_opva.Op_Muerte1 = muerte1_va0
    frm_a0_opva.Op_Muerte2 = muerte2_va0
    '5 Búsqueda de Cadena binaria
    If busqueda_cadena_binaria_va0 Then
        frm_a0_opva.Op_BusquedaCadena = 1
    Else
        frm_a0_opva.Op_BusquedaCadena = 0
    End If
    frm_a0_opva.CadenaBinaria = cadena_binaria_buscada_va0
    '6 Limite Muerte
    frm_a0_opva.LimiteMuerte = limite_muerte_va0


End Sub


Sub s_mapa_pintar_bordes_va0(formulario As Object)

    Dim p As Double
    Dim f As Double
    Dim c As Double

    p = 1
    'Pintamos bordes
    For f = 0 To mapa_filas_va0 + 1
        If ver_agentes_va0 Then
            s_pintar_objeto_ejv CTE_FORMULARIO, formulario, p, f, 0, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, formulario, p, f, mapa_columnas_va0 + 1, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
        End If
    Next f
    For c = 0 To mapa_columnas_va0 + 1
        If ver_agentes_va0 Then
            s_pintar_objeto_ejv CTE_FORMULARIO, formulario, p, 0, c, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
            s_pintar_objeto_ejv CTE_FORMULARIO, formulario, p, mapa_filas_va0 + 1, c, CTE_CUBO, cct_ejv(CTE_AZUL), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
        End If
    Next c
    
End Sub


Sub s_mostrar_mapa_actual_va0(pintar_todo As Boolean)

    Dim p As Double
    Dim f As Double
    Dim c As Double

    Screen.MousePointer = CTE_ARENA

    'pintar_todo dice si se pintan tambien los huecos de grisclaro
    If Not mapa_sin_obstaculos_ma0 Then
        If UBound(mapa_ma0, 1) > 0 Then
            For p = 1 To mapa_pisos_va0
            For f = 1 To mapa_filas_va0
            For c = 1 To mapa_columnas_va0
                If mapa_va0(p, f, c) = CTE_MAPA_OBSTACULO Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, p, f, c, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                Else
                    If pintar_todo And mapa_va0(p, f, c) = CTE_MAPA_VACIO Then
                        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, p, f, c, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                    End If
                End If
            Next c
            Next f
            Next p
        End If
    End If

    Screen.MousePointer = CTE_DEFECTO


End Sub

