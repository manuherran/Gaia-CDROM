Attribute VB_Name = "bas_aA_uva"
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


'Opciones
Global radio_pequenio_uva As Double
Global radio_grande_uva As Double

'cosas
Global actualizar_agente_mas_cercano As Boolean


Sub s_inicializar_ejemplo_elegido_uva()

    Dim f As Integer
    Dim c As Integer
    Dim p As Integer

    Dim exito_al_abrir As Boolean
    Dim filas As Integer
    Dim columnas As Integer
    
    actualizar_agente_mas_cercano = True
    
    'Carga de los tipos de agentes
    num_tipos_agentes_va0 = 1
    ReDim tendencia_rel_inicial_mov_tipo_agente_va0(CTE_8_DIR, num_tipos_agentes_va0) As Long
    ReDim tendencia_abs_inicial_mov_tipo_agente_va0(CTE_8_DIR, num_tipos_agentes_va0) As Long
    ReDim numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1 To num_tipos_agentes_va0) As Integer
    
    'OPCIONES I
    'GENERALES DE EJEMPLOS DE VIDA(DISTINTAS A LAS DE POR DEFECTO)
    '2 Grabar Resumen
    un_ej_grabar_gra_ejv = True
    un_ej_fichero_gra_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_GRA), "r_uva_" & num_ej_activo_ejv & ".gra")
    un_ej_grabar_resumen_txt_ejv = False
    un_ej_fichero_resumen_txt_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_TXT), "r_uva_" & num_ej_activo_ejv & ".txt")
    un_ej_grabar_resumen_xls_ejv = False
    un_ej_fichero_resumen_xls_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS), "r_uva_" & num_ej_activo_ejv & ".xls")
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
        nombre_fichero_mapa_va0 = "uvas.map"
        'ESPECIFICAS DE UVA
        '1 numero de agentes de cada tipo
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 10
        '2 radios
        radio_pequenio_uva = 8
        radio_grande_uva = 30
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe ese ejemplo"
    End Select

    'Cargo el mapa
    mapa_actual_ma0 = f_nombre_completo(path_largo_ejv(CTE_C_PRG_MAP), nombre_fichero_mapa_va0)
    s_aut_leer_mapa_ma0
    s_copiar_mapa_ma0_sobre_va0_va0
    

End Sub


Sub s_inicializar_uva()

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
    
    's_mapa_pintar_bordes_va0 frm_a0_va
    's_mostrar_mapa_actual_va0 False
    
    se_ha_empezado_a_crear_agentes_va0 = True
    s_crear_agentes_iniciales_va0

End Sub
Function f_es_mejor_posicion(num_agente As Integer, Ori_old() As Double)

    ReDim Ori_nuevo(1 To 3) As Double
    ReDim Des(1 To 3) As Double
    
    Ori_nuevo(1) = agente_z_va0(num_agente)
    Ori_nuevo(2) = agente_y_va0(num_agente)
    Ori_nuevo(3) = agente_x_va0(num_agente)
    Des(1) = agente_z_va0(agente_mas_cercano_uva(num_agente))
    Des(2) = agente_y_va0(agente_mas_cercano_uva(num_agente))
    Des(3) = agente_x_va0(agente_mas_cercano_uva(num_agente))
    If dist2ptos3d(Ori_nuevo(), Des()) < dist2ptos3d(Ori_old(), Des()) Then
        f_es_mejor_posicion = True
    Else
        f_es_mejor_posicion = False
    End If

End Function
Function f_esta_vacio_uva(ag_a_analizar As Integer, Z As Double, Y As Double, X As Double) As Boolean

    Dim cont As Integer
    Dim se_ha_tratado_el_primero As Boolean
    Dim vacio As Boolean
    ReDim Ori(1 To 3) As Double
    ReDim Des(1 To 3) As Double
    Dim dist As Double
    Dim dist_cercano As Double
    Dim cercano As Integer
    
    'Compruebo que que la distancia a cualquiera de las otros centros o agentes
    'no excede de 2R de la esfera grande en ese punto
    
    
    Ori(1) = Z
    Ori(2) = Y
    Ori(3) = X
    vacio = True
    se_ha_tratado_el_primero = False
    For cont = 1 To numero_total_de_agentes_ejv
        If cont <> ag_a_analizar Then
            Des(1) = agente_z_va0(cont)
            Des(2) = agente_y_va0(cont)
            Des(3) = agente_x_va0(cont)
            dist = dist2ptos3d(Ori(), Des())
            If se_ha_tratado_el_primero = False Or dist < dist_al_mas_cercano_uva(agente_actual_ejv) Then
                dist_cercano = dist
                cercano = cont
            End If
            If dist < 2 * radio_grande_uva Then
                vacio = False
                Exit For
            End If
            se_ha_tratado_el_primero = True
        End If
    Next cont
    
    'Si he encontrado lugar, eso es que lo voy a crear o mover ahí
    'asi que guardo uno de los mas cercanos para luego, en el siguiente movimiento
    'intentar acercarme a ese
    If actualizar_agente_mas_cercano Then
        If vacio Then
            agente_mas_cercano_uva(agente_actual_ejv) = cercano
            dist_al_mas_cercano_uva(agente_actual_ejv) = dist_cercano
        End If
    End If
    f_esta_vacio_uva = vacio
    

End Function
Sub s_mostrar_info_uva()
    
    Dim ultimo_agente_mas_cercano As Integer
    Dim cont As Integer
    
    'Pinto todas las uvas
    frm_a0_va.Refresh
    s_pintar_ejes3D CTE_FORMULARIO, frm_a0_va, 2
    For cont = 1 To numero_total_de_agentes_ejv
        s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(cont), agente_y_va0(cont), agente_x_va0(cont), ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio_pequenio_uva, 2
        s_pintar_esfera3D_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(cont), agente_y_va0(cont), agente_x_va0(cont), cct_ejv(CTE_BLANCO), cct_ejv(CTE_TRANSPARENTE), CTE_DIRECC_NINGUNA, radio_grande_uva, 2
        'pinto las lineas a los ejes
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(cont), agente_y_va0(cont), agente_x_va0(cont), 0, agente_y_va0(cont), agente_x_va0(cont), ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, 0, 0, agente_x_va0(cont), 0, agente_y_va0(cont), agente_x_va0(cont), ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, 0, agente_y_va0(cont), 0, 0, agente_y_va0(cont), agente_x_va0(cont), ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(cont), agente_y_va0(cont), agente_x_va0(cont), agente_z_va0(cont), 0, agente_x_va0(cont), ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, 0, 0, agente_x_va0(cont), agente_z_va0(cont), 0, agente_x_va0(cont), ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(cont), 0, 0, agente_z_va0(cont), 0, agente_x_va0(cont), ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(cont), agente_y_va0(cont), agente_x_va0(cont), agente_z_va0(cont), agente_y_va0(cont), 0, ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, 0, agente_y_va0(cont), 0, agente_z_va0(cont), agente_y_va0(cont), 0, ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(cont), 0, 0, agente_z_va0(cont), agente_y_va0(cont), 0, ccs_ejv(f_SumCirc(ncs_i_ejv, cont, 0)), 1
        
        
        ultimo_agente_mas_cercano = agente_mas_cercano_uva(cont)
        If ultimo_agente_mas_cercano > 0 Then
            s_pintar_linea3D_ejv CTE_FORMULARIO, frm_a0_va, agente_z_va0(cont), agente_y_va0(cont), agente_x_va0(cont), agente_z_va0(ultimo_agente_mas_cercano), agente_y_va0(ultimo_agente_mas_cercano), agente_x_va0(ultimo_agente_mas_cercano), cct_ejv(CTE_NEGRO), 1
        End If
    Next cont
    


End Sub

Sub s_grabar_resumen_uva()

    Dim cont As Integer
    Dim i As Integer
    Dim linea As String

    'Los 1 tipos de agentes
    linea = ""
    linea = linea & f_comillas(CStr(ciclo_ejv)) ' el ciclo actual
    For i = 1 To num_tipos_agentes_va0
        linea = linea & ";" & f_comillas(CStr(num_agentes_tipo_va0(i)))
    Next i
    s_grabar_dato_fichero_salida_ejv CTE_FIC_23W_1EJGRA, linea
      

End Sub

Function f_ahora_esta_mas_cercano_uva(old_Z As Double, old_Y As Double, old_X As Double)

    ReDim Ori(1 To 3) As Double
    ReDim Des(1 To 3) As Double
    
    Dim ultimo_agente_mas_cercano As Integer

    Ori(1) = old_Z
    Ori(2) = old_Y
    Ori(3) = old_X
    
    ultimo_agente_mas_cercano = agente_mas_cercano_uva(agente_actual_ejv)
    If ultimo_agente_mas_cercano = 0 Then
        f_ahora_esta_mas_cercano_uva = True
    Else
        'Pruebo si la nueva posicion esta mas cerca que la que tuvo en el anterior ciclo
        Des(1) = agente_z_va0(ultimo_agente_mas_cercano)
        Des(2) = agente_y_va0(ultimo_agente_mas_cercano)
        Des(3) = agente_x_va0(ultimo_agente_mas_cercano)
        
        If dist2ptos3d(Ori(), Des()) > dist_al_mas_cercano_uva(agente_actual_ejv) Then
            f_ahora_esta_mas_cercano_uva = False
        Else
            f_ahora_esta_mas_cercano_uva = True
        End If
    End If

End Function

Sub s_grabar_opciones_uva()

    numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = CInt(frm_aA_opuva.numeroAgentes)
    radio_pequenio_uva = CInt(frm_aA_opuva.radioPequenio)
    radio_grande_uva = CInt(frm_aA_opuva.radioGrande)

End Sub

Sub s_cargar_opciones_uva()

    frm_aA_opuva.numeroAgentes = numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1)
    frm_aA_opuva.radioPequenio = radio_pequenio_uva
    frm_aA_opuva.radioGrande = radio_grande_uva

End Sub

