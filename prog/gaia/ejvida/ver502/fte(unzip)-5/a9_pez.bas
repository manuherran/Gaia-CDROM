Attribute VB_Name = "bas_a9_pez"
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
Sub s_inicializar_ejemplo_elegido_pez()

    Dim i As Integer
    Dim j As Integer
    
    Dim p As Integer
    Dim f As Integer
    Dim c As Integer

    Dim exito_al_abrir As Boolean
    
    
    'Carga de los tipos de agentes
    num_tipos_agentes_va0 = 2
    ReDim tendencia_rel_inicial_mov_tipo_agente_va0(CTE_8_DIR, num_tipos_agentes_va0) As Long
    ReDim tendencia_abs_inicial_mov_tipo_agente_va0(CTE_8_DIR, num_tipos_agentes_va0) As Long
    ReDim numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1 To num_tipos_agentes_va0) As Integer
    
    
    'OPCIONES I
    'GENERALES DE EJEMPLOS DE VIDA(DISTINTAS A LAS DE POR DEFECTO)
    '2 Grabar Resumen
    un_ej_grabar_gra_ejv = True
    un_ej_fichero_gra_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_GRA), "r_pez_" & num_ej_activo_ejv & ".gra")
    un_ej_grabar_resumen_txt_ejv = False
    un_ej_fichero_resumen_txt_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_TXT), "r_pez_" & num_ej_activo_ejv & ".txt")
    un_ej_grabar_resumen_xls_ejv = False
    un_ej_fichero_resumen_xls_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS), "r_pez_" & num_ej_activo_ejv & ".xls")
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
        nombre_fichero_mapa_va0 = "ej01.map"
        'ESPECIFICAS DE EXPLORANDO PECES
        '1 Jugadores se repelen
        repulsion_exp = True
        '2 Compartir mapas
        compartir_mapas_exp = True
        '3 numero de agentes de cada tipo
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 10
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 10
        '4 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 3
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 3
        Next i
        For i = 1 To num_tipos_agentes_va0
            tendencia_abs_inicial_mov_tipo_agente_va0(1, i) = 0
            tendencia_abs_inicial_mov_tipo_agente_va0(2, i) = 1
            tendencia_abs_inicial_mov_tipo_agente_va0(3, i) = 400
            tendencia_abs_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_abs_inicial_mov_tipo_agente_va0(5, i) = 0
            tendencia_abs_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_abs_inicial_mov_tipo_agente_va0(7, i) = 400
            tendencia_abs_inicial_mov_tipo_agente_va0(8, i) = 1
        Next i
        

        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe ese ejemplo"
    End Select

    'Cargo el mapa
    mapa_actual_ma0 = f_nombre_completo(path_largo_ejv(CTE_C_PRG_MAP), nombre_fichero_mapa_va0)
    s_aut_leer_mapa_ma0
    s_copiar_mapa_ma0_sobre_va0_va0
    

End Sub
Sub s_mostrar_info_pez()

End Sub

Sub s_grabar_resumen_pez()

End Sub

Sub s_inicializar_pez()

End Sub
