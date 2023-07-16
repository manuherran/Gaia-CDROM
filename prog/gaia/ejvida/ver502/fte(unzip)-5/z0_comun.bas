Attribute VB_Name = "bas_z0_comun"
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

'Nombre y Version de la Aplicacion
Global nombre_aplicacion_ejv As String
Global version_aplicacion_ejv As String

'Paths
Global path_largo_ejv(1 To CTE_TOTAL_CARPETAS) As String
Global path_corto_ejv(1 To CTE_TOTAL_CARPETAS) As String


'=================================
'Configuracion de todo el programa
'=================================
Global version_ejv As String '1
Global idioma_ejv As String '2
Global elegir_idioma_ejv As Boolean '3
Global control_errores_de_programacion_ejv As Boolean '4
Global mostrar_logo_ejv As Boolean '5
Global algoritmo_ordenacion_ejv As String '6
Global sistema_operativo_ejv As String '7
Global pedir_confirmacion_ejv As Boolean '8
Global resolucion_pantalla_ejv As String '9
Global grabar_configuracion_ejv As Boolean '10
Global grabar_config_defecto_ejv As Boolean '11
'Estos ficheros de salida son unicos para una ejecución y se asignan en
'el fichero de configuracion
Global grabar_log_ejv As Boolean '12
Global fichero_log_ejv As String '13

Global grabar_resumen_txt_ejv As Boolean '14
Global fichero_resumen_txt_ejv As String '15

Global grabar_resumen_xls_ejv As Boolean '16
Global fichero_resumen_xls_ejv As String '17

Global automatico_ejv As Boolean '18
Global fichero_aut_ejv() As String '19

Global finalizacion_usuario_ejv As Boolean
Global cancelar_mostrar_grafico_ejv As Boolean

'Grabar resumen en Excel
Global HojaResumenExcel As Object 'Resumen de la ejecución (1)
Global HojaUnEjResumenExcel As Object 'Resumen de los programas lanzados en la ejecución (N)

Global ContFilasHojaResumenExcel As Long
Global ContFilasHojaUnEjResumenExcel As Long

'Ejemplo actual
Global num_prg_activo_ejv As Integer 'Programa
Global num_prg_anterior_activo_ejv As Integer  'Programa anterior
Global num_ej_activo_ejv As Integer 'Ejemplo
Global num_ficheros_aut_ejv As Long 'Número total de AUT
Global indice_auto As Long 'AUT actual de definicion de más características
Global num_iteraciones_ejv As Long 'Número total de iteraciones
Global indice_iteraciones As Long 'Iteración actual
Global seg_ej_actual_ejv As Long 'Numero de segundos transcurridos en una ejecución


'=====================================================================================
'Opciones Generales de Vida Artificial y Computacion Evolutiva para el programa activo
'=====================================================================================
    '1 Condición de Parada
Global CondParadaNumMaxCiclos_ejv As Boolean
Global CondParadaNumMaxCiclosFinal_ejv As Long
Global CondParadaFechaHora_ejv As Boolean
Global CondParadaFecha_ejv As Date
Global CondParadaHora_ejv As Date
Global CondParadaPeso_ejv As Boolean
Global CondParadaPesoNecesario_ejv As Double
    '2 Grabar Resumen (Ejecución de un ejemplo)
Global un_ej_grabar_gra_ejv As Boolean
Global un_ej_fichero_gra_ejv As String

Global un_ej_grabar_resumen_txt_ejv As Boolean
Global un_ej_fichero_resumen_txt_ejv As String

Global un_ej_grabar_resumen_xls_ejv As Boolean
Global un_ej_fichero_resumen_xls_ejv As String

Global autoguardado_ejv As Long
Global max_guardado_ejv As Long
Global cabeceras_ejv As Boolean
Global reemplazar_fic_ejv As Boolean
    '3 Funcion de Azar
Global tipo_funcion_azar_ejv As Integer
Global randomize_ejv As Boolean
Global azar_carpeta_ejv As String
Global azar_fichero_ejv As String
Global azar_fichero_num_char_ejv As String
    '4 Color de Fondo
Global cfondo_ejv As Integer
Global eliminar_cfondo_ejv As Boolean


'Resumen para estadísticas con los datos del ultimo ciclo en el ultimo instante
Global resumen_actual() As Long

'Ciclos de ejecucion
Global ciclo_ejv As Long

'Agentes (aspectos comunes as todos los programas)
Global agente_actual_ejv As Integer 'agente que se esta evaluando o ordenando o creando o reproduciendo etc
Global numero_total_de_agentes_ejv As Integer 'Numero total de agentes (suma de todos los tipos)
Global agente_contrario_pri As Integer


'Mostrar una imagen al comenzar
Global mostrar_aviso_imagen_ejv As Boolean

Global alfabeto_binario_ejv() As String
'Lista de ejemplos de vida
Global nombre_programa_ejv() As String

'Fichero de Azar
Global digitos_azar() As Integer 'chorro de numeros
Global azar_en_memoria_ejv As Boolean 'dice si efectivamente esta cargada o no la serie de numeros
Global indice_azar As Long 'puntero al valor actual del chorro de numeros

'Gauss
Global hay_solucion_anterior_gauss  As Boolean
Global solucion_anterior_gauss As Double


Global separacion_grafico_gra As Integer

Global ant_CX_gra As Integer
Global ant_CY_gra As Integer

'Colores: "serie" son todos menos el fondo, en el caso de estar activada la opcion
Global nct_i_ejv As Integer 'numero colores todos
Global nct_ejv() As String 'nombre colores todos
Global cct_ejv() As Long 'codigo colores todos
Global ncs_i_ejv As Integer 'numero colores serie
Global ncs_ejv() As String 'nombre colores serie
Global ccs_ejv() As Long 'codigo colores serie

'Selector
Global selector_max_der_sel As Integer
Global modificar_resultado_selector_sel As Boolean
Global resultado_selector_sel() As String

'Posicion graficas
Global horiz_graf_gra() As Integer

'Estado del menu ejecutar
Global estado_ejecutar_ejv() As Boolean
'Estado del menu ver
Global estado_ver_ejv() As Boolean

'Estado de la ejecución:
Global hay_que_detener_ejv As Boolean
Global hay_que_terminar_ejv As Boolean
Global esta_detenido_ejv As Boolean
Global esta_terminado_ejv As Boolean
Global es_la_primera_vez_ejv As Boolean

'Abrir ficheros
Global nombre_fichero_ejv As String
Global cancelar_operacion_fichero_ejv As Boolean
Global tipo_operacion_formulario_fic_ejv As Integer
Global nombre_fichero_ejv_es_solo_un_path_ejv As Boolean
Global lista_ficheros_sin_path_ejv() As String
Global cont_fic_lista_ejv As Integer


Function f_control_cerrar_va0() As Boolean
'1 4 5 7 9
Dim ret As Boolean

ret = False
If esta_terminado_ejv = False Then
    If hay_que_detener_ejv Then
        MsgBox "Se está deteniendo el proceso de Vida Artificial. Espere, por favor.", vbCritical
    Else
        MsgBox "Ha de Terminar(F7) el proceso de Vida Artificial antes de cerrar la ventana, o cerrar la ventana principal si desea terminar ya su sesión con Ejemplos de Vida.", vbCritical
    End If
        ret = True
Else
    If esta_detenido_ejv = False Then
        MsgBox "Se ha de detener el proceso de Vida Artificial antes de cerrar la ventana", vbCritical
        ret = True
    End If
End If


f_control_cerrar_va0 = ret

End Function

Function f_control_cerrar_pal() As Boolean
'2
Dim ret As Boolean

ret = False
If hay_que_detener_ejv = True Then
    MsgBox "Se está deteniendo el proceso de Palabras y Frases. Espere, por favor.", vbCritical
    ret = True
Else
    If esta_detenido_ejv = False Then
        MsgBox "Se ha de terminar (F7) el proceso de Palabras y Frases antes de cerrar esta ventana, o cerrar la ventana principal si desea terminar ya su sesión con Ejemplos de Vida.", vbCritical
        ret = True
    End If
End If


f_control_cerrar_pal = ret


End Function
Function f_control_cerrar_ce0() As Boolean
'3 8
Dim ret As Boolean

ret = False
If esta_terminado_ejv = False Then
    If hay_que_detener_ejv Then
        MsgBox "Se está deteniendo el proceso de Computación Evolutiva. Espere, por favor.", vbCritical
    Else
        MsgBox "Ha de terminar(F7) el proceso de Computación Evolutiva antes de cerrar esta ventana, o cerrar la ventana principal si desea terminar ya su sesión con Ejemplos de Vida.", vbCritical
    End If
        ret = True
Else
    If esta_detenido_ejv = False Then
        MsgBox "Se ha de detener el proceso de Computación Evolutiva antes de cerrar la ventana", vbCritical
        ret = True
    End If
End If

f_control_cerrar_ce0 = ret


End Function

Function f_control_cerrar_gai() As Boolean
'6
Dim ret As Boolean

ret = False
If esta_terminado_ejv = False Then
    If hay_que_detener_ejv Then
        MsgBox "Se está deteniendo el proceso Gaia. Espere, por favor.", vbCritical
    Else
        MsgBox "Ha de terminar(F7) el proceso Gaia antes de cerrar esta ventana, o cerrar la ventana principal si desea terminar ya su sesión con Ejemplos de Vida.", vbCritical
    End If
        ret = True
Else
    If esta_detenido_ejv = False Then
        MsgBox "Se ha de detener el proceso Gaia antes de cerrar la ventana", vbCritical
        ret = True
    End If
End If

f_control_cerrar_gai = ret


End Function


Function f_peticion_unload_mdi_ejv() As Integer

    Dim ret As Integer
    Dim txt1 As String
    Dim txt2 As String
    
    Dim dev As Integer
        
    'Defino mensajes
    txt1 = "Hay procesos no finalizados ¿Desea cerrar el programa de todas formas?"
    txt2 = "Esto finalizará su sesión con Ejemplos de Vida"
    If idioma_ejv = CTE_INGLES Then
        txt1 = "There are not stopped jobs. ¿Are you sure that you want to close all the windows and exit?"
        txt2 = "This will finish your session with Ejemplos de Vida"
    End If
    
    If Not pedir_confirmacion_ejv Then
        s_fin_todo
    End If
    
    If f_hay_procesos_no_finalizados Then
        ret = MsgBox(txt1, vbYesNoCancel + vbQuestion)
        If ret = vbCancel Then
            dev = True
        Else
            If ret = vbYes Then
                If MsgBox(txt2, vbOKCancel + vbInformation) = vbCancel Then
                    dev = True
                Else
                    s_fin_todo
                End If
            Else
                dev = f_control_cerrar_va0 '1 4
                If Not dev Then
                    dev = f_control_cerrar_pal '2
                End If
                If Not dev Then
                    dev = f_control_cerrar_ce0 '3
                End If
                If Not dev Then
                    dev = f_control_cerrar_gai '6
                End If
            End If
        End If
    Else
        If idioma_ejv = CTE_INGLES Then
            txt1 = "This will end your session with Ejemplos de Vida"
        Else
            txt1 = "Esto finalizará su sesión con Ejemplos de Vida"
        End If
        If MsgBox(txt1, vbOKCancel + vbInformation) = vbCancel Then
            dev = True
        Else
            s_fin_todo
        End If
    End If
    

    f_peticion_unload_mdi_ejv = dev

End Function
Sub s_inicializar_estado_menus_ejv()

    Dim i As Integer
    Dim j As Integer

    '==============================================
    'Estado de los menus de ejecutar (4 operaciones, 11 programas)
    'El programa 0 es si no hay ninguno activo
    ReDim estado_ejecutar_ejv(1 To CTE_EXE_num_total, 0 To CTE_PROG_num_total) As Boolean
    'Por defecto todo prohibido
    For i = 1 To CTE_EXE_num_total
        For j = 0 To CTE_PROG_num_total
            estado_ejecutar_ejv(i, j) = False
        Next j
        s_cambiar_estado_enabled_ejecutar_ejv i, False
    Next i
    'Excepciones a EJECUTAR
    '==============================================
    
    
    '==============================================
    'Estado de los menus ver (16 operaciones, 11 programas)
    'El programa 0 es si no hay ninguno activo
    ReDim estado_ver_ejv(1 To CTE_VER_num_total, 0 To CTE_PROG_num_total) As Boolean
    'Por defecto todo prohibido
    For i = 1 To CTE_VER_num_total
        For j = 0 To CTE_PROG_num_total
            estado_ver_ejv(i, j) = False
        Next j
        s_cambiar_estado_enabled_menus_ejv i, False
    Next i
    
    'Excepciones a VER completas
    For j = 0 To CTE_PROG_num_total
        estado_ver_ejv(CTE_VER_DICCIONARIO, j) = True 'todos
        estado_ver_ejv(CTE_VER_GRAFICO, j) = True 'todos
        estado_ver_ejv(CTE_VER_OPCIONES1, j) = True 'todos
    Next j
    
    'Excepciones a VER en MDI
    frm_z0_mdi.mn_Diccionario.Enabled = True
    frm_z0_mdi.mn_Grafico.Enabled = True
    frm_z0_mdi.H_Grafico.Enabled = True
    frm_z0_mdi.mn_Opciones1.Enabled = True
    
    'Excepciones a VER con programas abiertos, todos
    For j = 1 To CTE_PROG_num_total
        estado_ver_ejv(CTE_VER_ESTADO_EJECUCION, j) = True 'todos
        estado_ver_ejv(CTE_VER_DICCIONARIO, j) = True 'todos
        estado_ver_ejv(CTE_VER_GRAFICO, j) = True 'todos
    Next j
    
    'Excepciones a VER especificas
    estado_ver_ejv(CTE_VER_REFRESCAR, CTE_HYP) = True
    estado_ver_ejv(CTE_VER_REFRESCAR, CTE_PRI) = True
    estado_ver_ejv(CTE_VER_REFRESCAR, CTE_CEL) = True
    estado_ver_ejv(CTE_VER_REFRESCAR, CTE_EXP) = True
    estado_ver_ejv(CTE_VER_REFRESCAR, CTE_PEZ) = True
    '==============================================
    

End Sub
Sub s_inicializar_arrays_color_ejv()

    nct_i_ejv = 33
    ncs_i_ejv = nct_i_ejv
    
    ReDim nct_ejv(1 To nct_i_ejv) As String
    nct_ejv(1) = "Rojo"
    nct_ejv(2) = "Rosa"
    nct_ejv(3) = "Naranja"
    nct_ejv(4) = "Amarillo"
    nct_ejv(5) = "Verde Brillante"
    nct_ejv(6) = "Verde Claro"
    nct_ejv(7) = "Verde Pálido"
    nct_ejv(8) = "Azul"
    nct_ejv(9) = "Negro"
    nct_ejv(10) = "Blanco"
    nct_ejv(11) = "Gris Claro"
    nct_ejv(12) = "Gris Oscuro"
    nct_ejv(13) = "Azul Claro"
    nct_ejv(14) = "Azulón"
    nct_ejv(15) = "Verde Oscuro"
    nct_ejv(16) = "Salmón"
    nct_ejv(17) = "Rosa Claro"
    nct_ejv(18) = "Morado"
    nct_ejv(19) = "Azul Oscuro"
    nct_ejv(20) = "Marrón Claro"
    nct_ejv(21) = "Naranja Oscuro"
    nct_ejv(22) = "Vino Tinto"
    nct_ejv(23) = "Verde Manzana"
    nct_ejv(24) = "Verde Plastico"
    nct_ejv(25) = "Azul Grisáceo"
    nct_ejv(26) = "Rojo Tierra"
    nct_ejv(27) = "Verde Azulón"
    nct_ejv(28) = "Hueso"
    nct_ejv(29) = "Rosaceo"
    nct_ejv(30) = "Oro"
    nct_ejv(31) = "Violeta"
    nct_ejv(32) = "Lavanda"
    nct_ejv(33) = "Dorado"



    ReDim cct_ejv(-3 To nct_i_ejv) As Long
    'Colores
    'Los colores se obtienen haciendo
    'cada numero es 0-255
    'rojo-verde-azul
    cct_ejv(CTE_DEGRADADOCOLOR) = -3
    cct_ejv(CTE_DEGRADADOGRIS) = -2
    cct_ejv(CTE_TRANSPARENTE) = -1
    
    cct_ejv(CTE_ROJO) = 255 'rojo = RGB(255 * 1, 255 * 0, 255 * 0)
    cct_ejv(CTE_ROSA) = 16711935 'rosa = RGB(255 * 1, 255 * 0, 255 * 1)
    cct_ejv(CTE_NARANJA) = 33023  'naranja = RGB(255 * 1, 255 * 0.5, 255 * 0)
    cct_ejv(CTE_AMARILLO) = 65535  'amarillo = RGB(255 * 1, 255 * 1, 255 * 0)
    cct_ejv(CTE_VERDEBRILLANTE) = 65280  'verde = RGB(255 * 0, 255 * 1, 255 * 0)
    cct_ejv(CTE_VERDECLARO) = 45568  'verde_claro = RGB(255 * 0, 255 * 0.7, 255 * 0)
    cct_ejv(CTE_VERDEPALIDO) = 6736998  'verde_oscuro = RGB(255 * 0.4, 255 * 0.8, 255 * 0.4)
    cct_ejv(CTE_AZUL) = 16711680  'azul = RGB(255 * 0, 255 * 0, 255 * 1)
    cct_ejv(CTE_NEGRO) = 0  'negro = RGB(255 * 0, 255 * 0, 255 * 0)
    cct_ejv(CTE_BLANCO) = 16777215  'blanco = RGB(255 * 1, 255 * 1, 255 * 1)
    cct_ejv(CTE_GRISCLARO) = 12566463  'grisclaro = RGB(255 * 0.75, 255 * 0.75, 255 * 0.75)
    cct_ejv(CTE_GRISOSCURO) = 5000268  'grisoscuro = RGB(255 * 0.3, 255 * 0.3, 255 * 0.3)
    cct_ejv(CTE_AZULCLARO) = 16776960 'azulclaro = RGB(255 * 0.5, 2555 * 0.5, 255 * 1)
    cct_ejv(CTE_AZULON) = 16744576 'morado = RGB(255 * 1, 255 * 0, 255 * 1)
    cct_ejv(CTE_VERDEOSCURO) = 3368499 'RGB(255 * 0.2, 255 * 0.4, 255 * 0.2)
    cct_ejv(CTE_SALMON) = 8356095
    cct_ejv(CTE_ROSACLARO) = 16744703
    cct_ejv(CTE_MORADO) = 8323200
    cct_ejv(CTE_AZULOSCURO) = 16711808
    cct_ejv(CTE_MARRONCLARO) = 3777215
    cct_ejv(CTE_NARANJAOSCURO) = 22222
    cct_ejv(CTE_VINOTINTO) = 4334215
    cct_ejv(CTE_VERDEMANZANA) = 323232
    cct_ejv(CTE_VERDEPLASTICO) = 6545166
    cct_ejv(CTE_AZULGRISACEO) = 8081698
    cct_ejv(CTE_ROJOTIERRA) = 128
    cct_ejv(CTE_VERDEAZULON) = 11180288
    cct_ejv(CTE_HUESO) = 15794175
    cct_ejv(CTE_ROSACEO) = 12496890
    cct_ejv(CTE_ORO) = 1548500
    cct_ejv(CTE_VIOLETA) = 13187213
    cct_ejv(CTE_LAVANDA) = 16442595
    cct_ejv(CTE_DORADO) = 61695


    s_copiar_array1str nct_ejv, ncs_ejv
    s_copiar_array1lng cct_ejv, ccs_ejv


End Sub

Function f_hay_procesos_no_finalizados() As Boolean

    Dim dev As Boolean

    dev = False
    
    If hay_que_detener_ejv = True Or esta_detenido_ejv = False Then
        dev = True
    End If
    If hay_que_terminar_ejv = True Or esta_terminado_ejv = False Then
        dev = True
    End If
    
    f_hay_procesos_no_finalizados = dev

End Function

Sub s_load_menu_ejv()

    s_centrar_ventana_ejv frm_z0_menu
    s_tratamiento_idioma_menu

    'Cargo las opciones
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 2"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 3"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 4"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 5"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 6"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 7"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 8"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 9"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 10"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 11"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 12"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_PAL '2
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 2"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 3"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 4"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_3R '3
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 2"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 3"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_PRI '4
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 2"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_CEL '5
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_GAI '6
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_EXP '7
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_CAD '8
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_PEZ '9
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_UVA '10
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case CTE_YXY '11
            frm_z0_menu.Cb_ejemplo.Clear
            frm_z0_menu.Cb_ejemplo.AddItem "Ejemplo 1"
            frm_z0_menu.Cb_ejemplo.ListIndex = 0
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select

    frm_z0_menu.mensaje.Visible = False
    frm_z0_menu.Timer1.Enabled = False

End Sub
Sub s_cerrar_prg(prg As Integer)

    Select Case prg
        Case CTE_HYP '1
            Unload frm_a0_va
            Unload frm_a1_inhyp
            Unload frm_a1_tiposhyp 'Estas se cargan aunque no se vean!
            Unload frm_a1_ophyp 'Estas se cargan aunque no se vean!
        Case CTE_PAL '2
            Unload frm_b2_pal
            Unload frm_b2_inpal
        Case CTE_3R '3
            Unload frm_c0_ce
            Unload frm_c3_in3r
        Case CTE_PRI '4
            Unload frm_a0_va
            Unload frm_a4_inpri
            Unload frm_a4_tipospri
        Case CTE_CEL '5
            Unload frm_a0_va
            Unload frm_a5_incel
        Case CTE_GAI '6
            Unload frm_a0_va
            'Fi_Cerrar_Base_Datos ya se cierra en el unload
        Case CTE_EXP '7
            Unload frm_a0_va
            Unload frm_a7_inexp
        Case CTE_CAD '8
            Unload frm_a0_va
        Case CTE_PEZ '9
            Unload frm_a0_va
            Unload frm_a9_inpez
        Case CTE_UVA '10
            Unload frm_a0_va
        Case CTE_YXY '11
            Unload frm_a0_va
        Case Else
            'No hay ninguno, no hace falta descargarlo
    End Select

End Sub

Sub s_ejemplo_click_ejv(ej As String)

    Dim s_mensaje As String
    
    s_mensaje = ""
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            'Ejecutar un ejemplo
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1: Se recomiendan unos 8.000 ciclos."
                Case "Ejemplo 2"
                    s_mensaje = "Ejemplo 2: Se recomiendan unos 10.000 ciclos."
                Case "Ejemplo 3"
                    s_mensaje = "Ejemplo 3: Este ejemplo es bastante rápido."
                Case "Ejemplo 4"
                    s_mensaje = "Ejemplo 4: Parece que siempre mueren todas."
                Case "Ejemplo 5"
                    s_mensaje = "Ejemplo 5: Ejemplo muy lento, que comienza con 300 hormigas rojas y 300 verdes. Se recomiendan unos 5.000 ciclos."
                Case "Ejemplo 6"
                    s_mensaje = "Ejemplo 6: Comienza con 50 rojas. Muy interesante. Se recomiendan unos 5.000 ciclos."
                Case "Ejemplo 7"
                    s_mensaje = "Ejemplo 7: Comienza con 50 rojas. Ciclos muy marcados. La población no siempre sobrevive. Se recomiendan más de 1.000 ciclos."
                Case "Ejemplo 8"
                    s_mensaje = "Ejemplo 8: Ejemplo muy rápido."
                Case "Ejemplo 9"
                    s_mensaje = "Ejemplo 9: Muy interesante. Se recomiendan unos 4.000 ciclos."
                Case "Ejemplo 10"
                    s_mensaje = "Ejemplo 10: En este caso se observa cómo se sucenden ciclos de hormigas verdes y amarillas"
                Case "Ejemplo 11"
                    s_mensaje = "Ejemplo 11: "
                Case "Ejemplo 12"
                    s_mensaje = "Ejemplo 12: Despues de un tiempo, todas las hormigas se cruzan sin tocarse."
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_PAL '2
            Select Case ej
                Case "Ejemplo 1"
                   's_mensaje=  "Ejemplo 1: Una frase de siete letras"
                Case "Ejemplo 2"
                    's_mensaje=  "Ejemplo 2: Una frase más larga"
                Case "Ejemplo 3"
                    's_mensaje=  "Ejemplo 3: Una frase más corta"
                Case "Ejemplo 4"
                    s_mensaje = "Cadenas binarias"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_3R '3
            Select Case ej
                Case "Ejemplo 1"
                    's_mensaje=  "Ejemplo 1: En este ejemplo las entidades tienen muy pocas reglas por lo que nunca llegan a poseer un gran conocimiento, pero es muy útil para tener una primera visión del funcionamiento del programa.", vbInformation
                Case "Ejemplo 2"
                    's_mensaje=  "Ejemplo 2:", vbInformation
                Case "Ejemplo 3"
                    's_mensaje=  "Ejemplo 3:", vbInformation
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_PRI '4
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1: Algunos agentes muy distintos"
                Case "Ejemplo 2"
                    s_mensaje = "Ejemplo 2: Todos los agentes"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_CEL '5
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_GAI '6
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_EXP '7
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_CAD '8
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_PEZ '9
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_UVA '10
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case CTE_YXY '11
            Select Case ej
                Case "Ejemplo 1"
                    s_mensaje = "Ejemplo 1"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
    
    If s_mensaje <> "" Then
        frm_z0_menu.Timer1.Enabled = True
        frm_z0_menu.mensaje.Visible = True
        frm_z0_menu.mensaje.Caption = s_mensaje
    End If


End Sub

Sub s_mostrar_aviso_imagen()

    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_PAL '2
            frm_b2_pal.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_b2_pal.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_3R '3
            frm_c0_ce.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_c0_ce.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_PRI '4
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_CEL '5
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_GAI '6
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_EXP '7
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_CAD '8
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_PEZ '9
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_UVA '10
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case CTE_YXY '11
            frm_a0_va.aviso_ejecutar.Visible = mostrar_aviso_imagen_ejv
            frm_a0_va.Imagen.Visible = mostrar_aviso_imagen_ejv
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select


End Sub

Sub s_terminar_va0(terminar As Boolean)

    Screen.MousePointer = CTE_ARENA
    s_botones_enabled_va0 (True)
        
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, False
    
    
    
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            s_mostrar_estado_semaforo frm_a1_inhyp, CTE_DETENIENDO
        Case CTE_PRI '4
            s_mostrar_estado_semaforo frm_a4_inpri, CTE_DETENIENDO
        Case CTE_CEL '5
            s_mostrar_estado_semaforo frm_a5_incel, CTE_DETENIENDO
        Case CTE_GAI '6
            s_mostrar_estado_semaforo frm_a6_ingaia, CTE_DETENIENDO
        Case CTE_EXP '7
            s_mostrar_estado_semaforo frm_a7_inexp, CTE_DETENIENDO
        Case CTE_PEZ '9
            s_mostrar_estado_semaforo frm_a9_inpez, CTE_DETENIENDO
        Case CTE_UVA '10
            s_mostrar_estado_semaforo frm_aA_inuva, CTE_DETENIENDO
        Case CTE_YXY '11
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
    
    
        
    If terminar Then    'es terminar
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES3, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_MAPA, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPOS_AGENTES, True
        
        hay_que_detener_ejv = True
        hay_que_terminar_ejv = True
    End If
    
    If Not terminar Then
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES3, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_MAPA, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPOS_AGENTES, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_JUGAR_CONTRA_ORDENADOR, True
        hay_que_detener_ejv = True
        hay_que_terminar_ejv = False
    End If

End Sub

Sub s_terminar_pal()

    Screen.MousePointer = CTE_ARENA

    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_DICCIONARIO, True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, False
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, False
    
    'Inicializamos a vacio
    frm_b2_pal.Refresh

    hay_que_detener_ejv = True
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True


End Sub

Sub s_terminar_ce0(terminar As Boolean)
    
    Screen.MousePointer = CTE_ARENA
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, False
    
    s_mostrar_estado_semaforo frm_c3_in3r, CTE_DETENIENDO
    
    If terminar Then    'es terminar
        
        'Inicializamos a vacio
        frm_c0_ce.Refresh
        
        s_botones_activos_3r False
        
        's_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
        's_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, False
        's_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, False
        's_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES3, False
        's_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION, False
        '    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_EVALUACION, False
        '    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_SELECCION, False
        '    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION, False
        '        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES, False
        '        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO, False
        
        ver_agentes_3r = False
        hay_que_detener_ejv = True
        hay_que_terminar_ejv = True
    End If
    
    If Not terminar Then 'es detener
        hay_que_detener_ejv = True
        hay_que_terminar_ejv = False
    End If

    frm_c3_in3r.Refresh
    frm_c3_in3r.SetFocus

End Sub


Function f_sobrecruzamiento_ejv(cadena_padre As String, cadena_madre As String, long_elemento As Integer, metodo As Integer)

    Dim i As Long
    Dim num_elementos As Integer

    Select Case metodo
        Case CTE_ALTERNOS
            If Len(cadena_padre) <> Len(cadena_madre) Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End If
            num_elementos = Len(cadena_padre) / long_elemento
            If Len(cadena_padre) <> Len(cadena_madre) Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End If
            For i = 1 To Len(cadena_padre) / long_elemento
                f_sobrecruzamiento_ejv = f_extraer_subcadena_ejv(cadena_padre, i, long_elemento)
            Next i
        Case CTE_1_PTO_CORTE
            
        Case Else
    End Select

    f_sobrecruzamiento_ejv = 1

End Function

Sub s_grabar_opciones_generales_ejv()

    '1 Color de Fondo
    cfondo_ejv = frm_z0_op.Cb_Color.ListIndex + 1
    If frm_z0_op.Ch_EliminarColor = 1 Then
        eliminar_cfondo_ejv = True
    Else
        eliminar_cfondo_ejv = False
    End If
    
    '2 Condición de Parada
    If frm_z0_op.Ch_CondParadaNumMaxCiclos = 1 Then
        CondParadaNumMaxCiclos_ejv = True
    Else
        CondParadaNumMaxCiclos_ejv = False
    End If
    CondParadaNumMaxCiclosFinal_ejv = frm_z0_op.Op_CondParadaNumMaxCiclos
    If frm_z0_op.Ch_CondParadaFechaHora = 1 Then
        CondParadaFechaHora_ejv = True
    Else
        CondParadaFechaHora_ejv = False
    End If
    CondParadaFecha_ejv = frm_z0_op.Op_CondParadaFecha
    CondParadaHora_ejv = frm_z0_op.Op_CondParadaHora
    If frm_z0_op.Ch_CondParadaPeso = 1 Then
        CondParadaPeso_ejv = True
    Else
        CondParadaPeso_ejv = False
    End If
    CondParadaPesoNecesario_ejv = frm_z0_op.Op_CondParadaPeso
    
    '3 Grabar Resumen
        'Grafico
    If frm_z0_op.Ch_GrabarResumenGra = 1 Then
        un_ej_grabar_gra_ejv = True
    Else
        un_ej_grabar_gra_ejv = False
    End If
    un_ej_fichero_gra_ejv = f_nombre_completo(frm_z0_op.Op_GrabarResumenGraC, frm_z0_op.Op_GrabarResumenGraF)
        'Texto
    If frm_z0_op.Ch_GrabarResumenTxt = 1 Then
        un_ej_grabar_resumen_txt_ejv = True
    Else
        un_ej_grabar_resumen_txt_ejv = False
    End If
    un_ej_fichero_resumen_txt_ejv = f_nombre_completo(frm_z0_op.Op_GrabarResumenTxtC, frm_z0_op.Op_GrabarResumenTxtF)
        'Excel
    If frm_z0_op.Ch_GrabarResumenExcel = 1 Then
        un_ej_grabar_resumen_xls_ejv = True
    Else
        un_ej_grabar_resumen_xls_ejv = False
    End If
    un_ej_fichero_resumen_xls_ejv = f_nombre_completo(frm_z0_op.Op_GrabarResumenExcelC, frm_z0_op.Op_GrabarResumenExcelF)
        
        'Autoguardado, máximo guardado, cabeceras y reemplazar
    autoguardado_ejv = frm_z0_op.Op_Autoguardado
    max_guardado_ejv = frm_z0_op.Op_MaxGuardado
    If frm_z0_op.Op_Cabeceras = 1 Then
        cabeceras_ejv = True
    Else
        cabeceras_ejv = False
    End If
    If frm_z0_op.Op_Reemplazar = 1 Then
        reemplazar_fic_ejv = True
    Else
        reemplazar_fic_ejv = False
    End If
    
        
    '4 Funcion de Azar
    If frm_z0_op.Op_AzarVB = True Then
        'Azar de VB
        tipo_funcion_azar_ejv = CTE_AZARVB
        If frm_z0_op.Ch_Randomize.Enabled = True Then
            randomize_ejv = True
        Else
            randomize_ejv = False
        End If
    Else
        'Azar de fichero
        tipo_funcion_azar_ejv = CTE_AZARFIC
        randomize_ejv = False
        azar_carpeta_ejv = frm_z0_op.Op_AzarFicC
        azar_fichero_ejv = frm_z0_op.Op_AzarFicF
        azar_fichero_num_char_ejv = frm_z0_op.Op_AzarFicNumCh
    End If
    

End Sub


Sub s_cargar_opciones_generales_ejv()

    Dim i As Integer
    
    '1 Color de Fondo
    For i = 1 To nct_i_ejv
        frm_z0_op.Cb_Color.AddItem nct_ejv(i)
    Next i
    frm_z0_op.Cb_Color.ListIndex = cfondo_ejv - 1
    If eliminar_cfondo_ejv Then
        frm_z0_op.Ch_EliminarColor = 1
    Else
        frm_z0_op.Ch_EliminarColor = 0
    End If
    
    '2 Condición de Parada
    If CondParadaNumMaxCiclos_ejv Then
        frm_z0_op.Ch_CondParadaNumMaxCiclos = 1
    Else
        frm_z0_op.Ch_CondParadaNumMaxCiclos = 0
    End If
    frm_z0_op.Op_CondParadaNumMaxCiclos = CondParadaNumMaxCiclosFinal_ejv
    If CondParadaFechaHora_ejv Then
        frm_z0_op.Ch_CondParadaFechaHora = 1
    Else
        frm_z0_op.Ch_CondParadaFechaHora = 0
    End If
    frm_z0_op.Op_CondParadaFecha = CondParadaFecha_ejv
    frm_z0_op.Op_CondParadaHora = CondParadaHora_ejv
    If CondParadaPeso_ejv Then
        frm_z0_op.Ch_CondParadaPeso = 1
    Else
        frm_z0_op.Ch_CondParadaPeso = 0
    End If
    frm_z0_op.Op_CondParadaPeso = CondParadaPesoNecesario_ejv
    
    '3 Grabar Resumen
        'Grafico
    If un_ej_grabar_gra_ejv Then
        frm_z0_op.Ch_GrabarResumenGra = 1
    Else
        frm_z0_op.Ch_GrabarResumenGra = 0
    End If
    frm_z0_op.Op_GrabarResumenGraC = f_path_fichero(un_ej_fichero_gra_ejv, CTE_C_SAL_GRA)
    frm_z0_op.Op_GrabarResumenGraF = f_nombre_fichero(un_ej_fichero_gra_ejv)
        'Texto
    If un_ej_grabar_resumen_txt_ejv Then
        frm_z0_op.Ch_GrabarResumenTxt = 1
    Else
        frm_z0_op.Ch_GrabarResumenTxt = 0
    End If
    frm_z0_op.Op_GrabarResumenTxtC = f_path_fichero(un_ej_fichero_resumen_txt_ejv, CTE_C_SAL_TXT)
    frm_z0_op.Op_GrabarResumenTxtF = f_nombre_fichero(un_ej_fichero_resumen_txt_ejv)
        'Excel
    If un_ej_grabar_resumen_xls_ejv Then
        frm_z0_op.Ch_GrabarResumenExcel = 1
    Else
        frm_z0_op.Ch_GrabarResumenExcel = 0
    End If
    frm_z0_op.Op_GrabarResumenExcelC = f_path_fichero(un_ej_fichero_resumen_xls_ejv, CTE_C_SAL_XLS)
    frm_z0_op.Op_GrabarResumenExcelF = f_nombre_fichero(un_ej_fichero_resumen_xls_ejv)
        'Autoguardado y máximo guardado
    frm_z0_op.Op_Autoguardado = autoguardado_ejv
    frm_z0_op.Op_MaxGuardado = max_guardado_ejv
    If cabeceras_ejv Then
        frm_z0_op.Op_Cabeceras = 1
    Else
        frm_z0_op.Op_Cabeceras = 0
    End If
    If reemplazar_fic_ejv Then
        frm_z0_op.Op_Reemplazar = 1
    Else
        frm_z0_op.Op_Reemplazar = 0
    End If
    

    '4 Funcion de Azar
    If tipo_funcion_azar_ejv = CTE_AZARVB Then
        'Azar de VB
        frm_z0_op.Op_AzarVB = True
        frm_z0_op.Op_AzarFic = False
        
        frm_z0_op.Ch_Randomize.Enabled = True
        If randomize_ejv Then
            frm_z0_op.Ch_Randomize = 1
        Else
            frm_z0_op.Ch_Randomize = 0
        End If
        
        frm_z0_op.Op_AzarFicC.Enabled = False
        frm_z0_op.Op_AzarFicF.Enabled = False
        frm_z0_op.Op_AzarFicNumCh.Enabled = False
        frm_z0_op.Fic_Azar.Enabled = False
        
        frm_z0_op.Op_AzarFicC = azar_carpeta_ejv
        frm_z0_op.Op_AzarFicF = f_nombre_fichero(azar_fichero_ejv)
    Else
        'Azar de fichero
        frm_z0_op.Op_AzarVB = False
        frm_z0_op.Op_AzarFic = True
        
        frm_z0_op.Ch_Randomize.Enabled = False
        frm_z0_op.Ch_Randomize = 0
        
        frm_z0_op.Op_AzarFicC.Enabled = True
        frm_z0_op.Op_AzarFicF.Enabled = True
        frm_z0_op.Op_AzarFicNumCh.Enabled = True
        frm_z0_op.Fic_Azar.Enabled = True
        
        frm_z0_op.Op_AzarFicC = azar_carpeta_ejv
        frm_z0_op.Op_AzarFicF = f_nombre_fichero(azar_fichero_ejv)
        frm_z0_op.Op_AzarFicNumCh = azar_fichero_num_char_ejv
    End If
    

End Sub

Sub s_aceptar_menu_ejv(ej As String, opcion As Integer)

    Dim s_titulo As String
    Dim tmp As Integer
    
    s_titulo = ""
        
    'Si habia antes un programa abierto, lo cierro
    tmp = num_prg_activo_ejv
    If num_prg_anterior_activo_ejv <> CTE_NINGUNO Then
        s_cerrar_prg num_prg_anterior_activo_ejv
        'como el unload pone a NINGUNO la variable num_prg_activo_ejv, y ahora
        'no descargo el activo sino el anterior, salvo en tmp el activo para
        'que no se borre (en cambio si lo pongo al anterior)
        num_prg_activo_ejv = tmp
        num_prg_anterior_activo_ejv = CTE_NINGUNO
    End If
    
        
    
    s_titulo = nombre_programa_ejv(num_prg_activo_ejv) & ". " & ej
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case "Ejemplo 3"
                            num_ej_activo_ejv = 3
                        Case "Ejemplo 4"
                            num_ej_activo_ejv = 4
                        Case "Ejemplo 5"
                            num_ej_activo_ejv = 5
                        Case "Ejemplo 6"
                            num_ej_activo_ejv = 6
                        Case "Ejemplo 7"
                            num_ej_activo_ejv = 7
                        Case "Ejemplo 8"
                            num_ej_activo_ejv = 8
                        Case "Ejemplo 9"
                            num_ej_activo_ejv = 9
                        Case "Ejemplo 10"
                            num_ej_activo_ejv = 10
                        Case "Ejemplo 11"
                            num_ej_activo_ejv = 11
                        Case "Ejemplo 12"
                            num_ej_activo_ejv = 12
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_a0_va.WindowState = CTE_MAXIMIZED
            frm_a0_va.Show
            frm_a0_va.Caption = s_titulo
        Case CTE_PAL '2
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                           num_ej_activo_ejv = 1
                           's_mensaje=  "Ejemplo 1: Una frase de siete letras", vbInformation
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                            's_mensaje=  "Ejemplo 2: Una frase más larga", vbInformation
                        Case "Ejemplo 3"
                            num_ej_activo_ejv = 3
                            's_mensaje=  "Ejemplo 3: Una frase más corta", vbInformation
                        Case "Ejemplo 4"
                            num_ej_activo_ejv = 4
                            's_mensaje=  "Ejemplo 4: Una frase más corta", vbInformation
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_b2_pal.Show
            frm_b2_pal.Caption = s_titulo
        Case CTE_3R '3
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                            If pedir_confirmacion_ejv Then
                                MsgBox "Este ejemplo es para pruebas, los agentes sólo tienen una regla y no pueden aprender mucho", vbInformation
                            End If
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case "Ejemplo 3"
                            num_ej_activo_ejv = 3
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_c0_ce.Show
            frm_c0_ce.Caption = s_titulo
        Case CTE_PRI '4
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_a0_va.WindowState = CTE_MAXIMIZED
            frm_a0_va.Show
            frm_a0_va.Caption = s_titulo
        Case CTE_CEL '5
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_a0_va.WindowState = CTE_MAXIMIZED
            frm_a0_va.Show
            frm_a0_va.Caption = s_titulo
        Case CTE_GAI '6
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_a0_va.WindowState = CTE_MAXIMIZED
            frm_a0_va.Show
            frm_a0_va.Caption = s_titulo
        Case CTE_EXP '7
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_a0_va.WindowState = CTE_MAXIMIZED
            frm_a0_va.Show
            frm_a0_va.Caption = s_titulo
        Case CTE_CAD '8
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            'frm_a0_va.WindowState = CTE_MAXIMIZED
            'frm_a0_va.Show
            'frm_a0_va.Caption = s_titulo
        Case CTE_PEZ '9
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_a0_va.Show
            frm_a0_va.WindowState = CTE_MAXIMIZED
            frm_a0_va.Caption = s_titulo
        Case CTE_UVA '10
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_a0_va.WindowState = CTE_MAXIMIZED
            frm_a0_va.Show
            frm_a0_va.Caption = s_titulo
        Case CTE_YXY '11
            Select Case opcion
                Case 1
                    'Ejecutar un ejemplo
                    Select Case ej
                        Case "Ejemplo 1"
                            num_ej_activo_ejv = 1
                        Case "Ejemplo 2"
                            num_ej_activo_ejv = 2
                        Case Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
                    End Select
                Case 2
                    'Crear un ejemplo
                    num_ej_activo_ejv = 1
                Case 3
                    'Cargar un ejemplo
                    num_ej_activo_ejv = 1
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
            End Select
            Unload frm_z0_menu
            frm_a0_va.WindowState = CTE_MAXIMIZED
            frm_a0_va.Show
            frm_a0_va.Caption = s_titulo
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
End Sub


Sub s_click_programa_ejv(programa As Integer)

    Dim texto As String

    num_prg_anterior_activo_ejv = num_prg_activo_ejv
    
    num_prg_activo_ejv = programa
    texto = nombre_programa_ejv(programa)
    If Not automatico_ejv Then
        'Muestro el menu para elegir ejemplo
        frm_z0_menu.Caption = texto
        frm_z0_menu.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    End If

End Sub

Sub s_activar_opciones_generales_ejv()

    '1 Condición de Parada
    '2 Grabar Resumen (Ejecución de un ejemplo)
    '3 Funcion de Azar
    If tipo_funcion_azar_ejv = CTE_AZARFIC Then
        'Cargo el array de azar
        s_aut_leer_digitos_azar
    End If
    If randomize_ejv Then
        Randomize
        If Not automatico_ejv Then
            If Rnd = 0.6131098 Then
                MsgBox "Felicidades, se acaba de producir un suceso altamente improbable. De todas formas no es gran cosa. La probabilidad de este suceso es casi tan pequeña como su importancia, así que esta noche puede dormir tranquilo. De todas formas, si así lo desea, no deje de pensar que el " & Date & " fué un gran día. Habitualmente sobran razones.", vbExclamation
            End If
        End If
    End If
    '4 Color de Fondo
    If eliminar_cfondo_ejv Then
        s_inicializar_arrays_color_ejv
        f_borra_elemento_array_string cfondo_ejv, ncs_ejv()
        f_borra_elemento_array_long cfondo_ejv, ccs_ejv()
        ncs_i_ejv = ncs_i_ejv - 1
    End If
    
    
End Sub

Sub s_inicializar_opciones1()

    'OPCIONES I
    'GENERALES DE EJEMPLOS DE VIDA(POR DEFECTO)
    '1 Condición de Parada
    CondParadaNumMaxCiclos_ejv = False
    CondParadaNumMaxCiclosFinal_ejv = 2
    CondParadaFechaHora_ejv = False
    CondParadaFecha_ejv = Date
    CondParadaHora_ejv = Time
    CondParadaPeso_ejv = False
    CondParadaPesoNecesario_ejv = 0.8
    '2 Grabar Resumen
    un_ej_grabar_gra_ejv = True
    un_ej_fichero_gra_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_GRA), "res.gra")
    un_ej_grabar_resumen_txt_ejv = False
    un_ej_fichero_resumen_txt_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_TXT), "res.txt")
    un_ej_grabar_resumen_xls_ejv = False
    un_ej_fichero_resumen_xls_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS), "res.xls")
    'Autoguardado, máximo guardado, cabeceras y reemplazar
    max_guardado_ejv = 1000000
    autoguardado_ejv = 100
    cabeceras_ejv = False
    reemplazar_fic_ejv = True
    '3 Función de Azar
    tipo_funcion_azar_ejv = CTE_AZARVB
    randomize_ejv = True
    azar_carpeta_ejv = path_largo_ejv(CTE_C_ENT_RAN)
    azar_fichero_ejv = "pi49999.ran"
    azar_fichero_num_char_ejv = "50"
    '4 Color de Fondo
    cfondo_ejv = CTE_GRISCLARO
    eliminar_cfondo_ejv = True


    '============================

    '1 Condición de Parada
    '2 Grabar Resumen
    '3 Función de Azar
    azar_en_memoria_ejv = False
    indice_azar = 1
    '4 Color de Fondo
    s_inicializar_arrays_color_ejv
    cfondo_ejv = CTE_GRISCLARO

End Sub

Sub s_inicializar_arrays_programa_ejv()

    Screen.MousePointer = CTE_ARENA
    
    s_inicializar_opciones1
    s_cargar_apellidos_va0
    ha_cambiado_el_diccionaro_pal = False
    
    num_menu_viejos_ejv = 1
    separacion_grafico_gra = 14
    hay_solucion_anterior_gauss = False
    habilitar_change_zoom_va0 = False
    
    ReDim alfabeto_binario_ejv(1 To 2) As String
    alfabeto_binario_ejv(1) = "0"
    alfabeto_binario_ejv(2) = "1"

    ReDim nombre_programa_ejv(0 To CTE_PROG_num_total) As String
    nombre_programa_ejv(CTE_NINGUNO) = "Ninguno"
    nombre_programa_ejv(CTE_HYP) = "Hormigas y Plantas"
    nombre_programa_ejv(CTE_PAL) = "Palabras y Frases"
    nombre_programa_ejv(CTE_3R) = "Tres en Raya"
    nombre_programa_ejv(CTE_PRI) = "El dilema del prisionero"
    nombre_programa_ejv(CTE_CEL) = "Celdas"
    nombre_programa_ejv(CTE_GAI) = "Plataforma Gaia"
    nombre_programa_ejv(CTE_EXP) = "Explorando Mapas"
    nombre_programa_ejv(CTE_CAD) = "Cadenas"
    nombre_programa_ejv(CTE_PEZ) = "Peces"
    nombre_programa_ejv(CTE_UVA) = "Universo"
    nombre_programa_ejv(CTE_YXY) = "YXY"
    
    
End Sub

Sub s_error_num_prog(num As Integer)

    If idioma_ejv = CTE_INGLES Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error: wrong program number: " & num
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: número de programa no valido: " & num
    End If

End Sub

Sub s_error_ejv(opcion_finalizar As Boolean, mensaje As String)

    If automatico_ejv Then
        s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, mensaje
    Else
        If idioma_ejv = CTE_INGLES Then
            If opcion_finalizar = CON_OPCION_FINALIZAR Then
                If MsgBox(mensaje & " Do you want to continue the applicaction?", vbYesNo + vbCritical) = vbNo Then
                    s_fin_todo
                End If
            Else
                MsgBox mensaje, vbCritical
            End If
        Else
            If opcion_finalizar = CON_OPCION_FINALIZAR Then
                If MsgBox(mensaje & " ¿Desea continuar la aplicación?", vbYesNo + vbCritical) = vbNo Then
                    s_fin_todo
                End If
            Else
                MsgBox mensaje, vbCritical
            End If
        End If
    End If

End Sub


Sub s_condiciones_parada_ejv(tipo As Boolean)

    Dim i As Integer
    
    'Condicion de parada especial: ya han jugado todos contra todos (solo para este caso especial)
    If num_prg_activo_ejv = CTE_PRI And todos_contra_todos_pri Then
        If agente_actual_ejv > numero_total_de_agentes_ejv Then 'Podria ser mayor que +1 si la ultima accion fue eliminar agente
            finalizacion_usuario_ejv = False
            s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
            Exit Sub
        End If
    End If
    'Condicion de parada: ciclo maximo
    If CondParadaNumMaxCiclos_ejv Then
        Select Case tipo
            Case CTE_PARADA_POR_IGUAL
                If ciclo_ejv = CondParadaNumMaxCiclosFinal_ejv Then
                    finalizacion_usuario_ejv = False
                    s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                    Exit Sub
                End If
            Case CTE_PARADA_POR_MAYOR
                If ciclo_ejv > CondParadaNumMaxCiclosFinal_ejv Then
                    finalizacion_usuario_ejv = False
                    s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                    Exit Sub
                End If
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: "
        End Select
    End If
    'Condicion de parada: tiempo
    If CondParadaFechaHora_ejv Then
        If Date > CondParadaFecha_ejv Or (Date = CondParadaFecha_ejv And Time > CondParadaHora_ejv) Then
            finalizacion_usuario_ejv = False
            s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
            Exit Sub
        End If
    End If
    'Condicion de parada: peso (energia maxima)
    If CondParadaPeso_ejv Then
        Select Case num_prg_activo_ejv
            Case CTE_HYP '1
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_va0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_PAL '2
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_pal(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_3R '3
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_ce0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_PRI '4
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_va0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_CEL '5
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_va0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_GAI '6
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_va0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_EXP '7
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_va0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_CAD '8
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_ce0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_PEZ '9
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_va0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_UVA '10
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_va0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case CTE_YXY '11
                For i = 1 To numero_total_de_agentes_ejv
                    If peso_agente_va0(i) >= CondParadaPesoNecesario_ejv Then
                        finalizacion_usuario_ejv = False
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                        Exit Sub
                    End If
                Next i
            Case Else
                s_error_num_prog num_prg_activo_ejv
        End Select
    
    End If

End Sub

Sub s_inicializar_ejemplo_elegido_ejv()
    
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            s_inicializar_ejemplo_elegido_hyp
        Case CTE_PAL '2
            s_inicializar_ejemplo_elegido_pal
            s_activar_opciones_pal 'por el diccionario
        Case CTE_3R '3
            s_inicializar_ejemplo_elegido_3r
        Case CTE_PRI '4
            s_inicializar_ejemplo_elegido_pri
        Case CTE_CEL '5
            s_inicializar_ejemplo_elegido_cel
        Case CTE_GAI '6
            s_inicializar_ejemplo_elegido_gai
        Case CTE_EXP '7
            s_inicializar_ejemplo_elegido_exp
        Case CTE_CAD '8
            s_inicializar_ejemplo_elegido_cad
        Case CTE_PEZ '9
            s_inicializar_ejemplo_elegido_pez
        Case CTE_UVA '10
            s_inicializar_ejemplo_elegido_uva
        Case CTE_YXY '11
            s_inicializar_ejemplo_elegido_yxy
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
    

End Sub

Sub s_fin_bucle_general_ejv()


    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            s_mostrar_estado_semaforo frm_a1_inhyp, CTE_DETENIDO
            frm_a1_inhyp.Show
            frm_a1_inhyp.SetFocus
        Case CTE_PAL '2
            s_mostrar_estado_semaforo frm_b2_inpal, CTE_DETENIDO
            frm_b2_inpal.Show
            frm_b2_inpal.SetFocus
        Case CTE_3R '3
            s_mostrar_estado_semaforo frm_c3_in3r, estado_3r
            frm_c3_in3r.Show
            frm_c3_in3r.SetFocus
        Case CTE_PRI '4
            s_mostrar_estado_semaforo frm_a4_inpri, CTE_DETENIDO
            frm_a4_inpri.Show
            frm_a4_inpri.SetFocus
        Case CTE_CEL '5
            s_mostrar_estado_semaforo frm_a5_incel, CTE_DETENIDO
            frm_a5_incel.Show
            frm_a5_incel.SetFocus
        Case CTE_GAI '6
            s_mostrar_estado_semaforo frm_a6_ingaia, CTE_DETENIDO
            frm_a6_ingaia.Show
            frm_a6_ingaia.SetFocus
        Case CTE_EXP '7
            s_mostrar_estado_semaforo frm_a7_inexp, CTE_DETENIDO
            frm_a7_inexp.Show
            frm_a7_inexp.SetFocus
        Case CTE_PEZ '9
            s_mostrar_estado_semaforo frm_a9_inpez, CTE_DETENIDO
            frm_a9_inpez.Show
            frm_a9_inpez.SetFocus
        Case CTE_UVA '10
            s_mostrar_estado_semaforo frm_aA_inuva, CTE_DETENIDO
            frm_aA_inuva.Show
            frm_aA_inuva.SetFocus
        Case CTE_YXY '11
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select

    If hay_que_terminar_ejv Then
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
        esta_terminado_ejv = True
    Else
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, False
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, True
        esta_terminado_ejv = False
    End If

    If esta_terminado_ejv Then
        'Pongo habilitado todos los programas
        s_cambiar_estado_enabled_programas_todos_ejv True
        'La siguiente vez que entre aqui si sera la primera
        es_la_primera_vez_ejv = True
    Else
        'La siguiente vez que entre aqui ya no sera la primera
        es_la_primera_vez_ejv = False
    End If
    s_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, True
    esta_detenido_ejv = True
    hay_que_detener_ejv = False
    hay_que_terminar_ejv = False
    'Salvo y cierro los ficheros
    s_cerrar_ficheros_un_ejemplo_ejv
    Screen.MousePointer = CTE_DEFECTO

End Sub

Sub s_mostrar_docum_html_ejv()

    Dim fic As String
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            fic = f_nombre_completo(path_largo_ejv(CTE_C_DOC_WEB), "ejv1_c.htm")
        Case CTE_PAL '2
            fic = f_nombre_completo(path_largo_ejv(CTE_C_DOC_WEB), "ejv2_c.htm")
        Case CTE_3R '3
            fic = f_nombre_completo(path_largo_ejv(CTE_C_DOC_WEB), "ejv3_c.htm")
        Case CTE_PRI '4
            fic = f_nombre_completo(path_largo_ejv(CTE_C_DOC_WEB), "ejv4_c.htm")
        'Case CTE_CEL '5
        Case CTE_GAI '6
            fic = f_nombre_completo(path_largo_ejv(CTE_C_DOC_WEB), "ejv6_c.htm")
        Case CTE_EXP '7
            fic = f_nombre_completo(path_largo_ejv(CTE_C_DOC_WEB), "ejv7_c.htm")
        'Case CTE_CAD '8
        'Case CTE_PEZ '9
        'Case CTE_UVA '10
        'Case CTE_YXY '11
        Case Else
            fic = f_nombre_completo(path_largo_ejv(CTE_C_DOC_WEB), "ejv_c.htm")
    End Select
    

On Error GoTo abrir_error
    
    Dim Lsalida As Long
    Lsalida = Shell("start " & fic)
    
    Exit Sub
abrir_error:
    MsgBox "No ha sido posible abrir automáticamente el fichero de documentación html: " & fic, vbInformation

End Sub
