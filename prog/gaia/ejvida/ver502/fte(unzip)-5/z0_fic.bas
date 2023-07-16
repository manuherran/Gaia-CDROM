Attribute VB_Name = "bas_z0_fic"
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

Sub s_usr_guardar_mapa_ma0()

    Dim linea As String
    Dim p As Double
    Dim f As Double
    Dim c As Double
    Dim dato As Integer
    Dim s_dato As String
    Dim primera_coma As Integer
    Dim fila_vacia As Boolean
    
    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_MAP)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Guardar Fichero de Mapa"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "Guardar"
    frm_z0_fic.File1.Pattern = "*.map"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path-fichero
    frm_a0_mapa.Caption = nombre_fichero_ejv
    
    'Compruebo que no existe
    If f_existe_fichero(nombre_fichero_ejv) Then
        If MsgBox("El fichero ya existe. ¿Desea reemplazarlo?", vbQuestion + vbOKCancel) = vbCancel Then
            'Ha elegido cancelar la operación
            GoTo fin
        End If
    End If
    
    'Abro y leo
    Open nombre_fichero_ejv For Output As #CTE_FIC_03_MAP
    'Pisos
    linea = "PISOS=" & mapa_pisos_ma0
    Print #CTE_FIC_03_MAP, linea 'El Write graba entre "" y el Print no
    'Filas
    linea = "FILAS=" & mapa_filas_ma0
    Print #CTE_FIC_03_MAP, linea 'El Write graba entre "" y el Print no
    'Columnas
    linea = "COLUMNAS=" & mapa_columnas_ma0
    Print #CTE_FIC_03_MAP, linea
    'Modo de presentacion
    Select Case ver_zoom_ma0
        Case CTE_ZOOM_DETALLE
            linea = "ZOOM=" & CTE_tZOOM_DETALLE
        Case CTE_ZOOM_PANORAMICA
            linea = "ZOOM=" & CTE_tZOOM_PANORAMICA
        Case CTE_ZOOM_PIXELS
            linea = "ZOOM=" & CTE_tZOOM_PIXELS
        Case CTE_ZOOM_3D
            linea = "ZOOM=" & CTE_tZOOM_3D
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
    End Select
    Print #CTE_FIC_03_MAP, linea
    'Celdas de obstaculos
    Screen.MousePointer = CTE_ARENA
    For p = 1 To mapa_pisos_ma0
    linea = "PISO=" & p
    Print #CTE_FIC_03_MAP, linea 'El Write graba entre "" y el Print no
    fila_vacia = True
    For f = 1 To mapa_filas_ma0
        linea = ""
        For c = 1 To mapa_columnas_ma0
            If mapa_ma0(p, f, c) = True Then
                linea = linea & "1,"
                fila_vacia = False
            Else
                linea = linea & "0,"
            End If
        Next c
       'quitamos la coma final
        linea = Left(linea, Len(linea) - 1)
        If Not fila_vacia Then
            Print #CTE_FIC_03_MAP, linea
        End If
    Next f
    Next p
    Close #CTE_FIC_03_MAP
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha grabado el mapa " & nombre_fichero_ejv
    Screen.MousePointer = CTE_DEFECTO

fin:

    'Como la pantalla de "Guardar Como.." habra borrado el mapa, lo muestro otra vez
    s_refrescar_mapa_actual_ma0

End Sub


Sub s_usr_guardar_jugadores_pri()

    Dim linea As String
    Dim texto As String
    Dim corte As Integer
    Dim exito_al_abrir As Boolean
    
    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_PRI)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Guardar Fichero de Tipos de Jugadores al Prisionero"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "Guardar"
    frm_z0_fic.File1.Pattern = "*.pri"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path-fichero
    frm_a4_tipospri.Caption = "Tipos de Jugadores " & nombre_fichero_ejv
    frm_a4_tipospri.FicJugActual.Caption = nombre_fichero_ejv
    
    'Las lineas estan separadas por vbCrLf
    'es decir, por Chr$(10) + Chr$(13)
    
    'Compruebo que no existe
    If f_existe_fichero(nombre_fichero_ejv) Then
        If MsgBox("El fichero ya existe. ¿Desea reemplazarlo?", vbQuestion + vbOKCancel) = vbCancel Then
            'Ha elegido cancelar la operación
            GoTo fin
        End If
    End If
    
    
    'Abro y leo
    Open nombre_fichero_ejv For Output As #CTE_FIC_04_PRI
    texto = frm_a4_tipospri.txt_tipos.Text
    Screen.MousePointer = CTE_ARENA
    While Len(texto) > 0
        corte = InStr(texto, Chr$(13))
        If corte = 0 Then
            linea = texto
            texto = ""
        Else
            'cojo la linea
            linea = Left(texto, corte - 1)
            'quito lo que ya he usado
            texto = Right(texto, Len(texto) - corte)
            'Quito un chr$13 que queda por ahi
            texto = Right(texto, Len(texto) - 1)
        End If
        Print #CTE_FIC_04_PRI, linea
    Wend
    Close #CTE_FIC_04_PRI
    fich_jug_modificado_pri = False
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha grabado el fichero de jugadores al prisionero " & nombre_fichero_ejv
    Screen.MousePointer = CTE_DEFECTO
    
fin:
    'No hago nada

End Sub


Sub s_usr_guardar_ejecucion_3r()
    
    Dim linea As String
    Dim cont_ag As Integer
    Dim cont_re As Integer
    
    MsgBox "En esta versión se pueden guardar y recuperar los agentes con sus reglas, pero no las opciones elegidas", vbInformation
        
    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_3R)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Guardar Fichero de 3R"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "Guardar"
    frm_z0_fic.File1.Pattern = "*.3r"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path-fichero
    frm_c0_ce.Caption = nombre_fichero_ejv
    
    'Lo cargo en la pantalla
    cont_ag = 0
    
    'Compruebo que no existe
    If f_existe_fichero(nombre_fichero_ejv) Then
        If MsgBox("El fichero ya existe. ¿Desea reemplazarlo?", vbQuestion + vbOKCancel) = vbCancel Then
            'Ha elegido cancelar la operación
            GoTo fin
        End If
    End If
    
    
    'Abro y leo
    Open nombre_fichero_ejv For Output As #CTE_FIC_05_3R
    'guardo el actual
    Screen.MousePointer = CTE_ARENA
    For cont_ag = 1 To numero_total_de_agentes_ejv
        linea = agente_3r(cont_ag)
        Print #CTE_FIC_05_3R, linea
        
        linea = Format(peso_agente_ce0(cont_ag), "0.00000000")
        Print #CTE_FIC_05_3R, linea
        
        linea = ""
        For cont_re = 1 To numero_de_reglas_por_agente_3r
            linea = linea & CStr(peso_regla_agente_3r(cont_ag, cont_re)) & ","
        Next cont_re
       'quitamos la coma final
        linea = Left(linea, Len(linea) - 1)
        Print #CTE_FIC_05_3R, linea
        
        linea = ""
        For cont_re = 1 To numero_de_reglas_por_agente_3r
            linea = linea & CStr(prioridad_regla_agente_3r(cont_ag, cont_re)) & ","
        Next cont_re
       'quitamos la coma final
        linea = Left(linea, Len(linea) - 1)
        Print #CTE_FIC_05_3R, linea

    Next cont_ag
    Close #CTE_FIC_05_3R
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha grabado el fichero de 3R " & nombre_fichero_ejv
    Screen.MousePointer = CTE_DEFECTO

fin:
    'No hago nada


End Sub

Sub s_usr_abrir_mapa_ma0()

On Error GoTo abrir_error

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_MAP)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Abrir Fichero de Mapa"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Abrir"
    frm_z0_fic.File1.Pattern = "*.map"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    
    'Muestro el nuevo path-fichero
    frm_a0_mapa.Caption = nombre_fichero_ejv
    frm_a0_mapa.Refresh
    
        
    'Indico cual es el mapa que hay que cargar y lo cargo
    mapa_actual_ma0 = nombre_fichero_ejv
    s_aut_leer_mapa_ma0
    s_refrescar_mapa_actual_ma0
    
    'Muestro las dimensiones del mapa
    habilitar_change_zoom_va0 = False
    frm_a0_mapa.Op_MapaMaxEjeZ = mapa_pisos_ma0
    frm_a0_mapa.Op_MapaMaxEjeY = mapa_filas_ma0
    frm_a0_mapa.Op_MapaMaxEjeX = mapa_columnas_ma0
    s_cargar_tipo_zoom_ma0
    habilitar_change_zoom_va0 = True
    
'=======================================================
    Exit Sub
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero " & nombre_fichero_ejv & " o es incorrecto"
    Close #CTE_FIC_03_MAP
    s_poner_un_mapa_cualquiera_ma0
    
    

End Sub
Sub s_poner_un_mapa_cualquiera_ma0()

    'Pongo un mapa cualquiera
    mapa_pisos_ma0 = 1
    mapa_filas_ma0 = 100
    mapa_columnas_ma0 = 100
    ReDim mapa_ma0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Boolean
    ReDim nodo_visitado_va0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Integer
    'Modo de presentacion
    ver_zoom_ma0 = CTE_ZOOM_PANORAMICA
    s_fijar_separacion_mapa_ma0
    'Celdas de obstaculos
    'Inicializo todas y asi si falta algo queda como vacio
    s_mapa_inicializar_va0 mapa_ma0, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0, False

End Sub
Sub s_usr_abrir_ejecucion_3r()


On Error GoTo abrir_error
    
    Dim linea As String
    Dim cont_ag As Integer
    Dim cont_re As Integer
    
    Dim s_dato As String
    Dim dato As Integer
    Dim primera_coma As Integer
    
    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_3R)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Abrir Fichero de 3R"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Abrir"
    frm_z0_fic.File1.Pattern = "*.3r"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path-fichero
    frm_c0_ce.Caption = "Tres en Raya " & nombre_fichero_ejv
    
    'Lo cargo en la pantalla
    cont_ag = 0
    Open nombre_fichero_ejv For Input As #CTE_FIC_05_3R
    Screen.MousePointer = CTE_ARENA
    While Not EOF(CTE_FIC_05_3R) And cont_ag <= numero_total_de_agentes_ejv
        cont_ag = cont_ag + 1
        
        linea = f_leer_linea(CTE_FIC_05_3R)
        If Len(linea) > 0 Then
            agente_3r(cont_ag) = linea
        End If
        
        linea = f_leer_linea(CTE_FIC_05_3R)
        If Len(linea) > 0 Then
            peso_agente_ce0(cont_ag) = CLng(linea)
        End If
    
        linea = f_leer_linea(CTE_FIC_05_3R)
        For cont_re = 1 To numero_de_reglas_por_agente_3r
            primera_coma = InStr(linea, ",")
            If primera_coma = 0 Then
                s_dato = linea
                If IsNumeric(s_dato) Then
                    dato = Int(s_dato)
                Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: no numérico"
                End If
                linea = ""
            Else
                s_dato = Left(linea, primera_coma - 1)
                If IsNumeric(s_dato) Then
                    dato = Int(s_dato)
                Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: no numérico"
                End If
                linea = Right(linea, Len(linea) - primera_coma)
            End If
            peso_regla_agente_3r(cont_ag, cont_re) = dato
        Next cont_re
    
    
        linea = f_leer_linea(CTE_FIC_05_3R)
        For cont_re = 1 To numero_de_reglas_por_agente_3r
            primera_coma = InStr(linea, ",")
            If primera_coma = 0 Then
                s_dato = linea
                If IsNumeric(s_dato) Then
                    dato = Int(s_dato)
                Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: no numérico"
                End If
                linea = ""
            Else
                s_dato = Left(linea, primera_coma - 1)
                If IsNumeric(s_dato) Then
                    dato = Int(s_dato)
                Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: no numérico"
                End If
                linea = Right(linea, Len(linea) - primera_coma)
            End If
            prioridad_regla_agente_3r(cont_ag, cont_re) = dato
        Next cont_re
    
    
    Wend
    Close #CTE_FIC_05_3R
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha leido el fichero de 3R " & nombre_fichero_ejv
    Screen.MousePointer = CTE_DEFECTO
    
    
'=======================================================
    Exit Sub
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero " & nombre_fichero_ejv & " o es incorrecto"
    Screen.MousePointer = CTE_DEFECTO
    Close #CTE_FIC_05_3R

End Sub


Sub s_usr_abrir_jugadores_pri()

    Dim exito_al_abrir As Boolean
    
    'Elijo path por defecto
    nombre_fichero_jugadores_pri = path_largo_ejv(CTE_C_PRG_PRI)
    nombre_fichero_ejv = nombre_fichero_jugadores_pri
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Abrir Fichero de Tipos de Jugadores al Prisionero"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Abrir"
    frm_z0_fic.File1.Pattern = "*.pri"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path-fichero
    nombre_fichero_jugadores_pri = nombre_fichero_ejv
    frm_a4_tipospri.Caption = "Tipos de Jugadores " & nombre_fichero_jugadores_pri
    frm_a4_tipospri.FicJugActual.Caption = nombre_fichero_jugadores_pri
    
    exito_al_abrir = s_aut_abrir_jugadores_pri(nombre_fichero_jugadores_pri)
    If exito_al_abrir Then
        s_analisis_sintactico_tipos_jugadores_pri
        s_mostrar_fichero_tipos_jugadores_pri
        fich_jug_modificado_pri = False
    End If

End Sub

Sub s_aut_leer_fichero_automatico_ejv(indice_auto As Long)

On Error GoTo abrir_error


    Dim linea As String
    
    'Lo cargo
    Open f_nombre_completo(path_largo_ejv(CTE_C_PRG_AUT), fichero_aut_ejv(indice_auto)) For Input As #CTE_FIC_02_AUT
    s_leer_parametros_fichero_aut_ejv indice_auto
    Close #CTE_FIC_02_AUT
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha leido el fichero automático " & fichero_aut_ejv(indice_auto)
    
    
    Exit Sub
abrir_error:
    
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "No se encuentra el fichero " & fichero_aut_ejv(indice_auto) & " o hay errores en su contenido"
    Close #CTE_FIC_02_AUT

End Sub


Sub s_aut_leer_inicio_txt()

On Error GoTo abrir_error

    'Inicializo a los valores por defecto, por si luego falla algo
    s_asignar_valores_por_defecto_variables_config_ejv

principio:
    'Lo cargo
    Open nombre_fichero_ejv For Input As #CTE_FIC_01_INICIO
    s_leer_parametros_inicio_txt
    Close #CTE_FIC_01_INICIO
    
    Exit Sub
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero " & CTE_nombreINICIO_TXT & " o hay errores en su contenido"
    Close #CTE_FIC_01_INICIO

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_RAIZ)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    If Not automatico_ejv Then
        tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
        frm_z0_fic.Caption = "Abrir Fichero de configuración"  'Esto provoca la llamada, igual que un show
        frm_z0_fic.Aceptar.Caption = "&Abrir"
        frm_z0_fic.File1.Pattern = "inicio.txt"
        frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
        frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        If cancelar_operacion_fichero_ejv Then Exit Sub
    End If
GoTo principio


End Sub

Sub s_aut_grabar_automatico_ejv(indice_auto As Long, por_defecto As Boolean)

    'Este codigo aun no es usado por ningun programa
    'Ya que los aut se generan previamente con la herramienta
    'generador de aut del menu, y no desde el propio programa

    Dim linea As String
    Dim fic As String

    If por_defecto Then
        '1: AUTOMATICO NUMERO PROGRAMA
        num_prg_activo_ejv = 0
        '2: AUTOMATICO NUMERO EJEMPLO
        num_ej_activo_ejv = 0
        '3: FICHERO RESULTADOS
        un_ej_fichero_resumen_xls_ejv = "resultados.xls"
    End If

    fic = f_nombre_completo(path_largo_ejv(CTE_C_PRG_AUT), fichero_aut_ejv(indice_auto))
    Open fic For Output As #CTE_FIC_02_AUT
    
    '0 Cabecera
    linea = "'Parametros Comunes de un Ejemplo Automático"
    Print #CTE_FIC_02_AUT, linea
    linea = "'==========================================="
    Print #CTE_FIC_02_AUT, linea
    
    '1: AUTOMATICO NUMERO PROGRAMA
    linea = "AUTOMATICO NUMERO PROGRAMA:"
    linea = linea & num_prg_activo_ejv
    Print #CTE_FIC_02_AUT, linea
        
    '2: AUTOMATICO NUMERO EJEMPLO
    linea = "AUTOMATICO NUMERO EJEMPLO:"
    linea = linea & num_ej_activo_ejv
    Print #CTE_FIC_02_AUT, linea
        
    '3: FICHERO RESULTADOS
    linea = "FICHERO RESULTADOS:"
    linea = linea & un_ej_fichero_resumen_xls_ejv
    Print #CTE_FIC_02_AUT, linea
    
    '0 Cabecera
    linea = ""
    Print #CTE_FIC_02_AUT, linea
    linea = "'Parametros Específicos del Ejemplo Automático 1"
    Print #CTE_FIC_02_AUT, linea
    linea = "'==============================================="
    Print #CTE_FIC_02_AUT, linea
    
    
    'Escribo la parte especifica
    '=======================
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
        Case CTE_PAL '2
        Case CTE_3R '3
        Case CTE_PRI '4
        Case CTE_CEL '5
        Case CTE_GAI '6
        Case CTE_EXP '7
        Case CTE_CAD '8
        Case CTE_PEZ '9
        Case CTE_UVA '10
        Case CTE_YXY '11
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
    
    Close #CTE_FIC_02_AUT
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha escrito el fichero automático " & fic


End Sub
Sub s_asignar_valores_por_defecto_variables_config_ejv()

    '1: VERSION
    version_ejv = App.Major & "." & App.Minor & "." & App.Revision
    '2: IDIOMA
    idioma_ejv = CTE_CASTELLANO
    '3: ELEGIR IDIOMA
    elegir_idioma_ejv = True
    '4: CONTROL DE ERRORES
    control_errores_de_programacion_ejv = False
    '5: MOSTRAR_LOGO
    mostrar_logo_ejv = True
    '6: ALGORITMO DE ORDENACION
    algoritmo_ordenacion_ejv = CTE_BURBUJA
    '7: SISTEMA_OPERATIVO
    sistema_operativo_ejv = CTE_WINDOWS95
    '8: PEDIR_CONFIRMACION
    pedir_confirmacion_ejv = True
    '9: RESOLUCION PANTALLA
    resolucion_pantalla_ejv = CTE_800X600OSUPERIOR
    '10: GRABAR CONFIGURACION
    grabar_configuracion_ejv = True
    '11: GRABAR CONFIG POR DEFECTO
    grabar_config_defecto_ejv = False
    '12: GRABAR LOG
    grabar_log_ejv = False
    '13: FICHERO LOG
    fichero_log_ejv = "resum.log"
    '14: GRABAR RESUMEN TXT
    grabar_resumen_txt_ejv = False
    '15: FICHERO RESUMEN TXT
    fichero_resumen_txt_ejv = "resum.txt"
    '16: GRABAR RESUMEN EXCEL
    grabar_resumen_xls_ejv = False
    '17: FICHERO RESUMEN EXCEL
    fichero_resumen_xls_ejv = "resum.xls"
    '18: REEMPLAZAR FICHEROS EXISTENTES
    reemplazar_fic_ejv = False
    '19: AUTOMATICO
    automatico_ejv = False
    '20: FICHERO AUTOMATICO
    num_ficheros_aut_ejv = 1
    ReDim fichero_aut_ejv(1 To num_ficheros_aut_ejv) As String
    fichero_aut_ejv(1) = CTE_NOHAY

End Sub
Sub s_aut_grabar_inicio_txt(por_defecto As Boolean)

    Dim linea As String
    Dim indice_auto As Long
    
    nombre_fichero_ejv = f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), CTE_nombreINICIO_TXT)

    If por_defecto Then
        s_asignar_valores_por_defecto_variables_config_ejv
    End If

    Open nombre_fichero_ejv For Output As #CTE_FIC_01_INICIO
    
    '0 Cabecera
    linea = "'Parametros de Configuración generales de Ejemplos de Vida"
    Print #CTE_FIC_01_INICIO, linea
    linea = "'========================================================="
    Print #CTE_FIC_01_INICIO, linea
    
    '1: VERSION
    linea = CTE_VERSION
    linea = linea & version_ejv
    Print #CTE_FIC_01_INICIO, linea
    
    '2: IDIOMA
    linea = CTE_IDIOMA
    Select Case idioma_ejv
        Case CTE_CASTELLANO
            linea = linea & CTEm_CASTELLANO
        Case CTE_INGLES
            linea = linea & CTEm_INGLES
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: valor no valido: "
    End Select
    Print #CTE_FIC_01_INICIO, linea
    
    '3: ELEGIR IDIOMA
    linea = CTE_ELEGIRIDIOMA
    If elegir_idioma_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
    
    '4: CONTROL DE ERRORES
    linea = CTE_CTRLERRORES
    If control_errores_de_programacion_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
    
    '5: MOSTRAR_LOGO
    linea = CTE_MOSTRAR_LOGO
    If mostrar_logo_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '6: ALGORITMO DE ORDENACION
    linea = CTE_ALGORITMODEORDENACION
    Select Case algoritmo_ordenacion_ejv
        Case CTE_BURBUJA
            linea = linea & CTEm_BURBUJA
        Case CTE_QUICKSORT
            linea = linea & CTEm_QUICKSORT
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: valor no valido: "
    End Select
    Print #CTE_FIC_01_INICIO, linea
    
    '7: CTE_SISTEMA_OPERATIVO
    linea = CTE_SISTEMA_OPERATIVO
    Select Case sistema_operativo_ejv
        Case CTE_WINDOWS95
            linea = linea & CTEm_WINDOWS95
        Case CTE_WINDOWSNT
            linea = linea & CTEm_WINDOWSNT
        Case CTE_WINDOWS3X
            linea = linea & CTEm_WINDOWS3X
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: valor no valido: "
    End Select
    Print #CTE_FIC_01_INICIO, linea
    
    '8: CTE_PEDIR_CONFIRMACION
    linea = CTE_PEDIR_CONFIRMACION
    If pedir_confirmacion_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
    
    '9: RESOLUCION PANTALLA
    linea = CTE_RESOLUCIONPANTALLA
    Select Case resolucion_pantalla_ejv
        Case CTE_640X480
            linea = linea & CTEm_640X480
        Case CTE_800X600OSUPERIOR
            linea = linea & CTEm_800X600OSUPERIOR
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: valor no valido: "
    End Select
    Print #CTE_FIC_01_INICIO, linea
    
    '10: GRABAR CONFIGURACION
    linea = CTE_GRABAR_CONFIGURACION
    If grabar_configuracion_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '11: GRABAR CONFIG POR DEFECTO
    linea = CTE_GRABAR_CONFIG_POR_DEFECTO
    If grabar_config_defecto_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '12: GRABAR LOG
    linea = CTE_GRABAR_LOG
    If grabar_log_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '13: FICHERO LOG
    linea = CTE_FICHERO_LOG
    'Si esta en path_largo_ejv(CTE_C_SAL_LOG) no pongo el path y asi es mas generico para futuros cambios de version
    If f_path_iguales(f_path_fichero(fichero_log_ejv, CTE_C_SAL_LOG), path_largo_ejv(CTE_C_SAL_LOG)) Then
        linea = linea & f_nombre_fichero(fichero_log_ejv)
    Else
        linea = linea & fichero_log_ejv
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '14: GRABAR RESUMEN TXT
    linea = CTE_GRABAR_RESUMEN_TXT
    If grabar_resumen_txt_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '15: FICHERO RESUMEN TXT
    linea = CTE_FICHERO_RESUMEN_TXT
    'Si esta en path_largo_ejv(CTE_C_SAL_TXT) no pongo el path y asi es mas generico para futuros cambios de version
    If f_path_iguales(f_path_fichero(fichero_resumen_txt_ejv, CTE_C_SAL_TXT), path_largo_ejv(CTE_C_SAL_TXT)) Then
        linea = linea & f_nombre_fichero(fichero_resumen_txt_ejv)
    Else
        linea = linea & fichero_resumen_txt_ejv
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '16: GRABAR RESUMEN EXCEL
    linea = CTE_GRABAR_RESUMEN_EXCEL
    If grabar_resumen_xls_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '17: FICHERO RESUMEN EXCEL
    linea = CTE_FICHERO_RESUMEN_EXCEL
    'Si esta en path_largo_ejv(CTE_C_SAL_XLS) no pongo el path y asi es mas generico para futuros cambios de version
    If f_path_iguales(f_path_fichero(fichero_resumen_xls_ejv, CTE_C_SAL_XLS), path_largo_ejv(CTE_C_SAL_XLS)) Then
        linea = linea & f_nombre_fichero(fichero_resumen_xls_ejv)
    Else
        linea = linea & fichero_resumen_xls_ejv
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '18: REEMPLAZAR FICHEROS EXISTENTES
    linea = "REEMPLAZAR FICHEROS EXISTENTES="
    If reemplazar_fic_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
        
    '19: AUTOMATICO
    linea = "AUTOMATICO="
    If automatico_ejv Then
        linea = linea & CTE_txtTRUE
    Else
        linea = linea & CTE_txtFALSE
    End If
    Print #CTE_FIC_01_INICIO, linea
    
    '20: FICHERO AUTOMATICO
    For indice_auto = 1 To num_ficheros_aut_ejv
        linea = "FICHERO AUTOMATICO="
        If fichero_aut_ejv(indice_auto) = "" Or fichero_aut_ejv(indice_auto) = CTE_NOHAY Then
            fichero_aut_ejv(indice_auto) = CTEm_NOHAY
        End If
        linea = linea & fichero_aut_ejv(indice_auto)
        Print #CTE_FIC_01_INICIO, linea
    Next indice_auto
    
    Close #CTE_FIC_01_INICIO
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha escrito el fichero de configuración global " & nombre_fichero_ejv

End Sub


Function s_aut_abrir_jugadores_pri(fic_jug_pri_completo As String) As Boolean

On Error GoTo abrir_error

    Dim linea As String
    Dim n_lin As Integer
    
    'Inicializo
    ReDim fichero_tipos_jugadores_pri(1 To 1) As String
    
    Screen.MousePointer = CTE_ARENA
    'Lo cargo
    n_lin = 0
    Open fic_jug_pri_completo For Input As #CTE_FIC_04_PRI
    While Not EOF(CTE_FIC_04_PRI)
        linea = f_leer_linea(CTE_FIC_04_PRI)
        linea = UCase(linea)
        If Len(linea) > 0 Then
            n_lin = n_lin + 1
            ReDim Preserve fichero_tipos_jugadores_pri(1 To n_lin) As String
            fichero_tipos_jugadores_pri(n_lin) = linea
        End If
    Wend
    Close #CTE_FIC_04_PRI
    s_aut_abrir_jugadores_pri = True
    Screen.MousePointer = CTE_DEFECTO
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha leido el fichero de jugadores al prisionero " & fic_jug_pri_completo
   
'=======================================================
    Exit Function
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero " & fic_jug_pri_completo & " o es incorrecto"
    Screen.MousePointer = CTE_DEFECTO
    Close #CTE_FIC_04_PRI
    s_aut_abrir_jugadores_pri = False


End Function


Sub s_aut_leer_mapa_ma0()
    
On Error GoTo abrir_error
    
    'Abre el mapa indicado por mapa_actual_ma0 y lo carga en ma0
    
    Dim p As Double
    Dim f As Double
    Dim c As Double
    
    Dim linea As String
    Dim dato As Integer
    Dim s_dato As String
    Dim primera_coma As Integer
    
    mapa_sin_obstaculos_ma0 = True

    'Lo cargo en la pantalla
    f = 0
    Open mapa_actual_ma0 For Input As #CTE_FIC_03_MAP
    Screen.MousePointer = CTE_ARENA
    'Pisos
    linea = f_leer_linea(CTE_FIC_03_MAP)
    mapa_pisos_ma0 = CInt(Trim(Right(linea, Len(linea) - Len("PISOS="))))
    'Filas
    linea = f_leer_linea(CTE_FIC_03_MAP)
    mapa_filas_ma0 = CInt(Trim(Right(linea, Len(linea) - Len("FILAS="))))
    'Columnas
    linea = f_leer_linea(CTE_FIC_03_MAP)
    mapa_columnas_ma0 = CInt(Trim(Right(linea, Len(linea) - Len("COLUMNAS="))))
    
    ReDim mapa_ma0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Boolean
    ReDim nodo_visitado_va0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Integer
    
    'Modo de presentacion
    linea = f_leer_linea(CTE_FIC_03_MAP)
    Select Case UCase(Trim(Right(linea, Len(linea) - Len("ZOOM="))))
        Case CTE_tZOOM_DETALLE
            ver_zoom_ma0 = CTE_ZOOM_DETALLE
        Case CTE_tZOOM_PANORAMICA
            ver_zoom_ma0 = CTE_ZOOM_PANORAMICA
        Case CTE_tZOOM_PIXELS
            ver_zoom_ma0 = CTE_ZOOM_PIXELS
        Case CTE_tZOOM_3D
            ver_zoom_ma0 = CTE_ZOOM_3D
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
    End Select
    s_cargar_tipo_zoom_ma0
    s_fijar_separacion_mapa_ma0
    
    
    
    'Celdas de obstaculos
    'Inicializo todas y asi si falta algo queda como vacio
    s_mapa_inicializar_va0 mapa_ma0, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0, False
    'Leo el mapa
    p = 0
    f = 0
    While Not EOF(CTE_FIC_03_MAP) And f < mapa_filas_ma0
        If f = 0 Then
            p = p + 1
        End If
        f = f + 1
        linea = f_leer_linea(CTE_FIC_03_MAP)
        'Leo el piso, si pone piso
        If Left(linea, Len("PISO=")) = "PISO=" Then
            p = CInt(Trim(Right(linea, Len(linea) - Len("PISO="))))
            linea = ""
            If Not EOF(CTE_FIC_03_MAP) Then
                linea = f_leer_linea(CTE_FIC_03_MAP)
            End If
        End If
        If Len(linea) > 0 Then
            For c = 1 To mapa_columnas_ma0
                primera_coma = InStr(linea, ",")
                If primera_coma = 0 Then
                    s_dato = linea
                    If IsNumeric(s_dato) Then
                        dato = Int(s_dato)
                    Else
                        dato = 0
                    End If
                    linea = ""
                Else
                    s_dato = Left(linea, primera_coma - 1)
                    If IsNumeric(s_dato) Then
                        dato = Int(s_dato)
                    Else
                        dato = 0
                    End If
                    linea = Right(linea, Len(linea) - primera_coma)
                End If
                If dato = 1 Then
                    mapa_ma0(p, f, c) = True
                    mapa_sin_obstaculos_ma0 = False
                Else
                    mapa_ma0(p, f, c) = False
                End If
            Next c
        End If
        If f = mapa_filas_ma0 Then
            f = 0
        End If
    Wend
    Close #CTE_FIC_03_MAP
    Screen.MousePointer = CTE_DEFECTO
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha leido el fichero de mapa " & mapa_actual_ma0

'=======================================================
    Exit Sub
abrir_error:
    If automatico_ejv Then
        s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "No se encuentra el fichero " & mapa_actual_ma0 & " o es incorrecto"
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero " & mapa_actual_ma0 & " o es incorrecto"
    End If
    Screen.MousePointer = CTE_DEFECTO
    Close #CTE_FIC_03_MAP
    s_poner_un_mapa_cualquiera_ma0

End Sub

Sub s_accion_ficheros_va0(accion As Integer)

    Select Case frm_z0_mdi.ActiveForm.Name
        Case "frm_a0_mapa"
            Select Case accion
                Case CTE_FIC_ABRIR
                    s_usr_abrir_mapa_ma0
                Case CTE_FIC_GUARDAR
                    nombre_fichero_ejv = f_nombre_fichero(frm_a0_mapa.Caption)
                    s_usr_guardar_mapa_ma0
                Case CTE_FIC_GUARDARCOMO
                    nombre_fichero_ejv = f_nombre_fichero(frm_a0_mapa.Caption)
                    s_usr_guardar_mapa_ma0
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "La Accion de Fichero no existe"
            End Select
        Case "frm_c0_ce"
            Select Case accion
                Case CTE_FIC_ABRIR
                    s_usr_abrir_ejecucion_3r
                Case CTE_FIC_GUARDAR
                    'frm_z0_fic.n_fichero = f_nombre_fichero(nombre_fichero_ejv)
                    'esto no lo puedo poner porque provoca el load antes de tiempo
                    s_usr_guardar_ejecucion_3r
                Case CTE_FIC_GUARDARCOMO
                    s_usr_guardar_ejecucion_3r
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "La Accion de Fichero no existe"
            End Select
        Case "frm_a4_tipospri"
            Select Case accion
                Case CTE_FIC_ABRIR
                    s_usr_abrir_jugadores_pri
                Case CTE_FIC_GUARDAR
                    'frm_z0_fic.n_fichero = f_nombre_fichero(nombre_fichero_jugadores_pri)
                    'esto no lo puedo poner porque provoca el load antes de tiempo
                    s_usr_guardar_jugadores_pri
                Case CTE_FIC_GUARDARCOMO
                    s_usr_guardar_jugadores_pri
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "La Accion de Fichero no existe"
            End Select
'        Case "frm_z0_graf"
'            Select Case accion
'                Case CTE_FIC_ABRIR
'                    frm_z0_graf.s_usr_abrir_graf_ejv
'                Case CTE_FIC_GUARDAR
'                    frm_z0_fic.n_fichero = f_nombre_fichero(nombre_fichero_ejv)
'                    s_usr_guardar_graf_ejv
'                Case CTE_FIC_GUARDARCOMO
'                    s_usr_guardar_graf_ejv
'                Case Else
'                    s_error_ejv  CON_OPCION_FINALIZAR, "La Accion de Fichero no existe"
'            End Select
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "El formulario activo no admite operaciones con ficheros."
    End Select

End Sub


Function f_aut_leer_diccionario(cargar_todo As Boolean) As Long
    
    
'Aqui no debe haber ninguna referencia a frm_b2_pal.
'porque tb se usa en los graficos
    
On Error GoTo abrir_error

Dim linea As String
Dim indice As Long
Dim fic As String

principio:
    
    Screen.MousePointer = CTE_ARENA
    'Lo cargo
    fic = f_nombre_completo(dicc_carpeta_pal, dicc_fichero_pal)
    
    If cargar_todo Then
        numero_palabras_dicc_pal = f_numero_de_lineas_fic(fic, CTE_FIC_07_DICC)
    End If
    Open fic For Input As #CTE_FIC_07_DICC

    ReDim palabra_del_diccionario(1 To 1) As String
    
    indice = 0
    While Not EOF(CTE_FIC_07_DICC) And indice < numero_palabras_dicc_pal
        linea = f_leer_linea(CTE_FIC_07_DICC)
        indice = indice + 1
        ReDim Preserve palabra_del_diccionario(1 To indice) As String
        palabra_del_diccionario(indice) = Trim(linea)
    Wend
    Close #CTE_FIC_07_DICC
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha leido el fichero de diccionario " & fic
    Screen.MousePointer = CTE_DEFECTO
    
    'Control de errores de usuario
    If indice < numero_palabras_dicc_pal Then
        numero_palabras_dicc_pal = indice
        If Not automatico_ejv Then
            MsgBox "Solo se hay podido cargar " & numero_palabras_dicc_pal & " palabras", vbInformation
        End If
    End If

    f_aut_leer_diccionario = indice
    
    Exit Function
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero de diccionario o hay errores en su contenido"
    Close #CTE_FIC_07_DICC
    Screen.MousePointer = CTE_DEFECTO

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_ENT_DIC)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    If Not automatico_ejv Then
        tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
        frm_z0_fic.Caption = "Abrir Fichero de diccionario"  'Esto provoca la llamada, igual que un show
        frm_z0_fic.Aceptar.Caption = "&Abrir"
        frm_z0_fic.File1.Pattern = "*.dic"
        frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
        frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        'Como al princioio de esta funcion leo de 2 variables y no de nombre_fichero_ejv, tengo que pasar los nombres
        dicc_carpeta_pal = f_path_fichero(nombre_fichero_ejv, CTE_C_ENT_DIC)
        dicc_fichero_pal = f_nombre_fichero(nombre_fichero_ejv)
        If cancelar_operacion_fichero_ejv Then Exit Function
    End If
GoTo principio

End Function

Sub s_desencriptar(Fic_Enc As String, Fic_Des As String, s_clave As String)

Dim i As Long
Dim tama_bytes As Long
Dim pos_actual As Long
Dim arr_byte() As Byte
Dim byte_clave As Byte

Screen.MousePointer = CTE_ARENA

byte_clave = CByte(s_clave)


On Error GoTo abrir_error

    Open Fic_Enc For Binary Access Read As #CTE_FIC_15_ENC
    Open Fic_Des For Binary Access Write As #CTE_FIC_14_DES
    tama_bytes = LOF(CTE_FIC_15_ENC)
    ReDim arr_byte(1 To tama_bytes) As Byte
    For i = 1 To tama_bytes
        Get #CTE_FIC_15_ENC, , arr_byte(i)
        Put #CTE_FIC_14_DES, , mi_xor(arr_byte(i), byte_clave)
        pos_actual = Loc(CTE_FIC_15_ENC)
        If i Mod 100 = 0 Then
            frm_u0_encr.Porcentaje.Caption = Format((pos_actual / tama_bytes) * 100, "0.00") & "%"
            DoEvents
        End If
    Next i
    Close #CTE_FIC_15_ENC
    Close #CTE_FIC_14_DES

    frm_u0_encr.Porcentaje.Caption = "100%"
    Screen.MousePointer = CTE_DEFECTO
    MsgBox "Fichero Generado", vbInformation

    Exit Sub
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero binario"
    Close #CTE_FIC_15_ENC
    Close #CTE_FIC_14_DES
    Screen.MousePointer = CTE_DEFECTO


End Sub

Sub s_encriptar(Fic_Des As String, Fic_Enc As String, s_clave As String)

Dim i As Long
Dim tama_bytes As Long
Dim pos_actual As Long
Dim arr_byte() As Byte
Dim byte_clave As Byte

Screen.MousePointer = CTE_ARENA

byte_clave = CByte(s_clave)


'On Error GoTo abrir_error

    Open Fic_Des For Binary Access Read As #CTE_FIC_14_DES
    Open Fic_Enc For Binary Access Write As #CTE_FIC_15_ENC
    tama_bytes = LOF(CTE_FIC_14_DES)
    ReDim arr_byte(1 To tama_bytes) As Byte
    For i = 1 To tama_bytes
        Get #CTE_FIC_14_DES, , arr_byte(i)
        Put #CTE_FIC_15_ENC, , mi_xor(arr_byte(i), byte_clave)
        pos_actual = Loc(CTE_FIC_14_DES)
        If i Mod 100 = 0 Then
            frm_u0_encr.Porcentaje.Caption = Format((pos_actual / tama_bytes) * 100, "0.00") & "%"
            frm_u0_encr.TiempoEstimado = ""
            DoEvents
        End If
    Next i
    Close #CTE_FIC_14_DES
    Close #CTE_FIC_15_ENC

    frm_u0_encr.Porcentaje.Caption = "100%"
    Screen.MousePointer = CTE_DEFECTO
    MsgBox "Fichero Generado", vbInformation

    Exit Sub
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero binario"
    Close #CTE_FIC_14_DES
    Close #CTE_FIC_15_ENC
    Screen.MousePointer = CTE_DEFECTO


End Sub

Sub s_aut_leer_bkup()

On Error GoTo abrir_error

Dim linea As String
Dim indice As Long

principio:
    
    Screen.MousePointer = CTE_ARENA
    'Lo cargo
    Open nombre_fichero_ejv For Input As #CTE_FIC_08_BK_IN

    ReDim directorio_a_comp_bk(1 To 1) As String
    ReDim nombres_ficheros_backup_bk(1 To 1) As String

    indice = 0
    While Not EOF(CTE_FIC_08_BK_IN)
        linea = f_leer_linea(CTE_FIC_08_BK_IN)
        If Len(linea) > 0 Then
            indice = indice + 1
            ReDim Preserve directorio_a_comp_bk(1 To indice) As String
            ReDim Preserve nombres_ficheros_backup_bk(1 To indice) As String
            directorio_a_comp_bk(indice) = f_leer_campo(1, linea)
            nombres_ficheros_backup_bk(indice) = f_leer_campo(2, linea)
        End If
    Wend
    Close #CTE_FIC_08_BK_IN
    
    Screen.MousePointer = CTE_DEFECTO
    
    Exit Sub
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero de carpetas para hacer backup o hay errores en su contenido"
    Close #CTE_FIC_08_BK_IN
    Screen.MousePointer = CTE_DEFECTO

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_UTIL)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    If Not automatico_ejv Then
        tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
        frm_z0_fic.Caption = "Fichero de backup"  'Esto provoca la llamada, igual que un show
        frm_z0_fic.Aceptar.Caption = "&Abrir"
        frm_z0_fic.File1.Pattern = "*.txt"
        frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
        frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        If cancelar_operacion_fichero_ejv Then Exit Sub
    End If
GoTo principio


End Sub

Sub s_aut_leer_arbol()

On Error GoTo abrir_error

Dim linea As String
Dim indice As Long

principio:
    
    Screen.MousePointer = CTE_ARENA
    'Lo cargo
    Open nombre_fichero_ejv For Input As #CTE_FIC_16_ARB

    ReDim cod_arb(1 To 1) As String
    ReDim desc_arb(1 To 1) As String
    ReDim cod_padre_arb(1 To 1) As String
    indice = 0
    While Not EOF(CTE_FIC_16_ARB)
        linea = f_leer_linea(CTE_FIC_16_ARB)
        If Len(linea) > 0 Then
            indice = indice + 1
            ReDim Preserve cod_arb(1 To indice) As String
            ReDim Preserve desc_arb(1 To indice) As String
            ReDim Preserve cod_padre_arb(1 To indice) As String
            cod_arb(indice) = f_leer_campo(1, linea)
            desc_arb(indice) = f_leer_campo(2, linea)
            cod_padre_arb(indice) = f_leer_campo(3, linea)
        End If
    Wend
    Close #CTE_FIC_16_ARB
    Screen.MousePointer = CTE_DEFECTO
    
    Exit Sub
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero de árbol o hay errores en su contenido"
    Close #CTE_FIC_16_ARB
    Screen.MousePointer = CTE_DEFECTO

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_UTIL)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    If Not automatico_ejv Then
        tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
        frm_z0_fic.Caption = "Fichero de árbol"  'Esto provoca la llamada, igual que un show
        frm_z0_fic.Aceptar.Caption = "&Abrir"
        frm_z0_fic.File1.Pattern = "*.txt"
        frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
        frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        If cancelar_operacion_fichero_ejv Then Exit Sub
    End If
GoTo principio

End Sub


Sub s_aut_leer_digitos_azar()

On Error GoTo abrir_error

Dim linea As String
Dim indice As Long
Dim fic As String

principio:
    
    Screen.MousePointer = CTE_ARENA
    'Lo cargo
    fic = f_nombre_completo(azar_carpeta_ejv, azar_fichero_ejv)
    Open fic For Input As #CTE_FIC_06_AZAR

    ReDim digitos_azar(1 To CTE_numero_maximo_pi) As Integer
    indice = 1
    While Not EOF(CTE_FIC_06_AZAR)
        linea = f_leer_linea(CTE_FIC_06_AZAR)
        While Len(linea) > 0
            digitos_azar(indice) = Left(linea, 1)
            linea = Right(linea, Len(linea) - 1)
            indice = indice + 1
        Wend
    Wend
    azar_en_memoria_ejv = True
    Close #CTE_FIC_06_AZAR
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha leido el fichero de azar " & fic
    Screen.MousePointer = CTE_DEFECTO
    
    Exit Sub
abrir_error:
    s_error_ejv CON_OPCION_FINALIZAR, "No se encuentra el fichero de azar o hay errores en su contenido"
    Close #CTE_FIC_06_AZAR
    Screen.MousePointer = CTE_DEFECTO

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_ENT_RAN)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    If Not automatico_ejv Then
        tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
        frm_z0_fic.Caption = "Abrir Fichero de azar"  'Esto provoca la llamada, igual que un show
        frm_z0_fic.Aceptar.Caption = "&Abrir"
        frm_z0_fic.File1.Pattern = "*.ran"
        frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
        frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        'Como al princioio de esta funcion leo de 2 variables y no de nombre_fichero_ejv, tengo que pasar los nombres
        azar_carpeta_ejv = f_path_fichero(nombre_fichero_ejv, CTE_C_ENT_RAN)
        azar_fichero_ejv = f_nombre_fichero(nombre_fichero_ejv)
        If cancelar_operacion_fichero_ejv Then Exit Sub
    End If
GoTo principio

End Sub
Function f_posicionarse_en_seccion(seccion As String, num_fic As Integer, nom_fic_con_path As String) As Boolean

On Error Resume Next

    Dim linea As String

    'Lo cierro (si estaba abierto) y lo abro para posicionarme al principio
    Close #num_fic
    Open nom_fic_con_path For Input As #num_fic
    
    linea = ""
    While linea <> seccion And Not EOF(num_fic)
        linea = f_leer_linea(num_fic)
    Wend


End Function
Function f_cargar_seccion_actual(mi_array() As String, num_fic As Integer, nom_fic_con_path As String) As Boolean

    Dim linea As String
    Dim indice As Long

    ReDim Preserve mi_array(1 To 1) As String
    indice = 0
    
    linea = ""
    linea = f_leer_linea(num_fic)
    While Left(linea, 1) <> "[" And Not EOF(num_fic)
        If linea <> "" And Left(linea, 1) <> "[" Then
            indice = indice + 1
            ReDim Preserve mi_array(1 To indice) As String
            mi_array(indice) = linea
        End If
        linea = f_leer_linea(num_fic)
    Wend

    'Lo cierro y abro para posicionarme al principio otra vez
    Close #num_fic
    Open nom_fic_con_path For Input As #num_fic

End Function

Function f_nombre_completo_fichero_no_existente(ByVal fic As String, tipo As Integer) As String

    Dim path_f As String
    Dim nombre_mas_extension As String
    Dim nombre_sin_extension As String
    Dim extension_con_punto As String
    Dim cont As Long
    Dim s_cont As String

    'Obligo a que tenga un path
    fic = f_nombre_completo_obligatiorio_ejv(fic, tipo)
    
    If f_existe_fichero(fic) Then
        'Rompo
        path_f = f_path_fichero(fic, tipo)
        nombre_mas_extension = f_nombre_fichero(fic)
        nombre_sin_extension = f_nombre_sin_extension(nombre_mas_extension)
        extension_con_punto = f_extension_con_punto(nombre_mas_extension)
        'Añado el _00000001
        cont = 1
        s_cont = "_" & f_ceros_izquierda(CStr(cont), 8)
        'Incremento hasta que no exista
        While cont <= 99999999 And f_existe_fichero(f_nombre_completo(path_f, nombre_sin_extension & s_cont & extension_con_punto))
            cont = cont + 1
            s_cont = "_" & f_ceros_izquierda(CStr(cont), 8)
        Wend
        If cont > 99999999 Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error con el fichero " & path_f & nombre_sin_extension & s_cont & extension_con_punto
        Else
            f_nombre_completo_fichero_no_existente = f_nombre_completo(path_f, nombre_sin_extension & s_cont & extension_con_punto)
        End If
    Else
        f_nombre_completo_fichero_no_existente = fic
    End If

End Function
Function f_existe_fichero(fic As String) As Boolean

    On Error GoTo no_existe
    Open fic For Input As #CTE_FIC_99_EXISTE
    'Si no existe se produce un error en este punto
    
    'Si sigue por aqui es que si existe
    Close #CTE_FIC_99_EXISTE
    f_existe_fichero = True
    Exit Function
   
no_existe:
    f_existe_fichero = False
    Exit Function


End Function
Function restar_path_ejv(path1 As String, path2 As String) As String

    Dim temp1 As String
    Dim temp2 As String

    temp1 = UCase(path1)
    temp2 = UCase(path2)
    
    If temp1 = temp2 Then
        restar_path_ejv = ""
    Else
        While Len(temp2) > 0 And Left(temp1, 1) = Left(temp2, 1)
            temp1 = Right(temp1, Len(temp1) - 1)
            temp2 = Right(temp2, Len(temp2) - 1)
        Wend
        If Right(temp1, 1) = "\" Then
            temp1 = Left(temp1, Len(temp1) - 1)
        End If
        restar_path_ejv = temp1
    End If

End Function
Function f_path_iguales(path1 As String, path2 As String) As Boolean

    While Right(path1, 1) = "\" Or Right(path1, 1) = "/"
        path1 = Left(path1, Len(path1) - 1)
    Wend

    While Right(path2, 1) = "\" Or Right(path2, 1) = "/"
        path2 = Left(path2, Len(path2) - 1)
    Wend

    If path1 = path2 Then
        f_path_iguales = True
    Else
        f_path_iguales = False
    End If

End Function
Function f_extension_con_punto(nombre As String) As String

    Dim posicion_punto As Integer

    posicion_punto = InStr(nombre, ".")
    If posicion_punto = 0 Then
        'No hay extension
        f_extension_con_punto = ""
    Else
        f_extension_con_punto = Right(nombre, Len(nombre) - posicion_punto + 1)
    End If
    

End Function

Function f_nombre_sin_extension(nombre As String) As String

    Dim posicion_punto As Integer

    posicion_punto = InStr(nombre, ".")
    If posicion_punto = 0 Then
        'No hay extension
        f_nombre_sin_extension = nombre
    Else
        f_nombre_sin_extension = Left(nombre, posicion_punto - 1)
    End If
    

End Function
Function f_nombre_fichero(path_mas_nombre As String) As String

    Dim pos As Integer
    Dim tmp As String
    Dim cuenta_quitado As Integer
    
    cuenta_quitado = 0
    'Busco la posicion de la barra mas a la derecha
    tmp = path_mas_nombre
    'Mientras haya barras
    While InStr(tmp, "\") <> 0 Or InStr(tmp, "/") <> 0
        pos = InStr(tmp, "\")
        If InStr(tmp, "/") > pos Then
            pos = InStr(tmp, "/")
        End If
        'quito esa parte
        tmp = Right(tmp, Len(tmp) - pos)
        cuenta_quitado = cuenta_quitado + pos
    Wend
    
    'Ya se cual es la ultima barra
    f_nombre_fichero = Right(path_mas_nombre, Len(path_mas_nombre) - cuenta_quitado)


End Function

Function f_path_fichero(path_completo As String, tipo As Integer) As String

    Dim tmp As String
        
    If InStr(path_completo, "/") = 0 And InStr(path_completo, "\") = 0 Then
        'Es un nombre sin path
        f_path_fichero = path_largo_ejv(tipo) & "\"
    Else
        tmp = path_completo
        While Right(tmp, 1) <> "\" And Right(tmp, 1) <> "/"
            tmp = Left(tmp, Len(tmp) - 1)
        Wend
        tmp = Left(tmp, Len(tmp) - 1)
        f_path_fichero = tmp & "\"
    End If
    


End Function

Function f_nombre_completo_obligatiorio_ejv(referencia_fic As String, tipo As Integer) As String

    'Si solo viene el nombre y no el path completo, le añado
    'el path por defecto de ese tipo
    If f_nombre_fichero(referencia_fic) = referencia_fic Then
        If referencia_fic = "" Or referencia_fic = CTE_NOHAY Or referencia_fic = CTEm_NOHAY Then
            'Si no tiene referencia, le pongo el general
            f_nombre_completo_obligatiorio_ejv = f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), referencia_fic)
        Else
            'Si tiene referencia, le pongo su carpeta
            f_nombre_completo_obligatiorio_ejv = f_nombre_completo(path_largo_ejv(tipo), referencia_fic)
        End If
    Else
        'Ya tiene un path completo
        f_nombre_completo_obligatiorio_ejv = referencia_fic
    End If

End Function
Function f_unir_path(p_path1 As String, p_path2 As String) As String

    p_path1 = f_sustituir_subcadena(p_path1, "/", "\")
    p_path2 = f_sustituir_subcadena(p_path2, "/", "\")

    If Right(p_path1, 1) = "\" Or Right(p_path1, 1) = "/" Then
        If Left(p_path2, 1) = "\" Or Left(p_path2, 1) = "/" Then
            f_unir_path = p_path1 & Right(p_path2, Len(p_path2) - 1)
        Else
            f_unir_path = p_path1 & p_path2
        End If
    Else
        If Left(p_path2, 1) = "\" Or Left(p_path2, 1) = "/" Then
            f_unir_path = p_path1 & p_path2
        Else
            f_unir_path = p_path1 & "\" & p_path2
        End If
    End If

End Function

Function f_nombre_completo_existente(p_path As String, p_fichero As String) As String

    'Si en p_fichero viene el path completo y no solo el nombre,
    'uso el path completo
    If f_nombre_fichero(p_fichero) <> p_fichero Then
        f_nombre_completo_existente = p_fichero
    Else
        If Right(p_path, 1) = "\" Or Right(p_path, 1) = "/" Then
            f_nombre_completo_existente = p_path & p_fichero
        Else
            f_nombre_completo_existente = p_path & "\" & p_fichero
        End If
    End If
    
    'Compruebo que el fichero que devuelvo existe
    If Not f_existe_fichero(f_nombre_completo_existente) Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error: No existe el fichero " & f_nombre_completo_existente
    End If

End Function

Function f_nombre_completo(p_path As String, p_fichero As String) As String
    
    'Si en p_fichero viene el path completo y no solo el nombre,
    'uso el path completo
    If f_nombre_fichero(p_fichero) <> p_fichero Then
        f_nombre_completo = p_fichero
    Else
        If Right(p_path, 1) = "\" Or Right(p_path, 1) = "/" Then
            f_nombre_completo = p_path & p_fichero
        Else
            f_nombre_completo = p_path & "\" & p_fichero
        End If
    End If
    
End Function
Function f_leer_campo(orden As Integer, ByVal linea As String) As String

Dim cont_campos As Integer
Dim campo As String
Dim hay_comillas As Boolean

linea = Trim(linea)
'linea = f_quitar_comillas_dobles(linea)
If Len(linea) < 1 Then
    f_leer_campo = ""
Else
    cont_campos = 0
    While cont_campos < orden
        'Leo la primera " si existe
        hay_comillas = False
        If Left(linea, 1) = """" Then
            linea = Right(linea, Len(linea) - 1)
            hay_comillas = True 'cada campo esta rodeado por comillas
        End If
        'Leo el campo
        campo = ""
        While Left(linea, 1) <> ";" And Len(linea) > 0
            'ignoro las " dentro de los campos
            If Left(linea, 1) <> """" Then
                campo = campo & Left(linea, 1)
            End If
            linea = Right(linea, Len(linea) - 1)
        Wend
        'Quito el ;
        If Left(linea, 1) = ";" Then
            linea = Right(linea, Len(linea) - 1)
        Else
            If Len(linea) > 0 Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End If
        End If
        cont_campos = cont_campos + 1 'ya tengo un campo
        'quito la comillas final del campo si existe y si se abrio con comillas
        If hay_comillas And Right(campo, 1) = """" Then
            campo = Left(campo, Len(campo) - 1)
        End If
    Wend
    f_leer_campo = campo
End If

End Function


Function f_leer_linea(num_fic As Integer) As String

On Error Resume Next
    'Si es fin de fichero o hay error, devuelve ""
    'Esta funcion deberia usarse en todos los ficheros, y no haber
    'mas Line Input que este en todo el proyecto
    Dim linea As String

    linea = ""
    'Salto todos los comentarios y lineas vacias
    While (Len(linea) = 0 Or Left(linea, 1) = "'") And Not EOF(num_fic)
        Line Input #num_fic, linea
        'Devuelvo la linea sin espacios por delante y detrás
        linea = Trim(linea)
    Wend
    f_leer_linea = linea
    
End Function


Sub s_abrir_fichero_salida_ejv(fic As Integer, modo As Integer)
'======================================================================
'Esta funcion se puede llamar aun cuando no se va a usar cierto fichero
'y en ese caso se sale de la función sin hacer nada
'======================================================================
'On Error Resume Next

    Dim linea As String


    'Si existe y no hay que machacar pero tampoco anexar
    If Not reemplazar_fic_ejv And modo <> CTE_ABRIR_ANEXAR Then
        modo = CTE_ABRIR_NOBORRAR
    End If

    Select Case fic
        Case CTE_FIC_20_GLOLOG
            If grabar_log_ejv Then 'Solo si hay que tratar este fichero
                Select Case modo
                    Case CTE_ABRIR_BORRAR
                        Open fichero_log_ejv For Output As #fic
                    Case CTE_ABRIR_ANEXAR
                        Open fichero_log_ejv For Append As #fic
                    Case CTE_ABRIR_NOBORRAR
                        fichero_log_ejv = f_nombre_completo_fichero_no_existente(fichero_log_ejv, CTE_C_SAL_LOG)
                        Open fichero_log_ejv For Output As #fic
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error"
                End Select
                linea = "Fichero de Eventos (.log)"
                Print #fic, linea
                linea = "========================="
                Print #fic, linea
                linea = ""
                Print #fic, linea
            End If
        Case CTE_FIC_21_GLOTXT
            If grabar_resumen_txt_ejv Then 'Solo si hay que tratar este fichero
                Select Case modo
                    Case CTE_ABRIR_BORRAR
                        Open fichero_resumen_txt_ejv For Output As #fic
                    Case CTE_ABRIR_ANEXAR
                        Open fichero_resumen_txt_ejv For Append As #fic
                    Case CTE_ABRIR_NOBORRAR
                        fichero_resumen_txt_ejv = f_nombre_completo_fichero_no_existente(fichero_resumen_txt_ejv, CTE_C_SAL_TXT)
                        Open fichero_resumen_txt_ejv For Output As #fic
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error"
                End Select
                If cabeceras_ejv Then
                    linea = "Fichero Resumen (.txt)"
                    Print #fic, linea
                    linea = "======================"
                    Print #fic, linea
                    linea = ""
                    Print #fic, linea
                End If
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Fichero resumen (.txt) abierto con exito " & fichero_resumen_txt_ejv
            End If
        Case CTE_FIC_22_GLOXLS
            If grabar_resumen_xls_ejv Then 'Solo si hay que tratar este fichero
                Select Case modo
                    Case CTE_ABRIR_BORRAR
                    Case CTE_ABRIR_ANEXAR
                    Case CTE_ABRIR_NOBORRAR
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error"
                End Select
                'Creo la hoja Excel: la borro primero por si acaso existia
                'Esto no hace falta, ya que es posible generar un .xls sin que
                'excel este abierto
                'MsgBox "Ahora se va a intentar ejecutar el programa Excel ubicado en c:\Archivos de Programa\Microsoft Office\Office\EXCEL.EXE. Si Excel estuviera en otro directorio, se deberá abrir manualmente.", vbInformation
                'AppActivate Shell("c:\Archivos de Programa\Microsoft Office\Office\EXCEL.EXE", 1)    ' Activate Microsoft Excel
                'MsgBox "Comprobar que excel está abierto antes de pulsar este botón de aceptar. Si no está abierto, ejecutar Excel ahora.", vbInformation
                Set HojaResumenExcel = Nothing
                Set HojaResumenExcel = CreateObject("Excel.Sheet")
                HojaResumenExcel.Title = "Resumen Ejecución" 'esto no funciona
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Resumen Ejecución"
            End If
        Case CTE_FIC_23W_1EJGRA
            If un_ej_grabar_gra_ejv Then 'Solo si hay que tratar este fichero
                Select Case modo
                    Case CTE_ABRIR_BORRAR
                        Open un_ej_fichero_gra_ejv For Output As #fic
                    Case CTE_ABRIR_ANEXAR
                        Open un_ej_fichero_gra_ejv For Append As #fic
                    Case CTE_ABRIR_NOBORRAR
                        un_ej_fichero_gra_ejv = f_nombre_completo_fichero_no_existente(un_ej_fichero_gra_ejv, CTE_C_SAL_GRA)
                        Open un_ej_fichero_gra_ejv For Output As #fic
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error"
                End Select
                If cabeceras_ejv Then
                    linea = "Fichero de Gráfico (.gra)"
                    Print #fic, linea
                    linea = Date & " " & Time
                    Print #fic, linea
                    linea = "========================="
                    Print #fic, linea
                    linea = ""
                    Print #fic, linea
                End If
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Fichero gráfico (.gra) abierto con exito " & un_ej_fichero_gra_ejv
            End If
        Case CTE_FIC_24_1EJTXT
            If un_ej_grabar_resumen_txt_ejv Then 'Solo si hay que tratar este fichero
                Select Case modo
                    Case CTE_ABRIR_BORRAR
                        Open un_ej_fichero_resumen_txt_ejv For Output As #fic
                    Case CTE_ABRIR_ANEXAR
                        Open un_ej_fichero_resumen_txt_ejv For Append As #fic
                    Case CTE_ABRIR_NOBORRAR
                        un_ej_fichero_resumen_txt_ejv = f_nombre_completo_fichero_no_existente(un_ej_fichero_resumen_txt_ejv, CTE_C_SAL_TXT)
                        Open un_ej_fichero_resumen_txt_ejv For Output As #fic
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error"
                End Select
                If cabeceras_ejv Then
                    linea = "Fichero Resumen de un ejemplo (.txt)"
                    Print #fic, linea
                    linea = "===================================="
                    Print #fic, linea
                    linea = ""
                    Print #fic, linea
                 End If
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Fichero resumen de un ejemplo (.txt) abierto con exito " & un_ej_fichero_resumen_txt_ejv
            End If
        Case CTE_FIC_25_1EJXLS
            If un_ej_grabar_resumen_xls_ejv Then 'Solo si hay que tratar este fichero
                Select Case modo
                    Case CTE_ABRIR_BORRAR
                    Case CTE_ABRIR_ANEXAR
                    Case CTE_ABRIR_NOBORRAR
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error"
                End Select
                'Creo la hoja Excel: la borro primero por si acaso existia
                'Esto no hace falta, ya que es posible generar un .xls sin que
                'excel este abierto
                'MsgBox "Ahora se va a intentar ejecutar el programa Excel ubicado en c:\Archivos de Programa\Microsoft Office\Office\EXCEL.EXE. Si Excel estuviera en otro directorio, se deberá abrir manualmente.", vbInformation
                'AppActivate Shell("c:\Archivos de Programa\Microsoft Office\Office\EXCEL.EXE", 1)    ' Activate Microsoft Excel
                'MsgBox "Comprobar que excel está abierto antes de pulsar este botón de aceptar. Si no está abierto, ejecutar Excel ahora.", vbInformation
                Set HojaUnEjResumenExcel = Nothing
                Set HojaUnEjResumenExcel = CreateObject("Excel.Sheet")
                HojaUnEjResumenExcel.Title = "Resumen Ejecución de un Ejemplo" 'esto no funciona
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Resumen Ejecución de un Ejemplo", ContFilasHojaUnEjResumenExcel, 1
            End If
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select

End Sub

Sub s_grabar_dato_fichero_salida_ejv(fic As Integer, dato As Variant, Optional fila As Variant, Optional columna As Variant)
'======================================================================
'Esta funcion se puede llamar aun cuando no se va a usar cierto fichero
'y en ese caso se sale de la función sin hacer nada
'======================================================================
'On Error Resume Next

    Select Case fic
        Case CTE_FIC_20_GLOLOG
            If grabar_log_ejv Then 'Solo si hay que tratar este fichero
                dato = Date & " " & Time & " - " & dato
                Print #fic, dato
            End If
        Case CTE_FIC_21_GLOTXT
            If grabar_resumen_txt_ejv Then 'Solo si hay que tratar este fichero
                dato = dato
                Print #fic, dato
            End If
        Case CTE_FIC_22_GLOXLS
            If grabar_resumen_xls_ejv Then 'Solo si hay que tratar este fichero
                HojaResumenExcel.ActiveSheet.Cells(fila, columna).Value = dato
            End If
        Case CTE_FIC_23W_1EJGRA
            If un_ej_grabar_gra_ejv Then 'Solo si hay que tratar este fichero
                Print #fic, dato
            End If
        Case CTE_FIC_24_1EJTXT
            If un_ej_grabar_resumen_txt_ejv Then 'Solo si hay que tratar este fichero
                Print #fic, dato
            End If
        Case CTE_FIC_25_1EJXLS
            If un_ej_grabar_resumen_xls_ejv Then 'Solo si hay que tratar este fichero
                HojaUnEjResumenExcel.ActiveSheet.Cells(fila, columna).Value = dato
                ContFilasHojaUnEjResumenExcel = ContFilasHojaUnEjResumenExcel + 1
            End If
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select

End Sub
Sub s_abrir_ficheros_un_ejemplo_ejv(modo As Integer)

    'Abro los ficheros
    s_abrir_fichero_salida_ejv CTE_FIC_23W_1EJGRA, modo
    s_abrir_fichero_salida_ejv CTE_FIC_24_1EJTXT, modo
    s_abrir_fichero_salida_ejv CTE_FIC_25_1EJXLS, modo

End Sub

Sub s_cerrar_ficheros_globales_ejv()
    
    'Salvo y cierro los ficheros globales de toda la sesion
    s_cerrar_fichero_salida_ejv CTE_FIC_22_GLOXLS
    s_cerrar_fichero_salida_ejv CTE_FIC_21_GLOTXT
    s_cerrar_fichero_salida_ejv CTE_FIC_20_GLOLOG

End Sub
        

Sub s_cerrar_ficheros_un_ejemplo_ejv()
        
    'Salvo y cierro los ficheros relativos al ejemplo que se esta ejecutando
    s_cerrar_fichero_salida_ejv CTE_FIC_23W_1EJGRA
    s_cerrar_fichero_salida_ejv CTE_FIC_24_1EJTXT
    s_cerrar_fichero_salida_ejv CTE_FIC_25_1EJXLS

End Sub

Sub s_grabar_ficheros_un_ejemplo_ejv()
    'Grabo los ficheros sin cerrarlos, por si hay un corte de luz y esas cosas
    s_grabar_fichero_salida_ejv CTE_FIC_23W_1EJGRA
    s_grabar_fichero_salida_ejv CTE_FIC_24_1EJTXT
    s_grabar_fichero_salida_ejv CTE_FIC_25_1EJXLS

End Sub


Sub s_grabar_fichero_salida_ejv(fic As Integer)
'======================================================================
'Esta funcion se puede llamar aun cuando no se va a usar cierto fichero
'y en ese caso se sale de la función sin hacer nada
'======================================================================
'On Error Resume Next

    Select Case fic
        Case CTE_FIC_20_GLOLOG
            If grabar_log_ejv Then 'Solo si hay que tratar este fichero
                Close #fic
                Open fichero_log_ejv For Append As #fic
            End If
        Case CTE_FIC_21_GLOTXT
            If grabar_resumen_txt_ejv Then 'Solo si hay que tratar este fichero
                Close #fic
                Open fichero_resumen_txt_ejv For Append As #fic
            End If
        Case CTE_FIC_22_GLOXLS
            If grabar_resumen_xls_ejv Then 'Solo si hay que tratar este fichero
                HojaResumenExcel.SaveCopyAs fichero_resumen_xls_ejv
            End If
        Case CTE_FIC_23W_1EJGRA
            If un_ej_grabar_gra_ejv Then 'Solo si hay que tratar este fichero
                Close #fic
                Open un_ej_fichero_gra_ejv For Append As #fic
            End If
        Case CTE_FIC_24_1EJTXT
            If un_ej_grabar_resumen_txt_ejv Then 'Solo si hay que tratar este fichero
                Close #fic
                Open un_ej_fichero_resumen_txt_ejv For Append As #fic
            End If
        Case CTE_FIC_25_1EJXLS
            If un_ej_grabar_resumen_xls_ejv Then 'Solo si hay que tratar este fichero
                HojaUnEjResumenExcel.SaveCopyAs un_ej_fichero_resumen_xls_ejv
            End If
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select

End Sub

Sub s_cerrar_fichero_salida_ejv(fic As Integer)
'======================================================================
'Esta funcion se puede llamar aun cuando no se va a usar cierto fichero
'y en ese caso se sale de la función sin hacer nada
'======================================================================
'On Error Resume Next


    Select Case fic
        Case CTE_FIC_20_GLOLOG
            If grabar_log_ejv Then 'Solo si hay que tratar este fichero
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "El programa va a finalizar."
                Close #fic
            End If
        Case CTE_FIC_21_GLOTXT
            If grabar_resumen_txt_ejv Then 'Solo si hay que tratar este fichero
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se va a grabar y cerrar el resumen en TXT " & fichero_resumen_txt_ejv
                Close #fic
            End If
        Case CTE_FIC_22_GLOXLS
            If grabar_resumen_xls_ejv Then 'Solo si hay que tratar este fichero
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se va a grabar y cerrar el resumen en XLS " & fichero_resumen_xls_ejv
                s_grabar_fichero_salida_ejv fic
                HojaResumenExcel.Close
            End If
        Case CTE_FIC_23W_1EJGRA
            If un_ej_grabar_gra_ejv Then 'Solo si hay que tratar este fichero
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se va a grabar y cerrar el resumen para gráficos de un ejemplo " & un_ej_fichero_gra_ejv
                Close #fic
            End If
        Case CTE_FIC_24_1EJTXT
            If un_ej_grabar_resumen_txt_ejv Then 'Solo si hay que tratar este fichero
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se va a grabar y cerrar el resumen de un ejemplo en TXT " & un_ej_fichero_resumen_txt_ejv
                Close #fic
            End If
        Case CTE_FIC_25_1EJXLS
            If un_ej_grabar_resumen_xls_ejv Then 'Solo si hay que tratar este fichero
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se va a grabar y cerrar el resumen del ejemplo en XLS " & un_ej_fichero_resumen_xls_ejv
                s_grabar_fichero_salida_ejv fic
                HojaUnEjResumenExcel.Close
            End If
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select


End Sub

Function f_numero_de_lineas_fic(fic As String, indice_fic As Integer) As Long
    
    Dim linea As String
    
    Open fic For Input As #indice_fic
    f_numero_de_lineas_fic = 0
    While Not EOF(indice_fic)
        Line Input #indice_fic, linea
        f_numero_de_lineas_fic = f_numero_de_lineas_fic + 1
    Wend
    Close #indice_fic

End Function

Sub s_grabar_array_en_fic(mi_array() As String, fic As String, indice_fic As Integer)

    Dim i As Integer

    Open fic For Output As #indice_fic
    For i = LBound(mi_array, 1) To UBound(mi_array, 1)
        Print #indice_fic, mi_array(i)
        DoEvents
    Next i
    Close #indice_fic

End Sub

Sub s_cargar_fic_en_array(mi_array() As String, fic As String, indice_fic As Integer)

    Dim cont_lin As Integer
    Dim num_lin As Long
    Dim linea As String
    
    cont_lin = 0
    num_lin = f_numero_de_lineas_fic(fic, indice_fic)
    If num_lin = 0 Then
        ReDim mi_array(1 To 1) As String
        mi_array(1) = ""
    Else
        ReDim mi_array(1 To num_lin) As String
        Open fic For Input As #indice_fic
        While Not EOF(indice_fic)
            linea = f_leer_linea(indice_fic)
            If linea <> "" Then
                cont_lin = cont_lin + 1
                mi_array(cont_lin) = linea
            End If
        Wend
        Close #indice_fic
    End If
End Sub

Function f_todos_fin_fichero(num_indice_fic() As Integer) As Boolean

    Dim i As Integer

    f_todos_fin_fichero = True
    For i = 1 To UBound(num_indice_fic)
        If Not EOF(num_indice_fic(i)) Then
            f_todos_fin_fichero = False
            Exit Function
        End If
    Next i

End Function

Function f_ordenar_y_quitar_repetidos_gran_fichero_texto(fic As String)

    Dim num_max_fic As Integer
    Dim linea As String
    Dim diccionario() As String
    Dim palabra_anterior As String
    Dim num_pal_dicc As Long
    Dim max_num_pal_cada_temp As Integer
    Dim cont_pal_dicc As Integer
    Dim num_fic_temp As Integer
    Dim cont_fic_temp As Integer
    Dim fic_temp() As String
    Dim num_indice_fic_temp() As Integer
    Dim lin_fic() As String
    Dim fic_menor As Integer

    num_max_fic = CTE_FIC_511_ULTI_LISTA - CTE_FIC_100_ULTIMO '511-100=411 -> 101..511

    frm_u0_dicc.Op_Estado.Text = "Ordenando el fichero de salida" & vbCrLf
    frm_u0_dicc.Op_Estado2.Text = frm_u0_dicc.Op_Estado2.Text & "Ordenando el fichero de salida" & vbCrLf
    DoEvents

    'Calculo el numero de ficheros temporales
    num_pal_dicc = f_numero_de_lineas_fic(fic, CTE_FIC_100_ULTIMO)
    'Estimo un maximo de palabras que me genere aprox 411 ficheros pero obligando a que cada uno tenga al menos 100 palabras
    max_num_pal_cada_temp = Int(num_pal_dicc / num_max_fic)
    If max_num_pal_cada_temp < 100 Then
        max_num_pal_cada_temp = 100 'obligando a que cada uno tenga al menos 100 palabras
    End If
    If Int(num_pal_dicc / max_num_pal_cada_temp) = num_pal_dicc / 100 Then
        num_fic_temp = num_pal_dicc / max_num_pal_cada_temp
    Else
        num_fic_temp = Int(num_pal_dicc / max_num_pal_cada_temp) + 1
    End If
    'Por si acaso (no me fio ni de las matematicas ni de los programadores como yo)
    If num_fic_temp > num_max_fic Then
        num_fic_temp = 411
    End If
    If num_fic_temp < 1 Then
        num_fic_temp = 1111
    End If
    
    frm_u0_dicc.Op_Estado.Text = "Se van a generar " & num_fic_temp & " ficheros temporales" & vbCrLf
    frm_u0_dicc.Op_Estado2.Text = frm_u0_dicc.Op_Estado2.Text & "Se van a generar " & num_fic_temp & " ficheros temporales" & vbCrLf
    DoEvents
    
    'Defino los nombres de los ficheros temporales y los creo vacios
    ReDim fic_temp(1 To num_fic_temp) As String
    ReDim num_indice_fic_temp(1 To num_fic_temp) As Integer
    ReDim lin_fic(1 To num_fic_temp) As String
    For cont_fic_temp = 1 To num_fic_temp
        fic_temp(cont_fic_temp) = f_nombre_completo_fichero_no_existente(f_path_fichero(fic, CTE_C_PRG_UTIL) & "$_temporal.txt", CTE_C_PRG_UTIL)
        num_indice_fic_temp(cont_fic_temp) = CTE_FIC_100_ULTIMO + cont_fic_temp
        Open fic_temp(cont_fic_temp) For Output As #num_indice_fic_temp(cont_fic_temp)
        Close #num_indice_fic_temp(cont_fic_temp)
    Next cont_fic_temp

    'Abro el fichero de salida y lo divido en varios ficheros ordenados sin repetir palabras
    Open fic For Input As #CTE_FIC_100_ULTIMO
    cont_pal_dicc = 0
    cont_fic_temp = 0
    palabra_anterior = ""
    ReDim diccionario(1 To max_num_pal_cada_temp) As String
    While Not EOF(CTE_FIC_100_ULTIMO)
        linea = f_leer_linea(CTE_FIC_100_ULTIMO)
        'Detecto en que fichero debe ir
        If cont_pal_dicc = max_num_pal_cada_temp Then
            'Se graba en uno nuevo; antes, ordeno y grabo el anterior
            'ordeno y grabo el anterior
            cont_fic_temp = cont_fic_temp + 1
            frm_u0_dicc.Op_Estado.Text = "Escribiendo " & fic_temp(cont_fic_temp) & vbCrLf
            frm_u0_dicc.Op_Estado2.Text = frm_u0_dicc.Op_Estado2.Text & "Escribiendo " & fic_temp(cont_fic_temp) & vbCrLf
            DoEvents
            S_OrdenarArrayStrMinMax_bur diccionario()
            s_grabar_array_en_fic diccionario(), fic_temp(cont_fic_temp), num_indice_fic_temp(cont_fic_temp)
            'Se graba en uno nuevo
            ReDim diccionario(1 To max_num_pal_cada_temp) As String
            cont_pal_dicc = 1
            diccionario(cont_pal_dicc) = linea
        Else
            'Se graba en el actual
            cont_pal_dicc = cont_pal_dicc + 1
            diccionario(cont_pal_dicc) = linea
        End If
    Wend
    'ordeno y grabo el último
    cont_fic_temp = cont_fic_temp + 1
    frm_u0_dicc.Op_Estado.Text = "Escribiendo " & fic_temp(cont_fic_temp) & vbCrLf
    frm_u0_dicc.Op_Estado2.Text = frm_u0_dicc.Op_Estado2.Text & "Escribiendo " & fic_temp(cont_fic_temp) & vbCrLf
    DoEvents
    S_OrdenarArrayStrMinMax_bur diccionario()
    s_grabar_array_en_fic diccionario(), fic_temp(cont_fic_temp), num_indice_fic_temp(cont_fic_temp)
    Close #CTE_FIC_100_ULTIMO
    
    'Quito los repetidos de los ficheros temporales
    For cont_fic_temp = 1 To num_fic_temp
        s_cargar_fic_en_array diccionario(), fic_temp(cont_fic_temp), num_indice_fic_temp(cont_fic_temp)
        s_quitar_lineas_repetidas_en_array diccionario()
        s_grabar_array_en_fic diccionario(), fic_temp(cont_fic_temp), num_indice_fic_temp(cont_fic_temp)
    Next cont_fic_temp
    
    'Leo los ficheros temporales y escribo la salida eliminando repetidos
    frm_u0_dicc.Op_Estado.Text = frm_u0_dicc.Op_Estado.Text & "Escribiendo el fichero ordenado" & vbCrLf
    palabra_anterior = ""
    'Abro los ficheros
    Open fic For Output As #CTE_FIC_100_ULTIMO
    For cont_fic_temp = 1 To num_fic_temp
        Open fic_temp(cont_fic_temp) For Input As #num_indice_fic_temp(cont_fic_temp)
    Next cont_fic_temp
    'Leo la primera linea de todos
    For cont_fic_temp = 1 To num_fic_temp
        lin_fic(cont_fic_temp) = f_leer_linea(num_indice_fic_temp(cont_fic_temp))
    Next cont_fic_temp
    While Not f_todos_fin_fichero(num_indice_fic_temp())
        fic_menor = f_indice_elemento_menor_array_s(lin_fic())
        DoEvents
        'Escribo la menor
        If lin_fic(fic_menor) <> palabra_anterior Then
            Print #CTE_FIC_100_ULTIMO, lin_fic(fic_menor)
            palabra_anterior = lin_fic(fic_menor)
        End If
        'Vuelvo a leer del fichero que tiene la menor
        lin_fic(fic_menor) = f_leer_linea(num_indice_fic_temp(fic_menor))
    Wend
    
    'Cierro los ficheros
    Close #CTE_FIC_100_ULTIMO
    For cont_fic_temp = 1 To num_fic_temp
        Close #num_indice_fic_temp(cont_fic_temp)
    Next cont_fic_temp
        
    'Borro los temporales
    For cont_fic_temp = 1 To num_fic_temp
        Kill fic_temp(cont_fic_temp)
    Next cont_fic_temp

End Function


Function f_anaiadir_punto_extension(nombre As String, extension As String) As String

    Dim pos As Integer
    Dim resto As String

    f_anaiadir_punto_extension = nombre
    'Si no tiene punto extension se lo pongo
    If InStr(f_anaiadir_punto_extension, ".") = 0 And Len(f_anaiadir_punto_extension) > 0 Then
        'miro si el extension es punto algo
        pos = InStr(extension, ".")
        If pos = 0 Then
            'si no es punto algo y solo hay el algo, pongo el punto eso
            f_anaiadir_punto_extension = f_anaiadir_punto_extension & "." & extension
        ElseIf pos = 2 And Left(extension, 1) = "*" Then
            'Si es *. pongo el resto
            resto = Right(extension, Len(extension) - 1) 'todo desde el punto inclusive
            If InStr(resto, "*") <> 0 Then
                'Si hay caracteres raros es error
                s_error_ejv SIN_OPCION_FINALIZAR, "Error en el extension del fichero: " & extension
                Exit Function
            Else
                f_anaiadir_punto_extension = f_anaiadir_punto_extension & resto
            End If
        Else
            s_error_ejv SIN_OPCION_FINALIZAR, "Error: Error en el extension del fichero: " & extension
            Exit Function
        End If
        'con eso basta porque fichero se cambia en el evento
    End If

End Function

