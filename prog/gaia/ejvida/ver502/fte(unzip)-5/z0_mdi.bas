Attribute VB_Name = "bas_z0_mdi"
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

Global num_menu_viejos_ejv As Integer
Global terminar_todo_ejv As Boolean
Sub Main()

    Dim linea As String
    Dim parametros As String
    Dim i As Integer
    Dim s_duraciones As String
    
    
    ContFilasHojaResumenExcel = 1
    ContFilasHojaUnEjResumenExcel = 1
    
    terminar_todo_ejv = False
    
    'Nombre aplicación
    nombre_aplicacion_ejv = "Ejemplos de Vida " & App.Major & "." & App.Minor
    version_aplicacion_ejv = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    
    'Para prohibir varias instancias
    'If App.PrevInstance Then
    '    End
    'End If

    'Paths
    path_corto_ejv(CTE_C_RAIZ) = ""
    path_corto_ejv(CTE_C_PRG_3R) = "prg\3r"
    path_corto_ejv(CTE_C_PRG_AUT) = "prg\aut"
    path_corto_ejv(CTE_C_PRG_BMP) = "prg\bmp"
    path_corto_ejv(CTE_C_ENT_DIC) = "ent"
    path_corto_ejv(CTE_C_DOC) = "doc"
    path_corto_ejv(CTE_C_SAL_GRA) = "sal"
    path_corto_ejv(CTE_C_ENT) = "ent"
    path_corto_ejv(CTE_C_PRG_HYP) = "prg\hyp"
    path_corto_ejv(CTE_C_PRG_ICO) = "prg\ico"
    path_corto_ejv(CTE_C_SAL_LOG) = "sal"
    path_corto_ejv(CTE_C_PRG_MAP) = "prg\map"
    path_corto_ejv(CTE_C_DOC_WEB) = "doc\web"
    path_corto_ejv(CTE_C_PRG_PRI) = "prg\pri"
    path_corto_ejv(CTE_C_ENT_RAN) = "ent"
    path_corto_ejv(CTE_C_SAL_TXT) = "sal"
    path_corto_ejv(CTE_C_PRG_UTIL) = "prg\util"
    path_corto_ejv(CTE_C_SAL_XLS) = "sal"
    path_corto_ejv(CTE_C_PRG) = "prg"
    
    
    For i = 1 To CTE_TOTAL_CARPETAS
        path_largo_ejv(i) = f_unir_path(App.Path, path_corto_ejv(i))
    Next i

    parametros = Trim$(UCase(Command$))
    
    'No recuerdo si en la linea de comandos viene el
    'nombre del exe, asi que si viene se lo quito
    If Left(parametros, Len(App.EXEName)) = App.EXEName Then
        parametros = Trim(Right(parametros, Len(parametros) - Len(App.EXEName)))
    End If
    
    If Len(parametros) > 0 Then
        'Si hay parametros
        If f_existe_fichero(parametros) Then
            'El parametro es un fichero con path completo
            nombre_fichero_ejv = parametros
        Else
            If f_existe_fichero(f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), parametros)) Then
                'El parametro es un fichero con path relativo
                nombre_fichero_ejv = f_nombre_completo_existente(path_largo_ejv(CTE_C_RAIZ), parametros)
            Else
                'Vete a saber que es ese parametro
                nombre_fichero_ejv = f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), CTE_nombreINICIO_TXT)
            End If
        End If
    Else
        'No hay parametros
        nombre_fichero_ejv = f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), CTE_nombreINICIO_TXT)
    End If

    'Leo la configurcion indicada
    s_aut_leer_inicio_txt
    s_abrir_fichero_salida_ejv CTE_FIC_20_GLOLOG, CTE_ABRIR_BORRAR
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Se ha leido el fichero global de configuración " & nombre_fichero_ejv
       
    'Activo las acciones que se requieren debido a la configuracion
    If elegir_idioma_ejv Then
        frm_z0_leng.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    End If
       
    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Comienza una ejecución de " & CTE_LOGO_APLICACION & " " & version_ejv
    s_abrir_fichero_salida_ejv CTE_FIC_21_GLOTXT, CTE_ABRIR_BORRAR
    s_abrir_fichero_salida_ejv CTE_FIC_22_GLOXLS, CTE_ABRIR_BORRAR
       
       
    If Not automatico_ejv Then
        '============================================================================
        'Ejecución modo normal
        If mostrar_logo_ejv Then
            'Presentación
            frm_z0_logo.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Else
            s_inicializar_arrays_programa_ejv
        End If
        'Menu Principal de la Aplicacion
        frm_z0_mdi.Show CTE_AMODAL
        '============================================================================
    Else
        '============================================================================
        'Lanzamiento automatico desde excel o desde el propio ejecutable fijando parametros
        s_inicializar_arrays_programa_ejv
        'Menu Principal de la Aplicacion
        frm_z0_mdi.Show CTE_AMODAL
        'Como es automatico, leo el fichero de configuracion
        'Para cada uno de los ejemplos y los lanzo
        For indice_auto = 1 To num_ficheros_aut_ejv
If terminar_todo_ejv Then Exit For
            s_duraciones = ""
            s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Aut: " & f_espacios_izquierda(CStr(indice_auto), 8) & " Fic: " & fichero_aut_ejv(indice_auto)
            If fichero_aut_ejv(indice_auto) = "" Then
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Error: no se encuentra el parametro de configuracion FICHERO AUTOMATICO"
            Else
                If Not f_existe_fichero(fichero_aut_ejv(indice_auto)) Then
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Error: no se encuentra el fichero " & fichero_aut_ejv(indice_auto)
                Else
                    s_aut_leer_fichero_automatico_ejv (indice_auto)
                    'Es posible que cada ejemplo haya que lanzarlo varias veces
                    ContFilasHojaResumenExcel = ContFilasHojaResumenExcel + 1
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_22_GLOXLS, "P" & indice_auto, ContFilasHojaResumenExcel, 1
                    For indice_iteraciones = 1 To num_iteraciones_ejv
If terminar_todo_ejv Then Exit For
                        s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Comienza la ejecución del Aut " & indice_auto & " iteración " & indice_iteraciones
                        s_fijar_caption_mdi
                        s_lanzar_ejecucion_automatica_ejv
                        s_grabar_dato_fichero_salida_ejv CTE_FIC_22_GLOXLS, seg_ej_actual_ejv, ContFilasHojaResumenExcel, indice_iteraciones + 4
                        If indice_iteraciones <> 1 Then
                            s_duraciones = s_duraciones & ";"
                        End If
                        s_duraciones = s_duraciones & """" & CStr(seg_ej_actual_ejv) & """"
                    Next indice_iteraciones
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_21_GLOTXT, s_duraciones
                End If
            End If
            'Grabo los ficheros sin cerrarlos, por si hay un corte de luz y esas cosas
            s_grabar_fichero_salida_ejv CTE_FIC_20_GLOLOG
            s_grabar_fichero_salida_ejv CTE_FIC_21_GLOTXT
            s_grabar_fichero_salida_ejv CTE_FIC_22_GLOXLS
        Next indice_auto
        s_fin_todo
        '============================================================================
    End If

End Sub
Sub s_lanzar_ejecucion_automatica_ejv()

    Dim linea As String

    num_prg_activo_ejv = num_prg_activo_ejv
    s_click_programa_ejv num_prg_activo_ejv
    s_operacion_ejecutar_ejv CTE_EXE_COMENZAR
    

End Sub

Sub s_tratamiento_idioma_mdi()
    
    If idioma_ejv = CTE_INGLES Then
        'Archivo
        frm_z0_mdi.Mn_Archivo.Caption = "&File"
            frm_z0_mdi.mn_Abrir.Caption = "Open   F2"
            frm_z0_mdi.H_Abrir.ToolTipText = "Open File"
            frm_z0_mdi.mn_Guardar.Caption = "Save"
            frm_z0_mdi.H_Guardar.ToolTipText = "Save File"
            frm_z0_mdi.mn_GuardarComo.Caption = "Save As..."
            frm_z0_mdi.mn_generador_aut.Caption = "*.aut Generator..."
            frm_z0_mdi.Mn_Salir.Caption = "Exit   Alt+F4"
        'Editar
        frm_z0_mdi.mn_Edicion.Caption = "&Edit"
                frm_z0_mdi.mn_cortar.Caption = "Cut   Ctrl+X"
                frm_z0_mdi.mn_copiar.Caption = "Copy   Ctrl+C"
                frm_z0_mdi.mn_pegar.Caption = "Paste   Ctrl+V"
        'Ejemplos de Vida
        frm_z0_mdi.mn_Ejemplos.Caption = "&Life Samples"
                frm_z0_mdi.Mn_hyp.Caption = "Ants and Plants...   Ctrl+F1"
                frm_z0_mdi.H_HYP.ToolTipText = "Ants and Plants"
                frm_z0_mdi.mn_palyfras.Caption = "Words and Sentences (Genetic Algorithm)...   Ctrl+F2"
                frm_z0_mdi.H_PAL.ToolTipText = "Words and Sentences"
                frm_z0_mdi.mn_3r.Caption = "Tic-Tac-Toe (Genetic Classifier)...   Ctrl+F3"
                frm_z0_mdi.H_3R.ToolTipText = "Tic-Tac-Toe"
                frm_z0_mdi.mn_Prisionero.Caption = "Prisioner...   Ctrl+F4"
                frm_z0_mdi.H_PRI.ToolTipText = "Prisioner"
                frm_z0_mdi.mn_Explorando_Mapas.Caption = "Maps Explorer...   Ctrl+F7"
                frm_z0_mdi.mn_Editor_Mapas.Caption = "Map Editor..."
                frm_z0_mdi.H_Mapa.ToolTipText = "Map Editor"
            'En obras
            frm_z0_mdi.mn_obras.Caption = "At work"
                frm_z0_mdi.mn_Celdilla.Caption = "Cells..."
                frm_z0_mdi.mn_Cadenas.Caption = "Strings (Genetic Algorithm)..."
                frm_z0_mdi.mn_gaia.Caption = "Gaia Platform..."
        'Opciones
        frm_z0_mdi.mn_Opciones.Caption = "&Options"
            frm_z0_mdi.mn_Opciones1.Caption = "Options I: General Options"
            frm_z0_mdi.H_Opciones.ToolTipText = "Options I: General Options. Stop Conditions, Save Files, Random Function"
            frm_z0_mdi.mn_Opciones2.Caption = "Options II: Execution Mode..."
            frm_z0_mdi.mn_Opciones3.Caption = "Options III: Specific of the current program..."
            frm_z0_mdi.mn_Tipos_Agentes.Caption = "Options IV: Agent Kinds..."
            frm_z0_mdi.H_Agentes.ToolTipText = "Options IV: Agent Kinds"
            frm_z0_mdi.mn_Mapa.Caption = "Options V: Map Editor..."
            frm_z0_mdi.H_Mapa.ToolTipText = "Options V: Map Editor"
            frm_z0_mdi.mn_TipoEvolucion.Caption = "Options VI: Evolution Kind"
            frm_z0_mdi.mn_Metodo_Evaluacion.Caption = "Evaluation Method..."
            frm_z0_mdi.mn_Metodo_Seleccion.Caption = "Selection Method..."
            frm_z0_mdi.mn_Metodo_Reproduccion.Caption = "Reproduction Method..."
        'Ejecutar
        frm_z0_mdi.mn_Ejecutar.Caption = "&Run"
            frm_z0_mdi.mn_Comenzar.Caption = "&Start   F5"
            frm_z0_mdi.H_Comenzar.ToolTipText = "Run"
            frm_z0_mdi.mn_Continuar.Caption = "&Continue   F5"
            frm_z0_mdi.mn_Pausa.Caption = "&Pause   F6"
            frm_z0_mdi.H_Pausa.ToolTipText = "Pause"
            frm_z0_mdi.mn_terminar.Caption = "&Stop   F7"
            frm_z0_mdi.H_Terminar.ToolTipText = "Stop"
        'Ver
        frm_z0_mdi.mn_ver.Caption = "&View"
            frm_z0_mdi.mn_Refrescar.Caption = "Refresh World"
            frm_z0_mdi.H_Refrescar.ToolTipText = "Refresh World"
            frm_z0_mdi.mn_EstadoEjecucion.Caption = "Execution Status"
            frm_z0_mdi.H_Estado.ToolTipText = "Execution Status"
            frm_z0_mdi.mn_ListaAgentes.Caption = "Agents' List (All)"
            frm_z0_mdi.mn_MejoresAgentes.Caption = "Best Agents"
            frm_z0_mdi.mn_Apellidos.Caption = "Surnames"
            frm_z0_mdi.mn_Diccionario.Caption = "Dictionary"
            frm_z0_mdi.mn_JugarContraOrdenador.Caption = "Play against the Computer..."
            frm_z0_mdi.mn_ModificarAgente.Caption = "Change Agent..."
            frm_z0_mdi.mn_Grafico.Caption = "Graphic...   F8"
            frm_z0_mdi.H_Grafico.ToolTipText = "Graphic"
        'Ventana
        frm_z0_mdi.mn_Ventana.Caption = "&Window"
        frm_z0_mdi.mn_Mosaico_Horizontal.Caption = "Tile Horizontal"
        frm_z0_mdi.mn_MosaicoVertical.Caption = "Tile Vertical"
        frm_z0_mdi.mn_Cascada.Caption = "Cascade"
        frm_z0_mdi.mn_Organizar_Iconos.Caption = "Arrange Icons"
        'Ayuda
        frm_z0_mdi.Mn_Ayuda.Caption = "&Help"
        frm_z0_mdi.H_Ayuda.ToolTipText = "Help"
        frm_z0_mdi.Mn_AyudaHtm.Caption = "Documentation in htm..."
        frm_z0_mdi.Mn_Notas.Caption = "Version Notes"
        frm_z0_mdi.Mn_AyudaDoc.Caption = "in Word (.doc=..."
        frm_z0_mdi.Mn_AyudaTxt.Caption = "in text (.txt)..."
        frm_z0_mdi.Mn_Readme.Caption = "Readme.txt for programmers"
        frm_z0_mdi.Mn_Calculadora.Caption = "Calculator..."
        frm_z0_mdi.Mn_AcercaDe.Caption = "About..."
    Else
        'Archivo
        frm_z0_mdi.Mn_Archivo.Caption = "&Fichero"
            frm_z0_mdi.mn_Abrir.Caption = "Abrir   F2"
            frm_z0_mdi.H_Abrir.ToolTipText = "Abrir Fichero"
            frm_z0_mdi.mn_Guardar.Caption = "Guardar"
            frm_z0_mdi.H_Guardar.ToolTipText = "Guardar Fichero"
            frm_z0_mdi.mn_GuardarComo.Caption = "Guardar Como..."
            frm_z0_mdi.mn_generador_aut.Caption = "Generador de *.aut..."
            frm_z0_mdi.Mn_Salir.Caption = "Salir   Alt+F4"
        'Editar
        frm_z0_mdi.mn_Edicion.Caption = "E&dición"
                frm_z0_mdi.mn_cortar.Caption = "Cortar   Ctrl+X"
                frm_z0_mdi.mn_copiar.Caption = "Copiar   Ctrl+C"
                frm_z0_mdi.mn_pegar.Caption = "Pegar   Ctrl+V"
        'Ejemplos de Vida
        frm_z0_mdi.mn_Ejemplos.Caption = "Ejemplos de V&ida"
                frm_z0_mdi.Mn_hyp.Caption = "Hormigas y Plantas...   Ctrl+F1"
                frm_z0_mdi.H_HYP.ToolTipText = "Hormigas y Plantas"
                frm_z0_mdi.mn_palyfras.Caption = "Palabras y Frases (Algoritmo Genético)...   Ctrl+F2"
                frm_z0_mdi.H_PAL.ToolTipText = "Palabras y Frases"
                frm_z0_mdi.mn_3r.Caption = "Tres en Raya (Clasificador Genético)...   Ctrl+F3"
                frm_z0_mdi.H_3R.ToolTipText = "Tres en Raya"
                frm_z0_mdi.mn_Prisionero.Caption = "El Juego del Prisionero...   Ctrl+F4"
                frm_z0_mdi.H_PRI.ToolTipText = "El Juego del Prisionero"
                frm_z0_mdi.mn_Explorando_Mapas.Caption = "Explorando Mapas...   Ctrl+F7"
                frm_z0_mdi.mn_Editor_Mapas.Caption = "Editor de Mapas..."
                frm_z0_mdi.H_Mapa.ToolTipText = "Editor de Mapas"
            'En obras
            frm_z0_mdi.mn_obras.Caption = "En Obras"
                frm_z0_mdi.mn_Celdilla.Caption = "Celdilla..."
                frm_z0_mdi.mn_Cadenas.Caption = "Cadenas (Algoritmo Genético)..."
                frm_z0_mdi.mn_gaia.Caption = "Plataforma Gaia..."
        'Opciones
        frm_z0_mdi.mn_Opciones.Caption = "&Opciones"
            frm_z0_mdi.mn_Opciones1.Caption = "Opciones I: Opciones Generales"
            frm_z0_mdi.H_Opciones.ToolTipText = "Opciones I: Opciones Generales. Condiciones de Parada, Grabar Ficheros, Función de Azar"
            frm_z0_mdi.mn_Opciones2.Caption = "Opciones II: Modo de Ejecución..."
            frm_z0_mdi.mn_Opciones3.Caption = "Opciones III: Específicas del Programa Ejemplo..."
            frm_z0_mdi.mn_Tipos_Agentes.Caption = "Opciones IV: Tipos de Agentes..."
            frm_z0_mdi.H_Agentes.ToolTipText = "Opciones IV: Tipos de Agentes"
            frm_z0_mdi.mn_Mapa.Caption = "Opciones V: Editar Mapa..."
            frm_z0_mdi.H_Mapa.ToolTipText = "Opciones V: Editar Mapa"
            frm_z0_mdi.mn_TipoEvolucion.Caption = "Opciones VI: Tipo de Evolución"
            frm_z0_mdi.mn_Metodo_Evaluacion.Caption = "Método de Evaluación..."
            frm_z0_mdi.mn_Metodo_Seleccion.Caption = "Método de Selección..."
            frm_z0_mdi.mn_Metodo_Reproduccion.Caption = "Método de Reproducción..."
        'Ejecutar
        frm_z0_mdi.mn_Ejecutar.Caption = "&Ejecutar"
            frm_z0_mdi.mn_Comenzar.Caption = "&Comenzar   F5"
            frm_z0_mdi.H_Comenzar.ToolTipText = "Ejecutar"
            frm_z0_mdi.mn_Continuar.Caption = "&Continuar   F5"
            frm_z0_mdi.mn_Pausa.Caption = "&Pausa   F6"
            frm_z0_mdi.H_Pausa.ToolTipText = "Pausa"
            frm_z0_mdi.mn_terminar.Caption = "Terminar   F7"
            frm_z0_mdi.H_Terminar.ToolTipText = "Terminar"
        'Ver
        frm_z0_mdi.mn_ver.Caption = "&Ver"
            frm_z0_mdi.mn_Refrescar.Caption = "Refrescar Mundo"
            frm_z0_mdi.H_Refrescar.ToolTipText = "Refrescar Mundo"
            frm_z0_mdi.mn_EstadoEjecucion.Caption = "Estado de la Ejecución"
            frm_z0_mdi.H_Estado.ToolTipText = "Estado de la Ejecución"
            frm_z0_mdi.mn_ListaAgentes.Caption = "Lista de Agentes (Todos)"
            frm_z0_mdi.mn_MejoresAgentes.Caption = "Mejores Agentes"
            frm_z0_mdi.mn_Apellidos.Caption = "Apellidos"
            frm_z0_mdi.mn_Diccionario.Caption = "Diccionario"
            frm_z0_mdi.mn_JugarContraOrdenador.Caption = "Jugar contra el Ordenador..."
            frm_z0_mdi.mn_ModificarAgente.Caption = "Modificar Agente..."
            frm_z0_mdi.mn_Grafico.Caption = "Gráfico...   F8"
            frm_z0_mdi.H_Grafico.ToolTipText = "Gráfico"
        'Ventana
        frm_z0_mdi.mn_Ventana.Caption = "Ven&tana"
        frm_z0_mdi.mn_Mosaico_Horizontal.Caption = "Mosaico Horizontal"
        frm_z0_mdi.mn_MosaicoVertical.Caption = "Mosaico Vertical"
        frm_z0_mdi.mn_Cascada.Caption = "Cascada"
        frm_z0_mdi.mn_Organizar_Iconos.Caption = "Organizar Iconos"
        'Ayuda
        frm_z0_mdi.Mn_Ayuda.Caption = "&Ayuda"
        frm_z0_mdi.H_Ayuda.ToolTipText = "Ayuda"
        frm_z0_mdi.Mn_AyudaHtm.Caption = "Documentación en htm..."
        frm_z0_mdi.Mn_Notas.Caption = "Notas sobre la Versión"
        frm_z0_mdi.Mn_AyudaDoc.Caption = "en Word (.doc)..."
        frm_z0_mdi.Mn_AyudaTxt.Caption = "en Texto (.txt)..."
        frm_z0_mdi.Mn_Readme.Caption = "Readme.txt para programadores"
        frm_z0_mdi.Mn_Calculadora.Caption = "Calculadora..."
        frm_z0_mdi.Mn_AcercaDe.Caption = "Acerca de..."
    End If


End Sub

Sub s_leer_parametros_fichero_aut_ejv(indice_auto As Long)

    Dim linea As String
    Dim pos_igual As Integer
    
    Dim param As String
    Dim valor As String
    
    Dim esta_inicializado As Boolean
    Dim num_prog_ya_leido As Boolean
    Dim num_ej_ya_leido As Boolean
    esta_inicializado = False
    num_prog_ya_leido = False
    num_ej_ya_leido = False
    

    While Not EOF(CTE_FIC_02_AUT)
        linea = f_leer_linea(CTE_FIC_02_AUT)
        'Paso a mayúsculas
        linea = UCase(linea)
        'El nombre del parametro es la parte izquierda del igual
        pos_igual = InStr(linea, "=")
        If pos_igual = 0 Then
            s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Error en la línea" & linea
        End If
        param = Left(linea, pos_igual)
        valor = Right(linea, Len(linea) - Len(param))
        If InStr(linea, param) <> 0 Then
            Select Case param
                'Parametros opcionales comunes de cualquier ejemplo automático
                Case CTE_AUTOMATICO_NUMERO_PROGRAMA
                    num_prg_activo_ejv = CInt(valor)
                    num_prog_ya_leido = True
                Case CTE_AUTOMATICO_NUMERO_EJEMPLO
                    num_ej_activo_ejv = CInt(valor)
                    num_ej_ya_leido = True
                Case CTE_FICHERO_RESULTADOS
                    un_ej_fichero_resumen_xls_ejv = valor
                Case CTE_ITERACIONES
                    num_iteraciones_ejv = CLng(valor)
                'Parametros específicos del ejemplo automático
                Case CTE_FRASE_A_BUSCAR
                    frase_a_buscar_pal = valor
                Case CTE_CRITERIO_DE_PARADA
                    CondParadaPeso_ejv = True
                    CondParadaPesoNecesario_ejv = CDbl(valor)
                Case Else
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Error: no se encuentra el parametro de configuracion " & param & " en la línea " & linea
            End Select
        Else
            s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Error al leer el fichero de configuración " & fichero_aut_ejv(indice_auto) & ". La línea " & linea & " es incorrecta o se trata de una versión anterior."
        End If
        'Tengo que inicializar despues de saber cual es el programa y el ejemplo
        'pero antes del resto de parametros
        If Not esta_inicializado Then
            If num_prog_ya_leido And num_ej_ya_leido Then
                s_inicializar_ejemplo_elegido_ejv
                esta_inicializado = True
            End If
        End If
    Wend

End Sub

Sub s_leer_parametros_inicio_txt()

    Dim linea As String
    Dim res As Integer
    Dim pos_igual As Integer
    
    Dim param As String
    Dim valor As String

    linea = ""
    num_ficheros_aut_ejv = 0
    
    While Not EOF(CTE_FIC_01_INICIO)
        linea = f_leer_linea(CTE_FIC_01_INICIO)
        'Paso a mayúsculas
        linea = UCase(linea)
        'El nombre del parametro es la parte izquierda del igual
        pos_igual = InStr(linea, "=")
        'Control de errores
        If pos_igual = 0 Then
            s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Error en el fichero " & CTE_nombreINICIO_TXT & " en la línea" & linea
            'Sea automatico o no doy error porque es muy grave
            res = MsgBox("Error al leer el fichero de inicio " & nombre_fichero_ejv & ". La línea " & linea & " es incorrecta o se trata de una versión anterior. Una posible solución es crear un nuevo fichero de configuración por defecto actualizado y arrancar de nuevo el programa. Si desea crearlo y cerrar el programa, pulse SI. Si desea continuar, pulse NO. Si desea detener el programa en este instante pulse CANCELAR.", vbCritical + vbYesNoCancel)
            If res = vbYes Then
                Close #CTE_FIC_01_INICIO
                grabar_configuracion_ejv = True
                grabar_config_defecto_ejv = True
                s_fin_todo
            Else
                If res = vbCancel Then
                    End
                Else
                    grabar_configuracion_ejv = False
                    grabar_config_defecto_ejv = False
                End If
            End If
        End If
        'Hay igual
        param = Left(linea, pos_igual)
        valor = Right(linea, Len(linea) - Len(param))
        Select Case param
            '1: VERSION
            Case CTE_VERSION
                version_ejv = valor
            '2: IDIOMA
            Case CTE_IDIOMA
                idioma_ejv = valor
            '3: ELEGIR IDIOMA
            Case CTE_ELEGIRIDIOMA
                elegir_idioma_ejv = valor
            '4: CONTROL DE ERRORES
            Case CTE_CTRLERRORES
                If valor = UCase(CTE_txtFALSE) Then
                    control_errores_de_programacion_ejv = False
                Else
                    control_errores_de_programacion_ejv = True
                End If
            '5: MOSTRAR_LOGO
            Case CTE_MOSTRAR_LOGO
                If valor = UCase(CTE_txtFALSE) Then
                    mostrar_logo_ejv = False
                Else
                    mostrar_logo_ejv = True
                End If
            '6: ALGORITMO DE ORDENACION
            Case CTE_ALGORITMODEORDENACION
                algoritmo_ordenacion_ejv = valor
            '7: SISTEMA OPERATIVO
            Case CTE_SISTEMA_OPERATIVO
                sistema_operativo_ejv = valor
            '8: PEDIR CONFIRMACION
            Case CTE_PEDIR_CONFIRMACION
                If valor = UCase(CTE_txtFALSE) Then
                    pedir_confirmacion_ejv = False
                Else
                    pedir_confirmacion_ejv = True
                End If
            '9: RESOLUCION PANTALLA
            Case CTE_RESOLUCIONPANTALLA
                resolucion_pantalla_ejv = valor
            '10: GRABAR CONFIGURACION
                Case CTE_GRABAR_CONFIGURACION
                If valor = UCase(CTE_txtTRUE) Then
                    grabar_configuracion_ejv = True
                Else
                    grabar_configuracion_ejv = False
                End If
            '11: GRABAR CONFIG POR DEFECTO
            Case CTE_GRABAR_CONFIG_POR_DEFECTO
                If valor = UCase(CTE_txtTRUE) Then
                    grabar_config_defecto_ejv = True
                Else
                    grabar_config_defecto_ejv = False
                End If
            '12: GRABAR LOG
            Case CTE_GRABAR_LOG
                If valor = UCase(CTE_txtTRUE) Then
                    grabar_log_ejv = True
                Else
                    grabar_log_ejv = False
                End If
            '13: FICHERO LOG
            Case CTE_FICHERO_LOG
                fichero_log_ejv = valor
                fichero_log_ejv = f_nombre_completo_obligatiorio_ejv(fichero_log_ejv, CTE_C_SAL_LOG)
            '14: GRABAR RESUMEN TXT
            Case CTE_GRABAR_RESUMEN_TXT
                If valor = UCase(CTE_txtTRUE) Then
                    grabar_resumen_txt_ejv = True
                Else
                    grabar_resumen_txt_ejv = False
                End If
            '15: FICHERO RESUMEN TXT
            Case CTE_FICHERO_RESUMEN_TXT
                fichero_resumen_txt_ejv = valor
                fichero_resumen_txt_ejv = f_nombre_completo_obligatiorio_ejv(fichero_resumen_txt_ejv, CTE_C_SAL_TXT)
            '16: GRABAR RESUMEN EXCEL
            Case CTE_GRABAR_RESUMEN_EXCEL
                If valor = UCase(CTE_txtTRUE) Then
                    grabar_resumen_xls_ejv = True
                Else
                    grabar_resumen_xls_ejv = False
                End If
            '17: FICHERO RESUMEN EXCEL
            Case CTE_FICHERO_RESUMEN_EXCEL
                fichero_resumen_xls_ejv = valor
                fichero_resumen_xls_ejv = f_nombre_completo_obligatiorio_ejv(fichero_resumen_xls_ejv, CTE_C_SAL_XLS)
            '18: REEMPLAZAR FICHEROS EXISTENTES
            Case CTE_REEMPLAZAR_FICHEROS_EXISTENTES
                If valor = UCase(CTE_txtTRUE) Then
                    reemplazar_fic_ejv = True
                Else
                    reemplazar_fic_ejv = False
                End If
            '19: AUTOMATICO
            Case CTE_AUTOMATICO
                If valor = UCase(CTE_txtTRUE) Then
                    automatico_ejv = True
                Else
                    automatico_ejv = False
                    'Este control solo lo hago si no es automatico
                    If version_ejv <> App.Major & "." & App.Minor & "." & App.Revision Then
                        MsgBox "No coinciden los números de version " & version_ejv & " (del " & CTE_nombreINICIO_TXT & ") y " & App.Major & "." & App.Minor & "." & App.Revision & " del exe.", vbInformation
                    End If
                End If
            '20: FICHERO AUTOMATICO
            Case CTE_FICHEROAUTOMATICO
                If valor = "" Or valor = CTE_NOHAY Then
                    num_ficheros_aut_ejv = num_ficheros_aut_ejv + 1
                    ReDim Preserve fichero_aut_ejv(1 To num_ficheros_aut_ejv) As String
                    fichero_aut_ejv(num_ficheros_aut_ejv) = CTEm_NOHAY
                Else
                    'Ya he leido una, asi que la añado al array
                    num_ficheros_aut_ejv = num_ficheros_aut_ejv + 1
                    ReDim Preserve fichero_aut_ejv(1 To num_ficheros_aut_ejv) As String
                    fichero_aut_ejv(num_ficheros_aut_ejv) = valor
                    'fichero_aut_ejv(num_ficheros_aut_ejv) = f_nombre_completo_obligatiorio_ejv(fichero_aut_ejv(num_ficheros_aut_ejv), CTE_C_PRG_AUT)
                End If
            Case Else
                s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Error en el fichero " & CTE_nombreINICIO_TXT & " en la línea" & linea
                'Sea automatico o no doy error porque es muy grave
                res = MsgBox("Error al leer el fichero de inicio " & nombre_fichero_ejv & ". La línea " & linea & " para el parametro " & param & " es incorrecta o se trata de una versión anterior. Una posible solución es crear un nuevo fichero de configuración por defecto actualizado y arrancar de nuevo el programa. Si desea crearlo y cerrar el programa, pulse SI. Si desea continuar, pulse NO. Si desea detener el programa en este instante pulse CANCELAR.", vbCritical + vbYesNoCancel)
                If res = vbYes Then
                    Close #CTE_FIC_01_INICIO
                    grabar_configuracion_ejv = True
                    grabar_config_defecto_ejv = True
                    s_fin_todo
                Else
                    If res = vbCancel Then
                        End
                    Else
                        grabar_configuracion_ejv = False
                        grabar_config_defecto_ejv = False
                    End If
                End If
        End Select
    Wend
        


End Sub

Sub s_tratamiento_idioma_menu()
    
    If idioma_ejv = CTE_INGLES Then
        frm_z0_menu.quedesea.Caption = "What do you want to do?"
        frm_z0_menu.Op_1.Caption = "Execute an example"
        frm_z0_menu.Op_2.Caption = "Create a new example"
        frm_z0_menu.Op_3.Caption = "Load a saved example"
        frm_z0_menu.Aceptar.Caption = "OK"
        frm_z0_menu.Cancelar.Caption = "Cancel"
    End If


End Sub

Sub s_mostrar_auto_ejv(parametro As Long)

    On Error Resume Next
    Dim RetVal
    RetVal = Shell("notepad.exe " & fichero_aut_ejv(parametro), 1)
    

End Sub

Sub s_mostrar_config_ejv()

    Dim indice_auto As Long

    frm_z0_inic.List1.Clear
    frm_z0_inic.List2.Clear

    '1: VERSION
    frm_z0_inic.List1.AddItem "Version:"
    frm_z0_inic.List2.AddItem version_ejv
    
    '2: IDIOMA
    frm_z0_inic.List1.AddItem "Idioma:"
    Select Case idioma_ejv
        Case CTE_CASTELLANO
            frm_z0_inic.List2.AddItem CTEm_CASTELLANO
        Case CTE_INGLES
            frm_z0_inic.List2.AddItem CTEm_INGLES
        Case Else
            MsgBox "Error: el parámetro " & idioma_ejv & " es incorrecto. Se va a tomar el valor por defecto " & CTEm_CASTELLANO & ". Se recomienda activar la opción de guardar el fichero de configuración para evitar este error.", vbCritical
            idioma_ejv = CTE_CASTELLANO
            frm_z0_inic.List2.AddItem CTEm_CASTELLANO
    End Select
    
    '3: ELEGIR IDIOMA
    frm_z0_inic.List1.AddItem "Elegir idioma:"
    If elegir_idioma_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '4: CONTROL DE ERRORES
    frm_z0_inic.List1.AddItem "Control de errores:"
    If control_errores_de_programacion_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '5: MOSTRAR_LOGO
    frm_z0_inic.List1.AddItem "Mostrar logo:"
    If mostrar_logo_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '6: ALGORITMO DE ORDENACION
    frm_z0_inic.List1.AddItem "Algoritmo de ordenacion:"
    Select Case algoritmo_ordenacion_ejv
        Case CTE_BURBUJA
            frm_z0_inic.List2.AddItem CTEm_BURBUJA
        Case CTE_QUICKSORT
            frm_z0_inic.List2.AddItem CTEm_QUICKSORT
        Case Else
            MsgBox "Error: el parámetro " & algoritmo_ordenacion_ejv & " es incorrecto. Se va a tomar el valor por defecto " & CTEm_BURBUJA & ". Se recomienda activar la opción de guardar el fichero de configuración para evitar este error.", vbCritical
            algoritmo_ordenacion_ejv = CTE_BURBUJA
            frm_z0_inic.List2.AddItem CTEm_BURBUJA
    End Select
    
    '7: CTE_SISTEMA_OPERATIVO
    frm_z0_inic.List1.AddItem "Sistema Operativo:"
    Select Case sistema_operativo_ejv
        Case CTE_WINDOWS95
            frm_z0_inic.List2.AddItem CTEm_WINDOWS95
        Case CTE_WINDOWSNT
            frm_z0_inic.List2.AddItem CTEm_WINDOWSNT
        Case CTE_WINDOWS3X
            frm_z0_inic.List2.AddItem CTEm_WINDOWS3X
        Case Else
            MsgBox "Error: el parámetro " & sistema_operativo_ejv & " es incorrecto. Se va a tomar el valor por defecto " & CTEm_WINDOWS95 & ". Se recomienda activar la opción de guardar el fichero de configuración para evitar este error.", vbCritical
            algoritmo_ordenacion_ejv = CTE_BURBUJA
            frm_z0_inic.List2.AddItem CTEm_BURBUJA
    End Select
    
    '8: CTE_PEDIR_CONFIRMACION
    frm_z0_inic.List1.AddItem "Pedir Confirmación:"
    If pedir_confirmacion_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '9: RESOLUCION PANTALLA:
    frm_z0_inic.List1.AddItem "Resolución pantalla:"
    Select Case resolucion_pantalla_ejv
        Case CTE_640X480
            frm_z0_inic.List2.AddItem CTEm_640X480
        Case CTE_800X600OSUPERIOR
            frm_z0_inic.List2.AddItem CTEm_800X600OSUPERIOR
        Case Else
            MsgBox "Error: el parámetro " & resolucion_pantalla_ejv & " es incorrecto. Se va a tomar el valor por defecto " & CTEm_800X600OSUPERIOR & ". Se recomienda activar la opción de guardar el fichero de configuración para evitar este error.", vbCritical
            resolucion_pantalla_ejv = CTE_800X600OSUPERIOR
            frm_z0_inic.List2.AddItem CTEm_800X600OSUPERIOR
    End Select
    
    '10: GRABAR CONFIGURACION
    frm_z0_inic.List1.AddItem "Grabar configuración al salir:"
    If grabar_configuracion_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '11: GRABAR CONFIG POR DEFECTO
    frm_z0_inic.List1.AddItem "Grabar config. por defecto al salir:"
    If grabar_config_defecto_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '12: GRABAR LOG
    frm_z0_inic.List1.AddItem "Grabar LOG:"
    If grabar_log_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '13: FICHERO LOG
    frm_z0_inic.List1.AddItem "Fichero LOG:"
    frm_z0_inic.List2.AddItem fichero_log_ejv
    
    '14: GRABAR RESUMEN TXT
    frm_z0_inic.List1.AddItem "Grabar resumen TXT:"
    If grabar_resumen_txt_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '15: FICHERO RESUMEN TXT
    frm_z0_inic.List1.AddItem "Fichero resumen TXT:"
    frm_z0_inic.List2.AddItem fichero_resumen_txt_ejv
    
    '16: GRABAR RESUMEN EXCEL
    frm_z0_inic.List1.AddItem "Grabar resumen XLS:"
    If grabar_resumen_xls_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '17: FICHERO RESUMEN EXCEL
    frm_z0_inic.List1.AddItem "Fichero resumen XLS:"
    frm_z0_inic.List2.AddItem fichero_resumen_xls_ejv
    
    '18: REEMPLAZAR FICHEROS EXISTENTES
    frm_z0_inic.List1.AddItem "Reemplazar Ficheros Existentes:"
    If reemplazar_fic_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '19: AUTOMATICO
    frm_z0_inic.List1.AddItem "Ejecución en modo automático:"
    If automatico_ejv Then
        frm_z0_inic.List2.AddItem CTE_txtTRUE
    Else
        frm_z0_inic.List2.AddItem CTE_txtFALSE
    End If
    
    '20: FICHERO AUTOMATICO
    For indice_auto = 1 To num_ficheros_aut_ejv
        frm_z0_inic.List1.AddItem "Fichero automático:"
        If fichero_aut_ejv(indice_auto) = "" Or fichero_aut_ejv(indice_auto) = CTE_NOHAY Then
            fichero_aut_ejv(indice_auto) = CTEm_NOHAY
        End If
        frm_z0_inic.List2.AddItem fichero_aut_ejv(indice_auto)
    Next indice_auto
    
    
    
End Sub


Sub s_fin_todo()
    
    'Grabar por defecto implica ya grabar
    If grabar_config_defecto_ejv Then
        s_aut_grabar_inicio_txt True
    Else
        If grabar_configuracion_ejv Then
            s_aut_grabar_inicio_txt False
        End If
    End If
    
    'Cierro los ficheros de resumen si estaban abiertos
    s_cerrar_ficheros_globales_ejv
    'Fin de todo
    Reset 'cierro ficheros si habia abiertos
    End   'bye
    
    
    
End Sub
Sub s_click_zoom_ejv(formulario As Object)

    Dim tipo_zoom As Integer
    Dim recipiente_va As Boolean
    
    If habilitar_change_zoom_va0 Then
    
        tipo_zoom = formulario.Cb_Zoom.ListIndex
        
        'Si esta abierto el editor de mapas...
        If formulario.Name = "frm_u0_font" Then
            recipiente_va = True
        Else
            Select Case frm_z0_mdi.ActiveForm.Name
                Case "frm_a0_va"
                    recipiente_va = True
                Case "frm_a0_mapa"
                    recipiente_va = False
                Case "frm_z0_graf"
                    recipiente_va = True
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End Select
        End If
        
        If recipiente_va Then
            '-------------------------
            'Cambio el zoom de VA
            ver_zoom_va0 = tipo_zoom
            'Muestro el nuevo mapa
            frm_a0_va.Cls
            s_cargar_tipo_zoom_va0
            s_fijar_separacion_mapa_va0
            s_mapa_pintar_bordes_va0 frm_a0_va
            's_mostrar_mapa_actual_va0 False
            's_operacion_ver_ejv CTE_VER_ESTADO_EJECUCION
            'Refresco el mapa
            If formulario.Name <> "frm_u0_font" Then
                s_operacion_ver_ejv CTE_VER_REFRESCAR
            End If
            '-------------------------
        Else
            '-------------------------
            'Cambio el zoom de MA
            ver_zoom_ma0 = tipo_zoom
            'Muestro el nuevo mapa
            s_fijar_separacion_mapa_ma0
            s_refrescar_mapa_actual_ma0
            '-------------------------
        End If
    End If

End Sub

Sub s_tecla_pulsada_ejv(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        'Key Code Constants
        'vbKeyLButton    1   Left mouse button
        'vbKeyRButton    2   Right mouse button
        'vbKeyCancel 3   CANCEL key
        'vbKeyMButton    4   Middle mouse button
        'vbKeyBack   8   BACKSPACE key
        'vbKeyTab    9   TAB key
        'vbKeyClear  12  CLEAR key
        'vbKeyReturn 13  ENTER key
        'vbKeyShift  16  SHIFT key
        'vbKeyControl    17  CTRL key
        'vbKeyMenu   18  MENU key
        'vbKeyPause  19  PAUSE key
        'vbKeyCapital    20  CAPS LOCK key
        'vbKeyEscape 27  ESC key
        'vbKeySpace  32  SPACEBAR key
        'vbKeyPageUp 33  PAGE UP key
        'vbKeyPageDown   34  PAGE DOWN key
        'vbKeyEnd    35  END key
        'vbKeyHome   36  HOME key
        'vbKeyLeft   37  LEFT ARROW key
        'vbKeyUp 38  UP ARROW key
        'vbKeyRight  39  RIGHT ARROW key
        'vbKeyDown   40  DOWN ARROW key
        'vbKeySelect 41  SELECT key
        'vbKeyPrint  42  PRINT SCREEN key
        'vbKeyExecute    43  EXECUTE key
        'vbKeySnapshot   44  SNAPSHOT key
        'vbKeyInsert 45  INS key
        'vbKeyDelete 46  DEL key
        'vbKeyHelp   47  HELP key
        'vbKeyNumlock    144 NUM LOCK key
        
Dim kk_z As Double
Dim kk_x As Double
Dim kk_y As Double
Dim i As Integer


        
        Case vbKeyBack   '8   BACKSPACE key
            If num_prg_activo_ejv = CTE_NINGUNO Then Exit Sub
            If num_prg_activo_ejv = CTE_HYP Then
                kk_z = 0
                kk_y = 2.6
                kk_x = 0
                
                frm_a0_va.Cls
                For i = 1 To 12
                    Debug.Print "iteracion: " & i
                    Debug.Print "Posicion: " & kk_y & "  " & kk_x
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, kk_z, 15 + kk_y, 20 + kk_x, CTE_ESFERA, cct_ejv(CTE_NEGRO), cct_ejv(i), CTE_DIRECC_NINGUNA, CTE_ZOOM_DETALLE, 1
                    s_girar_sobre_eje "Z", CTE_1VUELTA / 12, kk_z, kk_y, kk_x
                Next i
                Beep
                
                's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, 1 ,1 ,1, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(fi_azar1(nct_i_ejv)), CTE_8_N, CTE_ZOOM_SUPER3D, 3
                's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, 1, 5, 5, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(fi_azar1(nct_i_ejv)), CTE_8_E, CTE_ZOOM_SUPER3D, 3
                's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, 1, 1, 5, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(fi_azar1(nct_i_ejv)), CTE_8_S, CTE_ZOOM_SUPER3D, 3
                's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, 1, 5, 1, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(fi_azar1(nct_i_ejv)), CTE_8_O, CTE_ZOOM_SUPER3D, 3
            
                's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, 1, 22, 5, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(fi_azar1(nct_i_ejv)), CTE_8_NE, CTE_ZOOM_SUPER3D, 3
                's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, 1, 11, 11, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(fi_azar1(nct_i_ejv)), CTE_8_SE, CTE_ZOOM_SUPER3D, 3
                's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, 1, 5, 11, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(fi_azar1(nct_i_ejv)), CTE_8_SO, CTE_ZOOM_SUPER3D, 3
                's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, 1, 11, 5, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(fi_azar1(nct_i_ejv)), CTE_8_NO, CTE_ZOOM_SUPER3D, 3
            End If
        
        Case vbKeyControl   '17  CTRL key
            If num_prg_activo_ejv = CTE_NINGUNO Then Exit Sub
            If Shift = 3 Then
                'Elijo path por defecto
                nombre_fichero_ejv = path_largo_ejv(CTE_C_RAIZ)
                nombre_fichero_ejv_es_solo_un_path_ejv = True
                'Elijo fichero
                tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
                frm_z0_fic.Caption = "Fichero de fondo"  'Esto provoca la llamada, igual que un show
                frm_z0_fic.Aceptar.Caption = "&Abrir"
                frm_z0_fic.File1.Pattern = "*.*"
                frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
                frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
                If Not cancelar_operacion_fichero_ejv Then
                    On Error Resume Next
                    frm_a0_va.Picture = LoadPicture(nombre_fichero_ejv)
                End If
            End If
        'Case 44     'se pulsó [ImprPant]
        Case vbKeyEscape '27  ESC key
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    Unload frm_a0_va
                Case CTE_PAL '2
                    Unload frm_b2_pal
                Case CTE_3R '3
                    Unload frm_c0_ce
                Case CTE_PRI '4
                    Unload frm_a0_va
                Case CTE_CEL '5
                    Unload frm_a0_va
                Case CTE_GAI '6
                    Unload frm_a0_va
                Case CTE_EXP '7
                    Unload frm_a0_va
                Case CTE_CAD '8
                    Unload frm_c0_ce
                Case CTE_PEZ '9
                    Unload frm_a0_va
                Case CTE_UVA '10
                    Unload frm_a0_va
                Case CTE_YXY '11
                    Unload frm_a0_va
                Case Else
                    'no descargo nada
            End Select
        Case vbKeyF1
            If Shift = 2 Then
                s_click_programa_ejv CTE_HYP
            Else
                s_mostrar_docum_html_ejv
            End If
        Case vbKeyF2
            If Shift = 2 Then
                s_click_programa_ejv CTE_PAL
            Else
                If frm_z0_mdi.mn_Abrir.Enabled Then
                    s_accion_ficheros_va0 CTE_FIC_ABRIR
                End If
            End If
        Case vbKeyF3
            If Shift = 2 Then
                s_click_programa_ejv CTE_3R
            End If
        Case vbKeyF4
            If Shift = 2 Then
                s_click_programa_ejv CTE_PRI
            End If
        Case vbKeyF5
            If Shift = 2 Then
                s_click_programa_ejv CTE_CEL
            Else
                If num_prg_activo_ejv <> CTE_NINGUNO Then
                    If estado_ejecutar_ejv(CTE_EXE_COMENZAR, num_prg_activo_ejv) Then
                        s_operacion_ejecutar_ejv CTE_EXE_COMENZAR
                    ElseIf estado_ejecutar_ejv(CTE_EXE_CONTINUAR, num_prg_activo_ejv) Then
                        s_operacion_ejecutar_ejv CTE_EXE_CONTINUAR
                    End If
                End If
            End If
        Case vbKeyF6
            If Shift = 2 Then
                s_click_programa_ejv CTE_GAI
            Else
                If num_prg_activo_ejv <> CTE_NINGUNO Then
                    If estado_ejecutar_ejv(CTE_EXE_PAUSA, num_prg_activo_ejv) Then
                        s_operacion_ejecutar_ejv CTE_EXE_PAUSA
                    End If
                End If
            End If
        Case vbKeyF7
            If Shift = 2 Then
                s_click_programa_ejv CTE_EXP
            Else
                If num_prg_activo_ejv <> CTE_NINGUNO Then
                    If estado_ejecutar_ejv(CTE_EXE_TERMINAR, num_prg_activo_ejv) Then
                        finalizacion_usuario_ejv = True
                        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                    End If
                End If
            End If
        Case vbKeyF8
            If Shift = 2 Then
                s_click_programa_ejv CTE_CAD
            Else
                frm_z0_graf.Show CTE_AMODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
            End If
        Case vbKeyF9
            If Shift = 2 Then
                s_click_programa_ejv CTE_PEZ
            End If
        Case vbKeyF10
            If Shift = 2 Then
                s_click_programa_ejv CTE_UVA
            End If
        Case vbKeyF11
            If Shift = 2 Then
                s_click_programa_ejv CTE_YXY
            End If
        'Case vbKeyF12
        'Case vbKeyF13
        'Case vbKeyF14
        'Case vbKeyF15
        'Case vbKeyF16
        Case vbKeyLeft   'LEFT ARROW key
            If Screen.ActiveForm.Name = "frm_a0_mapa" Then
                s_mover_cursor CTE_IZQUIERDA
            End If
        Case vbKeyUp     'UP ARROW key
            If Screen.ActiveForm.Name = "frm_a0_mapa" Then
                s_mover_cursor CTE_DEFRENTE
            End If
        Case vbKeyRight  'RIGHT ARROW key
            If Screen.ActiveForm.Name = "frm_a0_mapa" Then
                s_mover_cursor CTE_DERECHA
            End If
        Case vbKeyDown   'DOWN ARROW key
            If Screen.ActiveForm.Name = "frm_a0_mapa" Then
                s_mover_cursor CTE_ATRAS
            End If
        Case Else
            'Si es otra no hago nada, es algo que se teclea
    End Select


End Sub


Sub s_cambiar_estado_enabled_operaciones_ficheros_ejv(nuevo_estado As Boolean)

    frm_z0_mdi.mn_Abrir.Enabled = nuevo_estado
    frm_z0_mdi.H_Abrir.Enabled = nuevo_estado
    frm_z0_mdi.mn_Guardar = nuevo_estado
    frm_z0_mdi.H_Guardar.Enabled = nuevo_estado
    frm_z0_mdi.mn_GuardarComo = nuevo_estado

End Sub
Sub s_cambiar_estado_enabled_programas_todos_ejv(nuevo_estado As Boolean)

    Dim i As Integer

    s_cambiar_estado_enabled_programa_ejv CTE_HYP, nuevo_estado '1
    s_cambiar_estado_enabled_programa_ejv CTE_PAL, nuevo_estado '2
    s_cambiar_estado_enabled_programa_ejv CTE_3R, nuevo_estado '3
    s_cambiar_estado_enabled_programa_ejv CTE_PRI, nuevo_estado '4
    s_cambiar_estado_enabled_programa_ejv CTE_CEL, nuevo_estado '5
    s_cambiar_estado_enabled_programa_ejv CTE_GAI, nuevo_estado '6
    s_cambiar_estado_enabled_programa_ejv CTE_EXP, nuevo_estado '7
    s_cambiar_estado_enabled_programa_ejv CTE_CAD, nuevo_estado '8
    s_cambiar_estado_enabled_programa_ejv CTE_PEZ, nuevo_estado '9
    s_cambiar_estado_enabled_programa_ejv CTE_UVA, nuevo_estado '10
    s_cambiar_estado_enabled_programa_ejv CTE_YXY, nuevo_estado '11

    For i = 0 To num_menu_viejos_ejv - 1
        frm_z0_mdi.mn_listaviejos(i).Enabled = nuevo_estado
    Next i
    
End Sub
Sub s_cambiar_estado_enabled_programa_ejv(programa As Integer, nuevo_estado As Boolean)
    
    Select Case programa
        Case CTE_HYP '1
            frm_z0_mdi.Mn_hyp.Enabled = nuevo_estado
            frm_z0_mdi.H_HYP.Enabled = nuevo_estado
        Case CTE_PAL '2
            frm_z0_mdi.mn_palyfras.Enabled = nuevo_estado
            frm_z0_mdi.H_PAL.Enabled = nuevo_estado
        Case CTE_3R '3
            frm_z0_mdi.mn_3r.Enabled = nuevo_estado
            frm_z0_mdi.H_3R.Enabled = nuevo_estado
        Case CTE_PRI '4
            frm_z0_mdi.mn_Prisionero.Enabled = nuevo_estado
            frm_z0_mdi.H_PRI.Enabled = nuevo_estado
        Case CTE_CEL '5
            frm_z0_mdi.mn_Celdilla.Enabled = nuevo_estado
        Case CTE_GAI '6
            frm_z0_mdi.mn_gaia.Enabled = nuevo_estado
        Case CTE_EXP '7
            frm_z0_mdi.mn_Explorando_Mapas.Enabled = nuevo_estado
        Case CTE_CAD '8
            frm_z0_mdi.mn_Cadenas.Enabled = nuevo_estado
        Case CTE_PEZ '9
            frm_z0_mdi.mn_Peces.Enabled = nuevo_estado
        Case CTE_UVA '10
            frm_z0_mdi.mn_Universo.Enabled = nuevo_estado
        Case CTE_YXY '11
            frm_z0_mdi.mn_yxy.Enabled = nuevo_estado
        Case Else
            s_error_num_prog programa
    End Select

End Sub

Sub s_cambiar_estado_enabled_ejecutar_ejv(operacion As Integer, nuevo_estado As Boolean)

    Select Case operacion
        Case CTE_EXE_COMENZAR
            frm_z0_mdi.mn_Comenzar.Enabled = nuevo_estado
            If Not frm_z0_mdi.mn_Continuar.Enabled Then
                frm_z0_mdi.H_Comenzar.Enabled = nuevo_estado
            End If
        Case CTE_EXE_CONTINUAR
            frm_z0_mdi.mn_Continuar.Enabled = nuevo_estado
            If Not frm_z0_mdi.mn_Comenzar.Enabled Then
                frm_z0_mdi.H_Comenzar.Enabled = nuevo_estado
            End If
        Case CTE_EXE_PAUSA
            frm_z0_mdi.mn_Pausa.Enabled = nuevo_estado
            frm_z0_mdi.H_Pausa.Enabled = nuevo_estado
        Case CTE_EXE_TERMINAR
            frm_z0_mdi.mn_terminar.Enabled = nuevo_estado
            frm_z0_mdi.H_Terminar.Enabled = nuevo_estado
        Case Else
            MsgBox "No existe esa operación de ejecución", vbError
    End Select
    If operacion = CTE_EXE_COMENZAR Or operacion = CTE_EXE_CONTINUAR Then
        If nuevo_estado Then
            frm_z0_mdi.H_Comenzar.Enabled = True
        End If
    End If
    If num_prg_activo_ejv <> CTE_NINGUNO Then
        estado_ejecutar_ejv(operacion, num_prg_activo_ejv) = nuevo_estado
    End If
End Sub
Sub s_cambiar_estado_enabled_menus_ejv(operacion As Integer, nuevo_estado As Boolean)

    Select Case operacion
        Case CTE_VER_OPCIONES1
            frm_z0_mdi.mn_Opciones1.Enabled = nuevo_estado
            frm_z0_mdi.H_Opciones.Enabled = nuevo_estado
        Case CTE_VER_OPCIONES2
            frm_z0_mdi.mn_Opciones2.Enabled = nuevo_estado
        Case CTE_VER_OPCIONES3
            frm_z0_mdi.mn_Opciones3.Enabled = nuevo_estado
        Case CTE_VER_TIPOS_AGENTES
            frm_z0_mdi.mn_Tipos_Agentes.Enabled = nuevo_estado
            frm_z0_mdi.H_Agentes.Enabled = nuevo_estado
        Case CTE_VER_MAPA
            frm_z0_mdi.mn_Mapa.Enabled = nuevo_estado
            'frm_z0_mdi.H_Mapa.Enabled = nuevo_estado esto se queda asi por el tema de que sirve para 2 cosas
        Case CTE_VER_TIPO_EVOLUCION
            frm_z0_mdi.mn_TipoEvolucion.Enabled = nuevo_estado
            Case CTE_VER_TIPO_EVOLUCION_EVALUACION
                frm_z0_mdi.mn_Metodo_Evaluacion.Enabled = nuevo_estado
            Case CTE_VER_TIPO_EVOLUCION_SELECCION
                frm_z0_mdi.mn_Metodo_Seleccion.Enabled = nuevo_estado
            Case CTE_VER_TIPO_EVOLUCION_REPRODUCCION
                frm_z0_mdi.mn_Metodo_Reproduccion.Enabled = nuevo_estado
                Case CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES
                    frm_z0_mdi.mn_Tipo_Mutaciones.Enabled = nuevo_estado
                Case CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO
                    frm_z0_mdi.mn_Tipo_Sobrecruzamiento.Enabled = nuevo_estado
        Case CTE_VER_APELLIDOS
            frm_z0_mdi.mn_Apellidos.Enabled = nuevo_estado
        Case CTE_VER_REFRESCAR
            frm_z0_mdi.mn_Refrescar.Enabled = nuevo_estado
            frm_z0_mdi.H_Refrescar.Enabled = nuevo_estado
        Case CTE_VER_ESTADO_EJECUCION
            frm_z0_mdi.mn_EstadoEjecucion.Enabled = nuevo_estado
            frm_z0_mdi.H_Estado.Enabled = nuevo_estado
        Case CTE_VER_AGENTES_TODOS
            frm_z0_mdi.mn_ListaAgentes.Enabled = nuevo_estado
        Case CTE_VER_AGENTES_MEJORES
            frm_z0_mdi.mn_MejoresAgentes.Enabled = nuevo_estado
        Case CTE_VER_DICCIONARIO
            frm_z0_mdi.mn_Diccionario.Enabled = nuevo_estado
        Case CTE_VER_JUGAR_CONTRA_ORDENADOR
            frm_z0_mdi.mn_JugarContraOrdenador.Enabled = nuevo_estado
        Case CTE_VER_MODIFICAR_AGENTE
            frm_z0_mdi.mn_ModificarAgente.Enabled = nuevo_estado
        Case CTE_VER_GRAFICO
            frm_z0_mdi.mn_Grafico.Enabled = nuevo_estado
            'frm_z0_mdi.H_Grafico.Enabled = nuevo_estado esto se queda asi por el tema de que sirve para 2 cosas
        Case Else
            MsgBox "No existe esa operación de ver", vbInformation
    End Select
    If num_prg_activo_ejv <> CTE_NINGUNO Then
        estado_ver_ejv(operacion, num_prg_activo_ejv) = nuevo_estado
    End If

End Sub
Sub s_estado_enabled_ejecucion_ejv()

    frm_z0_mdi.mn_Comenzar.Enabled = estado_ejecutar_ejv(CTE_EXE_COMENZAR, num_prg_activo_ejv)
    frm_z0_mdi.mn_Continuar.Enabled = estado_ejecutar_ejv(CTE_EXE_CONTINUAR, num_prg_activo_ejv)
    If estado_ejecutar_ejv(CTE_EXE_COMENZAR, num_prg_activo_ejv) Or estado_ejecutar_ejv(CTE_EXE_CONTINUAR, num_prg_activo_ejv) Then
        frm_z0_mdi.H_Comenzar.Enabled = True
    Else
        frm_z0_mdi.H_Comenzar.Enabled = False
    End If
    
    frm_z0_mdi.mn_Pausa.Enabled = estado_ejecutar_ejv(CTE_EXE_PAUSA, num_prg_activo_ejv)
    frm_z0_mdi.H_Pausa.Enabled = estado_ejecutar_ejv(CTE_EXE_PAUSA, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_terminar.Enabled = estado_ejecutar_ejv(CTE_EXE_TERMINAR, num_prg_activo_ejv)
    frm_z0_mdi.H_Terminar.Enabled = estado_ejecutar_ejv(CTE_EXE_TERMINAR, num_prg_activo_ejv)
    
End Sub


Sub s_estado_enabled_ver_ejv()

    frm_z0_mdi.mn_Opciones1.Enabled = estado_ver_ejv(CTE_VER_OPCIONES1, num_prg_activo_ejv)
    frm_z0_mdi.H_Opciones.Enabled = estado_ver_ejv(CTE_VER_OPCIONES1, num_prg_activo_ejv)
    frm_z0_mdi.mn_Opciones2.Enabled = estado_ver_ejv(CTE_VER_OPCIONES2, num_prg_activo_ejv)
    frm_z0_mdi.mn_Opciones3.Enabled = estado_ver_ejv(CTE_VER_OPCIONES3, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_Tipos_Agentes.Enabled = estado_ver_ejv(CTE_VER_TIPOS_AGENTES, num_prg_activo_ejv)
    frm_z0_mdi.H_Agentes.Enabled = estado_ver_ejv(CTE_VER_TIPOS_AGENTES, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_Mapa.Enabled = estado_ver_ejv(CTE_VER_MAPA, num_prg_activo_ejv)
    frm_z0_mdi.H_Mapa.Enabled = estado_ver_ejv(CTE_VER_MAPA, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_TipoEvolucion.Enabled = estado_ver_ejv(CTE_VER_TIPO_EVOLUCION, num_prg_activo_ejv)
        frm_z0_mdi.mn_Metodo_Evaluacion.Enabled = estado_ver_ejv(CTE_VER_TIPO_EVOLUCION_EVALUACION, num_prg_activo_ejv)
        frm_z0_mdi.mn_Metodo_Seleccion.Enabled = estado_ver_ejv(CTE_VER_TIPO_EVOLUCION_SELECCION, num_prg_activo_ejv)
        frm_z0_mdi.mn_Metodo_Reproduccion.Enabled = estado_ver_ejv(CTE_VER_TIPO_EVOLUCION_REPRODUCCION, num_prg_activo_ejv)
            frm_z0_mdi.mn_Tipo_Mutaciones.Enabled = estado_ver_ejv(CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES, num_prg_activo_ejv)
            frm_z0_mdi.mn_Tipo_Sobrecruzamiento.Enabled = estado_ver_ejv(CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_Apellidos.Enabled = estado_ver_ejv(CTE_VER_APELLIDOS, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_Refrescar.Enabled = estado_ver_ejv(CTE_VER_REFRESCAR, num_prg_activo_ejv)
    frm_z0_mdi.H_Refrescar.Enabled = estado_ver_ejv(CTE_VER_REFRESCAR, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_EstadoEjecucion.Enabled = estado_ver_ejv(CTE_VER_ESTADO_EJECUCION, num_prg_activo_ejv)
    frm_z0_mdi.H_Estado.Enabled = estado_ver_ejv(CTE_VER_ESTADO_EJECUCION, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_ListaAgentes.Enabled = estado_ver_ejv(CTE_VER_AGENTES_TODOS, num_prg_activo_ejv)
    frm_z0_mdi.mn_MejoresAgentes.Enabled = estado_ver_ejv(CTE_VER_AGENTES_MEJORES, num_prg_activo_ejv)
    frm_z0_mdi.mn_Diccionario.Enabled = estado_ver_ejv(CTE_VER_DICCIONARIO, num_prg_activo_ejv)
    frm_z0_mdi.mn_JugarContraOrdenador.Enabled = estado_ver_ejv(CTE_VER_JUGAR_CONTRA_ORDENADOR, num_prg_activo_ejv)
    frm_z0_mdi.mn_ModificarAgente.Enabled = estado_ver_ejv(CTE_VER_MODIFICAR_AGENTE, num_prg_activo_ejv)
    
    frm_z0_mdi.mn_Grafico.Enabled = estado_ver_ejv(CTE_VER_GRAFICO, num_prg_activo_ejv)
    frm_z0_mdi.H_Grafico.Enabled = estado_ver_ejv(CTE_VER_GRAFICO, num_prg_activo_ejv)

End Sub


Sub s_informar_parametro_config_ejv(parametro As Integer)

    Select Case parametro
        '1: VERSION
        Case 1
            frm_z0_inic.mensaje = "Para modificar este parámetro, edite el fichero " & CTE_nombreINICIO_TXT & "."
        '2: IDIOMA
        Case 2
            frm_z0_inic.mensaje = "Elección de idioma."
        '3: ELEGIR IDIOMA
        Case 3
            frm_z0_inic.mensaje = "Este parámetro determina si aparece o no una ventana de elección de idioma al arrancar el programa."
        '4: CONTROL DE ERRORES
        Case 4
            frm_z0_inic.mensaje = "El modo control de errores sirve para detectar errores en la programación. Si se fija el modo control de errores con el valor True, la ejecución del programa puede ser mucho mas lenta. Se aconseja fijar este parámetro con el valor False"
        '5: MOSTRAR_LOGO
        Case 5
            frm_z0_inic.mensaje = "Este parámetro determina si aparece o no el logotipo al arrancar el programa. Si se elige False, el arranque del programa será mucho más rapido."
        '6: ALGORITMO DE ORDENACION
        Case 6
            frm_z0_inic.mensaje = "Este parámetro no se puede modificar."
        '7: CTE_SISTEMA_OPERATIVO
        Case 7
            frm_z0_inic.mensaje = "Sistema operativo en el que está ejecutándose el programa."
        '8: CTE_PEDIR_CONFIRMACION
        Case 8
            frm_z0_inic.mensaje = "Si se activa, el programa preguntará antes de realizar acciones como sobreescibir un fichero de datos o detener una simulación."
        '9: RESOLUCION PANTALLA
        Case 9
            frm_z0_inic.mensaje = "Este programa se ha desarrollado para una resolución de 800x600 o superior y color de alta densidad (16 bits) o superior. Para cambiar estos valores en Windows, ir a Inicio-Configuración-Panel de Control-Pantalla-Configuración   "
        '10: GRABAR CONFIGURACION
        Case 10
            frm_z0_inic.mensaje = "Este parámetro indica si se va a grabar la configuración al terminar normalmente el programa."
        '11: GRABAR CONFIG POR DEFECTO
        Case 11
            frm_z0_inic.mensaje = "En el caso de activar la grabación de la configuración, este parámetro indica si se va a grabar la configuración predeterminada en vez de la que ahora existe."
        '12: GRABAR LOG
        Case 12
            frm_z0_inic.mensaje = "El .log es un pequeño fichero que contiene la secuencia de los eventos más importantes que se han producido durante la ejecución de la aplicación."
        '13: FICHERO LOG
        Case 13
            frm_z0_inic.mensaje = "El .log es un pequeño fichero que contiene la secuencia de los eventos más importantes que se han producido durante la ejecución de la aplicación."
        '14: GRABAR RESUMEN TXT
        Case 14
            frm_z0_inic.mensaje = "El .txt es un fichero que contiene un resumen general de la ejecución del programa, dando un resumen de cada uno de los .aut."
        '15: FICHERO RESUMEN TXT
        Case 15
            frm_z0_inic.mensaje = "El .txt es un fichero que contiene un resumen general de la ejecución del programa, dando un resumen de cada uno de los .aut."
        '16: GRABAR RESUMEN EXCEL
        Case 16
            frm_z0_inic.mensaje = "El .xls es un fichero Excel que contiene un resumen general de la ejecución del programa, dando un resumen de cada uno de los .aut."
        '17: FICHERO RESUMEN EXCEL
        Case 17
            frm_z0_inic.mensaje = "El .xls es un fichero Excel que contiene un resumen general de la ejecución del programa, dando un resumen de cada uno de los .aut."
        '18: REEMPLAZAR FICHEROS EXISTENTES
        Case 18
            frm_z0_inic.mensaje = "Si se elige False, se añadirá a los nombres de fichero una extensión que comienza con la cadena 00000001."
        '19: AUTOMATICO
        Case 19
            frm_z0_inic.mensaje = "El modo automático sirve para ejecutar el programa sin intervención del usuario. Si se elige ahora el valor True, la próxima vez el programa arrancará en modo automático, y para modificar este parámetro, se deberá cambiar editar el fichero " & CTE_nombreINICIO_TXT & ", por ejemplo con el block de notas (notepad)."
        '20: FICHERO AUTOMATICO
        Case Else
            frm_z0_inic.mensaje = "Cada uno de estos ficheros define una ejecución de tipo automático, sin intervención del usuario. Para modificar estas líneas en " & CTE_nombreINICIO_TXT & ", edite el propio " & CTE_nombreINICIO_TXT & " con el botón Editar. Para modificar los ficheros de tipo automático, editelos haciendo Clik en la columna derecha, en el nombre del fichero."
    End Select

End Sub
Sub s_modificar_parametro_config_ejv(parametro As Integer)

    Dim num_fichero_automatico As Long
    Dim mensaje As String

    Select Case parametro
        '1: VERSION
        Case 1
        '2: IDIOMA
        Case 2
            frm_z0_leng.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
            s_tratamiento_idioma_mdi
        '3: ELEGIR IDIOMA
        Case 3
            elegir_idioma_ejv = Not elegir_idioma_ejv
        '4: CONTROL DE ERRORES
        Case 4
            control_errores_de_programacion_ejv = Not control_errores_de_programacion_ejv
        '5: MOSTRAR_LOGO
        Case 5
            mostrar_logo_ejv = Not mostrar_logo_ejv
        '6: ALGORITMO DE ORDENACION
        Case 6
        '7: CTE_SISTEMA_OPERATIVO
        Case 7
            If sistema_operativo_ejv = CTE_WINDOWS95 Then
                sistema_operativo_ejv = CTE_WINDOWSNT
            ElseIf sistema_operativo_ejv = CTE_WINDOWSNT Then
                sistema_operativo_ejv = CTE_WINDOWS3X
            ElseIf sistema_operativo_ejv = CTE_WINDOWS3X Then
                sistema_operativo_ejv = CTE_WINDOWS95
            End If
        '8: CTE_PEDIR_CONFIRMACION
        Case 8
            pedir_confirmacion_ejv = Not pedir_confirmacion_ejv
        '9: RESOLUCION PANTALLA
        Case 9
            If resolucion_pantalla_ejv = CTE_640X480 Then
                resolucion_pantalla_ejv = CTE_800X600OSUPERIOR
            ElseIf resolucion_pantalla_ejv = CTE_800X600OSUPERIOR Then
                resolucion_pantalla_ejv = CTE_640X480
            End If
        '10: GRABAR CONFIGURACION
        Case 10
            grabar_configuracion_ejv = Not grabar_configuracion_ejv
        '11: GRABAR CONFIG POR DEFECTO
        Case 11
            grabar_config_defecto_ejv = Not grabar_config_defecto_ejv
        '12: GRABAR LOG
        Case 12
            If grabar_log_ejv Then
                mensaje = "Esta acción cerrará el fichero de log que se esta grabando en esos momentos. ¿Está seguro de que quiere cerrar el fichero de log?"
            Else
                mensaje = "Esta acción abrirá un nuevo fichero de log en este mismo momento. ¿Está seguro de que quiere abrir un nuevo fichero de log, borrando algun otro posible fichero con el mismo nombre?"
            End If
            If MsgBox(mensaje, vbQuestion + vbYesNo) = vbYes Then
                grabar_log_ejv = Not grabar_log_ejv
                If grabar_log_ejv Then
                    s_abrir_fichero_salida_ejv CTE_FIC_20_GLOLOG, CTE_ABRIR_BORRAR
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Este fichero de Log comienza a grabarse a partir de una acción de usuario que produce una modificación de la configuración"
                Else
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_20_GLOLOG, "Este fichero de Log se cierra debido a una acción de usuario que produce una modificación de la configuración"
                    s_cerrar_fichero_salida_ejv CTE_FIC_20_GLOLOG
                End If
            End If
        '13: FICHERO LOG
        Case 13
            fichero_log_ejv = InputBox("Introduza el nuevo nombre del fichero", "Seleccionar Fichero", fichero_log_ejv)
        '14: GRABAR RESUMEN TXT
        Case 14
            If grabar_resumen_txt_ejv Then
                mensaje = "Esta acción cerrará el fichero de resumen .txt que se esta grabando en esos momentos. ¿Está seguro de que quiere cerrar el fichero de resumen txt?"
            Else
                mensaje = "Esta acción abrirá un nuevo fichero de resumen .txt en este mismo momento. ¿Está seguro de que quiere abrir un nuevo fichero de resumen .txt, borrando algun otro posible fichero con el mismo nombre?"
            End If
            If MsgBox(mensaje, vbQuestion + vbYesNo) = vbYes Then
                grabar_resumen_txt_ejv = Not grabar_resumen_txt_ejv
                If grabar_resumen_txt_ejv Then
                    s_abrir_fichero_salida_ejv CTE_FIC_21_GLOTXT, CTE_ABRIR_BORRAR
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_21_GLOTXT, "Este fichero de resumen .txt comienza a grabarse a partir de una acción de usuario que produce una modificación de la configuración"
                Else
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_21_GLOTXT, "Este fichero de resumen .txt se cierra debido a una acción de usuario que produce una modificación de la configuración"
                    s_cerrar_fichero_salida_ejv CTE_FIC_21_GLOTXT
                End If
            End If
        '15: FICHERO RESUMEN TXT
        Case 15
            fichero_resumen_txt_ejv = InputBox("Introduza el nuevo nombre del fichero", "Seleccionar Fichero", fichero_resumen_txt_ejv)
        '16: GRABAR RESUMEN EXCEL
        Case 16
            If grabar_resumen_xls_ejv Then
                mensaje = "Esta acción cerrará el fichero de resumen .xls que se esta grabando en esos momentos. ¿Está seguro de que quiere cerrar el fichero de resumen xls?"
            Else
                mensaje = "Esta acción abrirá un nuevo fichero de resumen .xls en este mismo momento. ¿Está seguro de que quiere abrir un nuevo fichero de resumen .xls, borrando algun otro posible fichero con el mismo nombre?"
            End If
            If MsgBox(mensaje, vbQuestion + vbYesNo) = vbYes Then
                grabar_resumen_xls_ejv = Not grabar_resumen_xls_ejv
                If grabar_resumen_xls_ejv Then
                    s_abrir_fichero_salida_ejv CTE_FIC_22_GLOXLS, CTE_ABRIR_BORRAR
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_22_GLOXLS, "Este fichero de resumen .xls comienza a grabarse a partir de una acción de usuario que produce una modificación de la configuración", 3, 3
                Else
                    s_grabar_dato_fichero_salida_ejv CTE_FIC_22_GLOXLS, "Este fichero de resumen .xls se cierra debido a una acción de usuario que produce una modificación de la configuración", 3, 3
                    s_cerrar_fichero_salida_ejv CTE_FIC_22_GLOXLS
                End If
            End If
        '17: FICHERO RESUMEN EXCEL
        Case 17
            fichero_resumen_xls_ejv = InputBox("Introduza el nuevo nombre del fichero", "Seleccionar Fichero", fichero_resumen_xls_ejv)
        '18: REEMPLAZAR FICHEROS EXISTENTES
        Case 18
            reemplazar_fic_ejv = Not reemplazar_fic_ejv
        '19: AUTOMATICO
        Case 19
            automatico_ejv = Not automatico_ejv
        '20: FICHERO AUTOMATICO
        Case Else
            num_fichero_automatico = parametro - CTE_INDICE_ULTIMO_PARAMETRO_REPETIDO + 1
            If fichero_aut_ejv(num_fichero_automatico) <> "" And fichero_aut_ejv(num_fichero_automatico) <> CTE_NOHAY And fichero_aut_ejv(num_fichero_automatico) <> CTEm_NOHAY Then
                'Es 19 o mayor de 19, es uno de auto
                'Creo un nuevo formulario como este mismo
                s_mostrar_auto_ejv num_fichero_automatico
            End If
    End Select
    s_mostrar_config_ejv

End Sub

Sub s_operacion_ejecutar_ejv(operacion As Integer)

    Dim txt As String
    
    mostrar_aviso_imagen_ejv = False
    s_mostrar_aviso_imagen
    
    'Al detener/parar o continuar/finalizar se cierra/abre siempre los 3 ficheros de resumen de un ejemplo
    'pero para cerrar, en vez de hacerlo aqui se hace la final del bucle_general porque
    'se detiene con doevents cuando a veces todavia le falta algo por escribir en el fichero
    
    Select Case operacion
        Case CTE_EXE_COMENZAR
            s_activar_opciones_generales_ejv
            s_abrir_ficheros_un_ejemplo_ejv CTE_ABRIR_BORRAR
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_mostrar_estado_semaforo frm_a1_inhyp, CTE_FUNCIONANDO
                    s_comenzar_va0
                Case CTE_PAL '2
                    s_mostrar_estado_semaforo frm_b2_inpal, CTE_FUNCIONANDO
                    s_comenzar_pal
                Case CTE_3R '3
                    s_mostrar_estado_semaforo frm_c3_in3r, CTE_FUNCIONANDO
                    s_comenzar_ce0
                Case CTE_PRI '4
                    s_mostrar_estado_semaforo frm_a4_inpri, CTE_FUNCIONANDO
                    s_comenzar_va0
                Case CTE_CEL '5
                    s_mostrar_estado_semaforo frm_a5_incel, CTE_FUNCIONANDO
                    s_comenzar_va0
                Case CTE_GAI '6
                    s_mostrar_estado_semaforo frm_a6_ingaia, CTE_FUNCIONANDO
                    GI_modo_de_ejecucion = 4
                    s_comenzar_gai
                Case CTE_EXP '7
                    s_mostrar_estado_semaforo frm_a7_inexp, CTE_FUNCIONANDO
                    s_comenzar_va0
                Case CTE_CAD '8
                    s_mostrar_estado_semaforo frm_c8_incad, CTE_FUNCIONANDO
                    s_comenzar_ce0
                Case CTE_PEZ '9
                    s_mostrar_estado_semaforo frm_a9_inpez, CTE_FUNCIONANDO
                    s_comenzar_va0
                Case CTE_UVA '10
                    s_mostrar_estado_semaforo frm_aA_inuva, CTE_FUNCIONANDO
                    s_comenzar_va0
                Case CTE_YXY '11
                    s_comenzar_va0
                Case Else
                    s_error_num_prog num_prg_activo_ejv
             End Select
        Case CTE_EXE_CONTINUAR
            s_activar_opciones_generales_ejv
            s_abrir_ficheros_un_ejemplo_ejv CTE_ABRIR_ANEXAR
            s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
            s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, True
            s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_bucle_general_va0
                Case CTE_PAL '2
                    s_comenzar_pal
                Case CTE_3R '3
                    s_bucle_general_ce0
                Case CTE_PRI '4
                    s_bucle_general_va0
                Case CTE_CEL '5
                    s_bucle_general_va0
                Case CTE_GAI '6
                    s_comenzar_gai
                Case CTE_EXP '7
                    s_bucle_general_va0
                Case CTE_CAD '8
                    s_bucle_general_ce0
                Case CTE_PEZ '9
                    s_bucle_general_va0
                Case CTE_UVA '10
                    s_bucle_general_va0
                Case CTE_YXY '11
                    s_bucle_general_va0
                Case Else
                    s_error_num_prog num_prg_activo_ejv
                    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, True
                    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, False
                    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, False
            End Select
        Case CTE_EXE_PAUSA
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_terminar_va0 False
                Case CTE_PAL '2
                    s_terminar_pal
                Case CTE_3R '3
                    s_terminar_ce0 False
                Case CTE_PRI '4
                    s_terminar_va0 False
                Case CTE_CEL '5
                    s_terminar_va0 False
                Case CTE_GAI '6
                    's_estado_detenido_gai
                    GI_Finalizar = True
                    s_terminar_va0 False
                Case CTE_EXP '7
                    s_terminar_va0 False
                Case CTE_CAD '8
                    s_terminar_ce0 False
                Case CTE_PEZ '9
                    s_terminar_va0 False
                Case CTE_UVA '10
                    s_terminar_va0 False
                Case CTE_YXY '11
                    s_terminar_va0 False
                Case Else
                    s_error_num_prog num_prg_activo_ejv
            End Select
            'Grabo y cierro el fichero actual, pero lo hago al final de bucle_general
        Case CTE_EXE_TERMINAR
            If Not automatico_ejv And finalizacion_usuario_ejv And pedir_confirmacion_ejv Then
                txt = "¿Está seguro de que desea terminar el proceso?"
                If idioma_ejv = CTE_INGLES Then
                    txt = "¿Are you sure that you want to finish the process?"
                End If
                If MsgBox(txt, vbQuestion + vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            mostrar_aviso_imagen_ejv = True
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_terminar_va0 True
                Case CTE_PAL '2
                    s_terminar_pal
                Case CTE_3R '3
                    s_terminar_ce0 True
                Case CTE_PRI '4
                    s_terminar_va0 True
                Case CTE_CEL '5
                    s_terminar_va0 True
                Case CTE_GAI '6
                    's_estado_detenido_gai
                    GI_Finalizar = True
                    s_terminar_va0 True
                Case CTE_EXP '7
                    s_terminar_va0 True
                Case CTE_CAD '8
                    s_terminar_ce0 True
                Case CTE_PEZ '9
                    s_terminar_va0 True
                Case CTE_UVA '10
                    s_terminar_va0 True
                Case CTE_YXY '11
                    s_terminar_va0 True
                Case Else
                    s_error_num_prog num_prg_activo_ejv
            End Select
            'Grabo y cierro el fichero actual, pero lo hago al final de bucle_general
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "No existe esa operación de ejecución"
    End Select


End Sub

Sub s_operacion_ver_ejv(operacion As Integer)

    Select Case operacion
        Case CTE_VER_OPCIONES1
            If num_prg_activo_ejv = CTE_NINGUNO Then
                MsgBox "No hay ningún programa activo en este momento. Algunas de las opciones que se modifiquen ahora se mantendrán hasta que se elija un nuevo programa (por ejemplo ""Hormigas y Plantas"") y entonces podrán ser sustituidas por las opciones por defecto del ejemplo elegido. Para activar un programa, elegir uno del menú ""Ejemplos de Vida"".", vbInformation
                frm_z0_op.Caption = "Opciones Generales de Ejemplos de Vida - (No hay programa activo)"
            Else
                frm_z0_op.Caption = "Opciones Generales de Ejemplos de Vida - " & nombre_programa_ejv(num_prg_activo_ejv)
            End If
            frm_z0_op.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case CTE_VER_OPCIONES2
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_ver_opciones_va0
                Case CTE_PAL '2
                    s_ver_opciones_pal
                Case CTE_3R '3
                    s_ver_opciones_ce0
                Case CTE_PRI '4
                    s_ver_opciones_va0
                Case CTE_CEL '5
                    s_ver_opciones_va0
                Case CTE_GAI '6
                Case CTE_EXP '7
                    s_ver_opciones_va0
                Case CTE_CAD '8
                    s_ver_opciones_ce0
                Case CTE_PEZ '9
                    s_ver_opciones_va0
                Case CTE_UVA '10
                    s_ver_opciones_va0
                Case CTE_YXY '11
                    s_ver_opciones_va0
                Case Else
                    s_error_num_prog num_prg_activo_ejv
             End Select
        Case CTE_VER_OPCIONES3 'Opciones específicas del programa ejemplo
            Select Case num_prg_activo_ejv
                Case CTE_HYP
                    'aqui habria que actualizar la suma de todos los tipos
                    s_centrar_ventana_ejv frm_a1_ophyp
                    frm_a1_ophyp.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
                Case CTE_3R
                    s_centrar_ventana_ejv frm_c3_op3r
                    frm_c3_op3r.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
                Case CTE_PRI
                    s_centrar_ventana_ejv frm_a4_oppri
                    frm_a4_oppri.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
                Case CTE_EXP
                    s_centrar_ventana_ejv frm_a7_opexp
                    frm_a7_opexp.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
                Case CTE_UVA
                    s_centrar_ventana_ejv frm_aA_opuva
                    frm_aA_opuva.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
                Case Else
                    s_error_num_prog num_prg_activo_ejv
            End Select
            
            If ciclo_ejv > 0 Then
                s_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, True
            End If
        
        Case CTE_VER_TIPOS_AGENTES '2
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_ver_tipos_va0
                Case CTE_PAL '2
                Case CTE_3R '3
                Case CTE_PRI '4
                    s_ver_tipos_va0
                Case CTE_CEL '5
                    s_ver_tipos_va0
                Case CTE_GAI '6
                    s_ver_tipos_va0
                Case CTE_EXP '7
                    s_ver_tipos_va0
                Case CTE_CAD '8
                Case CTE_PEZ '9
                    s_ver_tipos_va0
                Case CTE_UVA '10
                    s_ver_tipos_va0
                Case CTE_YXY '11
                    s_ver_tipos_va0
                Case Else
                    s_error_num_prog num_prg_activo_ejv
             End Select
        Case CTE_VER_MAPA '3
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_ver_mapa_va0
                Case CTE_PAL '2
                Case CTE_3R '3
                Case CTE_PRI '4
                    s_ver_mapa_va0
                Case CTE_CEL '5
                    s_ver_mapa_va0
                Case CTE_GAI '6
                Case CTE_EXP '7
                    s_ver_mapa_va0
                Case CTE_CAD '8
                Case CTE_PEZ '9
                    s_ver_mapa_va0
                Case CTE_UVA '10
                    s_ver_mapa_va0
                Case CTE_YXY '11
                    s_ver_mapa_va0
                Case Else
                    s_error_num_prog num_prg_activo_ejv
             End Select
        Case CTE_VER_TIPO_EVOLUCION '4
            s_error_ejv CON_OPCION_FINALIZAR, "Torpedo"
        Case CTE_VER_TIPO_EVOLUCION_EVALUACION '5
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                Case CTE_PAL '2
                Case CTE_3R '3
                    frm_c3_ev3r.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
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
        Case CTE_VER_TIPO_EVOLUCION_SELECCION '6
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                Case CTE_PAL '2
                Case CTE_3R '3
                    frm_c3_sel3r.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
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
        Case CTE_VER_TIPO_EVOLUCION_REPRODUCCION '7
            s_error_ejv CON_OPCION_FINALIZAR, "Torpedo"
        Case CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES '8
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                Case CTE_PAL '2
                Case CTE_3R '3
                    frm_c3_rm.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
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
        Case CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO '9
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                Case CTE_PAL '2
                Case CTE_3R '3
                    frm_c3_rs.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
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
        Case CTE_VER_APELLIDOS '10
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_mostrar_apellidos_va0
                Case CTE_PAL '2
                Case CTE_3R '3
                Case CTE_PRI '4
                    s_mostrar_apellidos_va0
                Case CTE_CEL '5
                Case CTE_GAI '6
                Case CTE_EXP '7
                Case CTE_CAD '8
                Case CTE_PEZ '9
                    s_mostrar_apellidos_va0
                Case CTE_UVA '10
                Case CTE_YXY '11
                Case Else
                    s_error_num_prog num_prg_activo_ejv
             End Select
        Case CTE_VER_REFRESCAR '11
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_ver_refrescar_va0
                    'frm_a1_inhyp.Show
                    'frm_a1_inhyp.SetFocus
                Case CTE_PAL '2
                Case CTE_3R '3
                Case CTE_PRI '4
                    s_ver_refrescar_va0
                    'frm_a4_inpri.Show
                    'frm_a4_inpri.SetFocus
                Case CTE_CEL '5
                Case CTE_GAI '6
                Case CTE_EXP '7
                    s_ver_refrescar_va0
                    'frm_a7_inexp.Show
                    'frm_a7_inexp.SetFocus
                Case CTE_CAD '8
                Case CTE_PEZ '9
                    s_ver_refrescar_va0
                    'frm_a9_inpez.Show
                    'frm_a9_inpez.SetFocus
                Case CTE_UVA '10
                    s_ver_refrescar_va0
                Case CTE_YXY '11
                    s_ver_refrescar_va0
                Case Else
                    s_error_num_prog num_prg_activo_ejv
             End Select
        Case CTE_VER_ESTADO_EJECUCION '12
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    frm_a1_inhyp.Show CTE_AMODAL
                    frm_a1_inhyp.Caption = "Información Hormigas"
                    s_mostrar_estado_semaforo frm_a1_inhyp, CTE_FUNCIONANDO
                Case CTE_PAL '2
                    frm_b2_inpal.Show CTE_AMODAL
                    frm_b2_inpal.Caption = "Información Frases"
                    s_mostrar_estado_semaforo frm_b2_inpal, CTE_FUNCIONANDO
                Case CTE_3R '3
                    frm_c3_in3r.Show CTE_AMODAL
                    frm_c3_in3r.Caption = "Información Jugadores"
                    s_mostrar_estado_semaforo frm_c3_in3r, CTE_FUNCIONANDO
                Case CTE_PRI '4
                    frm_a4_inpri.Show CTE_AMODAL
                    frm_a4_inpri.Caption = "Información Prisionero"
                    s_mostrar_estado_semaforo frm_a4_inpri, CTE_FUNCIONANDO
                Case CTE_CEL '5
                    frm_a5_incel.Show CTE_AMODAL
                    frm_a5_incel.Caption = "Información Celdilla"
                    s_mostrar_estado_semaforo frm_a5_incel, CTE_FUNCIONANDO
                Case CTE_GAI '6
                    frm_a6_ingaia.Show CTE_AMODAL
                    frm_a6_ingaia.Caption = "Información Gaia"
                    s_mostrar_estado_semaforo frm_a6_ingaia, CTE_FUNCIONANDO
                Case CTE_EXP '7
                    frm_a7_inexp.Show CTE_AMODAL
                    frm_a7_inexp.Caption = "Información Exploradores"
                    s_mostrar_estado_semaforo frm_a7_inexp, CTE_FUNCIONANDO
                Case CTE_CAD '8
                    frm_c8_incad.Show CTE_AMODAL
                    frm_c8_incad.Caption = "Información Cadenas"
                    s_mostrar_estado_semaforo frm_c8_incad, CTE_FUNCIONANDO
                Case CTE_PEZ '9
                    frm_a9_inpez.Show CTE_AMODAL
                    frm_a9_inpez.Caption = "Información Peces"
                    s_mostrar_estado_semaforo frm_a9_inpez, CTE_FUNCIONANDO
                Case CTE_UVA '10
                    frm_aA_inuva.Show CTE_AMODAL
                    frm_aA_inuva.Caption = "Información Universo"
                    s_mostrar_estado_semaforo frm_aA_inuva, CTE_FUNCIONANDO
                Case CTE_YXY '11
                Case Else
                    s_error_num_prog num_prg_activo_ejv
            End Select
        Case CTE_VER_AGENTES_TODOS '13
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                    s_mostrar_agentes_vivos_va0
                Case CTE_PAL '2
                    s_ver_frases_pal
                Case CTE_3R '3
                    s_ver_todos_los_agentes_3r
                Case CTE_PRI '4
                    s_mostrar_agentes_vivos_va0
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
        Case CTE_VER_AGENTES_MEJORES '14
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                Case CTE_PAL '2
                Case CTE_3R '3
                    s_ver_mejores_3r
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
        Case CTE_VER_DICCIONARIO '15
            frm_z0_lista.Show CTE_AMODAL
            s_mostrar_diccionario_pal
        Case CTE_VER_JUGAR_CONTRA_ORDENADOR '16
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                Case CTE_PAL '2
                Case CTE_3R '3
                    s_ver_jugar_contra_ordenador_3r
                Case CTE_PRI '4
                    s_ver_jugar_contra_ordenador_pri
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
        Case CTE_VER_MODIFICAR_AGENTE '17
            Select Case num_prg_activo_ejv
                Case CTE_HYP '1
                Case CTE_PAL '2
                Case CTE_3R '3
                    frm_c0_ce.Fr_ModificarAgente.Visible = True
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
        Case CTE_VER_GRAFICO '18
            frm_z0_graf.Show CTE_AMODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "No existe esa operación de ver"
    End Select


End Sub

Sub s_mostrar_estado_semaforo(mi_form As Object, estado As Integer)

    Select Case estado
        Case CTE_FUNCIONANDO
            mi_form.Label_estado.Caption = "Funcionando"
            mi_form.semaforo.Picture = LoadPicture(f_nombre_completo(path_largo_ejv(CTE_C_PRG_BMP), "semafV.bmp"))
            mi_form.Label_warning.Visible = False
            mi_form.Label_estado.ForeColor = cct_ejv(CTE_VERDEOSCURO)
        Case CTE_DETENIDO
            mi_form.Label_estado.Caption = "Detenido"
            mi_form.semaforo.Picture = LoadPicture(f_nombre_completo(path_largo_ejv(CTE_C_PRG_BMP), "semafR.bmp"))
            mi_form.Label_warning.Visible = False
            mi_form.Label_estado.ForeColor = cct_ejv(CTE_ROJO)
        Case CTE_DETENIENDO
            mi_form.Label_estado.Caption = "Deteniendo..."
            mi_form.semaforo.Picture = LoadPicture(f_nombre_completo(path_largo_ejv(CTE_C_PRG_BMP), "semafA.bmp"))
            mi_form.Label_warning.Visible = True
            mi_form.Label_estado.ForeColor = cct_ejv(CTE_AMARILLO)
        Case CTE_MOSTRANDO
            mi_form.Label_estado.Caption = "Mostrando..."
            mi_form.semaforo.Picture = LoadPicture(f_nombre_completo(path_largo_ejv(CTE_C_PRG_BMP), "semafA.bmp"))
            mi_form.Label_warning.Visible = True
            mi_form.Label_estado.ForeColor = cct_ejv(CTE_AMARILLO)
        Case CTE_JUGANDO
            mi_form.Label_estado.Caption = "Jugando"
            mi_form.semaforo.Picture = LoadPicture(f_nombre_completo(path_largo_ejv(CTE_C_PRG_BMP), "semafV.bmp"))
            mi_form.Label_warning.Visible = False
            mi_form.Label_estado.ForeColor = cct_ejv(CTE_ROSA)
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select

End Sub

Sub ajuste_color_controles_formulario_ejv(Panel As Form)

    Dim control_actual As Control

    Panel.BackColor = cct_ejv(cfondo_ejv)
    For Each control_actual In Panel
        If Left(control_actual.Name, Len("Line")) <> "Line" Then
        If Left(control_actual.Name, Len("Timer")) <> "Timer" Then
        If Left(control_actual.Name, Len("Op_")) <> "Op_" Then
        If Left(control_actual.Name, Len("Cb_")) <> "Cb_" Then
        'If Left(control_actual.Name, Len("Txt_")) <> "Txt_" Then
            control_actual.BackColor = cct_ejv(cfondo_ejv)
        'End If
        End If
        End If
        End If
        End If
    Next control_actual

End Sub
Sub s_inicializar_combo_zoom_ejv(formulario As Object)

    formulario.Cb_Zoom.Clear
    formulario.Cb_Zoom.AddItem "100%"
    formulario.Cb_Zoom.AddItem "50%"
    formulario.Cb_Zoom.AddItem "10%"
    formulario.Cb_Zoom.AddItem "3D"
    formulario.Cb_Zoom.ListIndex = 0

End Sub

Sub s_mdi_load_ejv()

    frm_z0_mdi.WindowState = CTE_MAXIMIZED
    
    s_inicializar_estado_menus_ejv
    
    '1 hyp 2 pri 5 cel 7 exp 9 pez
    esta_detenido_ejv = True
    esta_terminado_ejv = True
    '2 pal
    esta_detenido_ejv = True
    esta_terminado_ejv = True
    '3 3r 8 cad
    esta_detenido_ejv = True
    esta_terminado_ejv = True
    '6 gai
    esta_detenido_ejv = True
    esta_terminado_ejv = True
    
    s_cambiar_estado_enabled_operaciones_ficheros_ejv False
    s_tratamiento_idioma_mdi
    s_fijar_caption_mdi
    
    If Not automatico_ejv Then
        If control_errores_de_programacion_ejv Then
            MsgBox "El modo de control de errores de programación está activado. Esto puede afectar a las simulaciones, haciéndolas mucho más lentas. Para desctivarlo, editar el fichero de inicio en el menú Opciones - Opciones I: Opciones Generales.", vbInformation
        End If
    End If
    
    s_inicializar_combo_zoom_ejv frm_z0_mdi
    
    frm_z0_mdi.mn_listaviejos(0).Caption = nombre_programa_ejv(CTE_HYP) & " - Ej 1"
    
    Load frm_z0_mdi.mn_listaviejos(1)
    num_menu_viejos_ejv = num_menu_viejos_ejv + 1
    frm_z0_mdi.mn_listaviejos(1).Caption = nombre_programa_ejv(CTE_PRI) & " - Ej 1"
    frm_z0_mdi.mn_listaviejos(1).Visible = True
    

End Sub


Sub s_fijar_caption_mdi()

    If automatico_ejv Then
        frm_z0_mdi.Caption = nombre_aplicacion_ejv & " - Modo Automático (" & num_prg_activo_ejv & "," & num_ej_activo_ejv & "," & indice_auto & "," & indice_iteraciones & ")"
        frm_z0_mdi.mn_terminar_todo.Visible = True
    Else
        frm_z0_mdi.Caption = nombre_aplicacion_ejv
        frm_z0_mdi.mn_terminar_todo.Visible = False
    End If


End Sub

