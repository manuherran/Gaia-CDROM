Attribute VB_Name = "bas_a1_hyp"
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
Global energia_proporcionada_al_comer_hyp As Double
Global energia_consumida_al_pelearse_hyp As Double
Global numero_veces_regar_hyp As Integer
Global energia_consumida_al_regar_hyp As Double
Global numero_de_posiciones_alejar_regar_hyp As Integer
Global numero_de_posiciones_alejar_pelear_hyp As Integer
Global agua_inicial_planta_hyp As Double
Global hermafroditas_hyp As Boolean

'Datos de la Planta
Global planta_z() As Double
Global planta_y() As Double
Global planta_x() As Double
Global planta_agua() As Integer

Global num_inic_horm_hyp As Integer


'Numero de plantas que se deben crear en el inicio
Global numero_plantas_que_se_deben_crear_inicio_hyp As Integer
'Numero de plantas actual
Global numero_plantas_hyp As Integer

'Matriz de tipos de hormigas
Global tipo_hyp(1 To 20, 1 To 7) As String


Global suma_riegan_hyp As Integer
Global suma_nacen_va0 As Integer
Global suma_mueren_va0 As Integer
Global suma_mueren_vejez_va0 As Integer
Global suma_pelean_hyp As Integer


Sub s_crear_cajas_tipos_hyp()

Dim i As Integer

For i = 2 To 20
    Load frm_a1_tiposhyp.Etiq1(i)
    Load frm_a1_tiposhyp.Etiq2(i)
    Load frm_a1_tiposhyp.Etiq3(i)
    Load frm_a1_tiposhyp.Caja1(i)
    Load frm_a1_tiposhyp.Caja2(i)
    Load frm_a1_tiposhyp.Caja3(i)
    Load frm_a1_tiposhyp.Caja4(i)
Next i

End Sub
Sub s_crear_plantas_iniciales_hyp()
    
    Dim p As Double
    Dim f As Double
    Dim c As Double
    Dim indice As Integer
    Dim Prueba As Integer
    Dim exito_secuencial As Boolean
    
    'Inicializamos
    numero_plantas_hyp = 0
    
   'Creamos las plantas
    p = 1
    For indice = 1 To numero_plantas_que_se_deben_crear_inicio_hyp
        f = fi_azar1(CInt(mapa_filas_va0))
        c = fi_azar1(CInt(mapa_columnas_va0))
        
        If f_esta_vacio_va0(1, f, c) Then
            s_crear_una_planta_hyp p, f, c, agua_inicial_planta_hyp
        Else
            'La celda está ocupada, probamos a ponerla en otro lugar
            Prueba = 0
            While Prueba < 10
                f = fi_azar1(CInt(mapa_filas_va0))
                c = fi_azar1(CInt(mapa_columnas_va0))
                If f_esta_vacio_va0(p, f, c) Then
                    s_crear_una_planta_hyp p, f, c, agua_inicial_planta_hyp
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
                For c = 1 To mapa_columnas_va0
                    If f_esta_vacio_va0(1, f, c) Then
                         s_crear_una_planta_hyp p, f, c, agua_inicial_planta_hyp
                         exito_secuencial = True
                         Exit For
                     End If
                Next c
                If exito_secuencial Then Exit For
                Next f
                If Not exito_secuencial Then
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: No es posible crear tantas plantas"
                    Exit For
                End If
            End If
        End If
    Next indice

If numero_plantas_hyp <> numero_plantas_que_se_deben_crear_inicio_hyp Then
    s_error_ejv CON_OPCION_FINALIZAR, "Error en la creación de plantas"
End If


End Sub
Sub s_mostrar_info_hyp()

    Dim i As Integer
    Dim sumae As Double
    Dim cont_hembras As Long
    Dim cont_machos As Long
    
    'Los 5 tipos
    frm_a1_inhyp.rojas.Caption = num_agentes_tipo_va0(1)
    frm_a1_inhyp.rosas.Caption = num_agentes_tipo_va0(2)
    frm_a1_inhyp.naranjas.Caption = num_agentes_tipo_va0(3)
    frm_a1_inhyp.amarillas.Caption = num_agentes_tipo_va0(4)
    frm_a1_inhyp.verdes.Caption = num_agentes_tipo_va0(5)
    
    'Totales de hormigas y las plantas
    frm_a1_inhyp.txt_hormigas.Caption = numero_total_de_agentes_ejv
    frm_a1_inhyp.txt_plantas.Caption = numero_plantas_hyp
    
    'suma de toda la energía de todas las hormigas y sexos
    sumae = 0
    cont_hembras = 0
    cont_machos = 0
    For i = 1 To numero_total_de_agentes_ejv
        sumae = sumae + peso_agente_va0(i)
        If sexo_va0(i) = CTE_HEMBRA Then
            cont_hembras = cont_hembras + 1
        Else
            cont_machos = cont_machos + 1
        End If
    Next i
    frm_a1_inhyp.txt_senergía.Caption = Format(sumae, "0.0000")
    frm_a1_inhyp.hembras = cont_hembras
    frm_a1_inhyp.machos = cont_machos

    
    'suma de todas la que riegan en ese ciclo
    frm_a1_inhyp.txt_riegan.Caption = suma_riegan_hyp
    'suma de todas la que nacen en ese ciclo
    frm_a1_inhyp.txt_nacen.Caption = suma_nacen_va0
    'suma de todas la que mueren en ese ciclo
    frm_a1_inhyp.txt_mueren.Caption = suma_mueren_va0
    'suma de todas la que mueren por vejez en ese ciclo
    frm_a1_inhyp.txt_MuerenVejez.Caption = suma_mueren_vejez_va0
    'suma de todas la que pelean en ese ciclo
    frm_a1_inhyp.txt_pelean.Caption = suma_pelean_hyp
    

End Sub
Sub s_grabar_resumen_hyp()

    Dim i As Integer
    Dim linea As String
    
    'Los 5 tipos y las plantas
    linea = ""
    linea = linea & f_comillas(CStr(ciclo_ejv)) ' el ciclo actual
    For i = 1 To num_tipos_agentes_va0
        linea = linea & ";" & f_comillas(CStr(num_agentes_tipo_va0(i))) ' los 5 bichos
    Next i
    linea = linea & ";" & f_comillas(CStr(numero_plantas_hyp)) ' el numero de plantas
    s_grabar_dato_fichero_salida_ejv CTE_FIC_23W_1EJGRA, linea
    
    

End Sub
Sub s_crear_mas_plantas_hyp()

    Dim p As Double
    Dim f As Double
    Dim c As Double
    Dim indice As Integer
    Dim Prueba As Integer
    Dim exito_secuencial As Boolean
    Dim numero_de_pruebas As Long
    Dim numero_de_seres As Long
    Dim numero_de_celdas As Long

    'Creamos las plantas
    For indice = 1 To numero_de_plantas_nacen_ciclo_va0
    p = 1
    f = fi_azar1(CInt(mapa_filas_va0))
    c = fi_azar1(CInt(mapa_columnas_va0))
    
    If f_esta_vacio_va0(p, f, c) Then
        s_crear_una_planta_hyp p, f, c, agua_inicial_planta_hyp
    Else
        'La celda está ocupada, probamos a ponerla en otro lugar
        numero_de_seres = numero_plantas_hyp + numero_total_de_agentes_ejv
        numero_de_celdas = CLng(mapa_columnas_va0) * CLng(mapa_filas_va0)
        numero_de_pruebas = numero_de_seres / numero_de_celdas  '0..1
        numero_de_pruebas = 1 - numero_de_pruebas               '0..1
        numero_de_pruebas = Int(numero_de_pruebas * 10) + 1     '1..11
        Prueba = 0
        While Prueba < numero_de_pruebas
            f = fi_azar1(CInt(mapa_filas_va0))
            c = fi_azar1(CInt(mapa_columnas_va0))
            If f_esta_vacio_va0(p, f, c) Then
                s_crear_una_planta_hyp p, f, c, agua_inicial_planta_hyp
                Prueba = numero_de_pruebas + 1
            Else
                Prueba = Prueba + 1
            End If
        Wend
            If Prueba = numero_de_pruebas Then
                'No ha encontrado ninguna vacia en numero_de_pruebas pruebas
                'La buscamos secuencialmente
                exito_secuencial = False
                If numero_de_celdas > numero_de_seres Then
                    For f = 1 To mapa_filas_va0
                    For c = 1 To mapa_columnas_va0
                        If f_esta_vacio_va0(p, f, c) Then
                            s_crear_una_planta_hyp p, f, c, agua_inicial_planta_hyp
                            exito_secuencial = True
                            Exit For
                         End If
                    Next c
                    If exito_secuencial Then Exit For
                    Next f
                End If
                If Not exito_secuencial Then
                    'La última planta no se ha podido crear
                    'porque esta completamente todo lleno de plantas
                    'y no hay huecos, asi que dejamos que se pulse detener
                    DoEvents
                End If
            End If
    End If
    
    Next indice


End Sub
Function f_hay_planta_hyp(Z As Double, Y As Double, X As Double) As Boolean

    Dim i As Integer
    Dim dev As Boolean
    Dim existe As Boolean
    
    Dim nueva_z As Integer
    Dim nueva_y As Integer
    Dim nueva_x As Integer
        
    If numero_plantas_hyp = 0 Then
        f_hay_planta_hyp = False
        Exit Function
    End If
        
    'Tal vez se llame a esta funcion desde el juego del prisionero o otro
    If num_prg_activo_ejv <> CTE_HYP Then
        f_hay_planta_hyp = False
        'Control errores de programacion
        If control_errores_de_programacion_ejv Then
            s_error_ejv CON_OPCION_FINALIZAR, "Aviso: se llama a f_hay_planta_hyp desde un mundo que no es HYP"
        End If
        Exit Function
    End If
        
    nueva_z = Z
    nueva_y = Y
    nueva_x = X
    
    'Si nos toca inspecionar una celda en la barrera
    'esto significa inspecionar en realidad la que
    'está al otro lado
    If Y = 0 Then nueva_y = mapa_filas_va0
    If Y = mapa_filas_va0 + 1 Then nueva_y = 1
    If X = 0 Then nueva_x = mapa_columnas_va0
    If X = mapa_columnas_va0 + 1 Then nueva_x = 1
    
    If mapa_va0(nueva_z, nueva_y, nueva_x) = CTE_MAPA_PLANTA Then
        dev = True
    Else
        dev = False
    End If
    
    'Control errores de programacion
    If control_errores_de_programacion_ejv Then
        existe = False
        For i = 1 To numero_plantas_hyp
            If planta_z(i) = nueva_z And planta_y(i) = nueva_y And planta_x(i) = nueva_x Then
                existe = True
                Exit For
            End If
            DoEvents
        Next i
        If existe <> dev Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Planta existe en una zona vacía"
        End If
    End If
    
    f_hay_planta_hyp = dev


End Function
Sub s_visualizar_tipos_hormigas_hyp()

    Dim i As Integer
    
    frm_a1_tiposhyp.Etiq1(1).Caption = tipo_hyp(1, 1)
    frm_a1_tiposhyp.Etiq2(1).Caption = tipo_hyp(1, 2)
    frm_a1_tiposhyp.Etiq3(1).Caption = tipo_hyp(1, 3)
    
    frm_a1_tiposhyp.Caja1(1).Text = tipo_hyp(1, 4)
    frm_a1_tiposhyp.Caja2(1).Text = tipo_hyp(1, 5)
    frm_a1_tiposhyp.Caja3(1).Text = tipo_hyp(1, 6)
    frm_a1_tiposhyp.Caja4(1).Text = tipo_hyp(1, 7)
    
    frm_a1_tiposhyp.Etiq1(1).BackColor = 112200
    
    For i = 2 To 20
        frm_a1_tiposhyp.Etiq1(i).Top = frm_a1_tiposhyp.Etiq1(i - 1).Top + 300
        frm_a1_tiposhyp.Etiq1(i).Visible = True
        frm_a1_tiposhyp.Etiq1(i).Caption = tipo_hyp(i, 1)
            
        frm_a1_tiposhyp.Etiq2(i).Top = frm_a1_tiposhyp.Etiq2(i - 1).Top + 300
        frm_a1_tiposhyp.Etiq2(i).Visible = True
        frm_a1_tiposhyp.Etiq2(i).Caption = tipo_hyp(i, 2)
            
        frm_a1_tiposhyp.Etiq3(i).Top = frm_a1_tiposhyp.Etiq3(i - 1).Top + 300
        frm_a1_tiposhyp.Etiq3(i).Visible = True
        frm_a1_tiposhyp.Etiq3(i).Caption = tipo_hyp(i, 3)
            
        frm_a1_tiposhyp.Caja1(i).Top = frm_a1_tiposhyp.Caja1(i - 1).Top + 300
        frm_a1_tiposhyp.Caja1(i).Visible = True
        frm_a1_tiposhyp.Caja1(i).Text = tipo_hyp(i, 4)
       
        frm_a1_tiposhyp.Caja2(i).Top = frm_a1_tiposhyp.Caja2(i - 1).Top + 300
        frm_a1_tiposhyp.Caja2(i).Visible = True
        frm_a1_tiposhyp.Caja2(i).Text = tipo_hyp(i, 5)
        
        frm_a1_tiposhyp.Caja3(i).Top = frm_a1_tiposhyp.Caja3(i - 1).Top + 300
        frm_a1_tiposhyp.Caja3(i).Visible = True
        frm_a1_tiposhyp.Caja3(i).Text = tipo_hyp(i, 6)
       
        frm_a1_tiposhyp.Caja4(i).Top = frm_a1_tiposhyp.Caja4(i - 1).Top + 300
        frm_a1_tiposhyp.Caja4(i).Visible = True
        frm_a1_tiposhyp.Caja4(i).Text = tipo_hyp(i, 7)
       
    Next i
    
    For i = 1 To 4
        frm_a1_tiposhyp.Etiq1(i).BackColor = &HFF& 'rojo
        frm_a1_tiposhyp.Etiq2(i).BackColor = &HFF&
        frm_a1_tiposhyp.Etiq3(i).BackColor = &HFF&
    Next i
    For i = 5 To 8
        frm_a1_tiposhyp.Etiq1(i).BackColor = &H8080FF 'rosa
        frm_a1_tiposhyp.Etiq2(i).BackColor = &H8080FF
        frm_a1_tiposhyp.Etiq3(i).BackColor = &H8080FF
    Next i
    For i = 9 To 12
        frm_a1_tiposhyp.Etiq1(i).BackColor = &H80FF& 'naranja
        frm_a1_tiposhyp.Etiq2(i).BackColor = &H80FF&
        frm_a1_tiposhyp.Etiq3(i).BackColor = &H80FF&
    Next i
    For i = 13 To 16
        frm_a1_tiposhyp.Etiq1(i).BackColor = &HFFFF& 'amarillo
        frm_a1_tiposhyp.Etiq2(i).BackColor = &HFFFF&
        frm_a1_tiposhyp.Etiq3(i).BackColor = &HFFFF&
    Next i
    For i = 17 To 20
        frm_a1_tiposhyp.Etiq1(i).BackColor = &HFF00& 'verde
        frm_a1_tiposhyp.Etiq2(i).BackColor = &HFF00&
        frm_a1_tiposhyp.Etiq3(i).BackColor = &HFF00&
    Next i
    

End Sub
Sub s_inicializar_tipos_hormigas_hyp()

    Dim i As Integer
    
    
    tipo_hyp(1, 2) = "No"
    tipo_hyp(1, 3) = "No"
    tipo_hyp(1, 4) = "100"
    tipo_hyp(1, 5) = "0"
    tipo_hyp(1, 6) = "0"
    tipo_hyp(1, 7) = "0"
    
    
    'Tipos de hormiga
    For i = 1 To 4
        tipo_hyp(i, 1) = "1"
    Next i
    For i = 5 To 8
        tipo_hyp(i, 1) = "2"
    Next i
    For i = 9 To 12
        tipo_hyp(i, 1) = "3"
    Next i
    For i = 13 To 16
        tipo_hyp(i, 1) = "4"
    Next i
    For i = 17 To 20
        tipo_hyp(i, 1) = "5"
    Next i
    
    'Vecinos
    For i = 1 To 17 Step 4
        tipo_hyp(i, 2) = "No"
        tipo_hyp(i, 3) = "No"
    Next i
    For i = 2 To 18 Step 4
        tipo_hyp(i, 2) = "No"
        tipo_hyp(i, 3) = "Si"
    Next i
    For i = 3 To 19 Step 4
        tipo_hyp(i, 2) = "Si"
        tipo_hyp(i, 3) = "No"
    Next i
    For i = 4 To 20 Step 4
        tipo_hyp(i, 2) = "Si"
        tipo_hyp(i, 3) = "Si"
    Next i
    
    
    
    'Caso No No
    For i = 1 To 17 Step 4
        tipo_hyp(i, 4) = "100"
        tipo_hyp(i, 5) = "0"
        tipo_hyp(i, 6) = "0"
    Next i
    
    
    'Reproducción
    For i = 1 To 19 Step 2
        tipo_hyp(i, 7) = "0"
    Next i
    For i = 2 To 20 Step 2
        tipo_hyp(i, 7) = "20"
    Next i
    
    'Casos restantes:
    tipo_hyp(2, 4) = "0"
    tipo_hyp(2, 5) = "0"
    tipo_hyp(2, 6) = "80"
    
    tipo_hyp(3, 4) = "100"
    tipo_hyp(3, 5) = "0"
    tipo_hyp(3, 6) = "0"
    
    tipo_hyp(4, 4) = "0"
    tipo_hyp(4, 5) = "0"
    tipo_hyp(4, 6) = "80"
    
    tipo_hyp(6, 4) = "20"
    tipo_hyp(6, 5) = "0"
    tipo_hyp(6, 6) = "60"
    
    tipo_hyp(7, 4) = "75"
    tipo_hyp(7, 5) = "25"
    tipo_hyp(7, 6) = "0"
    
    tipo_hyp(8, 4) = "20"
    tipo_hyp(8, 5) = "20"
    tipo_hyp(8, 6) = "40"
    
    tipo_hyp(10, 4) = "40"
    tipo_hyp(10, 5) = "0"
    tipo_hyp(10, 6) = "40"
    
    tipo_hyp(11, 4) = "50"
    tipo_hyp(11, 5) = "50"
    tipo_hyp(11, 6) = "0"
    
    tipo_hyp(12, 4) = "20"
    tipo_hyp(12, 5) = "30"
    tipo_hyp(12, 6) = "30"
    
    tipo_hyp(14, 4) = "60"
    tipo_hyp(14, 5) = "0"
    tipo_hyp(14, 6) = "20"
    
    tipo_hyp(15, 4) = "25"
    tipo_hyp(15, 5) = "75"
    tipo_hyp(15, 6) = "0"
    
    tipo_hyp(16, 4) = "20"
    tipo_hyp(16, 5) = "40"
    tipo_hyp(16, 6) = "20"
    
    tipo_hyp(18, 4) = "80"
    tipo_hyp(18, 5) = "0"
    tipo_hyp(18, 6) = "0"
    
    tipo_hyp(19, 4) = "0"
    tipo_hyp(19, 5) = "100"
    tipo_hyp(19, 6) = "0"
    
    tipo_hyp(20, 4) = "0"
    tipo_hyp(20, 5) = "80"
    tipo_hyp(20, 6) = "0"
    
    
End Sub
Sub s_crear_una_planta_hyp(pla_z As Double, pla_y As Double, pla_x As Double, agua As Double)

    'Y:filas
    'X:col

    numero_plantas_hyp = numero_plantas_hyp + 1
    ReDim Preserve planta_z(1 To numero_plantas_hyp) As Double
    ReDim Preserve planta_y(1 To numero_plantas_hyp) As Double
    ReDim Preserve planta_x(1 To numero_plantas_hyp) As Double
    ReDim Preserve planta_agua(1 To numero_plantas_hyp) As Integer
    planta_z(numero_plantas_hyp) = pla_z
    planta_y(numero_plantas_hyp) = pla_y
    planta_x(numero_plantas_hyp) = pla_x
    planta_agua(numero_plantas_hyp) = agua
    'Control errores de programacion
    If control_errores_de_programacion_ejv Then
       If mapa_va0(pla_z, pla_y, pla_x) <> CTE_MAPA_VACIO Then
           s_error_ejv CON_OPCION_FINALIZAR, "Error: Planta creada en una zona no vacía"
       End If
    End If
    mapa_va0(pla_z, pla_y, pla_x) = CTE_MAPA_PLANTA
    'Pintamos la planta
    If ver_agentes_va0 Then
        s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, pla_z, pla_y, pla_x, CTE_PLANTA, cct_ejv(CTE_VERDEBRILLANTE), cct_ejv(CTE_VERDEBRILLANTE), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
    End If


End Sub

Sub s_inicializar_ejemplo_elegido_hyp()

    Dim i As Integer
    Dim j As Integer
    
    Dim f As Integer
    Dim c As Integer

    'Carga de los tipos de agentes
    num_tipos_agentes_va0 = 5
    ReDim tendencia_rel_inicial_mov_tipo_agente_va0(CTE_8_DIR, num_tipos_agentes_va0) As Long
    ReDim tendencia_abs_inicial_mov_tipo_agente_va0(CTE_8_DIR, num_tipos_agentes_va0) As Long
    ReDim numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1 To num_tipos_agentes_va0) As Integer

    'OPCIONES I
    'GENERALES DE EJEMPLOS DE VIDA(DISTINTAS A LAS DE POR DEFECTO)
    '2 Grabar Resumen
    un_ej_grabar_gra_ejv = True
    un_ej_fichero_gra_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_GRA), "r_hyp_" & num_ej_activo_ejv & ".gra")
    un_ej_grabar_resumen_txt_ejv = False
    un_ej_fichero_resumen_txt_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_TXT), "r_hyp_" & num_ej_activo_ejv & ".txt")
    un_ej_grabar_resumen_xls_ejv = False
    un_ej_fichero_resumen_xls_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS), "r_hyp_" & num_ej_activo_ejv & ".xls")
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
    probb_mutacion_tipo_inicial_va0 = 0.01
    probb_mutacion_mov_inicial_va0 = 0.01
    probb_mutacion_pm_inicial_va0 = 0.01
    PMPMCte_va0 = True
    '4 Lugar de nacimiento
    nacimiento_cerca_va0 = True
    '5 Búsqueda de Cadena binaria
    busqueda_cadena_binaria_va0 = False
    cadena_binaria_buscada_va0 = "000000000100000000010000000001"
    long_cadena_buscada_va0 = Len(cadena_binaria_buscada_va0)
    '6 Limite Muerte
    limite_muerte_va0 = 0


    
    Select Case num_ej_activo_ejv
        Case 1
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej01.map"
        'OPCIONES III
        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 0
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 5
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 1
        '5 energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 15
        '6 energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 0.01
        '7 energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 50
        '8 energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 20
        '9 energia_inicial_agente_va0
        energia_inicial_agente_va0 = 20
        '10 numero_veces_regar_hyp
        numero_veces_regar_hyp = 3
        '11 energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 0.2
        '12 numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 1
        '13 numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        '14 numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 0
        '15 numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 0
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        tendencia_rel_inicial_mov_tipo_agente_va0(1, 1) = 40
        tendencia_rel_inicial_mov_tipo_agente_va0(2, 1) = 20
        tendencia_rel_inicial_mov_tipo_agente_va0(3, 1) = 1
        tendencia_rel_inicial_mov_tipo_agente_va0(4, 1) = 1
        tendencia_rel_inicial_mov_tipo_agente_va0(5, 1) = 1
        tendencia_rel_inicial_mov_tipo_agente_va0(6, 1) = 1
        tendencia_rel_inicial_mov_tipo_agente_va0(7, 1) = 1
        tendencia_rel_inicial_mov_tipo_agente_va0(8, 1) = 20
        For i = 2 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 3
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 3
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True
        
        
        Case 2
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej02.map"
        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 150
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 200
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 40
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 40
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 40
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 40
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 40
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 30
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 0.01
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 2
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 10
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 1
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 1
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 20
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 5
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 5
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 5
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0

        Case 3
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej03.map"
        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 1
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 20
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 10
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 0
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 0
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 0
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 10
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 5
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 1
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 2
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 5
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 0
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 0
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 1
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 0
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 0
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True
    
    
        Case 4
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej04.map"
        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 1
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 15
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 3
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 3
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 3
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 3
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 3
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 5
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 1
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 2
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 5
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 1
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 1
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 1
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 1
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 2
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 2
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True

        
        Case 5
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej05.map"
        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 20
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 603
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 300
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 300
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 10
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 2
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 2
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 5
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 2
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 2
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 8
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 0
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 0
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True

        
        Case 6
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej06.map"
        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 10
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 54
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 50
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 1
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 10
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 2
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 2
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 5
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 0
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 2
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 2
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 0
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 0
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True

        
        Case 7
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej07.map"

        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 10
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 54
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 50
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 1
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 1
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 10
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 2
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 2
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 5
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 0
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 2
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 2
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 5
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 5
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True

    
        Case 8
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej08.map"

        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 5
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 20
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 4
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 15
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 2
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 2
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 10
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 1
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 1
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 2
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 0
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 0
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True

    
    
        Case 9
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej09.map"

        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 5
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 20
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 4
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 15
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 2
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 10
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 10
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 0
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 1
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 5
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 0
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 0
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True

    
    
        Case 10
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej10.map"

        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 20
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 20
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 10
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 0
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 0
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 0
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 10
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 20
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 1
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 10
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 3
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 20
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 1
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 10
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 1
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 3
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 3
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 3
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True

    
        Case 11
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej11.map"

        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 0
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 20
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 4
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 4
        'energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 0
        'energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 0
        'energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 999
        'energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        'energia_inicial_agente_va0
        energia_inicial_agente_va0 = 1
        'numero_veces_regar_hyp
        numero_veces_regar_hyp = 0
        'energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 0
        'numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 0
        'numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        'numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 0
        'numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 0
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 40
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 20
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 20
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True
    
        Case 12
        
        'MAPA
        '1 nombre mapa
        nombre_fichero_mapa_va0 = "ej12.map"
        'GENERALES DE VIDA ARTIFICIAL
        '3 tasas de mutación
        probb_mutacion_tipo_inicial_va0 = 0
        probb_mutacion_mov_inicial_va0 = 0
        probb_mutacion_pm_inicial_va0 = 0
        PMPMCte_va0 = True
        'ESPECIFICAS DE HORMIGAS Y PLANTAS
        '1 numero_inicial_de_plantas
        numero_plantas_que_se_deben_crear_inicio_hyp = 0
        '2 numero_inicial_de_hormigas
        num_inic_horm_hyp = 25
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = 5
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = 5
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = 5
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = 5
        numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = 5
        '5 energia_proporcionada_al_comer_hyp
        energia_proporcionada_al_comer_hyp = 0
        '6 energia_consumida_al_mover_va0
        energia_consumida_al_mover_va0 = 0
        '7 energia_consumida_al_reproducirse_va0
        energia_consumida_al_reproducirse_va0 = 99999
        '8 energia_consumida_al_pelearse_hyp
        energia_consumida_al_pelearse_hyp = 0
        '9 energia_inicial_agente_va0
        energia_inicial_agente_va0 = 20
        '10 numero_veces_regar_hyp
        numero_veces_regar_hyp = 0
        '11 energia_consumida_al_regar_hyp
        energia_consumida_al_regar_hyp = 0
        '12 numero_de_plantas_nacen_ciclo_va0
        numero_de_plantas_nacen_ciclo_va0 = 0
        '13 numero_de_posiciones_alejar_regar_hyp
        numero_de_posiciones_alejar_regar_hyp = 0
        '14 numero_de_posiciones_alejar_pelear_hyp
        numero_de_posiciones_alejar_pelear_hyp = 0
        '15 numero_de_posiciones_alejar_reproducirse_va0
        numero_de_posiciones_alejar_reproducirse_va0 = 0
        '21 Algoritmo de busqueda de un espacio libre cercano
        algoritmo_busqueda_va0 = 3
        '22 Tendencias del movimiento
        For i = 1 To num_tipos_agentes_va0
            tendencia_rel_inicial_mov_tipo_agente_va0(1, i) = 99999
            tendencia_rel_inicial_mov_tipo_agente_va0(2, i) = 3
            tendencia_rel_inicial_mov_tipo_agente_va0(3, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(4, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(5, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(6, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(7, i) = 1
            tendencia_rel_inicial_mov_tipo_agente_va0(8, i) = 3
        Next i
        For i = 2 To num_tipos_agentes_va0
            For j = 1 To CTE_8_DIR
                tendencia_abs_inicial_mov_tipo_agente_va0(j, i) = 0
            Next j
        Next i
        '23 Cantidad inicial de agua que posee cada planta
        agua_inicial_planta_hyp = 0
        hermafroditas_hyp = True
        
    
    Case Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe ese ejemplo"
    End Select
        
        
    
    Load frm_a1_tiposhyp
    s_crear_cajas_tipos_hyp
    s_inicializar_tipos_hormigas_hyp
    s_inicializar_arrays_va0
    s_visualizar_tipos_hormigas_hyp
    frm_a1_tiposhyp.Hide
    
    'Cargo el mapa
    mapa_actual_ma0 = f_nombre_completo_existente(path_largo_ejv(CTE_C_PRG_MAP), nombre_fichero_mapa_va0)
    s_aut_leer_mapa_ma0
    s_copiar_mapa_ma0_sobre_va0_va0
    nombre_fichero_ejv = nombre_fichero_mapa_va0
        
        
End Sub
Function f_calcular_accion_hyp(fila As Integer) As String

    Dim p1 As Integer
    Dim p2 As Integer
    Dim p3 As Integer
    Dim p4 As Integer
    Dim azar As Integer
    Dim tope1 As Integer
    Dim tope2 As Integer
    Dim tope3 As Integer
    Dim tope4 As Integer
    
    'Leo las probabilidades de la tabla
    p1 = tipo_hyp(fila, 4)
    p2 = tipo_hyp(fila, 5)
    p3 = tipo_hyp(fila, 6)
    p4 = tipo_hyp(fila, 7)
    
    azar = fi_azar1(100)
    
    tope1 = p1
    tope2 = tope1 + p2
    tope3 = tope2 + p3
    tope4 = tope3 + p4
    
    If tope4 <> 100 Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error: Llamada a la función de azar incorrecta. La suma de las probabilidades no es 100. Es posible que se hayan asignado probabilidades a las acciones de los agentes cuya suma no sea 100."
    End If
    
    If azar < tope1 Then
        f_calcular_accion_hyp = CTE_ACC_MOVER
    Else
        If azar < tope2 Then
            f_calcular_accion_hyp = CTE_ACC_REGAR
        Else
            If azar < tope3 Then
                f_calcular_accion_hyp = CTE_ACC_PELEAR
            Else
                f_calcular_accion_hyp = CTE_ACC_REPRODUCIRSE
            End If
        End If
    End If
    

End Function
Sub s_regar_planta_hyp(agente As Integer, lugar As Integer)

    Dim ag_y As Double
    Dim ag_x As Double
    
    Dim p_z As Double
    Dim p_y As Double
    Dim p_x As Double
    
    Dim indice As Integer
    Dim cont As Integer
    
    
    p_z = 1
    ag_y = agente_y_va0(agente)
    ag_x = agente_x_va0(agente)
        
    'Localizo la planta
    Select Case lugar
        Case CTE_8_N
            p_x = ag_x
            p_y = ag_y - 1
        Case CTE_8_NE
            p_x = ag_x + 1
            p_y = ag_y - 1
        Case CTE_8_E
            p_x = ag_x + 1
            p_y = ag_y
        Case CTE_8_SE
            p_x = ag_x + 1
            p_y = ag_y + 1
        Case CTE_8_S
            p_x = ag_x
            p_y = ag_y + 1
        Case CTE_8_SO
            p_x = ag_x - 1
            p_y = ag_y + 1
        Case CTE_8_O
            p_x = ag_x - 1
            p_y = ag_y
        Case CTE_8_NO
            p_x = ag_x - 1
            p_y = ag_y - 1
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
    End Select
    
    If p_y = 0 Then p_y = mapa_filas_va0
    If p_y = mapa_filas_va0 + 1 Then p_y = 1
    
    If p_x = 0 Then p_x = mapa_columnas_va0
    If p_x = mapa_columnas_va0 + 1 Then p_x = 1
    
    
    indice = fi_indice_planta_hyp(p_z, p_y, p_x)
    If indice <> 0 Then
        'Añado agua a la planta
        planta_agua(indice) = planta_agua(indice) + 1
        suma_riegan_hyp = suma_riegan_hyp + 1
       'Si es comesible lo pinto
        If planta_agua(indice) = numero_veces_regar_hyp Then
            If ver_agentes_va0 Then s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, p_z, p_y, p_x, CTE_PLANTALLENA, cct_ejv(CTE_VERDEBRILLANTE), cct_ejv(CTE_VERDEBRILLANTE), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
        End If
        
        'Quito energía a la hormiga
        If peso_agente_va0(agente) > 0 Then
            peso_agente_va0(agente) = peso_agente_va0(agente) - energia_consumida_al_regar_hyp
        End If
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error al regar planta"
    End If


End Sub
Function fi_indice_planta_hyp(Z As Double, Y As Double, X As Double) As Integer

    Dim encontrado As Boolean
    Dim cont As Integer
    Dim devolver As Integer

    encontrado = False
    devolver = 0
    For cont = 1 To numero_plantas_hyp
        If planta_z(cont) = Z And planta_y(cont) = Y And planta_x(cont) = X Then
            encontrado = True
            devolver = cont
            Exit For
        End If
    Next cont

    'Control errores de programacion
    If control_errores_de_programacion_ejv Then
        If Not encontrado Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: esa planta no existe"
        End If
        If devolver > 0 And mapa_va0(Z, Y, X) <> CTE_MAPA_PLANTA Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Planta en una zona no vacía"
        End If
    End If

    fi_indice_planta_hyp = devolver

End Function
Function f_intentar_comer_planta_hyp(ind_ag As Integer, lugar As Integer) As Boolean

    'Y:filas
    'X:col

    Dim ag_y As Double
    Dim ag_x As Double
    
    Dim p_z As Double
    Dim p_y As Double
    Dim p_x As Double
    
    Dim pl_borrar As Integer
    Dim cont As Integer
    
    p_z = 1
    ag_y = agente_y_va0(ind_ag)
    ag_x = agente_x_va0(ind_ag)
        
    'Localizo la planta
    Select Case lugar
        Case CTE_8_N
            p_x = ag_x
            p_y = ag_y - 1
        Case CTE_8_NE
            p_x = ag_x + 1
            p_y = ag_y - 1
        Case CTE_8_E
            p_x = ag_x + 1
            p_y = ag_y
        Case CTE_8_SE
            p_x = ag_x + 1
            p_y = ag_y + 1
        Case CTE_8_S
            p_x = ag_x
            p_y = ag_y + 1
        Case CTE_8_SO
            p_x = ag_x - 1
            p_y = ag_y + 1
        Case CTE_8_O
            p_x = ag_x - 1
            p_y = ag_y
        Case CTE_8_NO
            p_x = ag_x - 1
            p_y = ag_y - 1
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
    End Select
    
    If p_y = 0 Then p_y = mapa_filas_va0
    If p_y = mapa_filas_va0 + 1 Then p_y = 1
    
    If p_x = 0 Then p_x = mapa_columnas_va0
    If p_x = mapa_columnas_va0 + 1 Then p_x = 1
    
    pl_borrar = fi_indice_planta_hyp(p_z, p_y, p_x)
    If pl_borrar = 0 Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error al intentar comer planta"
    End If
    
    If planta_agua(pl_borrar) >= numero_veces_regar_hyp Then
        'Control errores de programacion
        If control_errores_de_programacion_ejv Then
            If mapa_va0(p_z, p_y, p_x) <> CTE_MAPA_PLANTA Then
                s_error_ejv CON_OPCION_FINALIZAR, "Error: Planta a comer no existe"
            End If
        End If
        mapa_va0(p_z, p_y, p_x) = CTE_MAPA_VACIO
        'Borro la planta
        If ver_agentes_va0 Then
           s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_va, p_z, p_y, p_x, CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
        End If
        'Si la planta a eliminar no es la ultima de todas
        'sustituyo sus valores con los de la ultima. Si es la ultima
        'puedo sustituirlos y luego borrarla, que da lo mismo
        'La elimino
        planta_z(pl_borrar) = planta_z(numero_plantas_hyp)
        planta_y(pl_borrar) = planta_y(numero_plantas_hyp)
        planta_x(pl_borrar) = planta_x(numero_plantas_hyp)
        planta_agua(pl_borrar) = planta_agua(numero_plantas_hyp)
        'Actualizo la cuenta de las plantas
        numero_plantas_hyp = numero_plantas_hyp - 1
        If numero_plantas_hyp > 0 Then
            ReDim Preserve planta_z(1 To numero_plantas_hyp) As Double
            ReDim Preserve planta_y(1 To numero_plantas_hyp) As Double
            ReDim Preserve planta_x(1 To numero_plantas_hyp) As Double
            ReDim Preserve planta_agua(1 To numero_plantas_hyp) As Integer
        End If
        'Aumento la fuerza
        peso_agente_va0(ind_ag) = peso_agente_va0(ind_ag) + energia_proporcionada_al_comer_hyp
        f_intentar_comer_planta_hyp = True
    Else
        f_intentar_comer_planta_hyp = False
    End If
    

End Function
Function f_intentar_reproducir_hormiga_hyp(hor_iniciativa As Integer, direccion As Integer) As Boolean

    'hor_iniciativa es la que intenta reproducirse, la actual, la que toma la iniciativa
    'hor_provocada es la que hemos encontrado, la vecina

    Dim hor_provocada As Integer
    
    Dim pm_tipo As Double
    Dim pm_mov As Double
    Dim pm_pm As Double
    
    Dim i As Integer
    
    'futuro hijo
    Dim hijo_x As Double
    Dim hijo_y As Double
    
    Dim provocada_z As Double
    Dim provocada_y As Double
    Dim provocada_x As Double

    Dim iniciativa_z As Double
    Dim iniciativa_y As Double
    Dim iniciativa_x As Double

    Dim tipo_provocada As Integer
    Dim tipo_iniciativa As Integer
    Dim tipo_h As Integer

    Dim azar As Integer
    Dim se_ha_encontrado As Boolean
    
    Dim cadena As String

    ReDim m_rel(1 To CTE_8_DIR) As Long
    ReDim m_abs(1 To CTE_8_DIR) As Long

    iniciativa_z = 1
    provocada_z = 1
    iniciativa_y = agente_y_va0(hor_iniciativa)
    iniciativa_x = agente_x_va0(hor_iniciativa)

    Select Case direccion
        Case CTE_8_N
            provocada_x = iniciativa_x
            provocada_y = iniciativa_y - 1
        Case CTE_8_NE
            provocada_x = iniciativa_x + 1
            provocada_y = iniciativa_y - 1
        Case CTE_8_E
            provocada_x = iniciativa_x + 1
            provocada_y = iniciativa_y
        Case CTE_8_SE
            provocada_x = iniciativa_x + 1
            provocada_y = iniciativa_y + 1
        Case CTE_8_S
            provocada_x = iniciativa_x
            provocada_y = iniciativa_y + 1
        Case CTE_8_SO
            provocada_x = iniciativa_x - 1
            provocada_y = iniciativa_y + 1
        Case CTE_8_O
            provocada_x = iniciativa_x - 1
            provocada_y = iniciativa_y
        Case CTE_8_NO
            provocada_x = iniciativa_x - 1
            provocada_y = iniciativa_y - 1
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
    End Select
    If provocada_y = 0 Then provocada_y = mapa_filas_va0
    If provocada_y = mapa_filas_va0 + 1 Then provocada_y = 1
    
    If provocada_x = 0 Then provocada_x = mapa_columnas_va0
    If provocada_x = mapa_columnas_va0 + 1 Then provocada_x = 1
    
    'Calculo el tipo del provocada con el que se reproduce
    hor_provocada = fi_indice_agente_va0(provocada_z, provocada_y, provocada_x)
     
    tipo_iniciativa = agente_tipo_va0(hor_iniciativa)
    tipo_provocada = agente_tipo_va0(hor_provocada)
    
    'Calculo el sexo del otro
    If sexo_va0(hor_provocada) = sexo_va0(hor_iniciativa) And hermafroditas_hyp = False Then
        'No son de sexos compatibles
        f_intentar_reproducir_hormiga_hyp = False
        Exit Function
    End If
    
    
    'Tipo===============================================================
    'Cojo la media del provocada y iniciativa
    tipo_h = Int(tipo_provocada + tipo_iniciativa) / 2
     
    'Hago que el usar la parte entera se desplace cada vez a un lado
    If tipo_provocada <> tipo_iniciativa Then
        azar = fi_azar1(2)
        If azar = 1 Then
           tipo_h = tipo_h + 1
        End If
    End If
    
    'Le añado una variación-mutación
    If agente_probb_mutacion_tipo_va0(hor_iniciativa) > 0 And agente_probb_mutacion_tipo_va0(hor_provocada) > 0 Then
        If f_analizar_probabilidad_ejv((agente_probb_mutacion_tipo_va0(hor_iniciativa) + agente_probb_mutacion_tipo_va0(hor_provocada)) / 2) Then
           tipo_h = fi_azar1(5)
        End If
    End If
    
    'Movimiento =========================================================================
    'Cojo la media del provocada y iniciativa
    For i = 1 To CTE_8_DIR
        m_rel(i) = (agente_tendencia_rel_mov_va0(i, hor_provocada) + agente_tendencia_rel_mov_va0(i, hor_iniciativa)) / 2
        'Le añado una variación-mutación
        If agente_probb_mutacion_mov_va0(hor_iniciativa) > 0 And agente_probb_mutacion_mov_va0(hor_provocada) > 0 Then
            If f_analizar_probabilidad_ejv((agente_probb_mutacion_mov_va0(hor_iniciativa) + agente_probb_mutacion_mov_va0(hor_provocada)) / 2) Then
               m_rel(i) = m_rel(i) + fi_azar2(-5, 5)
               If m_rel(i) < 0 Then m_rel(i) = 0
            End If
        End If
        m_abs(i) = (agente_tendencia_abs_mov_va0(i, hor_provocada) + agente_tendencia_abs_mov_va0(i, hor_iniciativa)) / 2
        'Le añado una variación-mutación
        If agente_probb_mutacion_mov_va0(hor_iniciativa) > 0 And agente_probb_mutacion_mov_va0(hor_provocada) > 0 Then
            If f_analizar_probabilidad_ejv((agente_probb_mutacion_mov_va0(hor_iniciativa) + agente_probb_mutacion_mov_va0(hor_provocada)) / 2) Then
               m_abs(i) = m_abs(i) + fi_azar2(-5, 5)
               If m_abs(i) < 0 Then m_abs(i) = 0
            End If
        End If
    Next i
    
    
    'Probb de Muta del tipo =============================================================
    'Cojo la media del provocada y iniciativa
    pm_tipo = (agente_probb_mutacion_tipo_va0(hor_provocada) + agente_probb_mutacion_tipo_va0(hor_iniciativa)) / 2
    'Le añado una variación-mutación
    If agente_probb_mutacion_pm_va0(hor_iniciativa) > 0 And agente_probb_mutacion_pm_va0(hor_provocada) > 0 Then
        If f_analizar_probabilidad_ejv((agente_probb_mutacion_pm_va0(hor_iniciativa) + agente_probb_mutacion_pm_va0(hor_provocada)) / 2) Then
            'modifico el +-10% de lo que ya habia
            pm_tipo = pm_tipo + fi_azar2(-(pm_tipo / 10), (pm_tipo / 10))
            If pm_tipo < 0 Then pm_tipo = 0
        End If
    End If
    
    'Probb de Muta del mov ==============================================================
    'Cojo la media del provocada y iniciativa
    pm_mov = (agente_probb_mutacion_mov_va0(hor_provocada) + agente_probb_mutacion_mov_va0(hor_iniciativa)) / 2
    'Le añado una variación-mutación
    If agente_probb_mutacion_pm_va0(hor_iniciativa) > 0 And agente_probb_mutacion_pm_va0(hor_provocada) > 0 Then
        If f_analizar_probabilidad_ejv((agente_probb_mutacion_pm_va0(hor_iniciativa) + agente_probb_mutacion_pm_va0(hor_provocada)) / 2) Then
            'modifico el +-10% de lo que ya habia
            pm_mov = pm_mov + fi_azar2(-(pm_mov / 10), (pm_mov / 10))
            If pm_mov < 0 Then pm_mov = 0
        End If
    End If
    
    'Probb de Muta de la Probb de Muta ==================================================
    'Cojo la media del provocada y iniciativa
    pm_pm = (agente_probb_mutacion_mov_va0(hor_provocada) + agente_probb_mutacion_mov_va0(hor_iniciativa)) / 2
    If Not PMPMCte_va0 Then
        'Le añado una variación-mutación
        If agente_probb_mutacion_pm_va0(hor_iniciativa) > 0 And agente_probb_mutacion_pm_va0(hor_provocada) > 0 Then
            If f_analizar_probabilidad_ejv((agente_probb_mutacion_pm_va0(hor_iniciativa) + agente_probb_mutacion_pm_va0(hor_provocada)) / 2) Then
                'modifico el +-10% de lo que ya habia
                pm_pm = pm_pm + fi_azar2(-(pm_pm / 10), (pm_pm / 10))
                If pm_pm < 0 Then pm_pm = 0
            End If
        End If
    End If
    
    'Lugar de nacimiento ================================================================
    If nacimiento_cerca_va0 Then
        'Busco un espacio libre o cerca de la iniciativa o cerca del provocada
        se_ha_encontrado = f_buscar_lugar_nacimiento_cerca_va0(iniciativa_y, iniciativa_x, hijo_y, hijo_x)
        If Not se_ha_encontrado Then
            se_ha_encontrado = f_buscar_lugar_nacimiento_cerca_va0(provocada_y, provocada_x, hijo_y, hijo_x)
        End If
    Else
        se_ha_encontrado = f_buscar_lugar_nacimiento_cualquiera_hyp(hijo_y, hijo_x)
    End If
    
    'Cadena Binaria
    cadena = f_sobrecruzamiento_ejv(cadena_binaria_va0(hor_provocada), cadena_binaria_va0(hor_iniciativa), 1, CTE_1_PTO_CORTE)
    
    If se_ha_encontrado Then
        'Creo la hormiga en ese lugar
        s_crear_un_agente_va0 1, hijo_y, hijo_x, tipo_h, energia_inicial_agente_va0, f_combinar_apellidos_va0(hor_provocada, hor_iniciativa), pm_tipo, pm_mov, pm_pm, m_abs(), m_rel(), cadena
        suma_nacen_va0 = suma_nacen_va0 + 1
        'Resto energía a la iniciativa
        peso_agente_va0(hor_iniciativa) = peso_agente_va0(hor_iniciativa) - energia_consumida_al_reproducirse_va0
        f_intentar_reproducir_hormiga_hyp = True
    Else
        'No se ha encontrado sitio para la hormiga
        f_intentar_reproducir_hormiga_hyp = False
        Exit Function
    End If


End Function
Function f_buscar_lugar_nacimiento_cualquiera_hyp(i As Double, j As Double)
    
'El resultado de la funcion lo devuelvo en los mismos parametros i j
    
    Dim Prueba As Integer
    Dim exito_secuencial As Boolean

    Dim numero_maximo_de_pruebas As Double
    Dim numero_de_seres As Integer
    Dim numero_de_celdas As Integer
    
    Dim se_ha_encontrado As Boolean
   
    se_ha_encontrado = False
    
    'Creo la hormiga en un lugar vacío
    'Si no hay lugares vacios, la hormiga no nace
    i = fi_azar1(CInt(mapa_filas_va0))
    j = fi_azar1(CInt(mapa_columnas_va0))
    If f_esta_vacio_va0(1, i, j) Then
        se_ha_encontrado = True
        'Ya lo he encontrado, me salgo
    Else
        'La celda está ocupada, probamos a ponerla en otro lugar
        numero_de_seres = numero_plantas_hyp + numero_total_de_agentes_ejv
        numero_de_celdas = mapa_columnas_va0 * mapa_filas_va0
        numero_maximo_de_pruebas = numero_de_seres / numero_de_celdas  '0..1
        numero_maximo_de_pruebas = 1 - numero_maximo_de_pruebas               '0..1
        numero_maximo_de_pruebas = Int(numero_maximo_de_pruebas * 10) + 1     '1..11
        Prueba = 0
        While Prueba < numero_maximo_de_pruebas
            i = fi_azar1(CInt(mapa_filas_va0))
            j = fi_azar1(CInt(mapa_columnas_va0))
            If f_esta_vacio_va0(1, i, j) Then
                'Ya lo he encontrado, me salgo
                se_ha_encontrado = True
                Prueba = numero_maximo_de_pruebas + 1 'para salirme, hago esto
            Else
                Prueba = Prueba + 1
            End If
        Wend
        If Prueba = numero_maximo_de_pruebas Then
            'No ha encontrado ninguna vacia en "numero_maximo_de_pruebas" pruebas
            'La buscamos secuencialmente
            exito_secuencial = False
            If numero_de_celdas > numero_de_seres Then
                For i = 1 To mapa_filas_va0
                For j = 1 To mapa_columnas_va0
                    If f_esta_vacio_va0(1, i, j) Then
                        'Ya lo he encontrado, me salgo
                        se_ha_encontrado = True
                        exito_secuencial = True
                        Exit For
                     End If
                Next j
                If exito_secuencial Then Exit For
                Next i
            End If
            If Not exito_secuencial Then
                se_ha_encontrado = False
                'No hay sitio: superpoblación
                'la hormiga no se crea
            End If
        End If
    End If
    
    f_buscar_lugar_nacimiento_cualquiera_hyp = se_ha_encontrado
    
    

End Function
Sub s_accion_reproducirse_va0(hor As Integer, mis_vecinos() As Integer)

    'hormiga actual
    Dim X As Integer
    Dim Y As Integer
    
    Dim cont As Integer
    Dim vecino_elegido As Integer
    Dim vecinos_desordenados(1 To CTE_8_DIR) As Integer

    Dim he_reproducido As Boolean

    'Calculo el vecino con el que se reproduce
    'LA hormigA que se reproduce encuentra siempre un hormigO
    'dispuesto a reproducirse. LA hormigA gastará energía
    'pero el hormigO no, ni se entera
    
    he_reproducido = False
    'Array ordenado
    For cont = 1 To CTE_8_DIR
        vecinos_desordenados(cont) = cont
    Next cont

    'hormiga actual
    Y = agente_y_va0(hor)
    X = agente_x_va0(hor)

    f_desordenar_array_i vecinos_desordenados()
    
    'Ahora el array está desordenado
    For cont = 1 To CTE_8_DIR
        vecino_elegido = vecinos_desordenados(cont)
        If mis_vecinos(vecino_elegido) = CTE_VEC_AGENTE Then
            'Es  "agente"
            'Esta ación solo tendrá exito si hay sitio libre
            he_reproducido = f_intentar_reproducir_hormiga_hyp(hor, vecino_elegido)
            Exit For
        End If
    Next cont
     
     
    If Not he_reproducido Then
        'Deberiamos reproducir esta hormiga pero no hay sitio para que nazca la hija
        'o la pareja no era del sexo adecuado
        'asi que no nace, pero no es un error
    End If
     

End Sub
Sub s_accion_pelear_hyp(hor As Integer, mis_vecinos() As Integer)

    
    Dim cont As Integer
    Dim vecino_elegido As Integer
    Dim vecinos_desordenados(1 To CTE_8_DIR) As Integer

    Dim he_peleado As Boolean

    'Calculo el vecino con el que se pelea
    'La hormiga que se pelea encuentra siempre otra
    'dispuesta a pelearse.
    'La hormiga que pierda, perderá la mitad de su energía
    'que pasará a la ganadora
    
    he_peleado = False
    'Array ordenado
    For cont = 1 To CTE_8_DIR
        vecinos_desordenados(cont) = cont
    Next cont

    f_desordenar_array_i vecinos_desordenados()
    
    'Ahora el array está desordenado
    For cont = 1 To CTE_8_DIR
        vecino_elegido = vecinos_desordenados(cont)
        If mis_vecinos(vecino_elegido) = CTE_VEC_AGENTE Then
            'Es "agente"
            s_pelear_hormiga_hyp hor, vecino_elegido
            he_peleado = True
            Exit For
        End If
    Next cont

    If Not he_peleado Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error: no se ha podido pelear"
    End If



End Sub
Sub s_pelear_hormiga_hyp(hor As Integer, direccion As Integer)

    Dim prob_ganar_hor As Double
    Dim prob_ganar_ene As Double
    
    Dim prob_ganar_esta_pelea_hor As Double
    Dim prob_ganar_esta_pelea_ene As Double
    
    Dim enemigo_z As Double
    Dim enemigo_y As Double
    Dim enemigo_x As Double

    Dim hor_enemigo As Integer

    Dim azar As Integer

    Dim gana_hormiga As Boolean
    Dim energia_robada As Double

    enemigo_z = 1
    Select Case direccion
        Case CTE_8_N
            enemigo_x = agente_x_va0(hor)
            enemigo_y = agente_y_va0(hor) - 1
        Case CTE_8_NE
            enemigo_x = agente_x_va0(hor) + 1
            enemigo_y = agente_y_va0(hor) - 1
        Case CTE_8_E
            enemigo_x = agente_x_va0(hor) + 1
            enemigo_y = agente_y_va0(hor)
        Case CTE_8_SE
            enemigo_x = agente_x_va0(hor) + 1
            enemigo_y = agente_y_va0(hor) + 1
        Case CTE_8_S
            enemigo_x = agente_x_va0(hor)
            enemigo_y = agente_y_va0(hor) + 1
        Case CTE_8_SO
            enemigo_x = agente_x_va0(hor) - 1
            enemigo_y = agente_y_va0(hor) + 1
        Case CTE_8_O
            enemigo_x = agente_x_va0(hor) - 1
            enemigo_y = agente_y_va0(hor)
        Case CTE_8_NO
            enemigo_x = agente_x_va0(hor) - 1
            enemigo_y = agente_y_va0(hor) - 1
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe esa dirección"
    End Select
    If enemigo_y = 0 Then enemigo_y = mapa_filas_va0
    If enemigo_y = mapa_filas_va0 + 1 Then enemigo_y = 1
    
    If enemigo_x = 0 Then enemigo_x = mapa_columnas_va0
    If enemigo_x = mapa_columnas_va0 + 1 Then enemigo_x = 1
    
    'Leo la probabilidad de ganar de cada uno
     prob_ganar_hor = hormiga_probabilidad_ganar_hyp(hor)
     hor_enemigo = fi_indice_agente_va0(enemigo_z, enemigo_y, enemigo_x)
     prob_ganar_ene = hormiga_probabilidad_ganar_hyp(hor_enemigo)
    
    'Calculo las probabilidades prácticas para esta pelea
    'Porque por ejemplo si las dos tienen probabilidad de ganar 80%
    'en realidad en este caso es como tener un 50%
     prob_ganar_esta_pelea_hor = (prob_ganar_hor / (prob_ganar_hor + prob_ganar_ene)) * 100
     prob_ganar_esta_pelea_ene = 100 - prob_ganar_esta_pelea_hor
    
    'Calculo quien gana
     azar = fi_azar1(100)
     If azar <= prob_ganar_esta_pelea_hor Then
        gana_hormiga = True
     End If
    
    suma_pelean_hyp = suma_pelean_hyp + 1
    'Pasamos energía de una a otra
    ' Cambiamos sus probabilidades, aumentando ligeramente las de la ganadora
    'y disminuyendo las de la perdedora. Para aumentar, aumento la mitad de lo
    'que le falta para llegar a 100. Para dismuinuir, disminuyo a la mitad.
    'Así pueden aumentar y disminuir hasta el infinito
     If gana_hormiga Then
         energia_robada = CDbl(peso_agente_va0(hor_enemigo) / 2)
         peso_agente_va0(hor_enemigo) = peso_agente_va0(hor_enemigo) - energia_robada
         peso_agente_va0(hor) = peso_agente_va0(hor) + energia_robada
         hormiga_probabilidad_ganar_hyp(hor) = hormiga_probabilidad_ganar_hyp(hor) + ((100 - hormiga_probabilidad_ganar_hyp(hor)) / 2)
         hormiga_probabilidad_ganar_hyp(hor_enemigo) = hormiga_probabilidad_ganar_hyp(hor_enemigo) / 2
     Else
         energia_robada = CDbl(peso_agente_va0(hor) / 2)
         peso_agente_va0(hor) = peso_agente_va0(hor) - energia_robada
         peso_agente_va0(hor_enemigo) = peso_agente_va0(hor_enemigo) + energia_robada
         hormiga_probabilidad_ganar_hyp(hor_enemigo) = hormiga_probabilidad_ganar_hyp(hor_enemigo) + ((100 - hormiga_probabilidad_ganar_hyp(hor_enemigo)) / 2)
         hormiga_probabilidad_ganar_hyp(hor) = hormiga_probabilidad_ganar_hyp(hor) / 2
     End If

    'Las dos pierden la energia determinada al pelearse
    peso_agente_va0(hor) = peso_agente_va0(hor) - energia_consumida_al_pelearse_hyp
    peso_agente_va0(hor_enemigo) = peso_agente_va0(hor_enemigo) - energia_consumida_al_pelearse_hyp


End Sub
Sub s_accion_regar_hyp(hor As Integer, mis_vecinos() As Integer)
    
    
    Dim cont As Integer
    Dim encontrado As Boolean

    encontrado = False
    For cont = 1 To CTE_8_DIR
        If mis_vecinos(cont) = CTE_VEC_PLANTA Then
            s_regar_planta_hyp hor, cont
            encontrado = True
            Exit For
        End If
    Next cont
    
    If Not encontrado Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error al regar planta. No hay planta."
    End If

End Sub
Sub s_grabar_tipos_hyp()

    Dim i As Integer

    frm_a1_tiposhyp.Refresh
    For i = 1 To 20
        tipo_hyp(i, 4) = frm_a1_tiposhyp.Caja1(i).Text
        tipo_hyp(i, 5) = frm_a1_tiposhyp.Caja2(i).Text
        tipo_hyp(i, 6) = frm_a1_tiposhyp.Caja3(i).Text
        tipo_hyp(i, 7) = frm_a1_tiposhyp.Caja4(i).Text
    Next i
    
    ReDim numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1 To num_tipos_agentes_va0) As Integer
    numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1) = frm_a1_tiposhyp.nhi1.Text
    numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2) = frm_a1_tiposhyp.nhi2.Text
    numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3) = frm_a1_tiposhyp.nhi3.Text
    numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4) = frm_a1_tiposhyp.nhi4.Text
    numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5) = frm_a1_tiposhyp.nhi5.Text

End Sub

Sub s_inicializar_hyp()
    
    hay_que_detener_ejv = False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
    
    frm_a0_va.Refresh
    
    se_ha_empezado_a_crear_agentes_va0 = False
    s_inicializar_arrays_va0
    
    s_copiar_mapa_ma0_sobre_va0_va0
    nombre_fichero_ejv = nombre_fichero_mapa_va0
    
    s_mapa_pintar_bordes_va0 frm_a0_va
    s_mostrar_mapa_actual_va0 False
    s_crear_plantas_iniciales_hyp
    se_ha_empezado_a_crear_agentes_va0 = True
    s_crear_agentes_iniciales_va0

End Sub

Sub s_cargar_opciones_hyp()


    frm_a1_ophyp.numInicHorm.Caption = num_inic_horm_hyp

    frm_a1_ophyp.numIniciPlantas.Text = numero_plantas_que_se_deben_crear_inicio_hyp
    frm_a1_ophyp.EnergiaAlComer.Text = energia_proporcionada_al_comer_hyp
    frm_a1_ophyp.EnergiaAlMover.Text = energia_consumida_al_mover_va0
    frm_a1_ophyp.EnergiaAlReproducir.Text = energia_consumida_al_reproducirse_va0
    frm_a1_ophyp.EnergiaAlPelear.Text = energia_consumida_al_pelearse_hyp
    frm_a1_ophyp.EnergiaInicialAgente.Text = energia_inicial_agente_va0
    frm_a1_ophyp.EnergiaAlRegar.Text = energia_consumida_al_regar_hyp
    frm_a1_ophyp.PlantasPorCiclo.Text = numero_de_plantas_nacen_ciclo_va0
    frm_a1_ophyp.PosicionesRegar.Text = numero_de_posiciones_alejar_regar_hyp
    frm_a1_ophyp.PosicionesPelear.Text = numero_de_posiciones_alejar_pelear_hyp
    frm_a1_ophyp.PosicionesReproducirse.Text = numero_de_posiciones_alejar_reproducirse_va0
    frm_a1_ophyp.CantInicialAgua.Text = agua_inicial_planta_hyp
    frm_a1_ophyp.VecesRegar.Text = numero_veces_regar_hyp
    frm_a1_ophyp.Op_Hermafroditas = hermafroditas_hyp
    frm_a1_ophyp.Op_nHermafroditas = Not hermafroditas_hyp


End Sub

Sub s_grabar_opciones_hyp()

    'Check:0,1
    'Option:true,false

    
    numero_plantas_que_se_deben_crear_inicio_hyp = CInt(frm_a1_ophyp.numIniciPlantas.Text)
    energia_proporcionada_al_comer_hyp = CDbl(frm_a1_ophyp.EnergiaAlComer.Text)
    energia_consumida_al_mover_va0 = CDbl(frm_a1_ophyp.EnergiaAlMover.Text)
    energia_consumida_al_reproducirse_va0 = CDbl(frm_a1_ophyp.EnergiaAlReproducir.Text)
    energia_consumida_al_pelearse_hyp = CDbl(frm_a1_ophyp.EnergiaAlPelear.Text)
    energia_inicial_agente_va0 = CDbl(frm_a1_ophyp.EnergiaInicialAgente.Text)
    energia_consumida_al_regar_hyp = CDbl(frm_a1_ophyp.EnergiaAlRegar.Text)
    numero_de_plantas_nacen_ciclo_va0 = CInt(frm_a1_ophyp.PlantasPorCiclo.Text)
    numero_de_posiciones_alejar_regar_hyp = CInt(frm_a1_ophyp.PosicionesRegar.Text)
    numero_de_posiciones_alejar_pelear_hyp = CInt(frm_a1_ophyp.PosicionesPelear.Text)
    numero_de_posiciones_alejar_reproducirse_va0 = CInt(frm_a1_ophyp.PosicionesReproducirse.Text)
    agua_inicial_planta_hyp = CInt(frm_a1_ophyp.CantInicialAgua.Text)
    numero_veces_regar_hyp = CInt(frm_a1_ophyp.VecesRegar.Text)
    hermafroditas_hyp = frm_a1_ophyp.Op_Hermafroditas
    


End Sub

