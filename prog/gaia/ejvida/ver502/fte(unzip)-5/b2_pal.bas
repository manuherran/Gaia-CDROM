Attribute VB_Name = "bas_b2_pal"
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


Global dicc_carpeta_pal As String
Global dicc_fichero_pal As String

Global numero_palabras_dicc_pal As Long
Global frase_a_buscar_pal As String
Global numero_de_frases_inicial As Integer
Global numero_de_progenitores_pal As Integer
Global numero_de_hijos_pal As Integer
Global tasa_de_mutacion_pal As Integer
Global tipo_reproduccion_pal As Integer

'Cadena buscada
Global numero_de_palabras_frase_buscada As Integer
Global palabra_cadena_buscada() As String

'Mutaciones
Global ha_habido_mutacion_anterior_pal As Boolean
Global mutaciones_acumuladas_pal As Boolean
Global tasa_de_mutacion_vieja_pal As Integer

'Matriz de elementos del patrón: cadenas generadas
Global elemento_palabra_frase() As String
Global peso_agente_pal() As Double

'otras opciones
Global eleccion_de_padres_al_azar_pal As Boolean
Global padres_identicos_producen_mutaciones_pal As Boolean
Global pesos_relativos_pal As Boolean

'Diccionario
Global palabra_del_diccionario() As String
Global c_d_pal As Integer



Global paso_pal As Integer

Sub s_mostrar_info_pal()

End Sub

Sub s_grabar_resumen_pal()

    Dim linea As String

    'Grabamos datos
    linea = ""
    linea = linea & f_comillas(CStr(ciclo_ejv)) ' el ciclo actual
    linea = linea & ";" & f_comillas(CStr(peso_agente_pal(1)))
    linea = linea & ";" & f_comillas(CStr(peso_agente_pal(numero_total_de_agentes_ejv)))
    s_grabar_dato_fichero_salida_ejv CTE_FIC_23W_1EJGRA, linea

End Sub
Sub s_pintar_todas_las_frases_pal()

    Dim FRASE As String
    Dim cont_frase As Integer
    Dim cont_palabra As Integer

    frm_b2_pal.fr_Todas.Visible = True
    
    frm_b2_pal.List5.Clear
    frm_b2_pal.List6.Clear
    
    'mostramos el contenido de todas
    For cont_frase = 1 To numero_total_de_agentes_ejv
        FRASE = ""
        'de numero_de_palabras_por_frase elementos
        For cont_palabra = 1 To numero_de_palabras_frase_buscada
            FRASE = FRASE & elemento_palabra_frase(cont_palabra, cont_frase) & " "
        Next cont_palabra
        frm_b2_pal.List5.AddItem FRASE
        frm_b2_pal.List6.AddItem Format(peso_agente_pal(cont_frase), "0.00000000")
DoEvents
    Next cont_frase

End Sub
Sub s_pintar_frases_pal()

    Dim FRASE As String
    Dim cont_frase As Integer
    Dim cont_palabra As Integer
    
    
    frm_b2_inpal.txt_ciclo = ciclo_ejv
    frm_b2_inpal.maximo = Left(CStr(peso_agente_pal(1)), 10)
    frm_b2_inpal.minimo = Left(CStr(peso_agente_pal(numero_total_de_agentes_ejv)), 10)
    
    
    If frm_b2_pal.Op_VerFrases Then
        frm_b2_pal.fr_Todas.Visible = False
        frm_b2_pal.Fr_Opciones.Visible = False
        frm_b2_pal.Fr_Ejecucion.Visible = True
    
        frm_b2_pal.List1.Visible = True
        frm_b2_pal.List2.Visible = True
        frm_b2_pal.List3.Visible = True
        frm_b2_pal.List4.Visible = True
        
        frm_b2_pal.List1.Clear
        frm_b2_pal.List2.Clear
        frm_b2_pal.List3.Clear
        frm_b2_pal.List4.Clear
        'mostramos el contenido de las 10 primeras
        For cont_frase = 1 To 10
            If cont_frase > numero_total_de_agentes_ejv Then Exit For
            FRASE = ""
            'de numero_de_palabras_por_frase elementos
            For cont_palabra = 1 To numero_de_palabras_frase_buscada
                FRASE = FRASE & elemento_palabra_frase(cont_palabra, cont_frase) & " "
DoEvents
            Next cont_palabra
            frm_b2_pal.List1.AddItem FRASE
            frm_b2_pal.List2.AddItem Format(peso_agente_pal(cont_frase), "0.00000000")
        Next cont_frase
        'mostramos el contenido de las 10 últimas
        For cont_frase = numero_total_de_agentes_ejv - 10 + 1 To numero_total_de_agentes_ejv
            If cont_frase > numero_total_de_agentes_ejv Or cont_frase < 0 Then Exit For
            FRASE = ""
            'de numero_de_palabras_por_frase elementos
            For cont_palabra = 1 To numero_de_palabras_frase_buscada
                FRASE = FRASE & elemento_palabra_frase(cont_palabra, cont_frase) & " "
DoEvents
            Next cont_palabra
            frm_b2_pal.List3.AddItem FRASE
            frm_b2_pal.List4.AddItem Format(peso_agente_pal(cont_frase), "0.00000000")
        Next cont_frase
    End If

    If peso_agente_pal(1) = 1 Then
        'hay_que_detener_ejv = True
        finalizacion_usuario_ejv = False
        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
    End If


End Sub
Function f_funcion_ajuste_pesos_frases_pal(numero_de_frase_actual As Integer) As Double


    Dim cont_frase As Integer
    Dim cont_palabra As Integer
    Dim peso As Double
    
    
    Dim numero_de_palabras_distintas_frase_mas_parecida As Integer
    Dim numero_de_palabras_iguales_frase_mas_parecida As Integer
    Dim frase_mas_parecida As Integer
    Dim numero_de_palabras_iguales_curso As Integer

    'El peso se calcula en función de si existen más entidades de mas peso
    'que la actual y parecidas a la actual
    'Busco la que mas se parece de todas de las de mas peso que ella
    'es decir, de todas sus anteriores, ya que está ordenado
    'Si esa entidad..
    'es identica, el peso es 0
    'es identica, salvo por un elemento, el peso es = al peso calculado menos la mitad del peso calculado, osea:
    ' 1 distinta  -->  p = p - (p/2)
    ' 2 distintas -->  p = p - (p/4)
    ' 3 distintas  -->  p = p - (p/8)
    ' 4 distintas  -->  p = p - (p/16)
    '...
    'igual, menos n -->  p = p - (p/(2^n))
    'Si es completamente distinta, el peso es p
        
    'La primera no la trato pq estan ordenadas
    'Trato todas desde la primera hasta una anterior a la actual
    'y veo si se parecen y si hay alguna muy parecida
    'disminuyo el peso a la actual
    'si el peso de la
    
    peso = peso_agente_pal(numero_de_frase_actual)
    frase_mas_parecida = 0
    numero_de_palabras_iguales_frase_mas_parecida = 0
    For cont_frase = 1 To numero_de_frase_actual - 1
        numero_de_palabras_iguales_curso = 0
        For cont_palabra = 1 To numero_de_palabras_frase_buscada
            If elemento_palabra_frase(cont_palabra, cont_frase) = elemento_palabra_frase(cont_palabra, numero_de_frase_actual) Then
                numero_de_palabras_iguales_curso = numero_de_palabras_iguales_curso + 1
            End If
DoEvents
        Next cont_palabra
        If numero_de_palabras_iguales_curso > numero_de_palabras_iguales_frase_mas_parecida Then
            'Si la frase que me hace la competencia, la frase mas parecida a la actual, tiene
            'menor o igual peso que la actual, entonces no se aplica la disminución de peso a
            'la actual
            'esto es pq aunque al principio están dsordenadas, al ir disminuyendo
            'algunos pesos, algunas por delante de una se quedan con poco peso
            If peso_agente_pal(cont_frase) > peso_agente_pal(numero_de_frase_actual) Then
                numero_de_palabras_iguales_frase_mas_parecida = numero_de_palabras_iguales_curso
                frase_mas_parecida = cont_frase
            End If
        End If
    Next cont_frase
    
    If frase_mas_parecida = 0 Then
        'No hay ninguna igual
        peso = peso
    Else
        If numero_de_palabras_iguales_frase_mas_parecida = numero_de_palabras_frase_buscada Then
            'Hay una que es exactamente igual
            peso = 0
        Else
            numero_de_palabras_distintas_frase_mas_parecida = numero_de_palabras_frase_buscada - numero_de_palabras_iguales_frase_mas_parecida
            peso = peso - (peso / (2 ^ numero_de_palabras_distintas_frase_mas_parecida))
        End If
    End If
    
    f_funcion_ajuste_pesos_frases_pal = peso
    


End Function
Function f_funcion_evaluacion_frases_pal(numero_de_frase As Integer) As Double

    Dim cont_frase As Integer
    Dim cont_palabra As Integer
    Dim peso As Double
    
    
    peso = 0
    For cont_palabra = 1 To numero_de_palabras_frase_buscada
        If elemento_palabra_frase(cont_palabra, numero_de_frase) = palabra_cadena_buscada(cont_palabra) Then
            peso = peso + 1
        End If
DoEvents
    Next cont_palabra
    peso = peso / numero_de_palabras_frase_buscada 'valor entre 0 y 1
    
    
    f_funcion_evaluacion_frases_pal = peso
    
End Function
Sub s_ajustar_pesos_frases_pal()

    Dim cont_frase As Integer
    
    'La de mas peso no se ajusta, ajustar siempre es disminuir
    'y se tienen en cuenta las anteriores a una dada
    For cont_frase = 2 To numero_total_de_agentes_ejv
        peso_agente_pal(cont_frase) = f_funcion_ajuste_pesos_frases_pal(cont_frase)
DoEvents
    Next cont_frase

End Sub
Sub s_evaluar_agentes_pal()

    Dim cont_frase As Integer
    
    For cont_frase = 1 To numero_total_de_agentes_ejv
        peso_agente_pal(cont_frase) = f_funcion_evaluacion_frases_pal(cont_frase)
        If ciclo_ejv = 1 Then
            frm_b2_pal.List2.AddItem Format(peso_agente_pal(cont_frase), "0.00000000")
        End If
DoEvents
    Next cont_frase


End Sub
Function f_palabra_esta_en_temp_pal(palabra_a_buscar As String, temp() As String, cont_temp As Integer) As Boolean


Dim i As Integer
Dim encontrado As Boolean

encontrado = False

For i = 1 To cont_temp
    If temp(i) = palabra_a_buscar Then
        encontrado = True
        Exit For
    End If
DoEvents
Next

f_palabra_esta_en_temp_pal = encontrado

End Function
Function f_existe_palabra_en_diccionario_pal(palabra_a_buscar As String) As Boolean

Dim i As Long
Dim encontrado As Boolean

encontrado = False

For i = 1 To numero_palabras_dicc_pal
    If palabra_del_diccionario(i) = palabra_a_buscar Then
        encontrado = True
        Exit For
    End If
DoEvents
Next

f_existe_palabra_en_diccionario_pal = encontrado


End Function
Sub s_comenzar_pal()

    Dim frase_actual As String
    Dim palabra_actual As String
    Dim cuenta_borrados As Integer
    
    Dim palabras_distintas As Integer
    
    Dim cont_frase As Integer
    Dim cont_palabra As Integer
    
    Dim texto As String
    Dim letra As String
    
    Dim i As Integer
    Dim temp() As String
    
        
    hay_que_detener_ejv = False
    esta_detenido_ejv = False

    frm_b2_inpal.Show
    frm_b2_inpal.Caption = "Información Frases"

    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True

    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, False
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, False
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, False
    s_cambiar_estado_enabled_menus_ejv CTE_VER_DICCIONARIO, False
    
    s_borrar_tiempo_comienzo
    
    s_grabar_dato_fichero_salida_ejv CTE_FIC_22_GLOXLS, CStr(f_contar_elementos(frase_a_buscar_pal, " ")), ContFilasHojaResumenExcel, 2
    s_grabar_dato_fichero_salida_ejv CTE_FIC_22_GLOXLS, CStr(CondParadaPesoNecesario_ejv), ContFilasHojaResumenExcel, 3
    
        
    'Control de errores de usuario
    If Len(frase_a_buscar_pal) < 1 Then
        MsgBox "Error: introduzca la frase a buscar", vbInformation
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, False
        'frm_b2_pal.Grafico.Enabled = False
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
        'Salvo y cierro los ficheros
        s_cerrar_ficheros_un_ejemplo_ejv
        Exit Sub
    End If
    If InStr(frase_a_buscar_pal, ",") Or InStr(frase_a_buscar_pal, ".") <> 0 Or InStr(frase_a_buscar_pal, ";") <> 0 Then
        MsgBox "Error: no se admiten signos de como "","", ""."", "";"", etc.", vbInformation
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, False
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, False
        'frm_b2_pal.Grafico.Enabled = False
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
        'Salvo y cierro los ficheros
        s_cerrar_ficheros_un_ejemplo_ejv
        Exit Sub
    End If
        
    cuenta_borrados = 0
    numero_de_palabras_frase_buscada = 0
    texto = frase_a_buscar_pal
    texto = texto & " "
    palabra_actual = ""
    While Len(texto) > 0
        'cogemos letra
        letra = Left(texto, 1)
        texto = Right(texto, Len(texto) - 1)
        'la añadimos
        If letra = " " Then
            numero_de_palabras_frase_buscada = numero_de_palabras_frase_buscada + 1
            ReDim Preserve palabra_cadena_buscada(1 To numero_de_palabras_frase_buscada) As String
            palabra_cadena_buscada(numero_de_palabras_frase_buscada) = palabra_actual
            'la añadimos también al diccionario, si no existe, pero machacando una para mantener
            'el tamaño del diccionario que se ha especificado
            If Not f_existe_palabra_en_diccionario_pal(palabra_actual) Then
                cuenta_borrados = cuenta_borrados + 1
                If numero_palabras_dicc_pal < cuenta_borrados Then
                    MsgBox "No es posible ejecutar. El número de palabras del diccionario debe ser mayor o igual que el numero de palabras diferentes de la frase a buscar", vbInformation
                    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, False
                    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, False
                    'frm_b2_pal.Grafico.Enabled = False
                    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True
                    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
                    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
                    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
                    s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
                    s_mostrar_estado_semaforo frm_b2_inpal, CTE_DETENIDO
                    Screen.MousePointer = CTE_DEFECTO
                    'Salvo y cierro los ficheros
                    s_cerrar_ficheros_un_ejemplo_ejv
                    Exit Sub
                Else
                    palabra_del_diccionario(cuenta_borrados) = palabra_actual
                End If
            End If
            palabra_actual = ""
        Else
            palabra_actual = palabra_actual & letra
        End If
    Wend
    
    
    'Recorremos la frase a buscar y contamos sus palabras distintas
    palabras_distintas = 0
    For i = 1 To numero_de_palabras_frase_buscada
        If Not f_palabra_esta_en_temp_pal(palabra_cadena_buscada(i), temp(), palabras_distintas) Then
            palabras_distintas = palabras_distintas + 1
            ReDim Preserve temp(1 To palabras_distintas) As String
            temp(palabras_distintas) = palabra_cadena_buscada(i)
        End If
DoEvents
    Next i
    
    If numero_palabras_dicc_pal < palabras_distintas Then
        MsgBox "Error: el número de palabras diferentes de la frase a buscar debe ser menor o igual que el numero de palabras del diccionario", vbInformation
        'frm_b2_pal.Grafico.Enabled = False
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True
        s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
        Exit Sub
        'Salvo y cierro los ficheros
        s_cerrar_ficheros_un_ejemplo_ejv
    End If
    
    
    'Creamos numero_de_frases_inicial cadenas y pesos
    ReDim peso_agente_pal(1 To numero_de_frases_inicial) As Double
    ReDim elemento_palabra_frase(1 To numero_de_palabras_frase_buscada, 1 To numero_de_frases_inicial) As String
    For cont_frase = 1 To numero_de_frases_inicial
        peso_agente_pal(cont_frase) = 0
        'de numero_de_palabras_por_frase elementos
        For cont_palabra = 1 To numero_de_palabras_frase_buscada
            elemento_palabra_frase(cont_palabra, cont_frase) = palabra_del_diccionario(fl_azar1(numero_palabras_dicc_pal))
DoEvents
        Next cont_palabra
    Next cont_frase
    
    
    numero_total_de_agentes_ejv = numero_de_frases_inicial
    
    s_bucle_general_pal
    Screen.MousePointer = CTE_DEFECTO


End Sub
Sub s_ordenar_pesos_frases_pal()


    Select Case algoritmo_ordenacion_ejv
        Case CTE_BURBUJA
            s_ordenar_pesos_frases_bur_pal
        Case CTE_QUICKSORT
            s_ordenar_pesos_frases_qui_pal
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Algoritmo de ordenación inexistente"
    End Select
    

End Sub
Sub s_ordenar_pesos_frases_bur_pal()

    Dim I_n As Integer
    Dim I_x As Integer
    Dim I_i As Integer
    Dim primero As Integer
    Dim ultimo As Integer
    Dim i_temp As String
    
    ' de > a <
    primero = 1
    ultimo = numero_total_de_agentes_ejv
    'Comparo cada elemento (todos menos el último)....
    For I_n = primero To ultimo - 1
       'con el siguiente y todos los demás hasta el último
        For I_x = I_n + 1 To ultimo
            If peso_agente_pal(I_x) > peso_agente_pal(I_n) Then
                'Cambio todos sus elementos
                For I_i = 1 To numero_de_palabras_frase_buscada
                    i_temp = elemento_palabra_frase(I_i, I_x)
                    elemento_palabra_frase(I_i, I_x) = elemento_palabra_frase(I_i, I_n)
                    elemento_palabra_frase(I_i, I_n) = i_temp
                Next I_i
                'Cambio su peso
                i_temp = peso_agente_pal(I_x)
                peso_agente_pal(I_x) = peso_agente_pal(I_n)
                peso_agente_pal(I_n) = i_temp
                
            End If
DoEvents
        Next I_x
    Next I_n

End Sub
Sub s_ordenar_pesos_frases_qui_pal()

    s_error_ejv CON_OPCION_FINALIZAR, "Quick Sort no programado"

    s_ordenar_qui_dos peso_agente_pal(), elemento_palabra_frase()


End Sub
Function f_hay_mutacion_pal() As Boolean
    
    'Ejemplos
    ' -7  es no
    '  0  es si
    '  1  es si
    '  2  es 50 % de las veces
    ' 10  es 10 % de las veces
    '100  es  1 % de las veces
    
    'independiente: la probb de mutación se calcula independientemente
    'de lo que haya pasado antes
    
    'Acumulada: la probabilidad de que se produzca una mutación es
    'el doble de la probabilidad de que se haya producido anteriormente,
    'hasta el límite de llegar a que la probabilidad sea la mitad (2)
    'donde se detiene. Si en algun caso no se produce mutación, la
    'probabilidad de mutación vuelve a ser la inicial
    
    
    Dim devolver As Boolean
    Dim temp As Integer

    devolver = False

    If tasa_de_mutacion_pal < 0 Then
        'Tasas de mutación negativas son sin mutación
         devolver = False
    Else
        If Not (mutaciones_acumuladas_pal) Then
            'mutación normal
             If tasa_de_mutacion_pal = 0 Or tasa_de_mutacion_pal = 1 Then
                 'Tasas de mutación 0 es siempre mutación
                  devolver = True
             Else
                 'Posible mutación
                 If fi_azar1(tasa_de_mutacion_pal) = 1 Then
                    'hay mutación
                     devolver = True
                 End If
            End If
       Else
            'mutación acumumulada
            'Si es la primera y antes no ha habido, guardamos el valor de la vieja
            If Not ha_habido_mutacion_anterior_pal Then
                tasa_de_mutacion_vieja_pal = tasa_de_mutacion_pal
            End If
             If tasa_de_mutacion_pal = 0 Or tasa_de_mutacion_pal = 1 Then
                 'Tasas de mutación 0 es siempre mutación
                  devolver = True
             Else
                 'Posible mutación
                 If fi_azar1(tasa_de_mutacion_pal) = 1 Then
                    'hay mutación
                      devolver = True
                      ha_habido_mutacion_anterior_pal = True
                      'hacemos que la proxima sea el doble mas probable, hasta 2 limite
                        temp = Int(tasa_de_mutacion_pal / 2) + 1
                        If temp >= 2 Then
                              tasa_de_mutacion_pal = temp
                        Else
                              tasa_de_mutacion_pal = 2
                        End If
                 Else
                      devolver = False
                      ha_habido_mutacion_anterior_pal = False
                      'recuperamos la tasa vieja
                      tasa_de_mutacion_pal = tasa_de_mutacion_vieja_pal
                 End If
            End If

       End If
    End If

    f_hay_mutacion_pal = devolver
    
End Function
Function f_son_padres_iguales_pal(padre1, padre2) As Boolean


Dim i As Integer
Dim son_iguales As Boolean

son_iguales = True

For i = 1 To numero_de_palabras_frase_buscada
    If elemento_palabra_frase(i, padre1) <> elemento_palabra_frase(i, padre2) Then
        son_iguales = False
        Exit For
    End If
DoEvents
Next

f_son_padres_iguales_pal = son_iguales

    

End Function
Sub s_reproduccion_frases_caso_pal(caso)


    Dim I_Indice1 As Integer
    Dim I_Indice2 As Integer
    
    Dim cont_palabra As Integer
    Dim i As Integer
    Dim generados As Integer
        
    Dim frontera As Integer
    Dim destino As Integer
    Dim ultimo As Integer
    
    Dim alterno As Boolean

    I_Indice1 = 1
    
    alterno = True
    
    Select Case caso
        Case "A"
            generados = 2
            frontera = Int(numero_total_de_agentes_ejv / 2)
            I_Indice2 = frontera + 1
        Case "B"
            generados = 18
            Select Case numero_total_de_agentes_ejv
                Case 8
                    generados = 6 'numero de hijos por cada pareja
                    frontera = 2 'ultimo de los supervivientes
                Case 20
                    frontera = 2
                Case 40
                    frontera = 4
                Case 80
                    frontera = 8
                Case 160
                    frontera = 16
                Case 320
                    frontera = 32
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: número de frases inicial no contemplado"
            End Select
        Case "C"
            generados = 8
            Select Case numero_total_de_agentes_ejv
                Case 8
                    generados = 6
                    frontera = 2
                Case 20
                    frontera = 4
                Case 40
                    frontera = 8
                Case 80
                    frontera = 16
                Case 160
                    frontera = 32
                Case 320
                    frontera = 64
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: número de frases inicial no contemplado"
            End Select
        Case "D"
            Select Case numero_total_de_agentes_ejv
                'destino = 40%
                'frontera = 40% + 50%
                Case 8
                    destino = 2
                    frontera = 2 + 4
                    ultimo = 8
                Case 20
                    destino = 8
                    frontera = 8 + 10
                    ultimo = 20
                Case 40
                    destino = 16
                    frontera = 16 + 20
                    ultimo = 40
                Case 80
                    destino = 32
                    frontera = 32 + 40
                    ultimo = 80
                Case 160
                    destino = 64
                    frontera = 64 + 80
                    ultimo = 160
                Case 320
                    destino = 128
                    frontera = 128 + 160
                    ultimo = 320
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error: número de frases inicial no contemplado"
            End Select
            I_Indice2 = frontera + 1
            'Copio el 10% final a partir del 40% y hago como en el caso A
            For i = I_Indice2 To ultimo
                destino = destino + 1
                For cont_palabra = 1 To numero_de_palabras_frase_buscada
                    elemento_palabra_frase(cont_palabra, destino) = elemento_palabra_frase(cont_palabra, i)
DoEvents
                Next cont_palabra
            Next i
            generados = 2
            frontera = Int(numero_total_de_agentes_ejv / 2)
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: caso de reproducción inexistente"
    End Select
    
    I_Indice2 = frontera + 1
    
   
    
    'Si los padres son al azar, los desordeno primero
    'solo los padres
    If eleccion_de_padres_al_azar_pal Then
        Si_DesordenMedioArray_2D_S elemento_palabra_frase(), numero_de_palabras_frase_buscada, 1, I_Indice2 - 1
    End If
    
    
    'Copio dos padres en dos hijos alternando uno si uno no
    While I_Indice1 < frontera
        If padres_identicos_producen_mutaciones_pal And f_son_padres_iguales_pal(I_Indice1, I_Indice1 + 1) Then
            'Si los padres son identicos, y se ha elegido la opción, los hijos son todo mutaciones
            'Cada pareja de supervivientes genera generados hijos
            For i = 0 To generados - 1
                For cont_palabra = 1 To numero_de_palabras_frase_buscada
                    elemento_palabra_frase(cont_palabra, I_Indice2 + i) = palabra_del_diccionario(fl_azar1(numero_palabras_dicc_pal))
DoEvents
                Next cont_palabra
            Next i
            I_Indice1 = I_Indice1 + 2
            I_Indice2 = I_Indice2 + generados
        Else
            'Reproduzco normalmente
            'Cada pareja de supervivientes genera 2 hijos
            For i = 0 To generados - 1
                'Cada frase tiene numero_de_palabras_frase_buscada palabras
                For cont_palabra = 1 To numero_de_palabras_frase_buscada
                    If (fi_azar1(2) = 1) Then
                        alterno = True
                    Else
                        alterno = False
                    End If
                    If alterno Then
                        'lo tomo del padre
                        elemento_palabra_frase(cont_palabra, I_Indice2 + i) = elemento_palabra_frase(cont_palabra, I_Indice1 + 1)
                    Else
                        'lo tomo de la madre
                        elemento_palabra_frase(cont_palabra, I_Indice2 + i) = elemento_palabra_frase(cont_palabra, I_Indice1)
                    End If
                    'Le añado una variación-mutación
                    If f_hay_mutacion_pal Then
                        elemento_palabra_frase(cont_palabra, I_Indice2 + i) = palabra_del_diccionario(fl_azar1(numero_palabras_dicc_pal))
                    End If
DoEvents
                Next cont_palabra
            Next i
            I_Indice1 = I_Indice1 + 2
            I_Indice2 = I_Indice2 + generados
        End If
    Wend



End Sub
Sub s_reproducir_frases_pal()


    Select Case tipo_reproduccion_pal
        Case 1
            'la última mitad de las ordenadas se eliminan
            'Reproduzco la primera mitad de las entidades haciendo que la primera
            'mitad borre a la segunda, combinandose las entidades de la primera
            'mitad por parejas, y intercambiando palabras
            'También hago una mutación cada cierto tiempo
            s_reproduccion_frases_caso_pal ("A")
        Case 2
            'Se selecciona el 10% primero, se borra y rellena el resto
            'Cada entidad genera otras 9
            'cada pareja genera 9 parejas
            'Cada 2 entidades generan 18
            'También hago una mutación cada cierto tiempo
            s_reproduccion_frases_caso_pal ("B")
        Case 3
            s_reproduccion_frases_caso_pal ("C")
        Case 4
            s_reproduccion_frases_caso_pal ("D")
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: caso de reproducción inexistente"
    End Select




End Sub
Sub s_bucle_general_pal()

    esta_detenido_ejv = False
    esta_terminado_ejv = False
    hay_que_detener_ejv = False
    hay_que_terminar_ejv = False

    'Pongo inhabilitado todos los programas
    s_cambiar_estado_enabled_programas_todos_ejv False

    'Primera Vez: tomamos el tiempo (aqui es siempre es la primera vez, si no iria en otro sitio)
    s_leer_tiempo_inicial_ejv
    s_mostrar_fecha_hora_comienzo_ejv
    s_leer_tiempo_final_ejv
    s_mostrar_fecha_hora_actual_ejv
    s_mostrar_estado_semaforo frm_b2_inpal, CTE_FUNCIONANDO
    
    'Pintamos las 20 primeras frases
    s_pintar_frases_pal
    
    ciclo_ejv = 0
    While hay_que_detener_ejv = False
        DoEvents
        ciclo_ejv = ciclo_ejv + 1
        'Se evaluan las entidades - soluciones  existentes
        s_evaluar_agentes_pal
        'Se clasifican las entidades en función de sus pesos de > a <
        s_ordenar_pesos_frases_pal
        'Una vez calculado el peso en la forma normal -independiente-
        'Si esta la opcion de calcularlo relativo
        If pesos_relativos_pal Then
            'Se ajustan los pesos de las entidades es como volver a evaluar
            s_ajustar_pesos_frases_pal
            'Se clasifican de nuevo las entidades en función de sus nuevos pesos
            s_ordenar_pesos_frases_pal
        End If
        'Pintamos las 20 primeras frases y almacenamos las graficas y
        'miramos si se ha llegado al objetivo
        s_pintar_frases_pal
        'Reproducción y mutaciones
        s_reproducir_frases_pal
        s_grabar_resumen_ejv
        s_mostrar_info_ejv
        s_condiciones_parada_ejv CTE_PARADA_POR_IGUAL
        'Muestro el numero de segundos trascurrido y la fecha-hora actual y la media de la duración de los ciclos
        s_mostrar_tiempo_transcurrido_ejv
    Wend


    
    'En el proceso de reproducción se desordenan algunas frases
    'Como al reproducir las frases no se desordenan los pesos, ya
    'que no se modifican (se desordenan las frases pero no los pesos)
    'no hace falta ordenar, y aunque la serie aparece ordenada,
    'no corresponde exactamente con la realidad.
    'Para que sea correcta, hay que evaluar y ordenar
    'justo antes de acabar
    'Se evaluan las entidades - soluciones  existentes
    s_evaluar_agentes_pal
    'Se clasifican las entidades en función de sus pesos de > a <
    s_ordenar_pesos_frases_pal
    
    s_mostrar_fecha_hora_actual_ejv
   
    frm_b2_pal.fr_Todas.Visible = False
    frm_b2_pal.Fr_Ejecucion.Visible = False
    frm_b2_pal.Fr_Opciones.Visible = True
    
    
    'Pongo habilitado todos los programas
    s_cambiar_estado_enabled_programas_todos_ejv True
    
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
    
    s_fin_bucle_general_ejv
    
End Sub
Sub s_activar_opciones_pal()

    Dim indice As Long
    Dim cargar_todo As Boolean

    cargar_todo = False
    If ha_cambiado_el_diccionaro_pal Then
        If MsgBox("¿Desea utilizar todas las palabras del diccionario elegido?", vbQuestion + vbYesNo) = vbYes Then
            cargar_todo = True
        End If
    End If
    ha_cambiado_el_diccionaro_pal = False
    'Cargar Diccionario
    indice = f_aut_leer_diccionario(cargar_todo)
    If control_errores_de_programacion_ejv Then
        If indice < numero_palabras_dicc_pal Then
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
        End If
    End If
    frm_b2_pal.num_pal.Text = numero_palabras_dicc_pal
    
End Sub

Sub s_grabar_opciones_pal()

    '1 Número de palabras del diccionario que hay que cargar
    numero_palabras_dicc_pal = CLng(frm_b2_pal.num_pal.Text)
    dicc_carpeta_pal = frm_b2_pal.Op_DiccC
    dicc_fichero_pal = frm_b2_pal.Op_DiccF
    '2 Frase a buscar
    frase_a_buscar_pal = Trim(frm_b2_pal.Text2.Text)
    '3 Número de cadenas inicial
    numero_de_frases_inicial = CInt(frm_b2_pal.Cb_num_pal.Text)
    '4 Probabilidad de Mutación
    tasa_de_mutacion_pal = CDbl(frm_b2_pal.Text4.Text)
    '5 Tipo de Reproducción
    If frm_b2_pal.Op_s1 Then
        tipo_reproduccion_pal = 1
    Else
        If frm_b2_pal.Op_s2 Then
            tipo_reproduccion_pal = 2
        Else
            If frm_b2_pal.Op_s3 Then
                tipo_reproduccion_pal = 3
            Else
                tipo_reproduccion_pal = 4
            End If
        End If
    End If
    '6 Mutaciones acumuladas
    mutaciones_acumuladas_pal = frm_b2_pal.Op_Acumulada
    '7 Elección de Padres
    eleccion_de_padres_al_azar_pal = frm_b2_pal.Op_PadresAzar_pal
    '8 Padres iguales producen hijos que son mutaciones en todos sus elementos.
    padres_identicos_producen_mutaciones_pal = frm_b2_pal.Op_PadresIdenticos
    '9 El peso se calcula en función de si existen más entidades parecidas a la actual
    pesos_relativos_pal = frm_b2_pal.Op_Relativo


End Sub
Sub s_cargar_opciones_pal()

    'Diccionario
    frm_b2_pal.Op_DiccC = dicc_carpeta_pal
    frm_b2_pal.Op_DiccF = dicc_fichero_pal
    'Frase a buscar
    frm_b2_pal.Text2.Text = frase_a_buscar_pal
    'Número de cadenas inicial
    frm_b2_pal.Cb_num_pal.Text = numero_de_frases_inicial
    'Probabilidad de Mutación
    frm_b2_pal.Text4.Text = tasa_de_mutacion_pal
    'Tipo de Reproducción
    frm_b2_pal.Op_s1 = False
    frm_b2_pal.Op_s2 = False
    frm_b2_pal.Op_s3 = False
    frm_b2_pal.Op_s4 = False
    Select Case tipo_reproduccion_pal
        Case 1
            frm_b2_pal.Op_s1 = True
        Case 2
            frm_b2_pal.Op_s2 = True
        Case 3
            frm_b2_pal.Op_s3 = True
        Case 4
            frm_b2_pal.Op_s4 = True
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: "
    End Select
    'Mutaciones acumuladas
    frm_b2_pal.Op_Acumulada = mutaciones_acumuladas_pal
    'Elección de Padres
    frm_b2_pal.Op_PadresAzar_pal = eleccion_de_padres_al_azar_pal
    frm_b2_pal.Op_nPadresAzar_pal = Not eleccion_de_padres_al_azar_pal
    'Padres iguales producen hijos que son mutaciones en todos sus elementos.
    frm_b2_pal.Op_PadresIdenticos = padres_identicos_producen_mutaciones_pal
    frm_b2_pal.Op_nPadresIdenticos = Not padres_identicos_producen_mutaciones_pal
    'El peso se calcula en función de si existen más entidades parecidas a la actual
    frm_b2_pal.Op_Relativo = pesos_relativos_pal
    frm_b2_pal.Op_nRelativo = Not pesos_relativos_pal

End Sub
Sub s_inicializar_ejemplo_elegido_pal()


    'OPCIONES I
    'GENERALES DE EJEMPLOS DE VIDA(DISTINTAS A LAS DE POR DEFECTO)
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
    un_ej_fichero_gra_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_GRA), "r_pal_" & num_ej_activo_ejv & ".gra")
    un_ej_grabar_resumen_txt_ejv = False
    un_ej_fichero_resumen_txt_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_TXT), "r_pal_" & num_ej_activo_ejv & ".txt")
    un_ej_grabar_resumen_xls_ejv = False
    un_ej_fichero_resumen_xls_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS), "r_pal_" & num_ej_activo_ejv & ".xls")
    max_guardado_ejv = 1000000
    autoguardado_ejv = 100

    'OPCIONES II
    'GENERALES DE PALABRAS
    Select Case num_ej_activo_ejv
        Case 1
            'Diccionario
            dicc_carpeta_pal = path_largo_ejv(CTE_C_ENT_DIC)
            dicc_fichero_pal = "2092.dic"
            numero_palabras_dicc_pal = 20
            'Frase a buscar
            frase_a_buscar_pal = "esta es una frase de siete palabras"
            'Número de cadenas inicial
            numero_de_frases_inicial = 40
            'Probabilidad de Mutación
            tasa_de_mutacion_pal = 5
            'Tipo de Reproducción
            tipo_reproduccion_pal = 1
            'Mutaciones acumuladas
            mutaciones_acumuladas_pal = True
            'Elección de Padres
            eleccion_de_padres_al_azar_pal = False
            'Padres iguales producen hijos que son mutaciones en todos sus elementos.
            padres_identicos_producen_mutaciones_pal = False
            'El peso se calcula en función de si existen más entidades parecidas a la actual
            pesos_relativos_pal = False
        Case 2
            'Diccionario
            dicc_carpeta_pal = path_largo_ejv(CTE_C_ENT_DIC)
            dicc_fichero_pal = "2092.dic"
            numero_palabras_dicc_pal = 20
            'Frase a buscar
            frase_a_buscar_pal = "esta es una frase de dieciseis palabras que habla sobre una frase que tiene dieciseis palabras"
            'Número de cadenas inicial
            numero_de_frases_inicial = 80
            'Probabilidad de Mutación
            tasa_de_mutacion_pal = 5
            'Tipo de Reproducción
            tipo_reproduccion_pal = 2
            'Mutaciones acumuladas
            mutaciones_acumuladas_pal = True
            'Elección de Padres
            eleccion_de_padres_al_azar_pal = False
            'Padres iguales producen hijos que son mutaciones en todos sus elementos.
            padres_identicos_producen_mutaciones_pal = False
            'El peso se calcula en función de si existen más entidades parecidas a la actual
            pesos_relativos_pal = False

        Case 3
            'Diccionario
            dicc_carpeta_pal = path_largo_ejv(CTE_C_ENT_DIC)
            dicc_fichero_pal = "2092.dic"
            numero_palabras_dicc_pal = 2092
            'Frase a buscar
            frase_a_buscar_pal = "una frase pequeña"
            'Número de cadenas inicial
            numero_de_frases_inicial = 160
            'Probabilidad de Mutación
            tasa_de_mutacion_pal = 2
            'Tipo de Reproducción
            tipo_reproduccion_pal = 3
            'Mutaciones acumuladas
            mutaciones_acumuladas_pal = True
            'Elección de Padres
            eleccion_de_padres_al_azar_pal = False
            'Padres iguales producen hijos que son mutaciones en todos sus elementos.
            padres_identicos_producen_mutaciones_pal = False
            'El peso se calcula en función de si existen más entidades parecidas a la actual
            pesos_relativos_pal = False
        Case 4
            'Diccionario
            dicc_carpeta_pal = path_largo_ejv(CTE_C_ENT_DIC)
            dicc_fichero_pal = "binario.dic"
            numero_palabras_dicc_pal = 2
            'Frase a buscar
            frase_a_buscar_pal = "0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 0 0 1 0 0 0 0 0 0 0 0 0 1"
            'Número de cadenas inicial
            numero_de_frases_inicial = 40
            'Probabilidad de Mutación
            tasa_de_mutacion_pal = 5
            'Tipo de Reproducción
            tipo_reproduccion_pal = 1
            'Mutaciones acumuladas
            mutaciones_acumuladas_pal = True
            'Elección de Padres
            eleccion_de_padres_al_azar_pal = True
            'Padres iguales producen hijos que son mutaciones en todos sus elementos.
            padres_identicos_producen_mutaciones_pal = False
            'El peso se calcula en función de si existen más entidades parecidas a la actual
            pesos_relativos_pal = False
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
    End Select

End Sub


Sub s_ver_opciones_pal()
    
    frm_b2_pal.fr_Todas.Visible = False
    frm_b2_pal.Fr_Ejecucion.Visible = False
    frm_b2_pal.Fr_Opciones.Visible = True

    frm_b2_pal.List1.Visible = False
    frm_b2_pal.List2.Visible = False
    frm_b2_pal.List3.Visible = False
    frm_b2_pal.List4.Visible = False

End Sub

Sub s_ver_grafico_pal()


End Sub

Sub s_ver_frases_pal()

    frm_b2_pal.fr_Todas.Visible = True
    frm_b2_pal.Fr_Ejecucion.Visible = False
    frm_b2_pal.Fr_Opciones.Visible = False

    s_pintar_todas_las_frases_pal

End Sub
