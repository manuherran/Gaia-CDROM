Attribute VB_Name = "bas_c3_3r"
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


'A quien toca jugar: X o O
Global turno_3r As String * 1

'tablero
Global estado_del_tablero_3r(1 To 9) As String * 1
Global lista_de_casillas_libres_3r(1 To 9) As Integer
Global numero_casillas_libres_3r As Integer
Global ganador_3r As String * 1

'estado del programa
Global estado_3r As Integer
Global se_han_creado_los_agentes_3r As Boolean

'Tasas
Global tasa_de_mutacion_3r As Integer
Global tipo_reproduccion_3r As Integer

'otras opciones
Global eleccion_de_padres_al_azar_3r As Boolean
Global eleccion_de_jugadores_al_azar_3r As Boolean
Global padres_identicos_producen_mutaciones_3r As Boolean
Global pesos_relativos_3r As Boolean
Global compartir_conocimiento_3r As Boolean
Global reglas_azar_3r As Boolean
Global ver_agentes_3r As Boolean
Global heredar_regla_mas_peso_3r As Boolean
Global quitar_reglas_repetidas_3r As Boolean
Global sust1_3r As Integer
Global sust2_3r As Integer
Global sust3_3r As Integer
Global NumeroPartidas_3r As Integer
Global Numero_cercanos_relativo_3r As Integer
Global Metodo_de_asignar_pesos_3r As String * 1
Global modificar_pesos_3r As Boolean
Global num_vecinos_3r As Integer
Global numero_reglas_variable_3r As Boolean '22
Global var11_3r As Integer
Global var12_3r As Integer
Global var21_3r As Integer
Global var22_3r As Integer
Global var31_3r As Integer
Global var32_3r As Integer
Global var41_3r As Integer
Global var42_3r As Integer
Global var51_3r As Integer
Global coger_alternos_3r As Boolean '21
Global personas_por_grupo_3r As String '22
Global Tipo_Mutacion_3r As Integer '24
Global Pesos_Partir_Cero_3r As Boolean '25


'Reglas y agentes
Global agente_3r() As String
Global numero_de_reglas_por_agente_3r As Integer
Global numero_de_caracteres_por_regla_3r As Integer
Global numero_de_caracteres_parte_izq_regla_3r As Integer

'pesos y otras propiedades de cada agente
Global peso_agente_ce0() As Long
Global peso_regla_agente_3r() As Long
Global prioridad_regla_agente_3r() As Integer
Global ciclo_nacimiento_agente_3r() As Long


Global historico_cambios_tamanio_agentes_3r As String

'subconjuntos de reglas
Global posibles_reglas_a_usar_3r() As Integer
Global numero_de_posibles_reglas_a_usar_3r As Integer

Global reglas_usadas_por_el_primero_3r() As Integer
Global numero_de_reglas_usadas_por_el_primero_3r As Integer

Global reglas_usadas_por_el_segundo_3r() As Integer
Global numero_de_reglas_usadas_por_el_segundo_3r As Integer

'Mutaciones
Global ha_habido_mutacion_anterior_3r As Boolean
Global mutaciones_acumuladas_3r As Boolean
Global tasa_de_mutacion_vieja_3r As Integer



'partidas
Global jugadas_3r As Long
Global ganadasO_3r As Long
Global ganadasX_3r As Long
Global tablas_3r As Long
Global jugadasreglas_3r As Long
Global rcompartidas_3r As Long
Global jugadasazar_3r As Long
Global le_ha_tocado_empezar_a_3r As Integer

Global paso_3r As Integer

Function f_mutar_regla(regla) As String

    Dim azar As Integer
    Dim dev As String
    
    If Tipo_Mutacion_3r < 0 Or Tipo_Mutacion_3r > 100 Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error: El tipo de mutación debe estar entre 0 y 100"
    End If
    
    azar = fi_azar1(100)
    If azar <= Tipo_Mutacion_3r Then
        dev = f_crear_regla_al_azar_3r
    Else
        dev = f_modificar_un_elemento_regla_al_azar_3r(regla)
    End If

    f_mutar_regla = dev

End Function
Sub s_peso_regla_disminuir(agente As Integer, regla As Integer)

    peso_regla_agente_3r(agente, regla) = peso_regla_agente_3r(agente, regla) - 1
 '   peso_regla_agente_3r(agente, regla) = peso_regla_agente_3r(agente, regla) / 2

End Sub
Sub s_peso_regla_aumentar(agente As Integer, regla As Integer)
    
    peso_regla_agente_3r(agente, regla) = peso_regla_agente_3r(agente, regla) + 1
'    peso_regla_agente_3r(agente, regla) = peso_regla_agente_3r(agente, regla) + ((100 - peso_regla_agente_3r(agente, regla)) / 2)

End Sub

Sub s_inicializar_juego_3r()

If estado_3r = CTE_JUGANDO Then
    frm_c3_juego3r.Label1.Visible = True
    frm_c3_juego3r.empieza.Visible = True
    frm_c3_juego3r.turno.Visible = True
    frm_c3_juego3r.Nueva.Visible = True
    frm_c3_juego3r.Text1.Visible = True
    frm_c3_juego3r.Label11.Visible = True
End If
    
frm_c3_juego3r.mensaje.ForeColor = &H0& 'negro
frm_c3_juego3r.mensaje.Caption = ""
    
If estado_3r = CTE_JUGANDO Then
    Dim N As Integer
    
    ganador_3r = "N"
    'inicializamos el tablero
    For N = 1 To numero_de_caracteres_parte_izq_regla_3r
        frm_c3_juego3r.B(N).Caption = ""
        estado_del_tablero_3r(N) = "V"
        lista_de_casillas_libres_3r(N) = N
    Next N
    numero_casillas_libres_3r = numero_de_caracteres_parte_izq_regla_3r
    s_jugar_contra_ordenador_3r
End If


End Sub
Function f_modificar_un_elemento_regla_al_azar_3r(regla)


Dim conclusion As String * 1
Dim i As Integer
ReDim posicion(1 To 9) As String * 1
Dim copia_regla As String
Dim dev As String
Dim azar_pos As Integer
Dim azar_tipo As Integer

'la desmenuzo
copia_regla = regla 'para no perderla
conclusion = Right(copia_regla, 1)
copia_regla = Left(copia_regla, 9)
For i = 1 To 9
    posicion(i) = Left(copia_regla, i)
    copia_regla = Right(copia_regla, Len(copia_regla) - 1)
Next i

'cojo una que no sea la conclusion
azar_pos = fi_azar1(9)
While azar_pos = conclusion
    azar_pos = fi_azar1(9)
Wend

'La modifico
azar_tipo = fi_azar1(4)
Select Case azar_tipo
    Case 1
        posicion(azar_pos) = "V"
    Case 2
        posicion(azar_pos) = "C"
    Case 3
        posicion(azar_pos) = "P"
    Case 4
        posicion(azar_pos) = "*"
    Case Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
End Select

dev = ""
For i = 1 To 9
   dev = dev & posicion(i)
Next i
dev = dev & conclusion

f_modificar_un_elemento_regla_al_azar_3r = dev

End Function
Function f_crear_regla_al_azar_3r() As String

'En vez de crear la regla totalmente al azar, la creamos
'con unas probabilidades que sabemos que serán más útiles,
'y asi ayudamos un poco a la evolución

'ejemplos de reglas que se espera que aparezcan:
'si hay dos en linea mias, pongo la tercera:
'***PPV*** --> 6
'si esta todo vacio, pongo en el centro
'VVVVVVVVV --> 5
'si hay dos en linea del contrario, pongo la tercera
'PPV****** --> 3

'leyenda:
'1: V Vacio
'2: * Indiferente
'3: P Ficha Propia
'4: C Ficha del Contrario

Dim devolver As String * 10
Dim azar As Integer


If reglas_azar_3r Then
    devolver = f_crear_regla_azar_metodo0 'azar mas real, aqui no hay trampa
Else
    'de todo un poco
    azar = fi_azar1(3)
    Select Case azar
        Case 1
            devolver = f_crear_regla_azar_metodo0 'azar mas real
        Case 2
            devolver = f_crear_regla_azar_metodo1 'las reglas más típicas (hacer trampa)
        Case 3
            devolver = f_crear_regla_azar_metodo2 'V10% *70% P10% C10%
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
    End Select
End If

f_crear_regla_al_azar_3r = devolver



End Function
Function f_crear_regla_azar_metodo0() As String

Dim devolver
Dim i As Integer
Dim i_casilla As Integer
Dim s_casilla As String * 1
Dim lista_casillas(1 To 9) As String * 1
Dim lista_destinos() As String * 1
Dim numero_destinos As Integer
Dim posicion_V_obligatoria As Integer


'Estas reglas se generan mas al azar, y la probabilidad
'de que aparezca cada posible valor en una celda
'es siempre igual
'V: 25
'*: 25
'P: 25
'C: 25


'De todas formas, ponemos una V obligatioria para que
'no se generen reglas no-validas. Es sobre todo para no perder
'ciclos de ejecución,

posicion_V_obligatoria = fi_azar1(9)

devolver = ""
For i = 1 To 9
    If i = posicion_V_obligatoria Then
        s_casilla = "V"
    Else
        i_casilla = fi_azar4(25, 25, 25, 25)
        Select Case i_casilla
            Case 1
                s_casilla = "V"
            Case 2
                s_casilla = "*"
            Case 3
                s_casilla = "P"
            Case 4
                s_casilla = "C"
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
        End Select
    End If
    lista_casillas(i) = s_casilla
    devolver = devolver & s_casilla
Next i

'Hacemos que la ficha colocada se coloque siempre en una casilla vacia
'para evitar reglas no validas
'hacemos una lista de todas las V donde puede caer y elegimos una
numero_destinos = 0
For i = 1 To 9
    If lista_casillas(i) = "V" Then
        numero_destinos = numero_destinos + 1
        ReDim Preserve lista_destinos(1 To numero_destinos) As String * 1
        lista_destinos(numero_destinos) = CStr(i)
    End If
Next i

'Tenemos numero_destinos posibles sitios donde poner nuestra ficha
If numero_destinos > 0 Then
    devolver = devolver & lista_destinos(fi_azar1(numero_destinos))
Else
    s_error_ejv CON_OPCION_FINALIZAR, "Error: número de destinos es " & numero_destinos
End If



f_crear_regla_azar_metodo0 = devolver


End Function
Function f_crear_regla_azar_metodo1() As String

'Creamos las reglas mas tipicas, con dos fichas y
'el destino en una tercera vacia
'Creamos una regla con dos P ó dos C y un V y el resto *

Dim devolver
Dim letra As String * 1
Dim azar As Integer
Dim azar2 As Integer

Dim i As Integer

Dim posicion_v As Integer
Dim posicion_1l As Integer
Dim posicion_2l As Integer


azar = fi_azar1(2)
If azar = 1 Then
    letra = "P"
Else
    letra = "C"
End If

posicion_v = 0
posicion_1l = 0
posicion_2l = 0
'las tres deben estar en linea para que la regla sea interesante
'hay 8 lineas posibles
azar = fi_azar1(8)
    Select Case azar
        'horizontal 1
        Case 1
            'la vacia puede estar en 1 de 3 posiciones, las otras dos son iguales
            'y no importa
            azar2 = fi_azar1(3)
                Select Case azar2
                    Case 1
                        posicion_v = 1
                        posicion_1l = 2
                        posicion_2l = 3
                    Case 2
                        posicion_v = 2
                        posicion_1l = 1
                        posicion_2l = 3
                    Case 3
                        posicion_v = 3
                        posicion_1l = 2
                        posicion_2l = 1
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
                 End Select
        'horizontal 2
        Case 2
            azar2 = fi_azar1(3)
                Select Case azar2
                    Case 1
                        posicion_v = 4
                        posicion_1l = 5
                        posicion_2l = 6
                    Case 2
                        posicion_v = 5
                        posicion_1l = 4
                        posicion_2l = 6
                    Case 3
                        posicion_v = 6
                        posicion_1l = 5
                        posicion_2l = 4
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
                 End Select
        'horizontal 3
        Case 3
            'la vacia puede estar en 1 de 3 posiciones, las otras dos son iguales
            'y no importa
            azar2 = fi_azar1(3)
                Select Case azar2
                    Case 1
                        posicion_v = 7
                        posicion_1l = 8
                        posicion_2l = 9
                    Case 2
                        posicion_v = 8
                        posicion_1l = 7
                        posicion_2l = 9
                    Case 3
                        posicion_v = 9
                        posicion_1l = 8
                        posicion_2l = 7
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
                 End Select
        'vertical 1
        Case 4
            azar2 = fi_azar1(3)
                Select Case azar2
                    Case 1
                        posicion_v = 1
                        posicion_1l = 4
                        posicion_2l = 7
                    Case 2
                        posicion_v = 4
                        posicion_1l = 1
                        posicion_2l = 7
                    Case 3
                        posicion_v = 7
                        posicion_1l = 4
                        posicion_2l = 1
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
                 End Select
        'vertical 2
        Case 5
            azar2 = fi_azar1(3)
                Select Case azar2
                    Case 1
                        posicion_v = 2
                        posicion_1l = 5
                        posicion_2l = 8
                    Case 2
                        posicion_v = 5
                        posicion_1l = 2
                        posicion_2l = 8
                    Case 3
                        posicion_v = 8
                        posicion_1l = 5
                        posicion_2l = 2
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
                 End Select
        'vertical 3
        Case 6
            azar2 = fi_azar1(3)
                Select Case azar2
                    Case 1
                        posicion_v = 3
                        posicion_1l = 6
                        posicion_2l = 9
                    Case 2
                        posicion_v = 6
                        posicion_1l = 3
                        posicion_2l = 9
                    Case 3
                        posicion_v = 9
                        posicion_1l = 6
                        posicion_2l = 3
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
                 End Select
        'diagonal 1
        Case 7
            azar2 = fi_azar1(3)
                Select Case azar2
                    Case 1
                        posicion_v = 1
                        posicion_1l = 5
                        posicion_2l = 9
                    Case 2
                        posicion_v = 5
                        posicion_1l = 1
                        posicion_2l = 9
                    Case 3
                        posicion_v = 9
                        posicion_1l = 5
                        posicion_2l = 1
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
                 End Select
        'diagonal 2
        Case 8
            azar2 = fi_azar1(3)
                Select Case azar2
                    Case 1
                        posicion_v = 3
                        posicion_1l = 5
                        posicion_2l = 7
                    Case 2
                        posicion_v = 5
                        posicion_1l = 3
                        posicion_2l = 7
                    Case 3
                        posicion_v = 7
                        posicion_1l = 5
                        posicion_2l = 3
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
                 End Select
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
    End Select

devolver = ""
For i = 1 To 9
    If i = posicion_v Then
        devolver = devolver & "V"
    Else
        If i = posicion_1l Or i = posicion_2l Then
            devolver = devolver & letra
        Else
            devolver = devolver & "*"
        End If
    End If
Next i

devolver = devolver & posicion_v


f_crear_regla_azar_metodo1 = devolver


End Function
Function f_crear_regla_azar_metodo2() As String

Dim devolver
Dim i As Integer
Dim i_casilla As Integer
Dim s_casilla As String * 1
Dim lista_casillas(1 To 9) As String * 1
Dim lista_destinos() As String * 1
Dim numero_destinos As Integer
Dim posicion_V_obligatoria As Integer


'Estas  reglas se generan mas al azar, pero la probabilidad
'de que aparezca cada posible valor en una celda
'es distinto:
'V: 10
'*: 70
'P: 10
'C: 10


posicion_V_obligatoria = fi_azar1(9)

devolver = ""
For i = 1 To 9
    If i <> posicion_V_obligatoria Then
        i_casilla = fi_azar4(10, 70, 10, 10)
        Select Case i_casilla
            Case 1
                s_casilla = "V"
            Case 2
                s_casilla = "*"
            Case 3
                s_casilla = "P"
            Case 4
                s_casilla = "C"
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: azar imposible"
        End Select
    Else
        s_casilla = "V"
    End If
    lista_casillas(i) = s_casilla
    devolver = devolver & s_casilla
Next i

'Hacemos que la ficha colocada se coloque siempre en una casilla vacia
'para evitar en lo posible reglas no validas
'hacemos una lista de todas las V donde puede caer y elegimos una
numero_destinos = 0
For i = 1 To 9
    If lista_casillas(i) = "V" Then
        numero_destinos = numero_destinos + 1
        ReDim Preserve lista_destinos(1 To numero_destinos) As String * 1
        lista_destinos(numero_destinos) = CStr(i)
    End If
Next i


'Tenemos numero_destinos posibles sitios donde poner nuestra ficha
If numero_destinos > 0 Then
    devolver = devolver & lista_destinos(fi_azar1(numero_destinos))
Else
    s_error_ejv CON_OPCION_FINALIZAR, "Error: número de destinos es " & numero_destinos
End If



f_crear_regla_azar_metodo2 = devolver


End Function
Sub s_crear_agentes_iniciales()

    Dim cont_agente As Integer
    Dim cont_regla As Integer
    
    'Creamos numero_total_de_agentes_ejv cadenas y pesos
    ReDim agente_3r(1 To numero_total_de_agentes_ejv) As String
    ReDim peso_agente_ce0(1 To numero_total_de_agentes_ejv) As Long
    ReDim ciclo_nacimiento_agente_3r(1 To numero_total_de_agentes_ejv) As Long
    ReDim peso_regla_agente_3r(1 To numero_total_de_agentes_ejv, 1 To numero_de_reglas_por_agente_3r) As Long
    ReDim prioridad_regla_agente_3r(1 To numero_total_de_agentes_ejv, 1 To numero_de_reglas_por_agente_3r) As Integer
    
    For cont_agente = 1 To numero_total_de_agentes_ejv
        DoEvents
        agente_3r(cont_agente) = ""
        peso_agente_ce0(cont_agente) = 0
        ciclo_nacimiento_agente_3r(cont_agente) = ciclo_ejv
        'de numero_de_reglas_por_agente_3r elementos
        For cont_regla = 1 To numero_de_reglas_por_agente_3r
            DoEvents
            agente_3r(cont_agente) = agente_3r(cont_agente) & f_crear_regla_al_azar_3r
            peso_regla_agente_3r(cont_agente, cont_regla) = f_media_pesos_primero_3r()
            prioridad_regla_agente_3r(cont_agente, cont_regla) = f_crear_prioridad_al_azar_3r()
        Next cont_regla
    Next cont_agente


End Sub
Sub s_pintar_todos_los_agentes_3r()

    'al final de la ejecución

    Dim agente As String
    Dim cont_agente As Integer
    Dim cont_regla As Integer
    Dim nlineas As Integer

    Dim mayor_valor_de_peso As Long
    Dim peso_regla As Long
    Dim prioridad_regla As Integer
    Dim regla_de_mayor_peso As Long
    Dim regla As String * 10

    Dim txt As String
    
    Dim uno As String
    Dim dos As String
    Dim tres As String
    
    Screen.MousePointer = CTE_ARENA
    s_mostrar_estado_semaforo frm_c3_in3r, CTE_MOSTRANDO
    frm_c0_ce.Lista5.Text = "" & vbCrLf
    
    'mostramos el contenido de todas
    nlineas = 0
    txt = ""
    For cont_agente = 1 To numero_total_de_agentes_ejv
        DoEvents
        If nlineas < CTE_MAX_LIN Then
            'agentes completos
            agente = agente_3r(cont_agente)
            If Len(agente) > CTE_MAX_CAR_LIN Then
                agente = Left(agente, CTE_MAX_CAR_LIN) & "..."
            End If
            txt = txt & vbCrLf
            nlineas = nlineas + 1
            txt = txt & "=======================================================================" & vbCrLf
            nlineas = nlineas + 1
            txt = txt & "Agente nº " & cont_agente & ", con peso " & Format(peso_agente_ce0(cont_agente), "0.00000000") & ", nac. " & ciclo_nacimiento_agente_3r(cont_agente) & vbCrLf
            nlineas = nlineas + 1
            txt = txt & agente & vbCrLf
            nlineas = nlineas + 1
            'la mejor regla de cada uno
            mayor_valor_de_peso = -1
            regla_de_mayor_peso = 0
            For cont_regla = 1 To numero_de_reglas_por_agente_3r
                DoEvents
                regla = f_tomar_regla_de_agente_3r(cont_agente, cont_regla)
                peso_regla = peso_regla_agente_3r(cont_agente, cont_regla)
                prioridad_regla = prioridad_regla_agente_3r(cont_agente, cont_regla)
                'todas sus reglas una por una
                uno = Left(regla, 3)
                dos = Mid(regla, 4, 3) & " --> " & Right(regla, 1) & " [ Regla nº " & cont_regla & ", con peso " & peso_regla & " y prioridad " & prioridad_regla & "]"
                tres = Mid(regla, 7, 3)
                txt = txt & vbCrLf
                nlineas = nlineas + 1
                txt = txt & uno & vbCrLf
                nlineas = nlineas + 1
                txt = txt & dos & vbCrLf
                nlineas = nlineas + 1
                txt = txt & tres & vbCrLf
                nlineas = nlineas + 1
                If peso_regla >= mayor_valor_de_peso Then
                    mayor_valor_de_peso = peso_regla
                    regla_de_mayor_peso = cont_regla
                End If
            Next cont_regla
            regla = f_tomar_regla_de_agente_3r(cont_agente, regla_de_mayor_peso)
            txt = txt & vbCrLf
            nlineas = nlineas + 1
            txt = txt & "-----------------------------------------------------------------------" & vbCrLf
            nlineas = nlineas + 1
            txt = txt & vbCrLf
            nlineas = nlineas + 1
            txt = txt & "La mejor regla de este agente es la nº " & regla_de_mayor_peso & ", de peso " & mayor_valor_de_peso & " y prioridad " & prioridad_regla_agente_3r(cont_agente, regla_de_mayor_peso) & vbCrLf
            nlineas = nlineas + 1
            uno = Left(regla, 3)
            dos = Mid(regla, 4, 3) & " --> " & Right(regla, 1) & " [ Regla nº " & regla_de_mayor_peso & ", con peso " & mayor_valor_de_peso & " ]"
            tres = Mid(regla, 7, 3)
            txt = txt & vbCrLf
            nlineas = nlineas + 1
            txt = txt & uno & vbCrLf
            nlineas = nlineas + 1
            txt = txt & dos & vbCrLf
            nlineas = nlineas + 1
            txt = txt & tres & vbCrLf
            nlineas = nlineas + 1
        End If
    Next cont_agente

    If Len(txt) > MAX_LISTA Then
        txt = Left(txt, MAX_LISTA)
    End If
    frm_c0_ce.Lista5.Text = txt

    Screen.MousePointer = CTE_DEFECTO


End Sub
Sub s_control_de_errores_en_programacion()
    
    Dim regla  As String
    Dim peso As Long
    
    
    Dim cont_agente As Integer
    Dim cont_regla As Integer


    For cont_agente = 1 To numero_total_de_agentes_ejv
        DoEvents
        If Len(agente_3r(cont_agente)) <> numero_de_reglas_por_agente_3r * numero_de_caracteres_por_regla_3r Then
            If 5 / 0 = 4 Then Beep
        End If
    
        For cont_regla = 1 To numero_de_reglas_por_agente_3r
            DoEvents
            peso = peso_regla_agente_3r(cont_agente, cont_regla)
            If peso > f_media_pesos_primero_3r() Then
               regla = Left(f_tomar_regla_de_agente_3r(cont_agente, cont_regla), 9)
               If Left(regla, 3) = "CCC" Then
                   If 5 / 0 = 4 Then Beep
               End If
               If Right(regla, 3) = "CCC" Then
                   If 5 / 0 = 4 Then Beep
               End If
            End If
        Next cont_regla
    Next cont_agente


End Sub
Sub s_botones_activos_3r(estado As Boolean)
    
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_MEJORES, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES3, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION, estado
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_EVALUACION, estado
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_SELECCION, estado
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION, estado
            s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES, estado
            s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO, estado
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_JUGAR_CONTRA_ORDENADOR, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_MEJORES, estado
    
    frm_c0_ce.super.Enabled = estado
    
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, Not estado
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, Not estado
    s_cambiar_estado_enabled_menus_ejv CTE_VER_MODIFICAR_AGENTE, estado
    
End Sub

Sub s_modificar_tamanio_agentes_3r()
    
    Dim hay_cambio As Boolean
    Dim viejo_numero As Integer
    Dim nuevo_numero As Integer
    
    Dim cont_agente As Integer
    Dim cont_regla As Integer
            
    hay_cambio = False
    viejo_numero = numero_de_reglas_por_agente_3r
    
    If numero_reglas_variable_3r = True Then
        If ciclo_ejv = 0 Then
            nuevo_numero = var11_3r
            hay_cambio = True
        End If
        If ciclo_ejv = var12_3r Then
            nuevo_numero = var21_3r
            hay_cambio = True
        End If
        If ciclo_ejv = var12_3r + var22_3r Then
            nuevo_numero = var31_3r
            hay_cambio = True
        End If
        If ciclo_ejv = var12_3r + var22_3r + var32_3r Then
            nuevo_numero = var41_3r
            hay_cambio = True
        End If
        If ciclo_ejv = var12_3r + var22_3r + var32_3r + var42_3r Then
            nuevo_numero = var51_3r
            hay_cambio = True
        End If
    End If


    If hay_cambio And ciclo_ejv <> 0 And nuevo_numero <> viejo_numero Then
    
        historico_cambios_tamanio_agentes_3r = historico_cambios_tamanio_agentes_3r & "En el ciclo " & CStr(ciclo_ejv) & " el viejo era " & CStr(viejo_numero) & " y pasamos a " & CStr(nuevo_numero) & ". "
        If nuevo_numero < viejo_numero Then
            numero_de_reglas_por_agente_3r = nuevo_numero
            ReDim Preserve peso_regla_agente_3r(1 To numero_total_de_agentes_ejv, 1 To numero_de_reglas_por_agente_3r) As Long
            ReDim Preserve prioridad_regla_agente_3r(1 To numero_total_de_agentes_ejv, 1 To numero_de_reglas_por_agente_3r) As Integer
        Else
            'Si los nuevos agentes son mas grandes, hay que añadir reglas
            ReDim Preserve peso_regla_agente_3r(1 To numero_total_de_agentes_ejv, 1 To nuevo_numero) As Long
            ReDim Preserve prioridad_regla_agente_3r(1 To numero_total_de_agentes_ejv, 1 To nuevo_numero) As Integer
            For cont_agente = 1 To numero_total_de_agentes_ejv
                For cont_regla = numero_de_reglas_por_agente_3r + 1 To nuevo_numero
                    DoEvents
                    agente_3r(cont_agente) = agente_3r(cont_agente) & f_crear_regla_al_azar_3r
                    peso_regla_agente_3r(cont_agente, cont_regla) = f_media_pesos_primero_3r()
                    prioridad_regla_agente_3r(cont_agente, cont_regla) = f_crear_prioridad_al_azar_3r()
                Next cont_regla
            Next cont_agente
            numero_de_reglas_por_agente_3r = nuevo_numero
        End If
    End If



End Sub
Sub s_quitar_reglas_peso_bajo()

    Dim cont_agente As Integer
    Dim cont_regla As Integer
    Dim agente_nuevo As String

    'Sustituir cada ciclo por mutaciones las reglas con un peso menor o igual que ___
    If sust1_3r > 0 Then
        'para cada agente
        For cont_agente = 1 To numero_total_de_agentes_ejv
            frm_c3_in3r.entidad = cont_agente
    
            agente_nuevo = ""
            'para cada regla de cada agente
            For cont_regla = 1 To numero_de_reglas_por_agente_3r
                DoEvents
                If peso_regla_agente_3r(cont_agente, cont_regla) <= ((f_media_pesos_primero_3r * sust1_3r) / 100) Then
                    agente_nuevo = agente_nuevo & f_mutar_regla(f_tomar_regla_de_agente_3r(cont_agente, cont_regla))
                    peso_regla_agente_3r(cont_agente, cont_regla) = f_media_pesos_primero_3r()
                    prioridad_regla_agente_3r(cont_agente, cont_regla) = f_crear_prioridad_al_azar_3r()
                Else
                    agente_nuevo = agente_nuevo & f_tomar_regla_de_agente_3r(cont_agente, cont_regla)
                End If
            Next cont_regla
        Next cont_agente
    End If
    
    'sustituir cada ___ ciclos por mutaciones las reglas con un peso menor o igual que ___
    If sust2_3r > 0 And sust3_3r > 0 Then
        If ciclo_ejv Mod sust2_3r = 0 Then
            'para cada agente
            For cont_agente = 1 To numero_total_de_agentes_ejv
                frm_c3_in3r.entidad = cont_agente
        
                agente_nuevo = ""
                'para cada regla de cada agente
                For cont_regla = 1 To numero_de_reglas_por_agente_3r
                    DoEvents
                    If peso_regla_agente_3r(cont_agente, cont_regla) <= ((f_media_pesos_primero_3r * sust3_3r) / 100) Then
                        agente_nuevo = agente_nuevo & f_mutar_regla(f_tomar_regla_de_agente_3r(cont_agente, cont_regla))
                        peso_regla_agente_3r(cont_agente, cont_regla) = f_media_pesos_primero_3r()
                        prioridad_regla_agente_3r(cont_agente, cont_regla) = f_crear_prioridad_al_azar_3r()
                    Else
                        agente_nuevo = agente_nuevo & f_tomar_regla_de_agente_3r(cont_agente, cont_regla)
                    End If
                Next cont_regla
            Next cont_agente
        End If
    End If


End Sub

Function f_crear_prioridad_al_azar_3r() As Integer
    f_crear_prioridad_al_azar_3r = fi_azar1(CTE_NUMERO_DE_NIVELES_DE_PRIORIDAD)
End Function

Sub s_mostrar_info_3r()

    frm_c3_juego3r.tipo1 = ""
    frm_c3_juego3r.tipo2 = ""
    frm_c3_juego3r.tipo3 = ""
    frm_c3_juego3r.tipo4 = ""

    
    'Muestro en pantalla los agentes 1 y 2
    If ver_agentes_3r Then
        frm_c3_juego3r.numtipo1_1 = resumen_actual(1) 'agente 1 buenas
        frm_c3_juego3r.numtipo1_2 = resumen_actual(2) 'agente 1 regulares
        frm_c3_juego3r.numtipo2_1 = resumen_actual(3) 'agente 2 buenas
        frm_c3_juego3r.numtipo2_2 = resumen_actual(4) 'agente 2 regulares
    End If
    

End Sub
Sub s_grabar_resumen_3r()

    'Cuento el numero de reglas del tipo (1) muy buenas
    ' PPV****** -> 3
    ' y el numero de reglas del tipo (2) regulares
    ' PPV***P** -> 3
    ' de las entidades 1 y 2 y lo guardo
    ' también lo muestro en la pantalla
    
    '1: 1 buenas
    '2: 1 regulares
    '3: 2 buenas
    '4: 2 regulares
    '5: numero de movimientos con reglas
    '6: peso del agente de mas peso
    
    ReDim resumen_actual(1 To 6) As Long
    Dim cont_agente As Integer
    Dim cont_regla As Integer
    Dim tipo As Integer
    Dim regla As String
    Dim i As Integer
    Dim linea As String
    

    'Preparamos datos
    'Inicializo
    For i = 1 To 6
        resumen_actual(i) = 0
    Next i
    'Para cada agente de los dos primeros
    For cont_agente = 1 To 2
        'Para cada regla de cada agente
        For cont_regla = 1 To numero_de_reglas_por_agente_3r
            DoEvents
            'Miro si es de las buenas
            regla = f_tomar_regla_de_agente_3r(cont_agente, cont_regla)
            tipo = f_dime_el_tipo_3r(regla)
            If cont_agente = 1 And tipo = 1 Then
                resumen_actual(1) = resumen_actual(1) + 1
                frm_c3_juego3r.tipo1 = frm_c3_juego3r.tipo1 & regla
            End If
            If cont_agente = 1 And tipo = 2 Then
                resumen_actual(2) = resumen_actual(2) + 1
                frm_c3_juego3r.tipo2 = frm_c3_juego3r.tipo2 & regla
            End If
            If cont_agente = 2 And tipo = 1 Then
                resumen_actual(3) = resumen_actual(3) + 1
                frm_c3_juego3r.tipo3 = frm_c3_juego3r.tipo3 & regla
            End If
            If cont_agente = 2 And tipo = 2 Then
                resumen_actual(4) = resumen_actual(4) + 1
                frm_c3_juego3r.tipo4 = frm_c3_juego3r.tipo4 & regla
            End If
        Next cont_regla
    Next cont_agente
    '5: numero de movimientos con reglas
    resumen_actual(5) = jugadasreglas_3r
    '6: peso del agente de mas peso
    resumen_actual(6) = Format(peso_agente_ce0(1), "0.00000000")
    'Grabamos datos
    linea = ""
    linea = linea & f_comillas(CStr(ciclo_ejv)) ' el ciclo actual
    For i = 1 To 6
        linea = linea & ";" & f_comillas(CStr(resumen_actual(i)))
    Next i
    s_grabar_dato_fichero_salida_ejv CTE_FIC_23W_1EJGRA, linea

    

End Sub
Function f_dime_el_tipo_3r(regla As String) As Integer

'Cuento el numero de reglas del tipo (1) muy buenas
' PPV****** -> 3
' y el numero de reglas del tipo (2) regulares
' PPV***P** -> 3
' de las entidades 1 y 2 y lo guardo
' también lo muestro en la pantalla

'si es mala devuelvo 0

'1: 1 buenas
'2: 1 regulares
'3: 2 buenas
'4: 2 regulares


Dim conclusion As String * 1
Dim i As Integer
ReDim posicion(1 To 9) As String * 1
Dim num_P As Integer
Dim num_C As Integer
Dim num_V As Integer
Dim cont_celda As Integer
Dim copia_regla As String

Dim cont_rayas As Integer
Dim c1 As Integer
Dim c2 As Integer
Dim c3 As Integer

Dim tipo As Integer

'la desmenuzo
copia_regla = regla 'para no perderla
conclusion = Right(copia_regla, 1)
copia_regla = Left(copia_regla, 9)
For i = 1 To 9
    posicion(i) = Left(copia_regla, i)
    copia_regla = Right(copia_regla, Len(copia_regla) - 1)
Next i

'Cuento las P y C, si no hay por lo menos 2 P y 1 V o 2 C y 1 V
'no vale; si hay eso puede ser de las buenas y si hay mas puede ser de las regulares
num_P = 0
num_C = 0
num_V = 0
For cont_celda = 1 To 9
    If posicion(cont_celda) = "P" Then
        num_P = num_P + 1
    Else
        If posicion(cont_celda) = "C" Then
            num_C = num_C + 1
        Else
            If posicion(cont_celda) = "V" Then
                num_V = num_V + 1
            End If
        End If
    End If
Next cont_celda


'Veo el tipo
tipo = 0

'Por cada linea
For cont_rayas = 1 To 8
    Select Case cont_rayas
        Case 1
            c1 = 1
            c2 = 2
            c3 = 3
        Case 2
            c1 = 4
            c2 = 5
            c3 = 6
        Case 3
            c1 = 7
            c2 = 8
            c3 = 9
        Case 4
            c1 = 1
            c2 = 4
            c3 = 7
        Case 5
            c1 = 2
            c2 = 5
            c3 = 8
        Case 6
            c1 = 3
            c2 = 6
            c3 = 9
        Case 7
            c1 = 1
            c2 = 5
            c3 = 9
        Case 8
            c1 = 3
            c2 = 5
            c3 = 7
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: rayas imposibles"
    End Select
    
    'Veo las combinaciones de ese caso
    'VPP
    'PVP
    'PPV
    If posicion(c1) = "V" And conclusion = c1 Then
        If posicion(c2) = posicion(c3) And (posicion(c2) = "P" Or posicion(c2) = "C") Then
            tipo = 2
        End If
    ElseIf posicion(c2) = "V" And conclusion = c2 Then
        If posicion(c1) = posicion(c3) And (posicion(c1) = "P" Or posicion(c1) = "C") Then
            tipo = 2
        End If
    ElseIf posicion(c3) = "V" And conclusion = c3 Then
        If posicion(c1) = posicion(c2) And (posicion(c1) = "P" Or posicion(c1) = "C") Then
            tipo = 2
        End If
    End If
Next cont_rayas


If tipo = 2 Then
    'es de tipo 2, lo tiene en algun sitio
    If num_P + num_C + num_V = 3 Then
        'si solo tiene eso bueno, es de tipo 1
        tipo = 1
    End If
End If


f_dime_el_tipo_3r = tipo


End Function
Sub s_pintar_mejores_agentes_3r()

    'durante la ejecución

    Dim agente As String
    Dim cont_agente As Integer
    Dim cont_regla As Integer
        
    Dim peso_de_la_regla_de_mas_peso_3r As Long
    Dim regla_de_mas_peso As String
    Dim regla As String
    
    Dim txt As String
    
    Dim peso As String
    Dim nacimiento As String
    Dim numero As String
    
    Dim num_mostrar As Integer
    
    num_mostrar = 20
    If numero_total_de_agentes_ejv < 20 Then
        num_mostrar = numero_total_de_agentes_ejv
    End If
    
    frm_c3_juego3r.Label1.Visible = False
    frm_c3_juego3r.empieza.Visible = False
    frm_c3_juego3r.turno.Visible = False
    frm_c3_juego3r.Nueva.Visible = False
    frm_c3_juego3r.Text1.Visible = False
    frm_c3_juego3r.Label11.Visible = False
    
    frm_c0_ce.fr_Todas.Visible = False
    frm_c0_ce.Fr_Ejecucion.Visible = True

    frm_c0_ce.Lista1.Visible = True
    frm_c0_ce.Lista3.Visible = True
    
    
    
    txt = "Num     Peso   Nacido <Regla 1 ><Regla 2 ><Regla 3 ><Regla 4 ><Regla 5 ><Regla 6 ><Regla 7 >" & vbCrLf
    'mostramos información sobre las 20 primeras entidades
    For cont_agente = 1 To num_mostrar
        DoEvents
        agente = agente_3r(cont_agente)
        If Len(agente) > CTE_MAX_CAR_LIN Then
            agente = Left(agente, CTE_MAX_CAR_LIN) & "..."
        End If
        numero = CStr(cont_agente)
        peso = Left(Format(peso_agente_ce0(cont_agente), "0.00000000"), 8)
        nacimiento = CStr(ciclo_nacimiento_agente_3r(cont_agente))
        While Len(numero) < 3
            numero = " " & numero
        Wend
        While Len(peso) < 8
            peso = " " & peso
        Wend
        While Len(nacimiento) < 8
            nacimiento = " " & nacimiento
        Wend
        txt = txt & numero & " " & peso & " " & nacimiento & " " & agente & "    " & vbCrLf
    Next cont_agente
    frm_c0_ce.Lista1.Text = txt
    
    txt = ""
    'mostramos el contenido de las 20 últimas
    For cont_agente = numero_total_de_agentes_ejv - num_mostrar + 1 To numero_total_de_agentes_ejv
        If numero_total_de_agentes_ejv <= 20 Then Exit For
        DoEvents
        agente = agente_3r(cont_agente)
        If Len(agente) > CTE_MAX_CAR_LIN Then
            agente = Left(agente, CTE_MAX_CAR_LIN) & "..."
        End If
        numero = CStr(cont_agente)
        peso = Left(Format(peso_agente_ce0(cont_agente), "0.00000000"), 8)
        nacimiento = CStr(ciclo_nacimiento_agente_3r(cont_agente))
        While Len(numero) < 3
            numero = " " & numero
        Wend
        While Len(peso) < 8
            peso = " " & peso
        Wend
        While Len(nacimiento) < 8
            nacimiento = " " & nacimiento
        Wend
        txt = txt & numero & " " & peso & " " & nacimiento & " " & agente & "    " & vbCrLf
    Next cont_agente
    frm_c0_ce.Lista3.Text = txt
        
   '(este codigo esta repetido)
        
   'Mostramos la regla de mas peso del agente 1
    peso_de_la_regla_de_mas_peso_3r = -1
    regla_de_mas_peso = ""
    For cont_regla = 1 To numero_de_reglas_por_agente_3r
        DoEvents
        regla = f_tomar_regla_de_agente_3r(1, cont_regla)
        If peso_regla_agente_3r(1, cont_regla) > peso_de_la_regla_de_mas_peso_3r Then
            peso_de_la_regla_de_mas_peso_3r = peso_regla_agente_3r(1, cont_regla)
            regla_de_mas_peso = regla
        End If
    Next cont_regla
    frm_c3_juego3r.r1 = regla_de_mas_peso
    frm_c3_juego3r.p1 = peso_de_la_regla_de_mas_peso_3r
    
   'Mostramos la regla de mas peso del agente 2
    peso_de_la_regla_de_mas_peso_3r = -1
    regla_de_mas_peso = ""
    For cont_regla = 1 To numero_de_reglas_por_agente_3r
        DoEvents
        regla = f_tomar_regla_de_agente_3r(2, cont_regla)
        If peso_regla_agente_3r(2, cont_regla) > peso_de_la_regla_de_mas_peso_3r Then
            peso_de_la_regla_de_mas_peso_3r = peso_regla_agente_3r(2, cont_regla)
            regla_de_mas_peso = regla
        End If
    Next cont_regla
    frm_c3_juego3r.r2 = regla_de_mas_peso
    frm_c3_juego3r.p2 = peso_de_la_regla_de_mas_peso_3r



End Sub
Sub s_evaluar_agentes_3r()
    
Dim ganador As Integer
Dim contrario As Integer

Dim primero As Integer
Dim segundo As Integer
    
    
'Evaluar es jugar N partidas de 2 en 2
'Primero los desordenamos y despues
'los cogemos de 2 en 2

'Inicializamos el juego de los agentes
'Estos contadores son relativos a todas las partidas en un ciclo
jugadas_3r = 0
ganadasO_3r = 0
ganadasX_3r = 0
tablas_3r = 0
jugadasreglas_3r = 0
rcompartidas_3r = 0
jugadasazar_3r = 0


'Si los jugadores se cogen al azar, los desordeno primero
'se desordenan todos
If eleccion_de_jugadores_al_azar_3r Then
    frm_c3_in3r.accion = "Desordenando"
    s_desordenar_agentes_3r
End If

frm_c3_in3r.accion = "Evaluando"


'Si los pesos parten de cero
If Pesos_Partir_Cero_3r Then
    For agente_actual_ejv = 1 To numero_total_de_agentes_ejv
        DoEvents
        peso_agente_ce0(agente_actual_ejv) = 0
    Next agente_actual_ejv
End If


If personas_por_grupo_3r = "2 Jugadores" Then
    'provoco partidas de 2 en 2
    For agente_actual_ejv = 1 To numero_total_de_agentes_ejv Step 2
        DoEvents
        primero = agente_actual_ejv
        segundo = agente_actual_ejv + 1
        DoEvents
        frm_c3_in3r.entidad = primero & "-" & segundo
        If ciclo_ejv = 1 And ver_agentes_3r Then
            frm_c3_juego3r.Show
        End If
        'aqui es donde juegan N partidas los 2 agentes
        'aqui ganador es el ganador
        'supuesto como el que mas partidas de una serie ha ganado
        'y es un dato que no se usa
        ganador = f_funcion_evaluacion_agentes_3r(primero, segundo)
        'Mostramos la regla de mas peso del agente primero y segundo
        s_mostrar_regla_de_mas_peso 1, primero
        s_mostrar_regla_de_mas_peso 2, segundo
        'Muestro el numero de segundos trascurrido y la fecha-hora actual
        s_mostrar_tiempo_transcurrido_ejv
    Next agente_actual_ejv
    
ElseIf personas_por_grupo_3r = "Todos contra todos" Then
    'provoco partidas todos con todos
    For agente_actual_ejv = 1 To numero_total_de_agentes_ejv - 1
        For contrario = agente_actual_ejv + 1 To numero_total_de_agentes_ejv
            primero = agente_actual_ejv
            segundo = contrario
            DoEvents
            frm_c3_in3r.entidad = primero & "-" & segundo
            If ver_agentes_3r And ciclo_ejv = 1 Then
                frm_c3_juego3r.Show
            End If
            'aqui es donde juegan N partidas los 2 agentes
            'aqui ganador es el ganador
            'supuesto como el que mas partidas de una serie ha ganado
            'y es un dato que no se usa
            ganador = f_funcion_evaluacion_agentes_3r(primero, segundo)
            'Mostramos la regla de mas peso del agente primero y segundo
            s_mostrar_regla_de_mas_peso 1, primero
            s_mostrar_regla_de_mas_peso 2, segundo
        Next contrario
        'Muestro el numero de segundos trascurrido y la fecha-hora actual
        s_mostrar_tiempo_transcurrido_ejv
    Next agente_actual_ejv
Else
    s_error_ejv CON_OPCION_FINALIZAR, "Error: Personas por grupo no existente"
End If


'Muestro resultados de estas partidas
frm_c3_in3r.jugadas = jugadas_3r
frm_c3_in3r.ganadasO = ganadasO_3r
frm_c3_in3r.ganadasX = ganadasX_3r
frm_c3_in3r.tablas = tablas_3r
frm_c3_in3r.jugadasreglas = jugadasreglas_3r
frm_c3_in3r.rcompartidas = rcompartidas_3r
frm_c3_in3r.jugadasazar = jugadasazar_3r


End Sub
Function f_funcion_evaluacion_agentes_3r(primero As Integer, segundo As Integer) As Integer

Dim i As Integer
Dim resultado As Integer

 
'Se juegan N partidas
For i = 1 To NumeroPartidas_3r
    DoEvents
    numero_de_reglas_usadas_por_el_primero_3r = 0
    numero_de_reglas_usadas_por_el_segundo_3r = 0
    ReDim reglas_usadas_por_el_primero_3r(1 To 1) As Integer
    ReDim reglas_usadas_por_el_segundo_3r(1 To 1) As Integer
    
    'Jugamos una
    resultado = f_jugar_una_partida_3r(primero, segundo)
    jugadas_3r = jugadas_3r + 1
    
    'Muestro resultados de estas partidas
    frm_c3_in3r.jugadas = jugadas_3r
    
    Select Case resultado
        Case CTE_GANA_EL_PRIMERO
           '1: Ha ganado el primero (que juega con O)
           'Nota: es el primero en la lista, no el primero que ha hecho un movimiento
            ganadasO_3r = ganadasO_3r + 1
           'ganador
            peso_agente_ce0(primero) = peso_agente_ce0(primero) + f_incremento_de_peso(CTE_GANAR, 1)
           'perdedor
            peso_agente_ce0(segundo) = peso_agente_ce0(segundo) + f_incremento_de_peso(CTE_PERDER, 2)
        Case CTE_GANA_EL_SEGUNDO
            '2: Ha ganado el segundo
             ganadasX_3r = ganadasX_3r + 1
            'perdedor
             peso_agente_ce0(primero) = peso_agente_ce0(primero) + f_incremento_de_peso(CTE_PERDER, 1)
            'ganador
             peso_agente_ce0(segundo) = peso_agente_ce0(segundo) + f_incremento_de_peso(CTE_GANAR, 2)
        Case CTE_TABLAS
            'es 0, tablas
            'no ha ganado nadie
             tablas_3r = tablas_3r + 1
             peso_agente_ce0(primero) = peso_agente_ce0(primero) + f_incremento_de_peso(CTE_EMPATAR, 1)
             peso_agente_ce0(segundo) = peso_agente_ce0(segundo) + f_incremento_de_peso(CTE_EMPATAR, 2)
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: resultado imposible"
    End Select

    's_compartir_conocimiento_3r    : modificar algunos vecinos, pero yo no
    's_modificar_pesos_reglas       : modificar a mi mismo

    'Si elegimos compartir el conocimiento, se
    'incrementan los pesos de las reglas que ha usado el ganador
    '(y se decrementan los de las que ha usado el perdedor)
    'de unos cuantos agentes (tal vez todos)
    'cercanos al actual, pero no las del propio ganador o perdedor
    If compartir_conocimiento_3r Then
        If resultado = CTE_GANA_EL_PRIMERO Then
            'Ha ganado el primero (que juega con O)
            s_compartir_conocimiento_3r primero, segundo, True
        Else
            If resultado = CTE_GANA_EL_SEGUNDO Then
            'Ha ganado el segundo
            s_compartir_conocimiento_3r primero, segundo, False
            End If
        End If
    End If
    
    'Modificamos los pesos de las reglas del propio agente que ha jugado
    If modificar_pesos_3r Then
        If resultado = CTE_GANA_EL_PRIMERO Then
            s_modificar_pesos_reglas primero, segundo, True
        Else
            If resultado = CTE_GANA_EL_SEGUNDO Then
            'Ha ganado el segundo
                s_modificar_pesos_reglas primero, segundo, False
            End If
        End If
    End If



Next i


'Esto no se usa, pero bueno, ya que esta...
If ganadasO_3r > ganadasX_3r Then
    f_funcion_evaluacion_agentes_3r = CTE_GANA_EL_PRIMERO '1
Else
    If ganadasX_3r > ganadasO_3r Then
        f_funcion_evaluacion_agentes_3r = CTE_GANA_EL_SEGUNDO '2
    Else
        f_funcion_evaluacion_agentes_3r = CTE_TABLAS '0
    End If
End If
    
End Function
Sub s_compartir_conocimiento_3r(primero As Integer, segundo As Integer, ha_ganado_el_primero As Boolean)
           
Dim i As Integer
Dim cont_agente As Integer
Dim cont_regla As Integer

Dim izquierda As Integer
Dim derecha As Integer
Dim mi_vecino As Integer


frm_c3_in3r.coco.Visible = True
frm_c3_in3r.coco2.Visible = True
frm_c3_in3r.cc.Caption = "Compartiendo Conocimiento"
frm_c3_in3r.cc.Visible = True

'aumento el peso de las reglas de las entidades cercanas
'que coinciden con las reglas usadas por el que gana
'pero las propias reglas del que gana no las modifico
'ya que eso se hace en s_modificar_pesos_reglas
 
 
'En cuanto al primero (puede ser ganador o perdedor)
 For i = 1 To numero_de_reglas_usadas_por_el_primero_3r
    frm_c3_in3r.coco = reglas_usadas_por_el_primero_3r(i)
    rcompartidas_3r = rcompartidas_3r + 1
    izquierda = primero - (num_vecinos_3r / 2)
    derecha = primero + (num_vecinos_3r / 2)
    For cont_agente = izquierda To derecha
        DoEvents
        mi_vecino = cont_agente
        'Hago un circulo y los primeros son vecinos de los ultimos
        If mi_vecino < 1 Then
            mi_vecino = numero_total_de_agentes_ejv - mi_vecino
        End If
        If mi_vecino > numero_total_de_agentes_ejv Then
            mi_vecino = mi_vecino - numero_total_de_agentes_ejv
        End If
        frm_c3_in3r.coco2 = mi_vecino
        DoEvents
        If mi_vecino <> primero Then
            For cont_regla = 1 To numero_de_reglas_por_agente_3r
                DoEvents
                If f_tomar_regla_de_agente_3r(mi_vecino, cont_regla) = f_tomar_regla_de_agente_3r(primero, reglas_usadas_por_el_primero_3r(i)) Then
                    If ha_ganado_el_primero Then
                        s_peso_regla_aumentar mi_vecino, cont_regla
                    Else
                        s_peso_regla_disminuir mi_vecino, cont_regla
                    End If
                End If
            Next cont_regla
        End If
    Next cont_agente
 Next i
 
 For i = 1 To numero_de_reglas_usadas_por_el_segundo_3r
    frm_c3_in3r.coco = reglas_usadas_por_el_segundo_3r(i)
    rcompartidas_3r = rcompartidas_3r + 1
    izquierda = segundo - (num_vecinos_3r / 2)
    derecha = segundo + (num_vecinos_3r / 2)
    For cont_agente = izquierda To derecha
        DoEvents
        mi_vecino = cont_agente
        'Hago un circulo y los primeros son vecinos de los ultimos
        If mi_vecino < 1 Then
            mi_vecino = numero_total_de_agentes_ejv - mi_vecino
        End If
        If mi_vecino > numero_total_de_agentes_ejv Then
            mi_vecino = mi_vecino - numero_total_de_agentes_ejv
        End If
        frm_c3_in3r.coco2 = mi_vecino
        DoEvents
        If mi_vecino <> segundo Then
            For cont_regla = 1 To numero_de_reglas_por_agente_3r
                DoEvents
                If f_tomar_regla_de_agente_3r(mi_vecino, cont_regla) = f_tomar_regla_de_agente_3r(segundo, reglas_usadas_por_el_segundo_3r(i)) Then
                    If Not ha_ganado_el_primero Then
                        s_peso_regla_aumentar mi_vecino, cont_regla
                    Else
                        s_peso_regla_disminuir mi_vecino, cont_regla
                    End If
                End If
            Next cont_regla
        End If
    Next cont_agente
 Next i

frm_c3_in3r.coco.Visible = False
frm_c3_in3r.coco2.Visible = False
frm_c3_in3r.cc.Visible = False


End Sub
Sub s_modificar_pesos_reglas(primero As Integer, segundo As Integer, ha_ganado_el_primero As Boolean)
   
Dim i As Integer
  
   

'Después de jugar, las reglas usadas por el ganador mejoran su peso
  
frm_c3_in3r.coco.Visible = True
frm_c3_in3r.cc.Caption = "Modificando Pesos R."
frm_c3_in3r.cc.Visible = True
  
  
If ha_ganado_el_primero Then
      For i = 1 To numero_de_reglas_usadas_por_el_primero_3r
         frm_c3_in3r.coco = reglas_usadas_por_el_primero_3r(i)
         s_peso_regla_aumentar primero, reglas_usadas_por_el_primero_3r(i)
      Next i
      For i = 1 To numero_de_reglas_usadas_por_el_segundo_3r
         frm_c3_in3r.coco = reglas_usadas_por_el_segundo_3r(i)
         s_peso_regla_disminuir segundo, reglas_usadas_por_el_segundo_3r(i)
      Next i
 Else
      For i = 1 To numero_de_reglas_usadas_por_el_segundo_3r
         frm_c3_in3r.coco = reglas_usadas_por_el_segundo_3r(i)
         s_peso_regla_aumentar segundo, reglas_usadas_por_el_segundo_3r(i)
      Next i
      For i = 1 To numero_de_reglas_usadas_por_el_primero_3r
         frm_c3_in3r.coco = reglas_usadas_por_el_primero_3r(i)
         s_peso_regla_disminuir primero, reglas_usadas_por_el_primero_3r(i)
      Next i
 End If


frm_c3_in3r.coco.Visible = False
frm_c3_in3r.cc.Visible = False


End Sub
Sub s_ordenar_agentes_3r()

    Select Case algoritmo_ordenacion_ejv
        Case CTE_BURBUJA
            s_ordenar_agentes_bur_3r
        Case CTE_QUICKSORT
            s_ordenar_agentes_qui_3r
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Algoritmo de ordenación inexistente"
    End Select
    
End Sub
Sub s_ordenar_agentes_bur_3r()

    Dim I_n As Integer
    Dim I_x As Integer
    Dim I_i As Integer
    Dim primero As Integer
    Dim ultimo As Integer
    Dim Temp_str As String
    Dim temp_int As Integer
    Dim temp_long As Long

    ' de > a <
    primero = 1
    ultimo = numero_total_de_agentes_ejv
    'Comparo cada elemento (todos menos el último)....
    For I_n = primero To ultimo - 1
        frm_c3_in3r.entidad = I_n
        DoEvents
       'con el siguiente y todos los demás hasta el último
        For I_x = I_n + 1 To ultimo
            DoEvents
            If peso_agente_ce0(I_x) > peso_agente_ce0(I_n) Then
                'Cambio todos sus elementos
                Temp_str = agente_3r(I_x)
                agente_3r(I_x) = agente_3r(I_n)
                agente_3r(I_n) = Temp_str
                'Cambio su peso
                temp_long = peso_agente_ce0(I_x)
                peso_agente_ce0(I_x) = peso_agente_ce0(I_n)
                peso_agente_ce0(I_n) = temp_long
                'Cambio su ciclo de nacimiento
                temp_long = ciclo_nacimiento_agente_3r(I_x)
                ciclo_nacimiento_agente_3r(I_x) = ciclo_nacimiento_agente_3r(I_n)
                ciclo_nacimiento_agente_3r(I_n) = temp_long
                'Cambio los pesos de sus reglas
                For I_i = 1 To numero_de_reglas_por_agente_3r
                    DoEvents
                    temp_long = peso_regla_agente_3r(I_x, I_i)
                    peso_regla_agente_3r(I_x, I_i) = peso_regla_agente_3r(I_n, I_i)
                    peso_regla_agente_3r(I_n, I_i) = temp_long
                Next I_i
                'Cambio las prioridades de sus reglas
                For I_i = 1 To numero_de_reglas_por_agente_3r
                    DoEvents
                    temp_int = prioridad_regla_agente_3r(I_x, I_i)
                    prioridad_regla_agente_3r(I_x, I_i) = prioridad_regla_agente_3r(I_n, I_i)
                    prioridad_regla_agente_3r(I_n, I_i) = temp_int
                Next I_i
            End If
        Next I_x
    Next I_n

End Sub
Sub s_ordenar_agentes_qui_3r()

End Sub

Function f_ganador_3r() As String

'Devuelve X si la partida ha acabado y ha ganado X
'Devuelve O si la partida ha acabado y ha ganado O
'Devuelve T si la partida ha acabado y no ha ganado nadie
'Devuelve N si la partida no ha acabado todavía


Dim N As Integer
Dim devolver As String * 1

devolver = "N"


'filas
For N = 0 To 6 Step 3
    If (estado_del_tablero_3r(N + 1) = "X" And estado_del_tablero_3r(N + 2) = "X" And estado_del_tablero_3r(N + 3) = "X") Then
         devolver = "X"
    End If
    If (estado_del_tablero_3r(N + 1) = "O" And estado_del_tablero_3r(N + 2) = "O" And estado_del_tablero_3r(N + 3) = "O") Then
         devolver = "O"
    End If
Next N

' COLUMNAS
For N = 1 To 3
If (estado_del_tablero_3r(N) = "X" And estado_del_tablero_3r(N + 3) = "X" And estado_del_tablero_3r(N + 6) = "X") Then
     devolver = "X"
End If
If (estado_del_tablero_3r(N) = "O" And estado_del_tablero_3r(N + 3) = "O" And estado_del_tablero_3r(N + 6) = "O") Then
     devolver = "O"
End If
Next N

' DIAGONALES
If (estado_del_tablero_3r(1) = "X" And estado_del_tablero_3r(5) = "X" And estado_del_tablero_3r(9) = "X") Then
     devolver = "X"
End If
If (estado_del_tablero_3r(1) = "O" And estado_del_tablero_3r(5) = "O" And estado_del_tablero_3r(9) = "O") Then
     devolver = "O"
End If

If (estado_del_tablero_3r(3) = "X" And estado_del_tablero_3r(5) = "X" And estado_del_tablero_3r(7) = "X") Then
     devolver = "X"
End If

If (estado_del_tablero_3r(3) = "O" And estado_del_tablero_3r(5) = "O" And estado_del_tablero_3r(7) = "O") Then
     devolver = "O"
End If


If devolver = "N" Then
    'tablero lleno
    If numero_casillas_libres_3r < 1 Then
        devolver = "T"
    Else
    ' y si no, es N
        devolver = "N"
    End If
End If


f_ganador_3r = devolver

End Function
Sub s_es_el_turno_de(primero As Integer, segundo As Integer)

    Dim pos As Integer

    'el primer jugador juega con O          (numero_de_agente)
    'el segundo jugador juega con X         (numero_de_agente + 1)
    'el jugador que cominenza se elige al azar

    pos = f_elegir_posicion_3r(primero, segundo)
    s_colocar_ficha_3r (pos)
    DoEvents

End Sub
Sub s_colocar_ficha_3r(pos As Integer)

    Dim casilla_a_borrar As Integer

    estado_del_tablero_3r(pos) = turno_3r
    If ver_agentes_3r Or estado_3r = CTE_JUGANDO Then
        frm_c3_juego3r.B(pos).Caption = turno_3r
    End If

    'quito la casilla recien ocupada de la lista de casillas libres
    If numero_casillas_libres_3r > 0 Then
        'pongo la ultima en la que se borra
        'busco la que hay que borrar
        casilla_a_borrar = f_busca_elemento_array_integer(pos, lista_de_casillas_libres_3r(), 1, numero_casillas_libres_3r)
        If casilla_a_borrar > 0 Then
            lista_de_casillas_libres_3r(casilla_a_borrar) = lista_de_casillas_libres_3r(numero_casillas_libres_3r)
        Else
            Beep
            s_error_ejv CON_OPCION_FINALIZAR, "Error de programacion: se esta intentando ejecutar una regla no valida"
        End If
        numero_casillas_libres_3r = numero_casillas_libres_3r - 1
    End If



End Sub
Function f_elegir_posicion_3r(primero As Integer, segundo As Integer) As Integer


Dim agente_tratando As Integer
Dim posicion As String

'el primer (en cuanto a numero de agente) jugador juega con O          (numero_de_agente)
'el segundo jugador juega con X         (numero_de_agente + 1)
'el jugador que comienza se elige al azar
If turno_3r = "O" Then
    agente_tratando = primero
Else
    agente_tratando = segundo
End If

'busco en las partes izquierdas de las reglas
'una descripcion del tablero que coincida con la actual

Dim i As Integer
Dim regla As String * 10
Dim agente As String
Dim mayor_peso As Long
Dim mayor_prioridad As Integer
Dim mejor_regla As Integer
Dim indice_regla_actual As Integer
Dim cont_seleccionadas As Integer

numero_de_posibles_reglas_a_usar_3r = 0
For i = 1 To numero_de_reglas_por_agente_3r
    DoEvents
    regla = f_tomar_regla_de_agente_3r(agente_tratando, i)
    If f_coincide_el_estado_del_tablero_con_la_regla(regla) Then
        'La añado en la lista de posibles reglas a usar
        numero_de_posibles_reglas_a_usar_3r = numero_de_posibles_reglas_a_usar_3r + 1
        ReDim Preserve posibles_reglas_a_usar_3r(1 To numero_de_posibles_reglas_a_usar_3r) As Integer
        posibles_reglas_a_usar_3r(numero_de_posibles_reglas_a_usar_3r) = i
    End If
Next i

'Se aplica, del conjunto de todas las reglas, aquella que:
'1.- corresponde con el estado del tablero
'2.- si no hay ninguna, se juega al azar
'3.- si hay mas de una, la que tiene mayor prioridad
'4.- si hay mas de una, la que tiene mayor peso
'5.- si hay mas de una, se juega con una al azar entre esas


'Puede que no haya ninguna regla que se corresponda con el estado del
'tablero y en ese caso jugamos al azar y no se asigna peso a las reglas
'ya que no se usan
If numero_de_posibles_reglas_a_usar_3r < 1 Then
   '=======================================================================
   ' Juego al azar
   '=======================================================================
    jugadasazar_3r = jugadasazar_3r + 1
    'muestro la regla que se esta usando(ninguna: azar)
    If estado_3r = CTE_JUGANDO Then
        frm_c3_juego3r.peso = "."
        frm_c3_juego3r.prioridad = "."
        frm_c3_juego3r.Label3(0) = "."
        For i = 1 To 9
            frm_c3_juego3r.Label3(i) = "."
        Next i
    End If
    f_elegir_posicion_3r = lista_de_casillas_libres_3r(fi_azar1(numero_casillas_libres_3r))
Else
   '=======================================================================
   ' Juego con una regla
   '=======================================================================
    jugadasreglas_3r = jugadasreglas_3r + 1
   '=======================================================================
   'Primer filtro: las de mayor prioridad
   'Primero calculo cual es la mayor prioridad
    mayor_prioridad = -1
    For i = 1 To numero_de_posibles_reglas_a_usar_3r
        indice_regla_actual = posibles_reglas_a_usar_3r(i)
        If prioridad_regla_agente_3r(agente_tratando, indice_regla_actual) > mayor_prioridad Then
            mayor_prioridad = prioridad_regla_agente_3r(agente_tratando, indice_regla_actual)
        End If
    Next i
   
   'De todas las reglas posibles a usar, selecciono las de mas prioridad
    cont_seleccionadas = 0
    For i = 1 To numero_de_posibles_reglas_a_usar_3r
       indice_regla_actual = posibles_reglas_a_usar_3r(i)
       If prioridad_regla_agente_3r(agente_tratando, indice_regla_actual) = mayor_prioridad Then
           cont_seleccionadas = cont_seleccionadas + 1
           posibles_reglas_a_usar_3r(cont_seleccionadas) = indice_regla_actual
       End If
    Next i
    numero_de_posibles_reglas_a_usar_3r = cont_seleccionadas
    '=======================================================================
    'Segundo filtro: De todas las reglas que quedan, cojo las de mas peso
    'Primero calculo cual es el mayor peso
     mayor_peso = -1
     For i = 1 To numero_de_posibles_reglas_a_usar_3r
         indice_regla_actual = posibles_reglas_a_usar_3r(i)
         If peso_regla_agente_3r(agente_tratando, indice_regla_actual) > mayor_peso Then
             mayor_peso = peso_regla_agente_3r(agente_tratando, indice_regla_actual)
         End If
     Next i
     
     If numero_de_posibles_reglas_a_usar_3r > 1 Then
        cont_seleccionadas = 0
        For i = 1 To numero_de_posibles_reglas_a_usar_3r
           indice_regla_actual = posibles_reglas_a_usar_3r(i)
           If peso_regla_agente_3r(agente_tratando, indice_regla_actual) = mayor_peso Then
               cont_seleccionadas = cont_seleccionadas + 1
               posibles_reglas_a_usar_3r(cont_seleccionadas) = indice_regla_actual
           End If
        Next i
        numero_de_posibles_reglas_a_usar_3r = cont_seleccionadas
     End If
    '=======================================================================
    'Tercer filtro: De todas las reglas que quedan, cojo al azar
    mejor_regla = posibles_reglas_a_usar_3r(fi_azar1(numero_de_posibles_reglas_a_usar_3r))
    '=======================================================================
     
     
     If estado_3r = CTE_FUNCIONANDO Then
        'La añado en la lista de reglas usadas - solo durante la evolución, durante el
        'juego con el hombre no
        If turno_3r = "O" Then
            numero_de_reglas_usadas_por_el_primero_3r = numero_de_reglas_usadas_por_el_primero_3r + 1
            If numero_de_reglas_usadas_por_el_primero_3r = 1 Then
                ReDim reglas_usadas_por_el_primero_3r(1 To 1) As Integer
            Else
                ReDim Preserve reglas_usadas_por_el_primero_3r(1 To numero_de_reglas_usadas_por_el_primero_3r) As Integer
            End If
            reglas_usadas_por_el_primero_3r(numero_de_reglas_usadas_por_el_primero_3r) = mejor_regla
        Else
            numero_de_reglas_usadas_por_el_segundo_3r = numero_de_reglas_usadas_por_el_segundo_3r + 1
            If numero_de_reglas_usadas_por_el_segundo_3r = 1 Then
                ReDim reglas_usadas_por_el_segundo_3r(1 To 1) As Integer
            Else
                ReDim Preserve reglas_usadas_por_el_segundo_3r(1 To numero_de_reglas_usadas_por_el_segundo_3r) As Integer
            End If
            reglas_usadas_por_el_segundo_3r(numero_de_reglas_usadas_por_el_segundo_3r) = mejor_regla
        End If
     End If
    'Aplico la regla, es decir, cojo la parte derecha
    regla = f_tomar_regla_de_agente_3r(agente_tratando, mejor_regla)
    posicion = Right(regla, 1)
    'muestro la regla que se esta usando si juego contra el hombre
    If estado_3r = CTE_JUGANDO Then
        frm_c3_juego3r.peso = peso_regla_agente_3r(agente_tratando, mejor_regla)
        frm_c3_juego3r.prioridad = prioridad_regla_agente_3r(agente_tratando, mejor_regla)
        frm_c3_juego3r.Label3(0) = posicion
        For i = 1 To 9
            frm_c3_juego3r.Label3(i) = Mid(regla, i, 1)
        Next i
    End If
    
    f_elegir_posicion_3r = posicion
    
End If

End Function
Function f_coincide_el_estado_del_tablero_con_la_regla(regla As String) As Boolean

    Dim i As Integer
    Dim elemento_regla As String * 1
    Dim elemento_tablero As String * 1
    Dim devolver As Boolean
    
    devolver = True
    For i = 1 To numero_de_caracteres_parte_izq_regla_3r
        elemento_regla = Mid(regla, i, 1)
        elemento_tablero = estado_del_tablero_3r(i)
        Select Case elemento_regla
            Case "*" 'este da igual, es el indiferente
            Case "V"
                If elemento_tablero = "X" Or elemento_tablero = "O" Then
                    devolver = False
                    Exit For
                End If
            Case "P"
                If elemento_tablero = "V" Then
                    devolver = False
                    Exit For
                Else
                    If turno_3r = "X" Then
                        If elemento_tablero = "O" Then
                            devolver = False
                            Exit For
                        End If
                    End If
                    If turno_3r = "O" Then
                        If elemento_tablero = "X" Then
                            devolver = False
                            Exit For
                        End If
                    End If
                End If
            Case "C"
                If elemento_tablero = "V" Then
                    devolver = False
                    Exit For
                Else
                    If turno_3r = "X" Then
                        If elemento_tablero = "X" Then
                            devolver = False
                            Exit For
                        End If
                    End If
                    If turno_3r = "O" Then
                        If elemento_tablero = "O" Then
                            devolver = False
                            Exit For
                        End If
                    End If
                End If
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: elemento regla inexistente"
        End Select
    Next i


    f_coincide_el_estado_del_tablero_con_la_regla = devolver

End Function
Sub s_es_el_turno_de_O()


    Dim pos As Integer

    'jugar contra el ordenador es jugar contra el agente numero 1
    'agente_actual_ejv = 1
    agente_actual_ejv = frm_c3_juego3r.Text1.Text
    pos = f_elegir_posicion_3r(agente_actual_ejv, 2)
    s_colocar_ficha_3r (pos)
    
    'despues de cada movimiento se ha de comprobar si se ha acabado
    ganador_3r = "N"
    ganador_3r = f_ganador_3r

    'Las tres formas de finalizar
    If ganador_3r = "X" Or ganador_3r = "O" Then
        frm_c3_juego3r.mensaje.ForeColor = &HFF& 'rojo
        frm_c3_juego3r.mensaje.Caption = "El ganador es " & ganador_3r
    Else
        If ganador_3r = "T" Then
            'ganador es "T"
            frm_c3_juego3r.mensaje.Caption = "Tablas. No hay ganador"
        End If
    End If
    
    'cambio de turno
    turno_3r = "X"
    frm_c3_juego3r.turno.Caption = "TURNO: " & turno_3r



End Sub
Sub s_jugar_contra_ordenador_3r()

Dim i As Integer
Dim azar As Integer

'Elegir un jugador de comienzo al azar
azar = fi_azar1(2)
If azar = 1 Then
    turno_3r = "X"
Else
    turno_3r = "O"
End If

frm_c3_juego3r.empieza.Caption = "El jugador que comienza esta partida es " & turno_3r
frm_c3_juego3r.turno.Caption = "TURNO: " & turno_3r

If turno_3r = "O" Then
    s_es_el_turno_de_O
End If


End Sub
Sub s_ajustar_pesos_agentes_3r()

    Dim cont_agente As Integer
    
    'La de mas peso no se ajusta, ajustar siempre es disminuir
    'y se tienen en cuenta las anteriores a una dada
    For cont_agente = 2 To numero_total_de_agentes_ejv
        frm_c3_in3r.entidad = cont_agente

        DoEvents
        peso_agente_ce0(cont_agente) = f_funcion_ajuste_pesos_agentes_3r(cont_agente)
    Next cont_agente


End Sub
Sub s_reproducir_agentes_3r()

    Select Case tipo_reproduccion_3r
        Case 1
            'la última mitad de las ordenadas se eliminan
            'Reproduzco la primera mitad de las entidades haciendo que la primera
            'mitad borre a la segunda, combinandose las entidades de la primera
            'mitad por parejas, y intercambiando reglas
            'También hago una mutación cada cierto tiempo
            s_reproduccion_agentes_3r_caso ("A")
        Case 2
            'Se selecciona el 10% primero, se borra y rellena el resto
            'Cada entidad genera otras 9
            'cada pareja genera 9 parejas
            'Cada 2 entidades generan 18
            'También hago una mutación cada cierto tiempo
            s_reproduccion_agentes_3r_caso ("B")
        Case 3
            s_reproduccion_agentes_3r_caso ("C")
        Case 4
            s_reproduccion_agentes_3r_caso ("D")
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: caso de reproducción inexistente"
    End Select

End Sub
Function f_funcion_ajuste_pesos_agentes_3r(numero_de_agente_actual As Integer) As Integer


    Dim cont_agente As Integer
    Dim cont_regla As Integer
    Dim peso As Long
    
    Dim agente As String
    Dim agente_actual As String
    
    Dim regla  As String * 10
    Dim regla_del_agente_actual  As String * 10
    
    Dim numero_de_reglas_distintas_agente_mas_parecido As Integer
    Dim numero_de_reglas_iguales_agente_mas_parecido As Integer
    Dim agente_mas_parecido As Integer
    Dim numero_de_reglas_iguales_curso As Integer

    Dim primero_a_analizar As Integer

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
        
    'La primera entidad no la trato pq estan ordenadas
    'Trato todas desde la primera hasta una anterior a la actual
    'y veo si se parecen y si hay alguna muy parecida
    'disminuyo el peso a la actual
    'si el peso de la
    
    peso = peso_agente_ce0(numero_de_agente_actual)
    agente_mas_parecido = 0
    numero_de_reglas_iguales_agente_mas_parecido = 0
    
    'Como mucho analizo los Numero_cercanos_relativo_3r jugadores mas cercanos
    'y con un peso mayor que el actual
    If numero_de_agente_actual > Numero_cercanos_relativo_3r Then
        primero_a_analizar = numero_de_agente_actual - Numero_cercanos_relativo_3r
    Else
        primero_a_analizar = 1
    End If
    For cont_agente = primero_a_analizar To numero_de_agente_actual - 1
        numero_de_reglas_iguales_curso = 0
        For cont_regla = 1 To numero_de_reglas_por_agente_3r
            DoEvents
            'regla del agente del bucle
            regla = f_tomar_regla_de_agente_3r(cont_agente, cont_regla)
            'regla del agente actual
            regla_del_agente_actual = f_tomar_regla_de_agente_3r(numero_de_agente_actual, cont_regla)
            'miramos si son iguales
            If regla = regla_del_agente_actual Then
                numero_de_reglas_iguales_curso = numero_de_reglas_iguales_curso + 1
            End If
        Next cont_regla
        If numero_de_reglas_iguales_curso > numero_de_reglas_iguales_agente_mas_parecido Then
            'Si la agente que me hace la competencia, la agente mas parecido a la actual, tiene
            'menor o igual peso que la actual, entonces no se aplica la disminución de peso a
            'la actual
            'esto es pq aunque al principio están desordenadas, al ir disminuyendo
            'algunos pesos, algunas por delante de una se quedan con poco peso
            If peso_agente_ce0(cont_agente) > peso_agente_ce0(numero_de_agente_actual) Then
                numero_de_reglas_iguales_agente_mas_parecido = numero_de_reglas_iguales_curso
                agente_mas_parecido = cont_agente
            End If
        End If
    Next cont_agente
    
    If numero_de_reglas_iguales_agente_mas_parecido = 0 Then
        'En ningun agente hay ninguna regla igual a una del actual
        peso = peso
    Else
        If numero_de_reglas_iguales_agente_mas_parecido = numero_de_reglas_por_agente_3r Then
            'Hay una entidad que es exactamente igual
            peso = 0
        Else
            numero_de_reglas_distintas_agente_mas_parecido = numero_de_reglas_por_agente_3r - numero_de_reglas_iguales_agente_mas_parecido
            peso = peso - (peso / (2 ^ numero_de_reglas_distintas_agente_mas_parecido))
        End If
    End If
    
    f_funcion_ajuste_pesos_agentes_3r = peso
    


End Function
Sub s_quitar_reglas_repetidas()

    'no tiene en cuenta la prioridad de la regla
    'dos reglas iguales excepto por su prioridad se consideran iguales

    Dim cont_agente As Integer
    Dim cont_regla As Integer
    Dim anterior As Integer
    Dim agente_nuevo As String
    Dim ha_habido_repetida As Boolean
    
    'para cada agente
    For cont_agente = 1 To numero_total_de_agentes_ejv
        frm_c3_in3r.entidad = cont_agente

        agente_nuevo = ""
        agente_nuevo = f_tomar_regla_de_agente_3r(cont_agente, 1)
        'para cada regla de cada agente, excepto la primera
        For cont_regla = 2 To numero_de_reglas_por_agente_3r
            DoEvents
            ha_habido_repetida = False
            'la comparo con las anteriores
            For anterior = 1 To cont_regla - 1
                'si la actual es igual a alguna de las anteriores
                If f_tomar_regla_de_agente_3r(cont_agente, cont_regla) = f_tomar_regla_de_agente_3r(cont_agente, anterior) Then
                    'como peso de la regla dejo el mas alto de los dos
                    If peso_regla_agente_3r(cont_agente, cont_regla) > peso_regla_agente_3r(cont_agente, anterior) Then
                        peso_regla_agente_3r(cont_agente, anterior) = peso_regla_agente_3r(cont_agente, cont_regla)
                    End If
                    'no puede haber mas de una repetida
                    ha_habido_repetida = True
                    Exit For
                End If
            Next anterior
            If ha_habido_repetida Then
                'se modifica esa regla del agente nuevo
                agente_nuevo = agente_nuevo & f_mutar_regla(f_tomar_regla_de_agente_3r(cont_agente, cont_regla))
                peso_regla_agente_3r(cont_agente, cont_regla) = f_media_pesos_primero_3r()
                prioridad_regla_agente_3r(cont_agente, cont_regla) = f_crear_prioridad_al_azar_3r()
            Else
                'no se modifica esa regla del agente nuevo
                agente_nuevo = agente_nuevo & f_tomar_regla_de_agente_3r(cont_agente, cont_regla)
            End If
        Next cont_regla
        'grabamos el nuevo agente
        agente_3r(cont_agente) = agente_nuevo

    Next cont_agente


End Sub
Function f_jugar_una_partida_3r(primero As Integer, segundo As Integer) As Integer

    'Devuelve CTE_GANA_EL_PRIMERO 1 si ha ganado el primero, es decir O
    'Devuelve CTE_GANA_EL_SEGUNDO 2 si ha ganado el segundo, es decir X
    'Devuelve CTE_TABLAS 0 si es tablas
    
    'Ahora juegan una partida los dos agentes:
    Dim cont_agente As Integer
    Dim cont_regla As Integer
    Dim peso As Long
    Dim ganador As String * 1
    
    'el primer (orden array) jugador juega con O          (numero_de_agente)
    'el segundo jugador juega con X         (numero_de_agente + 1)
    'el jugador que cominenza se elige al azar
   
   
    'Elegir un jugador de comienzo al azar
    le_ha_tocado_empezar_a_3r = fi_azar1(2)
    If le_ha_tocado_empezar_a_3r = 1 Then
        turno_3r = "O"
    Else
        turno_3r = "X"
    End If
  
    'inicializamos el tablero
    s_inicializar_tablero_3r
    
    ganador = "N"
    While ganador = "N"
        s_es_el_turno_de primero, segundo
        ganador = f_ganador_3r()
        s_cambio_de_turno_3r
    Wend

    'el primer jugador juega con O          (numero_de_agente)
    'el segundo jugador juega con X         (numero_de_agente + 1)
    'el jugador que cominenza se elige al azar
    If ganador = "O" Then
        f_jugar_una_partida_3r = CTE_GANA_EL_PRIMERO
    Else
        If ganador = "X" Then
            f_jugar_una_partida_3r = CTE_GANA_EL_SEGUNDO
        Else
            'ganador es "T"
            f_jugar_una_partida_3r = CTE_TABLAS
        End If
    End If


End Function
Sub s_inicializar_tablero_3r()
    
    Dim i As Integer
    
    'inicializamos el tablero
    For i = 1 To numero_de_caracteres_parte_izq_regla_3r
        estado_del_tablero_3r(i) = "V"
        lista_de_casillas_libres_3r(i) = i
        If ver_agentes_3r Then
            frm_c3_juego3r.B(i).Caption = ""
        End If
    Next i
    numero_casillas_libres_3r = numero_de_caracteres_parte_izq_regla_3r

End Sub
Function f_incremento_de_peso(ha_ganado As Integer, numero_de_orden As Integer) As Long

'numero_de_orden es el numero de la entidad que estamos evaluando
'en el orden en el que estan en el array

Dim he_empezado_yo As Integer

If numero_de_orden = le_ha_tocado_empezar_a_3r Then
    he_empezado_yo = True
Else
    he_empezado_yo = False
End If

Select Case Metodo_de_asignar_pesos_3r
    Case "A"
        Select Case ha_ganado
            Case CTE_GANAR
                f_incremento_de_peso = 1
            Case CTE_PERDER
                f_incremento_de_peso = -1
            Case CTE_EMPATAR
                f_incremento_de_peso = 0
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: ha ganado inexistente"
        End Select
    Case "B"
        Select Case ha_ganado
            Case CTE_GANAR
                Select Case he_empezado_yo
                    Case True
                        f_incremento_de_peso = 2
                    Case False
                        f_incremento_de_peso = 4
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: ha ganado inexistente"
                End Select
            Case CTE_PERDER
                Select Case he_empezado_yo
                    Case True
                        f_incremento_de_peso = -4
                    Case False
                        f_incremento_de_peso = -2
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: ha empezado inexistente"
                End Select
            Case CTE_EMPATAR
                Select Case he_empezado_yo
                    Case True
                        f_incremento_de_peso = -1
                    Case False
                        f_incremento_de_peso = 1
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: ha empezado inexistente"
                End Select
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: ha ganado inexistente"
        End Select
    Case "C"
        Select Case ha_ganado
            Case CTE_GANAR
                Select Case he_empezado_yo
                    Case True
                        f_incremento_de_peso = 2
                    Case False
                        f_incremento_de_peso = 4
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: ha empezado inexistente"
                End Select
            Case CTE_PERDER
                Select Case he_empezado_yo
                    Case True
                        f_incremento_de_peso = -4
                    Case False
                        f_incremento_de_peso = -2
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: ha empezado inexistente"
                End Select
            Case CTE_EMPATAR
                Select Case he_empezado_yo
                    Case True
                        f_incremento_de_peso = -1
                    Case False
                        f_incremento_de_peso = 1
                    Case Else
                        s_error_ejv CON_OPCION_FINALIZAR, "Error: ha empezado inexistente"
                End Select
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: ha ganado inexistente"
        End Select
        If numero_de_orden = 1 Then
            f_incremento_de_peso = f_incremento_de_peso * numero_de_reglas_usadas_por_el_primero_3r
        Else
            If numero_de_orden = 2 Then
                f_incremento_de_peso = f_incremento_de_peso * numero_de_reglas_usadas_por_el_segundo_3r
            Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: num orden inexistente"
            End If
        End If
        
    Case Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: método asignar pesos inexistente"
End Select


End Function
Sub s_cambio_de_turno_3r()
    
    If turno_3r = "X" Then
        turno_3r = "O"
    Else
        turno_3r = "X"
    End If

End Sub
Function f_hay_mutacion_regla_3r() As Boolean
    
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

    If tasa_de_mutacion_3r < 0 Then
        'Tasas de mutación negativas son sin mutación
         devolver = False
    Else
        If Not (mutaciones_acumuladas_3r) Then
            'mutación normal
             If tasa_de_mutacion_3r = 0 Or tasa_de_mutacion_3r = 1 Then
                 'Tasas de mutación 0 o 1 es siempre mutación
                  devolver = True
             Else
                 'Posible mutación
                 If fi_azar1(tasa_de_mutacion_3r) = 1 Then
                    'hay mutación
                     devolver = True
                 End If
            End If
       Else
            'mutación acumumulada
            'Si es la primera y antes no ha habido, guardamos el valor de la vieja
            If Not ha_habido_mutacion_anterior_3r Then
                tasa_de_mutacion_vieja_3r = tasa_de_mutacion_3r
            End If
             If tasa_de_mutacion_3r = 0 Or tasa_de_mutacion_3r = 1 Then
                 'Tasas de mutación 0 es siempre mutación
                  devolver = True
             Else
                 'Posible mutación
                 If fi_azar1(tasa_de_mutacion_3r) = 1 Then
                    'hay mutación
                      devolver = True
                      ha_habido_mutacion_anterior_3r = True
                      'hacemos que la proxima sea el doble mas probable, hasta 2 limite
                        temp = Int(tasa_de_mutacion_3r / 2) + 1
                        If temp >= 2 Then
                              tasa_de_mutacion_3r = temp
                        Else
                              tasa_de_mutacion_3r = 2
                        End If
                 Else
                      devolver = False
                      ha_habido_mutacion_anterior_3r = False
                      'recuperamos la tasa vieja
                      tasa_de_mutacion_3r = tasa_de_mutacion_vieja_3r
                 End If
            End If

       End If
    End If

    f_hay_mutacion_regla_3r = devolver
    
End Function
Function f_son_padres_iguales_3r(padre1, padre2) As Boolean


Dim i As Integer
Dim son_iguales As Boolean

son_iguales = True

If agente_3r(padre1) <> agente_3r(padre2) Then
    son_iguales = False
End If

f_son_padres_iguales_3r = son_iguales

    

End Function
Sub s_reproduccion_agentes_3r_caso(caso)


 Dim I_Indice1 As Integer
 Dim I_Indice2 As Integer
 
 Dim cont_regla As Integer
 Dim i As Integer
 Dim I_i As Integer
 Dim generados As Integer
     
 Dim frontera As Integer
 Dim destino As Integer
 Dim ultimo As Integer
 
 Dim alterno As Boolean
 Dim regla_a_heredar As String
 
 Dim padre As Integer
 Dim madre As Integer
 Dim hijo As Integer

 Dim pasar_al_siguiente As Boolean

 I_Indice1 = 1
 alterno = True
 
 'generados es el mumero de agentes generados por cada 2 agentes
'frontera es el ultimo de los progenitores
 Select Case caso
     Case "A"
        '50-50 es generar dos por cada dos
         generados = 2
         frontera = Int(numero_total_de_agentes_ejv / 2)
         I_Indice2 = frontera + 1
     Case "B"
        '10-90 es generar 18 por cada dos
         generados = 18
         Select Case numero_total_de_agentes_ejv
            Case 8
                '10-90 es generar 18 por cada dos
                'pero en este caso generamos 6 por cada dos
                'que es lo que mas se parece
                 generados = 6
                 'frontera es el ultimo de los progenitores
                 frontera = 2
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
                s_error_ejv CON_OPCION_FINALIZAR, "Error: numero de agentes no contemplado"
         End Select
     Case "C"
        '20-80 es generar 8 por cada dos
         generados = 8
         Select Case numero_total_de_agentes_ejv
            Case 8
                 '20-80 es generar 8 por cada dos
                'pero en este caso generamos 6 por cada dos
                'que es lo que mas se parece
                 generados = 6
                 'frontera es el ultimo de los progenitores
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
                s_error_ejv CON_OPCION_FINALIZAR, "Error: numero de agentes no contemplado"
         End Select
     Case "D"
        '40-50-10 copiar los 10 ultimos como los 40-50 y hacer el caso A
         Select Case numero_total_de_agentes_ejv
                'destino = 40%
                'frontera = 40% + 50%
            Case 8
                'algo parecido es copiar los 2 ultimos 7-8 como los 3-4
                'destino es el ultimo del 40% seleccioando superior
                'destino+1 es el lugar donde se comienza a poner al 10% ultimo
                'frontera es el ultimo de los que moriran
                'frontera mas uno es el primero de los que hay que mover
                'ultimo es el ultimo de todos
                
                'muevo los 7-8 a 3-4
                 destino = 2
                 frontera = 2 + 4
                 ultimo = numero_total_de_agentes_ejv '8
            Case 20
                destino = 8
                frontera = 8 + 10
                ultimo = 20
            Case 40
                 destino = 16
                 frontera = 16 + 20
                 ultimo = numero_total_de_agentes_ejv '40
            Case 80
                 destino = 32
                 frontera = 32 + 40
                 ultimo = numero_total_de_agentes_ejv '80
            Case 160
                 destino = 64
                 frontera = 64 + 80
                 ultimo = numero_total_de_agentes_ejv '160
            Case 320
                 destino = 128
                 frontera = 128 + 160
                 ultimo = numero_total_de_agentes_ejv '320
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: numero de agentes no contemplado"
         End Select
         I_Indice2 = frontera + 1
         'Copio el 10% final a partir del 40% y hago como en el caso A
         For i = I_Indice2 To ultimo 'esto es solo para contar
            DoEvents
             destino = destino + 1
            'copio agente
             agente_3r(destino) = agente_3r(i)
             peso_agente_ce0(destino) = peso_agente_ce0(i)
             ciclo_nacimiento_agente_3r(destino) = ciclo_ejv
            'pesos de las reglas
             For I_i = 1 To numero_de_reglas_por_agente_3r
                DoEvents
                peso_regla_agente_3r(destino, I_i) = peso_regla_agente_3r(i, I_i)
             Next I_i
            'prioridades de las reglas
             For I_i = 1 To numero_de_reglas_por_agente_3r
                DoEvents
                prioridad_regla_agente_3r(destino, I_i) = prioridad_regla_agente_3r(i, I_i)
             Next I_i
         Next i
         generados = 2
         frontera = Int(numero_total_de_agentes_ejv / 2)
    Case Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: caso de reproducción no contemplado"
 End Select
 
 I_Indice2 = frontera + 1
 


 
 'Si los padres son al azar, los desordeno primero
 'solo los padres
 If eleccion_de_padres_al_azar_3r Then
     s_desordenar_medio_array_agentes_3r frontera
 End If
 
 
 'Copio dos padres en dos hijos alternando uno si uno no
 While I_Indice1 < frontera
     DoEvents
     padre = I_Indice1 + 1
     madre = I_Indice1
     
     frm_c3_in3r.entidad = madre & "-" & padre
     
     'miro si los padres son identicos
     pasar_al_siguiente = False
     If padres_identicos_producen_mutaciones_3r Then
         If f_son_padres_iguales_3r(madre, padre) Then
             'Si los padres son identicos, y se ha elegido la opción, los hijos son todo mutaciones
             'Cada pareja de supervivientes genera generados hijos
             For i = 0 To generados - 1
                DoEvents
                 hijo = I_Indice2 + i
                 agente_3r(hijo) = ""
                 'como peso del agente pongo la media de sus padres
                 peso_agente_ce0(hijo) = (peso_agente_ce0(padre) + peso_agente_ce0(madre)) / 2
                'pongo el ciclo de nacimiento
                 ciclo_nacimiento_agente_3r(hijo) = ciclo_ejv
                 'inicializo los pesos de todas las reglas a f_media_pesos_primero_3r()
                 For cont_regla = 1 To numero_de_reglas_por_agente_3r
                     DoEvents
                     agente_3r(hijo) = agente_3r(hijo) & f_crear_regla_al_azar_3r
                     peso_regla_agente_3r(hijo, cont_regla) = f_media_pesos_primero_3r()
                     prioridad_regla_agente_3r(hijo, cont_regla) = f_crear_prioridad_al_azar_3r()
                 Next cont_regla
             Next i
             pasar_al_siguiente = True
          End If
     End If
        
     
     'Caso normal, los hijos salen de los padres
     If Not (pasar_al_siguiente) Then
         'Reproduzco normalmente
         'Cada pareja de supervivientes genera 2... hijos (caso A)
         For i = 0 To generados - 1
            DoEvents
             hijo = I_Indice2 + i
             'Cada agente tiene numero_de_reglas_por_agente_3r reglas
             agente_3r(hijo) = ""
             'como peso del agente pongo la media de sus padres
             peso_agente_ce0(hijo) = (peso_agente_ce0(madre) + peso_agente_ce0(padre)) / 2
            'pongo el ciclo de nacimiento
             ciclo_nacimiento_agente_3r(hijo) = ciclo_ejv
            'inicializo los pesos de todas las reglas a f_media_pesos_primero_3r()
             For cont_regla = 1 To numero_de_reglas_por_agente_3r
                 DoEvents
                 alterno = f_elegir_progenitor_donante_3r(alterno)
                 If alterno Then
                    'en principio se toma del padre, pero...
                    'Heredar la regla de mas peso de los dos padres en vez de tomarla al azar
                    If heredar_regla_mas_peso_3r And peso_regla_agente_3r(madre, cont_regla) > peso_regla_agente_3r(padre, cont_regla) Then
                       'si el peso de la madre es mayor
                       'lo tomo de la madre
                       regla_a_heredar = f_tomar_regla_de_agente_3r(madre, cont_regla)
                       agente_3r(hijo) = agente_3r(hijo) & f_tomar_regla_de_agente_3r(madre, cont_regla)
                       peso_regla_agente_3r(hijo, cont_regla) = peso_regla_agente_3r(madre, cont_regla)
                       prioridad_regla_agente_3r(hijo, cont_regla) = prioridad_regla_agente_3r(madre, cont_regla)
                    Else
                       'lo tomo del padre
                       regla_a_heredar = f_tomar_regla_de_agente_3r(padre, cont_regla)
                       peso_regla_agente_3r(hijo, cont_regla) = peso_regla_agente_3r(padre, cont_regla)
                       prioridad_regla_agente_3r(hijo, cont_regla) = prioridad_regla_agente_3r(padre, cont_regla)
                    End If
                     'variación-mutación
                     If f_hay_mutacion_regla_3r() Then
                         regla_a_heredar = f_mutar_regla(regla_a_heredar)
                         peso_regla_agente_3r(hijo, cont_regla) = f_media_pesos_primero_3r()
                         prioridad_regla_agente_3r(hijo, cont_regla) = f_crear_prioridad_al_azar_3r()
                     End If
                    'Lo heredo
                    agente_3r(hijo) = agente_3r(hijo) & regla_a_heredar
                  Else 'del alterno
                    'en principio se toma de la madre, pero...
                    'Heredar la regla de mas peso de los dos padres en vez de tomarla al azar
                    If heredar_regla_mas_peso_3r And peso_regla_agente_3r(padre, cont_regla) > peso_regla_agente_3r(madre, cont_regla) Then
                       'lo tomo del padre
                       regla_a_heredar = f_tomar_regla_de_agente_3r(padre, cont_regla)
                       peso_regla_agente_3r(hijo, cont_regla) = peso_regla_agente_3r(padre, cont_regla)
                       prioridad_regla_agente_3r(hijo, cont_regla) = prioridad_regla_agente_3r(padre, cont_regla)
                    Else
                       'lo tomo de la madre
                       regla_a_heredar = f_tomar_regla_de_agente_3r(madre, cont_regla)
                       peso_regla_agente_3r(hijo, cont_regla) = peso_regla_agente_3r(madre, cont_regla)
                       prioridad_regla_agente_3r(hijo, cont_regla) = prioridad_regla_agente_3r(madre, cont_regla)
                    End If
                     'variación-mutación
                     If f_hay_mutacion_regla_3r() Then
                         regla_a_heredar = f_mutar_regla(regla_a_heredar)
                         peso_regla_agente_3r(hijo, cont_regla) = f_media_pesos_primero_3r()
                         prioridad_regla_agente_3r(hijo, cont_regla) = f_crear_prioridad_al_azar_3r()
                     End If
                    'Lo heredo
                    agente_3r(hijo) = agente_3r(hijo) & regla_a_heredar
                 End If 'del alterno
             Next cont_regla
         Next i
     End If
     
     'paso a los 2 padres siguientes y al siguiente grupo de hijos
     I_Indice1 = I_Indice1 + 2
     I_Indice2 = I_Indice2 + generados
     
 Wend


End Sub
Function f_elegir_progenitor_donante_3r(alterno As Boolean) As Boolean

    If coger_alternos_3r Then
        f_elegir_progenitor_donante_3r = Not (alterno)
    Else
        If (fi_azar1(2) = 1) Then
            f_elegir_progenitor_donante_3r = True
        Else
            f_elegir_progenitor_donante_3r = False
        End If
    End If

End Function
Function f_media_pesos_primero_3r() As Long
    
Dim cont_regla As Integer
Dim total As Long

total = 0
For cont_regla = 1 To numero_de_reglas_por_agente_3r
    total = total + peso_regla_agente_3r(1, cont_regla)
Next
total = total / numero_de_reglas_por_agente_3r

f_media_pesos_primero_3r = total

End Function
Sub s_desordenar_medio_array_agentes_3r(ultimo As Integer)

'Desordena los agentes
'copiado de Si_DesordenMedioArray_2D_S agente_3r(), numero_de_reglas_por_agente_3r, 1, I_Indice2 - 1

Dim I_n As Integer
Dim Temp_str As String
Dim temp_int As Integer
Dim temp_long As Long
Dim uno As String
Dim otro As String

Dim j As Integer
Dim I_i As Integer

For I_n = 1 To 2 * ultimo
    DoEvents
   'Intercambio 2 elementos a azar
    uno = fi_azar1(ultimo)
    otro = fi_azar1(ultimo)
    
   'Muevo cada elemento
   'agentes con sus reglas
    Temp_str = agente_3r(uno)
    agente_3r(uno) = agente_3r(otro)
    agente_3r(otro) = Temp_str
    
    'pesos de los agentes
    temp_long = peso_agente_ce0(uno)
    peso_agente_ce0(uno) = peso_agente_ce0(otro)
    peso_agente_ce0(otro) = temp_long

    'Cambio su ciclo de nacimiento
    temp_long = ciclo_nacimiento_agente_3r(uno)
    ciclo_nacimiento_agente_3r(uno) = ciclo_nacimiento_agente_3r(otro)
    ciclo_nacimiento_agente_3r(otro) = temp_long
    
    'pesos de las reglas
    For I_i = 1 To numero_de_reglas_por_agente_3r
        DoEvents
        temp_long = peso_regla_agente_3r(uno, I_i)
        peso_regla_agente_3r(uno, I_i) = peso_regla_agente_3r(otro, I_i)
        peso_regla_agente_3r(otro, I_i) = temp_long
    Next I_i
    
    'prioridades  de las reglas
    For I_i = 1 To numero_de_reglas_por_agente_3r
        DoEvents
        temp_int = prioridad_regla_agente_3r(uno, I_i)
        prioridad_regla_agente_3r(uno, I_i) = prioridad_regla_agente_3r(otro, I_i)
        prioridad_regla_agente_3r(otro, I_i) = temp_int
    Next I_i
    
Next I_n



End Sub
Sub s_desordenar_agentes_3r()

'Desordena los agentes
'copiado de Si_DesordenMedioArray_2D_S agente_3r(), numero_de_reglas_por_agente_3r, 1, I_Indice2 - 1

Dim I_n As Integer
Dim Temp_str As String
Dim temp_int As Integer
Dim temp_long As Long
Dim uno As String
Dim otro As String

Dim j As Integer
Dim I_i As Integer

For I_n = 1 To 2 * numero_total_de_agentes_ejv
    
    frm_c3_in3r.entidad = I_n
    
    DoEvents
   'Intercambio 2 elementos a azar
    uno = fi_azar1(numero_total_de_agentes_ejv)
    otro = fi_azar1(numero_total_de_agentes_ejv)
   
   'Muevo cada elemento
   'agentes con sus reglas
    Temp_str = agente_3r(uno)
    agente_3r(uno) = agente_3r(otro)
    agente_3r(otro) = Temp_str
    
    'pesos de los agentes
    temp_long = peso_agente_ce0(uno)
    peso_agente_ce0(uno) = peso_agente_ce0(otro)
    peso_agente_ce0(otro) = temp_long
    
    'ciclo de nacimiento
    temp_long = ciclo_nacimiento_agente_3r(uno)
    ciclo_nacimiento_agente_3r(uno) = ciclo_nacimiento_agente_3r(otro)
    ciclo_nacimiento_agente_3r(otro) = temp_long
    
    'pesos de las reglas
    For I_i = 1 To numero_de_reglas_por_agente_3r
        DoEvents
        temp_long = peso_regla_agente_3r(uno, I_i)
        peso_regla_agente_3r(uno, I_i) = peso_regla_agente_3r(otro, I_i)
        peso_regla_agente_3r(otro, I_i) = temp_long
    Next I_i
    
    'prioridades de las reglas
    For I_i = 1 To numero_de_reglas_por_agente_3r
        DoEvents
        temp_int = prioridad_regla_agente_3r(uno, I_i)
        prioridad_regla_agente_3r(uno, I_i) = prioridad_regla_agente_3r(otro, I_i)
        prioridad_regla_agente_3r(otro, I_i) = temp_int
    Next I_i
    
Next I_n


End Sub
Function f_extraer_subcadena_ejv(cadena, posicion_relativa, long_elemento) As String

    f_extraer_subcadena_ejv = Mid(cadena, 1 + (long_elemento * (posicion_relativa - 1)), long_elemento)

End Function

Function f_tomar_regla_de_agente_3r(ByVal num_agente As Integer, ByVal num_regla As Integer) As String

    Dim agente As String
    Dim regla As String
        
    agente = agente_3r(num_agente)
    regla = Mid(agente, 1 + (numero_de_caracteres_por_regla_3r * (num_regla - 1)), numero_de_caracteres_por_regla_3r)

    f_tomar_regla_de_agente_3r = regla

End Function
Sub s_cargar_opciones_ev_3r()
    '7 Elección de Jugadores
    If eleccion_de_jugadores_al_azar_3r Then
        frm_c3_ev3r.Op_JugadoresAzar = 1
    Else
        frm_c3_ev3r.Op_JugadoresAzar = 0
    End If
    
    '9 El peso se calcula en función de si existen más entidades parecidas a la actual
    If pesos_relativos_3r Then
        frm_c3_ev3r.Op_Relativo = 1
    Else
        frm_c3_ev3r.Op_Relativo = 0
    End If
    '17 Numero de Partidas que juega cada pareja
     frm_c3_ev3r.Op_NumeroPartidas_3r = NumeroPartidas_3r
    '18 Numero de entidades cercanas con las que comparo al ajustar pesos
     frm_c3_ev3r.Op_Numero_cercanos_relativo_3r = Numero_cercanos_relativo_3r
    '19 Método de modificar los pesos despues de una partida
    If Metodo_de_asignar_pesos_3r = "A" Then
        frm_c3_ev3r.Op_Metodo_de_asignar_pesos_3r.ListIndex = 0
    ElseIf Metodo_de_asignar_pesos_3r = "B" Then
        frm_c3_ev3r.Op_Metodo_de_asignar_pesos_3r.ListIndex = 1
    ElseIf Metodo_de_asignar_pesos_3r = "C" Then
        frm_c3_ev3r.Op_Metodo_de_asignar_pesos_3r.ListIndex = 2
    ElseIf Metodo_de_asignar_pesos_3r = "D" Then
        frm_c3_ev3r.Op_Metodo_de_asignar_pesos_3r.ListIndex = 3
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: método de asignar pesos no contemplado"
    End If
    '22 Grupos que juegan partidas
    If personas_por_grupo_3r = "2 Jugadores" Then
        frm_c3_ev3r.Op_Personas_por_grupo.ListIndex = 0
    ElseIf personas_por_grupo_3r = "Todos contra todos" Then
        frm_c3_ev3r.Op_Personas_por_grupo.ListIndex = 1
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: agentes por grupo no contemplado"
    End If
    '25 Pesos_Partir_Cero
    If Pesos_Partir_Cero_3r Then
        frm_c3_ev3r.Op_Pesos_Partir_Cero = 1
    Else
        frm_c3_ev3r.Op_Pesos_Partir_Cero = 0
    End If
    

End Sub

Sub s_cargar_opciones_sel_3r()

    '4 Tipo de Selección-Reproducción
    If tipo_reproduccion_3r = 1 Then
        frm_c3_sel3r.Op_s1 = True
    Else
        If tipo_reproduccion_3r = 2 Then
            frm_c3_sel3r.Op_s2 = True
        Else
            If tipo_reproduccion_3r = 3 Then
                frm_c3_sel3r.Op_s3 = True
            Else
                frm_c3_sel3r.Op_s4 = True
            End If
        End If
    End If

End Sub

Sub s_inicializar_ejemplo_elegido_3r()


    'OPCIONES I
    'GENERALES DE EJEMPLOS DE VIDA(DISTINTAS A LAS DE POR DEFECTO)
    '2 Grabar Resumen
    un_ej_grabar_gra_ejv = True
    un_ej_fichero_gra_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_GRA), "r_3r_" & num_ej_activo_ejv & ".gra")
    un_ej_grabar_resumen_txt_ejv = False
    un_ej_fichero_resumen_txt_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_TXT), "r_3r_" & num_ej_activo_ejv & ".txt")
    un_ej_grabar_resumen_xls_ejv = False
    un_ej_fichero_resumen_xls_ejv = f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS), "r_3r_" & num_ej_activo_ejv & ".xls")
    max_guardado_ejv = 1000000
    autoguardado_ejv = 100

    'OPCIONES II
    'GENERALES DE CE
    Select Case num_ej_activo_ejv
        Case 1
            '1 Número de agentes inicial
            numero_total_de_agentes_ejv = 40
            '2 Probabilidad de Mutación
            tasa_de_mutacion_3r = 5
            '3 Número de reglas por agente
            'El numero de reglas de los agentes puede variar de unos ciclos a otros
            numero_de_reglas_por_agente_3r = 1
            numero_reglas_variable_3r = False
            var11_3r = 2
            var12_3r = 5
            var21_3r = 4
            var22_3r = 10
            var31_3r = 6
            var32_3r = 15
            var41_3r = 8
            var42_3r = 20
            var51_3r = 10
            '4 Tipo de Selección-Reproducción
            tipo_reproduccion_3r = 1
            '5 Mutaciones acumuladas
            mutaciones_acumuladas_3r = False
            '6 Elección de Padres
            eleccion_de_padres_al_azar_3r = True
            '7 Elección de Jugadores
            eleccion_de_jugadores_al_azar_3r = 1
            '8 Padres iguales producen hijos que son mutaciones en todos sus elementos.
            padres_identicos_producen_mutaciones_3r = False
            '9 El peso se calcula en función de si existen más entidades parecidas a la actual
            pesos_relativos_3r = False
            '10 Las entidades modifican sus propios pesos de reglas despues de jugar
            modificar_pesos_3r = False
            '11 La generación de reglas se hace totalmente al azar
            reglas_azar_3r = True
            '12 ver agentes
            ver_agentes_3r = True
            '13 Heredar la regla de mas peso de los dos padres en vez de tomarla al azar
            heredar_regla_mas_peso_3r = False
            '14 sustituir reglas repetidas por mutaciones
            quitar_reglas_repetidas_3r = False
            '15 sustituir cada ciclo por mutaciones las reglas con un peso menor o igual que ___
            sust1_3r = 0
            '16 sustituir cada ___ ciclos por mutaciones las reglas con un peso menor o igual que ___
            sust2_3r = 0
            sust3_3r = 0
            '17 Numero de Partidas que juega cada pareja
            NumeroPartidas_3r = 1
            '18 Numero de entidades cercanas con las que comparo al ajustar pesos
            Numero_cercanos_relativo_3r = 0
            '19 Método de modificar los pesos despues de una partida
            Metodo_de_asignar_pesos_3r = "C"
            '20 Compartir el conocimiento (modificar pesos reglas) con los num_vecinos_3r mas cercanos pero no de uno mismo
            compartir_conocimiento_3r = False
            num_vecinos_3r = 2
            '21 coger genes de padres alternos
            coger_alternos_3r = False
            '22 Forma de agrupar los agentes para jugar partidas en la evaluación
            personas_por_grupo_3r = "2 Jugadores"
            'personas_por_grupo_3r = "Todos contra todos"
            '24 Tipo_Mutacion
            Tipo_Mutacion_3r = 50
            '25 Pesos_Partir_Cero
            Pesos_Partir_Cero_3r = True

        Case 2
            '1 Número de agentes inicial
            numero_total_de_agentes_ejv = 8
            '2 Probabilidad de Mutación
            tasa_de_mutacion_3r = 5
            '3 Número de reglas por agente
            'El numero de reglas de los agentes puede variar de unos ciclos a otros
            numero_de_reglas_por_agente_3r = 130
            numero_reglas_variable_3r = False
            var11_3r = 2
            var12_3r = 5
            var21_3r = 4
            var22_3r = 10
            var31_3r = 6
            var32_3r = 15
            var41_3r = 8
            var42_3r = 20
            var51_3r = 10
            '4 Tipo de Selección-Reproducción
            tipo_reproduccion_3r = 1
            '5 Mutaciones acumuladas
            mutaciones_acumuladas_3r = False
            '6 Elección de Padres
            eleccion_de_padres_al_azar_3r = True
            '7 Elección de Jugadores
            eleccion_de_jugadores_al_azar_3r = 1
            '8 Padres iguales producen hijos que son mutaciones en todos sus elementos.
            padres_identicos_producen_mutaciones_3r = False
            '9 El peso se calcula en función de si existen más entidades parecidas a la actual
            pesos_relativos_3r = False
            '10 Las entidades modifican sus propios pesos de reglas despues de jugar
            modificar_pesos_3r = False
            '11 La generación de reglas se hace totalmente al azar
            reglas_azar_3r = True
            '12 ver agentes
            ver_agentes_3r = True
            '13 Heredar la regla de mas peso de los dos padres en vez de tomarla al azar
            heredar_regla_mas_peso_3r = False
            '14 sustituir reglas repetidas por mutaciones
            quitar_reglas_repetidas_3r = False
            '15 sustituir cada ciclo por mutaciones las reglas con un peso menor o igual que ___
            sust1_3r = 0
            '16 sustituir cada ___ ciclos por mutaciones las reglas con un peso menor o igual que ___
            sust2_3r = 0
            sust3_3r = 0
            '17 Numero de Partidas que juega cada pareja
            NumeroPartidas_3r = 4
            '18 Numero de entidades cercanas con las que comparo al ajustar pesos
            Numero_cercanos_relativo_3r = 0
            '19 Método de modificar los pesos despues de una partida
            Metodo_de_asignar_pesos_3r = "C"
            '20 Compartir el conocimiento (modificar pesos reglas) con los num_vecinos_3r mas cercanos pero no de uno mismo
            compartir_conocimiento_3r = False
            num_vecinos_3r = 2
            '21 coger genes de padres alternos
            coger_alternos_3r = False
            '22 Forma de agrupar los agentes para jugar partidas en la evaluación
            'personas_por_grupo_3r = "2 Jugadores"
            personas_por_grupo_3r = "Todos contra todos"
            '24 Tipo_Mutacion
            Tipo_Mutacion_3r = 50
            '25 Pesos_Partir_Cero
            Pesos_Partir_Cero_3r = True

            
        Case 3
            '1 Número de agentes inicial
            numero_total_de_agentes_ejv = 40
            '2 Probabilidad de Mutación
            tasa_de_mutacion_3r = 5
            '3 Número de reglas por agente
            'El numero de reglas de los agentes puede variar de unos ciclos a otros
            numero_de_reglas_por_agente_3r = 9
            numero_reglas_variable_3r = True
            var11_3r = 9
            var12_3r = 100
            var21_3r = 18
            var22_3r = 100
            var31_3r = 30
            var32_3r = 100
            var41_3r = 50
            var42_3r = 100
            var51_3r = 100
            '4 Tipo de Selección-Reproducción
            tipo_reproduccion_3r = 1
            '5 Mutaciones acumuladas
            mutaciones_acumuladas_3r = True
            '6 Elección de Padres
            eleccion_de_padres_al_azar_3r = True
            '7 Elección de Jugadores
            eleccion_de_jugadores_al_azar_3r = 1
            '8 Padres iguales producen hijos que son mutaciones en todos sus elementos.
            padres_identicos_producen_mutaciones_3r = False
            '9 El peso se calcula en función de si existen más entidades parecidas a la actual
            pesos_relativos_3r = 1
            '10 Las entidades modifican sus propios pesos de reglas despues de jugar
            modificar_pesos_3r = True
            '11 La generación de reglas se hace totalmente al azar
            reglas_azar_3r = False
            '12 ver agentes
            ver_agentes_3r = True
            '13 Heredar la regla de mas peso de los dos padres en vez de tomarla al azar
            heredar_regla_mas_peso_3r = False
            '14 sustituir reglas repetidas por mutaciones
            quitar_reglas_repetidas_3r = True
            '15 sustituir cada ciclo por mutaciones las reglas con un peso menor o igual que ___
            sust1_3r = 10
            '16 sustituir cada ___ ciclos por mutaciones las reglas con un peso menor o igual que ___
            sust2_3r = 25
            sust3_3r = 20
            '17 Numero de Partidas que juega cada pareja
            NumeroPartidas_3r = 2
            '18 Numero de entidades cercanas con las que comparo al ajustar pesos
            Numero_cercanos_relativo_3r = 6
            '19 Método de modificar los pesos despues de una partida
            Metodo_de_asignar_pesos_3r = "C"
            '20 Compartir el conocimiento (modificar pesos reglas) con los num_vecinos_3r mas cercanos pero no de uno mismo
            compartir_conocimiento_3r = True
            num_vecinos_3r = 2
            '21 coger genes de padres alternos
            coger_alternos_3r = False
            '22 Forma de agrupar los agentes para jugar partidas en la evaluación
            personas_por_grupo_3r = "Todos contra todos"
            '24 Tipo_Mutacion
            Tipo_Mutacion_3r = 50
            '25 Pesos_Partir_Cero
            Pesos_Partir_Cero_3r = False
    
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Ejemplo inexistente"
    End Select



End Sub


Sub s_cargar_opciones_3r()

    '1 Número de agentes inicial
    frm_c3_op3r.Cb_num_agentes.Text = numero_total_de_agentes_ejv
    '3 Número de reglas por agente
    frm_c3_op3r.Op_numero_reglas_variable_3r = numero_reglas_variable_3r
    frm_c3_op3r.Op_nnumero_reglas_variable_3r = Not (numero_reglas_variable_3r)
    frm_c3_op3r.Text5.Text = numero_de_reglas_por_agente_3r
    frm_c3_op3r.var11 = var11_3r
    frm_c3_op3r.var12 = var12_3r
    frm_c3_op3r.var21 = var21_3r
    frm_c3_op3r.var22 = var22_3r
    frm_c3_op3r.var31 = var31_3r
    frm_c3_op3r.var32 = var32_3r
    frm_c3_op3r.var41 = var41_3r
    frm_c3_op3r.var42 = var42_3r
    frm_c3_op3r.var51 = var51_3r
    var11_3r = numero_de_reglas_por_agente_3r
    '11 La generación de reglas se hace totalmente al azar
    frm_c3_op3r.Op_ReglasAzar = reglas_azar_3r
    frm_c3_op3r.Op_nReglasAzar = Not (reglas_azar_3r)
    '12 ver agentes
    frm_c3_op3r.Op_VerAgentes = ver_agentes_3r
    frm_c3_op3r.Op_nVerAgentes = Not (ver_agentes_3r)


End Sub

Sub s_grabar_opciones_3r()

    '1 Número de agentes inicial
    numero_total_de_agentes_ejv = CInt(frm_c3_op3r.Cb_num_agentes.Text)
    '3 Número de reglas por agente
    'El numero de reglas de los agentes puede variar de unos ciclos a otros
    numero_reglas_variable_3r = frm_c3_op3r.Op_numero_reglas_variable_3r
    If Not numero_reglas_variable_3r Then
        numero_de_reglas_por_agente_3r = CInt(frm_c3_op3r.Text5.Text)
    Else
        var11_3r = frm_c3_op3r.var11
        var12_3r = frm_c3_op3r.var12
        var21_3r = frm_c3_op3r.var21
        var22_3r = frm_c3_op3r.var22
        var31_3r = frm_c3_op3r.var31
        var32_3r = frm_c3_op3r.var32
        var41_3r = frm_c3_op3r.var41
        var42_3r = frm_c3_op3r.var42
        var51_3r = frm_c3_op3r.var51
        numero_de_reglas_por_agente_3r = var11_3r
    End If
    '11 La generación de reglas se hace totalmente al azar
    reglas_azar_3r = frm_c3_op3r.Op_ReglasAzar
    '12 ver agentes
    ver_agentes_3r = frm_c3_op3r.Op_VerAgentes


End Sub


Sub s_grabar_opciones_ev_3r()

    '7 Elección de Jugadores
    If frm_c3_ev3r.Op_JugadoresAzar = 1 Then
        eleccion_de_jugadores_al_azar_3r = True
    Else
        eleccion_de_jugadores_al_azar_3r = False
    End If
    '9 El peso se calcula en función de si existen más entidades parecidas a la actual
    If frm_c3_ev3r.Op_Relativo = 1 Then
        pesos_relativos_3r = True
    Else
        pesos_relativos_3r = False
    End If
    '17 Numero de Partidas que juega cada pareja
    NumeroPartidas_3r = frm_c3_ev3r.Op_NumeroPartidas_3r
    '18 Numero de entidades cercanas con las que comparo al ajustar pesos
    Numero_cercanos_relativo_3r = frm_c3_ev3r.Op_Numero_cercanos_relativo_3r
    '19 Método de modificar los pesos despues de una partida
    If frm_c3_ev3r.Op_Metodo_de_asignar_pesos_3r.ListIndex = 0 Then
        Metodo_de_asignar_pesos_3r = "A"
    ElseIf frm_c3_ev3r.Op_Metodo_de_asignar_pesos_3r.ListIndex = 1 Then
        Metodo_de_asignar_pesos_3r = "B"
    ElseIf frm_c3_ev3r.Op_Metodo_de_asignar_pesos_3r.ListIndex = 2 Then
        Metodo_de_asignar_pesos_3r = "C"
    ElseIf frm_c3_ev3r.Op_Metodo_de_asignar_pesos_3r.ListIndex = 3 Then
        Metodo_de_asignar_pesos_3r = "D"
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: método de asignar pesos no contemplado"
    End If
    '22 Grupos que juegan partidas
    If frm_c3_ev3r.Op_Personas_por_grupo.ListIndex = 0 Then
        personas_por_grupo_3r = "2 Jugadores"
    ElseIf frm_c3_ev3r.Op_Personas_por_grupo.ListIndex = 1 Then
        personas_por_grupo_3r = "Todos contra todos"
    Else
        s_error_ejv CON_OPCION_FINALIZAR, "Error: agentes por grupo no contemplado"
    End If
    '25 Pesos_Partir_Cero
    If frm_c3_ev3r.Op_Pesos_Partir_Cero = 1 Then
        Pesos_Partir_Cero_3r = True
    Else
        Pesos_Partir_Cero_3r = False
    End If

End Sub

Sub s_cargar_opciones_rep_mut_ce()

    '1 Probabilidad de Mutación
    frm_c3_rm.Op_TasaMutacion.Text = tasa_de_mutacion_3r
    '2 Mutaciones acumuladas
    frm_c3_rm.Op_Acumulada = mutaciones_acumuladas_3r
    frm_c3_rm.Op_nAcumulada = Not mutaciones_acumuladas_3r
    '3 Padres iguales producen hijos que son mutaciones en todos sus elementos.
    frm_c3_rm.Op_PadresIdenticos = padres_identicos_producen_mutaciones_3r
    frm_c3_rm.Op_nPadresIdenticos = Not padres_identicos_producen_mutaciones_3r
    '4 sustituir reglas repetidas por mutaciones
    frm_c3_rm.Op_CrearMutacionesEnRepetidas_3r = quitar_reglas_repetidas_3r
    frm_c3_rm.Op_nCrearMutacionesEnRepetidas_3r = Not quitar_reglas_repetidas_3r
    '5 sustituir cada ciclo por mutaciones las reglas con un peso menor o igual que ___
    frm_c3_rm.sust1 = sust1_3r
    '6 sustituir cada ___ ciclos por mutaciones las reglas con un peso menor o igual que ___
    frm_c3_rm.sust2 = sust2_3r
    frm_c3_rm.sust3 = sust3_3r
    '7 Tipo_Mutacion
     frm_c3_rm.Op_Tipo_Mutacion = CStr(Tipo_Mutacion_3r)

End Sub

Sub s_grabar_opciones_rep_mut_ce()

    '1 Probabilidad de Mutación
    tasa_de_mutacion_3r = frm_c3_rm.Op_TasaMutacion.Text
    '2 Mutaciones acumuladas
    mutaciones_acumuladas_3r = frm_c3_rm.Op_Acumulada
    '3 Padres iguales producen hijos que son mutaciones en todos sus elementos.
    padres_identicos_producen_mutaciones_3r = frm_c3_rm.Op_PadresIdenticos
    '4 sustituir reglas repetidas por mutaciones
    quitar_reglas_repetidas_3r = frm_c3_rm.Op_CrearMutacionesEnRepetidas_3r
    '5 sustituir cada ciclo por mutaciones las reglas con un peso menor o igual que ___
    sust1_3r = frm_c3_rm.sust1
    '6 sustituir cada ___ ciclos por mutaciones las reglas con un peso menor o igual que ___
    sust2_3r = frm_c3_rm.sust2
    sust3_3r = frm_c3_rm.sust3
    '7 Tipo_Mutacion
    Tipo_Mutacion_3r = CInt(frm_c3_rm.Op_Tipo_Mutacion)

End Sub

Sub s_cargar_opciones_rep_sob_ce()
    
    '1 Elección de Padres
    frm_c3_rs.Op_PadresAzar_3r = eleccion_de_padres_al_azar_3r
    frm_c3_rs.Op_nPadresAzar_3r = Not eleccion_de_padres_al_azar_3r
    '2 Las entidades modifican sus propios pesos de reglas despues de jugar
    frm_c3_rs.Op_ModificarPropiosPesos = modificar_pesos_3r
    frm_c3_rs.Op_nModificarPropiosPesos = Not modificar_pesos_3r
    '3 Heredar la regla de mas peso de los dos padres en vez de tomarla al azar
    frm_c3_rs.Op_HeredarReglaMasPeso_3r = heredar_regla_mas_peso_3r
    frm_c3_rs.Op_nHeredarReglaMasPeso_3r = Not heredar_regla_mas_peso_3r
    '4 Compartir el conocimiento (modificar pesos reglas) con los num_vecinos_3r mas cercanos pero no de uno mismo
    frm_c3_rs.Op_CompartirConocimiento = compartir_conocimiento_3r
    frm_c3_rs.Op_nCompartirConocimiento = Not compartir_conocimiento_3r
    frm_c3_rs.Op_CompartirNumVecinos = num_vecinos_3r
    '5 coger alternos los genes de cada padre
    frm_c3_rs.Op_CogerAlternos = coger_alternos_3r
    frm_c3_rs.Op_nCogerAlternos = Not coger_alternos_3r
    
    

End Sub


Sub s_grabar_opciones_rep_sob_ce()

    '1 Elección de Padres
    eleccion_de_padres_al_azar_3r = frm_c3_rs.Op_PadresAzar_3r
    '2 Las entidades modifican sus propios pesos de reglas despues de jugar
    modificar_pesos_3r = frm_c3_rs.Op_ModificarPropiosPesos
    '3 Heredar la regla de mas peso de los dos padres en vez de tomarla al azar
    heredar_regla_mas_peso_3r = frm_c3_rs.Op_HeredarReglaMasPeso_3r
    '4 Compartir el conocimiento (modificar pesos reglas) con los num_vecinos_3r mas cercanos pero no de uno mismo
    compartir_conocimiento_3r = frm_c3_rs.Op_CompartirConocimiento
    num_vecinos_3r = frm_c3_rs.Op_CompartirNumVecinos
    '5 coger alternos los genes de cada padre
    coger_alternos_3r = frm_c3_rs.Op_CogerAlternos
    
    
    
End Sub

Sub s_grabar_opciones_sel_3r()

    '4 Tipo de Selección-Reproducción
    If frm_c3_sel3r.Op_s1 Then
        tipo_reproduccion_3r = 1
    Else
        If frm_c3_sel3r.Op_s2 Then
            tipo_reproduccion_3r = 2
        Else
            If frm_c3_sel3r.Op_s3 Then
                tipo_reproduccion_3r = 3
            Else
                tipo_reproduccion_3r = 4
            End If
        End If
    End If

End Sub
Sub s_mostrar_regla_de_mas_peso(casilla As Integer, num_agente As Integer)

    Dim peso_de_la_regla_de_mas_peso_3r As Long
    Dim regla_de_mas_peso As String
    Dim regla As String
    Dim cont_regla As Integer

    'Mostramos la regla de mas peso del agente
    peso_de_la_regla_de_mas_peso_3r = -1
    regla_de_mas_peso = ""
    For cont_regla = 1 To numero_de_reglas_por_agente_3r
        regla = f_tomar_regla_de_agente_3r(num_agente, cont_regla)
        If peso_regla_agente_3r(num_agente, cont_regla) > peso_de_la_regla_de_mas_peso_3r Then
            peso_de_la_regla_de_mas_peso_3r = peso_regla_agente_3r(num_agente, cont_regla)
            regla_de_mas_peso = regla
        End If
    Next cont_regla
    
    Select Case casilla
        Case 1
            frm_c3_juego3r.r1 = regla_de_mas_peso
            frm_c3_juego3r.p1 = peso_de_la_regla_de_mas_peso_3r
        Case 2
            frm_c3_juego3r.r2 = regla_de_mas_peso
            frm_c3_juego3r.p2 = peso_de_la_regla_de_mas_peso_3r
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: casilla inexistente"
    End Select

End Sub


Sub s_load_carga_3r()

    frm_c0_ce.Fr_ModificarAgente.Visible = False
    s_botones_activos_3r True
    
    frm_c0_ce.super.Enabled = False
    
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, False
    
    paso_3r = 14
    
    se_han_creado_los_agentes_3r = False
    numero_de_caracteres_por_regla_3r = 10
    numero_de_caracteres_parte_izq_regla_3r = 9
    
    If Not automatico_ejv Then
        s_inicializar_ejemplo_elegido_ejv
    End If
    
    esta_detenido_ejv = True

End Sub

Sub s_ver_mejores_3r()

    frm_c0_ce.fr_Todas.Visible = False
    frm_c0_ce.Fr_Ejecucion.Visible = True
    
    s_pintar_mejores_agentes_3r
    ver_agentes_3r = True
    
    s_botones_activos_3r True
    's_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
    's_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
    's_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION, True
    '    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_EVALUACION, True
    '    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_SELECCION, True
    '    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION, True
    '        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES, True
    '        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO, True
    's_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, True
    's_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, True
    's_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_MEJORES, True
    
    s_mostrar_estado_semaforo frm_c3_in3r, CTE_DETENIDO

    'Con esto borro el grafico
    frm_c0_ce.Refresh

End Sub

Sub s_ver_todos_los_agentes_3r()

    Dim txt As String
    
    txt = "Ver todos los agentes puede ser un poco lento. ¿Está seguro de que desea listarlos?"
    If idioma_ejv = CTE_INGLES Then
        txt = "To see all the agents could be very a little slow. ¿Are you sure that you want to see all?"
    End If

    If MsgBox(txt, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = CTE_ARENA
        frm_c0_ce.Fr_Ejecucion.Visible = False
        frm_c0_ce.fr_Todas.Visible = True
        'Con esto borro el grafico
        frm_c0_ce.Refresh
        s_pintar_todos_los_agentes_3r
        Screen.MousePointer = CTE_DEFECTO
    End If

    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_EVALUACION, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_SELECCION, True
        s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION, True
            s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES, True
            s_cambiar_estado_enabled_menus_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO, True
    
    s_cambiar_estado_enabled_menus_ejv CTE_VER_GRAFICO, True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_MEJORES, True
    s_mostrar_estado_semaforo frm_c3_in3r, CTE_DETENIDO


End Sub

Sub s_ver_jugar_contra_ordenador_3r()

Dim cont_regla As Integer

If esta_detenido_ejv = True And hay_que_detener_ejv = False Then
    
    If Not se_han_creado_los_agentes_3r Then
        'todavia no se ha aprendido, y no hay nungun agente generado
        s_grabar_opciones_3r
        's_grabar_opciones_ev_3r
        's_grabar_opciones_rep_3r
        's_grabar_opciones_sel_3r
        'creamos uno al azar
        MsgBox "La evolución y el aprendizaje no se han producido todavía. Se va a jugar contra un jugador creado al azar.", vbInformation
        'Creamos 1 agente
        ReDim agente_3r(1 To 1) As String
        ReDim peso_regla_agente_3r(1 To 1, 1 To numero_de_reglas_por_agente_3r) As Long
        ReDim prioridad_regla_agente_3r(1 To 1, 1 To numero_de_reglas_por_agente_3r) As Integer
        'de numero_de_reglas_por_agente_3r elementos
        For cont_regla = 1 To numero_de_reglas_por_agente_3r
            agente_3r(1) = agente_3r(1) & f_crear_regla_al_azar_3r
            peso_regla_agente_3r(1, cont_regla) = 1 'cualquier cosa
            prioridad_regla_agente_3r(1, cont_regla) = 1 'cualquier cosa
        Next cont_regla
    End If
    s_mostrar_estado_semaforo frm_c3_in3r, estado_3r
   'Mostramos la pantalla
    Unload frm_c3_juego3r
    estado_3r = CTE_JUGANDO
    frm_c3_juego3r.Show

End If

End Sub
