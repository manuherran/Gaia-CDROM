Attribute VB_Name = "bas_z0_arra"
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


Sub Si_DesordenMedioArray_2D_S(mi_array() As String, maximo_dim1 As Integer, primero As Integer, ultimo As Integer)

'Desordena el segmento de un array que va desde
'primero hasta ultimo

Dim I_n As Integer


Dim longitud As Integer
ReDim temp(1 To maximo_dim1) As String

Dim uno As String
Dim otro As String

Dim j As Integer

longitud = ultimo - primero + 1

For I_n = 1 To 2 * longitud
   'Intercambio 2 elementos a azar
    uno = fi_azar2(primero, ultimo)
    otro = fi_azar2(primero, ultimo)
    
   'Muevo cada elemento de la primera dimensión, que son estaticos
    For j = 1 To maximo_dim1
        temp(j) = mi_array(j, uno)
        mi_array(j, uno) = mi_array(j, otro)
        mi_array(j, otro) = temp(j)
    Next j
Next I_n



End Sub

Sub Si_DesordenMedioArrayI(mi_array() As Integer, primero As Integer, ultimo As Integer)

'Desordena el segmento de un array que va desde
'primero hasta ultimo

Dim I_n As Integer


Dim longitud As Integer
Dim temp As Integer

Dim uno As Integer
Dim otro As Integer


longitud = ultimo - primero + 1

For I_n = 1 To 2 * longitud
   'Intercambio 2 elementos a azar
    uno = fi_azar2(primero, ultimo)
    otro = fi_azar2(primero, ultimo)
    temp = mi_array(uno)
    mi_array(uno) = mi_array(otro)
    mi_array(otro) = temp
Next I_n


End Sub


Sub f_desordenar_array_l(mi_array() As Long)
    'mi_array() es un parametro de entrada-salida

    'Desordena el array completo
    Dim I_n As Integer
    
    Dim primero As Integer
    Dim ultimo As Integer
    
    Dim longitud As Integer
    Dim temp As Long
    
    Dim uno As Long
    Dim otro As Long
    
    
    primero = LBound(mi_array)
    ultimo = UBound(mi_array)
    longitud = ultimo - primero + 1
    
    For I_n = 1 To 2 * longitud
       'Intercambio 2 elementos a azar
        uno = fi_azar2(primero, ultimo)
        otro = fi_azar2(primero, ultimo)
        
        temp = mi_array(uno)
        mi_array(uno) = mi_array(otro)
        mi_array(otro) = temp
    Next I_n


End Sub
Sub f_desordenar_array_i(mi_array() As Integer)
    'mi_array() es un parametro de entrada-salida

    'Desordena el array completo
    Dim I_n As Integer
    
    Dim primero As Integer
    Dim ultimo As Integer
    
    Dim longitud As Integer
    Dim temp As Integer
    
    Dim uno As Integer
    Dim otro As Integer
    
    
    primero = LBound(mi_array)
    ultimo = UBound(mi_array)
    longitud = ultimo - primero + 1
    
    For I_n = 1 To 2 * longitud
       'Intercambio 2 elementos a azar
        uno = fi_azar2(primero, ultimo)
        otro = fi_azar2(primero, ultimo)
        
        temp = mi_array(uno)
        mi_array(uno) = mi_array(otro)
        mi_array(otro) = temp
    Next I_n


End Sub
Function f_borra_elemento_array_string(indice As Integer, mi_array() As String) As Integer

    Dim i As Integer
    Dim encontrado As Boolean
    'Dim li As Integer
    'Dim ls As Integer
    
    
    encontrado = False
    For i = indice + 1 To UBound(mi_array)
        mi_array(i - 1) = mi_array(i)
        encontrado = True
    Next i
    
    'li = LBound(mi_array)
    'ls = UBound(mi_array) - 1
    
    ReDim Preserve mi_array(LBound(mi_array) To UBound(mi_array) - 1) As String
    
    If encontrado Then
        f_borra_elemento_array_string = True
    Else
        f_borra_elemento_array_string = False
    End If

End Function

Function f_borra_elemento_array_long(indice As Integer, mi_array() As Long) As Integer

    Dim i As Integer
    Dim encontrado As Boolean
    
    encontrado = False
    For i = indice + 1 To UBound(mi_array)
        mi_array(i - 1) = mi_array(i)
        encontrado = True
    Next i
    ReDim Preserve mi_array(LBound(mi_array) To UBound(mi_array) - 1) As Long
    
    If encontrado Then
        f_borra_elemento_array_long = True
    Else
        f_borra_elemento_array_long = False
    End If

End Function

Function f_busca_elemento_array_string(elemento As String, mi_array() As String, LI As Integer, LS As Integer) As Integer

    'Ejemplo llamada
    'casilla_a_borrar = f_busca_elemento_array_string("Azul", miarr(), 1, 30)
    
    Dim posicion_elemento  As Integer
    Dim i As Integer
    
    posicion_elemento = 0
    i = LI
    While mi_array(i) <> elemento And i <= LS
        i = i + 1
    Wend
    
    If i > LS Then
        f_busca_elemento_array_string = 0
    Else
        f_busca_elemento_array_string = i
    End If


End Function


Function f_busca_elemento_array_integer(elemento As Integer, mi_array() As Integer, LI As Integer, LS As Integer) As Integer

    'Ejemplo llamada
    'casilla_a_borrar = f_busca_elemento_array_integer(pos, lista_de_casillas_libres_3r(), 1, numero_casillas_libres_3r)
    
    Dim posicion_elemento  As Integer
    Dim i As Integer
    
    posicion_elemento = 0
    i = LI
    While mi_array(i) <> elemento And i <= LS
        i = i + 1
    Wend
    
    If i > LS Then
        f_busca_elemento_array_integer = 0
    Else
        f_busca_elemento_array_integer = i
    End If


End Function
Sub S_OrdenarArray3DIntMinMax(obj_pintar_p() As Double, obj_pintar_f() As Double, obj_pintar_c() As Double, obj_pintar_o() As Integer)

    Select Case algoritmo_ordenacion_ejv
        Case CTE_BURBUJA
            S_OrdenarArray3DIntMinMax_bur obj_pintar_p(), obj_pintar_f(), obj_pintar_c(), obj_pintar_o()
        Case CTE_QUICKSORT
            S_OrdenarArray3DIntMinMax_qui obj_pintar_p(), obj_pintar_f(), obj_pintar_c(), obj_pintar_o()
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Algoritmo de ordenación inexistente"
    End Select

End Sub

Sub S_OrdenarArray3DIntMinMax_bur(obj_pintar_p() As Double, obj_pintar_f() As Double, obj_pintar_c() As Double, obj_pintar_o() As Integer)

    Dim I_n As Long
    Dim I_x As Long
    Dim suma_n As Long
    Dim suma_x As Long
    Dim primero As Long
    Dim ultimo As Long
    Dim temp_p As Double
    Dim temp_f As Double
    Dim temp_c As Double
    Dim temp_o As Long
    
    ' de < a >
    primero = LBound(obj_pintar_p)
    ultimo = UBound(obj_pintar_p)
    'Comparo cada elemento (todos menos el último)....
    For I_n = primero To ultimo - 1
       'con el siguiente y todos los demás hasta el último
        For I_x = I_n + 1 To ultimo
            suma_n = obj_pintar_p(I_n) + obj_pintar_f(I_n) + obj_pintar_c(I_n)
            suma_x = obj_pintar_p(I_x) + obj_pintar_f(I_x) + obj_pintar_c(I_x)
            If suma_x < suma_n Then
                temp_p = obj_pintar_p(I_x)
                temp_f = obj_pintar_f(I_x)
                temp_c = obj_pintar_c(I_x)
                temp_o = obj_pintar_o(I_x)
                obj_pintar_p(I_x) = obj_pintar_p(I_n)
                obj_pintar_f(I_x) = obj_pintar_f(I_n)
                obj_pintar_c(I_x) = obj_pintar_c(I_n)
                obj_pintar_o(I_x) = obj_pintar_o(I_n)
                obj_pintar_p(I_x) = temp_p
                obj_pintar_f(I_x) = temp_f
                obj_pintar_c(I_x) = temp_c
                obj_pintar_o(I_x) = temp_o
    DoEvents
            End If
        Next I_x
    Next I_n


End Sub

Sub S_OrdenarArray3DIntMinMax_qui(obj_pintar_p() As Double, obj_pintar_f() As Double, obj_pintar_c() As Double, obj_pintar_o() As Integer)
    
    s_error_ejv CON_OPCION_FINALIZAR, "Quick Sort no programado"

End Sub

Sub S_OrdenarArray1IntMaxMin(mi_array() As Integer)

    Select Case algoritmo_ordenacion_ejv
        Case CTE_BURBUJA
            S_OrdenarArray1IntMaxMin_bur mi_array()
        Case CTE_QUICKSORT
            S_OrdenarArray1IntMaxMin_qui mi_array()
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Algoritmo de ordenación inexistente"
    End Select

End Sub

Sub S_OrdenarArray1LngMaxMin(mi_array() As Long)

    Select Case algoritmo_ordenacion_ejv
        Case CTE_BURBUJA
            S_OrdenarArray1LngMaxMin_bur mi_array()
        Case CTE_QUICKSORT
            S_OrdenarArray1LngMaxMin_qui mi_array()
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: Algoritmo de ordenación inexistente"
    End Select

End Sub

Sub S_OrdenarArray1LngMaxMin_bur(mi_array() As Long)

    Dim I_n As Long
    Dim I_x As Long
    Dim primero As Long
    Dim ultimo As Long
    Dim i_temp As Long

    ' de > a <
    primero = LBound(mi_array)
    ultimo = UBound(mi_array)
    'Comparo cada elemento (todos menos el último)....
    For I_n = primero To ultimo - 1
       'con el siguiente y todos los demás hasta el último
        For I_x = I_n + 1 To ultimo
            If mi_array(I_x) > mi_array(I_n) Then
                i_temp = mi_array(I_x)
                mi_array(I_x) = mi_array(I_n)
                mi_array(I_n) = i_temp
    DoEvents
            End If
        Next I_x
    Next I_n

End Sub

Sub S_OrdenarArrayStrMinMax_bur(mi_array() As String)

    Dim I_n As Integer
    Dim I_x As Integer
    Dim primero As Integer
    Dim ultimo As Integer
    Dim i_temp As String

    ' de < a >
    primero = LBound(mi_array)
    ultimo = UBound(mi_array)
    'Comparo cada elemento (todos menos el último)....
    For I_n = primero To ultimo - 1
       'con el siguiente y todos los demás hasta el último
        For I_x = I_n + 1 To ultimo
            If mi_array(I_x) < mi_array(I_n) Then
                i_temp = mi_array(I_x)
                mi_array(I_x) = mi_array(I_n)
                mi_array(I_n) = i_temp
    DoEvents
            End If
        Next I_x
    Next I_n

End Sub

Sub S_OrdenarArray1IntMaxMin_bur(mi_array() As Integer)

    Dim I_n As Integer
    Dim I_x As Integer
    Dim primero As Integer
    Dim ultimo As Integer
    Dim i_temp As Integer

    ' de > a <
    primero = LBound(mi_array)
    ultimo = UBound(mi_array)
    'Comparo cada elemento (todos menos el último)....
    For I_n = primero To ultimo - 1
       'con el siguiente y todos los demás hasta el último
        For I_x = I_n + 1 To ultimo
            If mi_array(I_x) > mi_array(I_n) Then
                i_temp = mi_array(I_x)
                mi_array(I_x) = mi_array(I_n)
                mi_array(I_n) = i_temp
    DoEvents
            End If
        Next I_x
    Next I_n

End Sub
Sub S_OrdenarArray1IntMaxMin_qui(mi_array() As Integer)

    s_error_ejv CON_OPCION_FINALIZAR, "Quick Sort no programado"
    

End Sub
Sub S_OrdenarArray1LngMaxMin_qui(mi_array() As Long)

    s_error_ejv CON_OPCION_FINALIZAR, "Quick Sort no programado"
    

End Sub

Function Fi_OrdenarEspecial2(ArrayGuia() As Integer, ArrayAOrdenar() As Integer, primero As Integer, ultimo As Integer)

'Esta función toma un array, realiza una copia, y
'ordena otro array (machacando la copia, pq la ordena)
'y lo hace entre 2 límites que se le dan como parámetros


    Dim I_n As Integer
    Dim I_x As Integer
    Dim i_temp As Integer
    Dim Copia() As Integer

    s_copiar_array1int ArrayGuia(), Copia()
    ' de > a <
    'Comparo cada elemento (todos menos el último)....
    For I_n = primero To ultimo - 1
       'con el siguiente y todos los demás hasta el último
        For I_x = I_n + 1 To ultimo
    DoEvents
            If Copia(I_x) > Copia(I_n) Then
               'en este orden
               'ordeno el array de patrones
                i_temp = Copia(I_x)
                Copia(I_x) = Copia(I_n)
                Copia(I_n) = i_temp
    
               'ordeno los indices
                i_temp = ArrayAOrdenar(I_x)
                ArrayAOrdenar(I_x) = ArrayAOrdenar(I_n)
                ArrayAOrdenar(I_n) = i_temp
            End If
        Next I_x
    Next I_n

End Function

Function Fi_Posicion(AR_I_Secuencia() As Integer, AR2_I_TotArbol() As Integer, AR2_I_Arbol() As Integer, I_Nivel As Integer, AR_I_Columna() As Integer) As Integer
'IN:    Secuencia que identifica una posición en el arbol: 2, 1, 4
'OUT:   Posición en el string: 47

Dim I_suma As Integer
Dim I_UltimoNivel As Integer
Dim I_CuentaNivel As Integer
Dim I_Col As Integer
Dim I_Niv As Integer
Dim I_x As Integer
Dim I_Sumado As Integer
Dim I_Sumado2 As Integer



'Inicializaciones
I_suma = 0
I_UltimoNivel = UBound(AR_I_Secuencia)




'Si estoy en el primer nivel inicializo el array de columnas
'por si se olvida refrescar los ceros!!
If I_Nivel = 1 Then
    For I_CuentaNivel = 1 To I_UltimoNivel
        AR_I_Columna(I_CuentaNivel) = 0
    DoEvents
    Next I_CuentaNivel
End If

'AR_I_Columna contiene el nº datos dejados a la izquierda



'Si estoy en el último nivel.. no hay hijos
If I_Nivel = I_UltimoNivel Then
   'devuelvo el valor,
    I_suma = AR2_I_Arbol(I_Nivel, AR_I_Columna(I_Nivel) + 1)

Else

   'Sumo todos los hijos anteriores al hijo buscado
    For I_Col = 1 To AR_I_Secuencia(I_Nivel) - 1


'        AR_I_Columna(I_nivel) = AR_I_Columna(I_nivel) + I_Col
'        I_suma = I_suma + Fi_Posicion(AR_I_Secuencia(), AR2_I_Arbol(), I_nivel + 1, AR_I_Columna())
        I_suma = I_suma + AR2_I_TotArbol(I_Nivel + 1, AR_I_Columna(I_Nivel + 1) + I_Col)

    Next I_Col

'   Actualizo las columnas
    I_Sumado = AR_I_Secuencia(I_Nivel) - 1
    AR_I_Columna(I_Nivel + 1) = AR_I_Columna(I_Nivel + 1) + I_Sumado
    For I_Niv = I_Nivel + 2 To I_UltimoNivel
        I_Sumado2 = 0
        For I_x = 1 To I_Sumado
            I_Sumado2 = I_Sumado2 + AR2_I_Arbol(I_Niv - 1, I_x)
        Next I_x
        AR_I_Columna(I_Niv) = AR_I_Columna(I_Niv) + I_Sumado2
        I_Sumado = I_Sumado2

    Next I_Niv




   'Para el hijo buscado..
   'Desciendo una posición y llamo
    I_Nivel = I_Nivel + 1
    If I_Nivel < I_UltimoNivel Then
        I_suma = I_suma + Fi_Posicion(AR_I_Secuencia(), AR2_I_TotArbol(), AR2_I_Arbol(), I_Nivel, AR_I_Columna())
    Else

        I_suma = I_suma + AR_I_Secuencia(I_UltimoNivel)
    End If



End If


Fi_Posicion = I_suma

End Function

Function Fi_Posicion_OLD(AR_I_Secuencia() As Integer, AR2_I_Arbol() As Integer, I_Nivel As Integer, AR_I_Columna() As Integer)

'IN:    Secuencia que identifica una posición en el arbol: 2, 1, 4
'OUT:   Posición en el string: 47

Dim I_suma As Integer
Dim I_UltimoNivel As Integer
Dim I_CuentaNivel As Integer
Dim I_Col As Integer



'Inicializaciones
I_suma = 0
I_UltimoNivel = UBound(AR_I_Secuencia)




'Si estoy en el primer nivel inicializo el array de columnas
If I_Nivel = 1 Then
    For I_CuentaNivel = 1 To I_UltimoNivel
        AR_I_Columna(I_CuentaNivel) = 1
    DoEvents
    Next I_CuentaNivel
End If




'Si estoy en el último nivel.. no hay hijos
If I_Nivel = I_UltimoNivel Then
   'devuelvo el valor,
    I_suma = AR2_I_Arbol(I_Nivel, AR_I_Columna(I_Nivel))

Else

   'Sumo todos los hijos anteriores al hijo buscado
    For I_Col = 1 To AR_I_Secuencia(I_Nivel) - 1
        AR_I_Columna(I_Nivel) = AR_I_Columna(I_Nivel) + I_Col
        I_suma = I_suma + Fi_Posicion_OLD(AR_I_Secuencia(), AR2_I_Arbol(), I_Nivel + 1, AR_I_Columna())
    Next I_Col

   'Para el hijo buscado..
   'Desciendo una posición y llamo
    I_Nivel = I_Nivel + 1
    If I_Nivel < I_UltimoNivel Then
        I_suma = I_suma + Fi_Posicion_OLD(AR_I_Secuencia(), AR2_I_Arbol(), I_Nivel, AR_I_Columna())
    Else
        I_suma = I_suma + 1
    End If



End If


Fi_Posicion_OLD = I_suma

End Function

Sub s_copiar_array1lng(Array1() As Long, Array2() As Long)

   'Declaración de variables
    Dim L_n As Long

   'Redimensionamos el Array2
    ReDim Array2(LBound(Array1) To UBound(Array1)) As Long

    'Bucle
     For L_n = LBound(Array1, 1) To UBound(Array1, 1)
        Array2(L_n) = Array1(L_n)
        DoEvents
    Next L_n

End Sub

Sub s_copiar_array1int(Array1() As Integer, Array2() As Integer)

   'Declaración de variables
    Dim I_n As Integer

   'Redimensionamos el Array2
    ReDim Array2(LBound(Array1) To UBound(Array1)) As Integer

    'Bucle
     For I_n = LBound(Array1, 1) To UBound(Array1, 1)
        Array2(I_n) = Array1(I_n)
        DoEvents
    Next I_n

End Sub

Sub s_copiar_array1str(Array1() As String, Array2() As String)

   'Declaración de variables
    Dim I_n As Integer

   'Redimensionamos el Array2
    ReDim Array2(LBound(Array1) To UBound(Array1)) As String

    'Bucle
     For I_n = LBound(Array1, 1) To UBound(Array1, 1)
        Array2(I_n) = Array1(I_n)
        DoEvents
    Next I_n

End Sub

Sub S_OrdenarEspecialLng(ArrayGuia() As Long, ArrayAOrdenar() As Long)

    '=======================================================
    'Ordena un array en funcion de los valores de otro array
    '=======================================================


    'Esta función toma un array, realiza una copia, y
    'ordena otro array (machacando la copia, pq la ordena)
    
    Dim L_n As Long
    Dim I_x As Long
    Dim primero As Long
    Dim ultimo As Long
    Dim L_Temp As Long
    Dim Copia() As Long

    
    s_copiar_array1lng ArrayGuia(), Copia()
    ' de > a <
    primero = LBound(Copia)
    ultimo = UBound(Copia)
    'Comparo cada elemento (todos menos el último)....
    For L_n = primero To ultimo - 1
       'con el siguiente y todos los demás hasta el último
        For I_x = L_n + 1 To ultimo
    DoEvents
            If Copia(I_x) > Copia(L_n) Then
               'en este orden
               'ordeno el array de patrones
                L_Temp = Copia(I_x)
                Copia(I_x) = Copia(L_n)
                Copia(L_n) = L_Temp
               'ordeno los indices
                L_Temp = ArrayAOrdenar(I_x)
                ArrayAOrdenar(I_x) = ArrayAOrdenar(L_n)
                ArrayAOrdenar(L_n) = L_Temp
            End If
        Next I_x
    Next L_n

End Sub


Sub S_OrdenarEspecialInt(ArrayGuia() As Integer, ArrayAOrdenar() As Integer)

    '=======================================================
    'Ordena un array en funcion de los valores de otro array
    '=======================================================


    'Esta función toma un array, realiza una copia, y
    'ordena otro array (machacando la copia, pq la ordena)
    
    Dim I_n As Integer
    Dim I_x As Integer
    Dim primero As Integer
    Dim ultimo As Integer
    Dim i_temp As Integer
    Dim Copia() As Integer

    
    s_copiar_array1int ArrayGuia(), Copia()
    ' de > a <
    primero = LBound(Copia)
    ultimo = UBound(Copia)
    'Comparo cada elemento (todos menos el último)....
    For I_n = primero To ultimo - 1
       'con el siguiente y todos los demás hasta el último
        For I_x = I_n + 1 To ultimo
    DoEvents
            If Copia(I_x) > Copia(I_n) Then
               'en este orden
               'ordeno el array de patrones
                i_temp = Copia(I_x)
                Copia(I_x) = Copia(I_n)
                Copia(I_n) = i_temp
    
               'ordeno los indices
                i_temp = ArrayAOrdenar(I_x)
                ArrayAOrdenar(I_x) = ArrayAOrdenar(I_n)
                ArrayAOrdenar(I_n) = i_temp
    
    
            End If
        Next I_x
    Next I_n



End Sub

Function f_array_l_a_listacomas(miarray() As Long) As String

    Dim i As Integer
    Dim res As String
    
    res = ""
    For i = LBound(miarray()) To UBound(miarray())
        res = res & miarray(i) & ","
    Next i
    'quito la ultima coma
    res = Left(res, Len(res) - 1)

    f_array_l_a_listacomas = res

End Function


Sub s_ordenar_qui_dos(ar_decision() As Double, arr_doble() As String)

    Dim i As Integer
    Dim j As Integer
    
    Dim cont2 As Integer
    Dim c3 As Integer
    

    Dim primero As Integer
    Dim ultimo As Integer
    Dim medio As Integer
    
    Dim primero_bis As Integer
    Dim ultimo_bis As Integer
    
    primero = UBound(ar_decision())
    ultimo = LBound(ar_decision())
    medio = CLng((primero + ultimo) / 2)
    
    primero_bis = UBound(arr_doble(), 2)
    ultimo_bis = LBound(arr_doble(), 2)
    
    ReDim ar_tmp2_dec(primero To medio) As Double
    ReDim ar_tmp2_dob(primero To medio, primero_bis To ultimo_bis) As String

    ReDim ar_tmp3_dec(medio To ultimo) As Double
    ReDim ar_tmp3_dob(medio To ultimo, primero_bis To ultimo_bis) As String

    'Copio los arrays en los temporales
    For i = primero To medio
        'El de decision
        ar_tmp2_dec(i) = ar_decision(i)
        'El doble
        For j = primero_bis To ultimo_bis
            ar_tmp2_dob(j, i) = arr_doble(j, i)
        Next j
    Next i
    For i = medio + 1 To ultimo
        'El de decision
        ar_tmp3_dec(i) = ar_decision(i)
        'El doble
        For j = primero_bis To ultimo_bis
            ar_tmp2_dob(j, i) = arr_doble(j, i)
        Next j
    Next i


    s_ordenar_qui_dos ar_tmp2_dec(), ar_tmp2_dob()
    s_ordenar_qui_dos ar_tmp3_dec(), ar_tmp3_dob()

    i = primero
    cont2 = primero
    c3 = medio + 1
    While i <= ultimo
        If ar_tmp2_dec(cont2) <= ar_tmp2_dec(cont2) Then
            ar_decision(i) = ar_tmp2_dec(cont2)
            For j = primero_bis To ultimo_bis
                arr_doble(j, i) = ar_tmp2_dob(j, cont2)
            Next j
            i = i + 1
            cont2 = cont2 + 1
        Else
            ar_decision(i) = ar_tmp3_dec(c3)
            For j = primero_bis To ultimo_bis
                arr_doble(j, i) = ar_tmp3_dob(j, c3)
            Next j
            i = i + 1
            c3 = c3 + 1
        End If
    Wend


End Sub


Function f_multiline2array(lista As String, mi_array() As String) As Integer

    'Recibe una variable de tipo texto con el contenido de una caja de texto de VB
    'de tipo multiline que contiene lineas separadas por vbCrLf
    'y devuelve en un array las lineas de esa caja, retornando
    'el numero de lineas
    
    'elimina las lineas vacias
    Dim cont_lineas As Integer
    Dim pos As Integer
    Dim linea As String
    

    cont_lineas = 0
    While Len(lista) > 0
        pos = InStr(lista, vbCrLf)
        If pos = 0 Then
            linea = lista
            lista = ""
        Else
            linea = Left(lista, pos - 1)
            lista = Right(lista, Len(lista) - pos - 1)
            pos = InStr(linea, Chr$(10))
            If pos = 1 Then
                linea = Right(linea, Len(linea) - 1)
            End If
        End If
        linea = Trim(linea)
        If Len(linea) > 0 And Left(linea, 1) <> "'" Then
            cont_lineas = cont_lineas + 1
            ReDim Preserve mi_array(1 To cont_lineas) As String
            mi_array(cont_lineas) = linea
        End If
    Wend

    f_multiline2array = cont_lineas
    
End Function


'
'Realiza una copia del Array1 sobre el Array2, ambos
'arrays de longs de dimensión 2
'
Sub s_CopiarArrayLng2Dim(Array1() As Long, Array2() As Long)
  
    'Declaración de variables
     Dim L_n As Long
     Dim L_x As Long

    'Redimensionamos el Array2
     ReDim Array2(LBound(Array1, 1) To UBound(Array1, 1), LBound(Array1, 2) To UBound(Array1, 2)) As Long

    'Bucle
     For L_n = LBound(Array1, 1) To UBound(Array1, 1)
     For L_x = LBound(Array1, 2) To UBound(Array1, 2)
        Array2(L_n, L_x) = Array1(L_n, L_x)
     Next L_x
    Next L_n

End Sub

'
'Realiza una copia del Array1 sobre el Array2, ambos
'arrays de strings de dimensión 2
'
Sub S_CopiarArrayStr2Dim(Array1() As String, Array2() As String)
  
    'Declaración de variables
     Dim L_n As Long
     Dim L_x As Long

    'Redimensionamos el Array2
     ReDim Array2(LBound(Array1, 1) To UBound(Array1, 1), LBound(Array1, 2) To UBound(Array1, 2)) As String

    'Bucle
     For L_n = LBound(Array1, 1) To UBound(Array1, 1)
     For L_x = LBound(Array1, 2) To UBound(Array1, 2)
        Array2(L_n, L_x) = Array1(L_n, L_x)
     Next L_x
    Next L_n

End Sub


Function f_indice_elemento_menor_array_s(mi_array() As String) As Integer
    
    Dim i As Integer
    Dim s_menor As String
    Dim pos_menor As Integer

    For i = 1 To UBound(mi_array)
        If mi_array(i) <> "" Then
            pos_menor = i
            s_menor = mi_array(i)
            Exit For
        End If
    Next i
    For i = pos_menor To UBound(mi_array)
        If mi_array(i) <> "" Then
            If mi_array(i) < s_menor Then
                s_menor = mi_array(i)
                pos_menor = i
            End If
        End If
    Next i
    f_indice_elemento_menor_array_s = pos_menor
    
End Function

Sub s_quitar_lineas_repetidas_en_array(mi_array() As String)

    Dim cont_destino As Integer
    Dim cont_origen As Integer
    
    If LBound(mi_array, 1) >= UBound(mi_array, 1) Then Exit Sub

    cont_destino = 1
    For cont_origen = LBound(mi_array, 1) + 1 To UBound(mi_array, 1)
        If mi_array(cont_destino) <> mi_array(cont_origen) Or mi_array(cont_destino) = "" Then
            cont_destino = cont_destino + 1
            mi_array(cont_destino) = mi_array(cont_origen)
        End If
        DoEvents
    Next cont_origen

    ReDim Preserve mi_array(1 To cont_destino) As String

End Sub

