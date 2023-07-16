Attribute VB_Name = "bas_z0_func"
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

Function f_mayor(a, B, c) As Integer

    If a > B Then
        If a > c Then
            f_mayor = 1
        Else
            f_mayor = 3
        End If
    Else
        If B > c Then
            f_mayor = 2
        Else
            f_mayor = 3
        End If
    End If

End Function

Function f_buscar_en_array(mi_array() As String, elemento As String) As Integer

    Dim i As Integer
    Dim encontrado As Boolean
    Dim indice_encontrado As Integer
    
    encontrado = False
    
    For i = 1 To UBound(mi_array)
        If Trim$(mi_array(i)) = Trim$(elemento) Then
            encontrado = True
            indice_encontrado = i
            Exit For
        End If
    Next i

    If encontrado Then
        f_buscar_en_array = indice_encontrado
    Else
        f_buscar_en_array = -1
    End If

End Function

Function f_buscar_en_lista(S_Literal As String, lista As Control, I_Posicion As Long, I_Longitud As Long) As Double

    'Control del error producido por Visual Basic
     On Error GoTo Error_BuscarEnLista
    

     Dim i            As Long
     Dim I_Encontrado As Double

    'Inicializamos las variables necesarias
     I_Encontrado = 0

    'Bucle de busqueda del registro solicitado en la lista
     For i = 0 To lista.ListCount - 1
         If (I_Posicion = 0 Or I_Longitud = 0 And Trim$(lista.List(i)) = Trim$(S_Literal)) Or Trim$(Mid$(lista.List(i), I_Posicion, I_Longitud)) = Trim$(S_Literal) Then
             I_Encontrado = i
             Exit For
         End If
     Next i

    'Inicializamos el valor devuelto por la busqueda al valor devuelto por la funcion
     f_buscar_en_lista = I_Encontrado
     Exit Function

Error_BuscarEnLista:
    'Inicializamos el valor devuelto por la funcion a 0
     f_buscar_en_lista = 0
     Exit Function
End Function
Sub s_campo_seleccionado(MiCampo As Control)
    
    'Posicionamos el cursor al principio del contenido del campo
     MiCampo.SelStart = 0
    'Resaltamos la longitud total del campo
     MiCampo.SelLength = Len(Trim$(MiCampo.Text))

End Sub

Sub s_centrar_ventana_ejv(ventana As Form)

    'Solo los form a tamaño normal puden ajustarse
    'si esta maximizado o minimizado da error
    'Ponemos el estado a Normal
    ventana.WindowState = CTE_NORMAL

    If ventana.MDIChild = True Then
        'Calculamos la altura del MDIChild, con respecto al MDI
        ventana.Top = (frm_z0_mdi.ScaleHeight - ventana.Height) / 2
        'Calculamos la longitud del MDIChild conrespecto al MDI
        ventana.Left = (frm_z0_mdi.ScaleWidth - ventana.Width) / 2
    Else
        'Calculamos la altura a mostrar la pantalla
        ventana.Top = (Screen.Height - ventana.Height) / 2
        'Calculamos la dimension de la pantalla a mostrar
        ventana.Left = (Screen.Width - ventana.Width) / 2
    End If

End Sub
Function f_comillas(linea As String) As String

    f_comillas = """" & linea & """"
    
End Function

Function f_quitar_comillas_dobles(linea As String)
    
    Dim dev As String
    
    dev = linea
    While Left(dev, 1) = """"
        dev = Right(dev, Len(dev) - 1)
    Wend
    While Right(dev, 1) = """"
        dev = Left(dev, Len(dev) - 1)
    Wend
    
    f_quitar_comillas_dobles = dev
    
End Function
Function f_espacios_derecha_td(ByVal cadena As String, ByVal max_long_texto As Integer, ByVal max_long_total As Integer) As String
    
    cadena = Left(cadena, max_long_texto)
    
    'Trunca por la derecha
    While Len(cadena) < max_long_total
        cadena = cadena & " "
    Wend

    f_espacios_derecha_td = cadena
    

End Function

Function f_espacios_izquierda_td(ByVal cadena As String, ByVal max_long_texto As Integer, ByVal max_long_total As Integer) As String
    
    cadena = Left(cadena, max_long_texto)
    
    'Trunca por la derecha
    While Len(cadena) < max_long_total
        cadena = " " & cadena
    Wend

    f_espacios_izquierda_td = cadena
    

End Function

Function f_espacios_izquierda(ByVal cadena As String, ByVal max_long As Integer) As String

    While Len(cadena) < max_long
        cadena = " " & cadena
    Wend

    f_espacios_izquierda = cadena

End Function
Function f_ceros_izquierda(ByVal cadena As String, ByVal max_long As Integer) As String

    While Len(cadena) < max_long
        cadena = "0" & cadena
    Wend

    f_ceros_izquierda = cadena

End Function


Function redondear_d(i As Double) As Integer

    'Trunca por encima o por debajo

    Dim i_entero As Integer
    Dim d_entero As Double
    i_entero = Int(i)
    d_entero = CDbl(i_entero)
    If (Abs(d_entero - i)) > 0.5 Then
        redondear_d = i_entero + 1
    Else
        redondear_d = i_entero
    End If

End Function


Function f_ocurrencias_cadena(cadena As String, c_buscar As String) As Integer
    
    Dim tmp As String
    Dim cont As Integer
    Dim pos As Integer
    
    'Cuenta el número de veces que aparece una subcadena en una cadena
    If Len(c_buscar) > Len(cadena) Then
        f_ocurrencias_cadena = 0
    Else
        If Len(c_buscar) = Len(cadena) Then
            If c_buscar = cadena Then
                f_ocurrencias_cadena = 1
            Else
                f_ocurrencias_cadena = 0
            End If
        Else
            tmp = cadena
            cont = 0
            pos = InStr(tmp, c_buscar)
            While pos > 0
                tmp = Right(tmp, Len(tmp) - pos)
                cont = cont + 1
                pos = InStr(tmp, c_buscar)
            Wend
            f_ocurrencias_cadena = cont
        End If
    End If

End Function

Function f_elemento_listacomas(lista As String, pos As Long) As String

    Dim tmp As String
    Dim cont As Long
    Dim lugar As Long
    Dim es_el_ultimo As Boolean
    Dim se_acabo As Boolean
    
    
    tmp = lista
    'Cont son los que he quitado
    cont = 0
    se_acabo = False
    es_el_ultimo = False
    While cont < pos - 1 And Not se_acabo
        'Quito todo hasta la primera coma
        lugar = InStr(tmp, ",")
        tmp = Right(tmp, Len(tmp) - lugar)
        cont = cont + 1
        If es_el_ultimo Then
            se_acabo = True
        End If
        If InStr(tmp, ",") = 0 Then
            es_el_ultimo = True
        End If
    Wend

    If InStr(tmp, ",") = 0 Then
        If se_acabo Then
            f_elemento_listacomas = ""
        Else
            'Es el ultimo
            f_elemento_listacomas = tmp
        End If
    Else
        f_elemento_listacomas = Left(tmp, InStr(tmp, ",") - 1)
    End If

End Function

Function f_SumCirc(ByVal maximo As Long, ByVal valor1 As Long, ByVal valor2 As Long) As Long

    'Suma de forma circular dos numeros
    'Siempre devuelve un numero entre 1 y maximo
    'por ejemplo,
    'f_SumCirc(8, 5, 5)
    '5+5 = 10 = 8+2 => devuelve 2
    '1 -> 1
    '2 -> 2
    '3 -> 3
    '4 -> 4
    '5 -> 5
    '6 -> 6
    '7 -> 7
    '8 -> 8
    '9 -> 1
    '10 -> 2
    '...
    
    'Es lo mismo hacer
    
    'Dim suma As Integer
    'suma = valor1 + valor2
    'While suma > maximo
    '    suma = suma - maximo
    'Wend
    'f_SumCirc = suma

    Dim suma As Long
    suma = ((valor1 + valor2) Mod maximo)
    If suma = 0 Then suma = maximo
    f_SumCirc = suma


End Function

Function f_repetir_cadena(cadena As String, separador As String, veces As Long) As String

    Dim i As Long
    
    f_repetir_cadena = ""
    For i = 1 To veces
        If i <> veces Then
            f_repetir_cadena = f_repetir_cadena & cadena & separador
        Else
            f_repetir_cadena = f_repetir_cadena & cadena
        End If
    Next i

End Function


Function f_contar_elementos(cadena As String, separador As String) As String

    Dim mi_cad As String
    Dim c As String * 1
    Dim estado As Integer
    
    If Len(separador) <> 1 Then
        s_error_ejv CON_OPCION_FINALIZAR, "Error en f_contar_elementos"
        Exit Function
    End If
    
    estado = 0
    f_contar_elementos = 0
    mi_cad = cadena
    While Len(mi_cad) > 0
        c = Left(mi_cad, 1)
        mi_cad = Right(mi_cad, Len(mi_cad) - 1)
        Select Case estado
            Case 0 'Inicio
                If c <> separador Then
                    f_contar_elementos = f_contar_elementos + 1
                    estado = 1
                End If
            Case 1 'Primer caracter leido
                If c = separador Then
                    estado = 0
                End If
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error: "
       End Select
    Wend

End Function

Function f_quitar_tags(linea As String) As String

    Dim temp As String
    Dim c As String * 1
    Dim estado As Integer
    Dim salida As String
    
    estado = 0
    temp = linea
    While Len(temp) > 0
        c = Left(temp, 1)
        temp = Right(temp, Len(temp) - 1)
        If estado = 0 Then
            If c = "<" Then
                estado = 1
            Else
                salida = salida & c
            End If
        Else
            If c = ">" Then
                estado = 0
            End If
        End If
    Wend

    f_quitar_tags = salida

End Function

Function f_quitar_caracteres(linea As String, car As String, sustituir As String) As String

    'Busca todos los car en linea
    'si lo encuentra, lo sustituye por sustituir
    'para que solo lo quite, poner "" en sustituir

    Dim temp As String
    Dim c As String * 1
    Dim salida As String
    Dim num_car As Integer
    Dim i As Integer
    Dim hay_que_quitar As Boolean
    
    num_car = Len(car)
    temp = linea
    While Len(temp) > 0
        c = Left(temp, 1)
        temp = Right(temp, Len(temp) - 1)
        hay_que_quitar = False
        For i = 1 To num_car
            If c = Mid(car, i, 1) Then
                hay_que_quitar = True
                salida = salida & sustituir
            End If
        Next i
        If Not hay_que_quitar Then
            salida = salida & c
        End If
    Wend

    f_quitar_caracteres = salida

End Function


Function f_sustituir_subcadena(linea As String, buscar As String, sustituir As String) As String
    
    Dim pos As Integer
    Dim izq As String
    Dim der As String
    Dim temp As String

    temp = linea
    While InStr(temp, buscar) <> 0
        pos = InStr(temp, buscar)
        izq = Left(temp, pos - 1)
        der = Right(temp, Len(temp) - pos - Len(buscar) + 1)
        temp = izq & sustituir & der
    Wend
    f_sustituir_subcadena = temp
    
End Function
Function bytebin2int(s_byte As String) As Integer

    Dim i As Integer

    bytebin2int = 0
    
    For i = 1 To 8
        If Mid(s_byte, i, 1) = "1" Then
            bytebin2int = bytebin2int + 2 ^ (8 - i)
        End If
    Next i

End Function

Function bytehex2bin(s_byte As String) As String
    
    Dim i As Integer


    s_byte = f_ceros_izquierda(s_byte, 2)
    bytehex2bin = ""
    
    For i = 1 To 2
        Select Case Mid(s_byte, i, 1)
            Case "0"
                bytehex2bin = bytehex2bin & "0000"
            Case "1"
                bytehex2bin = bytehex2bin & "0001"
            Case "2"
                bytehex2bin = bytehex2bin & "0010"
            Case "3"
                bytehex2bin = bytehex2bin & "0011"
            Case "4"
                bytehex2bin = bytehex2bin & "0100"
            Case "5"
                bytehex2bin = bytehex2bin & "0101"
            Case "6"
                bytehex2bin = bytehex2bin & "0110"
            Case "7"
                bytehex2bin = bytehex2bin & "0111"
            Case "8"
                bytehex2bin = bytehex2bin & "1000"
            Case "9"
                bytehex2bin = bytehex2bin & "1001"
            Case "A"
                bytehex2bin = bytehex2bin & "1010"
            Case "B"
                bytehex2bin = bytehex2bin & "1011"
            Case "C"
                bytehex2bin = bytehex2bin & "1100"
            Case "D"
                bytehex2bin = bytehex2bin & "1101"
            Case "E"
                bytehex2bin = bytehex2bin & "1110"
            Case "F"
                bytehex2bin = bytehex2bin & "1111"
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
        End Select
    Next i

End Function

Function mi_xor(byte1 As Byte, byte2 As Byte) As Byte

    Dim s_byte1 As String
    Dim s_byte2 As String
    Dim s_byte_out As String
    Dim i As Integer
    
    s_byte1 = Hex(CInt(byte1))
    s_byte2 = Hex(CInt(byte2))
    
    s_byte1 = bytehex2bin(s_byte1)
    s_byte2 = bytehex2bin(s_byte2)
    
    s_byte_out = ""
    For i = 1 To 8
        If Mid(s_byte1, i, 1) = "1" Xor Mid(s_byte2, i, 1) = "1" Then
            s_byte_out = s_byte_out & "1"
        Else
            s_byte_out = s_byte_out & "0"
        End If
    Next i
    mi_xor = CByte(bytebin2int(s_byte_out))

End Function

Function positivo(numero As Double) As Double
    If numero > 0 Then
        positivo = numero
    Else
        positivo = -numero
    End If

End Function
