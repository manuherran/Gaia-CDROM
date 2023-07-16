Attribute VB_Name = "bas_z0_azar"
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

Function f_rnd() As Double

    'Devuelve un numero entre 0 y 1 no incluidos
    'osea, desde
    '0,000000000000000..1
    'hasta
    '0,999999999999999...
    'y todos con igual probabilidad, es decir, es una distibución de
    'azar uniforme
    
    Dim i As Integer
    
    'Devuelvo los valores 0,1  0,2  0,3 hasta 0,9
    'Esto es para probar que pasa si se usa una funcion de azar mala
    'facilmente predecible. Todas las funciones de azar son en realidad
    'pseudoaleatorias, ya que no son realmente azar, todas son mas o menos
    'predecibles, asi que toda serie de numeros contiene azar en cierto grado
    'aunque en realidad ninguna lo es...
    '...hasta ahora sólo la fisica cuantica afirma haber detectado
    'la existencia de una verdadera funcion de azar no predecible,
    'en nuestro universo. Si esto fuera cierto, las implicaciones
    'serían asombrosas: algo en nuestro universo sin una causa, cuyo
    'comportamiento no depende de nada de nuestro universo. La puerta
    'a aquello que queda "fuera de nuestro universo" y otro indicio mas
    'de que nuestro universo puede ser entendido como algo parecido a
    'un ordenador que esta realizando una simulación de "Vida Artificial"
    'como las de este programa
    
    Select Case tipo_funcion_azar_ejv
        Case CTE_AZARVB
            f_rnd = Rnd
        Case CTE_AZARFIC
            'Uso los CTE_NUM_DECIMALES_EXACTITUD decimales de pi
            'a partir de cierta posición
            'y pongo un numero que es "0,esos decimales"
            f_rnd = 0
            For i = 1 To CTE_NUM_DECIMALES_EXACTITUD
                f_rnd = f_rnd + (digitos_azar(indice_azar) / (10 ^ i))
                'Actualizo el indice que apunta al azar
                indice_azar = f_SumCirc(azar_fichero_num_char_ejv, indice_azar, 1)
            Next i
            'Una vez ejecutado el bucle,
            'he saltado CTE_NUM_DECIMALES_EXACTITUD digitos, los que he usado como decimales
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Generador de Pseudoaleatorios no existente"
    End Select

End Function

Function f_gauss_m1(LI As Long, LS As Long) As Double

'------------------------------------------------------------------------
'Devuelve un numero al azar entre LI y LS
'segun una distribución normal o gaussiana (es lo mismo)
'con una cierta media y desviación tipica
'con forma de campana de gauss 1 creo
'de forma que devuelve con mayor probabilidad valores
'cercanos a la media que hemos fijado
'------------------------------------------------------------------------
'Metodo 1
'Propuesto por Phil Beffrey phil@pixar.com
'y Mattias Fagerlund matfa@acacia.se
'Thanks a lot!!!
'Calcular la media de varios numeros generados al azar con una distribucion uniforme
'bastan unos 12 ejemplos
'------------------------------------------------------------------------

    Dim suma_azar As Double
    Dim i As Integer
    
    suma_azar = 0
    For i = 1 To 12
        suma_azar = suma_azar + fl_azar2(LI, LS)
    Next i
    suma_azar = suma_azar / 12
    f_gauss_m1 = suma_azar
    
    
End Function

Function f_gauss_m2(LI As Double, LS As Double, media As Double, sigma As Double) As Double
    
    'Devuelve un numero al azar entre LI y LS
    'segun una distribución normal o gaussiana (es lo mismo)
    'con una cierta media y desviación tipica
    'con forma de campana de gauss
    'de forma que devuelve con mayor probabilidad valores
    'cercanos a la media que hemos fijado

    Dim suma_azar As Double
    Dim devolver As Double
    Dim azar As Double
    Dim suma As Double
    Dim valor1 As Double
    Dim valor2 As Double
    Dim factor As Double
    Dim i As Integer

    'Metodo 2
    
    'Posibles errores al traducir de C
    'Supongo que rand() devuelve entre 0 y RAND_MAX pero no se incluye el 0 ni el RAND_MAX
    'a mi me falta la funcion que devuelva entre a y b pero no enteros sino reales
    'Supongo que sqrt es sqr de VB
    
    'U1 = rand() / RAND_MAX
    'U2 = rand() / RAND_MAX
    'x1=2*(double)rand()/RAND_MAX-1;
    'x2=2*(double)rand()/RAND_MAX-1;
    
    'La media es 0
    'y la desviacion tipica es 1 o sigma
    
    'Obtenido de Box-Muller method
    'gracias a Antonio Carpintero Sanchez amcarpintero@polar.es
    
    If hay_solucion_anterior_gauss Then
        devolver = solucion_anterior_gauss
    Else
        suma = 1
        While suma >= 1
            valor1 = 2 * f_rnd - 1
            valor2 = 2 * f_rnd - 1
            suma = valor1 * valor1 + valor2 * valor2
        Wend
        factor = sigma * Sqr(-2 * Log(suma) / suma)
        devolver = valor1 * factor + media
        solucion_anterior_gauss = valor2 * factor + media
    End If
    hay_solucion_anterior_gauss = Not hay_solucion_anterior_gauss
    f_gauss_m2 = devolver
    
'===================================================
'
'Obtenido de:
'author = "William H. Press and Brian P. Flannery and Saul A. Teukolsky and William T. Vetterling",
'Title = "{NUMERICAL RECIPES} The Art Of Scientific Computing ({FORTRAN}Version)"
'publisher="Cambridge University Press"
'Year = "1989"
'Gracias a
'Evan Hughes E.J.Hughes@rmcs.cranfield.ac.uk
'http://www.rmcs.cranfield.ac.uk/~daps/daps/pgrads/e_h/hughes.htm
'    suma = 0
'    While suma = 0 Or suma >= 1
'        valor1 = 2 * f_rnd - 1
'        valor2 = 2 * f_rnd - 1
'        suma = valor1 * valor1 + valor2 * valor2
'    Wend
'    factor = sigma * sqrt(-2 * Log(suma) / suma)
'    solucion1 = valor1 * factor + media
'    solucion2 = valor2 * factor + media
'
'===================================================
    'Para obligar a unos limites tendre que hacer un bucle
    'que la llame constantemente hasta que caiga ahi

    

End Function
   

Function fi_azar1(fin As Integer) As Integer

    'Devuleve un número al azar entero
    'entre 1 y fin, inclusive ambos
    'y la probabilidad es igual para todos los números
    
    'Ojo es int (parte entera) y no cint (convertir a integer)
    fi_azar1 = Int(f_rnd * fin) + 1
   

End Function

Function fi_azar2(inicio As Integer, fin As Integer) As Integer

    'Devuleve un número al azar entero
    'entre Inicio y fin, inclusive ambos
    'y la probabilidad es igual para todos los números
    'y pueden ser valores negativos
    fi_azar2 = fi_azar1(fin - inicio + 1) + inicio - 1

End Function

Function fl_azar1(fin As Long) As Long

    'Devuleve un número al azar entero
    'entre inicio y fin, inclusive ambos
    'y la probabilidad es igual para todos los números
    'Ojo es int (parte entera) y no cint (convertir a integer)
    fl_azar1 = Int(f_rnd * fin) + 1
    

End Function

Function fl_azar2(inicio As Long, fin As Long) As Long

    'Devuleve un número al azar entero
    'entre inicio y fin, inclusive ambos
    'y la probabilidad es igual para todos los números
    'y pueden ser valores negativos
    fl_azar2 = fl_azar1(fin - inicio + 1) + inicio - 1


End Function


Function fi_azar4(p1 As Integer, p2 As Integer, p3 As Integer, p4 As Integer) As Integer

    'Devuelve un número al azar entero: 1 2 3 o 4
    'con una probabilidad asignada por los parametros
    '
    'Elegir un número al azar de un conjunto de 4 numeros
    'donde las probabilidades de que aparezca cada uno son:
    '1: 20
    '2: 60
    '3: 10
    '4: 10
    'ej:
    'fi_Azar4(20, 60, 10, 10)
    
    
    Dim azar As Integer
    Dim tope1 As Integer
    Dim tope2 As Integer
    Dim tope3 As Integer
    Dim tope4 As Integer
    
    tope1 = p1
    tope2 = tope1 + p2
    tope3 = tope2 + p3
    tope4 = tope3 + p4
    
    If tope4 <> 100 Then
        s_error_ejv CON_OPCION_FINALIZAR, "Llamada a la función de azar incorrecta. La suma de las probabilidades no es 100."
    End If
    
    azar = fi_azar1(100)
    
    If azar < tope1 Then
        fi_azar4 = 1
    Else
        If azar < tope2 Then
            fi_azar4 = 2
        Else
            If azar < tope3 Then
                fi_azar4 = 3
            Else
                fi_azar4 = 4
            End If
        End If
    End If


End Function
Function fl_AzarRangos(num_probb As Integer, p() As Long, suma As Long) As Integer

    'Devuelve un número al azar entero de 1 a num_probb
    'con una probabilidad asignada por los parametros p
    '
    'Ejemplo:
    'Elegir un número al azar de un conjunto de 4 numeros
    'donde las probabilidades de que aparezca cada uno son:
    '1: 20
    '2: 60
    '3: 10
    '4: 10
    'ej:
    'fi_Azar4(20, 60, 10, 10)
    'pues devolvera 1 2 3 o 4, probablemente devolvera un 2
    
    'La suma de todas las probabilidades debe dar suma
    
    
    Dim azar As Long
    Dim i As Integer
    ReDim tope(1 To num_probb) As Long
    
    
    If num_probb = 1 Then
        fl_AzarRangos = 1
        Exit Function
    End If
    
    tope(1) = p(1)
    
    For i = 2 To num_probb
        tope(i) = tope(i - 1) + p(i)
    Next i
    
    If tope(num_probb) <> suma Then
        s_error_ejv CON_OPCION_FINALIZAR, "Llamada a la función de azar incorrecta. La suma de las probabilidades no es " & suma
    End If
    
    azar = fl_azar1(suma)
    
    
    i = 1
    While azar >= tope(i) And i < num_probb
        i = i + 1
    Wend
    fl_AzarRangos = i


End Function

Function f_analizar_probabilidad_ejv(probb As Double) As Boolean

    Dim temp As Double
    Dim cont_desplaz As Integer

    'Paso a tanto por uno
    temp = probb * 100
    cont_desplaz = 2
    'Desplazo todos los decimales hasta que quede un numero de 0-100
    'Pero como mucho uso 5 decimales de exactitud, que no esta nada mal
    'si se tiene en cuenta que luego va de 1-100
    While CLng(temp) <> temp And cont_desplaz < 6
        cont_desplaz = cont_desplaz + 1
        temp = temp * 10
    Wend
    If fl_azar1(10 ^ cont_desplaz) <= CLng(temp) Then
        f_analizar_probabilidad_ejv = True
    Else
        f_analizar_probabilidad_ejv = False
    End If

End Function
