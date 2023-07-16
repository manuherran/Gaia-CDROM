Attribute VB_Name = "bas_a6_simples"
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

Function Fl_AumentarObjetivo(objetivo As Long, grado As Long) As Long

'CTE_Max_Long = 2147483647
'CTE_Min_Long = -2147483647
 
'b: lim max
'a: valor actual
 
'a1 = a + (b-a)/2
'a1 = (a+b) / 2
'a2 = (a1+b) /2
'...
'ai = {a + (2^i -1)b} / 2
 
 
Dim maximo_valor As Long
Dim minimo_valor As Long
Dim valor_medio As Long
 
 maximo_valor = 2000000000
 minimo_valor = 0
 valor_medio = 1000000000

 Fl_AumentarObjetivo = (objetivo + ((2 ^ grado) - 1) * maximo_valor) / 2


End Function

Function f_BorrarUniversoCompleto(p_Cod_Uni As Long)

    'Si solo queda uno o es el primero, no dejamos borrarlo
    If GL_Num_Uni > 1 And p_Cod_Uni <> 1 Then

    f_BorrarTabla CTE_TABLA_ENTIDAD, p_Cod_Uni, 0
    f_BorrarTabla CTE_TABLA_ACCION, p_Cod_Uni, 0
    f_BorrarTabla CTE_TABLA_MEMORIA, p_Cod_Uni, 0
    f_BorrarTabla CTE_TABLA_PARAMAC, p_Cod_Uni, 0
    f_BorrarTabla CTE_TABLA_REGLA, p_Cod_Uni, 0
    f_BorrarTabla CTE_TABLA_CONTEXTO, p_Cod_Uni, 0
    f_BorrarTabla CTE_TABLA_CONCLUSION, p_Cod_Uni, 0
    f_BorrarTabla CTE_TABLA_VBLE, p_Cod_Uni, 0
   'aparte
    f_BorrarTabla CTE_TABLA_UNIVERSO, p_Cod_Uni, 0

   'Reorganizamos los universos: al ultimo le damos el valor del que falta
    
   'Modificar GLOBAL
    GL_Num_Uni = GL_Num_Uni - 1
    If GL_Cod_Uni > GL_Num_Uni Then
        GL_Cod_Uni = 1
    End If
    S_EscribirGLOBAL

    End If

End Function


Function f_BorrarEntidadCompleta(p_Cod_Uni As Long, p_Cod_Ent)
    
    Dim num_ent As Long
    
    f_BorrarTabla CTE_TABLA_ENTIDAD, p_Cod_Uni, p_Cod_Ent
    f_BorrarTabla CTE_TABLA_ACCION, p_Cod_Uni, p_Cod_Ent
    f_BorrarTabla CTE_TABLA_MEMORIA, p_Cod_Uni, p_Cod_Ent
    f_BorrarTabla CTE_TABLA_PARAMAC, p_Cod_Uni, p_Cod_Ent
    f_BorrarTabla CTE_TABLA_REGLA, p_Cod_Uni, p_Cod_Ent
    f_BorrarTabla CTE_TABLA_CONTEXTO, p_Cod_Uni, p_Cod_Ent
    f_BorrarTabla CTE_TABLA_CONCLUSION, p_Cod_Uni, p_Cod_Ent
    f_BorrarTabla CTE_TABLA_VBLE, p_Cod_Uni, p_Cod_Ent

   'Modificamos los datos del universo
    num_ent = fl_Leer_Num_Ent(p_Cod_Uni)
   'Al borrar la entidad puede que el universo se haya quedado vacío
   'pero entonces no se ejecutará y no pasa nada
    'fl_Escribir_Num_Ent p_Cod_Uni, num_ent - 1
    fl_Escribir_Num_Ent p_Cod_Uni
   
   
   
   'Reorganizamos las entidades de ese universo
   'Pasamos la ultima a la posición de la actual
'    f_Cambiar_Cod_Ent p_Cod_Uni, p_Cod_Ent, p_Cod_Ent_Nuevo
   
   'Puede que hayamos borrado la entidad actual
   'entonces pasamos a la siguiente
    If p_Cod_Uni = GL_Cod_Uni And GL_Cod_Ent = p_Cod_Ent Then
        If p_Cod_Ent = GL_Cod_Ent Then
        'Si es la ultima la que hemos borrado, ponemos
        'la actual como la primera
        GL_Cod_Ent = 1
        'en caso contrario, como se ha reorganizado
        'no hace falta
        End If
    End If
   
    
   
   



End Function

Function f_BorrarTabla(p_TABLA As String, p_Cod_Uni As Long, p_Cod_Ent)



    'Miramos si hay que borrar solo una entidad o todas las de ese universo
     Dim SQL As String
     If p_Cod_Ent = 0 Then
        SQL = "SELECT * FROM " & p_TABLA & " WHERE " & CTE_ENTIDAD_Cod_Uni & " = " & Str$(p_Cod_Uni)
     Else
        SQL = "SELECT * FROM " & p_TABLA & " WHERE " & CTE_ENTIDAD_Cod_Uni & " = " & Str$(p_Cod_Uni) & " AND " & CTE_ENTIDAD_Cod_Ent & " = " & Str$(p_Cod_Ent)
     End If

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_BajaN
    'Esqueleto de SELECT elegido
     GS_BD_SQL = SQL
    'Acceso a la base de datos
     S_AccesoBD
    'Liberamos espacio en memoria datos de entrada
     ReDim GL_AR_BD_DatosEntrada(0) As Long
    'Control de error
     If GS_BD_Error <> CTE_ErrorNinguno Then
        'Tratamiento error acceso BD
         Beep
        'Visualizamos el error producido por el desarrollo de la funcion
         MsgBox ("Num Error: " & Err & ". Texto: " & error & ". Gaia: " & GS_BD_Error & ".")
        'Finalizamos la aplicación
         End
     End If


End Function

Sub S_EjecutarAccionSimple()
'
'Un acción simple puede ser una llamada a una función de Visual Basic
'de las que habrá que hacer todas las que se pueda, y crear nuevas
'dependiendo del objetivo del proyecto, o una llamada a otro programa,
'por ejemplo, un exe de MS-DOS ó otro programa de windows, ya se verá
'cómo, pero siempre que su ejecución sea siempre correcta, no demasiado
'larga y cuya finalización no depende de variables del entorno del
'proyecto. Es decir, puede haber while, pero sobre condiciones
'externas, de manera que siempre finalicen los while

'Vamos a ejecutar la acción:
'
'de la que conocemos:
'     GL_Num_Orden
'     GL_Tip
'     GL_Cod_Acc_Padre
'     GL_Num_Repetida_Padre


'La accion se define por cod_ent + cod_acc
'pero si es simple, referencia la entidad a ejecutar por el valor de
'
Select Case GL_Acc_simple
    Case 1
        S_S1_Predecir_Telefonica
    Case 2
        S_S2_Evaluar_Predictor_Telefonica
    Case 3
        S_S3_Predecir_Vallehermoso
    Case 4
        S_S4_Evaluar_Predictor_Vallehermoso
        
        
        
    Case 5
        S_S999_Accion_Simple_Vacia
    Case 6
        S_S999_Accion_Simple_Vacia
    Case 7
        S_S999_Accion_Simple_Vacia
    Case 8
        S_S999_Accion_Simple_Vacia
        
        
        
        
    Case 999
        S_S999_Accion_Simple_Vacia
    Case Else
        MsgBox "Error: Acción simple no existe"
        End
End Select



End Sub

Sub S_S1_Predecir_Telefonica()

'Un predictor tiene las siguientes variables:
'1:objetivo
'2,3,4:valor a predecir (Evento, año, mes)
'5:predicción

'1.- Leer valor a predecir
'2.- Leer reglas que representan lo que opina la entidad sobre el comportamiento
'de la ralidad
'3.- Aplicar reglas y obtener la predicción
'4.- Escribir la predicción

 Dim evento As String
 Dim ano As Integer
 Dim mes As Integer


'1.- Leer valor a predecir
 evento = fv_Leer_Vble(GL_Cod_Uni, GL_Cod_Ent, 2)
 ano = fv_Leer_Vble(GL_Cod_Uni, GL_Cod_Ent, 3)
 mes = fv_Leer_Vble(GL_Cod_Uni, GL_Cod_Ent, 4)
 
 
'2.- Leer reglas que representan lo que opina la entidad sobre el comportamiento
'de la realidad

'Regla completa:
'contexto + acción => conclusión
'
'contexto es un conjunto de var=valor
'accion es una accion simple o compleja
'conclusión es un conjunto de var=formula

'ejemplos
'si veo tales valores ctes y hago tal acción, luego tal variable toma el valor tal cte

'si veo tales valores y hago tal acción, luego tal variable toma el
'valor tal, que es una función de los valores de tal y tal

'acción puede ser 0 (no hace falta hacer nada para que ocurra, es inmediato)

'Parte izq puede ser 0: siempre que hago algo pasa algo

'Parte izq y acción puede ser 0: se define como calcular unas variables
'en función de otras y ya está



'3.- Aplicar reglas y obtener la predicción

'4.- Escribir la predicción


End Sub

Sub S_S2_Evaluar_Predictor_Telefonica()


'. Mirar si el valor predicho ya se conocia con seguridad
'.- Si es conocido,
'5.- Compararlo y en función de eso modificar la variable objetivo


End Sub
Sub S_S4_Evaluar_Predictor_Vallehermoso()

End Sub
Sub S_S6_Evaluar_Jugador()
    
    Dim dinero As Long
    
    'Lee el dinero que tiene
     dinero = fv_Leer_Vble(GL_Cod_Uni, GL_Cod_Ent, 6)
    
    'Lo guarda en la memoria
    
    'Lee de la memoria la evolución del dinero que ha tenido
    
    
    
    

End Sub

Sub S_S999_Accion_Simple_Vacia()

End Sub

Sub S_S3_Predecir_Vallehermoso()

End Sub

