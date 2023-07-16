Attribute VB_Name = "bas_a6_acc"
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
Sub S_Leer_Parametros()

    Dim codigo_var As Long
    Dim i As Long
    ReDim GV_Param(1 To GL_Num_Param) As Variant
    For i = 1 To GL_Num_Param
        codigo_var = fl_Leer_Parametro(GL_Cod_Uni, GL_Cod_Ent, GL_Cod_Acc, GL_Num_Repetida, i)
        GV_Param(i) = fv_Leer_Vble(GL_Cod_Uni, GL_Cod_Ent, codigo_var)
    Next


End Sub

Function fl_Leer_Num_Orden(p_Cod_Uni, p_Cod_Ent, p_Cod_Acc, p_Num_Repetida) As Long

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'Esqueleto de SELECT elegido
     GS_BD_SQL = "SELECT " & CTE_ACCION_Num_Orden & " FROM " & CTE_TABLA_ACCION & " WHERE " & CTE_ACCION_Cod_Uni & " = " & p_Cod_Uni & " AND " & CTE_ACCION_Cod_Acc & " = " & Str$(p_Cod_Acc) & " AND " & CTE_ACCION_Cod_Ent & " = " & Str$(p_Cod_Ent) & " AND " & CTE_ACCION_Num_Repetida & " = " & Str$(p_Num_Repetida)
    'Número de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = 1
    'Acceso a la base de datos
     S_AccesoBD
    'Control de error
     If GS_BD_Error <> CTE_ErrorNinguno Then
            'Tratamiento error acceso BD
             Beep
            'Visualizamos el error producido por el desarrollo de la funcion
             MsgBox ("Num Error: " & Err & ". Texto: " & error & ". Gaia: " & GS_BD_Error & ".")
            'Finalizamos la aplicación
             End
     End If
    
     GS_Des_Acc = GS_BD_DesSalida
    'Tratamiento acceso BD correcto
     fl_Leer_Num_Orden = GL_AR_BD_DatosSalida(0, 0)


End Function

Function fl_Leer_Padre(p_Cod_Uni, p_Cod_Ent, p_Cod_Acc, p_Num_Repetida) As Long

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'Esqueleto de SELECT elegido
     GS_BD_SQL = "SELECT " & CTE_ACCION_Cod_Acc_Padre & " FROM " & CTE_TABLA_ACCION & " WHERE " & CTE_ACCION_Cod_Uni & " = " & p_Cod_Uni & " AND " & CTE_ACCION_Cod_Acc & " = " & Str$(p_Cod_Acc) & " AND " & CTE_ACCION_Cod_Ent & " = " & Str$(p_Cod_Ent) & " AND " & CTE_ACCION_Num_Repetida & " = " & Str$(p_Num_Repetida)
    'Número de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = 1
    'Acceso a la base de datos
     S_AccesoBD
    'Control de error
     If GS_BD_Error <> CTE_ErrorNinguno Then
        If GS_BD_Error = CTE_ErrorCNE Then
             fl_Leer_Padre = 0
        Else
            'Tratamiento error acceso BD
             Beep
            'Visualizamos el error producido por el desarrollo de la funcion
             MsgBox ("Num Error: " & Err & ". Texto: " & error & ". Gaia: " & GS_BD_Error & ".")
            'Finalizamos la aplicación
             End
        End If
     End If
    
     GS_Des_Acc = GS_BD_DesSalida
    'Tratamiento acceso BD correcto
     fl_Leer_Padre = GL_AR_BD_DatosSalida(0, 0)

End Function

Function fl_Leer_Parametro(p_Cod_Uni, p_Cod_Ent, p_Cod_Acc, p_Num_Repetida, p_Cod_Param) As Long


    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'Esqueleto de SELECT elegido
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_PARAMAC & " WHERE " & CTE_PARAMAC_Cod_Uni & " = " & Str$(p_Cod_Uni) & " AND " & CTE_PARAMAC_Cod_Ent & " = " & Str$(GL_Cod_Ent) & " AND " & CTE_PARAMAC_Cod_Acc & " = " & Str$(p_Cod_Ent) & " AND " & CTE_PARAMAC_Num_Repetida & " = " & Str$(p_Num_Repetida) & " AND " & CTE_PARAMAC_Cod_Param & " = " & Str$(p_Cod_Param)
    'Número de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = CTE_N_PARAMAC
    'Array donde se recibe el resultado de la BD (un registro)
     ReDim GL_AR_BD_DatosSalida(0, GI_BD_NCamposConsulta - 1 - CTE_num_campos_des) As Long
    'Acceso a la base de datos
     S_AccesoBD
    'Control de error
     If GS_BD_Error <> CTE_ErrorNinguno Then
        'Tratamiento error acceso BD
         Beep
        'Visualizamos el error producido por el desarrollo de la funcion
         MsgBox ("Num Error: " & Err & ". Texto: " & error & ". Gaia: " & GS_BD_Error & ".")
        'Finalizamos la aplicación
         End
     End If
    
     GS_Des_Acc = GS_BD_DesSalida
    'Tratamiento acceso BD correcto
     fl_Leer_Parametro = GL_AR_BD_DatosSalida(0, 5)
    'Despues del tratamiento del acceso, liberamos la memoria ocupada
    'por el array que contiene los datos de salida de la consulta.
     ReDim GL_AR_BD_DatosSalida(0, 0) As Long


End Function

Sub S_CalcularAccionSgte()


'Es obligatorio que las acciones esten siempre ordenadas desde 1: 1,2,3
'no vale las acciones 2,5,7
'el numero de serie también debe estar ordenado, de forma que todas
'las acciones tienen numero de serie 1 salvo las que se repiten, y
'estas son: la primera en el arbol, serie 1, la segunda 2 etc en el orden
'de ejecución estatico primero el nivel 1, luego el segundo


'*****************************************************************
'viejo:
'1.- Acceder al hermano siguiente al actual, definido
'    por las variables globales
'2.- Ver si existe el hermano siguiente (si es simple)
'3.- Si no existe, o es compleja poner en las variables globales el valor
'    del padre
'nuevo:
'    Si es simple vemos si tiene hermano siguiente
        '- si tiene hermano siguiente, la siguiente es esa
        '- si no, ponemos en las variables globales el valor del padre
'    si es no simple (0,2,...) el sgte es el 1er hijo
'1.- Acceder al hermano siguiente al actual, definido
'    por las variables globales
'2.- Ver si existe el hermano siguiente (si es simple)
'3.- Si no existe, o es compleja poner en las variables globales el valor
'    del padre
'*****************************************************************

'refinitivo
'Al llamar a este proc, las variables globales apuntan a la
'acción actual, de la que ya se han leído sus datos
'Este proc calcula sobre las variables globales, la
'acción siguiente a ejecutar
' 0 compleja: con hijos
' 1 simple
' 2 etc de control
'hay un bucle implicito en la primera serie, pero no een las descomposiciones
'1.- vemos si la actual es simple
'1.1.- Si es simple, vemos si tiene un hermano siguiente
'1.1.1.- Si tiene un hermano siguiente, esta es la siguiente
'1.1.2.- Si no tiene un hermano siguiente,-es la ultima- vemos si es de nivel 1 o superior
'1.1.2.1.- Si es de nivel 1 -sin padre- entonces la siguiente es la primera de todas
'1.1.2.2.- Si tiene padre, entonces la siguiente es la siguiente al padre, y llamamos
'a esta misma función pasandole los datos del padre, pero haciendo que el padre no tenga
'hijos, es decir, diciendo que este padre es simple, sin hijos, aunque no sea cierto
'y de forma que después de hacer eso termine la función
'1.2.- si no es simple, puede ser con hijos o de control
'1.2.1.- Si tiene hijos, entonces la siguiente es el primero de los hijos
'1.2.2.- Si es de control, se ejecuta el control y este nos devuelve la acc siguiente
'*****************************************************************

    'Declaración de variables
     Dim I_Control As Integer
     Dim temp_Cod_Acc_Padre  As Long

    'Inicialización de variables
     I_Control = True
    
    
    
'1.- vemos si la actual es simple
 If GL_Tip = 1 Then
    '1.1.- Si es simple, vemos si tiene un hermano siguiente
    'Leemos el hermano siguiente
'          Para buscar el hermano, sabemos que:
'          CTE_ACCION_Cod_Ent = GL_Cod_Ent  ---> es la misma entidad
'          CTE_ACCION_Num_Orden = GL_Num_Orden +1  ---> es 1 + que el actual
'          ...y su padre es el mismo, así que:
'          CTE_ACCION_Cod_Acc_Padre = GL_Cod_Acc_Padre
'          y con todas estas restricciones, sólo puede aparecer una acción
'          -Es alguien que tiene el mismo padre que yo, es el
'           siguiente a mí, y todo esto dentro de una misma entidad-
'          y así obtenermos el código de acción
    'Pasamos a la siguiente acción
     GL_Num_Orden = GL_Num_Orden + 1
    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'SQL
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_ACCION & " WHERE " & CTE_ACCION_Cod_Uni & " = " & GL_Cod_Uni & " AND " & CTE_ACCION_Cod_Ent & " = " & GL_Cod_Ent & " AND " & CTE_ACCION_Num_Orden & " = " & GL_Num_Orden & " AND " & CTE_ACCION_Cod_Acc_Padre & " = " & GL_Cod_Acc_Padre
    'Número de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = CTE_N_ACCION
    'Array donde se recibe el resultado de la BD (un registro)
     ReDim GL_AR_BD_DatosSalida(0, GI_BD_NCamposConsulta - 1) As Long
    'Acceso a la base de datos
     S_AccesoBD
     If GS_BD_Error = CTE_ErrorNinguno Then
        '1.1.1.- Si tiene un hermano siguiente, esta es la siguiente
            'Tratamiento acceso BD correcto
             S_Leer_Datos_ACCION
            'Despues del tratamiento del acceso, liberamos la memoria ocupada
            'por el array que contiene los datos de salida de la consulta.
             ReDim GL_AR_BD_DatosSalida(0, 0) As Long
     Else
        If GS_BD_Error = CTE_ErrorCNE Then ' si no existia el siguiente
            'La que buscamos seguro que es la primera de su propia secuencia
             GL_Num_Orden = 1
            '1.1.2.- Si no tiene un hermano siguiente,-es la ultima- vemos si es de nivel 1 o superior
            If GL_Cod_Acc_Padre = 0 Then
                '1.1.2.1.- Si es de nivel 1 -sin padre- entonces la siguiente es la primera de todas
                GL_Cod_Acc = 1
                GL_Num_Repetida = 1
                
            Else
                '1.1.2.2.- Si tiene padre, entonces la siguiente es la siguiente al padre, y llamamos
                'a esta misma función pasandole los datos del padre, pero haciendo que el padre
                'sea uno más
                'no hecemos diciendo que este padre es simple, sin hijos, aunque no sea cierto
                'poruque lo lee y se da cuenta de que no.
                'y de forma que después de hacer eso termine la función
                'hacemos como si el padre no tiene hijos
                'GL_Tip = 1
                GL_Num_Repetida = 1 'los padres no se repiten
                temp_Cod_Acc_Padre = GL_Cod_Acc_Padre
                GL_Cod_Acc_Padre = fl_Leer_Padre(GL_Cod_Uni, GL_Cod_Ent, GL_Cod_Acc, 1)
                'GL_Num_Orden = fl_Leer_Num_Orden(GL_Cod_Uni, GL_Cod_Ent, temp_Cod_Acc_Padre, 1)
                GL_Num_Orden = temp_Cod_Acc_Padre
                S_CalcularAccionSgte 'este ya suma 1 a num_orden
            End If
         Else
            'es error de bd
            'Tratamiento error acceso BD
             Beep
            'Visualizamos el error producido por el desarrollo de la funcion
             MsgBox ("Num Error: " & Err & ". Texto: " & error & ". Gaia: " & GS_BD_Error & ".")
            'Finalizamos la aplicación
             End
         End If
    End If
 Else
    '1.2.- si no es simple, puede ser con hijos o de control
    If GL_Tip = 0 Then 'manue
        '1.2.1.- Si tiene hijos, entonces la siguiente es el primero de los hijos
             
            'Base de datos a la que se accede
             GI_BD_NumeroDeBD = 1
            'Operación a realizar: A,B,M,C
             GS_BD_Operacion = CTE_BD_Consulta1
            'SQL
             GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_ACCION & " WHERE " & CTE_ACCION_Cod_Uni & " = " & GL_Cod_Uni & " AND " & CTE_ACCION_Cod_Ent & " = " & GL_Cod_Ent & " AND " & CTE_ACCION_Num_Orden & " = 1 AND " & CTE_ACCION_Cod_Acc_Padre & " = " & GL_Cod_Acc
            'Número de campos de la tabla que se desean consultar
             GI_BD_NCamposConsulta = CTE_N_ACCION
            'Array donde se recibe el resultado de la BD (un registro)
             ReDim GL_AR_BD_DatosSalida(0, GI_BD_NCamposConsulta - 1) As Long
            'Acceso a la base de datos
             S_AccesoBD
             
            'si ha habido algun tipo de error
             If GS_BD_Error <> CTE_ErrorNinguno Then
                'Tratamiento error acceso BD
                 Beep
                'Visualizamos el error producido por el desarrollo de la funcion
                 MsgBox ("Num Error: " & Err & ". Texto: " & error & ". Gaia: " & GS_BD_Error & ".")
                'Finalizamos la aplicación
                 End
             Else
                'Tratamiento acceso BD correcto
                 S_Leer_Datos_ACCION
                'Despues del tratamiento del acceso, liberamos la memoria ocupada
                'por el array que contiene los datos de salida de la consulta.
                 ReDim GL_AR_BD_DatosSalida(0, 0) As Long
            End If
    Else
        '1.2.2.- Si es de control, se ejecuta el control y este nos devuelve la acc siguiente
        S_Ejecutar_Accion_Control
       
    End If
 End If
    
    

End Sub
Sub S_EjecutarPriAcciones()

'1.- Inicialización de variables
     GL_N_Acciones_Ejecutadas = 0
    
'2.- Mientras no se hayan ejecutado todas...
     While GL_N_Acciones_Ejecutadas < GL_Ent_Pri

'3.- Accedemos a la acción a ejecutar, definida por
    'GL_Cod_Ent + GL_Cod_Acc + GL_Num_Repetida
    'y tomamos sus datos sobre las variables globales
     S_Leer_ACCION

'4.- Vemos si es simple
     If GL_Tip = 1 Then 'es simple 1
        'Es simple
         S_EjecutarAccionSimple
     End If

'6.- Contador
     GL_N_Acciones_Ejecutadas = GL_N_Acciones_Ejecutadas + 1

'8.- Informamos al usuario
    If GI_modo_de_ejecucion >= 3 Then
       S_MostrarDatosUniverso
       S_MostrarDatosEntidad
       S_MostrarDatosAccion
    End If
    'Informamos al usuario de la entidad que se acaba de ejecutar
     If GI_modo_de_ejecucion >= 4 Then
         If MsgBox("Acción " & GL_Cod_Acc & " Ejecutada. ¿Mantener el modo de ejecución?", 1) = 2 Then GI_modo_de_ejecucion = GI_modo_de_ejecucion - 1
     End If

'7.- Calcular la siguiente acción en el arbol
     S_CalcularAccionSgte
     '"Ejecutar" una accion compleja o una de control consiste solo en calcular cual es la
     'siguiente accion a ejecutar, aunque cuenta como una ejecutada en pri

'9.- Vemos si hay que finalizar
     Wend


End Sub
Sub S_Leer_ACCION()

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'Esqueleto de SELECT elegido
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_ACCION & " WHERE " & CTE_ACCION_Cod_Uni & " = " & GL_Cod_Uni & " AND " & CTE_ACCION_Cod_Acc & " = " & Str$(GL_Cod_Acc) & " AND " & CTE_ACCION_Cod_Ent & " = " & Str$(GL_Cod_Ent) & " AND " & CTE_ACCION_Num_Repetida & " = " & Str$(GL_Num_Repetida)
    'Número de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = CTE_N_ACCION
    'Array donde se recibe el resultado de la BD (un registro)
     ReDim GL_AR_BD_DatosSalida(0, GI_BD_NCamposConsulta - 1 - CTE_num_campos_des) As Long
    'Acceso a la base de datos
     S_AccesoBD
    'Control de error
     If GS_BD_Error <> CTE_ErrorNinguno Then
        'Tratamiento error acceso BD
         Beep
        'Visualizamos el error producido por el desarrollo de la funcion
         MsgBox ("Num Error: " & Err & ". Texto: " & error & ". Gaia: " & GS_BD_Error & ".")
        'Finalizamos la aplicación
         End
     End If
    
     GS_Des_Acc = GS_BD_DesSalida
    'Tratamiento acceso BD correcto
     S_Leer_Datos_ACCION
    'Despues del tratamiento del acceso, liberamos la memoria ocupada
    'por el array que contiene los datos de salida de la consulta.
     ReDim GL_AR_BD_DatosSalida(0, 0) As Long



End Sub

Sub S_Leer_Datos_ACCION()

    GS_Des_Acc = GS_BD_DesSalida

    GL_Cod_Uni = GL_AR_BD_DatosSalida(0, 0)
    GL_Cod_Ent = GL_AR_BD_DatosSalida(0, 1)
    GL_Cod_Acc = GL_AR_BD_DatosSalida(0, 2)
    GL_Num_Repetida = GL_AR_BD_DatosSalida(0, 3)
    GL_Num_Orden = GL_AR_BD_DatosSalida(0, 4)
    GL_Tip = GL_AR_BD_DatosSalida(0, 5)
    GL_Cod_Acc_Padre = GL_AR_BD_DatosSalida(0, 6)
    GL_Acc_simple = GL_AR_BD_DatosSalida(0, 7)
    GL_Num_Param = GL_AR_BD_DatosSalida(0, 8)

End Sub

