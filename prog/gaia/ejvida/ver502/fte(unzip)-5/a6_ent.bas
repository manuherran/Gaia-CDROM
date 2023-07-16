Attribute VB_Name = "bas_a6_entidad"
Option Explicit
'------------------------------------------------------------------------
' Shareware desarrollado por Manuel de la Herr�n Gasc�n
' mherran@usa.net (Junio 1997 - Diciembre 1998) Madrid (Spain).
' http://www.geocities.com/SiliconValley/Vista/7491/
' -----------------------------------------------------------------------
' Este programa y sus ficheros fuente son gr�tis y de libre distribuci�n.
' El c�digo fuente est� disponible y puede ser modificado, distribuido,
' o utilizado en otros programas con entera libertad.
' -----------------------------------------------------------------------
' Para mantenerse informado de las sucesivas versiones del programa
' y d�nde conseguirlas, escriba un mail a mherran@usa.net
' Para sugerir posibles ampliaciones, enviar comentarios de cualquier tipo
' si se detectara alg�n error en la programaci�n o en la instalaci�n,
' o si se va a ampliar o utilizar una parte o todo este
' programa, no dude en ponerse en contacto con el autor.
'-----------------------------------------------------------------------
Sub S_EjecutarPriEntidades()

'1.- Inicializaci�n de variables
     GL_N_Entidades_Ejecutadas = 0
    
'2.- Mientras no se hayan ejecutado todas...
     While GL_N_Entidades_Ejecutadas < GL_Uni_Pri

'4.- La leemos, ejecutamos y tb escribimos la actual y calculamos la siguiente
     S_EjecutarEntidad

'6.- Contador
     GL_N_Entidades_Ejecutadas = GL_N_Entidades_Ejecutadas + 1
     
'8.- Vemos si hay que finalizar
     Wend



End Sub

Sub S_EjecutarEntidad()

' Subprocedimiento S_EjecutarEntidad
'
' OBJETIVO: La entidad GL_Cod_Ent tiene un �rbol de acciones
' que se han de ejecutar constantemente. Esta funci�n ejecuta
' las acciones que corresponden a un turno de ejecuci�n de
' una entidad. Se trata de:
'   1.- buscar el lugar donde se detuvo la ejecuci�n la �ltima vez,
'       en esa entidad
'   2.- ejecutar un n�mero de acciones en funci�n de la prioridad
'       correspondiente a ella
'   3.- marcar el lugar donde debe comenzar la pr�xima ejecuci�n
'       de esa entidad
'
' IN:   GL_Cod_Ent: C�digo de la entidad a ejecutar
'
'
' OUT:  No devuelve nada.
'


'suponemos que solo hay un universo, es decir gaia solo se usa
'para un unico proyecto cada vez, como puede ser la bolsa,
'y no se mezclan proyectos

'en global tenemos el c�digo de la entidad donde se detuvo
'la ejecuci�n de ese universo la ultima vez, por tanto tenemos
'un 1 en la primera ejecuci�n

'En la primera ejecuci�n de un proyecto, se ejecutan todas las entidades
'por orden comenzando por la primera. Pero si el proyecto se detiene, se
'almacena en global el c�digo de la entidad en curso para continuar la ejecuci�n desde
'esa y no desde el principio.

'Como cada entidad solo ejecuta N acciones cada vez, y luego se pasa
'a la ejecuci�n de otra entidad, cada entidad debe tener almacenado
'cual es el la instrucci�n que le tocar� ejecutar la siguiente vez
'que es cod ent + cod acc + num serie en la tabla entidad

'En todo momento, la lista de entidades es correlativa, es decir, existen
'entidades desde la uno hasta la N, todas, de una en una. No existe la
'entidad 3 y luego la 5. Paar ello, si se borra una entidad, se coge la
'ultima y se cambia su indice a la de la borrada, y a sus acciones
'variables, etc tambi�n... lo mejor ser� no borrar nunca, y simplemente
'darlas por muertas y incluir un proceso en el que se borren fisicamente
'las entidades muertas creadas a partir de cierta fecha...y cosas asi


'Despu�s de ejecutar una entidad, se incrementa el c�digo de la actual, y si
'esa no existe, es que ya no existen m�s, y se vuelve a comenzar desde
'la uno.

'No hay reutilizaci�n de acciones complejas. Cada entidad tiene su propia
'y completa descomposici�n de acciones en subacciones, aunque coincidan
'iguales en varias entidades.

'S� hay reutilizaci�n de acciones simples o b�sicas (directamente ejecutables)

'Para identificar a una acci�n simple o compleja
'se usa conjuntamente el c�digo de
'entidad + el c�digo de acci�n, pero pudiera ser que una entidad
'tuviera m�s de una vez en su arbol la misma acci�n, por lo
'que la clave de ACCION es Cod_Acci�n, Cod_Entidad y N�m_Serie
'siendo este �ltimo el n�mero que identifica a esa acci�n cuando
'hay varias, comenzando desde uno.

'Los tipos de acciones son:
'   0: compleja.
'   1: simple (directamente ejecutable)
'
'   el resto de acciones son de control de flujo:
'   2: IF condici�n THEN ejecutar el hermano n�mero n


     
    'Accedemos a la acci�n a ejecutar, definida por
    'GL_Cod_Ent + GL_Cod_Acc + GL_Num_Repetida
    'y tomamos sus datos sobre las variables globales
     S_Leer_ENTIDAD
     
    'Si la entidad est� viva, hay que ejecutarla
     If GL_Ent_Viv = 1 Then
        'Ejecutamos GL_Pri acciones en total (simples + complejas)
         S_EjecutarPriAcciones
         If GI_modo_de_ejecucion >= 2 Then
            S_MostrarDatosUniverso
            S_MostrarDatosEntidad
         End If
     Else
        If GI_modo_de_ejecucion >= 2 Then
            S_BorrarInformacionAccion
            S_MostrarDatosUniverso
            S_MostrarDatosEntidad
        End If
        If GI_modo_de_ejecucion >= 3 Then
           'Informamos al usuario de la entidad que se acaba de ejecutar
            If MsgBox("Entidad " & GL_Cod_Ent & " Ejecutada. �Mantener el modo de ejecuci�n?", 1) = 2 Then GI_modo_de_ejecucion = GI_modo_de_ejecucion - 1
        End If
     End If
     
    'Escribimos la entidad, porque al menos la acci�n a ejecutar
    'habr� cambiado
     S_Escribir_ENTIDAD
     
     
    'Pasamos a la siguiente entidad
    'Calculamos la siguiente entidad a ejecutar
    'Si no existe, la funci�n de leer entidad ya se encarga de
    'leer la siguiente
     GL_Cod_Ent = GL_Cod_Ent + 1
     If GL_Cod_Ent > GL_Num_Ent Then
         'Pasamos a la primera entidad
         GL_Cod_Ent = 1
     End If

        

End Sub

Sub S_Leer_ENTIDAD()


'1.- Accedemos a los datos de la entidad a ejecutar
    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operaci�n a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'Esqueleto de SELECT elegido
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_ENTIDAD & " WHERE " & CTE_ENTIDAD_Cod_Uni & " = " & GL_Cod_Uni & " AND " & CTE_ENTIDAD_Cod_Ent & " = " & Str$(GL_Cod_Ent)
    'N�mero de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = CTE_N_ENTIDAD 'todos
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
        'Finalizamos la aplicaci�n
         End
     Else
        'Tratamiento acceso BD correcto
         GS_Des_Ent = GS_BD_DesSalida
         GL_Cod_Uni = CLng(GL_AR_BD_DatosSalida(0, 0))
         GL_Cod_Ent = CLng(GL_AR_BD_DatosSalida(0, 1))
         GL_Ent_Viv = CLng(GL_AR_BD_DatosSalida(0, 2))
         GL_Ent_Pri = CLng(GL_AR_BD_DatosSalida(0, 3))
         GL_Cod_Obj = CLng(GL_AR_BD_DatosSalida(0, 4))
         GL_Cod_Acc = CLng(GL_AR_BD_DatosSalida(0, 5))
         GL_Num_Repetida = CLng(GL_AR_BD_DatosSalida(0, 6))
        'Despues del tratamiento del acceso, liberamos la memoria ocupada
        'por el array que contiene los datos de salida de la consulta.
         ReDim GL_AR_BD_DatosSalida(0, 0) As Long
     End If
     


End Sub
Sub S_Escribir_ENTIDAD()


    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operaci�n a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Modificacion1
    'Esqueleto de SELECT elegido
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_ENTIDAD & " WHERE " & CTE_ENTIDAD_Cod_Uni & " = " & GL_Cod_Uni & " AND " & CTE_ENTIDAD_Cod_Ent & " = " & Str$(GL_Cod_Ent)
    'Array de datos de entrada a la BD (un registro)
     'GS_BD_DesEntrada se mantiene su valor
     ReDim GL_AR_BD_DatosEntrada(0 To 6) As Long
     GL_AR_BD_DatosEntrada(0) = GL_Cod_Uni
     GL_AR_BD_DatosEntrada(1) = GL_Cod_Ent
     GL_AR_BD_DatosEntrada(2) = GL_Ent_Viv
     GL_AR_BD_DatosEntrada(3) = GL_Ent_Pri
     GL_AR_BD_DatosEntrada(4) = GL_Cod_Obj
     GL_AR_BD_DatosEntrada(5) = GL_Cod_Acc
     GL_AR_BD_DatosEntrada(6) = GL_Num_Repetida
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
        'Finalizamos la aplicaci�n
         End
     End If


End Sub
