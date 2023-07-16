Attribute VB_Name = "bas_a6_universo"
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


Function fl_Escribir_Num_Ent(p_Cod_Uni As Long) As Long

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Modificacion1
    'SQL
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_UNIVERSO & " WHERE " & CTE_UNIVERSO_Cod_Uni & " = " & p_Cod_Uni
    'Array de datos de entrada a la BD (un registro)
     ReDim GL_AR_BD_DatosEntrada(CTE_N_UNIVERSO - 1) As Long
     GL_AR_BD_DatosEntrada(4) = GL_Num_Ent
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

Function fl_Leer_Num_Ent(p_Cod_Uni As Long) As Long

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'SQL
     GS_BD_SQL = "SELECT " & CTE_UNIVERSO_Num_Ent & " FROM " & CTE_TABLA_UNIVERSO & " WHERE " & CTE_UNIVERSO_Cod_Uni & " = " & p_Cod_Uni
    'Número de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = 1
    'Array de datos de entrada a la BD (un registro)
     ReDim GL_AR_BD_DatosEntrada(CTE_N_UNIVERSO - 1) As Long
     GL_AR_BD_DatosEntrada(0) = GL_Cod_Uni
     GL_AR_BD_DatosEntrada(1) = GL_Cod_Ent
     GL_AR_BD_DatosEntrada(2) = GL_Uni_Viv
     GL_AR_BD_DatosEntrada(3) = GL_Uni_Pri
     GL_AR_BD_DatosEntrada(4) = GL_Num_Ent
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
    
    'Tratamiento acceso BD correcto
     fl_Leer_Num_Ent = CLng(GL_AR_BD_DatosSalida(0, 0))


End Function


Sub S_EscribirUNIVERSO()


'Ejemplo de modificación de un registro en una tabla

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Modificacion1
    'Esqueleto de SELECT elegido
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_UNIVERSO & " WHERE " & CTE_UNIVERSO_Cod_Uni & " = " & GL_Cod_Uni
    'Array de datos de entrada a la BD (un registro)
     ReDim GL_AR_BD_DatosEntrada(CTE_N_UNIVERSO - 1) As Long
     GL_AR_BD_DatosEntrada(0) = GL_Cod_Uni
     GL_AR_BD_DatosEntrada(1) = GL_Cod_Ent
     GL_AR_BD_DatosEntrada(2) = GL_Uni_Viv
     GL_AR_BD_DatosEntrada(3) = GL_Uni_Pri
     GL_AR_BD_DatosEntrada(4) = GL_Num_Ent
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

End Sub

Sub S_LeerUNIVERSO()

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'SQL
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_UNIVERSO & " WHERE " & CTE_UNIVERSO_Cod_Uni & " = " & GL_Cod_Uni
    'Número de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = CTE_N_UNIVERSO 'todos, incluido des
    'Array donde se recibe el resultado de la BD (un registro) excepto des, que va en GS_BD_DesSalida
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
    
    'Tratamiento acceso BD correcto
     GS_Des_Uni = GS_BD_DesSalida
     GL_Cod_Uni = CLng(GL_AR_BD_DatosSalida(0, 0))
     GL_Cod_Ent = CLng(GL_AR_BD_DatosSalida(0, 1))
     GL_Uni_Viv = CLng(GL_AR_BD_DatosSalida(0, 2))
     GL_Uni_Pri = CLng(GL_AR_BD_DatosSalida(0, 3))
     GL_Num_Ent = CLng(GL_AR_BD_DatosSalida(0, 4))
    'Despues del tratamiento del acceso, liberamos la memoria ocupada
    'por el array que contiene los datos de salida de la consulta.
     ReDim GL_AR_BD_DatosSalida(0, 0) As Long
  
     Exit Sub



End Sub

