Attribute VB_Name = "bas_z0_bd"
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


'----------BASES DE DATOS---------------
'Base de datos a la que se accede
Global GI_BD_NumeroDeBD As Integer
'Operación a realizar: A,B,M,C
Global GS_BD_Operacion  As String
'Número de campos de la tabla que se desean consultar
Global GI_BD_NCamposConsulta As Integer
'Esqueleto de SELECT elegido
Global GI_BD_Esqueleto As Integer
'Array de parámetros de SELECT
Global GS_AR_BD_Parametro() As String
'Array de datos de entrada a la BD
Global GL_AR_BD_DatosEntrada() As Long
Global GS_BD_DesEntrada As String
'Array donde se recibe el resultado de la BD y la des
Global GL_AR_BD_DatosSalida() As Long
Global GS_BD_DesSalida As String
'Númerode registros afectados
Global GL_BD_NRegistros As Long
'Error producido
Global GS_BD_Error As String


Function Fi_Cerrar_Base_Datos()

'Detección de error por VB
On Error GoTo trat_error


    'Cerrar Base de Datos
    G_DB_db1.Close

    Exit Function

'Tratamiento de Error
trat_error:

    MsgBox Err & " " & error, vbCritical

End Function


Function Fi_Abrir_Base_Datos() As Integer

'Detección de error por VB
On Error GoTo trat_error
    
    'Abrir de Base de Datos
    
    'Identificamos la base de datos a la que accedemos
    'normalmente esto va en un .INI
    'GS_path_bd = "x.mdb"
    'GS_dbms = Fs_Lee_Fichero_Ini(CTE_Path_ini, "BD", "dbms")
    'GS_path_bd = Fs_Lee_Fichero_Ini(CTE_Path_ini, "BD", "path")
     
    GS_dbms = "Access"
    Select Case GS_dbms
        Case "Access"
            Set G_WS_ws1 = DBEngine.Workspaces(0)
            Set G_DB_db1 = G_WS_ws1.OpenDatabase(GS_path_bd)
        Case "FoxPro"
        
        Case "ODBC"
            Set G_WS_ws1 = DBEngine.Workspaces(0)
            Set G_DB_db1 = G_WS_ws1.OpenDatabase(GS_path_bd, False, False, "ODBC;")
        Case Else
    
    End Select
   
   
    Exit Function
trat_error:
    MsgBox Err & " " & error & " .No se encuentra la base de datos"


End Function

Sub S_AccesoBD()

'Inicialización de variables
'Suponemos que no va a haber error
 GS_BD_Error = CTE_ErrorNinguno


 Select Case GI_BD_NumeroDeBD
     Case 1
      
          Select Case GS_BD_Operacion
              
              Case CTE_BD_Alta1
               'Alta de un registro
                S_AltaDe1Registro

              Case CTE_BD_Baja1
               'Baja de un registro
                S_BajaDe1Registro
              
              Case CTE_BD_Modificacion1
               'Modificacion de un registro
                S_ModificacionDe1Registro
              
              Case CTE_BD_Consulta1
               'Consulta de un registro
                S_ConsultaDe1Registro

              Case CTE_BD_AltaN
               'Alta de N registros
                S_AltaDeNRegistros

              Case CTE_BD_BajaN
               'Baja de N registros
                S_BajaDeNRegistros
              
              Case CTE_BD_ModificacionN
               'Modificacion de N registros
                S_ModificacionDeNRegistros
              
              Case CTE_BD_ConsultaN
               'Consulta de N registros
                S_ConsultaDeNRegistros
          
          End Select
 
 End Select

End Sub

Sub S_AltaDe1Registro()

'Control del error producido por Visual Basic
 On Error GoTo Error_AltaDe1Registro

'Declaración de variables
 Dim S_Sentencia        As String
 Dim I_Error              As Integer
 Dim i              As Integer
 
'Realizar consulta, es decir, abrir dynaset
 Set G_RS_ds1 = G_DB_db1.OpenRecordset(GS_BD_SQL, dbOpenDynaset)

'Control del error en RealizarConsulta
If I_Error <> CTE_NO_HAY_ERROR Then
    'Informamos al programador del error producido
    GS_BD_Error = CTE_ErrorRC
    'Cerramos el dynaset
    G_RS_ds1.Close
Exit Sub
End If

'El dyna ha de estar vacío, si no es alta existente (error)
If G_RS_ds1.RecordCount > 0 Then
    'Informamos al programador del error producido
    GS_BD_Error = CTE_ErrorAE
    'Cerramos el dynaset
    G_RS_ds1.Close
    Exit Sub
End If

'Comienza la transacción
BeginTrans

'Alta
G_RS_ds1.AddNew
'G_RS_ds1.Fields(0) = GS_BD_DesEntrada
For i = CTE_num_campos_des To UBound(GL_AR_BD_DatosEntrada)
    G_RS_ds1.Fields(i) = GL_AR_BD_DatosEntrada(i - 1)
Next i
G_RS_ds1.Update

'Control del error en EjecutarSentencia
If I_Error <> CTE_NO_HAY_ERROR Then
    'Informamos al programador del error producido
    GS_BD_Error = CTE_ErrorES
    'Anulamos la ejecucion realizada, por detectar un error
    Rollback
    'Cerramos el dynaset
    G_RS_ds1.Close
    Exit Sub
End If

'Finalización correcta
CommitTrans
'Cerramos el dynaset
G_RS_ds1.Close
Exit Sub


'==================================================================================
'Finalización errónea
Error_AltaDe1Registro:
'Anulamos la ejecucion realizada, por detectar un error
Rollback
'Informamos al programador del error producido
GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
G_RS_ds1.Close
End Sub

Sub S_AltaDeNRegistros()

'Pasos:

'Primera línea de código
'Control del error producido por Visual Basic
     On Error GoTo Error_AltaDeNRegistros
'Declaración de variables
 Dim S_Sentencia        As String
 Dim i_err              As Integer
 Dim I_Error              As Integer
 Dim I_n As Integer


'0.- Bucle de carga de los N registros
'For I_n = 0 To N - 1

'1.- Cargar los datos a dar de alta
    ReDim AR_PARAMETROS(11) As String
    AR_PARAMETROS(0) = "ddd"
          
'2.- Preparamos la sentencia de SQL(montamos la SELECT)
'Hacemos SELECT de lo que vamos a dar de alta
    ReDim ARS_MiArray(3) As String
    ARS_MiArray(0) = "ww"
    ARS_MiArray(1) = "Coh"
    ARS_MiArray(2) = "d"
'    S_Sentencia = Fs_PrepararSentencia(ARS_MiArray(), 0)

'3.- Realizamos la consulta: se abre el dyna sólo la primera vez
    If I_n = 0 Then
'        I_Error = Fi_RealizarConsulta(S_Sentencia, CTE_DY_BATCH)
    End If

'4.- El dyna ha de estar vacío, si no es alta existente (error)
'problema: sólo se valida la primera vez.

'5.- Ejecutamos el alta del registro
    'I_Error = Fi_EjecutarSentencia(CTE_TABLA_BATCH, CTE_DY_BATCH, AR_PARAMETROS(), CTE_ANADIR)

'6.- Fin bucle
'Next I_n





'X.- Inicializamos el valor devuelto por la funcion
     Exit Sub


'==================================================================================
'Finalización errónea
Error_AltaDeNRegistros:
'Anulamos la ejecucion realizada, por detectar un error
 Rollback
'Informamos al programador del error producido
 GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
 G_RS_ds1.Close
End Sub

Sub S_BajaDe1Registro()
    
'Control del error producido por Visual Basic
 On Error GoTo Error_BajaDe1Registro

'Declaración de variables
 Dim I_Error        As Integer

'Realizar consulta, es decir, abrir dynaset
 Set G_RS_ds1 = G_DB_db1.OpenRecordset(GS_BD_SQL, dbOpenDynaset)

'Miramos si hay registro
 If G_RS_ds1.RecordCount = 0 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorBNE
       'Cerramos el dynaset
        G_RS_ds1.Close
        Exit Sub
 Else
        If G_RS_ds1.RecordCount > 1 Then
            'Informamos al programador del error producido
             GS_BD_Error = CTE_ErrorSMR
            'Cerramos el dynaset
             G_RS_ds1.Close
             Exit Sub
        End If
 End If


'Comienza la transacción
BeginTrans
'Borramos el registro
G_RS_ds1.Delete
'Finalización correcta
CommitTrans
'Cerramos el dynaset
G_RS_ds1.Close
Exit Sub
'==================================================================================
'Finalización errónea
Error_BajaDe1Registro:
'Anulamos la ejecucion realizada, por detectar un error
Rollback
'Informamos al programador del error producido
GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
G_RS_ds1.Close
End Sub

Sub S_BajaDeNRegistros()

'Control del error producido por Visual Basic
 On Error GoTo Error_BajaDeNRegistros

'Declaración de variables
 Dim S_Sentencia        As String
 Dim I_Error            As Integer

'Realizar consulta, es decir, abrir dynaset
 Set G_RS_ds1 = G_DB_db1.OpenRecordset(GS_BD_SQL, dbOpenDynaset)

'Control del error en RealizarConsulta
If I_Error <> 0 Then
    'Informamos al programador del error producido
     GS_BD_Error = CTE_ErrorRC
    'Cerramos el dynaset
     G_RS_ds1.Close
     Exit Sub
End If

'El dyna ha de tener N registros, si no, es baja no existente (error)
If G_RS_ds1.RecordCount = 0 Then
    'Informamos al programador del error producido
    GS_BD_Error = CTE_ErrorBNE
    'Cerramos el dynaset
    G_RS_ds1.Close
    Exit Sub
End If
        
'Comienza la transacción
 BeginTrans
'Borramos los N registros
While Not G_RS_ds1.EOF
    'Borramos el registro seleccionado
    G_RS_ds1.Delete
    G_RS_ds1.MoveNext
Wend
'Finalización correcta
 CommitTrans
'Cerramos el dynaset
G_RS_ds1.Close
 Exit Sub
    
'==================================================================================
'Finalización errónea
Error_BajaDeNRegistros:
'Anulamos la ejecucion realizada, por detectar un error
Rollback
'Informamos al programador del error producido
GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
G_RS_ds1.Close
End Sub

Sub S_ConsultaDe1Registro()

'Control del error producido por Visual Basic
 On Error GoTo Error_ConsultaDe1Registro
'Declaración de variables
 Dim I_Error              As Integer
 Dim I_n              As Integer
'Realizar consulta, es decir, abrir dynaset
 Set G_RS_ds1 = G_DB_db1.OpenRecordset(GS_BD_SQL, dbOpenSnapshot)
'Miramos si hay registro
 If G_RS_ds1.RecordCount = 0 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorCNE
       'Cerramos el dynaset
        G_RS_ds1.Close
        Exit Sub
 End If
'Situar puntero en la primera; por precaución
'se hace antes al último. si la tabla está vacía,
'se produce el error 3021
 G_RS_ds1.MoveLast
 G_RS_ds1.MoveFirst
'Control del error en RealizarConsulta
     If I_Error <> 0 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorRC
       'Cerramos el dynaset
        G_RS_ds1.Close
        Exit Sub
     End If
'4.- Controlamos el número de registros devuelto debe ser solo 1
     If G_RS_ds1.RecordCount > 1 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorSMR
       'Cerramos el dynaset
        G_RS_ds1.Close
        Exit Sub
     End If
'5.- Devolvemos los datos pedidos
    'Evitamos los primeros registro que son des siempre y string
     For I_n = 1 To CTE_num_campos_des
         GS_BD_DesSalida = "" & G_RS_ds1.Fields(I_n - 1).Value
     Next I_n
     'el resto son todos numeros enteros
     For I_n = CTE_num_campos_des + 1 To GI_BD_NCamposConsulta
        GL_AR_BD_DatosSalida(0, I_n - 1 - CTE_num_campos_des) = G_RS_ds1.Fields(I_n - 1).Value
     Next I_n
'6.- Devolvemos el número de registros afectados
     GL_BD_NRegistros = G_RS_ds1.RecordCount
'7.- Finalización del procedimiento
       'Cerramos el dynaset
        G_RS_ds1.Close
 Exit Sub
    
'==================================================================================
'Finalización errónea
Error_ConsultaDe1Registro:
'Informamos al programador del error producido
GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
G_RS_ds1.Close
End Sub

Sub S_ConsultaDeNRegistros()

'Control del error producido por Visual Basic
 On Error GoTo Error_ConsultaDeNRegistros

'Declaración de variables
 Dim I_Error              As Integer
 Dim I_n              As Integer
 Dim I_m             As Integer
'Realizar consulta, es decir, abrir dynaset
 Set G_RS_ds1 = G_DB_db1.OpenRecordset(GS_BD_SQL, dbOpenSnapshot)
'Control del error en RealizarConsulta
If I_Error <> 0 Then
    'Informamos al programador del error producido
    GS_BD_Error = CTE_ErrorRC
    'Cerramos el dynaset
     G_RS_ds1.Close
    Exit Sub
End If

'Devolvemos los datos pedidos
     ReDim GL_AR_BD_DatosSalida(G_RS_ds1.RecordCount - 1, GI_BD_NCamposConsulta - 1)
     For I_m = 1 To G_RS_ds1.RecordCount
        For I_n = 1 To GI_BD_NCamposConsulta
            GL_AR_BD_DatosSalida(I_m - 1, I_n - 1) = G_RS_ds1.Fields(I_n - 1).Value
        Next I_n
        G_RS_ds1.MoveNext
     Next I_m

'Devolvemos el número de registros afectados
GL_BD_NRegistros = G_RS_ds1.RecordCount

'Cerramos el dynaset
G_RS_ds1.Close
Exit Sub
    
'==================================================================================
'Finalización errónea
Error_ConsultaDeNRegistros:
'Informamos al programador del error producido
 GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
G_RS_ds1.Close
End Sub

Sub S_ModificacionDe1Registro()

'Para modificar un único elemento, se modifica
'tal cual, con un update. En el caso de ser N no siempre...

'Control del error producido por Visual Basic
 On Error GoTo Error_ModificacionDe1Registro

'Declaración de variables
 Dim I_Error As Integer
 Dim i As Integer
 
'Realizar consulta, es decir, abrir dynaset
 Set G_RS_ds1 = G_DB_db1.OpenRecordset(GS_BD_SQL, dbOpenDynaset)

'Miramos si hay registro
 If G_RS_ds1.RecordCount = 0 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorMNE
       'Cerramos el dynaset
        G_RS_ds1.Close
        Exit Sub
 End If


'Control del error en RealizarConsulta
     If I_Error <> 0 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorRC
       'Cerramos el dynaset
        G_RS_ds1.Close
        Exit Sub
     End If

'4.- El dyna no ha de estar vacío, si no es modificación no existente (error)
     If G_RS_ds1.RecordCount = 0 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorMNE
       'Cerramos el dynaset
        G_RS_ds1.Close
        Exit Sub
     End If

'5.- Comienza la transacción
     BeginTrans

    G_RS_ds1.Edit
    'G_RS_ds1.Fields(0) = GS_BD_DesEntrada
    For i = CTE_num_campos_des To UBound(GL_AR_BD_DatosEntrada)
        G_RS_ds1.Fields(i) = GL_AR_BD_DatosEntrada(i - 1)
    Next i
    G_RS_ds1.Update


'6.- Ejecutamos el alta del registro y cerramos el dyna
 '    I_Error = Fi_EjecutarSentenciaGrabar(GS_BD_Tabla, CTE_DY_CERO, GL_AR_BD_DatosEntrada(), CTE_MODIFICAR)

'7.- Control del error en EjecutarSentencia
     If I_Error <> 0 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorES
       'Finalización correcta
        CommitTrans
       'Cerramos el dynaset
        G_RS_ds1.Close
        Exit Sub
     End If
    
'8.- Finalización del procedimiento
    
'Finalización correcta
 CommitTrans
'Cerramos el dynaset
 G_RS_ds1.Close
 Exit Sub
    
'==================================================================================
'Finalización errónea
Error_ModificacionDe1Registro:
'Anulamos la ejecucion realizada, por detectar un error
 Rollback
'Informamos al programador del error producido
 GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
 G_RS_ds1.Close
End Sub

Sub S_ModificacionDeNMRegistros()

'Problema: normalmente ocurrirá que la modificación
'de N registros supone en realidad borrar N registros y dar
'de alta M, siendo probablemente N<>M
'Hay dos fases:
    'A: borrar los viejos
    'B: dar de alta los nuevos

'Pasos:
'Declaración de variables
 Dim S_Sentencia        As String
 Dim i_err              As Integer
 Dim I_Error              As Integer
 Dim I_n As Integer

'Primera línea de código
'Control del error producido por Visual Basic
     On Error GoTo Error_ModificacionDeNMRegistros


'A.- Borramos los viejos. Es una baja de N registros
'1.- Preparamos la sentencia de SQL de los elementos a dar de baja
     ReDim ARS_MiArray(3) As String
   '  ARS_MiArray(0) = S_TablaABorrar
     ARS_MiArray(1) = "Cod_Batch"
'     ARS_MiArray(2) = GS_ClaveReg
   '  S_Sentencia = Fs_PrepararSentencia(ARS_MiArray(), 0)

'2.- Realizamos la consulta: se abre el dyna
'     i_err = Fi_RealizarConsulta(S_Sentencia, CTE_DY_BATCH)

'3.- Si el dyna esta vacío es un error

        'Control de la operacion realizada
         If i_err <> CTE_NO_HAY_ERROR Then
           'Inicializamos el valor devuelto por la funcion
         '   Fi_S_ModificacionDeNRegistros = CTE_HAY_ERROR
            Exit Sub
         End If
    'Control de registros seleccionados por la SELECT, si no encontramos
    'ningun registro, no hay nada que borrar
'
        
'5.- Borramos los N registros
'         While Not 'GDY_AR_Arraydynas(CTE_DY_BATCH).EOF
            'Borramos el registro seleccionado
             'GDY_AR_Arraydynas(CTE_DY_BATCH).Delete
             'GDY_AR_Arraydynas(CTE_DY_BATCH).MoveNext
'         Wend
'     End If

'B.- Damos de alta los nuevos: ¡El dyna ya esta abierto!

'Controlar que los registros no existan ya

'Bucle de carga de los N registros en BATCHFICHERO
'For I_n = 0 To Gi_temp_FICH_N - 1

    ReDim AR_PARAMETROS_BATCHFICHERO(2) As String
'    AR_PARAMETROS_BATCHFICHERO(0) = GS_ClaveReg
'Cod_Fic
    'AR_PARAMETROS_BATCHFICHERO(1) = GS_AR_Temp_FICH_Cod_Fic(I_n)
'Ejecutamos el alta del registro
    'I_Error = Fi_EjecutarSentenciaGrabar(CTE_TABLA_BATCHFICHERO, CTE_DY_BATCHFICHERO, AR_PARAMETROS_BATCHFICHERO(), CTE_ANADIR)
'Next I_n

'Cerramos el dynaset
 G_RS_ds1.Close






'X.- Inicializamos el valor devuelto por la funcion
  '   Fl_Eliminar = CTE_NO_HAY_ERROR
     Exit Sub


'==================================================================================
'Finalización errónea
Error_ModificacionDeNMRegistros:
'Anulamos la ejecucion realizada, por detectar un error
 Rollback
'Informamos al programador del error producido
 GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
 G_RS_ds1.Close
End Sub

Sub S_ModificacionDeNRegistros()

'Pasos:

'Primera línea de código
'Control del error producido por Visual Basic
     On Error GoTo Error_ModificacionDeNRegistros





'X.- Inicializamos el valor devuelto por la funcion
    ' Fl_Eliminar = CTE_NO_HAY_ERROR
     Exit Sub



'==================================================================================
'Finalización errónea
Error_ModificacionDeNRegistros:
'Anulamos la ejecucion realizada, por detectar un error
 Rollback
'Informamos al programador del error producido
 GS_BD_Error = CTE_ErrorVB
'Cerramos el dynaset
 G_RS_ds1.Close
End Sub



