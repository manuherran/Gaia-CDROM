Attribute VB_Name = "bas_a6_variable"
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

Function fv_Leer_Vble(p_Cod_Uni As Long, p_Cod_Ent As Long, p_Cod_Vble As Long) As Variant

'Lee una variable de la tabla VBLE
'definida por Cod_Uni + Cod_Ent + Cod_Vble
'y devuelve el valor leido
'que puede ser entero 1, string 2

'Control del error producido por Visual Basic
 On Error GoTo Error_fv_Leer_Vble

'Declaración de variables
 Dim I_Error              As Integer
 Dim I_n                  As Integer
 
 Dim valor As Variant

'SQL
 GS_BD_SQL = "SELECT " & CTE_VBLE_Valor & " FROM " & CTE_TABLA_VBLE & " WHERE " & CTE_VBLE_Cod_Uni & " = " & p_Cod_Uni & " AND " & CTE_VBLE_Cod_Ent & " = " & p_Cod_Ent & " AND " & CTE_VBLE_Cod_Vble & " = " & p_Cod_Vble
 
'Número de campos de la tabla que se desean consultar
 GI_BD_NCamposConsulta = 1

'Acceso a la base de datos
'Realizar consulta, es decir, abrir dynaset
 Set G_RS_ds1 = G_DB_db1.OpenRecordset(GS_BD_SQL, dbOpenDynaset)

'Miramos si hay registro
 If G_RS_ds1.RecordCount = 0 Then
       'Informamos al programador del error producido
        GS_BD_Error = CTE_ErrorCNE
       'Cerramos el dynaset
        G_RS_ds1.Close
        GoTo Error_fv_Leer_Vble
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
    GoTo Error_fv_Leer_Vble
 End If

'4.- Controlamos el número de registros devuelto debe ser solo 1
If G_RS_ds1.RecordCount > 1 Then
  'Informamos al programador del error producido
   GS_BD_Error = CTE_ErrorSMR
  'Cerramos el dynaset
   G_RS_ds1.Close
   GoTo Error_fv_Leer_Vble
End If

'5.- Devolvemos los datos pedidos
valor = G_RS_ds1.Fields(0).Value

'7.- Finalización del procedimiento
'Cerramos el dynaset
 G_RS_ds1.Close
 
'Devolvemos el valor
 fv_Leer_Vble = valor
 
 Exit Function
    
        
'==================================================================================
'Finalización errónea
Error_fv_Leer_Vble:
  
    'Tratamiento error acceso BD
     Beep
    'Visualizamos el error producido por el desarrollo de la funcion
     MsgBox ("Num Error: " & Err & ". Texto: " & error & ". Gaia: " & GS_BD_Error & ".")
    'Finalizamos la aplicación
     End

End Function

