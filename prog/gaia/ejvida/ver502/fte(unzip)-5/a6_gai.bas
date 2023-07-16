Attribute VB_Name = "bas_a6_gai"
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

'Declaraciones de acceso a base de datos
Global G_WS_ws1 As Workspace        'workspace
Global G_DB_db1 As Database         'database
Global G_RS_ds1 As Dynaset          'recordset de tipo dynaset o Snapshot segun sea ABM o C
Global GS_BD_SQL As String          'sentencia sql

Global GS_dbms As String            'admin de bd, (Access...)
Global GS_path_bd As String         'nombre fichero bd

'Modo de Ejecución
Global GI_modo_de_ejecucion As Integer

Global GI_Finalizar As Integer



'Número total de entidades ejecutadas hasta el momento
'por el universo actual
 Global GL_N_Entidades_Ejecutadas As Long

'Número total de acciones ejecutadas hasta el momento
'por la entidad actual
 Global GL_N_Acciones_Ejecutadas As Long

'Número de universos
 Global GL_Num_Uni As Long

'Universo actual
 Global GS_Des_Uni As String
 Global GL_Cod_Uni As Long
 Global GL_Uni_Viv As Long
 Global GL_Uni_Pri As Long
 Global GL_Num_Ent As Long 'Número de entidades del universo actual

'Entidad actual
 Global GS_Des_Ent As String
 Global GL_Cod_Ent As Long
     
'Datos de la entidad actual
 Global GL_Ent_Viv  As Long
 Global GL_Ent_Pri As Long
 Global GL_Cod_Obj As Long

'Acción actual
 Global GS_Des_Acc As String
 Global GL_Cod_Acc  As Long

'Datos de la acción actual
 Global GL_Num_Repetida  As Long
 Global GL_Num_Orden  As Long
 Global GL_Tip  As Long
 Global GL_Cod_Acc_Padre  As Long
 Global GL_Acc_simple  As Long
 Global GL_Num_Param As Long
 Global GV_Param() As Variant



Sub s_inicializar_ejemplo_elegido_gai()

    Select Case num_ej_activo_ejv
        Case 1
            GS_path_bd = path_largo_ejv(CTE_C_PRG) & "\gaia.mdb"
        Case 2
            GS_path_bd = path_largo_ejv(CTE_C_PRG) & "\gaia.mdb"
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: no existe ese ejemplo"
    End Select
    frm_a6_ingaia.Caption = "Plataforma Gaia " & GS_path_bd
    'Abro la base de datos, que se cierra al cerrar la ventana
    Fi_Abrir_Base_Datos

End Sub

Sub s_comenzar_gai()

    'Ejecutamos el proyecto
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, True
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, True
    s_ejecutar_proyecto_gai

End Sub


Sub s_cargar_etiquetas_gai()
    
    'Cargamos los valores fijos del Universo
    frm_a6_ingaia.Le_Uni.Clear
    frm_a6_ingaia.Le_Uni.AddItem "Universo"
    frm_a6_ingaia.Le_Uni.AddItem "Descripción"
    frm_a6_ingaia.Le_Uni.AddItem "Entidad a ejecutar"
    frm_a6_ingaia.Le_Uni.AddItem "Vivo"
    frm_a6_ingaia.Le_Uni.AddItem "Prioridad"
     
    'Cargamos los valores fijos de Entidad
    frm_a6_ingaia.Le_Ent.Clear
    frm_a6_ingaia.Le_Ent.AddItem "Entidad"
    frm_a6_ingaia.Le_Ent.AddItem "Descripción"
    frm_a6_ingaia.Le_Ent.AddItem "Viva"
    frm_a6_ingaia.Le_Ent.AddItem "Prioridad"
    frm_a6_ingaia.Le_Ent.AddItem "Objetivo"
    frm_a6_ingaia.Le_Ent.AddItem "Acción"
     
    'Cargamos los valores fijos del grid Acción
    frm_a6_ingaia.Le_Acc.Clear
    frm_a6_ingaia.Le_Acc.AddItem "Acción"
    frm_a6_ingaia.Le_Acc.AddItem "Descripción"
    frm_a6_ingaia.Le_Acc.AddItem "Acción Simple"

End Sub

Sub s_borrar_informacion_entidad_gai()
    
    'Entidad
    'frm_a6_ingaia.Le_Ent.Clear
    frm_a6_ingaia.Li_Ent.Clear
    
    'Acción
    'frm_a6_ingaia.Le_Acc.Clear
    frm_a6_ingaia.Li_Acc.Clear
    
    'frm_a6_ingaia.Refresh

End Sub

Sub S_BorrarInformacionAccion()
        
    'Acción
    frm_a6_ingaia.Li_Acc.Clear


End Sub

Sub S_MostrarDatosAccion()
    
    'Acción
    frm_a6_ingaia.Li_Acc.Clear
    frm_a6_ingaia.Li_Acc.AddItem GL_Cod_Acc
    frm_a6_ingaia.Li_Acc.AddItem GS_Des_Acc
    frm_a6_ingaia.Li_Acc.AddItem GL_Acc_simple
     

End Sub


Sub S_MostrarDatosEntidad()

    
    Dim texto As String
    
    'Entidad
    frm_a6_ingaia.Li_Ent.Clear
    frm_a6_ingaia.Li_Ent.AddItem GL_Cod_Ent
    frm_a6_ingaia.Li_Ent.AddItem GS_Des_Ent
    If GL_Ent_Viv = 0 Then
       texto = "0 (muerta)"
    Else
       texto = "1 (viva)"
    End If
    frm_a6_ingaia.Li_Ent.AddItem texto
    frm_a6_ingaia.Li_Ent.AddItem GL_Ent_Pri
    frm_a6_ingaia.Li_Ent.AddItem GL_Cod_Obj
    frm_a6_ingaia.Li_Ent.AddItem GL_Cod_Acc


End Sub

Sub S_MostrarDatosUniverso()


     Dim texto As String
    
    'Universo
    frm_a6_ingaia.Li_Uni.Clear
    frm_a6_ingaia.Li_Uni.AddItem GL_Cod_Uni
    frm_a6_ingaia.Li_Uni.AddItem GS_Des_Uni
    frm_a6_ingaia.Li_Uni.AddItem GS_Des_Ent
    If GL_Uni_Viv = 0 Then
       texto = "0 (muerto)"
    Else
       texto = "1 (vivo)"
    End If
    frm_a6_ingaia.Li_Uni.AddItem texto
    frm_a6_ingaia.Li_Uni.AddItem GL_Uni_Pri


End Sub


Sub S_EjecutarUniverso()


    'Escribir el universo cada vez que se cambia de universo
    'es porque los datos del universo normalmente han cambiado
    'por ejemplo, la entidad actual

    
    'Leemos la tabla UNIVERSO de Cod_Uni para saber la entidad actual
    'la prioridad de ese universo y si esta vivo
     S_LeerUNIVERSO
     
    'Si el universo está vivo y tiene entidades, hay que ejecutarlo
     If GL_Uni_Viv = 1 And GL_Num_Ent > 0 Then
        'Ejecutamos GL_Pri acciones en total (simples + complejas)
         S_EjecutarPriEntidades
         If GI_modo_de_ejecucion >= 1 Then
             S_MostrarDatosUniverso
         End If
     Else
         If GI_modo_de_ejecucion >= 1 Then
            S_MostrarDatosUniverso
            s_borrar_informacion_entidad_gai
            If GI_modo_de_ejecucion >= 2 Then
              'Informamos al usuario del universo ejecutado
               s_estado_detenido_gai
               If MsgBox("Universo " & GL_Cod_Uni & " Ejecutado. ¿Detener completamente la ejecución?", 1) = 1 Then
                   GI_Finalizar = True
                   GI_modo_de_ejecucion = 1
               End If
            End If
          End If
     End If
    
'    Ya se han ejecutado todas las que se debía:
'    se graba la accion actual, que todavia no se ha ejecutado
'    en la tabla entidad para preparar las próximas
'    ejecuciones en el siguiente ciclo
     S_EscribirUNIVERSO
    
    'Pasamos al siguiente universo
    'Calculamos el siguiente universo a ejecutar
     GL_Cod_Uni = GL_Cod_Uni + 1
     If GL_Cod_Uni > GL_Num_Uni Then
         'Pasamos al primer universo
         GL_Cod_Uni = 1
     End If



End Sub

Sub s_estado_detenido_gai()

    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, True
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, False

End Sub

Sub s_ejecutar_proyecto_gai()

'Todas las tablas tienen un campo des como primer campo
'de descripcion de la entidad, donde va entre otras cosas la
'fecha de creación de la entidad para limiezas de bd

'y por eso cuando se lee con
'select * en realidad, aunque se lee este campo, no se
'carga en las variables porque no se usa

'Al comenzar una ejecución se leen los datos de GLOBAL (1 registro)
'y se guardan en variables globales.
'Al detener la ejecución, habrá que actualizar la tabla de BD
'con los valores de estas variables

    frm_a6_ingaia.Show

    'Leemos la tabla GLOBAL para saber el universo actual Cod_Uni
    S_LeerGLOBAL
    
    'Inicialización de variables
    GI_Finalizar = False
    
    'Bucle principal del proyecto
    While GI_Finalizar = False
      'Leemos, Ejecutamos y Escribimos los valores actuales globales
      'en la tabla UNIVERSO par el actual
       S_EjecutarUniverso   'Aqui se calcula tb el siguiente
      'Solo aqui se permite que el usuario pulse pausa
       DoEvents
    
    Wend
    
    'Se ha detenido el proyecto
    S_EscribirGLOBAL

    s_fin_bucle_general_ejv

End Sub

Sub S_EscribirGLOBAL()

'Ejemplo de modificación de un registro en una tabla

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Modificacion1
    'Esqueleto de SELECT elegido
     GS_BD_SQL = "SELECT * FROM " & CTE_TABLA_GLOBAL
    'Array de datos de entrada a la BD (un registro)
     ReDim GL_AR_BD_DatosEntrada(CTE_N_GLOBAL - 1) As Long
     GL_AR_BD_DatosEntrada(0) = GL_Cod_Uni
     GL_AR_BD_DatosEntrada(1) = GL_Num_Uni
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
    
    'Tratamiento acceso BD correcto
     Exit Sub


End Sub


Sub S_LeerGLOBAL()

    'Base de datos a la que se accede
     GI_BD_NumeroDeBD = 1
    'Operación a realizar: A,B,M,C
     GS_BD_Operacion = CTE_BD_Consulta1
    'SQL
     GS_BD_SQL = "SELECT * FROM GLOBAL"
    'Número de campos de la tabla que se desean consultar
     GI_BD_NCamposConsulta = CTE_N_GLOBAL 'todos, incluido des
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
    
    'Tratamiento acceso BD correcto: Calculamos el universo actual y el numero total
     GS_Des_Uni = GS_BD_DesSalida
     GL_Cod_Uni = CLng(GL_AR_BD_DatosSalida(0, 0))
     GL_Num_Uni = CLng(GL_AR_BD_DatosSalida(0, 1))
    'Despues del tratamiento del acceso, liberamos la memoria ocupada
    'por el array que contiene los datos de salida de la consulta.
     ReDim GL_AR_BD_DatosSalida(0, 0) As Long
     Exit Sub


End Sub

Sub s_mostrar_info_gai()

End Sub

Sub s_grabar_resumen_gai()

End Sub
Sub s_inicializar_gai()

End Sub
