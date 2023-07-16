Attribute VB_Name = "bas_u0_util"
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
'Arboles
Global cod_arb() As String
Global desc_arb() As String
Global cod_padre_arb() As String


'Backup
Global directorio_a_comp_bk() As String
Global nombres_ficheros_backup_bk() As String


'Teclas
Global fic_teclas_completo_tec As String
Global fic_salida_completo_tec As String
Global l_identificativo_exe_tec As Long
Global ejecutable_tec As String
Global titulo_tec As String
Global hay_que_lanzar_tec As Boolean
Global iteraciones_tec As Long


Sub s_lanzar_teclas()
    
    On Error GoTo fin
    
    Dim linea As String
    
    Dim inicio() As String
    Dim bucle() As String
    Dim fin() As String
    
    Dim secuencia As String
    Dim cont As Long
    Dim it As Long
    
    fic_teclas_completo_tec = f_nombre_completo(path_largo_ejv(CTE_C_PRG_UTIL), "teclas.txt")
    fic_salida_completo_tec = f_nombre_completo(path_largo_ejv(CTE_C_PRG_UTIL), "salida.txt")
    
    Open fic_teclas_completo_tec For Input As #CTE_FIC_10_TECLAS_CFG
    
    'Leo el exe
    linea = f_leer_linea(CTE_FIC_10_TECLAS_CFG)
    ejecutable_tec = Right(linea, Len(linea) - Len("Exe="))
    
    'Leo el título
    linea = f_leer_linea(CTE_FIC_10_TECLAS_CFG)
    titulo_tec = Right(linea, Len(linea) - Len("Titulo="))
    
    'Leo si hay que abrir el ejecutable o esta ya abierto
    linea = f_leer_linea(CTE_FIC_10_TECLAS_CFG)
    If UCase(Right(linea, Len(linea) - Len("LlamarEjecutable="))) = UCase(CTE_txtFALSE) Then
        hay_que_lanzar_tec = False
    Else
        hay_que_lanzar_tec = True
    End If
    
    'Leo el número de iteraciones del bucle
    linea = f_leer_linea(CTE_FIC_10_TECLAS_CFG)
    iteraciones_tec = CLng(Right(linea, Len(linea) - Len("Iteraciones=")))
    
    'Leo la secuencia de inicio
    f_posicionarse_en_seccion "[Inicio]", CTE_FIC_10_TECLAS_CFG, fic_teclas_completo_tec
    f_cargar_seccion_actual inicio(), CTE_FIC_10_TECLAS_CFG, fic_teclas_completo_tec
    
    'Leo la secuencia a repetir
    f_posicionarse_en_seccion "[Bucle]", CTE_FIC_10_TECLAS_CFG, fic_teclas_completo_tec
    f_cargar_seccion_actual bucle(), CTE_FIC_10_TECLAS_CFG, fic_teclas_completo_tec
    
    'Leo la secuencia de fin
    f_posicionarse_en_seccion "[Fin]", CTE_FIC_10_TECLAS_CFG, fic_teclas_completo_tec
    f_cargar_seccion_actual fin(), CTE_FIC_10_TECLAS_CFG, fic_teclas_completo_tec
    
    Close #CTE_FIC_10_TECLAS_CFG
    
    If hay_que_lanzar_tec Then
        s_lanzar_y_activar_ejecutable ejecutable_tec, titulo_tec
    Else
        s_activar_ejecutable_ya_lanzado_por_titulo titulo_tec
    End If
    
    Open fic_salida_completo_tec For Output As #CTE_FIC_11_TECLAS_SAL
    Screen.MousePointer = CTE_ARENA
    
    For cont = 1 To UBound(inicio)
        s_genera_tecla inicio(cont)
    Next cont
    
    For it = 1 To iteraciones_tec
        For cont = 1 To UBound(bucle)
            s_genera_tecla bucle(cont)
        Next cont
        s_genera_tecla "<<GUARDAR_FICHERO_SALIDA>>"
    Next it
    
    For cont = 1 To UBound(fin)
        s_genera_tecla fin(cont)
    Next cont
    
    Close #CTE_FIC_11_TECLAS_SAL
    Screen.MousePointer = CTE_DEFECTO

fin:

End Sub

Sub s_genera_tecla(tecla As String)

    'Si se usan argumemntos con nombres, se pueden dar en cualquier orden
    'SendKeys String:="{TAB}", WAIT:=True
    'SendKeys "{TAB}", True
    
    Dim tmp As String
    Dim v_seg As Variant
    
    Select Case tecla
        Case ""
            'No hago nada
        Case "<<BORRAR_PORTAPAPELES>>"
            Clipboard.SetText ""
        Case "<<INCLUIR_#_EN_EL_PORTAPAPELES>>"
            tmp = Clipboard.GetText(vbCFText)
            Clipboard.SetText tmp & "#"
        Case "<<INCLUIR_;_EN_EL_PORTAPAPELES>>"
            tmp = Clipboard.GetText(vbCFText)
            Clipboard.SetText tmp & ";"
        Case "<<COPIAR_ACUMULANDO_EN_EL_PORTAPAPELES>>"
            tmp = Clipboard.GetText(vbCFText)
            SendKeys "^C", True
            Clipboard.SetText tmp & Clipboard.GetText(vbCFText)
        Case "<<ESCRIBIR_EN_SALIDA_EL_PORTAPAPELES>>"
            Print #CTE_FIC_11_TECLAS_SAL, Clipboard.GetText(vbCFText)
        Case "<<GUARDAR_FICHERO_SALIDA>>"
            Close #CTE_FIC_11_TECLAS_SAL
            Open fic_salida_completo_tec For Append As #CTE_FIC_11_TECLAS_SAL
        Case "<<SALTO_DE_LINEA>>"
            Print #CTE_FIC_11_TECLAS_SAL, ""
        Case "<<RETARDO_1_SEGUNDO>>"
            v_seg = Second(Time)
            While Second(Time) = v_seg
            Wend
        Case "<<RETARDO_5_SEGUNDOS>>"
            v_seg = f_SumCirc(60, Second(Time), 5)
            While Second(Time) <> v_seg
            Wend
        Case Else
            SendKeys tecla, True
    End Select
        
    
    
    'SendKeys "1", True
    'SendKeys "+{END}", True
    'SendKeys "^C", True
    'Print #CTE_FIC_11_TECLAS_SAL, Clipboard.GetText(vbCFText)
    'Print #CTE_FIC_11_TECLAS_SAL, ""
    'Print #CTE_FIC_11_TECLAS_SAL, Clipboard.GetText(vbCFText)
    'SendKeys "{ENTER}", True
    'SendKeys "%{F4}", True  ' Send ALT+F4


End Sub

Sub s_lanzar_y_activar_ejecutable(ejecutable As String, titulo As String)

On Error GoTo error

    l_identificativo_exe_tec = Shell(ejecutable, vbNormalNoFocus)
    AppActivate l_identificativo_exe_tec, False

    Exit Sub
error:
    'Se puede elegir el titulo de la ventana o el valor de retorno
    'Esto es por si ya esta abierto el programa, entonces no se tiene el
    'l_identificativo_exe_tec pero si se puede hacer referencia al titulo
    AppActivate titulo, False

End Sub

Sub s_activar_ejecutable_ya_lanzado_por_titulo(titulo As String)

    AppActivate titulo, False

End Sub

Sub s_activar_ejecutable_ya_lanzado_por_ident(ident As Long)

    AppActivate ident, False

End Sub

Function f_es_palabra_normal(palabra) As Boolean

    f_es_palabra_normal = True
    If InStr(palabra, "http") Or InStr(palabra, "HTTP") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "ftp") Or InStr(palabra, "FTP") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "www") Or InStr(palabra, "WWW") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "@") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "0") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "1") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "2") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "3") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "4") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "5") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "6") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "7") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "8") Then
        f_es_palabra_normal = False
    End If
    If InStr(palabra, "9") Then
        f_es_palabra_normal = False
    End If

End Function

Sub s_video()

    automatico_ejv = False
    s_click_programa_ejv CTE_HYP
    's_aceptar_menu_ejv "Ejemplo 1", 1
    automatico_ejv = True
    s_operacion_ejecutar_ejv CTE_EXE_COMENZAR

End Sub
