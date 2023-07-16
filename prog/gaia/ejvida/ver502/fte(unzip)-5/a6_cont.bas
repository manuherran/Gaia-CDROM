Attribute VB_Name = "bas_a6_control"
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

Sub S_S2_IfIgualThenGoto()
        'IF p1 = p2 THEN GOTO p3-p4 ELSE GOTO p5-p6
        'Lee la variable P1 y P2
        'Si son iguales, marca la siguiente instrucci�n como la P3
        S_Leer_Parametros
        If GV_Param(1) = GV_Param(2) Then
            GL_Cod_Acc = GV_Param(3)
            GL_Num_Repetida = GV_Param(4)
        Else
            GL_Cod_Acc = GV_Param(5)
            GL_Num_Repetida = GV_Param(6)
        End If

End Sub

Sub S_S4_IfMenorThenGoto()
        'IF p1 < p2 THEN GOTO p3-p4 ELSE GOTO p5-p6
        'Lee la variable P1 y P2
        'Si son iguales, marca la siguiente instrucci�n como la P3
        S_Leer_Parametros
        If GV_Param(1) < GV_Param(2) Then
            GL_Cod_Acc = GV_Param(3)
            GL_Num_Repetida = GV_Param(4)
        Else
            GL_Cod_Acc = GV_Param(5)
            GL_Num_Repetida = GV_Param(6)
        End If

End Sub

Sub S_S3_IfMayorThenGoto()
        'IF p1 > p2 THEN GOTO p3-p4 ELSE GOTO p5-p6
        'Lee la variable P1 y P2
        'Si son iguales, marca la siguiente instrucci�n como la P3
        S_Leer_Parametros
        If GV_Param(1) > GV_Param(2) Then
            GL_Cod_Acc = GV_Param(3)
            GL_Num_Repetida = GV_Param(4)
        Else
            GL_Cod_Acc = GV_Param(5)
            GL_Num_Repetida = GV_Param(6)
        End If

End Sub


Sub S_S6_IfMenorOIgualThenGoto()
        'IF p1 <= p2 THEN GOTO p3-p4 ELSE GOTO p5-p6
        'Lee la variable P1 y P2
        'Si son iguales, marca la siguiente instrucci�n como la P3
        S_Leer_Parametros
        If GV_Param(1) <= GV_Param(2) Then
            GL_Cod_Acc = GV_Param(3)
            GL_Num_Repetida = GV_Param(4)
        Else
            GL_Cod_Acc = GV_Param(5)
            GL_Num_Repetida = GV_Param(6)
        End If

End Sub

Sub S_S5_IfMayorOIgualThenGoto()
        'IF p1 >= p2 THEN GOTO p3-p4 ELSE GOTO p5-p6
        'Lee la variable P1 y P2
        'Si son iguales, marca la siguiente instrucci�n como la P3
        S_Leer_Parametros
        If GV_Param(1) >= GV_Param(2) Then
            GL_Cod_Acc = GV_Param(3)
            GL_Num_Repetida = GV_Param(4)
        Else
            GL_Cod_Acc = GV_Param(5)
            GL_Num_Repetida = GV_Param(6)
        End If

End Sub


Sub S_Ejecutar_Accion_Control()

Select Case GL_Tip
    Case 2
        S_S2_IfIgualThenGoto
    Case 3
        S_S3_IfMayorThenGoto
    Case 4
        S_S4_IfMenorThenGoto
    Case 5
        S_S5_IfMayorOIgualThenGoto
    Case 6
        S_S6_IfMenorOIgualThenGoto
        
        
    Case Else
        MsgBox "Error: Acci�n de control no existe"
        End

End Select

End Sub
