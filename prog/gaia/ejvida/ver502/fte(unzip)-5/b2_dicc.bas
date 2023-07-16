Attribute VB_Name = "bas_b2_dicc"
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


Sub s_mostrar_diccionario_pal()
 
    Dim i As Integer
    Dim txt As String
    Dim pal As String
    
    Screen.MousePointer = CTE_ARENA
    
    frm_z0_lista.Caption = "Diccionario"
    
    txt = ""
    For i = 1 To numero_palabras_dicc_pal
        pal = CStr(i) & ".- "
        While Len(pal) < 7
            pal = " " & pal
        Wend
        pal = pal & palabra_del_diccionario(i)
        txt = txt & pal & vbCrLf
    Next i
    frm_z0_lista.txt_lista.Text = txt
    
    Screen.MousePointer = CTE_DEFECTO

End Sub

