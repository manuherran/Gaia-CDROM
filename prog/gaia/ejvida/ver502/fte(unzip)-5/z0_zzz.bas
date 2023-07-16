Attribute VB_Name = "bas_z0_zzz"
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

Global numero_de_elem_dimension1 As Integer
Global numero_de_elem_dimension2 As Integer
Global numero_de_elem_dimension3 As Integer
Global numero_de_elem_dimension4 As Integer

Global mi_array() As String
Global objeto_excel As Object

Global config_min_HYM() As Integer 'minimo
Global config_max_HYM() As Integer 'maximo
Global config_sal_HYM() As Integer 'salto

Function dame_dimensiones_dato_HYM(posicion As Integer, dimension2 As Integer, dimension3 As Integer, dimension4 As Integer, dimension1 As String) As Boolean
    
    Dim temp As Long
    
    'El primer parametro es de entrada
    'Los 3 siguientes parametros son de salida
    dimension4 = f_SumCirc(numero_de_elem_dimension4, posicion, 0)
    
    temp = 1 + Int((posicion - 1) / (numero_de_elem_dimension4))
    dimension3 = f_SumCirc(numero_de_elem_dimension3, temp + numero_de_elem_dimension3, 0)
    'dimension3 = f_SumCirc(numero_de_elem_dimension3, f_dividir_por_encima_HYM(posicion / numero_de_elem_dimension4), 0)
    
    dimension2 = 1 + Int((posicion - 1) / (numero_de_elem_dimension3 * numero_de_elem_dimension4))
    'dimension2 = f_SumCirc(numero_de_elem_dimension3, posicion / (numero_de_elem_dimension3 * numero_de_elem_dimension4), 0)


End Function
Function dame_posicion_dato_HYM(dimension2 As Integer, dimension3 As Integer, dimension4 As Integer) As Integer

    dame_posicion_dato_HYM = 0
    dame_posicion_dato_HYM = dame_posicion_dato_HYM + ((dimension2 - 1) * numero_de_elem_dimension3 * numero_de_elem_dimension4)
    dame_posicion_dato_HYM = dame_posicion_dato_HYM + ((dimension3 - 1) * numero_de_elem_dimension4)
    dame_posicion_dato_HYM = dame_posicion_dato_HYM + dimension4

End Function


Function f_dame_fila_HYM(dimension2 As String, dimension3 As String, dimension1 As String) As Integer

    Dim exito As Boolean
    Dim fila As Integer
    
    Dim aux1, aux2, aux3, aux4, aux5, aux6 As String
    
       
    exito = False
    fila = 2
    aux1 = mi_array(1, fila)
    aux2 = mi_array(2, fila)
    aux3 = mi_array(3, fila)
    Do Until aux1 = "" And aux4 = ""
        If aux1 = dimension2 And aux2 = dimension3 And aux3 = dimension1 Then
            exito = True
            Exit Do
        End If
        fila = fila + 1
        aux1 = mi_array(1, fila)
        aux2 = mi_array(2, fila)
        aux3 = mi_array(3, fila)
        
        aux4 = mi_array(1, fila + 1)
    Loop

    If exito Then
        f_dame_fila_HYM = fila
    Else
        MsgBox "error"
        f_dame_fila_HYM = -1
    End If

End Function


Function f_crear_dato_azar_con_pos_HYM(posicion As Integer) As Long

    'Siempre que se pueda NO usar esta funcion, y usar en cambio
    'la funcion f_crear_dato_azar_con_dimHYM

    Dim dimension2 As Integer
    Dim dimension3 As Integer
    Dim dimension4 As Integer
    Dim dimension1 As String

    '----------------------------------------------
    'Sabido posicion que es la posición del dato
    'calculo su dimension2, dimension3, dimension4
    'dimension2, dimension3, dimension4 son parametros de salida
    dame_dimensiones_dato_HYM posicion, dimension2, dimension3, dimension4, dimension1
    f_crear_dato_azar_con_pos_HYM = f_crear_dato_azar_con_dimHYM(dimension2, dimension3, dimension4, dimension1)
    '----------------------------------------------

End Function

Function f_crear_dato_azar_con_dimHYM(dimension2 As Integer, dimension3 As Integer, dimension4 As Integer, dimension1 As String) As Long

    'Siempre que se pueda SI usar esta funcion en vez de
    'usar la funcion f_crear_dato_azar_con_pos_HYM

    Dim min As Long
    Dim max As Long
    Dim sal As Long
    
    Dim num_valores As Integer
    
    sal = config_sal_HYM(dimension2, dimension3)
    If sal = 0 Then
        f_crear_dato_azar_con_dimHYM = 0
    Else
        min = config_min_HYM(dimension2, dimension3)
        max = config_max_HYM(dimension2, dimension3)
        num_valores = Int((max - min) / sal)
        f_crear_dato_azar_con_dimHYM = min + (fi_azar1(num_valores + 1) - 1) * sal
    End If

End Function


Function f_dame_fila_busca_en_xls_HYM(dimension2 As String, dimension3 As String, dimension1 As String) As Integer

    Dim exito As Boolean
    Dim fila As Integer
    
    Dim aux1, aux2, aux3, aux4, aux5, aux6 As String
    
    aux1 = objeto_excel.ActiveSheet.Cells(2, 1).Value
    aux2 = objeto_excel.ActiveSheet.Cells(2, 2).Value
    aux3 = objeto_excel.ActiveSheet.Cells(2, 3).Value
       
    exito = False
    fila = 2
    Do Until aux1 = "" And aux2 = "" And aux3 = "" And aux4 = "" And aux5 = "" And aux6 = ""
        If aux1 = dimension2 And aux2 = dimension3 And aux3 = dimension1 Then
            exito = True
            Exit Do
        End If
        fila = fila + 1
        aux1 = objeto_excel.ActiveSheet.Cells(fila, 1).Value
        aux2 = objeto_excel.ActiveSheet.Cells(fila, 2).Value
        aux3 = objeto_excel.ActiveSheet.Cells(fila, 3).Value
        
        aux4 = objeto_excel.ActiveSheet.Cells(fila + 1, 1).Value
        aux5 = objeto_excel.ActiveSheet.Cells(fila + 1, 2).Value
        aux6 = objeto_excel.ActiveSheet.Cells(fila + 1, 3).Value
            
    Loop

    If exito Then
        f_dame_fila_busca_en_xls_HYM = fila
    Else
        MsgBox "error"
        f_dame_fila_busca_en_xls_HYM = -1
    End If

End Function

Sub s_cargar_datos_en_mem_HYM()
    
    Dim celda As String
    Dim f As Integer
    Dim c As Integer
    Dim celda_f_anterior As String
    Dim continuar As Boolean
    
    Screen.MousePointer = CTE_ARENA
    
    'Leo la primera f
    f = 2
    ReDim Preserve mi_array(1 To 18, 1 To f) As String
    For c = 1 To 18
        mi_array(c, f) = objeto_excel.ActiveSheet.Cells(f, c).Value
    Next c
    celda_f_anterior = mi_array(1, f)
    
    continuar = True
    While continuar
        f = f + 1
        'Leo la f
        ReDim Preserve mi_array(1 To 18, 1 To f) As String
        For c = 1 To 18
            mi_array(c, f) = objeto_excel.ActiveSheet.Cells(f, c).Value
        Next c
        If mi_array(1, f) = "" And celda_f_anterior = "" Then
            continuar = False
        Else
            celda_f_anterior = mi_array(1, f)
        End If
    Wend
    
    Screen.MousePointer = CTE_DEFECTO
    

End Sub

Sub s_llama_help(MiForm As Form, MiHLP As String)


     Dim I_i        As Integer
     Dim S_DummyVal As String
    
    'Inicializamos las variables utilizadas
     I_i = 0
     S_DummyVal = " "

    'Funcion SDK que nos permite abrir un fichero de ayuda especificado
     'I_i = WinHelp(MiForm.hWnd, MiHLP, HELP_PARTIALKEY, S_DummyVal)

End Sub

Sub s_mostrar_tiempo_transc_metodo_viejo_ejv()

    Dim s_media As String
    Dim d_media As Double
    
    Dim form_prg_activo As Object
    
    'Elijo el form actual=========Este código está repetido=========
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            Set form_prg_activo = frm_a1_inhyp
        Case CTE_PAL '2
            Set form_prg_activo = frm_b2_inpal
        Case CTE_3R '3
            Set form_prg_activo = frm_c3_in3r
        Case CTE_PRI '4
            Set form_prg_activo = frm_a4_inpri
        Case CTE_CEL '5
            Set form_prg_activo = frm_a5_incel
        Case CTE_GAI '6
            Set form_prg_activo = frm_a6_ingaia
        Case CTE_EXP '7
            Set form_prg_activo = frm_a7_inexp
        Case CTE_CAD '8
            Set form_prg_activo = frm_c8_incad
        Case CTE_PEZ '9
            Set form_prg_activo = frm_a9_inpez
        Case CTE_UVA '10
            Set form_prg_activo = frm_aA_inuva
        Case CTE_YXY '11
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select

    s_leer_tiempo_final_ejv
    form_prg_activo.fechaf = Gv_YYYY_Fin_ejv & "-" & Gv_MM_Mes_Fin_ejv & "-" & Gv_DD_Fin_ejv
    form_prg_activo.horaf = Gv_HH_Fin_ejv & ":" & Gv_MM_Min_Fin_ejv & ":" & Gv_SS_Fin_ejv
    
    'Cálculos
    form_prg_activo.segundosr = Gv_SS_Fin_ejv - Gv_SS_Comienzo_ejv
    form_prg_activo.minutosr = Gv_MM_Min_Fin_ejv - Gv_MM_Min_Comienzo_ejv
    form_prg_activo.horasr = Gv_HH_Fin_ejv - Gv_HH_Comienzo_ejv
    form_prg_activo.diasr = Gv_DD_Fin_ejv - Gv_DD_Comienzo_ejv
    form_prg_activo.mesesr = Gv_MM_Mes_Fin_ejv - Gv_MM_Mes_Comienzo_ejv
    form_prg_activo.anosr = Gv_YYYY_Fin_ejv - Gv_YYYY_Comienzo_ejv
    
    'Ajustes
    If form_prg_activo.segundosr < 0 Then
        'form_prg_activo.segundosr = 60 + form_prg_activo.segundosr
        form_prg_activo.segundosr = Gv_SS_Comienzo_ejv + form_prg_activo.segundosr
        form_prg_activo.minutosr = CInt(form_prg_activo.minutosr) - 1
    End If
    If form_prg_activo.minutosr < 0 Then
        'form_prg_activo.minutosr = 60 + form_prg_activo.minutosr
        form_prg_activo.minutosr = Gv_MM_Min_Comienzo_ejv + form_prg_activo.minutosr
        form_prg_activo.horasr = form_prg_activo.horasr - 1
    End If
    If form_prg_activo.horasr < 0 Then
        'form_prg_activo.horasr = 24 + form_prg_activo.horasr
        form_prg_activo.horasr = Gv_HH_Comienzo_ejv + form_prg_activo.horasr
        form_prg_activo.diasr = form_prg_activo.diasr - 1
    End If
    If form_prg_activo.diasr < 0 Then
        'form_prg_activo.diasr = 30 + form_prg_activo.diasr
        form_prg_activo.diasr = Gv_DD_Comienzo_ejv + form_prg_activo.diasr
        form_prg_activo.mesesr = form_prg_activo.mesesr - 1
    End If
    If form_prg_activo.mesesr < 0 Then
        'form_prg_activo.mesesr = 12 + form_prg_activo.mesesr
        form_prg_activo.mesesr = Gv_MM_Mes_Comienzo_ejv + form_prg_activo.mesesr
        form_prg_activo.anosr = form_prg_activo.mesesr - 1
    End If

    'Calculo el numero de segundos que ha tardado cada ciclo
    seg_ej_actual_ejv = 0
    seg_ej_actual_ejv = seg_ej_actual_ejv + form_prg_activo.segundosr
    seg_ej_actual_ejv = seg_ej_actual_ejv + form_prg_activo.minutosr * 60
    seg_ej_actual_ejv = seg_ej_actual_ejv + form_prg_activo.horasr * 3600
    seg_ej_actual_ejv = seg_ej_actual_ejv + form_prg_activo.diasr * 24 * 3600
    seg_ej_actual_ejv = seg_ej_actual_ejv + form_prg_activo.mesesr * 30 * 24 * 3600
    seg_ej_actual_ejv = seg_ej_actual_ejv + form_prg_activo.anosr * 12 * 30 * 24 * 3600


    'Calculo la media de segundos por ciclo
    If ciclo_ejv > 0 Then
        d_media = seg_ej_actual_ejv / ciclo_ejv
        s_media = Left(d_media, 10)
        If d_media < 0.1 Then
            s_media = Format(d_media, "0.0000000")
        End If
        form_prg_activo.media.Caption = s_media
    End If

End Sub

