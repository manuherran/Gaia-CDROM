Attribute VB_Name = "bas_z0_time"
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



'Control de tiempos: inicial
Global Gv_now_Comienzo_ejv As Variant
Global Gv_YYYY_Comienzo_ejv As Variant
Global Gv_MM_Mes_Comienzo_ejv As Variant
Global Gv_DD_Comienzo_ejv As Variant
Global Gv_HH_Comienzo_ejv As Variant
Global Gv_MM_Min_Comienzo_ejv As Variant
Global Gv_SS_Comienzo_ejv As Variant

'Control de tiempos: final
Global Gv_now_Fin_ejv As Variant
Global Gv_YYYY_Fin_ejv As Variant
Global Gv_MM_Mes_Fin_ejv As Variant
Global Gv_DD_Fin_ejv As Variant
Global Gv_HH_Fin_ejv As Variant
Global Gv_MM_Min_Fin_ejv As Variant
Global Gv_SS_Fin_ejv As Variant

Sub s_leer_tiempo_inicial_ejv()

    Gv_now_Comienzo_ejv = Now
    
    Gv_YYYY_Comienzo_ejv = Year(Date)
    Gv_MM_Mes_Comienzo_ejv = Month(Date)
    Gv_DD_Comienzo_ejv = Day(Date)
    Gv_HH_Comienzo_ejv = Hour(Time)
    Gv_MM_Min_Comienzo_ejv = Minute(Time)
    Gv_SS_Comienzo_ejv = Second(Time)

End Sub

Sub s_leer_tiempo_final_ejv()

    Gv_now_Fin_ejv = Now
    
    Gv_YYYY_Fin_ejv = Year(Date)
    Gv_MM_Mes_Fin_ejv = Month(Date)
    Gv_DD_Fin_ejv = Day(Date)
    Gv_HH_Fin_ejv = Hour(Time)
    Gv_MM_Min_Fin_ejv = Minute(Time)
    Gv_SS_Fin_ejv = Second(Time)


End Sub
    
Sub s_mostrar_tiempo_transcurrido_ejv()

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

    'Calculo el numero de segundos totales
    seg_ej_actual_ejv = DateDiff("s", Gv_now_Comienzo_ejv, Gv_now_Fin_ejv, vbMonday, vbFirstJan1)
    form_prg_activo.SegTot.Caption = seg_ej_actual_ejv
    

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

Sub s_borrar_tiempo_comienzo()

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

    form_prg_activo.fechaf = ""
    form_prg_activo.horaf = ""
    form_prg_activo.anosr = ""
    form_prg_activo.mesesr = ""
    form_prg_activo.diasr = ""
    form_prg_activo.horasr = ""
    form_prg_activo.minutosr = ""
    form_prg_activo.segundosr = ""

    Gv_SS_Comienzo_ejv = -1

End Sub

Sub s_mostrar_fecha_hora_actual_ejv()

    'La fecha actual y la de fin es la misma
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

    form_prg_activo.fechaf = Gv_YYYY_Fin_ejv & "-" & Gv_MM_Mes_Fin_ejv & "-" & Gv_DD_Fin_ejv
    form_prg_activo.horaf = Gv_HH_Fin_ejv & ":" & Gv_MM_Min_Fin_ejv & ":" & Gv_SS_Fin_ejv


End Sub
Sub s_mostrar_fecha_hora_comienzo_ejv()

    'La fecha actual y la de fin es la misma
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

    form_prg_activo.fechac = Gv_YYYY_Comienzo_ejv & "-" & Gv_MM_Mes_Comienzo_ejv & "-" & Gv_DD_Comienzo_ejv
    form_prg_activo.horac = Gv_HH_Comienzo_ejv & ":" & Gv_MM_Min_Comienzo_ejv & ":" & Gv_SS_Comienzo_ejv


End Sub

