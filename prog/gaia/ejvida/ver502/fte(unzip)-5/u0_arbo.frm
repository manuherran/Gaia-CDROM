VERSION 5.00
Begin VB.Form frm_u0_arbo 
   Caption         =   "Arbol"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4575
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   4575
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Aceptar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Btn_Arbol 
      Caption         =   "&Mostrar Arbol"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Etiqueta"
      Height          =   195
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frm_u0_arbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim cont_mostrados As Long
Dim mi_left As Long
Dim mi_top As Long

Dim avance_left As Long
Dim avance_top As Long

Dim media_caja As Long

Dim minimo_punto As Long

Private Sub Aceptar_Click()
    Unload Me

End Sub

Private Sub Btn_Arbol_Click()

    Dim i As Long
    

    Btn_Arbol.Visible = False
    Aceptar.Visible = False
    
    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_UTIL)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Fichero de árbol"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Abrir"
    frm_z0_fic.File1.Pattern = "*.txt"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If Not cancelar_operacion_fichero_ejv Then
        On Error Resume Next
        s_aut_leer_arbol
        cont_mostrados = 0
        mi_top = avance_top
        
        Screen.MousePointer = CTE_ARENA
        For i = 1 To UBound(cod_arb)
            If cod_padre_arb(i) = "" Then
                mi_left = avance_left
                f_expandir_nodo i
                DoEvents
            End If
        Next i
        Screen.MousePointer = CTE_DEFECTO
    End If
    
    
End Sub

Private Sub Form_Activate()

    Me.BackColor = cct_ejv(cfondo_ejv)
    ajuste_color_controles_formulario_ejv Me

End Sub

Private Sub Form_GotFocus()
    
    Me.BackColor = cct_ejv(cfondo_ejv)
    ajuste_color_controles_formulario_ejv Me

End Sub

Private Sub Form_Load()
    
    'Refresco automatico, consume muchos recursos
    'frm_u0_arbo.AutoRedraw = True
    frm_u0_arbo.AutoRedraw = False
    frm_u0_arbo.WindowState = CTE_MAXIMIZED
    
    frm_u0_arbo.DrawWidth = 2
    
    avance_left = 1200
    avance_top = 300
    
    media_caja = 120
    minimo_punto = 30
    
End Sub

Function f_expandir_nodo(nodo As Long) As Boolean
    
    Dim viejo_left As Long
    Dim numero_hijos As Long

    s_mostrar_etiqueta nodo
    viejo_left = mi_left
    numero_hijos = f_mostrar_hijos(nodo)
    mi_left = viejo_left

    If numero_hijos > 0 Then
        f_expandir_nodo = True
    Else
        f_expandir_nodo = False
    End If

End Function


Sub s_mostrar_etiqueta(nodo As Long)

    cont_mostrados = cont_mostrados + 1
    Load Etiqueta(cont_mostrados)
    Etiqueta(cont_mostrados).Left = mi_left
    Etiqueta(cont_mostrados).Top = mi_top
    Etiqueta(cont_mostrados).Visible = True
    Etiqueta(cont_mostrados).Caption = desc_arb(nodo)
    

End Sub

Function f_mostrar_hijos(nodo As Long) As Long

    Dim i As Long
    Dim primer_hijo As Boolean
    Dim cont_hijos As Long
    Dim pto_x As Long
    Dim pto_y As Long
    Dim ha_habido_hijos As Boolean
    
    primer_hijo = True
    cont_hijos = 0
    
    For i = 1 To UBound(cod_arb)
        If cod_padre_arb(i) = cod_arb(nodo) Then
            If primer_hijo Then
                cont_hijos = 1
                mi_left = mi_left + avance_left
                Me.Line (mi_left - avance_left + 20, mi_top + media_caja)-(mi_left, mi_top + media_caja), 0
            Else
                Me.Line (mi_left - avance_left / 4, mi_top + media_caja)-(mi_left, mi_top + media_caja), 0
                pto_y = mi_top + media_caja - minimo_punto
                pto_x = mi_left - avance_left / 4
                While Point(pto_x, pto_y) <> 0
                    PSet (pto_x, pto_y), 0
                    pto_y = pto_y - minimo_punto
                Wend
                PSet (pto_x, pto_y), 0
                'Me.Line (mi_left - avance_left / 4, mi_top + media_caja)-(mi_left - avance_left / 4, mi_top - (avance_top - media_caja - media_caja / 2) * cont_hijos - media_caja / 2), 0
                cont_hijos = cont_hijos + 1
            End If
            ha_habido_hijos = f_expandir_nodo(i)
            'Solo avanzo si ha tenido hijos
            If Not ha_habido_hijos Then
                mi_top = mi_top + avance_top
            End If
            primer_hijo = False
        End If
    Next i

    f_mostrar_hijos = cont_hijos

End Function

