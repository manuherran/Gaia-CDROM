VERSION 5.00
Begin VB.Form frm_a0_mapa 
   Caption         =   "Mapa de Obstáculos"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "A0_MAPA.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8400
   ScaleWidth      =   11880
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton test_movimiento 
      Caption         =   "Test &Movimiento..."
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Rellenar 
      Caption         =   "&Rellenar Todo"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Vaciar 
      Caption         =   "&Vaciar Todo"
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton invertir 
      Caption         =   "&Invertir Todo"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton bordeObs 
      Caption         =   "&Borde Obstaculos"
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Añadir Figura"
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   3255
      Begin VB.TextBox Op_MapaEjeZ 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox Op_MapaTipo 
         Height          =   315
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Op_MapaLong 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Op_MapaEjeX 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Op_MapaEjeY 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   17
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox Op_MapaFigura 
         Height          =   315
         Left            =   720
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Aniadir 
         Caption         =   "Añadir"
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "v"
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "^"
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Eje Z"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Piso"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Columna"
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Fila"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Longitud"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Eje X"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Eje Y"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Figura"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dimensiones del Mapa"
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      Begin VB.TextBox Op_MapaMaxEjeZ 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Ch_Parpadeo 
         Caption         =   "Parpadeo"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Cambiar 
         Caption         =   "Cambiar &Dimensiones"
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Op_MapaMaxEjeX 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Text            =   "20"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Op_MapaMaxEjeY 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Text            =   "20"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº de pisos (Z)"
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº de columnas (X)"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº de filas (Y)"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton derecha 
      Caption         =   ">"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton izquierda 
      Caption         =   "<"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton abajo 
      Caption         =   "v"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton arriba 
      Caption         =   "^"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2400
      Top             =   5760
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_a0_mapa"
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

Private Sub abajo_Click()
    
    s_mover_cursor CTE_ATRAS
    
End Sub

Private Sub Aceptar_Click()
    'Lo copio sobre los arrays de va0 si he venido por va0, asi modifico el
    'mapa actual de va0 y si no, (si estoy en el editor de mapas sin mas) no
    If copiar_mapa_a_va0_ma0 Then
        s_copiar_mapa_ma0_sobre_va0_va0
    End If
    copiar_mapa_a_va0_ma0 = False
    Unload Me

End Sub

Private Sub Aniadir_Click()
    
    If Op_MapaEjeZ > mapa_pisos_ma0 Then Op_MapaEjeZ = mapa_pisos_ma0
    If Op_MapaEjeZ < 1 Then Op_MapaEjeZ = 1
    
    If Op_MapaEjeY > mapa_columnas_ma0 Then Op_MapaEjeY = mapa_columnas_ma0
    If Op_MapaEjeY < 1 Then Op_MapaEjeY = 1
    
    If Op_MapaEjeX > mapa_filas_ma0 Then Op_MapaEjeX = mapa_filas_ma0
    If Op_MapaEjeX < 1 Then Op_MapaEjeX = 1
    
    s_mapa_pintar_figura_va0 Op_MapaTipo, Op_MapaFigura, Op_MapaLong, Op_MapaEjeZ, Op_MapaEjeY, Op_MapaEjeX
End Sub

Private Sub arriba_Click()

    s_mover_cursor CTE_DEFRENTE

End Sub

Private Sub bordeObs_Click()
    
    Dim p As Double
    Dim f As Double
    Dim c As Double

    Screen.MousePointer = CTE_ARENA
    'frm_a0_mapa.Cls

    For p = 1 To mapa_pisos_ma0
        For f = 1 To mapa_filas_ma0
            c = 1
            mapa_ma0(p, f, c) = True
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            c = mapa_columnas_ma0
            mapa_ma0(p, f, c) = True
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        Next f
    
        For c = 1 To mapa_columnas_ma0
            f = 1
            mapa_ma0(p, f, c) = True
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            f = mapa_filas_ma0
            mapa_ma0(p, f, c) = True
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        Next c
    Next p

    Screen.MousePointer = CTE_DEFECTO


End Sub

Private Sub Cambiar_Click()

If habilitar_change_zoom_va0 Then
    
    ReDim tmp(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Boolean
    Dim p As Integer
    Dim f As Integer
    Dim c As Integer
    Dim se_acabo As Boolean
    
    'Copio el mapa viejo a un temporal
    For p = 1 To mapa_pisos_ma0
    For f = 1 To mapa_filas_ma0
    For c = 1 To mapa_columnas_ma0
        tmp(p, f, c) = mapa_ma0(p, f, c)
    Next c
    Next f
    Next p
    
    'Redimensiono el nuevo mapa
    mapa_pisos_ma0 = CInt("0" & Op_MapaMaxEjeZ)
    If mapa_pisos_ma0 = 0 Then mapa_pisos_ma0 = 1
    
    mapa_filas_ma0 = CInt("0" & Op_MapaMaxEjeY)
    If mapa_filas_ma0 = 0 Then mapa_filas_ma0 = 1
    
    mapa_columnas_ma0 = CInt("0" & Op_MapaMaxEjeX)
    If mapa_columnas_ma0 = 0 Then mapa_columnas_ma0 = 1
    
    ReDim mapa_ma0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Boolean
    ReDim nodo_visitado_va0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Integer
    
    'Copio lo que ya tenía
    se_acabo = False
    For p = 1 To UBound(tmp, 1)
    For f = 1 To UBound(tmp, 2)
    For c = 1 To UBound(tmp, 3)
        If p <= mapa_pisos_ma0 Then
            If f <= mapa_filas_ma0 Then
                If c <= mapa_columnas_ma0 Then
                    mapa_ma0(p, f, c) = tmp(p, f, c)
                Else
                    Exit For
                End If
            Else
                se_acabo = True
            End If
        Else
            se_acabo = True
        End If
    Next c
    If se_acabo Then Exit For
    Next f
    If se_acabo Then Exit For
    Next p
    
    'Muestro el nuevo mapa
    s_refrescar_mapa_actual_ma0
    
End If

End Sub

Private Sub Cancelar_Click()
    Unload Me
End Sub

Private Sub Ch_Parpadeo_Click()

    If Ch_Parpadeo.Value = 1 Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If

End Sub

Private Sub Command5_Click()
    Op_MapaLong = Op_MapaLong - 1

End Sub

Private Sub Command6_Click()
    Op_MapaLong = Op_MapaLong + 1
End Sub

Private Sub derecha_Click()
    
    s_mover_cursor CTE_DERECHA

End Sub

Private Sub Form_Activate()
    
    Me.BackColor = cct_ejv(cfondo_ejv)
    ajuste_color_controles_formulario_ejv Me

    'Permito guardar y abrir
    s_cambiar_estado_enabled_operaciones_ficheros_ejv True
    'Pongo inhabilitado todos los programas que son incompatibles con este formulario
    frm_z0_mdi.mn_Editor_Mapas.Enabled = False
    
    'Muestro el mapa actual y inicializo el temporal
    s_refrescar_mapa_actual_ma0
    
    
End Sub
Private Sub Form_GotFocus()
    
    Me.BackColor = cct_ejv(cfondo_ejv)
    ajuste_color_controles_formulario_ejv Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift
    
End Sub

Private Sub Form_Load()
        
    'Refresco automatico, consume muchos recursos
    frm_a0_mapa.AutoRedraw = True
    frm_a0_mapa.WindowState = CTE_MAXIMIZED
        
    frm_a0_mapa.KeyPreview = True 'permito recibir teclas
    s_fijar_separacion_mapa_ma0
    s_cargar_tipo_zoom_ma0

    frm_a0_mapa.Op_MapaFigura.Clear
    frm_a0_mapa.Op_MapaFigura.AddItem "Horizontal"
    frm_a0_mapa.Op_MapaFigura.AddItem "Vertical"
    frm_a0_mapa.Op_MapaFigura.AddItem "Cuadrado"
    frm_a0_mapa.Op_MapaFigura.ListIndex = 0
    
    frm_a0_mapa.Op_MapaTipo.Clear
    frm_a0_mapa.Op_MapaTipo.AddItem "Relleno"
    frm_a0_mapa.Op_MapaTipo.AddItem "Vacío"
    frm_a0_mapa.Op_MapaTipo.ListIndex = 0

    frm_a0_mapa.Op_MapaLong = 5
    
    'Dimensiones del mapa
    habilitar_change_zoom_va0 = False
    frm_a0_mapa.Op_MapaMaxEjeZ = mapa_pisos_ma0
    frm_a0_mapa.Op_MapaMaxEjeY = mapa_filas_ma0
    frm_a0_mapa.Op_MapaMaxEjeX = mapa_columnas_ma0
    habilitar_change_zoom_va0 = True
    
    'Cursor
    cursor_x_va0 = 5
    cursor_y_va0 = 5
    habilitar_change_zoom_va0 = False
    frm_a0_mapa.Op_MapaEjeZ = 1
    frm_a0_mapa.Op_MapaEjeY = 5
    frm_a0_mapa.Op_MapaEjeX = 5
    habilitar_change_zoom_va0 = True

    'Muestro el path-fichero
    frm_a0_mapa.Caption = mapa_actual_ma0
    
    'Inicializo el test
    direccion_test_va0 = 0
    direccion_old_test_va0 = 0
    'Para el test, inicializo los nodos visitados para toda la serie
    s_inicializar_nodos_visitados_va0 mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0


End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim pos_z As Double
    Dim pos_y As Double
    Dim pos_x As Double
    Dim medio As Double
    
    medio = separacion_mapa_ma0 / 2
    
    '14 es es la separacion, el tamaño de los cuadritos
    '7 es la mitad de 14
    'Me muevo a esa posicion
    pos_z = CDbl("0" & Op_MapaEjeZ)
    pos_y = (Int(((Y - CTE_MAPA_INI_Y) - medio) / separacion_mapa_ma0)) + 1
    pos_x = (Int(((X - CTE_MAPA_INI_X) - medio) / separacion_mapa_ma0)) + 1
        
    s_ajustar_lugar_mapa CTE_MAPA_LIMITADO, pos_z, pos_y, pos_x, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
        
    Op_MapaEjeZ = CInt(pos_z)
    Op_MapaEjeY = CInt(pos_y)
    Op_MapaEjeX = CInt(pos_x)
    
    'Invierto ese punto
    'si es 3D pongo tambien en detalle
    If mapa_ma0(Op_MapaEjeZ, Op_MapaEjeY, Op_MapaEjeX) = True Then
        mapa_ma0(Op_MapaEjeZ, Op_MapaEjeY, Op_MapaEjeX) = False
        If ver_zoom_ma0 = CTE_ZOOM_3D Then
            s_refrescar_mapa_actual_ma0
        Else
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, Op_MapaEjeZ, Op_MapaEjeY, Op_MapaEjeX, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        End If
    Else
        mapa_ma0(Op_MapaEjeZ, Op_MapaEjeY, Op_MapaEjeX) = True
        mapa_sin_obstaculos_ma0 = False
        If ver_zoom_ma0 = CTE_ZOOM_3D Then
            s_refrescar_mapa_actual_ma0
        Else
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, Op_MapaEjeZ, Op_MapaEjeY, Op_MapaEjeX, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        End If
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    habilitar_change_zoom_va0 = False
    
    'Ahora no permito guardar y abrir
    s_cambiar_estado_enabled_operaciones_ficheros_ejv False
    
    'Pongo habilitado todos los programas que eran incompatibles con este formulario
     frm_z0_mdi.mn_Editor_Mapas.Enabled = True

End Sub

Private Sub invertir_Click()

    Dim p As Double
    Dim f As Double
    Dim c As Double

    Screen.MousePointer = CTE_ARENA
    frm_a0_mapa.Cls

    For p = 1 To mapa_pisos_ma0
    For f = 1 To mapa_filas_ma0
    For c = 1 To mapa_columnas_ma0
        If mapa_ma0(p, f, c) = True Then
            mapa_ma0(p, f, c) = False
            's_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        Else
            mapa_ma0(p, f, c) = True
            s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, p, f, c, CTE_CUBO, cct_ejv(CTE_DEGRADADOCOLOR), cct_ejv(CTE_DEGRADADOCOLOR), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
        End If
    Next c
    Next f
    Next p

    'Inicializo el test
    direccion_test_va0 = 0
    direccion_old_test_va0 = 0
    'Para el test, inicializo los nodos visitados para toda la serie
    s_inicializar_nodos_visitados_va0 mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0
    
    Screen.MousePointer = CTE_DEFECTO


End Sub

Private Sub izquierda_Click()
    
    s_mover_cursor CTE_IZQUIERDA
    
End Sub


Private Sub Op_MapaEjeX_Change()

    'Control de errores
    If Not IsNumeric(Op_MapaEjeX) Then
        Op_MapaEjeX = 0
        Exit Sub
    End If
    If Op_MapaEjeX < 0 Then
        Op_MapaEjeX = 0
        Exit Sub
    End If
    If Op_MapaEjeX > mapa_columnas_ma0 Then
        Op_MapaEjeX = mapa_columnas_ma0
        Exit Sub
    End If
    
    Op_MapaEjeX = Op_MapaEjeX + 0  'Esto evita el numero -0
    If habilitar_change_zoom_va0 And Ch_Parpadeo.Value = 1 Then
        s_mapa_pintar_cursor_va0
    End If

End Sub
Private Sub Op_MapaEjeY_Change()
    
    'Control de errores
    If Not IsNumeric(Op_MapaEjeY) Then
        Op_MapaEjeY = 0
        Exit Sub
    End If
    If Op_MapaEjeY < 0 Then
        Op_MapaEjeY = 0
        Exit Sub
    End If
    If Op_MapaEjeY > mapa_filas_ma0 Then
        Op_MapaEjeY = mapa_filas_ma0
        Exit Sub
    End If
    
    Op_MapaEjeY = Op_MapaEjeY + 0  'Esto evita el numero -0
    If habilitar_change_zoom_va0 And Ch_Parpadeo.Value = 1 Then
        s_mapa_pintar_cursor_va0
    End If

End Sub

Private Sub Op_MapaEjeZ_Change()
    
    'Control de errores
    If Not IsNumeric(Op_MapaEjeZ) Then
        Op_MapaEjeZ = 0
        Exit Sub
    End If
    If Op_MapaEjeZ < 0 Then
        Op_MapaEjeZ = 0
        Exit Sub
    End If
    If Op_MapaEjeZ > mapa_pisos_ma0 Then
        Op_MapaEjeZ = mapa_pisos_ma0
        Exit Sub
    End If
    
    Op_MapaEjeZ = Op_MapaEjeZ + 0  'Esto evita el numero -0
    If habilitar_change_zoom_va0 And Ch_Parpadeo.Value = 1 Then
        s_mapa_pintar_cursor_va0
    End If

End Sub

Private Sub Rellenar_Click()
    
    Screen.MousePointer = CTE_ARENA
    frm_a0_mapa.Cls
    s_mapa_pintar_figura_va0 "Rellenar Todo", Op_MapaFigura, Op_MapaLong, Op_MapaEjeZ, Op_MapaEjeY, Op_MapaEjeX
    Screen.MousePointer = CTE_DEFECTO

End Sub

Private Sub test_movimiento_Click()
    
    frm_a0_testm.Show
    
End Sub

Private Sub Timer1_Timer()
    s_mapa_parpadeo_cursor_va0
End Sub
Private Sub Vaciar_Click()
    
    Screen.MousePointer = CTE_ARENA
    frm_a0_mapa.Cls
    s_mapa_pintar_figura_va0 "Vaciar Todo", Op_MapaFigura, Op_MapaLong, Op_MapaEjeZ, Op_MapaEjeY, Op_MapaEjeX
    Screen.MousePointer = CTE_DEFECTO

End Sub
