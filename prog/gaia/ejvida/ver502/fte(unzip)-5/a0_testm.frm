VERSION 5.00
Begin VB.Form frm_a0_testm 
   Caption         =   "Test Movimiento"
   ClientHeight    =   5580
   ClientLeft      =   165
   ClientTop       =   3195
   ClientWidth     =   3795
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "a0_testm.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5580
   ScaleWidth      =   3795
   Begin VB.CommandButton Prueba3D7 
      Caption         =   "Prueba3D7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Prueba3D6 
      Caption         =   "Prueba3D6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Prueba3D5 
      Caption         =   "Prueba3D5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Prueba3D4 
      Caption         =   "Prueba3D4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Prueba3D3 
      Caption         =   "Prueba3D3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Prueba3D2 
      Caption         =   "Prueba3D2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Prueba3D 
      Caption         =   "Prueba3D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test de Crecimiento Fractal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3495
      Begin VB.ComboBox Cb_Metodo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox Op_CreandoObstaculosF 
         Caption         =   "Creando Obstaculos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton test_crecimiento 
         Caption         =   "Crecimiento Fractal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox num_ciclos 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Método"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº ciclos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Test de Búsqueda con Obstáculos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox num_mov 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "10"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton test_busqueda 
         Caption         =   "Probar funcion de búsqueda en espiral con obstáculos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox Op_CreandoObstaculosM 
         Caption         =   "Creando Obstaculos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Cb_Algoritmo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº mov."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Algoritmo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_a0_testm"
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

Private Sub Cb_Algoritmo_Change()
    'inicializo la direccion
    direccion_test_va0 = CTE_NORTE
    direccion_old_test_va0 = CTE_NORTE

End Sub



Private Sub Form_Load()

    'Cargo las opciones
    'Test de Búsqueda con obstáculos
    frm_a0_testm.Cb_Algoritmo.Clear
    frm_a0_testm.Cb_Algoritmo.AddItem "Algoritmo 1"
    frm_a0_testm.Cb_Algoritmo.AddItem "Algoritmo 2"
    frm_a0_testm.Cb_Algoritmo.AddItem "Algoritmo 3"
    If algoritmo_busqueda_va0 <> 0 Then
        frm_a0_testm.Cb_Algoritmo.ListIndex = algoritmo_busqueda_va0 - 1 'el primero es el cero
    Else
        frm_a0_testm.Cb_Algoritmo.ListIndex = 0
    End If
    frm_a0_testm.Cb_Algoritmo.Refresh
    
    
    'Test de Crecimiento Fractal
    frm_a0_testm.Cb_Metodo.Clear
    frm_a0_testm.Cb_Metodo.AddItem "Método 1"
    frm_a0_testm.Cb_Metodo.AddItem "Método 2"
    frm_a0_testm.Cb_Metodo.AddItem "Método 3"
    frm_a0_testm.Cb_Metodo.AddItem "Método 4"
    frm_a0_testm.Cb_Metodo.ListIndex = 0
    frm_a0_testm.Cb_Algoritmo.Refresh
    

End Sub

Private Sub Prueba3D_Click()

    Screen.MousePointer = CTE_ARENA
    frm_a0_mapa.Cls
    s_pintar_ejes3D CTE_FORMULARIO, frm_a0_mapa, 2
    s_genera_mundo_esferas
    Screen.MousePointer = CTE_DEFECTO
    

End Sub

Private Sub Prueba3D2_Click()
    
    Screen.MousePointer = CTE_ARENA
    frm_a0_mapa.Cls
    s_pintar_ejes3D CTE_FORMULARIO, frm_a0_mapa, 2
    s_genera_cubos
    Screen.MousePointer = CTE_DEFECTO

End Sub

Private Sub Prueba3D3_Click()
    
    Screen.MousePointer = CTE_ARENA
    frm_a0_mapa.Cls
    s_pintar_ejes3D CTE_FORMULARIO, frm_a0_mapa, 2
    s_genera_cubos3_1
    Screen.MousePointer = CTE_DEFECTO

End Sub

Private Sub Prueba3D4_Click()
    
    Screen.MousePointer = CTE_ARENA
    s_pintar_ejes3D CTE_FORMULARIO, frm_a0_mapa, 2
    s_genera_plano_cubo_esfera False, 4, 6, 6
    Screen.MousePointer = CTE_DEFECTO

End Sub

Private Sub Prueba3D5_Click()
    Screen.MousePointer = CTE_ARENA
    s_pintar_ejes3D CTE_FORMULARIO, frm_a0_mapa, 2
    s_genera_plano_cubo_esfera True, 6, 6, 6
    Screen.MousePointer = CTE_DEFECTO

End Sub

Private Sub Prueba3D6_Click()
    
    Screen.MousePointer = CTE_ARENA
    s_pintar_ejes3D CTE_FORMULARIO, frm_a0_mapa, 2
    s_genera_cosas
    Screen.MousePointer = CTE_DEFECTO

End Sub

Private Sub Prueba3D7_Click()

    s_genera_cubos2
End Sub

Private Sub test_busqueda_Click()
    Dim numero_elementos_pintar As Long
    
    Dim i As Long
    Dim estoy_rodeado As Integer
    
    Screen.MousePointer = CTE_ARENA
    'Pasamos a modo test
    estado_test_movimiento_va0 = True
    
    'Establezco la direccion inicial si es el primer clik
    If direccion_test_va0 = 0 Then
        direccion_test_va0 = CTE_NORTE
    End If

    'Guardo el numero de direcciones del algoritmo viejo antes de cambiarlo
    If num_direcc_algoritmo_va0 <> 0 Then
        num_direcc_old_algoritmo_va0 = num_direcc_algoritmo_va0
    End If
    Select Case Cb_Algoritmo.ListIndex
        Case 0
            num_direcc_algoritmo_va0 = 4
        Case 1
            num_direcc_algoritmo_va0 = 4
        Case 2
            num_direcc_algoritmo_va0 = 8
        Case Else
            MsgBox "Error: algoritmo inexistente", vbCritical
    End Select

    'Establezco el nodo inicial
    cursor_y_va0 = CInt(frm_a0_mapa.Op_MapaEjeY)
    cursor_x_va0 = CInt(frm_a0_mapa.Op_MapaEjeX)
    
    numero_elementos_pintar = CInt(frm_a0_testm.num_mov)

    'El nodo actual no lo tengo en cuenta
    'como nodo a visitar, y empiezo buscando uno
    'que no sea el actual
    For i = 1 To numero_elementos_pintar
        'Guardo la vieja dirección y posicion porque al moverme se va a cambiar
        direccion_old_test_va0 = direccion_test_va0
        cursor_old_y_va0 = cursor_y_va0
        cursor_old_x_va0 = cursor_x_va0
        num_direcc_old_algoritmo_va0 = num_direcc_algoritmo_va0
        'Aqui analizo el mapa temporal en vez del real
        'en este caso no hay hormigas no otras cosas mas que obstaculos
        'elijo nueva posicion
        Select Case Cb_Algoritmo.ListIndex
            Case 0
                'Algoritmo 1
                estoy_rodeado = f_alg1_calcular_siguiente_nodo_va0(cursor_z_va0, cursor_y_va0, cursor_x_va0, direccion_test_va0, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0)
                num_direcc_algoritmo_va0 = 4
            Case 1
                'Algoritmo 2
                estoy_rodeado = f_alg2_calcular_siguiente_nodo_va0(cursor_z_va0, cursor_y_va0, cursor_x_va0, direccion_test_va0, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0)
                num_direcc_algoritmo_va0 = 4
            Case 2
                'Algoritmo 3
                estoy_rodeado = f_alg3_calcular_siguiente_nodo_va0(cursor_z_va0, cursor_y_va0, cursor_x_va0, direccion_test_va0, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0)
                num_direcc_algoritmo_va0 = 8
            Case Else
                MsgBox "Error: algoritmo inexistente", vbCritical
                Exit Sub
        End Select
        If estoy_rodeado Then
            'Salgo de los bucles, no hay nada que hacer
            Exit For
        Else
            'control errores de programacion
            If control_errores_de_programacion_ejv Then
                If Not f_esta_vacio_va0(1, cursor_y_va0, cursor_x_va0) Then
                    s_error_ejv CON_OPCION_FINALIZAR, "Error"
                End If
            End If
            'Lo pongo como obstaculo si estoy en ese caso
            If frm_a0_testm.Op_CreandoObstaculosM = 1 Then
                mapa_ma0(1, cursor_y_va0, cursor_x_va0) = True
            End If
            'Me muevo a esa posicion, pero antes deshabilito el evento change
            'de estas cajas de texto, ya que lo que quiero hacer es cambiar
            'las dos a la vez, cosa que evidentemente no se puede
            habilitar_change_zoom_va0 = False
            frm_a0_mapa.Op_MapaEjeY.Text = cursor_y_va0
            frm_a0_mapa.Op_MapaEjeX.Text = cursor_x_va0
            habilitar_change_zoom_va0 = True
            'Borro el viejo (borro el cursor con su pirindolo)
            s_mapa_cursor_pintado_va0 cursor_old_z_va0, cursor_old_y_va0, cursor_old_x_va0, direccion_old_test_va0, num_direcc_old_algoritmo_va0, cct_ejv(cfondo_ejv)
            'pinto el nuevo: no hace falta el cursor porque se pone solo
            If frm_a0_testm.Op_CreandoObstaculosM = 1 Then
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, cursor_z_va0, cursor_y_va0, cursor_x_va0, CTE_CUBO, cct_ejv(CTE_NEGRO), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            Else
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_a0_mapa, cursor_z_va0, cursor_y_va0, cursor_x_va0, CTE_CUBO, cct_ejv(CTE_ROSA), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_ma0, 1
            End If
        End If
    Next i

    'Salimos de modo test
    estado_test_movimiento_va0 = False
    Screen.MousePointer = CTE_DEFECTO

End Sub

Private Sub test_crecimiento_Click()
    
    Screen.MousePointer = CTE_ARENA
    
    If frm_a0_testm.Cb_Metodo.ListIndex = 3 Then
        'Método 4: hoja
        s_insertar_hoja_ma0
    Else
        'Método 1
        'Método 2
        'Método 3
        s_test_crecimiento_ma0
    End If
    Screen.MousePointer = CTE_DEFECTO

End Sub


