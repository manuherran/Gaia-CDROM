VERSION 5.00
Begin VB.Form frm_u0_font 
   Caption         =   "El Big Bang en Fontainebleau"
   ClientHeight    =   6510
   ClientLeft      =   915
   ClientTop       =   1215
   ClientWidth     =   10035
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "u0_font.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6510
   ScaleWidth      =   10035
   Begin VB.TextBox Op_Resultado 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4200
      Width           =   9615
   End
   Begin VB.Frame Fr_Opciones 
      Caption         =   "Opciones"
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9615
      Begin VB.TextBox Op_CadenaInicial 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5565
         TabIndex        =   14
         Text            =   "1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opciones del movimiento"
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   9375
         Begin VB.TextBox Op_Columnas 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5715
            TabIndex        =   19
            Text            =   "40"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Op_Filas 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5685
            TabIndex        =   17
            Text            =   "40"
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox Cb_Zoom 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox Op_mostrarCamino 
            Caption         =   "Mostrar el camino seguido"
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.Frame Frame3 
            Caption         =   "Movimiento"
            Height          =   855
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   2295
            Begin VB.OptionButton Op_4dir 
               Caption         =   "4 direcciones"
               Height          =   195
               Left            =   240
               TabIndex        =   11
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton Op_8dir 
               Caption         =   "8 direcciones"
               Height          =   315
               Left            =   240
               TabIndex        =   10
               Top             =   480
               Value           =   -1  'True
               Width           =   1815
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Columnas"
            Height          =   195
            Left            =   4785
            TabIndex        =   20
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Filas"
            Height          =   195
            Left            =   5130
            TabIndex        =   18
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.TextBox Op_numIteraciones 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2085
         TabIndex        =   6
         Text            =   "16"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mostrar"
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2295
         Begin VB.OptionButton Op_nMostrarSerie 
            Caption         =   "Mostrar Movimiento"
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Op_MostrarSerie 
            Caption         =   "Mostrar Serie"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cadena Inicial"
         Height          =   195
         Left            =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Número de Iteraciones"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.CommandButton Salir 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Generar 
      Caption         =   "&Generar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "frm_u0_font"
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

Dim Z As Double
Dim Y As Double
Dim X As Double

Dim max_z As Double
Dim max_y As Double
Dim max_x As Double

Dim direccion As Integer

Private Sub Cb_Zoom_Click()
    
    s_click_zoom_ejv frm_u0_font

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
    'frm_u0_font.AutoRedraw = True
    frm_u0_font.AutoRedraw = False
    habilitar_change_zoom_va0 = True
    s_inicializar_combo_zoom_ejv frm_u0_font

End Sub

Private Sub Generar_Click()

    s_azar_fontainebleau
    'If MsgBox("¿Desea hacer visibles las opciones?", vbOKCancel + vbQuestion) = vbYes Then
    '    Fr_Opciones.Visible = True
    'End If
    
End Sub

Private Sub Salir_Click()
    Unload Me

End Sub


Sub s_azar_fontainebleau()

    Dim cadena As String
    Dim i As Integer

    s_fijar_separacion_mapa_va0
    
    cadena = Op_CadenaInicial
    Z = 1
    Y = 20
    X = 20
    
    max_z = 1
    max_y = Op_Filas
    max_x = Op_Columnas
    
    mapa_filas_va0 = max_y
    mapa_columnas_va0 = max_x
    direccion = CTE_8_N
    
    If Op_MostrarSerie Then
        Op_Resultado = ""
    Else
        Me.Fr_Opciones.Visible = False
        Me.Generar.Visible = False
        Me.Salir.Visible = False
        Me.Op_Resultado.Visible = False
        Me.WindowState = CTE_MAXIMIZED
        Me.Cls
        Me.Refresh
        s_mapa_pintar_bordes_va0 frm_u0_font
        s_pintar_objeto_ejv CTE_FORMULARIO, frm_u0_font, Z, Y, X, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(CTE_ROSA), direccion, ver_zoom_va0, 1
    End If
    For i = 1 To Op_numIteraciones
        If Len(cadena) > 5000 Then Exit For
        If Op_MostrarSerie Then
            Op_Resultado = Op_Resultado & cadena & vbCrLf
        Else
            s_mover_hormiga_fontainebleau cadena
        End If
        cadena = f_transforma_cadena_fontainebleau(cadena)
    Next i

End Sub


Function f_transforma_cadena_fontainebleau(cad As String) As String

    Dim ret As String
    Dim temp As String
    Dim c_anterior As String * 1
    Dim c As String * 1
    Dim num_C As Integer
    
    temp = cad

    c = Left(temp, 1)
    c_anterior = c
    temp = Right(temp, Len(temp) - 1)
    num_C = 1
    If Len(temp) = 0 Then
        ret = ret & CStr(num_C) & c
    End If
    While Len(temp) > 0
        c = Left(temp, 1)
        temp = Right(temp, Len(temp) - 1)
        If c = c_anterior Then
            num_C = num_C + 1
        Else
            ret = ret & CStr(num_C) & c_anterior
            c_anterior = c
            num_C = 1
        End If
    If Len(temp) = 0 Then
        ret = ret & CStr(num_C) & c
    End If
    Wend

    f_transforma_cadena_fontainebleau = ret

End Function


Sub s_mover_hormiga_fontainebleau(cadena As String)
    
    Dim temp As String
    Dim c As String * 1

    temp = cadena
    While Len(temp) > 0
        c = Left(temp, 1)
        temp = Right(temp, Len(temp) - 1)
        Select Case CInt(c)
            Case 1 'Avanzar
                If Op_mostrarCamino.Value = 0 Then
                    s_pintar_objeto_ejv CTE_FORMULARIO, frm_u0_font, Z, Y, X, CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                End If
                f_avanzo_8_dir_va0 Z, Y, X, direccion, max_z, max_y, max_x
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_u0_font, Z, Y, X, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(CTE_ROSA), direccion, ver_zoom_va0, 1
            Case 2 'Giro Izquierda
                If Op_4dir Then
                    direccion = f_giro_8_general_va0(direccion, CTE_8_IZQ)
                Else
                    direccion = f_giro_8_general_va0(direccion, CTE_8_DEF_IZQ)
                End If
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_u0_font, Z, Y, X, CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_u0_font, Z, Y, X, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(CTE_ROSA), direccion, ver_zoom_va0, 1
            Case 3 'Giro Derecha
                If Op_4dir Then
                    direccion = f_giro_8_general_va0(direccion, CTE_8_DER)
                Else
                    direccion = f_giro_8_general_va0(direccion, CTE_8_DEF_DER)
                End If
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_u0_font, Z, Y, X, CTE_ESFERA, cct_ejv(cfondo_ejv), cct_ejv(cfondo_ejv), CTE_DIRECC_NINGUNA, ver_zoom_va0, 1
                s_pintar_objeto_ejv CTE_FORMULARIO, frm_u0_font, Z, Y, X, CTE_HORMIGA, cct_ejv(CTE_NEGRO), cct_ejv(CTE_ROSA), direccion, ver_zoom_va0, 1
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
        End Select
        DoEvents
    Wend

End Sub
