VERSION 5.00
Begin VB.MDIForm frm_z0_mdi 
   BackColor       =   &H8000000C&
   Caption         =   "Vida"
   ClientHeight    =   4020
   ClientLeft      =   2475
   ClientTop       =   2565
   ClientWidth     =   7530
   Icon            =   "Z0_MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox BarraHerramientas 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7530
      TabIndex        =   3
      Top             =   0
      Width           =   7530
      Begin VB.ComboBox Cb_Zoom 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton H_Estado 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4920
         Picture         =   "Z0_MDI.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Estado de la Ejecución"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Refrescar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5280
         Picture         =   "Z0_MDI.frx":083C
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Refrescar Mundo"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Opciones 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4080
         Picture         =   "Z0_MDI.frx":0D6E
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ver Opciones"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_HYP 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   840
         Picture         =   "Z0_MDI.frx":12A0
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Hormigas y Plantas"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_3R 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1560
         Picture         =   "Z0_MDI.frx":17D2
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Tres en Raya"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_PAL 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1200
         Picture         =   "Z0_MDI.frx":1D04
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Palabras y Frases"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_PRI 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1920
         Picture         =   "Z0_MDI.frx":2236
         Style           =   1  'Graphical
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "El Dilema del Prisionero"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Ayuda 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7080
         Picture         =   "Z0_MDI.frx":2768
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ayuda"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Abrir 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Picture         =   "Z0_MDI.frx":2C9A
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Abrir Fichero"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Mapa 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2400
         Picture         =   "Z0_MDI.frx":31CC
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ver Mapa"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Grafico 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5640
         Picture         =   "Z0_MDI.frx":36FE
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ver Gráfico"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Terminar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3600
         Picture         =   "Z0_MDI.frx":3C30
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Stop"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Pausa 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3240
         Picture         =   "Z0_MDI.frx":4162
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Pausa"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton H_Agentes 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4440
         Picture         =   "Z0_MDI.frx":4694
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ver Tipos de Agentes"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Guardar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   360
         Picture         =   "Z0_MDI.frx":4BC6
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Guardar Fichero"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton H_Comenzar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2880
         Picture         =   "Z0_MDI.frx":50F8
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Comenzar"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Menu Mn_Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu mn_Abrir 
         Caption         =   "Abrir...   F2"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Guardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mn_GuardarComo 
         Caption         =   "Guardar Como..."
      End
      Begin VB.Menu linea11 
         Caption         =   "-"
      End
      Begin VB.Menu mn_generador_aut 
         Caption         =   "Generador de *.aut..."
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
      End
      Begin VB.Menu Mn_Salir 
         Caption         =   "Salir   Alt+F4"
      End
   End
   Begin VB.Menu mn_Edicion 
      Caption         =   "Edición"
      Begin VB.Menu mn_cortar 
         Caption         =   "Cortar   Ctrl+X"
      End
      Begin VB.Menu mn_copiar 
         Caption         =   "Copiar   Ctrl+C"
      End
      Begin VB.Menu mn_pegar 
         Caption         =   "Pegar   Ctrl+V"
      End
   End
   Begin VB.Menu mn_Ejemplos 
      Caption         =   "Ejemplos de Vida"
      Begin VB.Menu Mn_hyp 
         Caption         =   "Hormigas y Plantas..."
      End
      Begin VB.Menu mn_palyfras 
         Caption         =   "Palabras y Frases (Algoritmo Genético)..."
      End
      Begin VB.Menu mn_3r 
         Caption         =   "Tres en Raya (Clasificador Genético)..."
      End
      Begin VB.Menu mn_Prisionero 
         Caption         =   "El Juego del Prisionero..."
      End
      Begin VB.Menu mn_Explorando_Mapas 
         Caption         =   "Explorando Mapas..."
      End
      Begin VB.Menu mn_Universo 
         Caption         =   "Universo..."
      End
      Begin VB.Menu linea10 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Editor_Mapas 
         Caption         =   "Editor de Mapas..."
      End
      Begin VB.Menu linea5 
         Caption         =   "-"
      End
      Begin VB.Menu mn_obras 
         Caption         =   "En Obras"
         Begin VB.Menu mn_gaia 
            Caption         =   "Plataforma Gaia..."
         End
         Begin VB.Menu mn_Cadenas 
            Caption         =   "Cadenas (Algoritmo Genético)..."
         End
         Begin VB.Menu mn_Celdilla 
            Caption         =   "Celdilla..."
         End
         Begin VB.Menu mn_Peces 
            Caption         =   "Peces..."
         End
         Begin VB.Menu mn_yxy 
            Caption         =   "yxy..."
         End
         Begin VB.Menu linea13 
            Caption         =   "-"
         End
         Begin VB.Menu mn_EjecutarEjemplosDeVida 
            Caption         =   "Ejecutar otra instancia de Ejemplos de Vida"
         End
         Begin VB.Menu mn_ExplorarCarpetas 
            Caption         =   "Explorar Carpetas..."
         End
         Begin VB.Menu linea14 
            Caption         =   "-"
         End
         Begin VB.Menu mn_gaia_xls 
            Caption         =   "Gaia.xls"
         End
      End
      Begin VB.Menu linea15 
         Caption         =   "-"
      End
      Begin VB.Menu mn_listaviejos 
         Caption         =   "Hormigas y Plantas - Ej 1"
         Index           =   0
      End
   End
   Begin VB.Menu mn_Ejecutar 
      Caption         =   "Ejecutar"
      Begin VB.Menu mn_Comenzar 
         Caption         =   "&Comenzar   F5"
      End
      Begin VB.Menu mn_Continuar 
         Caption         =   "&Continuar   F5"
      End
      Begin VB.Menu linea3 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Pausa 
         Caption         =   "&Pausa   F6"
      End
      Begin VB.Menu mn_terminar 
         Caption         =   "&Terminar   F7"
      End
      Begin VB.Menu linea7 
         Caption         =   "-"
      End
      Begin VB.Menu mn_JugarContraOrdenador 
         Caption         =   "Jugar Contra el Ordenador..."
      End
   End
   Begin VB.Menu mn_Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu mn_Opciones1 
         Caption         =   "Opciones I ..."
      End
      Begin VB.Menu mn_Opciones2 
         Caption         =   "Opciones II ..."
      End
      Begin VB.Menu mn_Opciones3 
         Caption         =   "Opciones III ..."
      End
      Begin VB.Menu mn_Tipos_Agentes 
         Caption         =   "Opciones IV: Tipos de Agentes..."
      End
      Begin VB.Menu mn_Mapa 
         Caption         =   "Opciones V: Mapa..."
      End
      Begin VB.Menu mn_TipoEvolucion 
         Caption         =   "Opciones VI: Tipo de Evolución"
         Begin VB.Menu mn_Metodo_Evaluacion 
            Caption         =   "Método de Evaluación..."
         End
         Begin VB.Menu mn_Metodo_Seleccion 
            Caption         =   "Método de Selección..."
         End
         Begin VB.Menu mn_Metodo_Reproduccion 
            Caption         =   "Método de Reproducción"
            Begin VB.Menu mn_Tipo_Sobrecruzamiento 
               Caption         =   "Tipo de Sobrecruzamiento..."
            End
            Begin VB.Menu mn_Tipo_Mutaciones 
               Caption         =   "Tipo de Mutaciones..."
            End
         End
      End
   End
   Begin VB.Menu mn_ver 
      Caption         =   "Ver"
      Begin VB.Menu mn_EstadoEjecucion 
         Caption         =   "Estado de la Ejecución"
      End
      Begin VB.Menu mn_Refrescar 
         Caption         =   "Refrescar Mundo"
      End
      Begin VB.Menu mn_ListaAgentes 
         Caption         =   "Lista de Agentes (Todos)"
      End
      Begin VB.Menu mn_MejoresAgentes 
         Caption         =   "Mejores Agentes"
      End
      Begin VB.Menu mn_Diccionario 
         Caption         =   "Diccionario"
      End
      Begin VB.Menu mn_Apellidos 
         Caption         =   "Apellidos"
      End
      Begin VB.Menu mn_ModificarAgente 
         Caption         =   "Modificar Agente..."
      End
      Begin VB.Menu linea9 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Grafico 
         Caption         =   "Gráfico...   F8"
      End
   End
   Begin VB.Menu mn_Ventana 
      Caption         =   "Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mn_Mosaico_Horizontal 
         Caption         =   "Mosaico Horizontal"
      End
      Begin VB.Menu mn_MosaicoVertical 
         Caption         =   "Mosaico Vertical"
      End
      Begin VB.Menu mn_Cascada 
         Caption         =   "Cascada"
      End
      Begin VB.Menu mn_Organizar_Iconos 
         Caption         =   "Organizar Iconos"
      End
   End
   Begin VB.Menu Mn_Ayuda 
      Caption         =   "?"
      Begin VB.Menu Mn_AyudaHtm 
         Caption         =   "Documentación en htm..."
      End
      Begin VB.Menu Mn_Notas 
         Caption         =   "Notas sobre la versión"
         Begin VB.Menu Mn_AyudaDoc 
            Caption         =   "en Word (.doc)..."
         End
         Begin VB.Menu Mn_AyudaTxt 
            Caption         =   "en texto (.txt)..."
         End
      End
      Begin VB.Menu Mn_Readme 
         Caption         =   "Readme.txt para el programador"
      End
      Begin VB.Menu linea12 
         Caption         =   "-"
      End
      Begin VB.Menu Mn_Calculadora 
         Caption         =   "Calculadora..."
      End
      Begin VB.Menu linea6 
         Caption         =   "-"
      End
      Begin VB.Menu Mn_AcercaDe 
         Caption         =   "Acerca de..."
      End
   End
   Begin VB.Menu mn_terminar_todo 
      Caption         =   "Terminar Todo"
   End
End
Attribute VB_Name = "frm_z0_mdi"
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

Private Sub BarraHerramientas_KeyDown(KeyCode As Integer, Shift As Integer)
    
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub
Private Sub Cb_Zoom_Click()
    
    s_click_zoom_ejv frm_z0_mdi

End Sub

Private Sub Cb_Zoom_KeyDown(KeyCode As Integer, Shift As Integer)
    
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_3R_Click()
    s_click_programa_ejv CTE_3R

End Sub

Private Sub H_3R_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Abrir_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Agentes_Click()
    s_operacion_ver_ejv CTE_VER_TIPOS_AGENTES

End Sub

Private Sub H_Agentes_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Comenzar_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Estado_Click()
    s_operacion_ver_ejv CTE_VER_ESTADO_EJECUCION

End Sub

Private Sub H_Estado_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Grafico_Click()
    s_operacion_ver_ejv CTE_VER_GRAFICO

End Sub

Private Sub H_Grafico_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Guardar_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_HYP_Click()
    s_click_programa_ejv CTE_HYP

End Sub

Private Sub H_HYP_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Mapa_Click()
    If frm_z0_mdi.mn_Mapa.Enabled Then
        s_operacion_ver_ejv CTE_VER_MAPA
    Else
        copiar_mapa_a_va0_ma0 = False
        'Cargo el mapa por defecto
        mapa_actual_ma0 = f_nombre_completo(path_largo_ejv(CTE_C_PRG_MAP), "default.map")
        s_aut_leer_mapa_ma0
        frm_a0_mapa.Show CTE_AMODAL
        frm_a0_mapa.Caption = "Editor de Mapas"
    End If

End Sub


Private Sub H_Mapa_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Opciones_Click()
    s_operacion_ver_ejv CTE_VER_OPCIONES1

End Sub

Private Sub H_Opciones_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_PAL_Click()
    s_click_programa_ejv CTE_PAL

End Sub

Private Sub H_PAL_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Pausa_Click()
    s_operacion_ejecutar_ejv CTE_EXE_PAUSA

End Sub

Private Sub H_Pausa_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_PRI_Click()
    s_click_programa_ejv CTE_PRI

End Sub

Private Sub H_PRI_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Refrescar_Click()
    's_operacion_ver_ejv CTE_VER_ESTADO_EJECUCION
    s_operacion_ver_ejv CTE_VER_REFRESCAR

End Sub

Private Sub H_Refrescar_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub H_Terminar_Click()
    
    If automatico_ejv Then
        If MsgBox("Si pulsa ""Terminar"" se interrumpirá el ejemplo que actualmente se está ejecutando, pero después el programa continuará con el resto de los ejemplos definidos en los ficheros .aut referenciados en inicio.txt. Si lo que desea es finalizar todos los ejemplos, pulse el menú ""Terminar Todo"". ¿Desea interrumpir el ejemplo actual?", vbQuestion + vbYesNo + vbDefaultButton2) = vbOK Then
            finalizacion_usuario_ejv = True
            s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
        End If
    Else
        finalizacion_usuario_ejv = True
        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
    End If
    

End Sub


Private Sub H_Terminar_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub MDIForm_Load()

    BarraHerramientas.AutoSize = True
    mostrar_aviso_imagen_ejv = True
    s_mdi_load_ejv
    Screen.MousePointer = CTE_DEFECTO
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = f_peticion_unload_mdi_ejv

End Sub
Private Sub mn_3r_Click()
    
    s_click_programa_ejv CTE_3R

End Sub

Private Sub mn_Abrir_Click()

    s_accion_ficheros_va0 CTE_FIC_ABRIR

End Sub

Private Sub Mn_AcercaDe_Click()

    frm_z0_mdi.Caption = "Acerca de... " & nombre_aplicacion_ejv
    frm_z0_acer.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    s_fijar_caption_mdi

End Sub

Private Sub mn_Apellidos_Click()
    s_operacion_ver_ejv CTE_VER_APELLIDOS

End Sub

Private Sub Mn_AyudaDoc_Click()
    
    On Error Resume Next
    Dim Lsalida As Long
    Lsalida = Shell("start " & f_nombre_completo(path_largo_ejv(CTE_C_DOC), "version.doc"))

End Sub

Private Sub Mn_AyudaHtm_Click()

    s_mostrar_docum_html_ejv

End Sub

Private Sub Mn_AyudaSobre_Click()

     s_mostrar_docum_html_ejv

End Sub

Private Sub Mn_AyudaTxt_Click()
    
    On Error Resume Next
    Dim RetVal
    RetVal = Shell("notepad.exe " & f_nombre_completo(path_largo_ejv(CTE_C_DOC), "version.txt"), 1)

End Sub

Private Sub mn_Cadenas_Click()

    s_click_programa_ejv CTE_CAD

End Sub

Private Sub Mn_Calculadora_Click()

    On Error Resume Next
    Dim RetVal
    RetVal = Shell("calc.exe", 1)

End Sub

Private Sub mn_Cascada_Click()

     frm_z0_mdi.Arrange CTE_CASCADA

End Sub

Private Sub mn_Celdilla_Click()
    
    s_click_programa_ejv CTE_CEL

End Sub

Private Sub mn_Comenzar_Click()
    
    s_operacion_ejecutar_ejv CTE_EXE_COMENZAR

End Sub
Private Sub mn_Continuar_Click()
    
    s_operacion_ejecutar_ejv CTE_EXE_CONTINUAR

End Sub

Private Sub mn_copiar_Click()

    SendKeys "^(C)"
    'SHIFT   +
    'CTRL    ^
    'ALT %
    'Dim ReturnValue, I
    'ReturnValue = Shell("CALC.EXE", 1)  ' Run Calculator.
    'AppActivate ReturnValue     ' Activate the Calculator.
    'For I = 1 To 100    ' Set up counting loop.
    '    SendKeys I & "{+}", True    ' Send keystrokes to Calculator
    'Next I  ' to add each value of I.
    'SendKeys "=", True  ' Get grand total.
    'SendKeys "%{F4}", True  ' Send ALT+F4 to close Calculator.

End Sub

Private Sub mn_cortar_Click()
    
    SendKeys "^(X)"

End Sub

Private Sub mn_Diccionario_Click()
    
    s_operacion_ver_ejv CTE_VER_DICCIONARIO

End Sub


Private Sub mn_Editor_Mapas_Click()

    copiar_mapa_a_va0_ma0 = False
    'Cargo el mapa por defecto
    mapa_actual_ma0 = f_nombre_completo(path_largo_ejv(CTE_C_PRG_MAP), "default.map")
    s_aut_leer_mapa_ma0
    frm_a0_mapa.Show CTE_AMODAL
    frm_a0_mapa.Caption = "Editor de Mapas"
    
    
End Sub


Private Sub mn_EstadoEjecucion_Click()
    
    s_operacion_ver_ejv CTE_VER_ESTADO_EJECUCION

End Sub

Private Sub mn_Explorando_Mapas_Click()
    
    s_click_programa_ejv CTE_EXP

End Sub

Private Sub mn_ExplorarCarpetas_Click()
    
    On Error Resume Next
    Dim RetVal
    RetVal = Shell("Explorer " & path_largo_ejv(CTE_C_RAIZ), 1)

End Sub

Private Sub mn_Gaia_Click()
    
    s_click_programa_ejv CTE_GAI

End Sub


Private Sub mn_gaia_xls_Click()
    
    On Error Resume Next
    Dim Lsalida As Long
    Lsalida = Shell("start " & f_nombre_completo(path_largo_ejv(CTE_C_PRG), "gaia.xls"))

End Sub

Private Sub mn_generador_aut_Click()
     
     frm_z0_aut.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show

End Sub

Private Sub mn_Grafico_Click()
    s_operacion_ver_ejv CTE_VER_GRAFICO

End Sub

Private Sub mn_Guardar_Click()
    
    s_accion_ficheros_va0 CTE_FIC_GUARDAR
    
End Sub

Private Sub mn_GuardarComo_Click()
    
    s_accion_ficheros_va0 CTE_FIC_GUARDARCOMO

End Sub

Private Sub Mn_hyp_Click()
    
    s_click_programa_ejv CTE_HYP

End Sub


Private Sub mn_JugarContraOrdenador_Click()
    s_operacion_ver_ejv CTE_VER_JUGAR_CONTRA_ORDENADOR

End Sub

Private Sub mn_ListaAgentes_Click()
    s_operacion_ver_ejv CTE_VER_AGENTES_TODOS

End Sub

Private Sub mn_listaviejos_Click(Index As Integer)

    Dim txt As String
    Dim com As Integer
    Dim ej As Integer
    

    'Ejecuto ese programa
    txt = frm_z0_mdi.mn_listaviejos(Index).Caption
    com = InStr(txt, " - Ej ")
    ej = CInt(Right(txt, (Len(txt) - com - 5)))
    txt = Left(txt, com - 1)
    Select Case txt
        Case nombre_programa_ejv(CTE_HYP) '1
            num_prg_activo_ejv = CTE_HYP
        Case nombre_programa_ejv(CTE_PAL) '2
            num_prg_activo_ejv = CTE_PAL
        Case nombre_programa_ejv(CTE_3R) '3
            num_prg_activo_ejv = CTE_3R
        Case nombre_programa_ejv(CTE_PRI) '4
            num_prg_activo_ejv = CTE_PRI
        Case nombre_programa_ejv(CTE_CEL) '5
            num_prg_activo_ejv = CTE_CEL
        Case nombre_programa_ejv(CTE_GAI) '6
            num_prg_activo_ejv = CTE_GAI
        Case nombre_programa_ejv(CTE_EXP) '7
            num_prg_activo_ejv = CTE_EXP
        Case nombre_programa_ejv(CTE_CAD) '8
            num_prg_activo_ejv = CTE_CAD
        Case nombre_programa_ejv(CTE_PEZ) '9
            num_prg_activo_ejv = CTE_PEZ
        Case nombre_programa_ejv(CTE_UVA) '10
            num_prg_activo_ejv = CTE_UVA
        Case nombre_programa_ejv(CTE_YXY) '11
            num_prg_activo_ejv = CTE_YXY
       
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select
    
    s_aceptar_menu_ejv CStr("Ejemplo " & ej), 1

End Sub

Private Sub mn_Mapa_Click()
    s_operacion_ver_ejv CTE_VER_MAPA

End Sub

Private Sub mn_MejoresAgentes_Click()
    s_operacion_ver_ejv CTE_VER_AGENTES_MEJORES

End Sub

Private Sub mn_Metodo_Evaluacion_Click()
    s_operacion_ver_ejv CTE_VER_TIPO_EVOLUCION_EVALUACION

End Sub

Private Sub mn_Metodo_Seleccion_Click()
    s_operacion_ver_ejv CTE_VER_TIPO_EVOLUCION_SELECCION

End Sub


Private Sub mn_ModificarAgente_Click()

    s_operacion_ver_ejv CTE_VER_MODIFICAR_AGENTE

End Sub

Private Sub mn_Mosaico_Horizontal_Click()
     
     frm_z0_mdi.Arrange CTE_MOSAICO_HORIZONTAL

End Sub
Private Sub mn_MosaicoVertical_Click()

     frm_z0_mdi.Arrange CTE_MOSAICO_VERTICAL

End Sub

Private Sub mn_Opciones1_Click()
    s_operacion_ver_ejv CTE_VER_OPCIONES1

End Sub

Private Sub mn_Opciones2_Click()
    s_operacion_ver_ejv CTE_VER_OPCIONES2

End Sub

Private Sub mn_Opciones3_Click()
    s_operacion_ver_ejv CTE_VER_OPCIONES3

End Sub

Private Sub mn_Organizar_Iconos_Click()

     frm_z0_mdi.Arrange CTE_ORGANIZAR_ICONOS

End Sub

Private Sub mn_padreOpciones_Click()

End Sub

Private Sub mn_palyfras_Click()

    s_click_programa_ejv CTE_PAL

End Sub

Private Sub mn_Pausa_Click()
    
    s_operacion_ejecutar_ejv CTE_EXE_PAUSA

End Sub

Private Sub mn_Peces_Click()
    
    s_click_programa_ejv CTE_PEZ

End Sub

Private Sub mn_pegar_Click()
    
    SendKeys "^(V)"

End Sub

Private Sub mn_Prisionero_Click()
     
    s_click_programa_ejv CTE_PRI

End Sub

Private Sub Mn_Readme_Click()
    
    On Error Resume Next
    Dim RetVal
    RetVal = Shell("notepad.exe " & f_nombre_completo(path_largo_ejv(CTE_C_DOC), "readme.txt"), 1)

End Sub

Private Sub mn_Refrescar_Click()
    
    's_operacion_ver_ejv CTE_VER_ESTADO_EJECUCION
    s_operacion_ver_ejv CTE_VER_REFRESCAR

End Sub

Private Sub Mn_Salir_Click()
    Unload Me
End Sub

Private Sub mn_Terminar_Click()
    
    If automatico_ejv Then
        If MsgBox("Si pulsa ""Terminar"" se interrumpirá el ejemplo que actualmente se está ejecutando, pero después el programa continuará con el resto de los ejemplos definidos en los ficheros .aut referenciados en inicio.txt. Si lo que desea es finalizar todos los ejemplos, pulse el menú ""Terminar Todo"". ¿Desea interrumpir el ejemplo actual?", vbQuestion + vbYesNo + vbDefaultButton2) = vbOK Then
            finalizacion_usuario_ejv = True
            s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
        End If
    Else
        finalizacion_usuario_ejv = True
        s_operacion_ejecutar_ejv CTE_EXE_TERMINAR
    End If

End Sub


Private Sub mn_terminar_todo_Click()
    
    If MsgBox("¿Desea terminar completamente la ejecución?", vbQuestion + vbYesNo) = vbOK Then
        terminar_todo_ejv = True
    End If
    
End Sub

Private Sub mn_Tipo_Mutaciones_Click()

    s_operacion_ver_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES

End Sub

Private Sub mn_Tipo_Sobrecruzamiento_Click()
    
    s_operacion_ver_ejv CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO

End Sub

Private Sub mn_Tipos_Agentes_Click()
    s_operacion_ver_ejv CTE_VER_TIPOS_AGENTES

End Sub


Private Sub H_Abrir_Click()
    
    s_accion_ficheros_va0 CTE_FIC_ABRIR

End Sub

Private Sub H_Comenzar_Click()
    
    If num_prg_activo_ejv <> CTE_NINGUNO Then
        If estado_ejecutar_ejv(CTE_EXE_COMENZAR, num_prg_activo_ejv) Then
            s_operacion_ejecutar_ejv CTE_EXE_COMENZAR
        ElseIf estado_ejecutar_ejv(CTE_EXE_CONTINUAR, num_prg_activo_ejv) Then
            s_operacion_ejecutar_ejv CTE_EXE_CONTINUAR
        End If
    End If

End Sub

Private Sub H_Guardar_Click()
    s_accion_ficheros_va0 CTE_FIC_GUARDAR

End Sub

Private Sub mn_Universo_Click()
    s_click_programa_ejv CTE_UVA

End Sub

Private Sub mn_yxy_Click()
    s_click_programa_ejv CTE_YXY

End Sub
