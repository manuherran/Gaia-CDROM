VERSION 5.00
Begin VB.Form frm_c0_ce 
   Caption         =   "Tres en Raya"
   ClientHeight    =   7095
   ClientLeft      =   810
   ClientTop       =   405
   ClientWidth     =   10395
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "c0_ce.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   10395
   Begin VB.Frame Fr_ModificarAgente 
      Caption         =   "Modificar Agente"
      Height          =   4815
      Left            =   6600
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox superage 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   21
         Text            =   "1"
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton MRCancelar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton super 
         Caption         =   "Meter al Super Agente como nº"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Frame Frame333 
         Caption         =   "Modificar Regla"
         Height          =   2775
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   3015
         Begin VB.TextBox inc_pri 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox inc_peso 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   12
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton ModificarRegla 
            Caption         =   "Modificar"
            Height          =   375
            Left            =   960
            TabIndex        =   11
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox inc_pos 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox inc_agente 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox inc_regla 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "y prioridad"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "con peso"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "en posicion"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   960
            Width           =   975
         End
         Begin VB.Label label222 
            Alignment       =   1  'Right Justify
            Caption         =   "en agente"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   975
         End
         Begin VB.Label label333 
            Alignment       =   1  'Right Justify
            Caption         =   "Regla"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.Frame fr_Todas 
      Caption         =   "Todos los Agentes"
      Height          =   6495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox Lista5 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   6135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Frame Fr_Ejecucion 
      Caption         =   "Los 20 mejores y los 20 peores"
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox Lista3 
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
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3120
         Width           =   8895
      End
      Begin VB.TextBox Lista1 
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
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Image Imagen 
      Height          =   1680
      Left            =   600
      Picture         =   "c0_ce.frx":030A
      Top             =   600
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label aviso_ejecutar 
      BackStyle       =   0  'Transparent
      Caption         =   "Para comenzar pulse ""Comenzar"" en el menu ""Ejecutar"" (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frm_c0_ce"
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

Private Sub Form_GotFocus()
    Me.BackColor = cct_ejv(cfondo_ejv)

End Sub

Private Sub ModificarRegla_Click()

    Dim cont_regla  As Integer
    Dim pos As Integer
    Dim ag As Integer
    Dim peso As Integer
    Dim pri As Integer
    
    Dim nuevo_agente As String
    
    pos = CInt("0" & inc_pos)
    ag = CInt("0" & inc_agente)
    peso = CLng("0" & inc_peso)
    pri = CInt("0" & inc_pri)
    
    If pos = 0 Or ag = 0 Then Exit Sub
    
    nuevo_agente = ""
    For cont_regla = 1 To numero_de_reglas_por_agente_3r
        If cont_regla = pos Then
            nuevo_agente = nuevo_agente & inc_regla
            peso_regla_agente_3r(ag, pos) = peso
            prioridad_regla_agente_3r(ag, pos) = pri
        Else
            nuevo_agente = nuevo_agente & f_tomar_regla_de_agente_3r(ag, cont_regla)
        End If
    Next cont_regla
    agente_3r(ag) = nuevo_agente

End Sub


Private Sub Form_Activate()
    
    's_identificar_num_prg_activo_ejv

    'No permito refrescar mundo porque no hay
    s_cambiar_estado_enabled_menus_ejv CTE_VER_REFRESCAR, False
    
    'Sí permito guardar y abrir
    s_cambiar_estado_enabled_operaciones_ficheros_ejv True
    
    s_estado_enabled_ejecucion_ejv
    s_estado_enabled_ver_ejv
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub Form_Load()
    
    'Permito recibir teclas
    Me.KeyPreview = True
    
    'Mostramos la pantalla total
    Me.WindowState = CTE_MAXIMIZED
    
    'Imagen
    s_mostrar_aviso_imagen
    s_load_carga_3r

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = f_control_cerrar_ce0

End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Ahora no permito guardar y abrir
    s_cambiar_estado_enabled_operaciones_ficheros_ejv False

    Unload frm_c3_in3r
     
    'Actualizo el estado de enabled de ejecucion y ver
    num_prg_activo_ejv = CTE_NINGUNO
    
    'Cogiendolo de los arrays del num_prg_activo_ejv
    s_estado_enabled_ejecucion_ejv
    s_estado_enabled_ver_ejv
     
    'Pongo habilitado todos los programas
    s_cambiar_estado_enabled_programas_todos_ejv True
     
    'Grabo los ficheros sin cerrarlos, por si hay un corte de luz y esas cosas
    s_grabar_fichero_salida_ejv CTE_FIC_20_GLOLOG
    s_grabar_fichero_salida_ejv CTE_FIC_21_GLOTXT
    s_grabar_fichero_salida_ejv CTE_FIC_22_GLOXLS
     
End Sub

Private Sub MRCancelar_Click()
    Fr_ModificarAgente.Visible = False

End Sub

Private Sub super_Click()

    Dim i As Integer
    Dim cont_regla As Integer
    Dim max_regla As Integer
    Dim num_reglas_preparadas As Integer
    
        
    If MsgBox("Introducir el super-agente provocará probablemente una gran modificación de la población ¿Desea introducirlo?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    
    num_reglas_preparadas = 111
    
    ReDim super_regla(1 To num_reglas_preparadas) As String * 10
    ReDim super_peso_regla(1 To num_reglas_preparadas) As Long
    ReDim super_prioridad_regla(1 To num_reglas_preparadas) As Integer
    
    'Horizontales P
    super_regla(1) = "PPV******3"
    super_regla(2) = "PVP******2"
    super_regla(3) = "VPP******1"
    
    super_regla(4) = "***PPV***6"
    super_regla(5) = "***PVP***5"
    super_regla(6) = "***VPP***4"
    
    super_regla(7) = "******PPV9"
    super_regla(8) = "******PVP8"
    super_regla(9) = "******VPP7"
    
    'Verticales P
    super_regla(10) = "P**P**V**7"
    super_regla(11) = "P**V**P**4"
    super_regla(12) = "V**P**P**1"
    
    super_regla(13) = "*P**P**V*8"
    super_regla(14) = "*P**V**P*5"
    super_regla(15) = "*V**P**P*2"
    
    super_regla(16) = "**P**P**V9"
    super_regla(17) = "**P**V**P6"
    super_regla(18) = "**V**P**P3"
    
    'Diagonales P
    super_regla(19) = "P***P***V9"
    super_regla(20) = "P***V***P5"
    super_regla(21) = "V***P***P1"
    
    super_regla(22) = "**P*P*V**7"
    super_regla(23) = "**P*V*P**5"
    super_regla(24) = "**V*P*P**3"
    
    For cont_regla = 1 To 24
        super_peso_regla(cont_regla) = 1
        super_prioridad_regla(cont_regla) = 3
    Next cont_regla
    
    
    'Horizontales C
    super_regla(25) = "CCV******3"
    super_regla(26) = "CVC******2"
    super_regla(27) = "VCC******1"
    
    super_regla(28) = "***CCV***6"
    super_regla(29) = "***CVC***5"
    super_regla(30) = "***VCC***4"
    
    super_regla(31) = "******CCV9"
    super_regla(32) = "******CVC8"
    super_regla(33) = "******VCC7"
    
    'Verticales P y C
    super_regla(34) = "C**C**V**7"
    super_regla(35) = "C**V**C**4"
    super_regla(36) = "V**C**C**1"
    
    super_regla(37) = "*C**C**V*8"
    super_regla(38) = "*C**V**C*5"
    super_regla(39) = "*V**C**C*2"
    
    super_regla(40) = "**C**C**V9"
    super_regla(41) = "**C**V**C6"
    super_regla(42) = "**V**C**C3"
    
    'Diagonales C
    super_regla(43) = "C***C***V9"
    super_regla(44) = "C***V***C5"
    super_regla(45) = "V***C***C1"
    
    super_regla(46) = "**C*C*V**7"
    super_regla(47) = "**C*V*C**5"
    super_regla(48) = "**V*C*C**3"
    
    For cont_regla = 25 To 48
        super_peso_regla(cont_regla) = 1
        super_prioridad_regla(cont_regla) = 2
    Next cont_regla
    
    
    super_regla(49) = "VVVVVVVVV5"
    super_peso_regla(49) = 50
    super_prioridad_regla(cont_regla) = 3
    
    super_regla(50) = "****V****5"
    super_peso_regla(cont_regla) = 1
    super_prioridad_regla(cont_regla) = 1
    
    
    super_regla(51) = "CVVVVVVVV5"
    super_regla(52) = "VCVVVVVVV5"
    super_regla(53) = "VVCVVVVVV5"
    super_regla(54) = "VVVCVVVVV5"
    super_regla(55) = "VVVVVCVVV5"
    super_regla(56) = "VVVVVVCVV5"
    super_regla(57) = "VVVVVVVCV5"
    super_regla(58) = "VVVVVVVVC5"
    For cont_regla = 51 To 58
        super_peso_regla(cont_regla) = 1
        super_prioridad_regla(cont_regla) = 3
    Next cont_regla
    
    
    super_regla(59) = "CVVVPVVVV3"
    super_regla(60) = "VCVVPVVVV3"
    super_regla(61) = "VVCVPVVVV6"
    super_regla(62) = "VVVCPVVVV7"
    super_regla(63) = "VVVVPCVVV3"
    super_regla(64) = "VVVVPVCVV9"
    super_regla(65) = "VVVVPVVCV9"
    super_regla(66) = "VVVVPVVVC7"
    For cont_regla = 59 To 66
        super_peso_regla(cont_regla) = 1
        super_prioridad_regla(cont_regla) = 3
    Next cont_regla
    
    
    'Mas reglas
    super_regla(67) = "P*V******3"
    super_regla(68) = "P*****V**7"
    super_regla(69) = "P*******V9"
    super_regla(70) = "V*P******1"
    super_regla(71) = "**P*****V9"
    super_regla(72) = "**P***V**7"
    For cont_regla = 67 To 72
        super_peso_regla(cont_regla) = 1
        super_prioridad_regla(cont_regla) = 1
    Next cont_regla
    
    'Reglas de Ender
    super_regla(72) = "CVVVPVVVC2"
    super_regla(73) = "CVVVPVVVC4"
    super_regla(74) = "CVVVPVVVC6"
    super_regla(75) = "CVVVPVVVC8"
    
    super_regla(76) = "VVCVPVCVV2"
    super_regla(77) = "VVCVPVCVV4"
    super_regla(78) = "VVCVPVCVV6"
    super_regla(79) = "VVCVPVCVV8"
    
    super_regla(80) = "VVVCPVVVV7"
    super_regla(81) = "VCVVPVVVV1"
    super_regla(82) = "VVVVPCVVV3"
    super_regla(83) = "VVVVPVVCV9"
    
    super_regla(84) = "VVVCPVVVV1"
    super_regla(85) = "VCVVPVVVV3"
    super_regla(86) = "VVVVPCVVV9"
    super_regla(87) = "VVVVPVVCV7"
    
    super_regla(88) = "VVCCPVPVV9"
    super_regla(89) = "PCVVPVVVC7"
    super_regla(90) = "VVPVPCCVV1"
    super_regla(91) = "CVVVPVVCP3"
    
    super_regla(92) = "VVCVCVPVV1"
    super_regla(93) = "VVCVCVPVV9"
    
    super_regla(94) = "PVVVCVVVC3"
    super_regla(95) = "PVVVCVVVC7"
    
    super_regla(96) = "VVPVCVCVV1"
    super_regla(97) = "VVPVCVCVV9"
    
    super_regla(98) = "CVVVCVVVP3"
    super_regla(99) = "CVVVCVVVP7"
    
    super_regla(100) = "VVVVCVVVV1"
    super_regla(101) = "VVVVCVVVV3"
    super_regla(102) = "VVVVCVVVV7"
    super_regla(103) = "VVVVCVVVV9"
    
    super_regla(104) = "PVVCPVVVC7"
    super_regla(105) = "PVVCPVVVC3"
    
    super_regla(106) = "VCPVPVCVV1"
    super_regla(107) = "VCPVPVCVV9"
    
    super_regla(108) = "CVVVPCVVP3"
    super_regla(109) = "CVVVPCVVP7"
    
    super_regla(110) = "VVCVPVPCV1"
    super_regla(111) = "VVCVPVPCV9"
    
    
    For cont_regla = 72 To 111
        super_peso_regla(cont_regla) = 1
        super_prioridad_regla(cont_regla) = 3
    Next cont_regla
    
    
    
    
    '======================
    
    max_regla = numero_de_reglas_por_agente_3r
    If numero_de_reglas_por_agente_3r > num_reglas_preparadas Then
        max_regla = num_reglas_preparadas
    End If
    
    agente_3r(superage.Text) = ""
    
    For cont_regla = 1 To max_regla
        agente_3r(superage.Text) = agente_3r(superage.Text) & super_regla(cont_regla)
        peso_regla_agente_3r(superage.Text, cont_regla) = super_peso_regla(cont_regla)
        prioridad_regla_agente_3r(superage.Text, cont_regla) = super_prioridad_regla(cont_regla)
    Next cont_regla
    
    'Si no he rellenado todo
    If numero_de_reglas_por_agente_3r > num_reglas_preparadas Then
        For cont_regla = max_regla + 1 To numero_de_reglas_por_agente_3r
            agente_3r(superage.Text) = agente_3r(superage.Text) & "VVVVVVVVV5"
            peso_regla_agente_3r(superage.Text, cont_regla) = 1
            prioridad_regla_agente_3r(superage.Text, cont_regla) = 1
        Next cont_regla
    End If
    
    
    peso_agente_ce0(superage.Text) = peso_agente_ce0(superage.Text) + 500
    ciclo_nacimiento_agente_3r(superage.Text) = ciclo_ejv
    
    
    
    MsgBox "El super-agente has sido introducido como agente nº " & superage.Text, vbInformation

End Sub



