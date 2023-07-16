VERSION 5.00
Begin VB.Form frm_a1_tiposhyp 
   Caption         =   "Tipos de Hormigas"
   ClientHeight    =   7455
   ClientLeft      =   720
   ClientTop       =   810
   ClientWidth     =   10575
   Icon            =   "A1_TIHYP.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7455
   ScaleWidth      =   10575
   Begin VB.CommandButton Cambiar_a 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   4
      Left            =   9120
      TabIndex        =   43
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_a 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   9120
      TabIndex        =   42
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_a 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   3
      Left            =   9120
      TabIndex        =   41
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_a 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   9120
      TabIndex        =   40
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_a 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   2
      Left            =   9120
      TabIndex        =   39
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_a 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   9120
      TabIndex        =   38
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_a 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   37
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_a 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   9120
      TabIndex        =   36
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_a 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   0
      Left            =   9120
      TabIndex        =   34
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_a 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   9120
      TabIndex        =   33
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox nhi5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   21
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox nhi4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   20
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox nhi3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   19
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox nhi2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   18
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox nhi1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   17
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Caja4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   16
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Caja3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Caja2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Caja1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   13
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Tendencias_r 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   7800
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_r 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   0
      Left            =   7800
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_r 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   7800
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_r 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   1
      Left            =   7800
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_r 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   7800
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_r 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   2
      Left            =   7800
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_r 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   7800
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_r 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   3
      Left            =   7800
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Tendencias_r 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   7800
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Cambiar_r 
      Caption         =   "&Cambiar..."
      Height          =   255
      Index           =   4
      Left            =   7800
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Etiq 
      Alignment       =   2  'Center
      Caption         =   "Tendencias Absolutas Movimiento"
      Height          =   615
      Index           =   2
      Left            =   9120
      TabIndex        =   35
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Etiq 
      Alignment       =   2  'Center
      Caption         =   "Número de Hormigas Inicial"
      Height          =   495
      Index           =   0
      Left            =   6480
      TabIndex        =   32
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Etiq3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   31
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Etiq2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   30
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Etiq1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   29
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Etiq 
      Alignment       =   2  'Center
      Caption         =   "Reproducirse"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   28
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Etiq 
      Alignment       =   2  'Center
      Caption         =   "Luchar"
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   27
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Etiq 
      Alignment       =   2  'Center
      Caption         =   "Regar"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   26
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Etiq 
      Alignment       =   2  'Center
      Caption         =   "Mover"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   25
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Etiq_3 
      Alignment       =   2  'Center
      Caption         =   "¿Hay Hormiga?"
      Height          =   495
      Index           =   0
      Left            =   2160
      TabIndex        =   24
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Etiq_2 
      Alignment       =   2  'Center
      Caption         =   "¿Hay Planta?"
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Etiq_1 
      Alignment       =   2  'Center
      Caption         =   "Tipo de Hormiga"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Etiq 
      Alignment       =   2  'Center
      Caption         =   "Tendencias Relativas Movimiento"
      Height          =   615
      Index           =   1
      Left            =   7800
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frm_a1_tiposhyp"
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

Private Sub Aceptar_Click()
    
    Dim i As Integer
    Dim cont As Long
    
    Dim tot As Integer
    
    'Copio de los textos para hacer efectivo el cambio (si ha habido) de
    'tendencias de mov, asi permito modificar directamente el texto
    For i = 1 To num_tipos_agentes_va0
        'Copio el array
        For cont = 1 To CTE_8_DIR
            tendencia_rel_inicial_mov_tipo_agente_va0(cont, i) = f_elemento_listacomas(Tendencias_r(i - 1).Text, cont)
        Next cont
    Next i
    For i = 1 To num_tipos_agentes_va0
        'Copio el array
        For cont = 1 To CTE_8_DIR
            tendencia_abs_inicial_mov_tipo_agente_va0(cont, i) = f_elemento_listacomas(Tendencias_a(i - 1).Text, cont)
        Next cont
    Next i
    
    'Grabo el numero de hormigas que hay que crear por tipo
    s_grabar_tipos_hyp

    'Calculo el total que se debera crear de cada tipo y lo visualizo
    tot = 0
    For i = 1 To num_tipos_agentes_va0
        tot = tot + numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(i)
    Next i
    num_inic_horm_hyp = tot
    
    Me.Hide

End Sub

Private Sub Cambiar_a_Click(Index As Integer)
    
    s_cambiar_tendencias_mov_va0 CTE_ABSOLUTAS, Index

End Sub

Private Sub Cambiar_r_Click(Index As Integer)

    s_cambiar_tendencias_mov_va0 CTE_RELATIVAS, Index

End Sub

Private Sub Cancelar_Click()

    Me.Hide

End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    Dim cont As Integer
    
    ReDim temp(1 To CTE_8_DIR) As Long
    
    frm_a1_tiposhyp.nhi1.Text = numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(1)
    frm_a1_tiposhyp.nhi2.Text = numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(2)
    frm_a1_tiposhyp.nhi3.Text = numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(3)
    frm_a1_tiposhyp.nhi4.Text = numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(4)
    frm_a1_tiposhyp.nhi5.Text = numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(5)
    
    
    '22 Tendencias del movimiento
    For i = 1 To num_tipos_agentes_va0
        'Copio el array
        For cont = 1 To CTE_8_DIR
            temp(cont) = tendencia_rel_inicial_mov_tipo_agente_va0(cont, i)
        Next cont
        frm_a1_tiposhyp.Tendencias_r(i - 1).Text = f_array_l_a_listacomas(temp())
        'Copio el array
        For cont = 1 To CTE_8_DIR
            temp(cont) = tendencia_abs_inicial_mov_tipo_agente_va0(cont, i)
        Next cont
        frm_a1_tiposhyp.Tendencias_a(i - 1).Text = f_array_l_a_listacomas(temp())
    Next i
    frm_a1_tiposhyp.Refresh
    
End Sub

