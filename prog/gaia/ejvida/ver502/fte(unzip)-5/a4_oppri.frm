VERSION 5.00
Begin VB.Form frm_a4_oppri 
   Caption         =   "Opciones de Prisionero"
   ClientHeight    =   5670
   ClientLeft      =   225
   ClientTop       =   975
   ClientWidth     =   10710
   Icon            =   "A4_OPPRI.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5670
   ScaleWidth      =   10710
   Begin VB.Frame Frame1 
      Caption         =   "Puntos Obtenidos"
      Height          =   1815
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   3975
      Begin VB.TextBox el_que_defrauda 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2880
         TabIndex        =   35
         Text            =   "5"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox el_que_coopera 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   32
         Text            =   "0"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox ambos_defraudan 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Text            =   "0"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox ambos_cooperan 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Text            =   "3"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "El que Defrauda obtiene"
         Height          =   615
         Left            =   2880
         TabIndex        =   36
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "El que Coopera obtiene"
         Height          =   615
         Left            =   2160
         TabIndex        =   34
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Si uno Coopera, y el otro Defrauda"
         Height          =   495
         Left            =   2160
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Si ambos Defraudan, ambos obtienen"
         Height          =   855
         Left            =   1200
         TabIndex        =   31
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Si ambos Cooperan, ambos obtienen"
         Height          =   1095
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox Op_ProbbError 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   9840
      TabIndex        =   25
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox Cb_TipoAgente 
      BackColor       =   &H8000000E&
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox nuevo_numero 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4320
      TabIndex        =   18
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton todos 
      Caption         =   "Todos"
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox tipos 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   6735
   End
   Begin VB.TextBox Op_num_part 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   9840
      TabIndex        =   14
      Text            =   "10"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Op_EnergiaInicialAgente 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   9840
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Op_energiaConsumidaReproducirse 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   9840
      TabIndex        =   8
      Text            =   "2"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Op_EnergiaConsumUnaPos 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   9840
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox alejar 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   9840
      TabIndex        =   6
      Text            =   "10"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Modificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tipo de Juego"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton Op_nTodosContraTodos 
         Caption         =   "Mundo de agentes"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Op_TodosContraTodos 
         Caption         =   "Todos contra todos clasico"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Probabilidad de error en la decisión (%)"
      Height          =   255
      Left            =   6720
      TabIndex        =   26
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Modificar la cantidad de agentes de un tipo"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Código"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nueva cantidad"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre del tipo de Agente"
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Nº Partidas que juegan cada vez"
      Height          =   255
      Left            =   7320
      TabIndex        =   15
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía inicial que posee cada agente"
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía consumida por un agente al reproducirse"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía consumida por un agente al moverse una posición"
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de posiciones que se alejan dos agentes después de jugar una partida"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frm_a4_oppri"
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
    
           
    s_grabar_opciones_pri
    Unload Me


End Sub

Private Sub Cancelar_Click()
    Unload Me

End Sub

Private Sub Form_Load()
    
    'Cargo las opciones y los tipos (son mas opciones)
    s_cargar_opciones_pri
    s_cargar_tipos_agentes_pri

End Sub

Private Sub Modificar_Click()

    Dim pos As Integer
    Dim numero As Integer
    esta_modificado_num_agen_tipo_pri = True
        
    pos = InStr(Cb_TipoAgente.Text, ":")
    numero = CInt(Left(Cb_TipoAgente.Text, pos - 1))
    numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(numero) = CInt(nuevo_numero)
    
    'Cargo los nuevos tipos
    s_cargar_tipos_agentes_pri
End Sub

Private Sub Modificar_Mapa_Click()
    frm_a0_mapa.Show CTE_AMODAL

End Sub

Private Sub todos_Click()
    
    Dim i As Integer
    
    If IsNumeric(nuevo_numero.Text) Then
        esta_modificado_num_agen_tipo_pri = True
        For i = 1 To num_tipos_agentes_va0
            numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(i) = CInt(nuevo_numero.Text)
        Next i
        
        'Cargo los nuevos tipos
        s_cargar_tipos_agentes_pri
    End If

End Sub
