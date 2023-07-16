VERSION 5.00
Begin VB.Form frm_a0_opva 
   Caption         =   "Opciones de Vida Artificial"
   ClientHeight    =   4710
   ClientLeft      =   1155
   ClientTop       =   975
   ClientWidth     =   9795
   Icon            =   "a0_opva.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4710
   ScaleWidth      =   9795
   Begin VB.TextBox LimiteMuerte 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7080
      TabIndex        =   32
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.Frame Frame5 
      Caption         =   "Otro Aprendizaje a nivel de especie"
      Height          =   2535
      Left            =   6360
      TabIndex        =   25
      Top             =   120
      Width           =   3255
      Begin VB.CheckBox Ch_MetodoRecompensaNHijos 
         Caption         =   "Con la misma energía requerida, el número de hijos es igual al número de bits coincidentes con la cadena buscada"
         Height          =   1095
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox CadenaBinaria 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Text            =   "10001011010001000111011110011010101010010101010101010111101100010110010110101"
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox Op_BusquedaCadena 
         Caption         =   "Búsqueda de la cadena binaria"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Value           =   1  'Checked
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Probabilidades de Mutación iniciales"
      Height          =   2535
      Left            =   3000
      TabIndex        =   15
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Op_nPMPMCte 
         Caption         =   "La PM del resto de PM puede mutar según una PM que es ella misma"
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton Op_PMPMCte 
         Caption         =   "La PM del resto de PM es un valor constante"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.TextBox Op_PMPM 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Text            =   "10"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Op_PMMov 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Text            =   "10"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Op_PMColor 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Text            =   "10"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "PM del resto de PM"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "%"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "PM del movimiento"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "%"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "PM del tipo de agente"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "%"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reemplazamiento de Agentes"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2775
      Begin VB.TextBox Op_Muerte2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Text            =   "10"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Op_Muerte1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "10"
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Op_Inmortales 
         Caption         =   "Agentes Inmortales"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Op_nInmortales 
         Caption         =   "Los agentes mueren después"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "con una distribución gaussiana"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "y"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "de un número de ciclos entre"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Modo de Ejecución"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton Op_VerAgentes 
         Caption         =   "Ver Agentes"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Op_nVerAgentes 
         Caption         =   "Rápido"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lugar de Nacimiento"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
      Begin VB.OptionButton Op_nNacimientoCerca 
         Caption         =   "Al azar"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Op_NacimientoCerca 
         Caption         =   "Cerca de los padres"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Una agente muere si su energía es menor o igual que"
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   2760
      Width           =   3855
   End
End
Attribute VB_Name = "frm_a0_opva"
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
    
    s_grabar_opciones_generales_va0
    Unload Me
    'Me.Hide


End Sub

Private Sub Cancelar_Click()
    Unload Me
    'Me.Hide

End Sub

Private Sub Form_Load()
    
    s_cargar_opciones_generales_va0

End Sub

