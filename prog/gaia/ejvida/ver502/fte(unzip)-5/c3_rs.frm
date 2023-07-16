VERSION 5.00
Begin VB.Form frm_c3_rs 
   Caption         =   "Método de Reproducción - Sobrecruzamiento"
   ClientHeight    =   4365
   ClientLeft      =   810
   ClientTop       =   1260
   ClientWidth     =   9990
   Icon            =   "c3_rs.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   9990
   Begin VB.Frame Frame3 
      Caption         =   "Los genes se heredan de uno u otro progenitor"
      Height          =   855
      Left            =   4680
      TabIndex        =   22
      Top             =   1800
      Width           =   3615
      Begin VB.OptionButton Op_nCogerAlternos 
         Caption         =   "al azar de uno u otro"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Op_CogerAlternos 
         Caption         =   "de forma alterna"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton Ayuda 
         Caption         =   "Ayuda"
         Height          =   255
         Index           =   10
         Left            =   2640
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Las entidades comparten su conocimiento"
      Height          =   855
      Left            =   3960
      TabIndex        =   15
      Top             =   2760
      Width           =   3735
      Begin VB.TextBox Op_CompartirNumVecinos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Op_nCompartirConocimiento 
         Caption         =   "No"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton Op_CompartirConocimiento 
         Caption         =   "Sí"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton Ayuda 
         Caption         =   "Ayuda"
         Height          =   255
         Index           =   9
         Left            =   2760
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "vecinos"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Con"
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Las entidades modifican los pesos de sus reglas"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   3735
      Begin VB.CommandButton Ayuda 
         Caption         =   "Ayuda"
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Op_ModificarPropiosPesos 
         Caption         =   "Sí"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Op_nModificarPropiosPesos 
         Caption         =   "No"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Heredar regla de más peso"
      Height          =   855
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
      Begin VB.CommandButton Ayuda 
         Caption         =   "Ayuda"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Op_nHeredarReglaMasPeso_3r 
         Caption         =   "No"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton Op_HeredarReglaMasPeso_3r 
         Caption         =   "Sí"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Elección de Padres"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
      Begin VB.CommandButton Ayuda 
         Caption         =   "Ayuda"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Op_nPadresAzar_3r 
         Caption         =   "Secuencial"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Op_PadresAzar_3r 
         Caption         =   "Al azar"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txt_Ayuda 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "frm_c3_rs"
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
    s_grabar_opciones_rep_sob_ce
    Unload Me

End Sub
Private Sub Cancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    s_centrar_ventana_ejv Me
    txt_Ayuda.Text = vbCrLf & vbCrLf & vbCrLf & vbCrLf & "          Pulse Ayuda para una mejor explicacion de cada opción."
    s_cargar_opciones_rep_sob_ce

End Sub
