VERSION 5.00
Begin VB.Form frm_c3_rm 
   Caption         =   "Método de Reproducción - Mutaciones"
   ClientHeight    =   6105
   ClientLeft      =   810
   ClientTop       =   1260
   ClientWidth     =   9990
   Icon            =   "c3_rm.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   9990
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   11
      Left            =   9240
      TabIndex        =   34
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Op_Tipo_Mutacion 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      TabIndex        =   32
      Text            =   "20"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   30
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   7
      Left            =   8400
      TabIndex        =   29
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   28
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Op_TasaMutacion 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3600
      TabIndex        =   25
      Text            =   "20"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame Frame11 
      Caption         =   "Crear mutaciones en reglas repetidas"
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   2895
      Begin VB.CommandButton Ayuda 
         Caption         =   "Ayuda"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Op_CrearMutacionesEnRepetidas_3r 
         Caption         =   "Sí"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Op_nCrearMutacionesEnRepetidas_3r 
         Caption         =   "No"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.TextBox sust3 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5760
      TabIndex        =   18
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox sust1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox sust2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Top             =   3600
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "De padres idénticos hijos mutantes"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   2775
      Begin VB.CommandButton Ayuda 
         Caption         =   "Ayuda"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Op_nPadresIdenticos 
         Caption         =   "No"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Op_PadresIdenticos 
         Caption         =   "Sí"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frecuencia de Mutación"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   2295
      Begin VB.CommandButton Ayuda 
         Caption         =   "Ayuda"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Op_Acumulada 
         Caption         =   "Acumulada"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Op_nAcumulada 
         Caption         =   "Independiente"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   5640
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "% de las veces, y modificar un elemento de la regla el resto"
      Height          =   255
      Left            =   4920
      TabIndex        =   33
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Una mutación consiste en Generar una nueva regla el"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      Caption         =   "generaciones de reglas"
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label35 
      Caption         =   "Hay 1 mutación en la regla generada por cada "
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label16 
      Caption         =   "Sustituir por mutaciones las reglas con un peso menor o igual que"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label17 
      Caption         =   "ciclos, sustituir por mutaciones las reglas con un peso menor"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Cada"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "o igual que el"
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "% de la media de los pesos del mejor"
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label21 
      Caption         =   "% de la media de los pesos del mejor"
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "el"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   3000
      Width           =   255
   End
End
Attribute VB_Name = "frm_c3_rm"
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
    s_grabar_opciones_rep_mut_ce
    Unload Me

End Sub
Private Sub Cancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    s_centrar_ventana_ejv Me
    txt_Ayuda.Text = vbCrLf & vbCrLf & vbCrLf & vbCrLf & "          Pulse Ayuda para una mejor explicacion de cada opción."
    s_cargar_opciones_rep_mut_ce
    
End Sub
