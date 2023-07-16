VERSION 5.00
Begin VB.Form frm_c3_sel3r 
   Caption         =   "Método de Selección"
   ClientHeight    =   5595
   ClientLeft      =   1155
   ClientTop       =   1995
   ClientWidth     =   9480
   Icon            =   "C3_SEL3R.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   9480
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Seleción - ¿Quienes se reproducen?"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   6375
      Begin VB.CheckBox Check1 
         Caption         =   "Todos los que sobreviven se reproducen con el mismo número de hijos"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   5415
      End
   End
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   1
      Top             =   3480
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Seleción - ¿Quienes sobreviven?"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   6375
      Begin VB.OptionButton Op_s2 
         Caption         =   "B)  10% mejor sobreviven y se reproducen sobre 90%"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   5535
      End
      Begin VB.OptionButton Op_s1 
         Caption         =   "A)  50% mejor sobreviven y se reproducen sobre el otro 50%"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   5535
      End
      Begin VB.OptionButton Op_s3 
         Caption         =   "C)  20% mejor sobreviven y se reproducen sobre 80%"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   5175
      End
      Begin VB.OptionButton Op_s4 
         Caption         =   "D)  40% mejor y 10% peor sobre 50%"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   3615
      End
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   5160
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
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frm_c3_sel3r"
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
    s_grabar_opciones_sel_3r
    Unload Me
End Sub

Private Sub Cancelar_Click()
    Unload Me

End Sub
Private Sub Form_Load()
   
    s_centrar_ventana_ejv Me
    txt_Ayuda.Text = vbCrLf & vbCrLf & vbCrLf & vbCrLf & "          Pulse Ayuda para una mejor explicacion de cada opción."
    s_cargar_opciones_sel_3r

End Sub

