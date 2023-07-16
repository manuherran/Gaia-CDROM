VERSION 5.00
Begin VB.Form frm_c3_ev3r 
   Caption         =   "Método de Evaluación"
   ClientHeight    =   6045
   ClientLeft      =   1155
   ClientTop       =   1425
   ClientWidth     =   9480
   Icon            =   "C3_EV3R.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   9480
   Begin VB.CheckBox Op_Pesos_Partir_Cero 
      Caption         =   "Calcular los pesos cada vez partiendo de cero"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Width           =   3615
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
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   9135
   End
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   3
      Left            =   8640
      TabIndex        =   16
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   15
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   14
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Ayuda 
      Caption         =   "Ayuda"
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox Op_Relativo 
      Caption         =   "Aplicar a los pesos resultantes el metodo del peso relativo en función de si existen otras entidades de mayor peso"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   8295
   End
   Begin VB.ComboBox Op_Metodo_de_asignar_pesos_3r 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ComboBox Op_Personas_por_grupo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Op_Numero_cercanos_relativo_3r 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Text            =   "1"
      Top             =   5160
      Width           =   495
   End
   Begin VB.CheckBox Op_JugadoresAzar 
      Caption         =   "Desordenar todas las entidades antes de evaluarlas"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.TextBox Op_NumeroPartidas_3r 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Text            =   "1"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "parecidas a la que se está evaluando, teniendo en cuenta las "
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "entidades mas cercanas"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Para evaluar las entidades, jugar"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "y asignar un peso a cada entidad segun el método:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Label mensaje 
      Caption         =   "partidas en grupos de"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "frm_c3_ev3r"
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
    s_grabar_opciones_ev_3r
    Unload Me
End Sub

Private Sub Ayuda_Click(Index As Integer)

Dim texto As String
texto = ""
Select Case Index
    Case 0
        texto = texto & "Se recomienda desordenar las entidades antes de evaluarlas."
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "Hacerlo produce una tendencia a un comportamento elitista, ya que las que han sido buenas jugarán con una mayor probabilidad, contra las que han sido malas y tendrán mas probabilidades de ganar."
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "No hacerlo favorece a las peores entidades y produce una mayor variedad: una entidad buena podría jugar contra otra todavía mejor y perder, y en cambio una mala jugar contra otra peor y ganar."
    Case 1
        texto = texto & "El método básico para evaluar dos entidades es hacer que jueguen una única partida."
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "Sin embargo, cuantas más partidas juegue una entidad buena, más facilmente podrá demostrar que lo es, por lo que es interesante que cada pareja juegue varias partidas contra su contrincante en vez de una sóla."
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "Existe el inconveniente del calculo adicional que esto supone, por lo que el número de partidas no podrá ser muy elevado."
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "Por otra parte, en vez de evaluar las entidades de dos en dos, es posible hacer pequeños torneos en un reducido grupo para la asignación de los nuevos pesos."
        texto = texto & vbCrLf & vbCrLf
    Case 2
        texto = texto & "El método A suma en cada partida, 1 al peso de la entidad ganadora y resta 1 a la perdedora, quedando como estaban en caso de tablas."
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "El método B es más complejo, y dependiendo del resultado de cada partida y de quien la haya empezado, modifica los valores de los pesos de la siguiente forma."
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "No empieza y Sí gana         --> +4"
        texto = texto & vbCrLf
        texto = texto & "Sí empieza y Sí gana         --> +2"
        texto = texto & vbCrLf
        texto = texto & "No empieza y queda en tablas --> +1"
        texto = texto & vbCrLf
        texto = texto & "Sí empieza y queda en tablas --> -1"
        texto = texto & vbCrLf
        texto = texto & "No empieza y pierde          --> -2"
        texto = texto & vbCrLf
        texto = texto & "Sí empieza y pierde          --> -4"
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "El método C es como el B multiplicando dicho valor por el número de reglas usadas."
        texto = texto & vbCrLf & vbCrLf
    Case 3
        texto = texto & "El método del peso relativo consiste en disminuir el peso de aquellas entidades que sean parecidas a alguna otra."
        texto = texto & vbCrLf & vbCrLf
        texto = texto & "De esta forma se asegura la variedad pero se hace el proceso mucho mas lento, por lo que si el numero de entidades es medio/alto se debe restringir la comparación a las entidades mas cercanas."
        texto = texto & vbCrLf & vbCrLf
    
    Case Else
        MsgBox "Error: color inexistente", vbCritical
End Select

txt_Ayuda.Text = texto

End Sub

Private Sub Cancelar_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    
    s_centrar_ventana_ejv Me
    txt_Ayuda.Text = vbCrLf & vbCrLf & vbCrLf & vbCrLf & "          Pulse Ayuda para una mejor explicacion de cada opción."

    Op_Personas_por_grupo.Clear
    Op_Personas_por_grupo.AddItem "2 Jugadores"
    Op_Personas_por_grupo.AddItem "Todos contra todos"
    Op_Personas_por_grupo.ListIndex = 0

    Op_Metodo_de_asignar_pesos_3r.Clear
    Op_Metodo_de_asignar_pesos_3r.AddItem "Método A"
    Op_Metodo_de_asignar_pesos_3r.AddItem "Método B"
    Op_Metodo_de_asignar_pesos_3r.AddItem "Método C"
    Op_Metodo_de_asignar_pesos_3r.ListIndex = 0
    
    s_cargar_opciones_ev_3r


End Sub

Private Sub MasOpciones_Click()

End Sub
