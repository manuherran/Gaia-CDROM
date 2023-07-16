VERSION 5.00
Begin VB.Form frm_a4_inpri 
   Caption         =   "El juego del Dilema del Prisionero"
   ClientHeight    =   7230
   ClientLeft      =   5625
   ClientTop       =   1470
   ClientWidth     =   8325
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "a4_inpri.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7230
   ScaleWidth      =   8325
   Begin VB.PictureBox semaforo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   8040
      Picture         =   "a4_inpri.frx":0442
      ScaleHeight     =   735
      ScaleWidth      =   375
      TabIndex        =   42
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton OrdenarPorTipo 
      Caption         =   "Ordenar por Tipos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   15
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton OrdenarPorPeso 
      Caption         =   "Ordenar por Pesos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Top             =   0
      Width           =   1695
   End
   Begin VB.ListBox tipos 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Segundos totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6480
      TabIndex        =   44
      Top             =   5040
      Width           =   1230
   End
   Begin VB.Label SegTot 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   43
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   7680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label media 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   41
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Duraci�n de los ciclos (en segundos)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6480
      TabIndex        =   40
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label segundosr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   39
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Segundos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   38
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label minutosr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   37
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Minutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   36
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label horasr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Horas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   34
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label diasr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   33
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "D�as"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   32
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "A�os"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   31
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label anosr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   30
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Meses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   29
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label mesesr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   28
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label horaf 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5160
      TabIndex        =   27
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Hora Fin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label horac 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   5160
      TabIndex        =   25
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Hora Comienzo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6480
      TabIndex        =   24
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label fechaf 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5160
      TabIndex        =   23
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Fin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label fechac 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label27 
      Caption         =   "Fecha Comienzo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label txt_agente 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Agente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label txt_MuerenVejez 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lbl555 
      Caption         =   "N� Mueren Vejez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label_titulo_estado 
      Caption         =   "Estado:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label_estado 
      Caption         =   "Detenido"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label_warning 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Espere, por favor..."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label33 
      Caption         =   "N� mueren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label32 
      Caption         =   "N� nacen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label30 
      Caption         =   "Energ�a Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label29 
      Caption         =   "Ciclo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "Total Agentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label txt_mueren 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label txt_nacen 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label txt_senerg�a 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label txt_total_agentes 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label txt_ciclo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frm_a4_inpri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------
' Shareware desarrollado por Manuel de la Herr�n Gasc�n
' mherran@usa.net (Junio 1997 - Diciembre 1998) Madrid (Spain).
' http://www.geocities.com/SiliconValley/Vista/7491/
' -----------------------------------------------------------------------
' Este programa y sus ficheros fuente son gr�tis y de libre distribuci�n.
' El c�digo fuente est� disponible y puede ser modificado, distribuido,
' o utilizado en otros programas con entera libertad.
' -----------------------------------------------------------------------
' Para mantenerse informado de las sucesivas versiones del programa
' y d�nde conseguirlas, escriba un mail a mherran@usa.net
' Para sugerir posibles ampliaciones, enviar comentarios de cualquier tipo
' si se detectara alg�n error en la programaci�n o en la instalaci�n,
' o si se va a ampliar o utilizar una parte o todo este
' programa, no dude en ponerse en contacto con el autor.
'-----------------------------------------------------------------------
Private Sub Form_Load()
    
    Me.KeyPreview = True 'permito recibir teclas
    ajuste_color_controles_formulario_ejv Me
    frm_a4_inpri.tipos.Clear
    frm_a4_inpri.Refresh

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub OrdenarPorPeso_Click()
    
    mostrar_por_orden_de_pesos_pri = True
    s_mostrar_resumen_pesos_pri

End Sub

Private Sub OrdenarPorTipo_Click()
    mostrar_por_orden_de_pesos_pri = False
    s_mostrar_resumen_pesos_pri

End Sub
