VERSION 5.00
Begin VB.Form frm_a6_ingaia 
   Caption         =   "Plataforma Gaia"
   ClientHeight    =   3600
   ClientLeft      =   1095
   ClientTop       =   1155
   ClientWidth     =   11265
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
   Icon            =   "a6_ingai.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   11265
   Begin VB.PictureBox semaforo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   10920
      Picture         =   "a6_ingai.frx":0442
      ScaleHeight     =   735
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.ListBox Le_Acc 
      BackColor       =   &H8000000F&
      Height          =   645
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ListBox Le_Ent 
      BackColor       =   &H8000000F&
      Height          =   1230
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ListBox Le_Uni 
      BackColor       =   &H8000000F&
      Height          =   1035
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.ListBox Li_Acc 
      BackColor       =   &H8000000F&
      Height          =   645
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   5895
   End
   Begin VB.ListBox Li_Ent 
      BackColor       =   &H8000000F&
      Height          =   1230
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
   End
   Begin VB.ListBox Li_Uni 
      BackColor       =   &H8000000F&
      Height          =   1035
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   5895
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
      Left            =   9840
      TabIndex        =   17
      Top             =   1560
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
      Left            =   8520
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label_estado 
      Caption         =   "Detenido"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9240
      TabIndex        =   15
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label_titulo_estado 
      Caption         =   "Estado:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label_warning 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Espere, por favor..."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   8520
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label txt_senergía 
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
      Left            =   8520
      TabIndex        =   11
      Top             =   1320
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
      Left            =   9840
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label30 
      Caption         =   "Energía Total"
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
      Left            =   9840
      TabIndex        =   9
      Top             =   1320
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
      Left            =   8520
      TabIndex        =   8
      Top             =   1080
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
      Left            =   9840
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frm_a6_ingaia"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub Form_Load()

    Me.KeyPreview = True 'permito recibir teclas
    ajuste_color_controles_formulario_ejv Me

    'Mostramos la pantalla en el centro
    s_centrar_ventana_ejv Me
    
    s_cargar_etiquetas_gai

End Sub

