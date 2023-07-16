VERSION 5.00
Begin VB.Form frm_a4_juegopri 
   Caption         =   "Juego al Dilema del Prisionero"
   ClientHeight    =   3210
   ClientLeft      =   1350
   ClientTop       =   990
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3210
   ScaleWidth      =   7020
   Begin VB.CommandButton Defraudar 
      Caption         =   "&Defraudar"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox jugador 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "1"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Cerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Cooperar 
      Caption         =   "&Cooperar"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Agente Contrario"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label r1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label r2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Partida"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Jugar contra el nº"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Mis Acciones"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frm_a4_juegopri"
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

Private Sub Cerrar_Click()
   
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True 'permito recibir teclas

End Sub
