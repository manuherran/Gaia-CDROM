VERSION 5.00
Begin VB.Form frm_z0_agra 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5460
   ClientLeft      =   2250
   ClientTop       =   1785
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   -1  'True
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5460
   ScaleWidth      =   6735
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton Aceptar 
         Cancel          =   -1  'True
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label mensaje 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-----------------------------------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   600
         TabIndex        =   3
         Top             =   840
         Width           =   5175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         DrawMode        =   14  'Copy Pen
         X1              =   480
         X2              =   6000
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Li_Linea_1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         DrawMode        =   14  'Copy Pen
         X1              =   480
         X2              =   6000
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Muchas gracias por vuestra ayuda y sugerencias:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   4230
      End
   End
End
Attribute VB_Name = "frm_z0_agra"
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

    frm_z0_acer.Timer1.Enabled = True
    Unload Me

End Sub

Private Sub Form_Load()
   
   Dim txt As String
   txt = ""
   txt = txt & "Bego Revenga, por la imagen original del papiro en ""Computación Evolutiva"", Edu Herrán por  la imagen de la hormiga rosa en ""Vida Artificial"" y pruebas de Hormigas y Plantas. "
   txt = txt & "Charles Dumont (Algoritmos Fractales de hojas), Phil Beffrey and Mattias Fagerlund in comp.ai.genetic (Gaussian Functions). "
   txt = txt & "Koldo Gotzon Ayuso (Experimentos sobre opciones del Tres en Raya), Juan Antonio Tubío (Nuevas opciones del Prisionero y varios jugadores), Natacha (Pruebas de ficheros de Ayuda). "
   txt = txt & "Dani y Pedro Gascón (Opciones de ""Hormigas y Plantas""), Nacho Estevas (Pruebas de Instalación), Ender y Chipiron (Diseño del super Agente del ""Tres en Raya"", relaciones públicas, propaganda y kalimotxo), Mariano Revenga (Funciones Aleatorias y número PI). "
   txt = txt & "Toda la gente de es.comp.lenguajes.visual-basic, y a todos los que habeis diseñado jugadores para el prisionero: jmfm, roger, cristina, julia, tona, canamares, star, mfdez, lancelot, night, mgalera, acayetano, maguilos, LittleJohn, miguel, Francisco, polo, jet, buso, margol, xouba, wen, Nactus, g_oteiza (geos), pinky... "
   txt = txt & "¡Si me dejo a alguien avisarme!"
   mensaje.Caption = txt
    
   s_centrar_ventana_ejv Me
    
End Sub


