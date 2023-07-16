VERSION 5.00
Begin VB.Form frm_c3_op3r 
   Caption         =   "Opciones del Tres en Raya"
   ClientHeight    =   6585
   ClientLeft      =   1155
   ClientTop       =   975
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6585
   ScaleWidth      =   7515
   Begin VB.ComboBox Cb_num_agentes 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Modo de Ejecución"
      Height          =   855
      Left            =   3480
      TabIndex        =   27
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Op_nVerAgentes 
         Caption         =   "Rápido"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Op_VerAgentes 
         Caption         =   "Ver Agentes"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Creación de Reglas"
      Height          =   855
      Left            =   5400
      TabIndex        =   24
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Op_nReglasAzar 
         Caption         =   "Preparadas"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Op_ReglasAzar 
         Caption         =   "Al azar"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Número de reglas por agente"
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6975
      Begin VB.TextBox var51 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox var41 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   13
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox var42 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5520
         TabIndex        =   12
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox var31 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox var32 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5520
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox var21 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox var22 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5520
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox var12 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5520
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox var11 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Op_numero_reglas_variable_3r 
         Caption         =   "Variable"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Op_nnumero_reglas_variable_3r 
         Caption         =   "Constante"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   2520
         Y1              =   240
         Y2              =   2520
      End
      Begin VB.Label Label26 
         Caption         =   "durante el resto de los ciclos"
         Height          =   255
         Left            =   3840
         TabIndex        =   23
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label24 
         Caption         =   "durante los siguientes"
         Height          =   255
         Left            =   3840
         TabIndex        =   22
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "ciclos"
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "durante los siguientes"
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "ciclos"
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "durante los siguientes"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "ciclos"
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "ciclos"
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "durante los primeros"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de agentes inicial"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   $"c3_op3r.frx":0000
      Height          =   1935
      Left            =   360
      TabIndex        =   31
      Top             =   3960
      Width           =   6615
   End
End
Attribute VB_Name = "frm_c3_op3r"
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
    
    Dim i As Integer
    Dim suma As Integer
   
    s_grabar_opciones_3r
    
    Unload Me

End Sub

Private Sub Cancelar_Click()
    
    Unload Me

End Sub


Private Sub Form_Load()
    
    Cb_num_agentes.Clear
    Cb_num_agentes.AddItem 8
    Cb_num_agentes.AddItem 20
    Cb_num_agentes.AddItem 40
    Cb_num_agentes.AddItem 80
    Cb_num_agentes.AddItem 160
    Cb_num_agentes.AddItem 320
    
    s_cargar_opciones_3r

End Sub
