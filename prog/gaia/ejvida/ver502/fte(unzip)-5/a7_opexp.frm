VERSION 5.00
Begin VB.Form frm_a7_opexp 
   Caption         =   "Opciones de Explorando Mapas"
   ClientHeight    =   3375
   ClientLeft      =   1155
   ClientTop       =   975
   ClientWidth     =   9795
   Icon            =   "a7_opexp.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3375
   ScaleWidth      =   9795
   Begin VB.Frame Frame1 
      Caption         =   "Cuando dos agentes se encuentran"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      Begin VB.CheckBox Op_CompartirMapas 
         Caption         =   "Comparten sus mapas"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Op_Repulsion 
         Caption         =   "Se repelen"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   3
      Text            =   "15"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   2
      Text            =   "5"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de Agentes tipo Helicóptero"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de Agentes tipo Explorador"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frm_a7_opexp"
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
    
    s_grabar_opciones_exp
    Unload Me
    

End Sub

Private Sub Cancelar_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    'Cargo las opciones
    s_cargar_opciones_exp

End Sub
