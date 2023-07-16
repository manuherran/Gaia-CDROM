VERSION 5.00
Begin VB.Form frm_aA_opuva 
   Caption         =   "Opciones de Universo"
   ClientHeight    =   3375
   ClientLeft      =   1155
   ClientTop       =   975
   ClientWidth     =   9795
   Icon            =   "aA_opuva.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3375
   ScaleWidth      =   9795
   Begin VB.TextBox radioGrande 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   6
      Text            =   "5"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox radioPequenio 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   4
      Text            =   "5"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox numeroAgentes 
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tamaño esfera grande"
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tamaño esfera pequeña"
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de Agentes"
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frm_aA_opuva"
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
    
    s_grabar_opciones_uva
    Unload Me
    

End Sub

Private Sub Cancelar_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    'Cargo las opciones
    s_cargar_opciones_uva

End Sub

Private Sub radio_pequenio_uva_Change()

End Sub
