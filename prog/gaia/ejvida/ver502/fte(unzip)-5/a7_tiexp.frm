VERSION 5.00
Begin VB.Form frm_a7_tiposexp 
   Caption         =   "Tipos de Agentes"
   ClientHeight    =   7455
   ClientLeft      =   1380
   ClientTop       =   795
   ClientWidth     =   9405
   Icon            =   "a7_tiexp.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7455
   ScaleWidth      =   9405
   Begin VB.Frame Fr_Grafico 
      Caption         =   "Agente número..."
      Height          =   2655
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   4695
      Begin VB.ListBox lst_Datos_a_mostrar 
         BackColor       =   &H8000000F&
         Height          =   1815
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst_Color 
         BackColor       =   &H8000000F&
         Height          =   1815
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox lst_Tipo 
         BackColor       =   &H8000000F&
         Height          =   1815
         Left            =   2880
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Movimiento"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label txt1 
         Caption         =   "Color"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.ComboBox Cb_Gr_Especial 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton VerOpcionesAgente 
      Caption         =   "&Ver Agente"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Cb_eje_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Ver opciones del agente"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "de tipo"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frm_a7_tiposexp"
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
    
    Me.Hide

End Sub

Private Sub Cancelar_Click()
    Me.Hide

End Sub

