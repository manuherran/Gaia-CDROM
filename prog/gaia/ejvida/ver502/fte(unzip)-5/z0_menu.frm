VERSION 5.00
Begin VB.Form frm_z0_menu 
   Caption         =   "Menú"
   ClientHeight    =   4440
   ClientLeft      =   2640
   ClientTop       =   2115
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "z0_menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4440
   ScaleWidth      =   7005
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3240
      Top             =   3960
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
      Begin VB.ComboBox Cb_ejemplo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Op_2 
         Caption         =   "Crear un nuevo ejemplo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Op_1 
         Caption         =   "Ejecutar un ejemplo de demostración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.OptionButton Op_3 
         Caption         =   "Cargar un ejemplo guardado previamente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   4335
      End
   End
   Begin VB.CommandButton Aceptar 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Mensaje 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label quedesea 
      Alignment       =   2  'Center
      Caption         =   "¿Qué desea hacer?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frm_z0_menu"
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

    Dim opcion As Integer
    
    If frm_z0_menu.Op_1 Then
        opcion = 1
    ElseIf frm_z0_menu.Op_2 Then
        opcion = 2
    ElseIf frm_z0_menu.Op_3 Then
        opcion = 3
        MsgBox "Esta opción todavia no está disponible", vbInformation
    End If

    s_aceptar_menu_ejv frm_z0_menu.Cb_ejemplo, opcion

End Sub

Private Sub Cancelar_Click()
    num_prg_activo_ejv = num_prg_anterior_activo_ejv
    Unload Me
End Sub

Private Sub Cb_ejemplo_Click()
    
    s_ejemplo_click_ejv frm_z0_menu.Cb_ejemplo

End Sub

Private Sub Form_Load()
   
    s_load_menu_ejv
    
End Sub
Private Sub Timer1_Timer()
    
    mensaje.Visible = False
    Timer1.Enabled = False
End Sub
