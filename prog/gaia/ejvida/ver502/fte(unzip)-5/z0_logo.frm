VERSION 5.00
Begin VB.Form frm_z0_logo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4815
   ClientLeft      =   1845
   ClientTop       =   1725
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   3
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Z0_LOGO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4815
   ScaleWidth      =   8535
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   3720
         TabIndex        =   1
         Top             =   1800
         Width           =   4455
         Begin VB.Label La_Aplicacion 
            AutoSize        =   -1  'True
            Caption         =   "hghjgh"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   555
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1290
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7335
         Begin VB.PictureBox Pi_Imagen 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   3375
            Left            =   120
            Picture         =   "Z0_LOGO.frx":030A
            ScaleHeight     =   3375
            ScaleWidth      =   3735
            TabIndex        =   5
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label La_Version 
            AutoSize        =   -1  'True
            Caption         =   "fgfdgfdg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            TabIndex        =   7
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Para Windows 95 y NT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4320
            TabIndex        =   6
            Top             =   1200
            Width           =   2685
         End
      End
      Begin VB.Label La_Coment 
         AutoSize        =   -1  'True
         Caption         =   "Este programa y sus ficheros fuente son grátis y de libre distribución."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   4200
         Width           =   4800
      End
   End
   Begin VB.Timer Ti_Reloj 
      Interval        =   3000
      Left            =   840
      Top             =   4860
   End
End
Attribute VB_Name = "frm_z0_logo"
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

Private Sub Form_Activate()
    
    s_inicializar_arrays_programa_ejv

End Sub

Private Sub Form_Load()

    s_centrar_ventana_ejv Me
    La_Version.Caption = version_aplicacion_ejv
    
    'Activamos el timer
    Ti_Reloj.Enabled = True
    
    'Actualizamos el nombre de la Aplicación
    La_Aplicacion.Caption = CTE_LOGO_APLICACION


End Sub
Private Sub Form_Unload(Cancel As Integer)

    'Desactivamos el timer
    Ti_Reloj.Enabled = False
    
    Screen.MousePointer = CTE_DEFECTO

End Sub

Private Sub Ti_Reloj_Timer()

    'Desactivamos el timer
    Ti_Reloj.Enabled = False

    'Descarga
    Unload Me

End Sub


