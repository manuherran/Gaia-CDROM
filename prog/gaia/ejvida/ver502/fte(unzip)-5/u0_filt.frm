VERSION 5.00
Begin VB.Form frm_u0_filt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro"
   ClientHeight    =   4995
   ClientLeft      =   2280
   ClientTop       =   2055
   ClientWidth     =   4995
   ClipControls    =   0   'False
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
   Icon            =   "u0_filt.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4995
   ScaleWidth      =   4995
   Begin VB.CommandButton boton_guardar 
      Caption         =   "&Guardar"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton boton_abrir 
      Caption         =   "A&brir"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton boton_filtro 
      Caption         =   "&Filtro"
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox foto 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   1320
      ScaleHeight     =   2055
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Aceptar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "frm_u0_filt"
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
    
    Unload Me

End Sub


Private Sub boton_abrir_Click()
   
   
    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_UTIL)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Fichero de imagen"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Abrir"
    frm_z0_fic.File1.Pattern = "*.bmp;*.gif;*.jpg"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If Not cancelar_operacion_fichero_ejv Then
        On Error Resume Next
        Set foto.Picture = LoadPicture(nombre_fichero_ejv)
        foto.Visible = True
    End If

End Sub

Private Sub boton_filtro_Click()

    Dim filas As Integer
    Dim columnas As Integer
    
    'MsgBox "el color es " & foto.Point(100, 100)
    
    Screen.MousePointer = CTE_ARENA
    
    For filas = 1 To foto.Width / 8
        For columnas = 1 To foto.Height / 8
            'foto.PSet (filas, columnas), RGB(256, 0, 0)
            foto.PSet (filas, columnas), foto.Point(filas, columnas) * 2
        Next columnas
    Next filas
    
    'foto.Refresh
    Screen.MousePointer = CTE_DEFECTO

End Sub

