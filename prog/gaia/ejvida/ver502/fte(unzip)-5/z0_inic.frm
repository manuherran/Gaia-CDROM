VERSION 5.00
Begin VB.Form frm_z0_inic 
   Caption         =   "Editar Fichero de Inicio"
   ClientHeight    =   6825
   ClientLeft      =   3105
   ClientTop       =   2685
   ClientWidth     =   9630
   Icon            =   "z0_inic.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6825
   ScaleWidth      =   9630
   Begin VB.TextBox mensaje 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "z0_inic.frx":0442
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton EjecutarInstancia 
      Caption         =   "Ejecutar Instancia"
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox Expertos 
      Caption         =   "Ver Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   8040
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton CerrarTodo 
      Caption         =   "Cerrar Todo"
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Recargar 
      Caption         =   "Recargar"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton GrabarporDefecto 
      Caption         =   "Grabar por Defecto"
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Grabar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Editar 
      Caption         =   "Editar en Notepad"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   4545
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton Cerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Avanzadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frm_z0_inic"
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

Private Sub CerrarTodo_Click()

    s_fin_todo

End Sub

Private Sub Editar_Click()
    
    On Error Resume Next
    Dim RetVal
    RetVal = Shell("notepad.exe " & f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), CTE_nombreINICIO_TXT), 1)

End Sub

Private Sub EjecutarInstancia_Click()
    
    On Error Resume Next
    Dim RetVal
    RetVal = Shell(f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), App.EXEName), 1)

End Sub

Private Sub Expertos_Click()

    If Expertos = 1 Then
        Recargar.Visible = True
        Editar.Visible = True
        GrabarporDefecto.Visible = True
        GrabarporDefecto.Visible = True
        EjecutarInstancia.Visible = True
        CerrarTodo.Visible = True
    Else
        Recargar.Visible = False
        Editar.Visible = False
        GrabarporDefecto.Visible = False
        GrabarporDefecto.Visible = False
        EjecutarInstancia.Visible = False
        CerrarTodo.Visible = False
    End If

End Sub

Private Sub Form_Load()
   
frm_z0_inic.mensaje = ""
frm_z0_inic.mensaje = frm_z0_inic.mensaje & "Para obtener información acerca de un parámetro, hacer click en la columna izquierda." & vbCrLf
frm_z0_inic.mensaje = frm_z0_inic.mensaje & "Para modificarlo, hacer click en la columna derecha." & vbCrLf
frm_z0_inic.mensaje = frm_z0_inic.mensaje & "Para editar el fichero de configuración completo y otras opciones avanzadas, hacer click en Ver Opciones Avanzadas." & vbCrLf


s_centrar_ventana_ejv Me
s_mostrar_config_ejv

End Sub


Private Sub Grabar_Click()

    s_aut_grabar_inicio_txt False
    
End Sub

Private Sub GrabarporDefecto_Click()
    
    s_aut_grabar_inicio_txt True
    nombre_fichero_ejv = f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), CTE_nombreINICIO_TXT)
    s_aut_leer_inicio_txt
    s_mostrar_config_ejv

End Sub

Private Sub List1_Click()

    s_informar_parametro_config_ejv List1.ListIndex + 1

End Sub

Private Sub List2_Click()
    
    s_informar_parametro_config_ejv List2.ListIndex + 1
    s_modificar_parametro_config_ejv List2.ListIndex + 1
    
End Sub


Private Sub mensaje_Click()
    frm_z0_inic.mensaje = "Para obtener información acerca de un parámetro, hacer click en la columna izquierda. Para modificarlo, hacer click en la columna derecha. Para editar el fichero de configuración completo, hacer click en Editar."

End Sub

Private Sub Recargar_Click()
    
    nombre_fichero_ejv = f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), CTE_nombreINICIO_TXT)
    s_aut_leer_inicio_txt
    s_mostrar_config_ejv

End Sub
