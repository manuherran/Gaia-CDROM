VERSION 5.00
Begin VB.Form frm_z0_fic 
   Caption         =   "s_mensaje"
   ClientHeight    =   4770
   ClientLeft      =   2880
   ClientTop       =   2220
   ClientWidth     =   6165
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
   Icon            =   "Z0_FIC.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   6165
   Begin VB.TextBox tipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   3
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox txt_nom_fichero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3840
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3000
      Pattern         =   "*.mdb"
      TabIndex        =   7
      Top             =   480
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2895
   End
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
      Left            =   4440
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Aceptar 
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
      Left            =   4440
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lbl_fichero 
      Caption         =   "(Nombre completo de Archivo)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo de Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre de Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar en:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_z0_fic"
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
    
    Select Case tipo_operacion_formulario_fic_ejv
        Case CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
            If txt_nom_fichero = "" Then
                'Busco un fichero y no hay ninguno
                MsgBox "Error: Especifique un nombre de fichero.", vbCritical
                Exit Sub
            End If
            'Si no tiene punto extension se lo pongo
            If InStr(txt_nom_fichero, ".") = 0 And Len(txt_nom_fichero) > 0 Then
                txt_nom_fichero = f_anaiadir_punto_extension(txt_nom_fichero, tipo)
            End If
            If f_nombre_fichero(txt_nom_fichero) = txt_nom_fichero Then
                'supongo que es un nombre suelto, le añado el path
                lbl_fichero = f_nombre_completo(Dir1, txt_nom_fichero)
            Else
                'supongo que es un path completo
                lbl_fichero = txt_nom_fichero
            End If
            nombre_fichero_ejv = lbl_fichero
            nombre_fichero_ejv_es_solo_un_path_ejv = False
        Case CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
            'Si no tiene punto extension se lo pongo
            If InStr(txt_nom_fichero, ".") = 0 And Len(txt_nom_fichero) > 0 Then
                txt_nom_fichero = f_anaiadir_punto_extension(txt_nom_fichero, tipo)
            End If
            If txt_nom_fichero = "" Then
                nombre_fichero_ejv_es_solo_un_path_ejv = True
                nombre_fichero_ejv = lbl_fichero
            Else
                If f_nombre_fichero(txt_nom_fichero) = txt_nom_fichero Then
                    'supongo que es un nombre suelto, le añado el path
                    lbl_fichero = f_nombre_completo(Dir1, txt_nom_fichero)
                Else
                    'supongo que es un path completo
                    lbl_fichero = txt_nom_fichero
                End If
                nombre_fichero_ejv_es_solo_un_path_ejv = False
                nombre_fichero_ejv = lbl_fichero
            End If
        Case CTE_SELECCIONAR_CARPETA_OP_FICH
            nombre_fichero_ejv_es_solo_un_path_ejv = True
            If txt_nom_fichero = "" Then
                'No busco un fichero sino una carpeta, y hay una por defecto
                txt_nom_fichero = Dir1
                nombre_fichero_ejv = Dir1
            Else
                nombre_fichero_ejv = f_path_fichero(lbl_fichero, CTE_C_RAIZ)
            End If
        Case CTE_SELECCIONAR_LISTA_FICHEROS_OP_FICH
            nombre_fichero_ejv_es_solo_un_path_ejv = True
            If txt_nom_fichero = "" Then
                'No busco un fichero sino una carpeta, y hay una por defecto
                txt_nom_fichero = Dir1
                nombre_fichero_ejv = Dir1
            Else
                nombre_fichero_ejv = f_path_fichero(lbl_fichero, CTE_C_RAIZ)
            End If
            ReDim lista_ficheros_sin_path_ejv(1 To File1.ListCount + 1) As String
            For cont_fic_lista_ejv = 1 To File1.ListCount + 1
                lista_ficheros_sin_path_ejv(cont_fic_lista_ejv) = File1.List(cont_fic_lista_ejv - 1)
            Next cont_fic_lista_ejv
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: La operación a realizar con la ventana de fichero no existe"
    End Select
    
    Unload Me

End Sub
Private Sub Cancelar_Click()

    cancelar_operacion_fichero_ejv = True
    Unload Me

End Sub

Private Sub Dir1_Change()

    File1.Path = Dir1
    
    lbl_fichero = f_nombre_completo(Dir1, File1)
    txt_nom_fichero = File1

End Sub

Private Sub Dir1_Click()
    
    File1.Path = Dir1
    
    lbl_fichero = f_nombre_completo(Dir1, File1)
    txt_nom_fichero = File1

End Sub

Private Sub Drive1_Change()

On Error Resume Next
Dir1 = Drive1

End Sub

Private Sub File1_Click()

    lbl_fichero = f_nombre_completo(Dir1, File1)
    txt_nom_fichero = File1

End Sub
Private Sub File1_DblClick()
    File1_Click
    Aceptar_Click
End Sub

Private Sub Form_Activate()
    cancelar_operacion_fichero_ejv = False

End Sub

Private Sub Form_Load()

On Error GoTo carpeta_erronea

    Aceptar.Default = False
    'No puede ser default al enter porque es necesario
    'el enter en el tipo de archivo
    
    cancelar_operacion_fichero_ejv = False
    Drive1 = "c:\"
    
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Dir1 = nombre_fichero_ejv
        lbl_fichero = nombre_fichero_ejv
    Else
        Dir1 = f_path_fichero(nombre_fichero_ejv, CTE_C_RAIZ)
        'lbl_fichero = f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), File1)
        lbl_fichero = nombre_fichero_ejv
    End If
    
    Exit Sub
carpeta_erronea:
    
    Dir1 = path_largo_ejv(CTE_C_RAIZ)
    lbl_fichero = f_nombre_completo(path_largo_ejv(CTE_C_RAIZ), File1)

End Sub
Private Sub n_fichero_Change()

    lbl_fichero = f_nombre_completo(Dir1, txt_nom_fichero)

End Sub

Private Sub tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo error
    
    If KeyCode = vbKeyReturn Then '13  ENTER key
        File1.Pattern = tipo
    End If
    Exit Sub
error:
    
    MsgBox "El tipo de fichero es incorrecto. Un tipo correcto es ""*.bmp;*.gif"""

End Sub

