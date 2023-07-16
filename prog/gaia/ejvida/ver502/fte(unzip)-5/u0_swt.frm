VERSION 5.00
Begin VB.Form frm_u0_swt 
   Caption         =   "Generación de fichero swt para creación de programas de instalación en VB"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10035
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   10035
   Begin VB.TextBox Op_comenzar 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Text            =   "33"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Op_Extensiones 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Text            =   "*.*"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Op_Estado 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3720
      Width           =   9855
   End
   Begin VB.CommandButton Salir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Generar 
      Caption         =   "&Generar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Op_ListaOrigen 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton Fic_Ori 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Comenzar en"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Líneas a incluir en el fichero swt"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ficheros origen"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frm_u0_swt"
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



Private Sub Fic_Ori_Click()

    Dim i As Integer
    Dim texto As String

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_RAIZ)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_LISTA_FICHEROS_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Lista de Ficheros"  'Esto provoca la llamada, igual que un show  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    'frm_z0_fic.File1.Pattern = "*.*"
    frm_z0_fic.File1.Pattern = Op_Extensiones.Text
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    'Añado en la lista la lista de ficheros de esa carpeta

    texto = Op_ListaOrigen.Text
    For i = 1 To UBound(lista_ficheros_sin_path_ejv)
        If lista_ficheros_sin_path_ejv(i) <> "" Then
            texto = texto & f_nombre_completo_existente(nombre_fichero_ejv, lista_ficheros_sin_path_ejv(i)) & vbCrLf
        End If
    Next i
    Op_ListaOrigen.Text = texto

End Sub

Private Sub Form_Activate()

    Me.BackColor = cct_ejv(cfondo_ejv)
    ajuste_color_controles_formulario_ejv Me

End Sub

Private Sub Form_GotFocus()
    
    Me.BackColor = cct_ejv(cfondo_ejv)
    ajuste_color_controles_formulario_ejv Me

End Sub

Private Sub Generar_Click()

    Dim lista_ficheros() As String
    Dim num_fic As Integer
    Dim cont_fic As Integer
    Dim indice As Integer
    Dim linea As String
    Dim comenzar As Integer
    Dim texto As String
    Dim resto_path As String

    Screen.MousePointer = CTE_ARENA
    Op_Estado.Text = ""

    comenzar = CInt(Op_comenzar)
    'Paso la lista a un array
    num_fic = f_multiline2array(Op_ListaOrigen.Text, lista_ficheros())

    'Por cada fichero
    texto = ""
    For cont_fic = 1 To num_fic
        indice = comenzar + cont_fic - 1
        texto = texto & "File" & CStr(indice) & "=""" & lista_ficheros(cont_fic) & """,Verdadero,"""
        texto = texto & "$"
        If f_path_fichero(lista_ficheros(cont_fic), CTE_C_RAIZ) = path_largo_ejv(CTE_C_RAIZ) Then
            texto = texto & "(AppPath)"
        Else
            resto_path = restar_path_ejv(f_path_fichero(lista_ficheros(cont_fic), CTE_C_RAIZ), path_largo_ejv(CTE_C_RAIZ))
            If resto_path = "" Then
                texto = texto & "(AppPath)"
            Else
                texto = texto & "(AppPath)\" & resto_path
            End If
        End If
        texto = texto & """,,False" & vbCrLf
    Next cont_fic
    Op_Estado.Text = texto
'File32="C:\gaia\vida\fte_vb50\readme.txt",Verdadero,"$(AppPath)",,False
'File12="C:\gaia\vida\fte_vb50\map\ej07.map",Verdadero,"$(AppPath)\map",,False

    Screen.MousePointer = CTE_DEFECTO
    
    MsgBox "Las líneas para el swt han sido generadas.", vbInformation
    
    
End Sub

Private Sub Salir_Click()
    Unload Me

End Sub
