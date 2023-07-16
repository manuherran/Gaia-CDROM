VERSION 5.00
Begin VB.Form frm_u0_dicc 
   Caption         =   "Generación de diccionario a partir de documentos html"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10035
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "u0_dicc.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6240
   ScaleWidth      =   10035
   Begin VB.CommandButton Editar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Op_Extensiones 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Text            =   "*.htm;*.html"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Op_Estado2 
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
      Height          =   2895
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   7695
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
      Height          =   2895
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2280
      Width           =   7695
   End
   Begin VB.CommandButton Salir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Generar 
      Caption         =   "&Generar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   5520
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
      Height          =   1215
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton Fic_Ori 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Op_FicheroDestino 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Text            =   "dicc.txt"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Op_CarpetaDestino 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "c:\......"
      Top             =   1680
      Width           =   4695
   End
   Begin VB.CommandButton Fic_Des 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de la Ejecución"
      Height          =   195
      Left            =   285
      TabIndex        =   12
      Top             =   2280
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ficheros htm origen"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1380
   End
   Begin VB.Label Label6 
      Caption         =   "Carpeta"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Fichero"
      Height          =   255
      Left            =   6840
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fichero destino"
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   1680
      Width           =   1080
   End
End
Attribute VB_Name = "frm_u0_dicc"
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

Private Sub Editar_Click()
    
    On Error Resume Next
    Dim RetVal
    Dim fic_destino As String
    
    fic_destino = f_nombre_completo(Op_CarpetaDestino, Op_FicheroDestino)
    If f_existe_fichero(fic_destino) Then
        RetVal = Shell("notepad.exe " & fic_destino, 1)
    End If
End Sub

Private Sub Fic_Des_Click()

    'Fijo una carpeta por defecto
    If Len(Trim(Op_CarpetaDestino.Text)) > 0 Then
        nombre_fichero_ejv = Trim(Op_CarpetaDestino.Text)
    Else
        nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_UTIL)
    End If
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta o carpeta mas fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.txt"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Op_CarpetaDestino.Text = nombre_fichero_ejv
    Else
        Op_CarpetaDestino.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_PRG_UTIL)
        Op_FicheroDestino = f_nombre_fichero(nombre_fichero_ejv)
    End If

End Sub

Private Sub Fic_Ori_Click()

    Dim i As Integer
    Dim texto As String

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_DOC_WEB)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_LISTA_FICHEROS_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Lista de Ficheros"  'Esto provoca la llamada, igual que un show  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    'frm_z0_fic.File1.Pattern = "*.htm;*.html"
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

Private Sub Form_Load()
    
    Op_CarpetaDestino.Text = path_largo_ejv(CTE_C_PRG_UTIL)
    frm_u0_dicc.Op_Estado.Visible = True
    frm_u0_dicc.Op_Estado2.Visible = False
    
End Sub

Private Sub Generar_Click()

    Dim fic_destino As String
    Dim lista_ficheros() As String
    Dim num_fic As Integer
    Dim cont_fic As Integer
    Dim cont_fic_leidos As Integer
    Dim linea As String
    Dim c As String * 1
    Dim palabra As String

    'Compruebo que no existe el destino
    fic_destino = f_nombre_completo(Op_CarpetaDestino, Op_FicheroDestino)
    If f_existe_fichero(fic_destino) Then
        If MsgBox("El fichero ya existe. ¿Desea reemplazarlo?", vbQuestion + vbOKCancel) = vbCancel Then
            'Ha elegido cancelar la operación
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = CTE_ARENA
    Op_Estado.Text = ""
    Op_Estado2.Text = ""

    'Abro el fichero de salida
    Open fic_destino For Output As #CTE_FIC_13_DICC_SAL

    'Paso la lista a un array
    num_fic = f_multiline2array(Op_ListaOrigen.Text, lista_ficheros())

    'Por cada fichero
    cont_fic_leidos = 0
    For cont_fic = 1 To num_fic
        If f_existe_fichero(lista_ficheros(cont_fic)) Then
            cont_fic_leidos = cont_fic_leidos + 1
            Op_Estado.Text = "Tratando el fichero " & lista_ficheros(cont_fic) & vbCrLf
            Op_Estado2.Text = Op_Estado2.Text & "Tratando el fichero " & lista_ficheros(cont_fic) & vbCrLf
            DoEvents
            Open lista_ficheros(cont_fic) For Input As #CTE_FIC_12_DICC_ENT
            While Not EOF(CTE_FIC_12_DICC_ENT)
                linea = f_leer_linea(CTE_FIC_12_DICC_ENT)
                If Len(linea) > 0 Then
                    linea = f_quitar_tags(linea)
                    If Len(linea) > 0 Then
                        If InStr(linea, Chr$(9)) <> 0 Then linea = f_sustituir_subcadena(linea, Chr$(9), " ") 'tabulador
                        If InStr(linea, "&nbsp;") <> 0 Then linea = f_sustituir_subcadena(linea, "&nbsp;", " ")
                        If InStr(linea, "&aacute;") <> 0 Then linea = f_sustituir_subcadena(linea, "&aacute;", "á")
                        If InStr(linea, "&eacute;") <> 0 Then linea = f_sustituir_subcadena(linea, "&eacute;", "é")
                        If InStr(linea, "&iacute;") <> 0 Then linea = f_sustituir_subcadena(linea, "&iacute;", "í")
                        If InStr(linea, "&oacute;") <> 0 Then linea = f_sustituir_subcadena(linea, "&oacute;", "ó")
                        If InStr(linea, "&uacute;") <> 0 Then linea = f_sustituir_subcadena(linea, "&uacute;", "ú")
                        linea = f_quitar_caracteres(linea, "(),;.""'¡!|{}¿?ºª<>$%&*+-/\_~·#^-=:[]", " ")
                        'Ojo no quitar los caracteres @, http, www, (numeros), ftp que van por f_es_palabra_normal
                        palabra = ""
                        While Len(linea) > 0
                            c = Left(linea, 1)
                            linea = Right(linea, Len(linea) - 1)
                            If c = " " Then
                                If f_es_palabra_normal(palabra) Then
                                    Print #CTE_FIC_13_DICC_SAL, palabra
                                End If
                                palabra = ""
                            Else
                                palabra = palabra & c
                            End If
                        Wend
                        If palabra <> "" Then
                            If f_es_palabra_normal(palabra) Then
                                Print #CTE_FIC_13_DICC_SAL, palabra
                            End If
                            palabra = ""
                        End If
                    End If
                End If
            Wend
            Close #CTE_FIC_12_DICC_ENT
        Else
            Op_Estado.Text = "ERROR: No se ha podido abrir el fichero " & lista_ficheros(cont_fic) & vbCrLf
            Op_Estado2.Text = Op_Estado2.Text & "ERROR: No se ha podido abrir el fichero " & lista_ficheros(cont_fic) & vbCrLf
        End If
    Next cont_fic
    
    Op_Estado.Text = "--" & vbCrLf
    Op_Estado2.Text = Op_Estado2.Text & "--" & vbCrLf
    
    Op_Estado.Text = "Se ha leido " & cont_fic_leidos & " ficheros." & vbCrLf
    Op_Estado2.Text = Op_Estado2.Text & "Se ha leido " & cont_fic_leidos & " ficheros." & vbCrLf
    
    'Cierro el fichero de salida
    Close #CTE_FIC_13_DICC_SAL
    
    'Lo ordeno y quito repetidos
    f_ordenar_y_quitar_repetidos_gran_fichero_texto (fic_destino)
    
    Op_Estado.Text = "Ejecución finalizada." & vbCrLf
    Op_Estado2.Text = Op_Estado2.Text & "Ejecución finalizada." & vbCrLf
    
    Screen.MousePointer = CTE_DEFECTO
    Op_Estado.Visible = False
    Op_Estado2.Visible = True
    
    MsgBox "El fichero de diccionario ha sido generado.", vbInformation
    
    
End Sub

Private Sub Salir_Click()
    Unload Me

End Sub
