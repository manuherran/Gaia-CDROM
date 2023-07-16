VERSION 5.00
Begin VB.Form frm_z0_aut 
   Caption         =   "Generador de *.aut"
   ClientHeight    =   7680
   ClientLeft      =   1155
   ClientTop       =   975
   ClientWidth     =   9795
   Icon            =   "z0_aut.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7680
   ScaleWidth      =   9795
   Begin VB.TextBox lineas_config_ejv 
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
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5280
      Width           =   9255
   End
   Begin VB.TextBox min_digitos 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8880
      TabIndex        =   14
      Text            =   "4"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox parametros 
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
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3240
      Width           =   9255
   End
   Begin VB.TextBox PathGeneracion 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Text            =   "c:\......"
      Top             =   240
      Width           =   6735
   End
   Begin VB.TextBox indice 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8040
      TabIndex        =   9
      Text            =   "1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox nombre_fic 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "prueba_"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txt_cabecera 
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
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9255
   End
   Begin VB.CommandButton elegir_carpeta 
      Caption         =   "..."
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Generar 
      Caption         =   "&Generar"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Cerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Líneas a incluir en el fichero config.ejv"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Número mínimo de dígitos en las secuencias numericas"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "Parámetros a combinar, separados por el símbolo ""="""
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   ".aut"
      Height          =   255
      Left            =   8760
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Directorio de generación"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Nombres de los ficheros generados"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "concatenado con un indice que empieza en"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Cabecera Fija"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frm_z0_aut"
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
Private Sub Form_Load()
    
    Dim texto As String
    Dim cont As Long

    PathGeneracion = path_largo_ejv(CTE_C_PRG_AUT)
    
    texto = ""
    texto = texto & "'Parametros opcionales comunes de cualquier ejemplo automático" & vbCrLf
    texto = texto & "'=============================================================" & vbCrLf
    texto = texto & "AUTOMATICO NUMERO PROGRAMA=2" & vbCrLf
    texto = texto & "AUTOMATICO NUMERO EJEMPLO=4" & vbCrLf
    texto = texto & "'FICHERO RESULTADOS=RESULTADOS4.XLS" & vbCrLf
    texto = texto & "ITERACIONES=20" & vbCrLf
    texto = texto & "" & vbCrLf
    texto = texto & "'Parametros específicos del ejemplo automático" & vbCrLf
    texto = texto & "'=============================================" & vbCrLf
    texto = texto & "" & vbCrLf
    txt_cabecera = texto
    
    
    texto = ""
    texto = texto & "'TASA MUTACION=A=B" & vbCrLf
    texto = texto & "'NUMERO DE AGENTES=1=2=3" & vbCrLf
    texto = texto & "FRASE A BUSCAR="
    For cont = 1 To 39
        texto = texto & f_repetir_cadena("0 0 0 1", " ", cont) & "="
    Next cont
    texto = texto & f_repetir_cadena("0 0 0 1", " ", 40)
    texto = texto & vbCrLf
    texto = texto & "CRITERIO DE PARADA=0,8=0,9=0,95=0,98=1" & vbCrLf
    texto = texto & "" & vbCrLf
    parametros = texto
    
    texto = ""
    texto = texto & "(No es necesario escribir aquí)" & vbCrLf
    texto = texto & "En esta caja de texto aparecerán después de la generación" & vbCrLf
    texto = texto & "las lineas que se deben incluir al final del fichero " & CTE_nombreINICIO_TXT & vbCrLf
    texto = texto & "para que sea capaz de llamar a todos los *.aut generados" & vbCrLf
    texto = texto & "" & vbCrLf
    lineas_config_ejv = texto
    
    
    
End Sub

Private Sub Generar_Click()

    Dim texto As String
    Dim nombre_completo_fic_actual As String
    Dim linea As String
    Dim i As Integer
    Dim num_lineas As Integer
    Dim cont_fic As Integer
    Dim parametro_actual() As Integer
    Dim numero_de_parametros() As Integer
    Dim lista_parametros() As String
    Dim lista_etiquetas() As String
    Dim numero_total_ficheros_a_generar As Integer
    Dim param() As String
    Dim incrementado As Boolean
    
    'Trato los parametros
    num_lineas = f_multiline2array(parametros.Text, lista_parametros())
    'Ya tengo las lineas en un array
    ReDim numero_de_parametros(1 To num_lineas) As Integer
    ReDim lista_etiquetas(1 To num_lineas) As String
    ReDim param(1 To num_lineas) As String
    ReDim parametro_actual(1 To num_lineas) As Integer
    For i = 1 To num_lineas
        numero_de_parametros(i) = f_ocurrencias_cadena(lista_parametros(i), "=")
        lista_etiquetas(i) = Left(lista_parametros(i), InStr(lista_parametros(i), "="))
        parametro_actual(i) = 1
    Next i
    numero_total_ficheros_a_generar = 1
    For i = 1 To num_lineas
        numero_total_ficheros_a_generar = numero_total_ficheros_a_generar * numero_de_parametros(i)
    Next i

    If Len(numero_total_ficheros_a_generar) > CInt(min_digitos) Then
        MsgBox "El número mínimo de dígitos (" & min_digitos & ") es menor que el número de digitos (" & Len(numero_total_ficheros_a_generar) & ") del número de ficheros a generar (" & numero_total_ficheros_a_generar & "). Se recomienda cancelar la generación y aumentar el valor del parámetro número mínimo de dígitos al valor " & Len(numero_total_ficheros_a_generar) & " o superior.", vbInformation
    End If

    If MsgBox("Esto generará los " & numero_total_ficheros_a_generar & " ficheros en el path " & PathGeneracion & ", reemplazando los ficheros ya existentes si hubiese alguno con el mismo nombre. ¿Está seguro de querer generar los ficheros *.aut?", vbQuestion + vbOKCancel) = vbOK Then
        texto = ""
        cont_fic = 0
        Screen.MousePointer = CTE_ARENA
        While cont_fic < numero_total_ficheros_a_generar
            cont_fic = cont_fic + 1
            'Compongo los nombres
            nombre_completo_fic_actual = f_nombre_completo(PathGeneracion, nombre_fic & f_ceros_izquierda(CStr(cont_fic + CLng(indice.Text) - 1), CInt(min_digitos)) & ".aut")
            texto = texto & "FICHERO AUTOMATICO=" & nombre_completo_fic_actual & vbCrLf
            'Elijo parametros
            If f_tomar_parametros(param(), lista_parametros(), parametro_actual(), lista_etiquetas(), num_lineas) Then
                'Genero los ficheros
                Open nombre_completo_fic_actual For Output As #CTE_FIC_02_AUT
                linea = txt_cabecera
                Print #CTE_FIC_02_AUT, linea
                For i = 1 To num_lineas
                    linea = param(i)
                    Print #CTE_FIC_02_AUT, linea
                Next i
                Close #CTE_FIC_02_AUT
                'Paso al siguiente valor de parametro
                incrementado = False
                i = 1
                While Not incrementado
                    If parametro_actual(i) < numero_de_parametros(i) Then
                        'Incremento este
                        parametro_actual(i) = parametro_actual(i) + 1
                        incrementado = True
                    Else
                        'Ya he tratado el ultimo el ultimo, asi que compruebo otro
                        parametro_actual(i) = 1
                        i = i + 1
                        If i > num_lineas Then
                            'Ya he llegado al ultimo valor de todos
                            incrementado = True
                        End If
                    End If
                Wend
            End If
        Wend
        lineas_config_ejv = texto
        Screen.MousePointer = CTE_DEFECTO
        MsgBox "Los ficheros han sido generados.", vbInformation
    Else
        MsgBox "Operación cancelada.", vbInformation
    End If

    

End Sub

Private Sub elegir_carpeta_Click()

    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_RAIZ)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.*"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        frm_z0_aut.PathGeneracion.Text = nombre_fichero_ejv
    Else
        frm_z0_aut.PathGeneracion.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_RAIZ)
    End If

End Sub



Function f_tomar_parametros(param() As String, lista_parametros() As String, parametro_actual() As Integer, lista_etiquetas() As String, num_lineas As Integer) As Boolean

    'Devuelvo una linea normal en cada elemento de param()

    Dim pos_inicial_param As Integer
    Dim long_param As Integer
    Dim cont As Integer
    Dim i As Integer
    Dim tmp As String

    For i = 1 To num_lineas
        tmp = lista_parametros(i)
        cont = 0
        pos_inicial_param = 0
        'Salto varios =
        While cont < parametro_actual(i)
            cont = cont + 1
            pos_inicial_param = pos_inicial_param + InStr(tmp, "=")
            tmp = Right(tmp, Len(tmp) - InStr(tmp, "="))
        Wend
        'Cojo hasta el final o hasta el siguiente =
        If InStr(tmp, "=") = 0 Then
            'Es el ultimo valor
            param(i) = lista_etiquetas(i) & Right(lista_parametros(i), Len(lista_parametros(i)) - pos_inicial_param)
        Else
            'no es el ultimo valor
            long_param = InStr(tmp, "=") - 1
            param(i) = Right(lista_parametros(i), Len(lista_parametros(i)) - pos_inicial_param)
            param(i) = lista_etiquetas(i) & Left(param(i), long_param)
        End If
    Next i

    f_tomar_parametros = True


End Function


Private Sub Cerrar_Click()
    Unload Me

End Sub


