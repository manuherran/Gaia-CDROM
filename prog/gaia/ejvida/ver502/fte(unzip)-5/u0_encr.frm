VERSION 5.00
Begin VB.Form frm_u0_encr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encriptar/Desencriptar"
   ClientHeight    =   1935
   ClientLeft      =   1110
   ClientTop       =   3240
   ClientWidth     =   10095
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1935
   ScaleWidth      =   10095
   Begin VB.TextBox Op_Clave 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4680
      TabIndex        =   13
      Text            =   "253"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Fic_Enc 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Op_CarpetaEnc 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2040
      TabIndex        =   7
      Text            =   "c:\......"
      Top             =   1200
      Width           =   4695
   End
   Begin VB.TextBox Op_FicEnc 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6840
      TabIndex        =   6
      Text            =   "enc.txt"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Op_FicDes 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6840
      TabIndex        =   5
      Text            =   "desenc.txt"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Op_CarpetaDes 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2040
      TabIndex        =   4
      Text            =   "c:\......"
      Top             =   840
      Width           =   4695
   End
   Begin VB.CommandButton Fic_Des 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton boton_desencriptar 
      Caption         =   "&Desencriptar"
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
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton boton_encriptar 
      Caption         =   "&Encriptar"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1095
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label TiempoEstimado 
      AutoSize        =   -1  'True
      Caption         =   "0 seg."
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
      Left            =   8520
      TabIndex        =   16
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Porcentaje 
      AutoSize        =   -1  'True
      Caption         =   "0%"
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
      Left            =   8520
      TabIndex        =   15
      Top             =   120
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Clave (1-255)"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Encriptado"
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
      Left            =   8280
      TabIndex        =   12
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desencriptado"
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
      Left            =   8280
      TabIndex        =   11
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Carpeta"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fichero"
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
      Left            =   6840
      TabIndex        =   9
      Top             =   480
      Width           =   525
   End
End
Attribute VB_Name = "frm_u0_encr"
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


Private Sub boton_desencriptar_Click()
    
    Dim Fic_Des As String
    Dim Fic_Enc As String
    Dim clave As String

    Fic_Des = f_nombre_completo(frm_u0_encr.Op_CarpetaDes, frm_u0_encr.Op_FicDes)
    Fic_Enc = f_nombre_completo(frm_u0_encr.Op_CarpetaEnc, frm_u0_encr.Op_FicEnc)
    clave = Op_Clave.Text

    'Compruebo que existe el origen
    If Not f_existe_fichero(Fic_Enc) Then
        MsgBox "El fichero encriptado especificado """ & Fic_Enc & """ no existe. No es posible desencriptarlo.", vbInformation
    Else
        'Compruebo que no existe el destino
        If f_existe_fichero(Fic_Des) Then
            If MsgBox("El fichero desencriptado """ & Fic_Des & """ ya existe. ¿Desea reemplazarlo?", vbQuestion + vbOKCancel) = vbOK Then
                'Ha elegido reemplazar el fichero
                s_desencriptar Fic_Enc, Fic_Des, clave
            End If
        Else
            s_desencriptar Fic_Enc, Fic_Des, clave
        End If
    End If

End Sub

Private Sub boton_encriptar_Click()

    Dim Fic_Des As String
    Dim Fic_Enc As String
    Dim clave As String

    Fic_Des = f_nombre_completo(frm_u0_encr.Op_CarpetaDes, frm_u0_encr.Op_FicDes)
    Fic_Enc = f_nombre_completo(frm_u0_encr.Op_CarpetaEnc, frm_u0_encr.Op_FicEnc)
    clave = Op_Clave.Text

    'Compruebo que existe el origen
    If Not f_existe_fichero(Fic_Des) Then
        MsgBox "El fichero desencriptado especificado """ & Fic_Des & """ no existe. No es posible encriptarlo.", vbInformation
    Else
        'Compruebo que no existe el destino
        If f_existe_fichero(Fic_Enc) Then
            If MsgBox("El fichero encriptado """ & Fic_Enc & """ ya existe. ¿Desea reemplazarlo?", vbQuestion + vbOKCancel) = vbOK Then
                'Ha elegido reemplazar el fichero
                s_encriptar Fic_Des, Fic_Enc, clave
            End If
        Else
            s_encriptar Fic_Des, Fic_Enc, clave
        End If
    End If
    
End Sub

Private Sub Fic_Des_Click()

    'Fijo una carpeta por defecto
    If Len(Trim(Op_CarpetaDes.Text)) > 0 Then
        nombre_fichero_ejv = Trim(Op_CarpetaDes.Text)
    Else
        nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_UTIL)
    End If
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta o carpeta mas fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.*"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Op_CarpetaDes.Text = nombre_fichero_ejv
    Else
        Op_CarpetaDes.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_PRG_UTIL)
    End If

    'Si ha tecleado un nombre de fichero, lo muestro en su sitio
    If Len(f_nombre_fichero(nombre_fichero_ejv)) > 0 Then
        Op_FicDes = f_nombre_fichero(nombre_fichero_ejv)
    End If

End Sub

Private Sub Fic_Enc_Click()

    'Fijo una carpeta por defecto
    If Len(Trim(Op_CarpetaEnc.Text)) > 0 Then
        nombre_fichero_ejv = Trim(Op_CarpetaEnc.Text)
    Else
        nombre_fichero_ejv = path_largo_ejv(CTE_C_PRG_UTIL)
    End If
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta o carpeta mas fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.*"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Op_CarpetaEnc.Text = nombre_fichero_ejv
    Else
        Op_CarpetaEnc.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_PRG_UTIL)
    End If

    'Si ha tecleado un nombre de fichero, lo muestro en su sitio
    If Len(f_nombre_fichero(nombre_fichero_ejv)) > 0 Then
        Op_FicEnc = f_nombre_fichero(nombre_fichero_ejv)
    End If

End Sub

Private Sub Form_Load()

    Op_CarpetaDes.Text = path_largo_ejv(CTE_C_PRG_UTIL)
    Op_CarpetaEnc.Text = path_largo_ejv(CTE_C_PRG_UTIL)

End Sub

Private Sub Text1_Change()

End Sub
