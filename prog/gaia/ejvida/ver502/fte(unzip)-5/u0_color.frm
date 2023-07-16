VERSION 5.00
Begin VB.Form frm_u0_color 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4995
   ClientLeft      =   2250
   ClientTop       =   1785
   ClientWidth     =   4995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4995
   ScaleWidth      =   4995
   Begin VB.ComboBox Cb_Color 
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
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox escala 
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
      Left            =   1320
      TabIndex        =   15
      Text            =   "8"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox t_hex 
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
      Left            =   2880
      TabIndex        =   12
      Text            =   "#FF0000"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox t_rgb 
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
      Left            =   2880
      TabIndex        =   11
      Text            =   "255"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.VScrollBar V_Azul 
      Height          =   3255
      Left            =   1800
      TabIndex        =   9
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox T_Azul 
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
      Left            =   1680
      TabIndex        =   8
      Text            =   "255"
      Top             =   3840
      Width           =   495
   End
   Begin VB.VScrollBar V_Verde 
      Height          =   3255
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox T_Verde 
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
      Left            =   1080
      TabIndex        =   5
      Text            =   "255"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox T_Rojo 
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
      Left            =   480
      TabIndex        =   3
      Text            =   "255"
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   3000
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.VScrollBar V_Rojo 
      Height          =   3255
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Aceptar 
      Cancel          =   -1  'True
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
      Left            =   3720
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "B"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   1800
      TabIndex        =   22
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "G"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   1200
      TabIndex        =   21
      Top             =   4200
      Width           =   150
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "R"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   600
      TabIndex        =   20
      Top             =   4200
      Width           =   150
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "255"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Escala"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hex"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      Top             =   2760
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "VB"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2400
      TabIndex        =   13
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Azul"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Verde"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Rojo"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frm_u0_color"
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

Dim habilitar_change As Boolean

Private Sub Aceptar_Click()
    
    Unload Me

End Sub


Private Sub Cb_Color_Click()
    
    Dim nuevo As Long
        
    nuevo = cct_ejv(Cb_Color.ListIndex + 1)
    Picture1.BackColor = nuevo
    t_rgb.Text = nuevo

End Sub

Private Sub t_rgb_Change()

    Dim valor_hex As String
    Dim valor_long As Long
    
    Dim dec_R As Integer
    Dim dec_G As Integer
    Dim dec_B As Integer
    
    Dim hex_R As String
    Dim hex_G As String
    Dim hex_B As String

If habilitar_change Then

    'Si el valor es correcto, saco el color
    'si no es correcto, lo convierto a uno correcto
    If Len(t_rgb.Text) <= 8 And Len(t_rgb.Text) > 0 And IsNumeric(t_rgb.Text) Then
        valor_long = t_rgb.Text
        If valor_long <= 16777215 Then
            Picture1.BackColor = t_rgb.Text
        Else
            t_rgb.Text = "16777215"
            valor_long = 16777215
            Picture1.BackColor = t_rgb.Text
        End If
    Else
        t_rgb.Text = "0"
        valor_long = 0
        Picture1.BackColor = t_rgb.Text
    End If


    'Calculo los RGB
    s_colorVB2colorRGB valor_long, dec_R, dec_G, dec_B

    habilitar_change = False
    T_Rojo = dec_R
    T_Verde = dec_G
    T_Azul = dec_B
    habilitar_change = True
    
    hex_R = f_ceros_izquierda(CStr(Hex(dec_R)), 2)
    hex_G = f_ceros_izquierda(CStr(Hex(dec_G)), 2)
    hex_B = f_ceros_izquierda(CStr(Hex(dec_B)), 2)

    t_hex.Text = "#" & hex_R & hex_G & hex_B

End If
    
End Sub

Private Sub t_hex_Change()
    
    Dim cadena As String
    Dim valor_long As Long
    
    Dim hex_R As String
    Dim hex_G As String
    Dim hex_B As String

    Dim dec_R As Integer
    Dim dec_G As Integer
    Dim dec_B As Integer

    cadena = t_hex.Text
    If Left(cadena, 1) = "#" Then
        cadena = Right(cadena, Len(cadena) - 1)
    End If
    'cadena = f_ceros_izquierda(cadena, 6)

    hex_R = Left(f_ceros_izquierda(cadena, 6), 2)
    hex_G = Mid(f_ceros_izquierda(cadena, 6), 3, 2)
    hex_B = Right(f_ceros_izquierda(cadena, 6), 2)

    dec_R = Val("&H" & hex_R)
    dec_G = Val("&H" & hex_G)
    dec_B = Val("&H" & hex_B)
    
    valor_long = RGB(dec_R, dec_G, dec_B)
    
    habilitar_change = False
    T_Rojo = dec_R
    T_Verde = dec_G
    T_Azul = dec_B
    t_rgb.Text = valor_long
    habilitar_change = True
            
    Picture1.BackColor = valor_long


End Sub


Private Sub escala_Change()

    If Len(escala.Text) <= 3 Then
        V_Rojo.LargeChange = escala.Text
        V_Verde.LargeChange = escala.Text
        V_Azul.LargeChange = escala.Text
    End If
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    
    'Combo
    For i = 1 To nct_i_ejv
        Cb_Color.AddItem nct_ejv(i)
    Next i
    frm_z0_op.Cb_Color.ListIndex = 0
    
    
    habilitar_change = True
    escala.Text = 64
    
    V_Rojo.max = 255
    V_Rojo.min = 0
    V_Rojo.Value = 255
    V_Rojo.LargeChange = escala.Text
    
    V_Verde.max = 255
    V_Verde.min = 0
    V_Verde.Value = 255
    V_Verde.LargeChange = escala.Text
    
    V_Azul.max = 255
    V_Azul.min = 0
    V_Azul.Value = 255
    V_Azul.LargeChange = escala.Text
    
    s_centrar_ventana_ejv Me
    cambiar_color

End Sub


Private Sub T_Azul_Change()
    
    If habilitar_change Then
        cambiar_color
    End If

    habilitar_change = False
    V_Azul.Value = T_Azul.Text
    habilitar_change = True

End Sub


Private Sub T_Rojo_Change()
    
    If habilitar_change Then
        cambiar_color
    End If

    habilitar_change = False
    V_Rojo.Value = T_Rojo.Text
    habilitar_change = True
    

End Sub

Private Sub T_Verde_Change()
    If habilitar_change Then
        cambiar_color
    End If
    
    habilitar_change = False
    V_Verde.Value = T_Verde.Text
    habilitar_change = True

End Sub

Private Sub V_Azul_Change()
    T_Azul = V_Azul.Value

End Sub

Private Sub V_Rojo_Change()
    
    T_Rojo = V_Rojo.Value

End Sub

Private Sub V_Verde_Change()
    T_Verde = V_Verde.Value

End Sub


Sub cambiar_color()
    habilitar_change = False
    t_rgb.Text = RGB(T_Rojo, T_Verde, T_Azul)
    habilitar_change = True
    t_hex.Text = "#" & f_ceros_izquierda(Hex(CInt(T_Rojo.Text)), 2) & f_ceros_izquierda(Hex(CInt(T_Verde.Text)), 2) & f_ceros_izquierda(Hex(CInt(T_Azul.Text)), 2)
    Picture1.BackColor = RGB(T_Rojo, T_Verde, T_Azul)
    
End Sub
