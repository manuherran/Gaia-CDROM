VERSION 5.00
Begin VB.Form frm_a0_mov 
   Caption         =   "Tendencias en el movimiento"
   ClientHeight    =   6225
   ClientLeft      =   2415
   ClientTop       =   1200
   ClientWidth     =   7350
   FillStyle       =   0  'Solid
   Icon            =   "a0_mov.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   7350
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   195
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "a0_mov.frx":0442
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   195
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   38
      Text            =   "a0_mov.frx":044A
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   435
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   36
      Text            =   "a0_mov.frx":0455
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   435
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "a0_mov.frx":0460
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   435
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "a0_mov.frx":0469
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   435
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "a0_mov.frx":0474
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   435
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   32
      Text            =   "a0_mov.frx":047D
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   435
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "a0_mov.frx":0488
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   435
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "a0_mov.frx":0491
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   435
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "a0_mov.frx":049C
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox SE 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   26
      Text            =   "1"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox NE 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   25
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox SO 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      TabIndex        =   24
      Text            =   "1"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox NO 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      TabIndex        =   23
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox O 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      TabIndex        =   22
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox S 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   21
      Text            =   "1"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox N 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   20
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox E 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   16
      Left            =   3360
      TabIndex        =   37
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   11
      Left            =   2640
      TabIndex        =   15
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   10
      Left            =   2640
      TabIndex        =   14
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   13
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   4
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   6
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   9
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   8
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   7
      Top             =   3600
      Width           =   255
   End
   Begin VB.Line Line6 
      X1              =   1440
      X2              =   5520
      Y1              =   5280
      Y2              =   1680
   End
   Begin VB.Line Line5 
      X1              =   1440
      X2              =   5520
      Y1              =   1680
      Y2              =   5280
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   15
      Left            =   2760
      TabIndex        =   19
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   18
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   13
      Left            =   3360
      TabIndex        =   17
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   12
      Left            =   3360
      TabIndex        =   16
      Top             =   2760
      Width           =   255
   End
   Begin VB.Line Line10 
      X1              =   1440
      X2              =   5520
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line9 
      X1              =   5520
      X2              =   5520
      Y1              =   1680
      Y2              =   5280
   End
   Begin VB.Line Line8 
      X1              =   1440
      X2              =   5520
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line7 
      X1              =   1440
      X2              =   1440
      Y1              =   1680
      Y2              =   5280
   End
   Begin VB.Line Line4 
      X1              =   3480
      X2              =   3480
      Y1              =   1680
      Y2              =   5280
   End
   Begin VB.Line Line3 
      X1              =   1440
      X2              =   5520
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label2 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   28
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   1080
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   1080
      X2              =   1080
      Y1              =   1080
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   720
      X2              =   5520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label mensaje 
      Caption         =   $"a0_mov.frx":04A5
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frm_a0_mov"
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
    
    lista_tendencias_en_modificacion_va0(1) = CLng(N.Text)
    lista_tendencias_en_modificacion_va0(2) = CLng(NE)
    lista_tendencias_en_modificacion_va0(3) = CLng(E)
    lista_tendencias_en_modificacion_va0(4) = CLng(SE)
    lista_tendencias_en_modificacion_va0(5) = CLng(S)
    lista_tendencias_en_modificacion_va0(6) = CLng(SO)
    lista_tendencias_en_modificacion_va0(7) = CLng(O)
    lista_tendencias_en_modificacion_va0(8) = CLng(NO)
    
    ha_habido_cambio_lista_tendencias_va0 = True
    Unload Me

End Sub

Private Sub Cancelar_Click()
    Unload Me

End Sub

Private Sub Form_Activate()

    Dim CX As Integer
    Dim CY As Integer

    frm_a0_mov.Refresh
    
    'pinto una hormiga
    frm_a0_mov.ScaleMode = vbPixels   ' Set scale to pixels.
    CX = 232
    CY = 233

    frm_a0_mov.FillColor = cct_ejv(tipo_agente_cambiar_tendencias_va0)
    frm_a0_mov.Circle (CX, CY + 2), 3, cct_ejv(CTE_NEGRO)
    frm_a0_mov.Circle (CX, CY - 2), 3, cct_ejv(CTE_NEGRO)
    frm_a0_mov.Circle (CX + 1, CY - 6), 1, cct_ejv(CTE_NEGRO)
    frm_a0_mov.Circle (CX - 1, CY - 6), 1, cct_ejv(CTE_NEGRO)

End Sub

Private Sub Form_Load()
    
    frm_a0_mov.FillStyle = vbFSSolid 'solido
    s_centrar_ventana_ejv Me

    If tipo_tendencia_en_modificacion_va0 = CTE_RELATIVAS Then
        frm_a0_mov.Caption = "Tendencias relativas iniciales del movimiento"
        frm_a0_mov.mensaje = "Valores altos de tendencia indican una mayor probabilidad de ir en esa dirección. Estas direcciones son relativas, esto es, una tendencia hacia el norte es en realidad una tendencia hacia mantener la dirección actual. La suma de todos los valores no tiene por que corresponder con ningún número en especial."
    Else
        frm_a0_mov.Caption = "Tendencias absolutas iniciales del movimiento"
        frm_a0_mov.mensaje = "Valores altos de tendencia indican una mayor probabilidad de ir en esa dirección. Estas direcciones son absolutas: norte es arriba, etc. La suma de todos los valores no tiene por que corresponder con ningún número en especial."
    End If
    
    N = lista_tendencias_en_modificacion_va0(1)
    NE = lista_tendencias_en_modificacion_va0(2)
    E = lista_tendencias_en_modificacion_va0(3)
    SE = lista_tendencias_en_modificacion_va0(4)
    S = lista_tendencias_en_modificacion_va0(5)
    SO = lista_tendencias_en_modificacion_va0(6)
    O = lista_tendencias_en_modificacion_va0(7)
    NO = lista_tendencias_en_modificacion_va0(8)


End Sub
