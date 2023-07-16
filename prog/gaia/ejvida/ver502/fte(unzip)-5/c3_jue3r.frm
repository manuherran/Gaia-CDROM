VERSION 5.00
Begin VB.Form frm_c3_juego3r 
   Caption         =   "Tablero del Juego"
   ClientHeight    =   6840
   ClientLeft      =   1350
   ClientTop       =   990
   ClientWidth     =   7020
   Icon            =   "C3_JUE3R.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6840
   ScaleWidth      =   7020
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   39
      Text            =   "1"
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton Cerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Nueva 
      Caption         =   "&Nueva Partida"
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   960
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   1920
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   0
         TabIndex        =   7
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   960
         TabIndex        =   6
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   1920
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton B 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Label Label22 
      Caption         =   "Prioridad"
      Height          =   255
      Left            =   5400
      TabIndex        =   61
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label prioridad 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5400
      TabIndex        =   60
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label19 
      Caption         =   "Estas son las reglas de tipo 1 y 2 encontradas en los agentes 1 y 2"
      Height          =   255
      Left            =   360
      TabIndex        =   59
      Top             =   5640
      Width           =   6375
   End
   Begin VB.Label tipo4 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   6600
      Width           =   6615
   End
   Begin VB.Label tipo3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   6360
      Width           =   6615
   End
   Begin VB.Label tipo2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   6120
      Width           =   6615
   End
   Begin VB.Label tipo1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Top             =   5880
      Width           =   6615
   End
   Begin VB.Label r1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3480
      TabIndex        =   54
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label r2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5040
      TabIndex        =   53
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label p1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   52
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label p2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5040
      TabIndex        =   51
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label27 
      Caption         =   "Tipo 2"
      Height          =   255
      Left            =   3960
      TabIndex        =   50
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label26 
      Caption         =   "Tipo 1"
      Height          =   255
      Left            =   3960
      TabIndex        =   49
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label25 
      Caption         =   "Ag. 2"
      Height          =   255
      Left            =   5040
      TabIndex        =   48
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label24 
      Caption         =   "Ag. 1"
      Height          =   255
      Left            =   4440
      TabIndex        =   47
      Top             =   360
      Width           =   495
   End
   Begin VB.Label numtipo2_2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5040
      TabIndex        =   46
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label numtipo1_2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   45
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label numtipo2_1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5040
      TabIndex        =   44
      Top             =   720
      Width           =   495
   End
   Begin VB.Label numtipo1_1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   43
      Top             =   720
      Width           =   495
   End
   Begin VB.Label peso 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   42
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Peso"
      Height          =   255
      Left            =   4800
      TabIndex        =   41
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Jugar contra el nº"
      Height          =   255
      Left            =   840
      TabIndex        =   40
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*"
      Height          =   255
      Index           =   14
      Left            =   3480
      TabIndex        =   38
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "Indiferente"
      Height          =   255
      Left            =   3840
      TabIndex        =   37
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Ficha del Contrario"
      Height          =   255
      Left            =   3840
      TabIndex        =   36
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      Height          =   255
      Index           =   13
      Left            =   3480
      TabIndex        =   35
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "Ficha Propia"
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P"
      Height          =   255
      Index           =   12
      Left            =   3480
      TabIndex        =   33
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Casilla Vacía"
      Height          =   255
      Left            =   3840
      TabIndex        =   32
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V"
      Height          =   255
      Index           =   11
      Left            =   3480
      TabIndex        =   31
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Juego al Azar"
      Height          =   255
      Left            =   3840
      TabIndex        =   30
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
      Height          =   255
      Index           =   10
      Left            =   3480
      TabIndex        =   29
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Leyenda:"
      Height          =   255
      Left            =   3480
      TabIndex        =   28
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Última regla utilizada:"
      Height          =   255
      Left            =   3480
      TabIndex        =   27
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   3960
      TabIndex        =   26
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   25
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   24
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   23
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   22
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   21
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   20
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   19
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   18
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   17
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label empieza 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label mensaje 
      Caption         =   "Estado del juego: jugando"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Tú juegas con X. La máquina juega con O."
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label turno 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frm_c3_juego3r"
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

Private Sub b_Click(Index As Integer)

'Al hacer click en una casilla

If estado_3r = CTE_JUGANDO And ganador_3r = "N" Then

    Dim N As Integer
    
    If (Not (estado_del_tablero_3r(Index) = "V")) Then
        Beep
        frm_c3_juego3r.mensaje.ForeColor = &HFF& 'rojo
        frm_c3_juego3r.mensaje.Caption = "Esa casilla está ocupada. Partida Perdida. El ganador es 0"
        ganador_3r = "O"
    Else
        s_colocar_ficha_3r (Index)
        ganador_3r = f_ganador_3r()
        If ganador_3r = "X" Or ganador_3r = "O" Then
            frm_c3_juego3r.mensaje.ForeColor = &HFF& 'rojo
            frm_c3_juego3r.mensaje.Caption = "El ganador es " & ganador_3r
        Else
            If ganador_3r = "T" Then
                frm_c3_juego3r.mensaje.ForeColor = &HFF& 'rojo
                frm_c3_juego3r.mensaje.Caption = "Tablas. No hay ganador"
            Else
                If ganador_3r = "N" Then
                    turno_3r = "O"
                    frm_c3_juego3r.turno.Caption = "TURNO: " & turno_3r
                    s_es_el_turno_de_O
                End If
            End If
        End If
    End If
    
End If

End Sub

Private Sub Cerrar_Click()
   
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True 'permito recibir teclas
    s_inicializar_juego_3r

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    If estado_3r = CTE_FUNCIONANDO Then
        If ver_agentes_3r = True Then
            If MsgBox("¿Desea continuar la ejecución sin ver las partidas y sin ver los mejores agentes?", vbYesNo + vbQuestion) = vbYes Then
                ver_agentes_3r = False
                frm_c0_ce.fr_Todas.Visible = False
                frm_c0_ce.Fr_Ejecucion.Visible = False
            Else
                Cancel = True
            End If
        End If
    Else
        If estado_3r = CTE_JUGANDO Then
            estado_3r = CTE_DETENIDO
            s_mostrar_estado_semaforo frm_c3_in3r, estado_3r
        End If
    End If
    
End Sub

Private Sub Nueva_Click()
    Form_Load
End Sub

