VERSION 5.00
Begin VB.Form frm_a1_ophyp 
   Caption         =   "Opciones de Hormigas y Plantas"
   ClientHeight    =   6585
   ClientLeft      =   1155
   ClientTop       =   975
   ClientWidth     =   9795
   Icon            =   "A1_OPHYP.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6585
   ScaleWidth      =   9795
   Begin VB.Frame Frame1 
      Caption         =   "Reproducción sexual"
      Height          =   1215
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton Op_nHermafroditas 
         Caption         =   "Dos sexos. Hembras y machos"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Op_Hermafroditas 
         Caption         =   "Agentes hermafroditas. Cualquiera"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "puede reproducirse con cualquiera"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.TextBox PosicionesReproducirse 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   14
      Text            =   "2"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox PosicionesPelear 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   13
      Text            =   "2"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox PosicionesRegar 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   12
      Text            =   "2"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox PlantasPorCiclo 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   11
      Text            =   "2"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox EnergiaAlRegar 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   10
      Text            =   "1"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox VecesRegar 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   9
      Text            =   "3"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox EnergiaInicialAgente 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   8
      Text            =   "10"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox EnergiaAlPelear 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   7
      Text            =   "2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox EnergiaAlReproducir 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   6
      Text            =   "2"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox EnergiaAlMover 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   5
      Text            =   "2"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox EnergiaAlComer 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   4
      Text            =   "15"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox numIniciPlantas 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   3
      Text            =   "5"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox CantInicialAgua 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8760
      TabIndex        =   2
      Text            =   "10"
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label numInicHorm 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8760
      TabIndex        =   29
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de posiciones que se mueve la hormiga después de reproducirse"
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   4200
      Width           =   5295
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de posiciones que se mueve la hormiga después de pelear"
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de posiciones que se mueve la hormiga después de regar"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de plantas que nacen por ciclo"
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía consumida por una hormiga al regar una vez una planta"
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de veces que se ha de regar una planta para que pueda ser comida"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   4920
      Width           =   5415
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía inicial que posee cada hormiga"
      Height          =   255
      Left            =   5640
      TabIndex        =   22
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía consumida por una hormiga al pelearse"
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía consumida por una hormiga al reproducirse"
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía consumida por una hormiga al moverse una posición"
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Energía proporcionada a una hormiga cada vez que come una planta"
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Número inicial de hormigas"
      Height          =   255
      Left            =   6720
      TabIndex        =   17
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número inicial de plantas"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad inicial de agua que posee cada planta"
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   4560
      Width           =   3495
   End
End
Attribute VB_Name = "frm_a1_ophyp"
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
    
    Dim i As Integer
    Dim suma As Integer
   
    s_grabar_opciones_hyp
    
    suma = 0
    For i = 1 To num_tipos_agentes_va0
        suma = suma + numero_agentes_que_se_deben_crear_inicio_de_tipo_va0(i)
    Next i
    
    num_inic_horm_hyp = suma
    'pongo como viejas las actuales modificadas
    copia_dim_va0_2_viejo_va0
    
    Unload Me

End Sub

Private Sub Cancelar_Click()
    
    Unload Me

End Sub

Private Sub Mapa_Click()
    s_ver_mapa_va0
End Sub

Private Sub Form_Load()
    
    s_cargar_opciones_hyp

End Sub
