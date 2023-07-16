VERSION 5.00
Begin VB.Form frm_z0_graf 
   Caption         =   "Gráficos"
   ClientHeight    =   7320
   ClientLeft      =   1035
   ClientTop       =   1425
   ClientWidth     =   11370
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
   Icon            =   "z0_graf.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7320
   ScaleWidth      =   11370
   Begin VB.CommandButton Borrar 
      Caption         =   "Borrar"
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
      Left            =   1800
      TabIndex        =   55
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Fr_Opciones 
      Caption         =   "Opciones"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   7215
      Begin VB.ComboBox Cb_num_ejes 
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
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton VerPropiedadesEje 
         Caption         =   "Ver Propiedades"
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
         Left            =   5040
         TabIndex        =   14
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox Cb_Ejes 
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
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Número de ejes a mostrar"
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
         Left            =   600
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Modificar Opciones de"
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
         Left            =   720
         TabIndex        =   27
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Fr_Grafico 
      Caption         =   "Datos mostrados en el eje..."
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   7215
      Begin VB.TextBox txt_leyenda 
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
         Left            =   1200
         TabIndex        =   41
         Top             =   2640
         Width           =   3255
      End
      Begin VB.CommandButton ModificarPropiedadesEje 
         Caption         =   "Modificar"
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
         Left            =   5280
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox marcas_Y 
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
         Left            =   6240
         TabIndex        =   26
         Text            =   "20"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox marcas_X 
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
         Left            =   6240
         TabIndex        =   25
         Text            =   "20"
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox Cb_Escala 
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
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   600
         Width           =   2055
      End
      Begin VB.ListBox lst_Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2880
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox lst_Color 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox lst_Datos_a_mostrar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Leyenda"
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
         Left            =   240
         TabIndex        =   42
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Marcas en eje Y cada"
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
         Left            =   4560
         TabIndex        =   24
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Marcas en eje X cada"
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
         Left            =   4560
         TabIndex        =   23
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo"
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
         Left            =   2880
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label txt1 
         Caption         =   "Color"
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
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Datos a mostrar"
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
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton OtrosGraficos 
      Caption         =   "Otros Gráficos..."
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
      Left            =   5400
      TabIndex        =   4
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Refrescar 
      Caption         =   "Mostrar Gráfico"
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
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Opciones 
      Caption         =   "Opciones de los Ejes..."
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
      Left            =   3360
      TabIndex        =   3
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Salir 
      Cancel          =   -1  'True
      Caption         =   "C&errar"
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
      Left            =   10080
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Imprimir 
      Caption         =   "Imprimir..."
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
      Left            =   7080
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame FrOtros_Graficos 
      Caption         =   "Otros Gráficos"
      Height          =   2055
      Left            =   240
      TabIndex        =   31
      Top             =   840
      Width           =   7215
      Begin VB.TextBox txt_semilla 
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
         Left            =   3000
         TabIndex        =   53
         Text            =   "0,123456789"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox numDatosMostrar 
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
         Left            =   3000
         TabIndex        =   51
         Text            =   "20"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox Cb_eje_mod 
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
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton AniadirEspecia 
         Caption         =   "Añadir"
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
         Left            =   5760
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Cb_Gr_Especial 
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
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Semilla"
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
         Left            =   2160
         TabIndex        =   54
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Número de datos a mostrar"
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
         Left            =   840
         TabIndex        =   52
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "añadir en el eje"
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
         Left            =   3240
         TabIndex        =   49
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Gráfico"
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
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label max_x_eje 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "max_x_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   48
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label max_x_eje 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   10200
      TabIndex        =   47
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label max_x_eje 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   10200
      TabIndex        =   46
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label max_x_eje 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   10200
      TabIndex        =   45
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label max_x_eje 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   10200
      TabIndex        =   44
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label max_x_eje 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   10200
      TabIndex        =   43
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label max_y_eje 
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   40
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label max_y_eje 
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   39
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label max_y_eje 
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   38
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label max_y_eje 
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   37
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label max_y_eje 
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   36
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label max_y_eje 
      BackStyle       =   0  'Transparent
      Caption         =   "max_y_eje1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Et_gr 
      BackStyle       =   0  'Transparent
      Caption         =   "Et_gr6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   12
      Top             =   6240
      Width           =   6375
   End
   Begin VB.Label Et_gr 
      BackStyle       =   0  'Transparent
      Caption         =   "Et_gr5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   5160
      Width           =   6615
   End
   Begin VB.Label Et_gr 
      BackStyle       =   0  'Transparent
      Caption         =   "Et_gr4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   10
      Top             =   4200
      Width           =   6975
   End
   Begin VB.Label Et_gr 
      BackStyle       =   0  'Transparent
      Caption         =   "Et_gr3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   9
      Top             =   3120
      Width           =   6855
   End
   Begin VB.Label Et_gr 
      BackStyle       =   0  'Transparent
      Caption         =   "Et_gr2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Top             =   2040
      Width           =   6375
   End
   Begin VB.Label Et_gr 
      BackStyle       =   0  'Transparent
      Caption         =   "Et_gr1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   6975
   End
End
Attribute VB_Name = "frm_z0_graf"
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

Dim numero_total_ejes_gra As Integer

Dim eje_modificar_opciones_gra As Integer

Dim texto_dato_a_mostrar_gra() As String
Dim color_dato_a_mostrar_gra() As Integer
Dim tipo_pto_dato_a_mostrar_gra() As Integer

'Propiedades de los ejes
Dim marcas_X_cada_gra() As Integer
Dim marcas_Y_cada_gra() As Integer
Dim escala_gra() As Integer
Dim numero_de_graficos_a_mostrar_en_eje_gra() As Integer 'numero de graficos a mostrar en cada eje
Dim leyenda_gra() As String
Dim max_numero_de_graficos_a_mostrar_en_eje_gra As Integer 'el numero de graficos del eje con mas graficos

'Graficos disponibles en este momento
Dim num_graficos_diponibles_gra As Integer 'numero total
Dim grafico_disponible_gra() As String 'lista de todos los graficos disponibles, juntos todos los tipos

'Genericos, los que hay por cada programa en cualquier momento
Dim num_graficos_diponibles_genericos_gra() As Integer
Dim grafico_disponible_generico_gra() As String
'Ejemplo, si hay dos curvas # rojas y # rosas en el eje 1 seria:
'dato_a_mostrar_gra(eje 1, indice 1) = grafico 1 (# h rojas)
'dato_a_mostrar_programa_gra(eje 1, indice 2) = grafico de hyp

'Especiales
'Numero de datos que tengo de ese grafico especial
'Tengo lo que yo quiera, pero es fijo por programa
Dim num_datos_mostrar_especial_gra() As Long

Dim posicion_ejeYanterior As Integer
Dim posicion_ejeXanterior As Integer

Dim semilla As Double

Sub s_modificar_prop_eje_gra(dato_a_modificar As String)

    Dim i As Integer
    
    Select Case dato_a_modificar
        Case "Dato a mostrar"
            frm_z0_sele.Etiqueta_izq.Caption = "Datos Disponibles"
            frm_z0_sele.Etiqueta_der.Caption = "Datos Seleccionados"
            For i = 1 To num_graficos_diponibles_gra
                frm_z0_sele.txt_lista_izq.AddItem grafico_disponible_gra(i)
            Next i
            For i = 1 To lst_Datos_a_mostrar.ListCount
                frm_z0_sele.txt_lista_der.AddItem lst_Datos_a_mostrar.List(i - 1)
            Next i
        Case "Color"
            frm_z0_sele.Etiqueta_izq.Caption = "Colores Disponibles"
            frm_z0_sele.Etiqueta_der.Caption = "Colores Seleccionados"
            For i = 1 To nct_i_ejv
                frm_z0_sele.txt_lista_izq.AddItem nct_ejv(i)
            Next i
            For i = 1 To lst_Color.ListCount
                frm_z0_sele.txt_lista_der.AddItem lst_Color.List(i - 1)
            Next i
        Case "Tipo"
            frm_z0_sele.Etiqueta_izq.Caption = "Tipos Disponibles"
            frm_z0_sele.Etiqueta_der.Caption = "Tipos Seleccionados"
            frm_z0_sele.txt_lista_izq.AddItem "Puntos"
            frm_z0_sele.txt_lista_izq.AddItem "Líneas"
            For i = 1 To lst_Tipo.ListCount
                frm_z0_sele.txt_lista_der.AddItem lst_Tipo.List(i - 1)
            Next i
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select
            
    selector_max_der_sel = lst_Datos_a_mostrar.ListCount
    frm_z0_sele.Show CTE_MODAL
    frm_z0_sele.Caption = "Selección del parámetro " & dato_a_modificar
    frm_z0_sele.Etiqueta_izq.Caption = ""

    'Aqui ya ha elegido
    If modificar_resultado_selector_sel Then
        Select Case dato_a_modificar
            Case "Dato a mostrar"
                frm_z0_graf.lst_Datos_a_mostrar.Clear
                For i = 1 To UBound(resultado_selector_sel, 1)
                    frm_z0_graf.lst_Datos_a_mostrar.AddItem resultado_selector_sel(i)
                    texto_dato_a_mostrar_gra(eje_modificar_opciones_gra, i) = resultado_selector_sel(i)
                Next i
                
            Case "Color"
                frm_z0_graf.lst_Color.Clear
                For i = 1 To UBound(resultado_selector_sel, 1)
                    frm_z0_graf.lst_Color.AddItem resultado_selector_sel(i)
                    color_dato_a_mostrar_gra(eje_modificar_opciones_gra, i) = s_calcular_color_str2int_gra(resultado_selector_sel(i))
                Next i
            Case "Tipo"
                frm_z0_graf.lst_Tipo.Clear
                For i = 1 To UBound(resultado_selector_sel, 1)
                    frm_z0_graf.lst_Tipo.AddItem resultado_selector_sel(i)
                    If resultado_selector_sel(i) = "Puntos" Then
                        tipo_pto_dato_a_mostrar_gra(eje_modificar_opciones_gra, i) = 1
                    Else
                        tipo_pto_dato_a_mostrar_gra(eje_modificar_opciones_gra, i) = 2
                    End If
                Next i
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
        End Select
        frm_z0_graf.Refresh
    End If


End Sub

Private Sub AniadirEspecia_Click()
    
    Dim num_gra As Integer
    Dim eje_destino As Integer

    num_gra = Cb_Gr_Especial.ListIndex + 1
    eje_destino = Cb_eje_mod.ListIndex + 1
    
    num_datos_mostrar_especial_gra(num_gra) = CLng(numDatosMostrar.Text)
    Select Case num_gra
        Case CTE_ESPECIAL_1_PAL
            numero_palabras_dicc_pal = 2092
            dicc_carpeta_pal = path_largo_ejv(CTE_C_ENT_DIC)
            dicc_fichero_pal = "2092.dic"
            f_aut_leer_diccionario False
            If num_datos_mostrar_especial_gra(num_gra) > numero_palabras_dicc_pal Then
                num_datos_mostrar_especial_gra(num_gra) = numero_palabras_dicc_pal
                MsgBox "El máximo es " & num_datos_mostrar_especial_gra(num_gra), vbInformation
            End If
        Case CTE_ESPECIAL_2_APE
            If num_datos_mostrar_especial_gra(num_gra) > CTE_numero_maximo_apellidos Then
                num_datos_mostrar_especial_gra(num_gra) = CTE_numero_maximo_apellidos
                MsgBox "El máximo es " & num_datos_mostrar_especial_gra(num_gra), vbInformation
            End If
        Case CTE_ESPECIAL_3_RND
            If num_datos_mostrar_especial_gra(num_gra) > CTE_numero_maximo_rnd Then
                num_datos_mostrar_especial_gra(num_gra) = CTE_numero_maximo_rnd
                MsgBox "El máximo es " & num_datos_mostrar_especial_gra(num_gra), vbInformation
            End If
        Case CTE_ESPECIAL_4_PI
            If Not azar_en_memoria_ejv Then
                s_aut_leer_digitos_azar
                azar_en_memoria_ejv = True
            End If
            If num_datos_mostrar_especial_gra(num_gra) > CTE_numero_maximo_pi Then
                num_datos_mostrar_especial_gra(num_gra) = CTE_numero_maximo_pi
                MsgBox "El máximo es " & num_datos_mostrar_especial_gra(num_gra), vbInformation
            End If
        Case CTE_ESPECIAL_5_X2
            If num_datos_mostrar_especial_gra(num_gra) > CTE_numero_maximo_x2 Then
                num_datos_mostrar_especial_gra(num_gra) = CTE_numero_maximo_x2
                MsgBox "El máximo es " & num_datos_mostrar_especial_gra(num_gra), vbInformation
            End If
        Case CTE_ESPECIAL_6_X21
            If num_datos_mostrar_especial_gra(num_gra) > CTE_numero_maximo_x21 Then
                num_datos_mostrar_especial_gra(num_gra) = CTE_numero_maximo_x21
                MsgBox "El máximo es " & num_datos_mostrar_especial_gra(num_gra), vbInformation
            End If
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select
    
    
    'Hay un grafico mas disponible
    num_graficos_diponibles_gra = num_graficos_diponibles_gra + 1
    ReDim Preserve grafico_disponible_gra(1 To num_graficos_diponibles_gra) As String
    grafico_disponible_gra(num_graficos_diponibles_gra) = grafico_disponible_generico_gra(CTE_ESPECIAL, num_gra)
    'Ahora en ese eje hay un grafico mas
    numero_de_graficos_a_mostrar_en_eje_gra(eje_destino) = numero_de_graficos_a_mostrar_en_eje_gra(eje_destino) + 1
    'Actualizo el maximo
    If numero_de_graficos_a_mostrar_en_eje_gra(eje_destino) > max_numero_de_graficos_a_mostrar_en_eje_gra Then
        max_numero_de_graficos_a_mostrar_en_eje_gra = numero_de_graficos_a_mostrar_en_eje_gra(eje_destino)
    End If
    ReDim Preserve texto_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To max_numero_de_graficos_a_mostrar_en_eje_gra) As String
    texto_dato_a_mostrar_gra(eje_destino, numero_de_graficos_a_mostrar_en_eje_gra(eje_destino)) = grafico_disponible_gra(num_graficos_diponibles_gra)
    ReDim Preserve color_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To max_numero_de_graficos_a_mostrar_en_eje_gra) As Integer
    color_dato_a_mostrar_gra(eje_destino, numero_de_graficos_a_mostrar_en_eje_gra(eje_destino)) = f_SumCirc(5, num_gra, 0)
    ReDim Preserve tipo_pto_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To max_numero_de_graficos_a_mostrar_en_eje_gra) As Integer
    tipo_pto_dato_a_mostrar_gra(eje_destino, numero_de_graficos_a_mostrar_en_eje_gra(eje_destino)) = CTE_GRA_LINEA
    'Propiedades del eje
    leyenda_gra(eje_destino) = leyenda_gra(eje_destino) & " " & grafico_disponible_gra(num_graficos_diponibles_gra)



End Sub

Private Sub Borrar_Click()

    frm_z0_graf.Refresh

End Sub

Private Sub Cb_Gr_Especial_Change()

    numDatosMostrar = num_datos_mostrar_especial_gra(Cb_Gr_Especial.ListIndex + 1)

End Sub

Private Sub Cb_Gr_Especial_Click()
    numDatosMostrar = num_datos_mostrar_especial_gra(Cb_Gr_Especial.ListIndex + 1)

End Sub

Private Sub Cb_Gr_Especial_GotFocus()
    numDatosMostrar = num_datos_mostrar_especial_gra(Cb_Gr_Especial.ListIndex + 1)

End Sub

Private Sub Form_GotFocus()
    
    Me.BackColor = cct_ejv(cfondo_ejv)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cancelar_mostrar_grafico_ejv = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    s_cambiar_estado_enabled_operaciones_ficheros_ejv False
    
    If tipo_funcion_azar_ejv = CTE_AZARVB Then
        'Borro los arrays de azar, que no los voy a necesitar ya que uso VB
        ReDim digitos_azar(0 To 0) As Integer
        azar_en_memoria_ejv = False
    End If

End Sub

Private Sub lst_Color_Click()
    
    s_modificar_prop_eje_gra "Color"

End Sub

Private Sub lst_Datos_a_mostrar_Click()
    
    'lst_Datos_a_mostrar.ListIndex + 1
    
    s_modificar_prop_eje_gra "Dato a mostrar"

End Sub

Private Sub Imprimir_Click()

    If MsgBox("¿Desea imprimir el formulario?", vbQuestion + vbYesNo) = vbYes Then
        Me.PrintForm
    Else
        s_ver_grafico_gra CTE_IMPRESORA, Printer
    End If

End Sub

Private Sub lst_Tipo_Click()
    s_modificar_prop_eje_gra "Tipo"


End Sub

Private Sub ModificarPropiedadesEje_Click()
    
    Dim i As Integer
    Dim txt As String
    
    'Modificar eje actual
    If eje_modificar_opciones_gra > 0 Then
        marcas_X_cada_gra(eje_modificar_opciones_gra) = marcas_X
        marcas_Y_cada_gra(eje_modificar_opciones_gra) = marcas_Y
        escala_gra(eje_modificar_opciones_gra) = Cb_Escala.ListIndex + 1
        'Leyenda
        txt = ""
        If txt_leyenda = "" Then
            For i = 1 To numero_de_graficos_a_mostrar_en_eje_gra(eje_modificar_opciones_gra)
                txt = txt & texto_dato_a_mostrar_gra(eje_modificar_opciones_gra, i)
            Next i
        Else
            txt = txt_leyenda
        End If
        leyenda_gra(eje_modificar_opciones_gra) = txt
    'Modificar todos los ejes
    Else
        For i = 1 To numero_total_ejes_gra
            marcas_X_cada_gra(i) = CInt(0 & marcas_X)
            marcas_Y_cada_gra(i) = CInt(0 & marcas_Y)
            escala_gra(i) = Cb_Escala.ListIndex + 1
        Next i
    End If




End Sub

Private Sub Opciones_Click()
    
    frm_z0_graf.Refresh
    s_etiquetas_visibles_gra False
    FrOtros_Graficos.Visible = False
    Fr_Opciones.Visible = True
    Fr_Grafico.Visible = False
    frm_z0_graf.Refresh
    
End Sub

Private Sub OtrosGraficos_Click()
    
    frm_z0_graf.Refresh
    s_etiquetas_visibles_gra False
    Fr_Opciones.Visible = False
    Fr_Grafico.Visible = False
    FrOtros_Graficos.Visible = True
    frm_z0_graf.Refresh
    
End Sub

Private Sub Refrescar_Click()

'On Error Resume Next

cancelar_mostrar_grafico_ejv = False
'Si puedo pulsar pausa es que hay algo ejecutandose y entonces no lo permito
If estado_ejecutar_ejv(CTE_EXE_PAUSA, num_prg_activo_ejv) = False Then

    Fr_Opciones.Visible = False
    Fr_Grafico.Visible = False
    FrOtros_Graficos.Visible = False
    Fr_Grafico.Refresh

    s_etiquetas_visibles_gra False
    
    s_ver_grafico_gra CTE_FORMULARIO, frm_z0_graf
    
    Screen.MousePointer = CTE_DEFECTO
End If

End Sub

Private Sub Salir_Click()
'On Error Resume Next
    Close #CTE_FIC_23R_1EJGRA
    Unload Me
End Sub

Private Sub Form_Load()

    Dim i As Integer
    
    'Refresco automatico, consume muchos recursos
    'frm_z0_graf.AutoRedraw = True
    frm_z0_graf.AutoRedraw = False
    
    's_centrar_ventana_ejv Me
    Me.WindowState = CTE_MAXIMIZED
    
    s_etiquetas_visibles_gra False
    s_vaciar_etiquetas_gra
    
    s_cambiar_estado_enabled_operaciones_ficheros_ejv True
    
    Fr_Opciones.Visible = False
    Fr_Grafico.Visible = False
    FrOtros_Graficos.Visible = False
    
    'Inicializo las propiedades de los graficos especiales
    ReDim num_datos_mostrar_especial_gra(1 To CTE_GRA_ESP_num) As Long
    num_datos_mostrar_especial_gra(CTE_ESPECIAL_1_PAL) = 50 'numero_palabras_dicc_pal
    num_datos_mostrar_especial_gra(CTE_ESPECIAL_2_APE) = 50 'CTE_numero_maximo_apellidos
    num_datos_mostrar_especial_gra(CTE_ESPECIAL_3_RND) = 50 'CTE_numero_maximo_rnd
    num_datos_mostrar_especial_gra(CTE_ESPECIAL_4_PI) = 50 'CTE_numero_maximo_pi
    num_datos_mostrar_especial_gra(CTE_ESPECIAL_5_X2) = 50 'CTE_numero_maximo_x2
    num_datos_mostrar_especial_gra(CTE_ESPECIAL_6_X21) = 50 'CTE_numero_maximo_x21
    
    s_inicializar_graficos_disponibles_genericos_gra
    s_inicializar_ejes_gra
    
    Cb_num_ejes.Clear
    Cb_num_ejes.AddItem CTE_EJE_num_max_ejes
    Cb_num_ejes.ListIndex = 0
    
    Cb_Ejes.Clear
    Cb_eje_mod.Clear
    Cb_Ejes.AddItem "(Todos)"
    For i = 1 To CTE_EJE_num_max_ejes
        Cb_Ejes.AddItem "Eje " & i
        Cb_eje_mod.AddItem "Eje " & i
    Next i
    Cb_Ejes.ListIndex = 1
    Cb_eje_mod.ListIndex = 0
    
    Cb_Escala.Clear
    Cb_Escala.AddItem "Ajustada al máximo"
    Cb_Escala.AddItem "Real en Pixels"
    Cb_Escala.ListIndex = 0
    
    Cb_Gr_Especial.Clear
    For i = 1 To CTE_GRA_ESP_num
        Cb_Gr_Especial.AddItem grafico_disponible_generico_gra(CTE_ESPECIAL, i)
    Next i
    Cb_Gr_Especial.ListIndex = 0
    
    
    
End Sub

Sub s_inicializar_graficos_disponibles_genericos_gra()

    ReDim num_graficos_diponibles_genericos_gra(1 To CTE_PROG_num_total) As Integer
    ReDim grafico_disponible_generico_gra(1 To CTE_PROG_num_total, 1 To CTE_GRA_max_graficos) As String
    
    num_graficos_diponibles_genericos_gra(CTE_HYP) = CTE_GRA_HYP_num
    grafico_disponible_generico_gra(CTE_HYP, 1) = "Nº rojas"
    grafico_disponible_generico_gra(CTE_HYP, 2) = "Nº rosas"
    grafico_disponible_generico_gra(CTE_HYP, 3) = "Nº naranjas"
    grafico_disponible_generico_gra(CTE_HYP, 4) = "Nº amarillas"
    grafico_disponible_generico_gra(CTE_HYP, 5) = "Nº verdes"
    grafico_disponible_generico_gra(CTE_HYP, 6) = "Nº plantas"

    num_graficos_diponibles_genericos_gra(CTE_PAL) = CTE_GRA_PAL_num
    grafico_disponible_generico_gra(CTE_PAL, 1) = "Peso frase de más peso"
    grafico_disponible_generico_gra(CTE_PAL, 2) = "Peso frase de menos peso"

    num_graficos_diponibles_genericos_gra(CTE_3R) = CTE_GRA_3R_num
    grafico_disponible_generico_gra(CTE_3R, 1) = "Nº reglas tipo 1 del agente 1"
    grafico_disponible_generico_gra(CTE_3R, 2) = "Nº reglas tipo 2 del agente 1"
    grafico_disponible_generico_gra(CTE_3R, 3) = "Nº reglas tipo 1 del agente 2"
    grafico_disponible_generico_gra(CTE_3R, 4) = "Nº reglas tipo 2 del agente 2"
    grafico_disponible_generico_gra(CTE_3R, 5) = "Nº de movimientos con reglas"
    grafico_disponible_generico_gra(CTE_3R, 6) = "Peso del agente de mas peso (1)"
    

    num_graficos_diponibles_genericos_gra(CTE_ESPECIAL) = CTE_GRA_ESP_num
    grafico_disponible_generico_gra(CTE_ESPECIAL, CTE_ESPECIAL_1_PAL) = CTE_ESPECIALd_1_PAL
    grafico_disponible_generico_gra(CTE_ESPECIAL, CTE_ESPECIAL_2_APE) = CTE_ESPECIALd_2_APE
    grafico_disponible_generico_gra(CTE_ESPECIAL, CTE_ESPECIAL_3_RND) = CTE_ESPECIALd_3_RND
    grafico_disponible_generico_gra(CTE_ESPECIAL, CTE_ESPECIAL_4_PI) = CTE_ESPECIALd_4_PI
    grafico_disponible_generico_gra(CTE_ESPECIAL, CTE_ESPECIAL_5_X2) = CTE_ESPECIALd_5_X2
    grafico_disponible_generico_gra(CTE_ESPECIAL, CTE_ESPECIAL_6_X21) = CTE_ESPECIALd_6_X21


End Sub

Sub s_inicializar_ejes_gra()

    'Ejemplo, si hay dos curvas # rojas y # rosas en el eje 1 seria:
    'dato_a_mostrar_gra(eje 1, indice 1) = grafico 1 (# h rojas)
    'dato_a_mostrar_gra(eje 1, indice 2) = grafico 2 (# h rosas)

    Dim i As Integer
    Dim tipo_graf As Integer
    Dim eje_destino As Integer
    
    'Inicializo los ejes vacios
    numero_total_ejes_gra = CTE_EJE_num_max_ejes
    max_numero_de_graficos_a_mostrar_en_eje_gra = -1
    ReDim horiz_graf_gra(1 To numero_total_ejes_gra) As Integer
    horiz_graf_gra(1) = 60
    horiz_graf_gra(2) = 130
    horiz_graf_gra(3) = 200
    horiz_graf_gra(4) = 270
    horiz_graf_gra(5) = 340
    horiz_graf_gra(6) = 410
    
    'Inicializo todos los ejes a escala real y con 0 datos
    ReDim numero_de_graficos_a_mostrar_en_eje_gra(1 To numero_total_ejes_gra) As Integer 'numero de graficos a mostrar en cada eje
    ReDim escala_gra(1 To numero_total_ejes_gra) As Integer
    ReDim marcas_X_cada_gra(1 To numero_total_ejes_gra) As Integer
    ReDim marcas_Y_cada_gra(1 To numero_total_ejes_gra) As Integer
    ReDim leyenda_gra(1 To numero_total_ejes_gra) As String
    For i = 1 To numero_total_ejes_gra
        numero_de_graficos_a_mostrar_en_eje_gra(i) = 0
        escala_gra(i) = CTE_GRA_AJUSTADA
        marcas_X_cada_gra(i) = 20
        marcas_Y_cada_gra(i) = 20
        leyenda_gra(i) = ""
    Next i
    
    'Inicialmente supongo que no hay ningun grafico disponible
    num_graficos_diponibles_gra = 0
    ReDim grafico_disponible_gra(1 To 1) As String 'lista de todos los graficos disponibles, juntos todos los tipos
    
    ReDim texto_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To 1) As String
    ReDim color_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To 1) As Integer
    ReDim tipo_pto_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To 1) As Integer
    
    'Añado los graficos del programa activo
    If ciclo_ejv > 0 And num_prg_activo_ejv <> CTE_NINGUNO Then
        For i = 1 To num_graficos_diponibles_genericos_gra(num_prg_activo_ejv)
            'Hay un grafico mas disponible de los N que admite el programa actual
            num_graficos_diponibles_gra = num_graficos_diponibles_gra + 1
            ReDim Preserve grafico_disponible_gra(1 To num_graficos_diponibles_gra) As String
            grafico_disponible_gra(num_graficos_diponibles_gra) = grafico_disponible_generico_gra(num_prg_activo_ejv, i)
            'Asocio ese grafico a un eje
            eje_destino = f_SumCirc(numero_total_ejes_gra, i, 0)
            'Propiedades del eje
            leyenda_gra(eje_destino) = leyenda_gra(eje_destino) & " " & grafico_disponible_gra(num_graficos_diponibles_gra)
            'Ahora en ese eje hay un grafico mas, pero si hay mas graficos que ejes empiezo desde el primer eje
            numero_de_graficos_a_mostrar_en_eje_gra(eje_destino) = numero_de_graficos_a_mostrar_en_eje_gra(eje_destino) + 1
            'Actualizo el maximo
            If numero_de_graficos_a_mostrar_en_eje_gra(eje_destino) > max_numero_de_graficos_a_mostrar_en_eje_gra Then
                max_numero_de_graficos_a_mostrar_en_eje_gra = numero_de_graficos_a_mostrar_en_eje_gra(eje_destino)
            End If
            ReDim Preserve texto_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To max_numero_de_graficos_a_mostrar_en_eje_gra) As String
            texto_dato_a_mostrar_gra(eje_destino, numero_de_graficos_a_mostrar_en_eje_gra(eje_destino)) = grafico_disponible_gra(num_graficos_diponibles_gra)
            ReDim Preserve color_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To max_numero_de_graficos_a_mostrar_en_eje_gra) As Integer
            color_dato_a_mostrar_gra(eje_destino, numero_de_graficos_a_mostrar_en_eje_gra(eje_destino)) = f_SumCirc(5, i, 0)
            ReDim Preserve tipo_pto_dato_a_mostrar_gra(1 To numero_total_ejes_gra, 1 To max_numero_de_graficos_a_mostrar_en_eje_gra) As Integer
            tipo_pto_dato_a_mostrar_gra(eje_destino, numero_de_graficos_a_mostrar_en_eje_gra(eje_destino)) = CTE_GRA_LINEA
        Next i
    End If
    
    'Añado los especiales
    'tipo_prog = CTE_ESPECIAL Or

End Sub

Sub s_etiquetas_visibles_gra(estado As Boolean)

    Dim i As Integer
    
    For i = 1 To CTE_EJE_num_max_ejes
        Et_gr(i).Visible = estado
        max_x_eje(i).Visible = estado
        max_y_eje(i).Visible = estado
    Next i
    
End Sub

Sub s_ver_grafico_gra(tipo_soporte As Integer, obj_soporte As Object)

'On Error Resume Next

    'Posibles llamadas
    's_ver_grafico_gra CTE_FORMULARIO, frm_z0_graf
    's_ver_grafico_gra CTE_IMPRESORA, Printer
    
    Dim cont_ejes As Integer
    Dim cont_gra As Integer
    
    Select Case tipo_soporte
        Case CTE_FORMULARIO
            If obj_soporte.Name = "frm_z0_graf" Then
                'obj_soporte.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
                'obj_soporte.Caption = "Evolución Hormigas"
                obj_soporte.Refresh
                obj_soporte.ScaleMode = vbPixels   ' Set scale to pixels.
            End If
        Case CTE_IMPRESORA
            If MsgBox("La impresión puede tardar varios minutos. ¿Está seguro de que desea imprimir las gráficas?", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            Else
                If MsgBox("¿Desea imprimir en color?", vbQuestion + vbYesNo) = vbYes Then
                    Printer.ColorMode = 2
                Else
                   Printer.ColorMode = 1
                End If
                Printer.ScaleMode = vbPoints
                Printer.FillStyle = vbFSSolid
            End If
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error al imprimir"
    End Select
    
    
    Screen.MousePointer = CTE_ARENA
    
    'horizontal
    For cont_ejes = 1 To numero_total_ejes_gra
        obj_soporte.Line (separacion_grafico_gra, horiz_graf_gra(cont_ejes))-(maximo_ancho_ejes + separacion_grafico_gra, horiz_graf_gra(cont_ejes)), cct_ejv(CTE_NEGRO)
    Next cont_ejes
        
    'Vertical
    For cont_ejes = 1 To numero_total_ejes_gra
        obj_soporte.Line (separacion_grafico_gra, horiz_graf_gra(cont_ejes) - 50)-(separacion_grafico_gra, horiz_graf_gra(cont_ejes)), cct_ejv(CTE_NEGRO)
    Next cont_ejes
    
    For cont_ejes = 1 To numero_total_ejes_gra
        'Visualizo etiquetas
        Et_gr(cont_ejes).Caption = leyenda_gra(cont_ejes)
        Et_gr(cont_ejes).Visible = True
    Next cont_ejes
    
    'Aqui hay un error porque antes de pintar los graficos deberia
    'leer todos los graficos de un mismo eje calculanddo sus maximos en
    'y y en x para luego poder pintarlos ambos en la misma escala, la mas
    'restrictiva
    For cont_ejes = 1 To numero_total_ejes_gra
        For cont_gra = 1 To numero_de_graficos_a_mostrar_en_eje_gra(cont_ejes)
            If texto_dato_a_mostrar_gra(cont_ejes, cont_gra) <> "" Then
                s_pintar_grafico_gra tipo_soporte, obj_soporte, cont_gra, cont_ejes
                If cancelar_mostrar_grafico_ejv Then Exit For
            End If
        Next cont_gra
        If cancelar_mostrar_grafico_ejv Then Exit For
    Next cont_ejes


    'Borro los arrays
    'For cont_ejes = 1 To numero_total_ejes_gra
    '    numero_de_graficos_a_mostrar_en_eje_gra(cont_ejes) = 0
    'Next cont_ejes

    If tipo_soporte = CTE_IMPRESORA Then
        Printer.EndDoc
    End If
    Screen.MousePointer = CTE_DEFECTO
    
    'No poner DoEvents porque se borra el grafico


End Sub

Sub s_pintar_grafico_gra(tipo_soporte As Integer, obj_soporte As Object, num_gra As Integer, eje As Integer)

    Dim color_pto As Long
    Dim cont_dato As Long
    Dim campo As String
    Dim linea As String

    Dim valor As Double
    Dim maximo As Double

    Dim tipo_grafico As Integer
    Dim num_gra_res As Integer
    
    Dim posicion_ejeY As Integer
    Dim posicion_ejeX As Integer
    
    Dim ciclos_almacenados As Long
    
    'If esta_detenido_ejv And Not esta_terminado_ejv And un_ej_grabar_gra_ejv Then
    '    s_cerrar_fichero_salida_ejv CTE_FIC_23W_1EJGRA
    'End If
    
    'Calculo el orden del grafico del resumen, sabido el seleccionado
    Select Case texto_dato_a_mostrar_gra(eje, num_gra)
        'Hyp
        Case grafico_disponible_generico_gra(CTE_HYP, 1) '"Nº rojas"
            num_gra_res = 2
            tipo_grafico = CTE_HYP
        Case grafico_disponible_generico_gra(CTE_HYP, 2) '"Nº rosas"
            num_gra_res = 3
            tipo_grafico = CTE_HYP
        Case grafico_disponible_generico_gra(CTE_HYP, 3) '"Nº naranjas"
            num_gra_res = 4
            tipo_grafico = CTE_HYP
        Case grafico_disponible_generico_gra(CTE_HYP, 4) '"Nº amarillas"
            num_gra_res = 5
            tipo_grafico = CTE_HYP
        Case grafico_disponible_generico_gra(CTE_HYP, 5) '"Nº verdes"
            num_gra_res = 6
            tipo_grafico = CTE_HYP
        Case grafico_disponible_generico_gra(CTE_HYP, 6) '"Nº plantas"
            num_gra_res = 7
            tipo_grafico = CTE_HYP
            
        'Pal
        Case grafico_disponible_generico_gra(CTE_PAL, 1) '"Peso frase de más peso"
            num_gra_res = 2
            tipo_grafico = CTE_PAL
        Case grafico_disponible_generico_gra(CTE_PAL, 2) '"Peso frase de menos peso"
            num_gra_res = 3
            tipo_grafico = CTE_PAL
            
        '3r
        Case grafico_disponible_generico_gra(CTE_3R, 1) '"Nº reglas tipo 1 del agente 1"
            num_gra_res = 2
            tipo_grafico = CTE_3R
        Case grafico_disponible_generico_gra(CTE_3R, 2) ' "Nº reglas tipo 2 del agente 1"
            num_gra_res = 3
            tipo_grafico = CTE_3R
        Case grafico_disponible_generico_gra(CTE_3R, 3) ' "Nº reglas tipo 1 del agente 2"
            num_gra_res = 4
            tipo_grafico = CTE_3R
        Case grafico_disponible_generico_gra(CTE_3R, 4) ' "Nº reglas tipo 2 del agente 2"
            num_gra_res = 5
            tipo_grafico = CTE_3R
        Case grafico_disponible_generico_gra(CTE_3R, 5) ' "Nº de movimientos con reglas"
            num_gra_res = 6
            tipo_grafico = CTE_3R
        Case grafico_disponible_generico_gra(CTE_3R, 6) ' "Peso del agente de mas peso (1)"
            num_gra_res = 7
            tipo_grafico = CTE_3R
            
        'especial
        Case CTE_ESPECIALd_1_PAL
            num_gra_res = CTE_ESPECIAL_1_PAL
            tipo_grafico = CTE_ESPECIAL
        Case CTE_ESPECIALd_2_APE
            num_gra_res = CTE_ESPECIAL_2_APE
            tipo_grafico = CTE_ESPECIAL
        Case CTE_ESPECIALd_3_RND
            num_gra_res = CTE_ESPECIAL_3_RND
            tipo_grafico = CTE_ESPECIAL
        Case CTE_ESPECIALd_4_PI
            num_gra_res = CTE_ESPECIAL_4_PI
            tipo_grafico = CTE_ESPECIAL
        Case CTE_ESPECIALd_5_X2
            num_gra_res = CTE_ESPECIAL_5_X2
            tipo_grafico = CTE_ESPECIAL
        Case CTE_ESPECIALd_6_X21
            num_gra_res = CTE_ESPECIAL_6_X21
            tipo_grafico = CTE_ESPECIAL
        
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select
    
    'Pongo como anterior el 0,0
    ant_CX_gra = separacion_grafico_gra
    ant_CY_gra = horiz_graf_gra(eje)
    maximo = -1

    color_pto = cct_ejv(color_dato_a_mostrar_gra(eje, num_gra))
    Select Case tipo_grafico
        Case CTE_ESPECIAL
'==============================================================
            Select Case num_gra_res
                Case CTE_ESPECIAL_1_PAL
                    'Longitud Palabras'numero_palabras_dicc_pal
                    'Calculo el maximo en un bucle separado
                    If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                        For cont_dato = 1 To num_datos_mostrar_especial_gra(num_gra_res)
                            If Len(palabra_del_diccionario(cont_dato)) > maximo Then
                                maximo = Len(palabra_del_diccionario(cont_dato))
                            End If
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Exit Sub
End If
                        Next cont_dato
                        'Al final, si todo son ceros, pongo 1 como maximo (es arbitrario)
                        If maximo = 0 Then maximo = 1
                    End If
                    'Visualizo etiquetas maximo eje Y
                    If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                        max_y_eje(eje).Caption = maximo
                    Else
                        max_y_eje(eje).Caption = 50
                    End If
                    max_y_eje(eje).Visible = True
                    'Pinto puntos
                    For cont_dato = 1 To num_datos_mostrar_especial_gra(num_gra_res)
                        s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, num_datos_mostrar_especial_gra(num_gra_res), Len(palabra_del_diccionario(cont_dato)), cont_dato, maximo, color_pto
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Exit Sub
End If
                    Next cont_dato
                Case CTE_ESPECIAL_2_APE
                    'Longitud Apellidos'CTE_numero_maximo_apellidos
                    'Calculo el maximo en un bucle separado
                    If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                        For cont_dato = 1 To num_datos_mostrar_especial_gra(num_gra_res)
                            If Len(apellidos_posibles(cont_dato)) > maximo Then
                                maximo = Len(apellidos_posibles(cont_dato))
                            End If
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Exit Sub
End If
                        Next cont_dato
                        'Al final, si todo son ceros, pongo 1 como maximo (es arbitrario)
                        If maximo = 0 Then maximo = 1
                    End If
                    'Pinto puntos
                    For cont_dato = 1 To num_datos_mostrar_especial_gra(num_gra_res)
                        s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, num_datos_mostrar_especial_gra(num_gra_res), Len(apellidos_posibles(cont_dato)), cont_dato, maximo, color_pto
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Exit Sub
End If
                    Next cont_dato
                Case CTE_ESPECIAL_3_RND
                    'RND Visual Basic'CTE_numero_maximo_rnd
                    'Calculo el maximo
                    maximo = 9
                    'Visualizo etiquetas maximo eje Y
                    If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                        max_y_eje(eje).Caption = maximo
                    Else
                        max_y_eje(eje).Caption = 50
                    End If
                    max_y_eje(eje).Visible = True
                    'Pinto puntos
                    For cont_dato = 1 To num_datos_mostrar_especial_gra(num_gra_res)
                        s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, num_datos_mostrar_especial_gra(num_gra_res), Int(Rnd * 9), cont_dato, maximo, color_pto
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Exit Sub
End If
                    Next cont_dato
                Case CTE_ESPECIAL_4_PI
                    'Pi'CTE_numero_maximo_pi
                    'Calculo el maximo
                    maximo = 9
                    'Visualizo etiquetas maximo eje Y
                    If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                        max_y_eje(eje).Caption = maximo
                    Else
                        max_y_eje(eje).Caption = 50
                    End If
                    max_y_eje(eje).Visible = True
                    'Pinto puntos
                    For cont_dato = 1 To num_datos_mostrar_especial_gra(num_gra_res)
                        s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, num_datos_mostrar_especial_gra(num_gra_res), CLng(digitos_azar(cont_dato)), cont_dato, maximo, color_pto
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Exit Sub
End If
                    Next cont_dato
                Case CTE_ESPECIAL_5_X2
                    '"x[t]=1-x[t-1]^2"'CTE_numero_maximo_x2
                    'Calculo el maximo
                    maximo = 1
                    'Visualizo etiquetas maximo eje Y
                    If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                        max_y_eje(eje).Caption = maximo
                    Else
                        max_y_eje(eje).Caption = 50
                    End If
                    max_y_eje(eje).Visible = True
                    'Pinto puntos
                    semilla = CDbl(txt_semilla.Text)
                    cont_dato = 1
                    s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, num_datos_mostrar_especial_gra(num_gra_res), semilla, cont_dato, maximo, color_pto
                    For cont_dato = 2 To num_datos_mostrar_especial_gra(num_gra_res)
                        semilla = 1 - (semilla ^ 2)
                        'semilla = CDbl(Left(CStr(semilla), 4))
                        s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, num_datos_mostrar_especial_gra(num_gra_res), semilla, cont_dato, maximo, color_pto
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Exit Sub
End If
                    Next cont_dato
                Case CTE_ESPECIAL_6_X21
                    '"x[t]=(2*x[t-1]^2)-1"'CTE_numero_maximo_x21
                    'Calculo el maximo
                    maximo = 1
                    'Visualizo etiquetas maximo eje Y
                    If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                        max_y_eje(eje).Caption = maximo
                    Else
                        max_y_eje(eje).Caption = 50
                    End If
                    max_y_eje(eje).Visible = True
                    'Pinto puntos
                    semilla = CDbl(txt_semilla.Text)
                    cont_dato = 1
                    s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, num_datos_mostrar_especial_gra(num_gra_res), semilla, cont_dato, maximo, color_pto
                    For cont_dato = 2 To num_datos_mostrar_especial_gra(num_gra_res)
                        semilla = 2 * (semilla ^ 2) - 1
                        'semilla = CDbl(Left(CStr(semilla), 4))
                        s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, num_datos_mostrar_especial_gra(num_gra_res), semilla, cont_dato, maximo, color_pto
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Exit Sub
End If
                    Next cont_dato
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End Select
            max_x_eje(eje).Caption = num_datos_mostrar_especial_gra(num_gra_res)
            max_x_eje(eje).Visible = True
            'Pinto las marcas en horizontal, para todos los especiales
            If marcas_X_cada_gra(eje) > 0 Then
                For cont_dato = 1 To num_datos_mostrar_especial_gra(num_gra_res) Step marcas_X_cada_gra(eje)
                    posicion_ejeX = cont_dato * (CDbl(maximo_ancho_ejes) / CDbl(num_datos_mostrar_especial_gra(num_gra_res)))
                    Me.Line (separacion_grafico_gra + posicion_ejeX, horiz_graf_gra(eje) - 2)-(separacion_grafico_gra + posicion_ejeX, horiz_graf_gra(eje) + 2), cct_ejv(CTE_NEGRO)
                Next cont_dato
            End If
'==============================================================
'==============================================================
'==============================================================
'==============================================================
        Case Else
            If ciclo_ejv > 1 Then
                'Calculo el maximo en un bucle separado
                If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                    Open un_ej_fichero_gra_ejv For Input As #CTE_FIC_23R_1EJGRA
                    ciclos_almacenados = 0
                    campo = f_leer_campo(num_gra_res, f_leer_linea(CTE_FIC_23R_1EJGRA))
                    While Len(campo) > 0
                        ciclos_almacenados = ciclos_almacenados + 1
                        If IsNumeric(campo) Then
                            valor = CDbl(campo)
                        Else
                            s_error_ejv CON_OPCION_FINALIZAR, "Error"
                        End If
                        If valor > maximo Then
                            maximo = valor
                        End If
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Close #CTE_FIC_23R_1EJGRA
    Exit Sub
End If
                        campo = f_leer_campo(num_gra_res, f_leer_linea(CTE_FIC_23R_1EJGRA))
                    Wend
                    Close #CTE_FIC_23R_1EJGRA
                    'Al final, si todo son ceros, pongo 1 como maximo (es arbitrario)
                    If maximo = 0 Then maximo = 1
                End If
                'Visualizo etiquetas maximo eje Y
                If escala_gra(eje) = CTE_GRA_AJUSTADA Then
                    max_y_eje(eje).Caption = maximo
                Else
                    max_y_eje(eje).Caption = 50
                End If
                max_y_eje(eje).Visible = True
                'Pinto puntos
                Open un_ej_fichero_gra_ejv For Input As #CTE_FIC_23R_1EJGRA
                'For cont_dato = 1 To ciclos_almacenados
                'Next cont_dato
                cont_dato = 0
                While Not EOF(CTE_FIC_23R_1EJGRA) And cont_dato < max_guardado_ejv
                    linea = f_leer_linea(CTE_FIC_23R_1EJGRA)
                    If linea <> "" Then
                        valor = CDbl(f_leer_campo(num_gra_res, linea))
                        cont_dato = cont_dato + 1
                        s_pintar_valor_grafico_gra tipo_soporte, obj_soporte, num_gra, eje, ciclo_ejv, valor, cont_dato, maximo, color_pto
                    End If
DoEvents
If cancelar_mostrar_grafico_ejv Then
    Close #CTE_FIC_23R_1EJGRA
    Exit Sub
End If
                Wend
                Close #CTE_FIC_23R_1EJGRA
                max_x_eje(eje).Caption = ciclo_ejv
                If ciclo_ejv > max_guardado_ejv Then
                    max_x_eje(eje).Caption = max_guardado_ejv
                End If
                max_x_eje(eje).Visible = True
            End If
            'Pinto las marcas en horizontal
            If marcas_X_cada_gra(eje) > 0 Then
                For cont_dato = 1 To ciclos_almacenados Step marcas_X_cada_gra(eje)
                    posicion_ejeX = cont_dato * (CDbl(maximo_ancho_ejes) / CDbl(ciclos_almacenados))
                    Me.Line (separacion_grafico_gra + posicion_ejeX, horiz_graf_gra(eje) - 2)-(separacion_grafico_gra + posicion_ejeX, horiz_graf_gra(eje) + 2), cct_ejv(CTE_NEGRO)
                Next cont_dato
            End If
    
    End Select
        
    'Pinto las marcas en la vertical segun los maximos
    If marcas_Y_cada_gra(eje) > 0 Then
        For cont_dato = 1 To maximo Step marcas_Y_cada_gra(eje)
            posicion_ejeY = horiz_graf_gra(eje) - (CDbl((cont_dato - 1) * 50) / CDbl(maximo))
            Me.Line (separacion_grafico_gra - 2, posicion_ejeY)-(separacion_grafico_gra + 2, posicion_ejeY), cct_ejv(CTE_NEGRO)
        Next cont_dato
    End If
    
End Sub
Sub s_pintar_valor_grafico_gra(tipo_soporte As Integer, obj_soporte As Object, num_gra As Integer, eje As Integer, max_valor_ejeX As Long, valor_actual_ejeY As Double, valor_actual_ejeX As Long, maximo As Double, color_pto As Long)

    Dim posicion_ejeZ As Double
    Dim posicion_ejeY As Double
    Dim posicion_ejeX As Double


    'Ajustado o Real (el X siempre se ajusta)
    posicion_ejeX = valor_actual_ejeX * (CDbl(maximo_ancho_ejes) / CDbl(max_valor_ejeX))
    Select Case escala_gra(eje)
        Case CTE_GRA_AJUSTADA
            'Ajustado
            posicion_ejeY = horiz_graf_gra(eje) - valor_actual_ejeY * (CDbl(50) / CDbl(maximo))
        Case CTE_GRA_REAL
            'Real
            posicion_ejeY = horiz_graf_gra(eje) - valor_actual_ejeY
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error"
    End Select

    If posicion_ejeX <> posicion_ejeXanterior Or posicion_ejeY <> posicion_ejeYanterior Then
        'Punto o linea con anterior
        Select Case tipo_pto_dato_a_mostrar_gra(eje, num_gra)
            Case CTE_GRA_PTO
                s_pintar_objeto_ejv tipo_soporte, obj_soporte, posicion_ejeZ, posicion_ejeY, posicion_ejeX, CTE_PUNTO, color_pto, color_pto, CTE_DIRECC_NINGUNA, CTE_ZOOM_PIXELS, 1
            Case CTE_GRA_LINEA
                s_pintar_objeto_ejv tipo_soporte, obj_soporte, posicion_ejeZ, posicion_ejeY, posicion_ejeX, "linea_con_anterior", color_pto, color_pto, CTE_DIRECC_NINGUNA, CTE_ZOOM_PIXELS, 1
            Case Else
                s_error_ejv CON_OPCION_FINALIZAR, "Error"
        End Select
    End If
    
    posicion_ejeXanterior = posicion_ejeX
    posicion_ejeYanterior = posicion_ejeY
    
End Sub
Function s_calcular_color_str2int_gra(color As String) As Integer

    Dim indice As Integer

    indice = f_buscar_en_array(nct_ejv(), color)
    If indice = -1 Then
        MsgBox "Error: color no encontrado", vbCritical
    Else
        s_calcular_color_str2int_gra = indice
    End If
    
End Function

Sub s_ver_propiedades_eje_gra()

    Dim i As Integer
    
    eje_modificar_opciones_gra = Cb_Ejes.ListIndex
    If eje_modificar_opciones_gra > 0 Then
        Fr_Grafico.Caption = "Datos mostrados en el eje " & eje_modificar_opciones_gra
    Else
        Fr_Grafico.Caption = "Opciones de todos los ejes"
    End If
    Fr_Grafico.Visible = True
    lst_Datos_a_mostrar.Enabled = True
    lst_Color.Enabled = True
    lst_Tipo.Enabled = True
    
    If eje_modificar_opciones_gra = 0 Then
        lst_Datos_a_mostrar.Enabled = False
        lst_Datos_a_mostrar.BackColor = cct_ejv(cfondo_ejv)
        lst_Color.Enabled = False
        lst_Color.BackColor = cct_ejv(cfondo_ejv)
        lst_Tipo.Enabled = False
        lst_Tipo.BackColor = cct_ejv(cfondo_ejv)
        'Datos comunes al eje, pongo uno cualquiera, por ejemplo el primero
        Cb_Escala.ListIndex = escala_gra(1) - 1
        marcas_X = ""
        marcas_Y = ""
        txt_leyenda = ""
    Else
        'Muestro los datos del eje seleccionado para cada grafico de ese eje
        lst_Datos_a_mostrar.Clear
        lst_Datos_a_mostrar.Enabled = True
        lst_Datos_a_mostrar.BackColor = cct_ejv(CTE_BLANCO)
        lst_Color.Clear
        lst_Color.Enabled = True
        lst_Color.BackColor = cct_ejv(CTE_BLANCO)
        lst_Tipo.Clear
        lst_Tipo.Enabled = True
        lst_Tipo.BackColor = cct_ejv(CTE_BLANCO)
        For i = 1 To numero_de_graficos_a_mostrar_en_eje_gra(eje_modificar_opciones_gra)
            lst_Datos_a_mostrar.AddItem texto_dato_a_mostrar_gra(eje_modificar_opciones_gra, i)
            lst_Color.AddItem nct_ejv(color_dato_a_mostrar_gra(eje_modificar_opciones_gra, i))
            Select Case tipo_pto_dato_a_mostrar_gra(eje_modificar_opciones_gra, i)
                Case CTE_GRA_PTO
                    lst_Tipo.AddItem "Puntos"
                Case CTE_GRA_LINEA
                    lst_Tipo.AddItem "Líneas"
                Case Else
                    s_error_ejv CON_OPCION_FINALIZAR, "Error"
            End Select
        Next i
        'Datos comunes al eje
        Cb_Escala.ListIndex = escala_gra(eje_modificar_opciones_gra) - 1
        marcas_X = marcas_X_cada_gra(eje_modificar_opciones_gra)
        marcas_Y = marcas_Y_cada_gra(eje_modificar_opciones_gra)
        txt_leyenda = leyenda_gra(eje_modificar_opciones_gra)
    End If

End Sub


Private Sub VerPropiedadesEje_Click()
    s_ver_propiedades_eje_gra
End Sub

Sub s_vaciar_etiquetas_gra()

    Dim i As Integer
    
    For i = 1 To CTE_EJE_num_max_ejes
        Et_gr(i).Caption = ""
        max_x_eje(i).Caption = ""
        max_y_eje(i).Caption = ""
    Next i

End Sub

Sub s_aut_abrir_graf_ejv()

On Error GoTo abrir_error
    
    'Abre el gráfico indicado por nombre_fichero_ejv
    
    Dim p As Double
    Dim f As Double
    Dim c As Double
    Dim linea As String
    Dim dato As Integer
    Dim s_dato As String
    Dim primera_coma As Integer
    

    'Lo cargo en la pantalla
    p = 0
    Open nombre_fichero_ejv For Input As #CTE_FIC_23R_1EJGRA
    Screen.MousePointer = CTE_ARENA
    'Filas
    linea = f_leer_linea(CTE_FIC_23R_1EJGRA)
    mapa_filas_ma0 = CInt(Trim(Right(linea, Len(linea) - Len("FILAS="))))
    'Columnas
    linea = f_leer_linea(CTE_FIC_23R_1EJGRA)
    mapa_columnas_ma0 = CInt(Trim(Right(linea, Len(linea) - Len("COLUMNAS="))))
    
    ReDim mapa_ma0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Boolean
    ReDim nodo_visitado_va0(1 To mapa_pisos_ma0, 1 To mapa_filas_ma0, 1 To mapa_columnas_ma0) As Integer
    
    'Modo de presentacion
    linea = f_leer_linea(CTE_FIC_23R_1EJGRA)
    Select Case UCase(Trim(Right(linea, Len(linea) - Len("DETALLE:"))))
        Case CTE_tZOOM_DETALLE
            ver_zoom_ma0 = CTE_ZOOM_DETALLE
        Case CTE_tZOOM_PANORAMICA
            ver_zoom_ma0 = CTE_ZOOM_PANORAMICA
        Case CTE_tZOOM_PIXELS
            ver_zoom_ma0 = CTE_ZOOM_PIXELS
        Case CTE_tZOOM_3D
            ver_zoom_ma0 = CTE_ZOOM_3D
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
    End Select
    s_fijar_separacion_mapa_ma0
    
    'Celdas de obstaculos
    'Inicializo todas y asi si falta algo queda como vacio
    s_mapa_inicializar_va0 mapa_ma0, mapa_pisos_ma0, mapa_filas_ma0, mapa_columnas_ma0, False
    'Leo el mapa
    While Not EOF(CTE_FIC_23R_1EJGRA) And f < mapa_filas_ma0
        f = f + 1
        linea = f_leer_linea(CTE_FIC_23R_1EJGRA)
        If Len(linea) > 0 Then
            For c = 1 To mapa_columnas_ma0
                primera_coma = InStr(linea, ",")
                If primera_coma = 0 Then
                    s_dato = linea
                    If IsNumeric(s_dato) Then
                        dato = Int(s_dato)
                    Else
                        dato = 0
                    End If
                    linea = ""
                Else
                    s_dato = Left(linea, primera_coma - 1)
                    If IsNumeric(s_dato) Then
                        dato = Int(s_dato)
                    Else
                        dato = 0
                    End If
                    linea = Right(linea, Len(linea) - primera_coma)
                End If
                If dato = 1 Then
                    mapa_ma0(p, f, c) = True
                Else
                    mapa_ma0(p, f, c) = False
                End If
            Next c
        End If
    Wend
    
    Screen.MousePointer = CTE_DEFECTO
    Close #CTE_FIC_23R_1EJGRA

'=======================================================
    Exit Sub
abrir_error:
    MsgBox "No se encuentra el fichero " & nombre_fichero_ejv & " o es incorrecto", vbCritical
    Screen.MousePointer = CTE_DEFECTO
    Close #CTE_FIC_23R_1EJGRA
    

End Sub


Sub s_usr_guardar_graf_ejv()

    Dim linea As String
    Dim p As Double
    Dim f As Double
    Dim c As Double
    Dim dato As Integer
    Dim s_dato As String
    Dim primera_coma As Integer
    
    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_SAL_GRA)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Guardar Fichero de Gráfico"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "Guardar"
    frm_z0_fic.File1.Pattern = "*.gra"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path-fichero
    frm_a0_mapa.Caption = "Gráfico Almacenado " & nombre_fichero_ejv
    
    'Lo grabo
    f = 0
    
    'Compruebo que no existe
    If f_existe_fichero(nombre_fichero_ejv) Then
        If MsgBox("El fichero ya existe. ¿Desea reemplazarlo?", vbQuestion + vbOKCancel) = vbCancel Then
            'Ha elegido cancelar la operación
            GoTo fin
        End If
    End If
    
    'Abro y leo
    Open nombre_fichero_ejv For Output As #CTE_FIC_23W_1EJGRA
    'Filas
    linea = "FILAS=" & mapa_filas_ma0
    'El Write graba entre "" y el Print no
    Print #CTE_FIC_23W_1EJGRA, linea
    'Columnas
    linea = "COLUMNAS=" & mapa_columnas_ma0
    Print #CTE_FIC_23W_1EJGRA, linea
    'Modo de presentacion
    Select Case ver_zoom_ma0
        Case CTE_ZOOM_DETALLE
            linea = "ZOOM=" & CTE_tZOOM_DETALLE
        Case CTE_ZOOM_PANORAMICA
            linea = "ZOOM=" & CTE_tZOOM_PANORAMICA
        Case CTE_ZOOM_PIXELS
            linea = "ZOOM=" & CTE_tZOOM_PIXELS
        Case CTE_ZOOM_3D
            linea = "ZOOM=" & CTE_tZOOM_3D
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: zoom incorrecto"
    End Select
    Print #CTE_FIC_23W_1EJGRA, linea
    'Celdas de obstaculos
    For f = 1 To mapa_filas_ma0
        linea = ""
        For c = 1 To mapa_columnas_ma0
            If mapa_ma0(p, f, c) = True Then
                linea = linea & "1,"
            Else
                linea = linea & "0,"
            End If
        Next c
       'quitamos la coma final
        linea = Left(linea, Len(linea) - 1)
        Print #CTE_FIC_23W_1EJGRA, linea
    Next f
    Close #CTE_FIC_23W_1EJGRA

fin:

    'Como la pantalla de "Guardar Como.." habra borrado el mapa, lo muestro otra vez
    s_refrescar_mapa_actual_ma0

End Sub



Sub s_usr_abrir_graf_ejv()

On Error GoTo abrir_error

    
    'Elijo path por defecto
    nombre_fichero_ejv = path_largo_ejv(CTE_C_SAL_GRA)
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH
    frm_z0_fic.Caption = "Abrir Fichero de Gráfico"  'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Abrir"
    frm_z0_fic.File1.Pattern = "*.gra"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    
    'Muestro el nuevo path-fichero
    frm_a0_mapa.Caption = "Gráfico Almacenado " & nombre_fichero_ejv
    frm_a0_mapa.Refresh
    
        
    'Cargo nombre_fichero_ejv
    s_aut_abrir_graf_ejv
    
    
'=======================================================
    Exit Sub
abrir_error:
    MsgBox "No se encuentra el fichero " & nombre_fichero_ejv & " o es incorrecto", vbCritical
    Close #CTE_FIC_23R_1EJGRA
    

End Sub
