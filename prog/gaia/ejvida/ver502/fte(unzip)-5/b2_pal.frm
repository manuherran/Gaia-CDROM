VERSION 5.00
Begin VB.Form frm_b2_pal 
   Caption         =   "Palabras y Frases"
   ClientHeight    =   8595
   ClientLeft      =   810
   ClientTop       =   405
   ClientWidth     =   10395
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
   Icon            =   "B2_PAL.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8595
   ScaleWidth      =   10395
   Begin VB.Frame Fr_Opciones 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
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
         Left            =   4800
         TabIndex        =   45
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton Aceptar 
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
         Left            =   3600
         TabIndex        =   44
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox Op_DiccF 
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
         Left            =   7680
         TabIndex        =   42
         Text            =   "2092.dic"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Op_DiccC 
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
         Left            =   3360
         TabIndex        =   41
         Text            =   "c:\......"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CommandButton Fic_Dicc 
         Caption         =   "..."
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
         TabIndex        =   40
         Top             =   1800
         Width           =   255
      End
      Begin VB.Frame Frame6 
         Caption         =   "Cálculo del Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   28
         Top             =   3720
         Width           =   2415
         Begin VB.OptionButton Op_Relativo 
            Caption         =   "Relativo"
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
            TabIndex        =   30
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Op_nRelativo 
            Caption         =   "Independiente"
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
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Padres idénticos producen mutaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   4680
         Width           =   3135
         Begin VB.OptionButton Op_PadresIdenticos 
            Caption         =   "Sí"
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
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Op_nPadresIdenticos 
            Caption         =   "No"
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
            TabIndex        =   26
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Elección de Padres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   2415
         Begin VB.OptionButton Op_PadresAzar_pal 
            Caption         =   "Al azar"
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
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Op_nPadresAzar_pal 
            Caption         =   "Secuencial"
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
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frecuencia de Mutación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   2415
         Begin VB.OptionButton Op_ind 
            Caption         =   "Independiente"
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
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Op_Acumulada 
            Caption         =   "Acumulada"
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
            TabIndex        =   19
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Seleción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2760
         TabIndex        =   13
         Top             =   2760
         Width           =   6135
         Begin VB.OptionButton Op_s4 
            Caption         =   "D)  40% mejor y 10% peor sobre 50%"
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
            TabIndex        =   17
            Top             =   1080
            Width           =   5775
         End
         Begin VB.OptionButton Op_s3 
            Caption         =   "C)  20% mejor sobreviven y se reproducen sobre 80%"
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
            Top             =   840
            Width           =   5175
         End
         Begin VB.OptionButton Op_s1 
            Caption         =   "A)  50% mejor sobreviven y se reproducen sobre el otro 50%"
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
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   5535
         End
         Begin VB.OptionButton Op_s2 
            Caption         =   "B)  10% mejor sobreviven y se reproducen sobre 90%"
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
            TabIndex        =   14
            Top             =   600
            Width           =   5535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Modo de Ejecución"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2415
         Begin VB.OptionButton Option1 
            Caption         =   "Rápido"
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
            TabIndex        =   12
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Op_VerFrases 
            Caption         =   "Ver Frases"
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
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.ComboBox Cb_num_pal 
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
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text4 
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
         Left            =   6360
         TabIndex        =   6
         Text            =   "20"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text2 
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
         Height          =   525
         Left            =   2280
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   6735
      End
      Begin VB.TextBox num_pal 
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
         Left            =   7560
         TabIndex        =   2
         Text            =   "5"
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Diccionario"
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
         Left            =   3360
         TabIndex        =   43
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   $"B2_PAL.frx":0742
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3480
         TabIndex        =   39
         Top             =   4320
         Width           =   5895
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Hay 1 mutación en la palabra generada por cada "
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
         Left            =   2640
         TabIndex        =   8
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "generaciones de palabras"
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
         Left            =   6720
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Número de frases inicial"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Frase a buscar"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número total de palabras del diccionario"
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
         Left            =   4320
         TabIndex        =   1
         Top             =   2280
         Width           =   3015
      End
   End
   Begin VB.Frame Fr_Ejecucion 
      Caption         =   "Las 10 mejores y las 10 peores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   10215
      Begin VB.ListBox List4 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   8520
         TabIndex        =   37
         Top             =   3840
         Width           =   1575
      End
      Begin VB.ListBox List3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   120
         TabIndex        =   36
         Top             =   3840
         Width           =   8295
      End
      Begin VB.ListBox List2 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   8520
         TabIndex        =   35
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   8295
      End
   End
   Begin VB.Frame fr_Todas 
      Caption         =   "Todas las frases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   10215
      Begin VB.ListBox List6 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6105
         Left            =   8520
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.ListBox List5 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6105
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Image Imagen 
      Height          =   1680
      Left            =   600
      Picture         =   "B2_PAL.frx":08EC
      Top             =   600
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label aviso_ejecutar 
      BackStyle       =   0  'Transparent
      Caption         =   "Para comenzar pulse ""Comenzar"" en el menu ""Ejecutar"" (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frm_b2_pal"
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
    
    s_grabar_opciones_pal
    s_activar_opciones_pal
    frm_b2_pal.fr_Todas.Visible = False
    frm_b2_pal.Fr_Opciones.Visible = False
    frm_b2_pal.Fr_Ejecucion.Visible = True
    
End Sub

Private Sub Cancelar_Click()
    
    frm_b2_pal.fr_Todas.Visible = False
    frm_b2_pal.Fr_Opciones.Visible = False
    frm_b2_pal.Fr_Ejecucion.Visible = True

End Sub

Private Sub Fic_Dicc_Click()
    
    'Fijo una carpeta por defecto
    If Len(Trim(Op_DiccC.Text)) > 0 Then
        nombre_fichero_ejv = Trim(Op_DiccC.Text)
    Else
        nombre_fichero_ejv = path_largo_ejv(CTE_C_ENT_DIC)
    End If
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta o carpeta mas fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.dic"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Op_DiccC.Text = nombre_fichero_ejv
    Else
        Op_DiccC.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_ENT_DIC)
    End If

    'Si ha tecleado un nombre de fichero, lo muestro en su sitio
    If Len(f_nombre_fichero(nombre_fichero_ejv)) > 0 Then
        Op_DiccF = f_nombre_fichero(nombre_fichero_ejv)
    End If

End Sub

Private Sub Form_Activate()
    
    's_identificar_num_prg_activo_ejv
    ha_cambiado_el_diccionaro_pal = False
    
    'Ahora no permito guardar y abrir
    s_cambiar_estado_enabled_operaciones_ficheros_ejv False
    
    
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_PAUSA, False
    s_estado_enabled_ejecucion_ejv
    s_estado_enabled_ver_ejv
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True 'permito recibir teclas
    num_prg_activo_ejv = CTE_PAL
    s_tratamiento_idioma_pal
    
    'Mostramos la pantalla en el centro del monitor
    Me.Height = 7200
    Me.Width = 11000
    Me.WindowState = CTE_MAXIMIZED
    
    s_mostrar_aviso_imagen
    
    Fr_Opciones.Visible = True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES1, True
    s_cambiar_estado_enabled_menus_ejv CTE_VER_OPCIONES2, True
    
    Cb_num_pal.Clear
    Cb_num_pal.AddItem 8
    Cb_num_pal.AddItem 20
    Cb_num_pal.AddItem 40
    Cb_num_pal.AddItem 80
    Cb_num_pal.AddItem 160
    Cb_num_pal.AddItem 320
    
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_COMENZAR, True
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_TERMINAR, False
    'Grafico.Enabled = False
    
    paso_pal = 14
    

    If Not automatico_ejv Then
        s_inicializar_ejemplo_elegido_ejv
    End If
    
    s_cargar_opciones_pal
    
    esta_detenido_ejv = True
    

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = f_control_cerrar_pal

End Sub
Private Sub Form_Unload(Cancel As Integer)
     
    Unload frm_b2_inpal

    'Actualizo el estado de enabled de ejecucion y ver
    num_prg_activo_ejv = CTE_NINGUNO
    
    'Pongo habilitado todos los programas
    s_cambiar_estado_enabled_programas_todos_ejv True
    
    'Cogiendolo de los arrays del num_prg_activo_ejv
    s_estado_enabled_ejecucion_ejv
    s_estado_enabled_ver_ejv

End Sub

Sub s_tratamiento_idioma_pal()
    If idioma_ejv = CTE_INGLES Then
    Else
    
    End If
    
End Sub

Private Sub Op_DiccC_Change()

    ha_cambiado_el_diccionaro_pal = True

End Sub

Private Sub Op_DiccF_Change()
    
    ha_cambiado_el_diccionaro_pal = True

End Sub
