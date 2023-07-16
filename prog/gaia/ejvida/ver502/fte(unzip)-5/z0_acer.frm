VERSION 5.00
Begin VB.Form frm_z0_acer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5490
   ClientLeft      =   2250
   ClientTop       =   1785
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   -1  'True
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "z0_acer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   6735
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6495
      Begin VB.ComboBox Cb_Util 
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
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Util 
         Caption         =   "&Ejecutar Utilidad..."
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
         Left            =   2640
         TabIndex        =   2
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agra&decimientos..."
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
         Left            =   600
         TabIndex        =   0
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Text            =   "gaiasoft@geocities.com"
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Text            =   "mherran@usa.net"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Text            =   "http://www.geocities.com/SiliconValley/Vista/7491/"
         Top             =   3960
         Width           =   5415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "La Rana Blanca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4560
         TabIndex        =   15
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Image Image9 
         Height          =   810
         Left            =   600
         Picture         =   "z0_acer.frx":0442
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image Image8 
         Height          =   810
         Left            =   600
         Picture         =   "z0_acer.frx":43CC
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image Image7 
         Height          =   810
         Left            =   600
         Picture         =   "z0_acer.frx":8356
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image Image6 
         Height          =   810
         Left            =   600
         Picture         =   "z0_acer.frx":C2E0
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image Image5 
         Height          =   810
         Left            =   600
         Picture         =   "z0_acer.frx":1026A
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image Image4 
         Height          =   810
         Left            =   600
         Picture         =   "z0_acer.frx":141F4
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label ver 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
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
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   5460
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   810
         Left            =   600
         Picture         =   "z0_acer.frx":1817E
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         DrawMode        =   14  'Copy Pen
         X1              =   480
         X2              =   6000
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         DrawMode        =   14  'Copy Pen
         X1              =   480
         X2              =   6000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   255
         TabIndex        =   12
         Top             =   240
         Width           =   5925
      End
      Begin VB.Label La_CopyRight 
         AutoSize        =   -1  'True
         Caption         =   "Desarrollado por:"
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
         Left            =   600
         TabIndex        =   13
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Line Li_Linea_1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         DrawMode        =   14  'Copy Pen
         X1              =   480
         X2              =   6000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manu Herrán Gascón"
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
         Left            =   1920
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   $"z0_acer.frx":1C108
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2280
         TabIndex        =   9
         Top             =   2880
         Width           =   4095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Madrid (España)"
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
         Left            =   2640
         TabIndex        =   10
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   4320
         Picture         =   "z0_acer.frx":1C1AE
         Top             =   1200
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   6495
      Begin VB.Timer Ti_1segundo 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   240
      End
      Begin VB.Timer Ti_10segundos 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   960
         Top             =   240
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
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   4560
         Top             =   240
      End
   End
End
Attribute VB_Name = "frm_z0_acer"
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
    'Descargamos
     Unload Me

End Sub


Private Sub Command1_Click()
    
    frm_z0_agra.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show

End Sub



Private Sub Form_Load()
   
    'Permito recibir teclas
    Me.KeyPreview = True
    
    Dim num_instancia As Long
           
    Timer1.Enabled = False
    
    label.Caption = nombre_aplicacion_ejv & " para Windows 95, 98 y NT"
    
    'num_instancia = GetModuleHandle(App.EXEName & ".exe")
    'num_instancia = GetModuleUsage(handle)
    'hInstance = " & App.hInstance & ". ThreadID = " & App.ThreadID & ". Instancia de ejecutable " & num_instancia
    ver.Caption = "Revisión " & App.Major & "." & App.Minor & "." & App.Revision & ". ThreadID " & App.ThreadID

    'Mostramos la pantalla en el centro
     s_centrar_ventana_ejv Me

    Image3.Visible = True
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    Image8.Visible = False
    Image9.Visible = False
    
    
    Cb_Util.Clear
    Cb_Util.AddItem "Colores..."
    Cb_Util.AddItem "Info de Sistema..."
    Cb_Util.AddItem "Backup..."
    Cb_Util.AddItem "Fractal..."
    Cb_Util.AddItem "Árbol..."
    Cb_Util.AddItem "Filtro..."
    Cb_Util.AddItem "Diccionario..."
    Cb_Util.AddItem "Solitario..."
    Cb_Util.AddItem "Teclas..."
    Cb_Util.AddItem "Encriptar..."
    Cb_Util.AddItem "Video..."
    Cb_Util.AddItem "Fontainebleau..."
    Cb_Util.AddItem "Generador de SWT..."
    Cb_Util.ListIndex = 0
    
    
    
    

End Sub


Private Sub Image1_Click()
    Timer1.Enabled = Not Timer1.Enabled

End Sub

Private Sub Ti_10segundos_Timer()
    
    Ti_10segundos.Enabled = False

End Sub

Private Sub Ti_1segundo_Timer()

    Ti_1segundo.Enabled = False

End Sub

Private Sub Timer1_Timer()

    If Image3.Visible Then
        Image3.Visible = False
        Image4.Visible = True
    ElseIf Image4.Visible Then
        Image4.Visible = False
        Image5.Visible = True
    ElseIf Image5.Visible Then
        Image5.Visible = False
        Image6.Visible = True
    ElseIf Image6.Visible Then
        Image6.Visible = False
        Image7.Visible = True
    ElseIf Image7.Visible Then
        Image7.Visible = False
        Image8.Visible = True
    ElseIf Image8.Visible Then
        Image8.Visible = False
        Image9.Visible = True
    ElseIf Image9.Visible Then
        Image9.Visible = False
        Image3.Visible = True
    End If


End Sub

Private Sub Util_Click()
    
    Select Case Cb_Util
        Case "Colores..."
            frm_u0_color.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Info de Sistema..."
            frm_u0_info.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Backup..."
            frm_u0_bk.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Árbol..."
            frm_u0_arbo.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Fractal..."
            frm_u0_frac.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Filtro..."
            frm_u0_filt.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Diccionario..."
            frm_u0_dicc.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Solitario..."
            'frm_u0_soli.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Teclas..."
            If MsgBox("Esta utilidad envia teclas a otro programa. El comportamiento se define en el fichero teclas.txt Se recomienda consultar este fichero antes de ejecutar la utilidad", vbOKCancel + vbDefaultButton2 + vbExclamation) = vbOK Then
                s_lanzar_teclas
            End If
            'frm_u0_tecl.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Encriptar..."
            frm_u0_encr.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Video..."
            Unload Me
            s_video
        Case "Fontainebleau..."
            frm_u0_font.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case "Generador de SWT..."
            frm_u0_swt.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
        Case Else
            s_error_ejv CON_OPCION_FINALIZAR, "Error: "
    End Select

End Sub
