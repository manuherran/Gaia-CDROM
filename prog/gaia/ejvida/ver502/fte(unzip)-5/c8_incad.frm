VERSION 5.00
Begin VB.Form frm_c8_incad 
   ClientHeight    =   8145
   ClientLeft      =   8310
   ClientTop       =   600
   ClientWidth     =   3195
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
   Icon            =   "c8_incad.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8145
   ScaleWidth      =   3195
   Begin VB.PictureBox semaforo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2880
      Picture         =   "c8_incad.frx":0442
      ScaleHeight     =   735
      ScaleWidth      =   375
      TabIndex        =   55
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Segundos totales"
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
      Height          =   195
      Left            =   1560
      TabIndex        =   57
      Top             =   7800
      Width           =   1230
   End
   Begin VB.Label SegTot 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   56
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   2760
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label media 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   54
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Duración de los ciclos (en segundos)"
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
      Height          =   495
      Left            =   1560
      TabIndex        =   53
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label segundosr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Segundos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   51
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label minutosr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Minutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   49
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label horasr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   48
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Horas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   47
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label diasr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Días"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   45
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Años"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   44
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label anosr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Meses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   42
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label mesesr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label horaf 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Hora Fin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      TabIndex        =   39
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label horac 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Hora Comienzo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1560
      TabIndex        =   37
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label fechaf 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Fin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      TabIndex        =   35
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label fechac 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Comienzo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2760
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label txt_MuerenVejez 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   32
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label label444 
      Caption         =   "Nº Mueren Vejez"
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
      Left            =   1560
      TabIndex        =   31
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Agente"
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
      Left            =   1560
      TabIndex        =   30
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label txt_agente 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label_titulo_estado 
      Caption         =   "Estado:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label_estado 
      Caption         =   "Detenido"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   840
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label_warning 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Espere, por favor..."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label34 
      Caption         =   "Nº Pelean"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   1560
      TabIndex        =   25
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label33 
      Caption         =   "Nº Mueren"
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
      Left            =   1560
      TabIndex        =   24
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label32 
      Caption         =   "Nº Nacen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label31 
      Caption         =   "Nº Riegan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label30 
      Caption         =   "Energía Total"
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
      Left            =   1560
      TabIndex        =   21
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label29 
      Caption         =   "Ciclo"
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
      Left            =   1560
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "Total Hormigas"
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
      Left            =   1560
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label27 
      Caption         =   "Total Plantas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "Nº verdes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label25 
      Caption         =   "Nº amarillas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label24 
      Caption         =   "Nº naranjas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "Nº rosas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Nº rojas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label txt_pelean 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label txt_mueren 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label txt_nacen 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label txt_riegan 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label txt_senergía 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label txt_hormigas 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label txt_plantas 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label txt_ciclo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label verdes 
      Alignment       =   2  'Center
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label amarillas 
      Alignment       =   2  'Center
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label naranjas 
      Alignment       =   2  'Center
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label rosas 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label rojas 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frm_c8_incad"
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
    Me.KeyPreview = True 'permito recibir teclas
    ajuste_color_controles_formulario_ejv Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

