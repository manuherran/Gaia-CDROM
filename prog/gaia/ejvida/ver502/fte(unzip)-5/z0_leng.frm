VERSION 5.00
Begin VB.Form frm_z0_leng 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Idioma - Languagge"
   ClientHeight    =   2265
   ClientLeft      =   4080
   ClientTop       =   3300
   ClientWidth     =   3750
   Icon            =   "z0_leng.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2265
   ScaleWidth      =   3750
   Begin VB.CommandButton Aceptar 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line i3 
      X1              =   1920
      X2              =   2760
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line e3 
      X1              =   840
      X2              =   1680
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line i1 
      X1              =   1920
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line e1 
      X1              =   840
      X2              =   1680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line i2 
      X1              =   2760
      X2              =   2760
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Line i4 
      X1              =   1920
      X2              =   1920
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Line e2 
      X1              =   1680
      X2              =   1680
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Line e4 
      X1              =   840
      X2              =   840
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   2040
      Picture         =   "z0_leng.frx":0442
      Top             =   720
      Width           =   615
   End
   Begin VB.Label label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choose Languagge"
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
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   2385
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   960
      Picture         =   "z0_leng.frx":17E4
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frm_z0_leng"
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
    elegir_idioma_ejv = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyLeft   'LEFT ARROW key
            Image1_Click
        Case vbKeyRight  'RIGHT ARROW key
            Image2_Click
        Case Else
            'Si es otra no hago nada, es algo que se teclea
    End Select

End Sub

Private Sub Form_Load()
     
    s_centrar_ventana_ejv Me
    Me.KeyPreview = True 'permito recibir teclas
    
    Select Case idioma_ejv
        Case CTE_INGLES
            e1.Visible = True
            e2.Visible = True
            e3.Visible = True
            e4.Visible = True
            i1.Visible = False
            i2.Visible = False
            i3.Visible = False
            i4.Visible = False
        Case CTE_CASTELLANO
            e1.Visible = False
            e2.Visible = False
            e3.Visible = False
            e4.Visible = False
            i1.Visible = True
            i2.Visible = True
            i3.Visible = True
            i4.Visible = True
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If e1.Visible = True Then
        idioma_ejv = CTE_CASTELLANO
    Else
        idioma_ejv = CTE_INGLES
    End If
        e1.Visible = True


End Sub

Private Sub Image1_Click()
    e1.Visible = True
    e2.Visible = True
    e3.Visible = True
    e4.Visible = True
    i1.Visible = False
    i2.Visible = False
    i3.Visible = False
    i4.Visible = False
End Sub

Private Sub Image2_Click()
    e1.Visible = False
    e2.Visible = False
    e3.Visible = False
    e4.Visible = False
    i1.Visible = True
    i2.Visible = True
    i3.Visible = True
    i4.Visible = True

End Sub
