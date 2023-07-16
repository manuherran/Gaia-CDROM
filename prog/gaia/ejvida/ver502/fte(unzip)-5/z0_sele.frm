VERSION 5.00
Begin VB.Form frm_z0_sele 
   Caption         =   "Selector"
   ClientHeight    =   4650
   ClientLeft      =   1875
   ClientTop       =   1395
   ClientWidth     =   11175
   Icon            =   "z0_sele.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4650
   ScaleWidth      =   11175
   Begin VB.ListBox txt_lista_der 
      Height          =   3375
      Left            =   6000
      TabIndex        =   13
      Top             =   480
      Width           =   4575
   End
   Begin VB.ListBox txt_lista_izq 
      Height          =   3375
      Left            =   600
      TabIndex        =   12
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton subir_der 
      Caption         =   "^"
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton bajar_der 
      Caption         =   "v"
      Height          =   375
      Left            =   10680
      TabIndex        =   8
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton subir_izq 
      Caption         =   "^"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton bajar_izq 
      Caption         =   "v"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton a_la_izquierda_seleccionado 
      Caption         =   "<"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton a_la_derecha_todos 
      Caption         =   ">>"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton a_la_izquierda_todos 
      Caption         =   "<<"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton a_la_derecha_seleccionado 
      Caption         =   ">"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Etiqueta_der 
      Caption         =   "Etiqueta_der"
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
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Etiqueta_izq 
      Caption         =   "Etiqueta_izq"
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
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frm_z0_sele"
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
    
    
    If frm_z0_sele.txt_lista_der.ListCount = selector_max_der_sel Then
        ReDim resultado_selector_sel(1 To frm_z0_sele.txt_lista_der.ListCount) As String
        
        For i = 1 To frm_z0_sele.txt_lista_der.ListCount
            resultado_selector_sel(i) = frm_z0_sele.txt_lista_der.List(i - 1)
        Next i
        modificar_resultado_selector_sel = True
    
        Unload Me
    Else
        If selector_max_der_sel = 1 Then
            MsgBox "Debe seleccionar sólo un item en el cuadro derecho", vbInformation
        Else
            MsgBox "Debe seleccionar " & selector_max_der_sel & " items en el cuadro derecho", vbInformation
        End If
    End If
    
End Sub
Private Sub Cancelar_Click()
    
    modificar_resultado_selector_sel = False
    Unload Me

End Sub

Private Sub Form_Load()
    s_centrar_ventana_ejv Me

End Sub


Private Sub a_la_derecha_seleccionado_Click()
            
    frm_z0_sele.txt_lista_der.AddItem frm_z0_sele.txt_lista_izq.Text
    frm_z0_sele.txt_lista_izq.RemoveItem frm_z0_sele.txt_lista_izq.ListIndex
    
End Sub

Private Sub a_la_izquierda_seleccionado_Click()
    
    frm_z0_sele.txt_lista_izq.AddItem frm_z0_sele.txt_lista_der.Text
    frm_z0_sele.txt_lista_der.RemoveItem frm_z0_sele.txt_lista_der.ListIndex

End Sub



Private Sub a_la_derecha_todos_Click()

    Dim i As Integer
    
    For i = 0 To frm_z0_sele.txt_lista_izq.ListCount
        frm_z0_sele.txt_lista_der.AddItem frm_z0_sele.txt_lista_izq.List(i)
        frm_z0_sele.txt_lista_izq.RemoveItem i
    Next i

End Sub

Private Sub a_la_izquierda_todos_Click()
    
    Dim i As Integer
    
    For i = 0 To frm_z0_sele.txt_lista_izq.ListCount
        frm_z0_sele.txt_lista_izq.AddItem frm_z0_sele.txt_lista_der.List(i)
        frm_z0_sele.txt_lista_der.RemoveItem i
    Next i

End Sub


