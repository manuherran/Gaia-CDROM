VERSION 5.00
Begin VB.Form frm_a4_tipospri 
   Caption         =   "Tipos de Jugadores"
   ClientHeight    =   6000
   ClientLeft      =   630
   ClientTop       =   1485
   ClientWidth     =   8970
   Icon            =   "a4_tipri.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   8970
   Begin VB.CommandButton H_Abrir 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1320
      Picture         =   "a4_tipri.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Abrir Fichero"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Compilar 
      Caption         =   "Com&pilar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_tipos 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   8775
   End
   Begin VB.Label FicJugActual 
      Caption         =   "FicJugActual"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frm_a4_tipospri"
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
    
    Dim cont_mostrar As Integer
    Dim cont_real As Integer
    Dim pos As Integer
    Dim tmp As String
    Dim linea As String
            
    'me encargo de que este grabado,
    'Si se ha modificado, y no esta grabado, lo grabo
    'lo leo y lo cargo para que sea el actual
    If fich_jug_modificado_pri Then
        If MsgBox("¿Desea guardar los cambios en un fichero?", vbQuestion + vbYesNo) = vbYes Then
            Guardar_Click
        End If
    End If
    'Guardo estas lineas en el array
    ReDim fichero_tipos_jugadores_pri(1 To 1) As String
    ReDim tipos_jugadores_mostrar_pri(1 To 1) As String
    tmp = txt_tipos.Text
    cont_mostrar = 0
    cont_real = 0
    While Len(tmp) > 0
        pos = InStr(tmp, vbCrLf)
        If pos = 0 Then
            linea = tmp
            tmp = ""
        Else
            linea = Left(tmp, pos - 1)
            tmp = Right(tmp, Len(tmp) - pos - 1)
            pos = InStr(linea, Chr$(10))
            If pos = 1 Then
                linea = Right(linea, Len(linea) - 1)
            End If
        End If
        cont_mostrar = cont_mostrar + 1
        ReDim Preserve tipos_jugadores_mostrar_pri(1 To cont_mostrar) As String
        tipos_jugadores_mostrar_pri(cont_mostrar) = linea
        
        linea = Trim(linea)
        If Len(linea) > 0 Then
            cont_real = cont_real + 1
            ReDim Preserve fichero_tipos_jugadores_pri(1 To cont_real) As String
            fichero_tipos_jugadores_pri(cont_real) = linea
        End If
    Wend
    
    s_analisis_sintactico_tipos_jugadores_pri
    Unload Me
    
    
End Sub
Private Sub Cancelar_Click()
    Unload Me

End Sub

Private Sub Form_Activate()
    'Permito guardar y abrir
    s_cambiar_estado_enabled_operaciones_ficheros_ejv True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    s_tecla_pulsada_ejv KeyCode, Shift

End Sub

Private Sub Form_Load()

    
    Screen.MousePointer = CTE_ARENA
    's_centrar_ventana_ejv Me
    'Me.WindowState = CTE_NORMAL 'no funciona ninguno de los dos por ser mdichild!!!!
    
    'Muestro el nuevo path-fichero
    frm_a4_tipospri.Caption = "Tipos de Jugadores - " & nombre_fichero_jugadores_pri
    frm_a4_tipospri.FicJugActual.Caption = nombre_fichero_jugadores_pri
    s_mostrar_fichero_tipos_jugadores_pri
    fich_jug_modificado_pri = False
    Screen.MousePointer = CTE_DEFECTO
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Ahora no permito guardar y abrir
    s_cambiar_estado_enabled_operaciones_ficheros_ejv False

End Sub

Private Sub Guardar_Click()
    s_usr_guardar_jugadores_pri
End Sub

Private Sub H_Abrir_Click()
    
    s_accion_ficheros_va0 CTE_FIC_ABRIR

End Sub

Private Sub txt_tipos_Change()
    
    If habilitar_change_pri Then
        fich_jug_modificado_pri = True
    End If

End Sub
Private Sub Abrir_Click()
    s_usr_abrir_jugadores_pri
End Sub

