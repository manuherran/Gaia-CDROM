VERSION 5.00
Begin VB.Form frm_u0_bk 
   Caption         =   "Backup"
   ClientHeight    =   7035
   ClientLeft      =   2295
   ClientTop       =   2070
   ClientWidth     =   5355
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7035
   ScaleWidth      =   5355
   Begin VB.CommandButton CerrarTodo 
      Caption         =   "Cerrar Todo"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Carpetas 
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
      Left            =   1920
      TabIndex        =   7
      Text            =   "c:\......"
      Top             =   5520
      Width           =   3375
   End
   Begin VB.TextBox Backup 
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
      Left            =   1920
      TabIndex        =   6
      Text            =   "c:\backup\"
      Top             =   6240
      Width           =   3375
   End
   Begin VB.TextBox Basura 
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
      Left            =   1920
      TabIndex        =   5
      Text            =   "c:\basura\"
      Top             =   6600
      Width           =   3375
   End
   Begin VB.TextBox Compresor 
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
      Left            =   1920
      TabIndex        =   4
      Text            =   "c:\arj.exe"
      Top             =   5880
      Width           =   3375
   End
   Begin VB.CommandButton DiscoDuro 
      Caption         =   "&Disco Duro"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Diskette 
      Caption         =   "Dis&kette"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5010
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton Salir 
      Caption         =   "&Salir"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carpeta de Basura"
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
      Left            =   270
      TabIndex        =   11
      Top             =   6720
      Width           =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carpeta destino"
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
      TabIndex        =   10
      Top             =   6360
      Width           =   1110
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Compresor"
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
      Left            =   840
      TabIndex        =   9
      Top             =   6000
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carpetas a Comprimir"
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
      Left            =   165
      TabIndex        =   8
      Top             =   5640
      Width           =   1485
   End
End
Attribute VB_Name = "frm_u0_bk"
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

Dim carpeta_basura_bk As String
Dim numero_directorios_bk As Long
Dim fichero_bat_bk As String
Dim fichero_carpetas_bk As String
Dim carpeta_ficheros_backup_bk As String
Dim extension_ficheros_bk As String
Dim compresor_completo_bk As String

Private Sub CerrarTodo_Click()
    
    s_fin_todo

End Sub

Private Sub DiscoDuro_Click()
    Dim i As Long
    Dim X As Variant
    Dim cadena As String

    Open fichero_bat_bk For Output As #CTE_FIC_09_BK_BAT

    For i = 0 To numero_directorios_bk - 1
        If List1.Selected(i) Then
            Print #CTE_FIC_09_BK_BAT,
            Print #CTE_FIC_09_BK_BAT, "c:"
            Print #CTE_FIC_09_BK_BAT, "cls"
            Print #CTE_FIC_09_BK_BAT, "cd c:\"
            cadena = "IF EXIST " + carpeta_ficheros_backup_bk + nombres_ficheros_backup_bk(i + 1) + extension_ficheros_bk + "  DEL " + carpeta_ficheros_backup_bk + nombres_ficheros_backup_bk(i + 1) + extension_ficheros_bk
            Print #CTE_FIC_09_BK_BAT, cadena
            cadena = compresor_completo_bk + " A -RVA " + carpeta_ficheros_backup_bk + nombres_ficheros_backup_bk(i + 1) + extension_ficheros_bk + " " + directorio_a_comp_bk(i + 1)
            Print #CTE_FIC_09_BK_BAT, cadena
            Print #CTE_FIC_09_BK_BAT,
            'ejemplo:
            'c:\arj.exe A -RVA c:\backup\bea.arj c:\bea\
        End If
    Next i
    Close #CTE_FIC_09_BK_BAT
    X = Shell(fichero_bat_bk, 1)
    


End Sub

Private Sub Diskette_Click()
    Dim i As Long
    Dim X As Variant
    Dim cadena As String

    
    Open fichero_bat_bk For Output As #CTE_FIC_09_BK_BAT

    
    For i = 0 To numero_directorios_bk - 1
        If List1.Selected(i) Then
            Print #CTE_FIC_09_BK_BAT,
            Print #CTE_FIC_09_BK_BAT, "c:"
            Print #CTE_FIC_09_BK_BAT, "cls"
            Print #CTE_FIC_09_BK_BAT, "cd c:\"
            cadena = "IF EXIST a:\" + nombres_ficheros_backup_bk(i + 1) + extension_ficheros_bk + " DEL a:\" + nombres_ficheros_backup_bk(i + 1) + extension_ficheros_bk
            Print #CTE_FIC_09_BK_BAT, cadena
            'no sirve pq todavia no te ha pedido el diskette 2
            'For x = 1 To 9
            '    cadena = "IF EXIST a:\" + nombres_ficheros_backup_bk(i + 1) + "a0" + Trim(CStr(x)) + " DEL a:\" + nombres_ficheros_backup_bk(i + 1) + "a0" + Trim(CStr(x))
            '    Print #CTE_FIC_09_BK_BAT, cadena
            'Next x
            cadena = compresor_completo_bk + " A -RVA a:\" + nombres_ficheros_backup_bk(i + 1) + extension_ficheros_bk + " " + directorio_a_comp_bk(i + 1)
            Print #CTE_FIC_09_BK_BAT, cadena
            Print #CTE_FIC_09_BK_BAT,
            'ejemplo:
            'c:\arj.exe A -RVA a:\bea.arj c:\bea\
            'DoEvents
            'PerderTiempo
            'DoEvents
        End If
    Next i
    Close #CTE_FIC_09_BK_BAT
    X = Shell(fichero_bat_bk, 1)
    

End Sub

Private Sub Form_Load()

    Dim i As Integer

    'Mostramos la pantalla en el centro
     s_centrar_ventana_ejv Me

    extension_ficheros_bk = "arj"
    If f_existe_fichero(Compresor.Text) Then
        compresor_completo_bk = Compresor.Text
    Else
        compresor_completo_bk = "c:\arj.exe"
    End If
    If Basura.Text <> "" Then
        carpeta_basura_bk = Basura.Text
    Else
        carpeta_basura_bk = "c:\basura\"
    End If
    fichero_bat_bk = f_nombre_completo(path_largo_ejv(CTE_C_PRG_UTIL), "$fwbktmp.bat")
    If Backup.Text <> "" Then
        carpeta_ficheros_backup_bk = Backup.Text
    Else
        carpeta_ficheros_backup_bk = "c:\backup\"
    End If
    
    fichero_carpetas_bk = f_nombre_completo(path_largo_ejv(CTE_C_PRG_UTIL), "backup.txt")
    Carpetas.Text = fichero_carpetas_bk
    
    nombre_fichero_ejv = fichero_carpetas_bk
    s_aut_leer_bkup
    
    numero_directorios_bk = UBound(directorio_a_comp_bk)
    
    For i = 1 To numero_directorios_bk
        List1.AddItem directorio_a_comp_bk(i)
    Next i

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FillColor = QBColor(Rnd * 15)   ' Choose random FillColor.
    FillStyle = Int(Rnd * 8)    ' Choose random FillStyle.
    Circle (X, Y), 250  ' Draw a circle.
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Kill fichero_bat_bk

End Sub

Private Sub Salir_Click()
    
    Unload Me

End Sub
