VERSION 5.00
Begin VB.Form frm_z0_op 
   Caption         =   "Opciones Generales de ""Ejemplos de Vida"""
   ClientHeight    =   7530
   ClientLeft      =   1155
   ClientTop       =   975
   ClientWidth     =   9645
   Icon            =   "z0_op.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7530
   ScaleWidth      =   9645
   Begin VB.CommandButton Command1 
      Caption         =   "&Editar Fichero de Inicio"
      Height          =   375
      Left            =   4800
      TabIndex        =   48
      Top             =   360
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   2355
      TabIndex        =   46
      Top             =   360
      Width           =   2415
   End
   Begin VB.CheckBox Ch_EliminarColor 
      Caption         =   "Eliminar el color de fondo de la lista de colores posibles"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   720
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Función de Azar"
      Height          =   2055
      Left            =   120
      TabIndex        =   26
      Top             =   1080
      Width           =   9375
      Begin VB.TextBox Op_AzarFicNumCh 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   35
         Text            =   "49999"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox Ch_Randomize 
         Caption         =   "Con Randomize"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Fic_Azar 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox Op_AzarFicC 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   30
         Text            =   "c:\......"
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox Op_AzarFicF 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7800
         TabIndex        =   29
         Text            =   "pi49999.ran"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton Op_AzarFic 
         Caption         =   "Números tomados de un fichero de texto, por ejemplo, pi49999.ran que contiene los 49999 primeros decimales"
         Height          =   315
         Left            =   360
         TabIndex        =   28
         Top             =   840
         Width           =   8775
      End
      Begin VB.OptionButton Op_AzarVB 
         Caption         =   "RND de Visual Basic"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "caracteres del fichero."
         Height          =   255
         Left            =   3720
         TabIndex        =   36
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Usar únicamente los primeros"
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "del número Pi"
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Condiciónes de Parada"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   9375
      Begin VB.TextBox Op_CondParadaPeso 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   16
         Text            =   "100"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Ch_CondParadaPeso 
         Caption         =   "Cuando haya algun agente con peso (energía) >="
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Op_CondParadaHora 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         Text            =   "23:59:59"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Op_CondParadaFecha 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Text            =   "03/05/98"
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Ch_CondParadaFechaHora 
         Caption         =   "En la fecha/hora"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Op_CondParadaNumMaxCiclos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "5000"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox Ch_CondParadaNumMaxCiclos 
         Caption         =   "En el ciclo"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Grabar Resumen"
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   9375
      Begin VB.CheckBox Op_Reemplazar 
         Caption         =   "Reemplazar los ficheros si ya existen"
         Height          =   255
         Left            =   6120
         TabIndex        =   47
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox Op_Cabeceras 
         Caption         =   "Generar cabeceras en los ficheros"
         Height          =   255
         Left            =   6120
         TabIndex        =   45
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Op_MaxGuardado 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3960
         TabIndex        =   43
         Text            =   "5000"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Op_Autoguardado 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         TabIndex        =   40
         Text            =   "5000"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox Ch_GrabarResumenGra 
         Caption         =   "Grabar resumen Gráfico"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton Fic_Gra 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Op_GrabarResumenGraC 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   23
         Text            =   "c:\......"
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox Op_GrabarResumenGraF 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7800
         TabIndex        =   22
         Text            =   "resumen.gra"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Op_GrabarResumenTxtF 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7800
         TabIndex        =   21
         Text            =   "resumen.txt"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Op_GrabarResumenExcelF 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7800
         TabIndex        =   20
         Text            =   "resumen.xls"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Op_GrabarResumenTxtC 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   19
         Text            =   "c:\......"
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox Op_GrabarResumenExcelC 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Text            =   "c:\......"
         Top             =   1320
         Width           =   4695
      End
      Begin VB.CommandButton Fic_Txt 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Fic_Excel 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Ch_GrabarResumenTxt 
         Caption         =   "Grabar resumen en txt"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox Ch_GrabarResumenExcel 
         Caption         =   "Grabar resumen en Excel"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "ciclos"
         Height          =   255
         Left            =   5400
         TabIndex        =   44
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Como máximo, guardar información de los primeros "
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "ciclos"
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Autoguardado cada"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Fichero"
         Height          =   255
         Left            =   7800
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Carpeta"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox Cb_Color 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Color de Fondo"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frm_z0_op"
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
    
    s_grabar_opciones_generales_ejv
    s_activar_opciones_generales_ejv
    Unload Me


End Sub

Private Sub Cancelar_Click()
    Unload Me

End Sub

Private Sub Cb_Color_Click()
    
    Picture1.BackColor = cct_ejv(frm_z0_op.Cb_Color.ListIndex + 1)

End Sub

Private Sub Command1_Click()
     
     frm_z0_inic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show

End Sub

Private Sub Fic_Azar_Click()
    
    'Fijo una carpeta por defecto
    If Len(Trim(Op_AzarFicC.Text)) > 0 Then
        nombre_fichero_ejv = Trim(Op_AzarFicC.Text)
    Else
        nombre_fichero_ejv = path_largo_ejv(CTE_C_ENT_RAN)
    End If
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta o carpeta mas fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.ran"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Op_AzarFicC.Text = nombre_fichero_ejv
    Else
        Op_AzarFicC.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_ENT_RAN)
    End If
    
    'Si ha tecleado un nombre de fichero, lo muestro en su sitio
    If Len(f_nombre_fichero(nombre_fichero_ejv)) > 0 Then
        Op_AzarFicF.Text = f_nombre_fichero(nombre_fichero_ejv)
    End If
   
    
End Sub
    
Private Sub Fic_Excel_Click()
    
    'Fijo una carpeta por defecto
    If Len(Trim(Op_GrabarResumenExcelC.Text)) > 0 Then
        nombre_fichero_ejv = Trim(Op_GrabarResumenExcelC.Text)
    Else
        nombre_fichero_ejv = path_largo_ejv(CTE_C_SAL_XLS)
    End If
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta o carpeta mas fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.xls"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Op_GrabarResumenExcelC.Text = nombre_fichero_ejv
    Else
        Op_GrabarResumenExcelC.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_SAL_XLS)
    End If

    'Si ha tecleado un nombre de fichero, lo muestro en su sitio
    If Len(f_nombre_fichero(nombre_fichero_ejv)) > 0 Then
        Op_GrabarResumenExcelF = f_nombre_fichero(nombre_fichero_ejv)
    End If
    

End Sub

Private Sub Fic_Gra_Click()
    
    'Fijo una carpeta por defecto
    If Len(Trim(Op_GrabarResumenGraC.Text)) > 0 Then
        nombre_fichero_ejv = Trim(Op_GrabarResumenGraC.Text)
    Else
        nombre_fichero_ejv = path_largo_ejv(CTE_C_SAL_GRA)
    End If
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta o carpeta mas fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.gra"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Op_GrabarResumenGraC.Text = nombre_fichero_ejv
    Else
        Op_GrabarResumenGraC.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_SAL_GRA)
    End If

    'Si ha tecleado un nombre de fichero, lo muestro en su sitio
    If Len(f_nombre_fichero(nombre_fichero_ejv)) > 0 Then
        Op_GrabarResumenGraF = f_nombre_fichero(nombre_fichero_ejv)
    End If

End Sub

Private Sub Fic_Txt_Click()
    
    'Fijo una carpeta por defecto
    If Len(Trim(Op_GrabarResumenTxtC.Text)) > 0 Then
        nombre_fichero_ejv = Trim(Op_GrabarResumenTxtC.Text)
    Else
        nombre_fichero_ejv = path_largo_ejv(CTE_C_SAL_TXT)
    End If
    nombre_fichero_ejv_es_solo_un_path_ejv = True
    'Elijo carpeta o carpeta mas fichero
    tipo_operacion_formulario_fic_ejv = CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH
    frm_z0_fic.Caption = "Seleccionar Carpeta" 'Esto provoca la llamada, igual que un show
    frm_z0_fic.Aceptar.Caption = "&Seleccionar"
    frm_z0_fic.File1.Pattern = "*.txt"
    frm_z0_fic.tipo = frm_z0_fic.File1.Pattern
    frm_z0_fic.Show CTE_MODAL 'Un modal requiere cerrar la ventana para continuar con el código que viene despues del show
    If cancelar_operacion_fichero_ejv Then Exit Sub
    
    'Muestro el nuevo path
    If nombre_fichero_ejv_es_solo_un_path_ejv Then
        Op_GrabarResumenTxtC.Text = nombre_fichero_ejv
    Else
        Op_GrabarResumenTxtC.Text = f_path_fichero(nombre_fichero_ejv, CTE_C_SAL_TXT)
    End If

    'Si ha tecleado un nombre de fichero, lo muestro en su sitio
    If Len(f_nombre_fichero(nombre_fichero_ejv)) > 0 Then
        Op_GrabarResumenTxtF = f_nombre_fichero(nombre_fichero_ejv)
    End If

End Sub

Private Sub Form_Load()
    
    'Tengo que recargar el array, ya que puede haberse eliminado el
    'color de fondo actual de la lista
    s_inicializar_arrays_color_ejv
    s_cargar_opciones_generales_ejv

End Sub

Private Sub Op_AzarFic_Click()
    s_enabled_azar
End Sub

Private Sub Op_AzarVB_Click()
    s_enabled_azar
End Sub

Sub s_enabled_azar()

    If Op_AzarVB Then
        Ch_Randomize.Enabled = True
        frm_z0_op.Op_AzarFicC.Enabled = False
        frm_z0_op.Op_AzarFicF.Enabled = False
        frm_z0_op.Op_AzarFicNumCh.Enabled = False
        frm_z0_op.Fic_Azar.Enabled = False
    Else
        Ch_Randomize.Enabled = False
        frm_z0_op.Op_AzarFicC.Enabled = True
        frm_z0_op.Op_AzarFicF.Enabled = True
        frm_z0_op.Op_AzarFicNumCh.Enabled = True
        frm_z0_op.Fic_Azar.Enabled = True
    End If

End Sub
