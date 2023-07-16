VERSION 5.00
Begin VB.Form frm_a0_va 
   Caption         =   "Vida Artificial"
   ClientHeight    =   8475
   ClientLeft      =   810
   ClientTop       =   405
   ClientWidth     =   10395
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
   Icon            =   "A0_VA.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8475
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin VB.Image Imagen 
      Height          =   810
      Left            =   600
      Picture         =   "A0_VA.frx":030A
      Top             =   480
      Visible         =   0   'False
      Width           =   525
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
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frm_a0_va"
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

Private Sub Form_Activate()
    
    's_identificar_num_prg_activo_ejv
    'Ahora no permito guardar y abrir
    s_cambiar_estado_enabled_operaciones_ficheros_ejv False
    habilitar_change_zoom_va0 = True
    
    's_identificar_num_prg_activo_ejv
    'Actualizo el estado de enabled de ejecucion y ver
    'cogiendolo de los arrays
    s_estado_enabled_ejecucion_ejv
    s_estado_enabled_ver_ejv

End Sub

Private Sub Form_GotFocus()
    
    Me.BackColor = cct_ejv(cfondo_ejv)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    s_tecla_pulsada_ejv KeyCode, Shift
    
End Sub

Private Sub Form_Load()
        
    'Permito recibir teclas
    Me.KeyPreview = True
    
    'Refresco automatico, consume muchos recursos
    'frm_a0_va.AutoRedraw = True
    frm_a0_va.AutoRedraw = False
    
    'Mostramos la pantalla en el centro del monitor
    Me.Height = 7200
    Me.Width = 11000
    s_centrar_ventana_ejv Me
    
    'Imagen
    s_mostrar_aviso_imagen
    ciclo_ejv = 0
    s_tratamiento_idioma_va0
    esta_detenido_ejv = True
    s_botones_enabled_va0 (True)
    s_cambiar_estado_enabled_menus_ejv CTE_VER_AGENTES_TODOS, False
    s_cambiar_estado_enabled_ejecutar_ejv CTE_EXE_CONTINUAR, False
    esta_modificado_num_agen_tipo_pri = False
    If Not automatico_ejv Then
        s_inicializar_ejemplo_elegido_ejv
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = f_control_cerrar_va0

End Sub
Private Sub Form_Unload(Cancel As Integer)
     
    Dim prg As Integer
    
    habilitar_change_zoom_va0 = False
    
    'Si ya estoy abriendo otro, entonces son distintos
    'y no hay que preguntar por num_prg_activo_ejv sino por el anterior
    'porque es seguro que el que hay que cerrar es el anterior
    If num_prg_activo_ejv = num_prg_anterior_activo_ejv Or num_prg_anterior_activo_ejv = 0 Then
        prg = num_prg_activo_ejv
    Else
        prg = num_prg_anterior_activo_ejv
    End If
     
     
    Select Case num_prg_activo_ejv
        Case CTE_HYP '1
            Unload frm_a1_inhyp
            Unload frm_a1_tiposhyp 'Estas se cargan aunque no se vean!
            Unload frm_a1_ophyp
        Case CTE_PAL '2
            Unload frm_b2_inpal
        Case CTE_3R '3
            Unload frm_c3_in3r
        Case CTE_PRI '4
            Unload frm_a4_inpri
            Unload frm_a4_tipospri
        Case CTE_CEL '5
            Unload frm_a5_incel
        Case CTE_GAI '6
            Fi_Cerrar_Base_Datos
            Unload frm_a6_ingaia
        Case CTE_EXP '7
            Unload frm_a7_inexp
        Case CTE_CAD '8
            Unload frm_c8_incad
        Case CTE_PEZ '9
            Unload frm_a9_inpez
        Case CTE_UVA '10
            Unload frm_aA_inuva
        Case CTE_YXY '11
        Case Else
            s_error_num_prog num_prg_activo_ejv
    End Select
     
     
    'Pongo habilitado todos los programas
    s_cambiar_estado_enabled_programas_todos_ejv True
     
    'Actualizo el estado de enabled de ejecucion y ver
    num_prg_activo_ejv = CTE_NINGUNO
    'cogiendolo de los arrays del num_prg_activo_ejv
    s_estado_enabled_ejecucion_ejv
    s_estado_enabled_ver_ejv

    'Grabo los ficheros sin cerrarlos, por si hay un corte de luz y esas cosas
    s_grabar_fichero_salida_ejv CTE_FIC_20_GLOLOG
    s_grabar_fichero_salida_ejv CTE_FIC_21_GLOTXT
    s_grabar_fichero_salida_ejv CTE_FIC_22_GLOXLS
    

End Sub



Sub s_tratamiento_idioma_va0()
    If idioma_ejv = CTE_INGLES Then
    Else
    End If
    
End Sub
