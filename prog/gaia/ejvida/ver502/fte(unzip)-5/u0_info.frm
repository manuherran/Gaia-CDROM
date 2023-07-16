VERSION 5.00
Begin VB.Form frm_u0_info 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5340
   ClientLeft      =   2250
   ClientTop       =   1785
   ClientWidth     =   9075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5340
   ScaleWidth      =   9075
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
      Height          =   4545
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   5895
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
      Left            =   3960
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "frm_u0_info"
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
    
    Unload Me

End Sub

Private Sub Form_Activate()
    
    s_mostrar_info

End Sub

Private Sub Form_Load()
    
    s_centrar_ventana_ejv Me


End Sub

Sub s_mostrar_info()


    frm_u0_info.List1.Clear
    frm_u0_info.List2.Clear

    'Memoria
    Dim ms As MEMORYSTATUS
    ms.dwLength = Len(ms)
    GlobalMemoryStatus ms
    
    
    frm_u0_info.List1.AddItem "Memoria RAM total"
    frm_u0_info.List2.AddItem Format(ms.dwTotalPhys, "#,0") & " bytes = " & Format(ms.dwTotalPhys / 1024, "#,0") & " K = " & Format(ms.dwTotalPhys / 1024 / 1024, "#,0") & " Megas"
    
    frm_u0_info.List1.AddItem "Memoria RAM libre"
    frm_u0_info.List2.AddItem Format(ms.dwAvailPhys, "#,0") & " bytes = " & Format(ms.dwAvailPhys / 1024, "#,0") & " K = " & Format(ms.dwAvailPhys / 1024 / 1024, "#,0") & " Megas"
    
    frm_u0_info.List1.AddItem "% de RAM libre"
    frm_u0_info.List2.AddItem Format(ms.dwAvailPhys / ms.dwTotalPhys, "Percent")
    

   
    'Disco duro
    Dim tl&
    Dim S$
    Dim spaceloc%
    Dim SectorsPerCluster&, BytesPerSector&, NumberOfFreeClustors&, TotalNumberOfClustors&
    Dim BytesFree&, BytesTotal&
    Dim PercentFree&
    Dim TotalBytes, FreeBytes As Long
    S$ = "c:\"
    's$ = Drive1.Drive
    ' Is there a space? Strip off the volume name if so
    spaceloc = InStr(S$, " ")
    If spaceloc > 0 Then
        S$ = Left$(S$, spaceloc - 1)
    End If
    If Right$(S$, 1) <> "\" Then S$ = S$ & "\"
    tl& = GetDiskFreeSpace(S$, SectorsPerCluster, BytesPerSector, NumberOfFreeClustors, TotalNumberOfClustors)
    TotalBytes = TotalNumberOfClustors * SectorsPerCluster * BytesPerSector
    
    frm_u0_info.List1.AddItem "Espacio total en disco " & S$
    frm_u0_info.List2.AddItem Format(TotalBytes, "#,0") & " bytes = " & Format(TotalBytes / 1024, "#,0") & " K = " & Format(TotalBytes / 1024 / 1024, "#,0") & " Megas = " & Format(TotalBytes / 1024 / 1024 / 1024, "#,0") & " Gigas"
    
    FreeBytes = NumberOfFreeClustors * SectorsPerCluster * BytesPerSector
    
    frm_u0_info.List1.AddItem "Espacio libre en disco " & S$
    frm_u0_info.List2.AddItem Format(FreeBytes, "#,0") & " bytes = " & Format(FreeBytes / 1024, "#,0") & " K = " & Format(FreeBytes / 1024 / 1024, "#,0") & " Megas"
    
    frm_u0_info.List1.AddItem "% de espacio libre en disco " & S$
    frm_u0_info.List2.AddItem Format(FreeBytes / TotalBytes, "Percent")
    
    
    'Windows y fuentes
    Dim d$, r&
    d$ = String$(255, 0)
    r& = GetWindowsDirectory(d$, 254)
    d$ = Left$(d$, r)
    
    frm_u0_info.List1.AddItem "Directorio de Windows"
    frm_u0_info.List2.AddItem d$
   
    
    If GetVersion() > 0 Then
        d$ = d$ & "\system"
    Else
        d$ = d$ & "\fonts"
    End If
    
    frm_u0_info.List1.AddItem "Directorio de Fuentes"
    frm_u0_info.List2.AddItem d$
    
    
    Print ""
    Print ""
    
    'tiempo
    Dim tim As SYSTEMTIME
    Dim tim2 As FILETIME
    Dim dl&
  
    GetSystemTime tim
    
    dl& = SystemTimeToFileTime(tim, tim2)
    
    
    frm_u0_info.List1.AddItem "Fecha"
    frm_u0_info.List2.AddItem "Año: " & tim.wYear & " Mes: " & tim.wMonth & " Día de la semana: " & tim.wDayOfWeek & " Día: " & tim.wDay
    
    frm_u0_info.List1.AddItem "Hora"
    frm_u0_info.List2.AddItem "Hora: " & tim.wHour & " Minuto: " & tim.wMinute & " Segundo: " & tim.wSecond & " Milisegundos: " & tim.wMilliseconds
    
    
    frm_u0_info.List1.AddItem "Low-High"
    frm_u0_info.List2.AddItem "Low: " & tim2.dwLowDateTime & " High: " & tim2.dwHighDateTime
    


End Sub
