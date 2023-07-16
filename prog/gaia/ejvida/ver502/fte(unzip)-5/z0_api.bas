Attribute VB_Name = "bas_z0_api"
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


'Para hacer mi_doevents
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Sleep 0&

'De vb50
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


Type MEMORYSTATUS
    dwLength As Long    ' 32
    dwMemoryLoad As Long    ' percent of memory in use
    dwTotalPhys As Long ' bytes of physical memory
    dwAvailPhys As Long ' free physical memory bytes
    dwTotalPageFile As Long ' bytes of paging file
    dwAvailPageFile As Long ' free bytes of paging file
    dwTotalVirtual As Long  ' user bytes of address space
    dwAvailVirtual As Long  ' free user bytes
End Type

Type SYSTEMTIME
    wYear As Long
    wMonth As Long
    wDayOfWeek As Long
    wDay As Long
    wHour As Long
    wMinute As Long
    wSecond As Long
    wMilliseconds As Long
End Type

Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpmstMemStat As MEMORYSTATUS)
Declare Function GetFocus& Lib "user32" ()
Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal message&, ByVal wParam&, lParam As Any)
Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

#If Win32 Then
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetVersion Lib "kernel32" () As Long
#Else
Declare Function GetVersion& Lib "Kernel" ()
Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
#End If

Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long


