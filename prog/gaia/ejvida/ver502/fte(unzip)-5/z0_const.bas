Attribute VB_Name = "bas_z0_const"
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


'============================================
'Tipos usados:
'integer -> entero
'long    -> entero grande
'double  -> con decimales grande
'============================================
'Check:0,1
'Option:true,false
'============================================


'----------------------------------------
'Abrir y leer fichero excel de entrada
'    If Not f_existe_fichero(s_path_fichero_xls) Then
'        MsgBox "error"
'    End If
'Set o_hoja_excel = Nothing
'Set o_hoja_excel = GetObject(s_path_fichero_xls)
'variable = o_hoja_excel.ActiveSheet.Cells(fila, columna).Value
'o_hoja_excel.Close
'----------------------------------------


'----------------------------------------
'Abrir y escribir fichero excel de salida
'    If f_existe_fichero(s_path_fichero_xls) Then
'        'El fichero ya existe, va a ser reemplazado
'    End If
'Set o_hoja_excel = Nothing
'Set o_hoja_excel = CreateObject("Excel.Sheet")
'o_hoja_excel.ActiveSheet.Cells(fila, columna).Value = variable
'o_hoja_excel.SaveCopyAs s_path_fichero_xls
'o_hoja_excel.Close
'----------------------------------------



'============================================
'Al hacer esto todos los dibujos salen rellenos por defecto
'frm_a0_va.FillStyle = vbFSSolid 'solid
'Para dibujar algo sin relleno se pone y se quita
'FillStyle = vbFSTransparent 'transparent
'aqui dibujo
'FillStyle = vbFSSolid 'solid
'============================================
'Tipos no usados:
'single   -> con decimales pequeño
'============================================

'Dim NewForm as New Form1


'============================================
'Shell "command c/ c:\arj.exe -va", vbHide

'    Dim MyAppID, ReturnValue
'    AppActivate "Microsoft Word"    ' Activate Microsoft Word
    
'    Dim MyAppID, ReturnValue
'    AppActivate "Microsoft Excel"    ' Activate Microsoft Excel
    
    ' AppActivate can also use the return value of the Shell function.
'    MyAppID = Shell("C:\WORD\WINWORD.EXE", 1)   ' Run Microsoft Word.
'    AppActivate MyAppID ' Activate Microsoft Word
    
    ' You can also use the return value of the Shell function.
'    ReturnValue = Shell("c:\EXCEL\EXCEL.EXE", 1) ' Run Microsoft Excel.
'    AppActivate ReturnValue ' Activate Microsoft Excel.

    'ShellExecute "notepad.exe " & fichero_aut_ejv(parametro)
        'HojaExcel.SaveCopyAs filename:=f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS),"resultados.xls"), fileFormat:="xlnormal", Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        'HojaExcel.SaveCopyAs f_nombre_completo(path_largo_ejv(CTE_C_SAL_XLS),"resultados.xls"), "xlnormal", "", "", False, False

'============================================


'============================================
'Los arrays de VB pueden crecer o disminuir solo por la ultima dimension,
'pero aun asi, el valor de la ultima dimension debe ser el mismo para todos
'por eso, si se hace un redim preserve se deberia hacer poniendo como valor
'de la ultima dimension el maximo valor posible, aunque otros valores esten
'vacios
'me parece que un redim preserve a un valor menor p ej (2,1) cuando existen otros
'mayores (1,2) no se ejecuta
'============================================
    
Global Const CTE_PI = 3.14159265358979

Global Const CTE_1VUELTA = 6.28318530717958
Global Const CTE_MEDIAVUELTA = 3.14159265358979
Global Const CTE_CUARTODEVUELTA = 1.5707963267949

Global Const CTE_Max_Integer = 32767
Global Const CTE_Max_Long = 2147483647
Global Const CTE_Min_Long = -2147483647
Global Const MAX_LISTA = 32000

'Nombre de la Aplicacion
Global Const CTE_LOGO_APLICACION = "Ejemplos de Vida"

'Programas y num_prg_activo_ejv
Global Const CTE_PROG_num_total = 12
Global Const CTE_NINGUNO = 0
Global Const CTE_HYP = 1 'Hormigas y Plantas
Global Const CTE_PAL = 2 'Palabras y Frases
Global Const CTE_3R = 3  'Tres en Raya
Global Const CTE_PRI = 4 'Prisionero
Global Const CTE_CEL = 5 'Celdilla
Global Const CTE_GAI = 6 'Plataforma Gaia
Global Const CTE_EXP = 7 'Explorando Mapas
Global Const CTE_CAD = 8 'Cadenas
Global Const CTE_PEZ = 9 'Peces
Global Const CTE_UVA = 10 'Universo
Global Const CTE_YXY = 11 'YXY
Global Const CTE_ESPECIAL = 12


'Carpetas
Global Const CTE_C_RAIZ = 1
Global Const CTE_C_PRG_3R = 2
Global Const CTE_C_PRG_AUT = 3
Global Const CTE_C_PRG_BMP = 4
Global Const CTE_C_ENT_DIC = 5
Global Const CTE_C_DOC = 6
Global Const CTE_C_SAL_GRA = 7
Global Const CTE_C_ENT = 8
Global Const CTE_C_PRG_HYP = 9
Global Const CTE_C_PRG_ICO = 10
Global Const CTE_C_SAL_LOG = 11
Global Const CTE_C_PRG_MAP = 12
Global Const CTE_C_DOC_WEB = 13
Global Const CTE_C_PRG_PRI = 14
Global Const CTE_C_ENT_RAN = 15
Global Const CTE_C_SAL_TXT = 16
Global Const CTE_C_PRG_UTIL = 17
Global Const CTE_C_SAL_XLS = 18
Global Const CTE_C_PRG = 19

Global Const CTE_TOTAL_CARPETAS = 19




'El mayor numero de graficos de cualquier programa
Global Const CTE_GRA_max_graficos = 6


'El numero de graficos no siempre coincide con el numero de datos
'de los que se guarda resumen. Aunque mas o menos la idea es
'mostrar un grafico por cada variable de la que se guarda
'resumen, a veces el numero solo se sabe en tiempo de ejecución
Global Const CTE_GRA_HYP_num = 6
Global Const CTE_GRA_PAL_num = 2
Global Const CTE_GRA_3R_num = 6
Global Const CTE_GRA_ESP_num = 6
Global Const CTE_numero_maximo_rnd = 5000
Global Const CTE_numero_maximo_pi = 49999
Global Const CTE_numero_maximo_x2 = 5000
Global Const CTE_numero_maximo_x21 = 5000

Global Const CTE_ESPECIAL_1_PAL = 1
Global Const CTE_ESPECIAL_2_APE = 2
Global Const CTE_ESPECIAL_3_RND = 3
Global Const CTE_ESPECIAL_4_PI = 4
Global Const CTE_ESPECIAL_5_X2 = 5
Global Const CTE_ESPECIAL_6_X21 = 6

Global Const CTE_ESPECIALd_1_PAL = "Longitud Palabras"
Global Const CTE_ESPECIALd_2_APE = "Longitud Apellidos"
Global Const CTE_ESPECIALd_3_RND = "RND Visual Basic"
Global Const CTE_ESPECIALd_4_PI = "Pi"
Global Const CTE_ESPECIALd_5_X2 = "x[t]=1-x[t-1]^2"
Global Const CTE_ESPECIALd_6_X21 = "x[t]=(2*x[t-1]^2)-1"



'Estado de la ejecución
Global Const CTE_FUNCIONANDO = 1 'Aprendiendo
Global Const CTE_DETENIDO = 2
Global Const CTE_DETENIENDO = 3
Global Const CTE_MOSTRANDO = 4
Global Const CTE_JUGANDO = 5 'Jugando Hombre contra la Máquina

'Tipo de detección de parada por ciclo
Global Const CTE_PARADA_POR_IGUAL = 0
Global Const CTE_PARADA_POR_MAYOR = 1

'Acciones de ejecución
Global Const CTE_EXE_num_total = 4
Global Const CTE_EXE_COMENZAR = 1
Global Const CTE_EXE_CONTINUAR = 2
Global Const CTE_EXE_PAUSA = 3
Global Const CTE_EXE_TERMINAR = 4

'Acciones de Ver
Global Const CTE_VER_num_total = 20
Global Const CTE_VER_OPCIONES1 = 1
Global Const CTE_VER_OPCIONES2 = 2
Global Const CTE_VER_OPCIONES3 = 3
Global Const CTE_VER_TIPOS_AGENTES = 4
Global Const CTE_VER_MAPA = 5
Global Const CTE_VER_TIPO_EVOLUCION = 6
    Global Const CTE_VER_TIPO_EVOLUCION_EVALUACION = 7
    Global Const CTE_VER_TIPO_EVOLUCION_SELECCION = 8
    Global Const CTE_VER_TIPO_EVOLUCION_REPRODUCCION = 9
        Global Const CTE_VER_TIPO_EVOLUCION_REPRODUCCION_MUTACIONES = 10
        Global Const CTE_VER_TIPO_EVOLUCION_REPRODUCCION_SOBRECRUZAMIENTO = 11
Global Const CTE_VER_APELLIDOS = 12
Global Const CTE_VER_REFRESCAR = 13
Global Const CTE_VER_ESTADO_EJECUCION = 14
Global Const CTE_VER_AGENTES_TODOS = 15
Global Const CTE_VER_AGENTES_MEJORES = 16
Global Const CTE_VER_DICCIONARIO = 17
Global Const CTE_VER_JUGAR_CONTRA_ORDENADOR = 18
Global Const CTE_VER_MODIFICAR_AGENTE = 19
Global Const CTE_VER_GRAFICO = 20

'Colores
Global Const CTE_DEGRADADOCOLOR = -3
Global Const CTE_DEGRADADOGRIS = -2
Global Const CTE_TRANSPARENTE = -1

Global Const CTE_ROJO = 1
Global Const CTE_ROSA = 2
Global Const CTE_NARANJA = 3
Global Const CTE_AMARILLO = 4
Global Const CTE_VERDEBRILLANTE = 5
Global Const CTE_VERDECLARO = 6
Global Const CTE_VERDEPALIDO = 7
Global Const CTE_AZUL = 8
Global Const CTE_NEGRO = 9
Global Const CTE_BLANCO = 10
Global Const CTE_GRISCLARO = 11
Global Const CTE_GRISOSCURO = 12
Global Const CTE_AZULCLARO = 13
Global Const CTE_AZULON = 14
Global Const CTE_VERDEOSCURO = 15
Global Const CTE_SALMON = 16
Global Const CTE_ROSACLARO = 17
Global Const CTE_MORADO = 18
Global Const CTE_AZULOSCURO = 19
Global Const CTE_MARRONCLARO = 20
Global Const CTE_NARANJAOSCURO = 21
Global Const CTE_VINOTINTO = 22
Global Const CTE_VERDEMANZANA = 23
Global Const CTE_VERDEPLASTICO = 24
Global Const CTE_AZULGRISACEO = 25
Global Const CTE_ROJOTIERRA = 26
Global Const CTE_VERDEAZULON = 27
Global Const CTE_HUESO = 28
Global Const CTE_ROSACEO = 29
Global Const CTE_ORO = 30
Global Const CTE_VIOLETA = 31
Global Const CTE_LAVANDA = 32
Global Const CTE_DORADO = 33


'Figuras
Global Const CTE_PUNTO = "punto"
Global Const CTE_CUBO = "cubo"
Global Const CTE_CUADRADOCURSOR = "cuadradocursor"
Global Const CTE_ESFERA = "esfera"
Global Const CTE_PLANTA = "planta"
Global Const CTE_PLANTALLENA = "planta_llena"
Global Const CTE_HORMIGA = "hormiga"
Global Const CTE_HORMIMUERTA = "hormimuerta"
Global Const CTE_PRISIONERO = "prisionero"
Global Const CTE_ = ""


'Numero de ejes
Global Const CTE_EJE_num_max_ejes = 6
Global Const CTE_EJE_1 = 1
Global Const CTE_EJE_2 = 2
Global Const CTE_EJE_3 = 3
Global Const CTE_EJE_4 = 4
Global Const CTE_EJE_5 = 5
Global Const CTE_EJE_6 = 6

'Lugar de Impresion
Global Const CTE_FORMULARIO = 0
Global Const CTE_IMPRESORA = 1




'Escala
Global Const CTE_GRA_AJUSTADA = 1 '"Ajustada al máximo"
Global Const CTE_GRA_REAL = 2 '"Real en Pixels"

'Pto  o linea
Global Const CTE_GRA_PTO = 1 '"Puntos"
Global Const CTE_GRA_LINEA = 2 '"Líneas"

'Estado de las ventanas
Global Const CTE_NORMAL = 0
Global Const CTE_MINIMIZED = 1
Global Const CTE_MAXIMIZED = 2

'presentación ventanas
Global Const CTE_CASCADA = 0
Global Const CTE_MOSAICO_HORIZONTAL = 1
Global Const CTE_MOSAICO_VERTICAL = 2
Global Const CTE_ORGANIZAR_ICONOS = 3

'Constantes Globales que afectan al Ratón
Global Const CTE_DEFECTO = 0
Global Const CTE_ARENA = 11
'Global Const CTE_NO_DROP = 12

'Parámetros del Show
Global Const CTE_MODAL = 1
Global Const CTE_AMODAL = 0


'Tipos de azar
Global Const CTE_AZARVB = 0
Global Const CTE_AZARFIC = 1

'Numero de decimales que se toman a partir de la serie de numeros al azar, osea
'si tengo 437846374 genero por ejemplo 0.437 si CTE_NUM_DECIMALES_EXACTITUD=3
Global Const CTE_NUM_DECIMALES_EXACTITUD = 12


Global Const CTE_INDIFERENTE = 0


'Errores
Global Const CON_OPCION_FINALIZAR = True
Global Const SIN_OPCION_FINALIZAR = False


'graficas
Global Const maximo_ancho_ejes = 700


'Metodos de sobrecruzamiento
Global Const CTE_ALTERNOS = 1
Global Const CTE_1_PTO_CORTE = 2
Global Const CTE_2_PTO_CORTE = 3


'Configuracion de todo
'Generales
Global Const CTE_txtTRUE = "True"
Global Const CTE_txtFALSE = "False"
Global Const CTE_nombreINICIO_TXT = "inicio.txt"
'Especificas
Global Const CTE_VERSION = "VERSION=" '1
Global Const CTE_IDIOMA = "IDIOMA=" '2
    Global Const CTE_INGLES = "INGLES"
    Global Const CTEm_INGLES = "Ingles"
    Global Const CTE_CASTELLANO = "CASTELLANO"
    Global Const CTEm_CASTELLANO = "Castellano"
Global Const CTE_ELEGIRIDIOMA = "ELEGIR IDIOMA=" '3
Global Const CTE_CTRLERRORES = "CONTROL DE ERRORES=" '4
Global Const CTE_MOSTRAR_LOGO = "MOSTRAR LOGO=" '5
Global Const CTE_ALGORITMODEORDENACION = "ALGORITMO DE ORDENACION=" '6
    Global Const CTE_BURBUJA = "BURBUJA"
    Global Const CTEm_BURBUJA = "Burbuja"
    Global Const CTE_QUICKSORT = "QUICKSORT"
    Global Const CTEm_QUICKSORT = "QuickSort"
Global Const CTE_SISTEMA_OPERATIVO = "SISTEMA OPERATIVO=" '7
    Global Const CTE_WINDOWS95 = "WINDOWS95"
    Global Const CTEm_WINDOWS95 = "Windows95"
    Global Const CTE_WINDOWSNT = "WINDOWSNT"
    Global Const CTEm_WINDOWSNT = "WindowsNT"
    Global Const CTE_WINDOWS3X = "WINDOWS3X"
    Global Const CTEm_WINDOWS3X = "Windows3x"
Global Const CTE_PEDIR_CONFIRMACION = "PEDIR CONFIRMACION=" '8
Global Const CTE_RESOLUCIONPANTALLA = "RESOLUCION PANTALLA=" '9
    Global Const CTE_640X480 = "640X480"
    Global Const CTEm_640X480 = "640x480"
    Global Const CTE_800X600OSUPERIOR = "800X600 O SUPERIOR"
    Global Const CTEm_800X600OSUPERIOR = "800x600 o superior"
Global Const CTE_GRABAR_CONFIGURACION = "GRABAR CONFIGURACION=" '10
Global Const CTE_GRABAR_CONFIG_POR_DEFECTO = "GRABAR CONFIG POR DEFECTO=" '11
Global Const CTE_GRABAR_LOG = "GRABAR LOG=" '12
Global Const CTE_FICHERO_LOG = "FICHERO LOG=" '13
Global Const CTE_GRABAR_RESUMEN_TXT = "GRABAR RESUMEN TXT=" '14
Global Const CTE_FICHERO_RESUMEN_TXT = "FICHERO RESUMEN TXT=" '15
Global Const CTE_GRABAR_RESUMEN_EXCEL = "GRABAR RESUMEN EXCEL=" '16
Global Const CTE_FICHERO_RESUMEN_EXCEL = "FICHERO RESUMEN EXCEL=" '17
Global Const CTE_AUTOMATICO = "AUTOMATICO=" '18
Global Const CTE_REEMPLAZAR_FICHEROS_EXISTENTES = "REEMPLAZAR FICHEROS EXISTENTES=" '19
Global Const CTE_FICHEROAUTOMATICO = "FICHERO AUTOMATICO=" '20
    Global Const CTE_NOHAY = "(NO HAY)"
    Global Const CTEm_NOHAY = "(No hay)"
    Global Const CTE_INDICE_ULTIMO_PARAMETRO_REPETIDO = 20


'Configuracion de un ejemplo automatico
Global Const CTE_AUTOMATICO_NUMERO_PROGRAMA = "AUTOMATICO NUMERO PROGRAMA="
Global Const CTE_AUTOMATICO_NUMERO_EJEMPLO = "AUTOMATICO NUMERO EJEMPLO="
Global Const CTE_FICHERO_RESULTADOS = "FICHERO RESULTADOS="
Global Const CTE_ITERACIONES = "ITERACIONES="
Global Const CTE_FRASE_A_BUSCAR = "FRASE A BUSCAR="
Global Const CTE_CRITERIO_DE_PARADA = "CRITERIO DE PARADA="



'Acciones con ficheros
Global Const CTE_FIC_ABRIR = 1
Global Const CTE_FIC_GUARDAR = 2
Global Const CTE_FIC_GUARDARCOMO = 3

'Modos de Apertura
Global Const CTE_ABRIR_BORRAR = 0
Global Const CTE_ABRIR_ANEXAR = 1
Global Const CTE_ABRIR_NOBORRAR = 2


'Numeros para lectura y escritura de ficheros planos
Global Const CTE_FIC_01_INICIO = 1 'inicio.txt de configuracion global
Global Const CTE_FIC_02_AUT = 2 'definicion de un ejemplo automatico
Global Const CTE_FIC_03_MAP = 3 'mapa
Global Const CTE_FIC_04_PRI = 4 'definicion de jugadores al prisionero
Global Const CTE_FIC_05_3R = 5  '3 en raya
Global Const CTE_FIC_06_AZAR = 6 'fichero de azar
Global Const CTE_FIC_07_DICC = 7 'fichero de diccionario
Global Const CTE_FIC_08_BK_IN = 8 'backup, fichero de entrada con la lista de carpetas
Global Const CTE_FIC_09_BK_BAT = 9 'backup, fichero de salida bat generado automaticamente
Global Const CTE_FIC_10_TECLAS_CFG = 10 'configuración de teclas
Global Const CTE_FIC_11_TECLAS_SAL = 11 'Salida de teclas
Global Const CTE_FIC_12_DICC_ENT = 12 'entrada de utilidad de creacion de diccionarios
Global Const CTE_FIC_13_DICC_SAL = 13 'Salida de utilidad de creacion de diccionarios
Global Const CTE_FIC_14_DES = 14 'Fic. desencriptado
Global Const CTE_FIC_15_ENC = 15 'Fic. encriptado
Global Const CTE_FIC_16_ARB = 16 'arbol

Global Const CTE_FIC_20_GLOLOG = 20
Global Const CTE_FIC_21_GLOTXT = 21
Global Const CTE_FIC_22_GLOXLS = 22

Global Const CTE_FIC_23R_1EJGRA = 23
Global Const CTE_FIC_24_1EJTXT = 24
Global Const CTE_FIC_25_1EJXLS = 25

Global Const CTE_FIC_23W_1EJGRA = 33

Global Const CTE_FIC_99_EXISTE = 99 'cualquiera, es para comprobar si ya existe al grabar
Global Const CTE_FIC_100_ULTIMO = 100 'numero maximo de fichero indivudual
                                      'hasta el 255 incluido son no accesibles por otras aplic y el resto si
Global Const CTE_FIC_511_ULTI_LISTA = 511 'numero maximo de fichero total



'Operaciones que se pueden hacer con el formulario de fichero frm_z0_fic
Global Const CTE_SELECCIONAR_FICHERO_OBLIGATIORIO_OP_FICH = 1
Global Const CTE_SELECCIONAR_FICHERO_o_CARPETA_OP_FICH = 2
Global Const CTE_SELECCIONAR_CARPETA_OP_FICH = 3
Global Const CTE_SELECCIONAR_LISTA_FICHEROS_OP_FICH = 4


'0.- Vida Artificial
Global Const CTE_LONG_MAX_APELLIDOS = 300

'Sexos
Global Const CTE_HEMBRA = 1
Global Const CTE_MACHO = 2


'1.- Hormigas y Plantas

'Zoom
Global Const CTE_ZOOM_DETALLE = 0
Global Const CTE_ZOOM_PANORAMICA = 1
Global Const CTE_ZOOM_PIXELS = 2
Global Const CTE_ZOOM_3D = 3
Global Const CTE_ZOOM_SUPER3D = 4
Global Const CTE_tZOOM_DETALLE = "DETALLE"
Global Const CTE_tZOOM_PANORAMICA = "PANORAMICA"
Global Const CTE_tZOOM_PIXELS = "PIXELS"
Global Const CTE_tZOOM_3D = "3D"


Global Const CTE_8_DIR = 8
Global Const CTE_4_DIR = 4

'Funciones de movimiento para relleno de obstaculos con 4 direcciones
Global Const CTE_DIRECC_NINGUNA = 0
Global Const CTE_NORTE = 1
Global Const CTE_ESTE = 2
Global Const CTE_SUR = 3
Global Const CTE_OESTE = 4

'Funciones de movimiento de agentes
'Recorro desde el norte en el sentido de agujas del reloj
Global Const CTE_8_N = 1
Global Const CTE_8_NE = 2
Global Const CTE_8_E = 3
Global Const CTE_8_SE = 4
Global Const CTE_8_S = 5
Global Const CTE_8_SO = 6
Global Const CTE_8_O = 7
Global Const CTE_8_NO = 8

Global Const CTE_DEFRENTE = 1
Global Const CTE_DERECHA = 2
Global Const CTE_ATRAS = 3
Global Const CTE_IZQUIERDA = 4

Global Const CTE_8_DEF = 1
Global Const CTE_8_DEF_DER = 2
Global Const CTE_8_DER = 3
Global Const CTE_8_ATR_DER = 4
Global Const CTE_8_ATR = 5
Global Const CTE_8_ATR_IZQ = 6
Global Const CTE_8_IZQ = 7
Global Const CTE_8_DEF_IZQ = 8


'Tipos de mapas a la hora de hacer referencia a un punto
Global Const CTE_MAPA_ESFERICO = 0
Global Const CTE_MAPA_LIMITADO = 1

'Lo que puede haber en un solo punto de la pantalla
Global Const CTE_MAPA_VACIO = 0
Global Const CTE_MAPA_OBSTACULO = 1
Global Const CTE_MAPA_AGENTE = 2
Global Const CTE_MAPA_PLANTA = 3

'Lo que un agente puede tener rodeandole
Global Const CTE_VEC_NADA = 10
Global Const CTE_VEC_OBSTACULO = 11
Global Const CTE_VEC_PLANTA = 12
Global Const CTE_VEC_AGENTE = 13
Global Const CTE_VEC_AGENTEYPLANTA = 14

Global Const CTE_ACC_NADA = 99
Global Const CTE_ACC_MOVER = 100
Global Const CTE_ACC_REGAR = 101
Global Const CTE_ACC_PELEAR = 102
Global Const CTE_ACC_REPRODUCIRSE = 103
Global Const CTE_ACC_JUGAR = 104


Global Const CTE_MAPA_INI_Z = 0
Global Const CTE_MAPA_INI_Y = 70
Global Const CTE_MAPA_INI_X = 280


Global Const CTE_VA0_INI_Z = 0
Global Const CTE_VA0_INI_Y = 10 '80
Global Const CTE_VA0_INI_X = 10

Global Const CTE_numero_maximo_apellidos = 52

'Tendencias del movimiento
Global Const CTE_RELATIVAS = 1
Global Const CTE_ABSOLUTAS = 2

'2.- Palabras y frases

'3.- Tres en raya
Global Const CTE_GANAR = 1
Global Const CTE_PERDER = -1
Global Const CTE_EMPATAR = 0

Global Const CTE_GANA_EL_PRIMERO = 1
Global Const CTE_GANA_EL_SEGUNDO = 2
Global Const CTE_TABLAS = 0

'Los niveles de prioridad van de 1 a 3 y 3 es la maxima prioridad, las mejores
Global Const CTE_NUMERO_DE_NIVELES_DE_PRIORIDAD = 3
Global Const CTE_MAX_LIN = 1000
Global Const CTE_MAX_CAR_LIN = 500

'4.- Prisionero
Global Const CTE_BEGIN_JUGADOR = "BEGIN JUGADOR"
Global Const CTE_NOMBRE_JUGADOR = "NOMBRE JUGADOR"
Global Const CTE_NUMERO_AGENTES = "NUMERO AGENTES"
Global Const CTE_PARAMETROS_MOVIMIENTO = "PARAMETROS MOVIMIENTO"
Global Const CTE_BEGIN_REGLA = "BEGIN REGLA"
Global Const CTE_PRIORIDAD = "PRIORIDAD"
Global Const CTE_CONDICION = "CONDICION"
Global Const CTE_ACCION = "ACCION"
Global Const CTE_END_REGLA = "END REGLA"
Global Const CTE_END_JUGADOR = "END JUGADOR"

Global Const CTE_C = 0
Global Const CTE_D = 1


'Estado del analisis sintactico
Global Const CTE_INICIO = 0
Global Const CTE_B_JUGADOR_LEIDO = 1
Global Const CTE_N_JUGADOR_LEIDO = 2
Global Const CTE_B_REGLA_LEIDO = 3

'Trigonométricas
Global Const CTE_RAIZDE2 = 1.4142135623731 'SQR(2)
Global Const CTE_RAIZDE3 = 1.73205080756888  'SQR(3)
Global Const CTE_RAIZDE5 = 2.23606797749979 'SQR(5)

Global Const CTE_1ENTRERAIZDE2 = 0.707106781186547 '1/(SQR(2))
Global Const CTE_1ENTRERAIZDE3 = 0.577350269189626 '1/(SQR(3))
Global Const CTE_1ENTRERAIZDE5 = 0.447213595499958  '1/(SQR(5))

Global Const CTE_2ENTRERAIZDE3 = 1.15470053837925 '2/(SQR(3))
Global Const CTE_2ENTRERAIZDE5 = 0.894427190999916 '2/(SQR(5))

Global Const CTE_RAIZDE2ENTRERAIZDE3 = 0.816496580927726 '(SQR(2))/(SQR(3))


Global Const CTE_RAIZDE3porRAIZDE2_entre4 = 0.612372435695795 'SQR(3)*SQR(2)/4






'----------BASES DE DATOS: GAIA---------------
'Numero de campos de descripcion en cada tabla
Global Const CTE_num_campos_des = 1 'numero de campos descripción en las tablas

'Constantes propias de Base de Datos
Global Const CTE_ANADIR = 0
Global Const CTE_MODIFICAR = 1
Global Const CTE_LEER = 1

'Controlador de la base de datos
Global Const CTE_ACCESS = "Access"

'Tipos de Acceso
Global Const CTE_BD_Alta1 = "A1"
Global Const CTE_BD_Baja1 = "B1"
Global Const CTE_BD_Modificacion1 = "M1"
Global Const CTE_BD_Consulta1 = "C1"
Global Const CTE_BD_AltaN = "AN"
Global Const CTE_BD_BajaN = "BN"
Global Const CTE_BD_ModificacionN = "MN"
Global Const CTE_BD_ConsultaN = "CN"

'----------ERRORES---------------
'Constante de deteccion de errores
Global Const CTE_HAY_ERROR = 1
Global Const CTE_NO_HAY_ERROR = 0

'Errores de GS_BD_Error
Global Const CTE_ErrorNinguno = "Ninguno"
Global Const CTE_ErrorRC = "Error en RealizarConsulta"
Global Const CTE_ErrorCNE = "Consulta No Existente"
Global Const CTE_ErrorES = "Error en EjecutarSentencia"
Global Const CTE_ErrorAE = "Alta Existente"
Global Const CTE_ErrorBNE = "Baja No Existente"
Global Const CTE_ErrorSMR = "Selección de más de un Registro"
Global Const CTE_ErrorMNE = "Modificación No Existente"
Global Const CTE_ErrorVB = "Error Detectado por VB"


'**************************************
'***    Tablas relativas a Gaia    ****
'**************************************

'1.- Tabla GLOBAL: no tiene clave porque solo es un registro
'por cada universo y solo hay un universo
'Nombre
Global Const CTE_TABLA_GLOBAL = "GLOBAL"
'Campos
'des
Global Const CTE_GLOBAL_Cod_Uni = "Cod_Uni"
Global Const CTE_GLOBAL_Num_Uni = "Num_Uni"
'Número total de campos
Global Const CTE_N_GLOBAL = 3


'2.- Tabla UNIVERSO
'un registro por cada universo
'Nombre
Global Const CTE_TABLA_UNIVERSO = "UNIVERSO"
'Campos
'des
Global Const CTE_UNIVERSO_Cod_Uni = "Cod_Uni" 'clave
Global Const CTE_UNIVERSO_Cod_Ent = "Cod_Ent"
Global Const CTE_UNIVERSO_Uni_Viv = "Uni_Viv"
Global Const CTE_UNIVERSO_Uni_Pri = "Uni_Pri"
Global Const CTE_UNIVERSO_Num_Ent = "Num_Ent"
'Número total de campos
Global Const CTE_N_UNIVERSO = 6


'3.- Tabla ENTIDAD
'Nombre
Global Const CTE_TABLA_ENTIDAD = "ENTIDAD"
'Campos
'Des
Global Const CTE_ENTIDAD_Cod_Uni = "Cod_Uni"  'Clave
Global Const CTE_ENTIDAD_Cod_Ent = "Cod_Ent"  'Clave
Global Const CTE_ENTIDAD_Ent_Viv = "Ent_Viv"
Global Const CTE_ENTIDAD_Ent_Pri = "Ent_Pri"
Global Const CTE_ENTIDAD_Cod_Obj = "Cod_Obj"
Global Const CTE_ENTIDAD_Cod_Acc = "Cod_Acc"
Global Const CTE_ENTIDAD_Num_Repetida = "Num_Repetida"
'Número total de campos
Global Const CTE_N_ENTIDAD = 8


'4.- Tabla ACCION
'Nombre
Global Const CTE_TABLA_ACCION = "ACCION"
'Campos
'Des
Global Const CTE_ACCION_Cod_Uni = "Cod_Uni"      'Clave
Global Const CTE_ACCION_Cod_Acc = "Cod_Acc"      'Clave
Global Const CTE_ACCION_Cod_Ent = "Cod_Ent"      'Clave
Global Const CTE_ACCION_Num_Repetida = "Num_Repetida"  'Clave
Global Const CTE_ACCION_Num_Orden = "Num_Orden"
Global Const CTE_ACCION_Tip = "Tip"
Global Const CTE_ACCION_Cod_Acc_Padre = "Cod_Acc_Padre"
Global Const CTE_ACCION_Acc_simple = "Acc_Simple"
Global Const CTE_ACCION_Num_Param = "Num_Param"
'Número total de campos
Global Const CTE_N_ACCION = 10


'4.- Tabla PARAMAC
'Nombre
Global Const CTE_TABLA_PARAMAC = "PARAMAC"
'Campos
'Des
Global Const CTE_PARAMAC_Cod_Uni = "Cod_Uni"      'Clave
Global Const CTE_PARAMAC_Cod_Acc = "Cod_Acc"      'Clave
Global Const CTE_PARAMAC_Cod_Ent = "Cod_Ent"      'Clave
Global Const CTE_PARAMAC_Num_Repetida = "Num_Repetida"  'Clave
Global Const CTE_PARAMAC_Cod_Param = "Cod_Param"  'Clave
Global Const CTE_PARAMAC_Cod_Var_Param = "Cod_Var_Param"
'Número total de campos
Global Const CTE_N_PARAMAC = 7



'5.- Tabla REGLA
'Nombre
Global Const CTE_TABLA_REGLA = "REGLA"
'Campos
'Des
Global Const CTE_REGLA_Cod_Uni = "Cod_Uni"  'Clave
Global Const CTE_REGLA_Cod_Ent = "Cod_Ent"  'Clave
Global Const CTE_REGLA_Cod_Reg = "Cod_Reg"  'Clave
Global Const CTE_REGLA_Cod_Contexto = "Cod_Contexto"
Global Const CTE_REGLA_Cod_Acc = "Cod_Acc"
Global Const CTE_REGLA_Cod_Conclusion = "Cod_Conclusion"
'Número total de campos
Global Const CTE_N_REGLA = 7

'6.- Tabla VBLE
'Nombre
Global Const CTE_TABLA_VBLE = "VBLE"
'Campos
'Des
Global Const CTE_VBLE_Cod_Uni = "Cod_Uni"  'Clave
Global Const CTE_VBLE_Cod_Ent = "Cod_Ent"  'Clave
Global Const CTE_VBLE_Cod_Vble = "Cod_Vble"  'Clave
Global Const CTE_VBLE_Valor = "Valor"
Global Const CTE_VBLE_Tipo_Valor = "Tipo_Valor"
Global Const CTE_VBLE_Tipo = "Tipo"
Global Const CTE_VBLE_Conocido = "Conocido"
'Número total de campos
Global Const CTE_N_VBLE = 8

'7.- Tabla CONTEXTO
'Nombre
Global Const CTE_TABLA_CONTEXTO = "CONTEXTO"
'Campos
'Des
Global Const CTE_CONTEXTO_Cod_Uni = "Cod_Uni"  'Clave
Global Const CTE_CONTEXTO_Cod_Ent = "Cod_Ent"  'Clave
Global Const CTE_CONTEXTO_Cod_Contexto = "Cod_Contexto" 'Clave
Global Const CTE_CONTEXTO_Cod_Vble = "Cod_Vble"  'Clave
Global Const CTE_CONTEXTO_Valor = "Valor"
'Número total de campos
Global Const CTE_N_CONTEXTO = 6


'8.- Tabla CONCLUSION
'Nombre
Global Const CTE_TABLA_CONCLUSION = "CONCLUSION"
'Campos
'Des
Global Const CTE_CONCLUSION_Cod_Uni = "Cod_Uni"  'Clave
Global Const CTE_CONCLUSION_Cod_Ent = "Cod_Ent"  'Clave
Global Const CTE_CONCLUSION_Cod_Conclusion = "Cod_Conclusion" 'Clave
Global Const CTE_CONCLUSION_Cod_Vble = "Cod_Vble"  'Clave
Global Const CTE_CONCLUSION_Valor = "Valor"
'Número total de campos
Global Const CTE_N_CONCLUSION = 6


'5.- Tabla MEMORIA
'Nombre
Global Const CTE_TABLA_MEMORIA = "MEMORIA"
'Campos
Global Const CTE_MEMORIA_Cod_Ent = "Cod_Ent"  'Clave
Global Const CTE_MEMORIA_Instante = "Instante"  'Clave
Global Const CTE_MEMORIA_Cod_LI = "Cod_LI"
Global Const CTE_MEMORIA_Cod_Acc = "Cod_Acc"
Global Const CTE_MEMORIA_Cod_LD = "Cod_LD"
'Número total de campos
Global Const CTE_N_MEMORIA = 6



