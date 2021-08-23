VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInformacion 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informacion del Sistema"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12375
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmInformacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   12375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerCarga 
      Interval        =   100
      Left            =   5580
      Top             =   3960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5865
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   10345
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Sistema"
      TabPicture(0)   =   "frmInformacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvwEmpresas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Oracle"
      TabPicture(1)   =   "frmInformacion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Archivos del Sistema"
      TabPicture(2)   =   "frmInformacion.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Sucursales"
      TabPicture(3)   =   "frmInformacion.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Configuración Regional"
      TabPicture(4)   =   "frmInformacion.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView4"
      Tab(4).ControlCount=   1
      Begin MSComctlLib.ListView lvwEmpresas 
         Height          =   5460
         Left            =   30
         TabIndex        =   8
         Top             =   330
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   9631
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Empresa"
            Object.Width           =   19403
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CodEmpresa"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5460
         Left            =   -74970
         TabIndex        =   9
         Top             =   330
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   9631
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Archivo"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ubicación"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Archivo"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ubicación"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5460
         Left            =   -74970
         TabIndex        =   10
         Top             =   330
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   9631
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod. Sucursal"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Activa"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ultimo Lote"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ultima Transacción"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Turno Surtidor"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   5460
         Left            =   -74970
         TabIndex        =   11
         Top             =   330
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   9631
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Empresa"
            Object.Width           =   20285
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CodEmpresa"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   5460
         Left            =   -74970
         TabIndex        =   12
         Top             =   330
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   9631
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8490
      Top             =   6810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10110
      Top             =   6900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Utilidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   30
      TabIndex        =   3
      Top             =   7290
      Width           =   12315
      Begin VB.CommandButton Command1 
         Caption         =   "Prueba PROXY"
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
         Left            =   3240
         TabIndex        =   6
         Top             =   300
         Width           =   1530
      End
      Begin VB.CommandButton cmdVerLote 
         Caption         =   "Ver Lote"
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
         Left            =   1650
         TabIndex        =   5
         Top             =   300
         Width           =   1530
      End
      Begin VB.CommandButton cmdStopDllHost 
         Caption         =   "StopDllHost"
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
         Left            =   60
         TabIndex        =   4
         Top             =   300
         Width           =   1530
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9450
      Top             =   6930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   8085
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7408
            MinWidth        =   7408
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Terminar"
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
      Left            =   6060
      TabIndex        =   1
      Top             =   6930
      Width           =   1530
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Generar Archivo"
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
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6930
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   30
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
End
Attribute VB_Name = "frmInformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn  As ADODB.Connection
Private rst  As ADODB.Recordset
Private SQL  As String
Private itmX As ListItem
Private iEmpresas As Integer
Private strEmpresa As String

Private temp As String
Option Explicit

Private Const OS_ERROR = -1
Private Const OS_95 = 1

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
  ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
  (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400

'-----------GetRegistryValue
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Enum REGRootTypesEnum
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL = &H1F0000

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_READ = ((READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

'-----------GetRegistryValue

'-----------serial number
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias _
    "GetVolumeInformationA" (ByVal lpRootPathName As String, _
    ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long)

Const MAX_PATH = 260

Const FILE_CASE_SENSITIVE_SEARCH = &H1
Const FILE_CASE_PRESERVED_NAMES = &H2
Const FILE_UNICODE_ON_DISK = &H4
Const FILE_PERSISTENT_ACLS = &H8
Const FILE_FILE_COMPRESSION = &H10
Const FILE_VOLUME_IS_COMPRESSED = &H8000

Private DiskSerialNumber As Long
Private ActivationKey As String
'-----------serial number

'-----------Configuracion Regional
   'Declaraciones del Api
   Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
   
   Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
        ByVal Locale As Long, _
        ByVal LCType As Long, _
        ByVal lpLCData As String, _
        ByVal cchData As Long) As Long
   
   'Constante para obtener algunos de los simbolos de la configuración regional
   Private Const LOCALE_ICENTURY = &H24 ' especificador de formato de siglo
   Private Const LOCALE_ICOUNTRY = &H5 ' código del país
   Private Const LOCALE_ICURRDIGITS = &H19 ' número de dígitos de la moneda local
   Private Const LOCALE_ICURRENCY = &H1B ' modo de moneda positiva
   Private Const LOCALE_IDATE = &H21 ' orden del formato de fecha corta
   Private Const LOCALE_IDAYLZERO = &H26 ' número de ceros iniciales en el campo día
   Private Const LOCALE_IDEFAULTCODEPAGE = &HB ' página de códigos predeterminada
   Private Const LOCALE_IDEFAULTCOUNTRY = &HA ' código predeterminado del país
   Private Const LOCALE_IDEFAULTLANGUAGE = &H9 ' Id. predeterminado del idioma
   Private Const LOCALE_IDIGITS = &H11 ' número de dígitos fraccionarios
   Private Const LOCALE_IINTLCURRDIGITS = &H1A ' nº de dígitos de la moneda internacional
   Private Const LOCALE_ILANGUAGE = &H1 ' Id. de idioma
   Private Const LOCALE_ILDATE = &H22 ' orden del formato de fecha larga
   Private Const LOCALE_ILZERO = &H12 ' número de ceros iniciales de decimales
   Private Const LOCALE_IMEASURE = &HD ' 0 = métrico, 1 = EE.UU.
   Private Const LOCALE_IMONLZERO = &H27 ' número de ceros iniciales en el campo mes
   Private Const LOCALE_INEGCURR = &H1C ' modo de moneda negativa
   Private Const LOCALE_INEGSEPBYSPACE = &H57 ' símbolo de moneda separado por un espacio de cantidad negativa
   Private Const LOCALE_INEGSIGNPOSN = &H53 ' posición del signo negativo
   Private Const LOCALE_INEGSYMPRECEDES = &H56 ' símbolo de moneda precede a cantidad negativa
   Private Const LOCALE_IPOSSEPBYSPACE = &H55 ' símbolo de moneda separado por un espacio de cantidad positiva
   Private Const LOCALE_IPOSSIGNPOSN = &H52 ' posición del signo positivo
   Private Const LOCALE_IPOSSYMPRECEDES = &H54 ' símbolo de moneda precede a cantidad positiva
   Private Const LOCALE_ITIME = &H23 ' especificador de formato de hora
   Private Const LOCALE_ITLZERO = &H25 ' número de ceros iniciales en el campo hora
   Private Const LOCALE_NOUSEROVERRIDE = &H80000000 ' no usa sustituciones del usuario
   Private Const LOCALE_S1159 = &H28 ' designador de AM
   Private Const LOCALE_S2359 = &H29 ' designador de PM
   Private Const LOCALE_SABBREVCTRYNAME = &H7 ' nombre abreviado del país
   Private Const LOCALE_SABBREVDAYNAME1 = &H31 ' nombre abreviado para Lunes
   Private Const LOCALE_SABBREVDAYNAME2 = &H32 ' nombre abreviado para Martes
   Private Const LOCALE_SABBREVDAYNAME3 = &H33 ' nombre abreviado para Miércoles
   Private Const LOCALE_SABBREVDAYNAME4 = &H34 ' nombre abreviado para Jueves
   Private Const LOCALE_SABBREVDAYNAME5 = &H35 ' nombre abreviado para Viernes
   Private Const LOCALE_SABBREVDAYNAME6 = &H36 ' nombre abreviado para Sábado
   Private Const LOCALE_SABBREVDAYNAME7 = &H37 ' nombre abreviado para Domingo
   Private Const LOCALE_SABBREVLANGNAME = &H3 ' nombre del idioma abreviado
   Private Const LOCALE_SABBREVMONTHNAME1 = &H44 ' nombre abreviado para Enero
   Private Const LOCALE_SABBREVMONTHNAME10 = &H4D ' nombre abreviado para Octubre
   Private Const LOCALE_SABBREVMONTHNAME11 = &H4E ' nombre abreviado para Noviembre
   Private Const LOCALE_SABBREVMONTHNAME12 = &H4F ' nombre abreviado para Diciembre
   Private Const LOCALE_SABBREVMONTHNAME13 = &H100F
   Private Const LOCALE_SABBREVMONTHNAME2 = &H45 ' nombre abreviado para Febrero
   Private Const LOCALE_SABBREVMONTHNAME3 = &H46 ' nombre abreviado para Marzo
   Private Const LOCALE_SABBREVMONTHNAME4 = &H47 ' nombre abreviado para Abril
   Private Const LOCALE_SABBREVMONTHNAME5 = &H48 ' nombre abreviado para Mayo
   Private Const LOCALE_SABBREVMONTHNAME6 = &H49 ' nombre abreviado para Junio
   Private Const LOCALE_SABBREVMONTHNAME7 = &H4A ' nombre abreviado para Julio
   Private Const LOCALE_SABBREVMONTHNAME8 = &H4B ' nombre abreviado para Agosto
   Private Const LOCALE_SABBREVMONTHNAME9 = &H4C ' nombre abreviado para Septiembre
   Private Const LOCALE_SCOUNTRY = &H6 ' nombre traducido del país
   Private Const LOCALE_SCURRENCY = &H14 ' símbolo de moneda local
   Private Const LOCALE_SDATE = &H1D ' separador de fecha
   Private Const LOCALE_SDAYNAME1 = &H2A ' nombre largo para Lunes
   Private Const LOCALE_SDAYNAME2 = &H2B ' nombre largo para Martes
   Private Const LOCALE_SDAYNAME3 = &H2C ' nombre largo para Miércoles
   Private Const LOCALE_SDAYNAME4 = &H2D ' nombre largo para Jueves
   Private Const LOCALE_SDAYNAME5 = &H2E ' nombre largo para Viernes
   Private Const LOCALE_SDAYNAME6 = &H2F ' nombre largo para Sábado
   Private Const LOCALE_SDAYNAME7 = &H30 ' nombre largo para Domingo
   Private Const LOCALE_SDECIMAL = &HE ' separador de decimales
   Private Const LOCALE_SENGCOUNTRY = &H1002 ' nombre del país en inglés
   Private Const LOCALE_SENGLANGUAGE = &H1001 ' nombre del idioma en inglés
   Private Const LOCALE_SGROUPING = &H10 ' agrupación de dígitos
   Private Const LOCALE_SINTLSYMBOL = &H15 ' símbolo de moneda internacional
   Private Const LOCALE_SLANGUAGE = &H2 ' nombre traducido del idioma
   Private Const LOCALE_SLIST = &HC ' separador de elementos de lista
   Private Const LOCALE_SLONGDATE = &H20 ' cadena de formato de fecha larga
   Private Const LOCALE_SMONDECIMALSEP = &H16 ' separador de decimales en moneda
   Private Const LOCALE_SMONGROUPING = &H18 ' agrupación de moneda
   Private Const LOCALE_SMONTHNAME1 = &H38 ' nombre largo para Enero
   Private Const LOCALE_SMONTHNAME10 = &H41 ' nombre largo para Octubre
   Private Const LOCALE_SMONTHNAME11 = &H42 ' nombre largo para Noviembre
   Private Const LOCALE_SMONTHNAME12 = &H43 ' nombre largo para Diciembre
   Private Const LOCALE_SMONTHNAME2 = &H39 ' nombre largo para Febrero
   Private Const LOCALE_SMONTHNAME3 = &H3A ' nombre largo para Marzo
   Private Const LOCALE_SMONTHNAME4 = &H3B ' nombre largo para Abril
   Private Const LOCALE_SMONTHNAME5 = &H3C ' nombre largo para Mayo
   Private Const LOCALE_SMONTHNAME6 = &H3D ' nombre largo para Junio
   Private Const LOCALE_SMONTHNAME7 = &H3E ' nombre largo para Julio
   Private Const LOCALE_SMONTHNAME8 = &H3F ' nombre largo para Agosto
   Private Const LOCALE_SMONTHNAME9 = &H40 ' nombre largo para Septiembre
   Private Const LOCALE_SMONTHOUSANDSEP = &H17 ' separador de miles en moneda
   Private Const LOCALE_SNATIVECTRYNAME = &H8 ' nombre nativo del país
   Private Const LOCALE_SNATIVEDIGITS = &H13 ' ASCII 0-9 nativo
   Private Const LOCALE_SNATIVELANGNAME = &H4 ' nombre nativo del idioma
   Private Const LOCALE_SNEGATIVESIGN = &H51 ' signo negativo
   Private Const LOCALE_SPOSITIVESIGN = &H50 ' signo positivo
   Private Const LOCALE_SSHORTDATE = &H1F ' cadena de formato de fecha corta
   Private Const LOCALE_STHOUSAND = &HF ' separador de miles
   Private Const LOCALE_STIME = &H1E ' separador de hora
   Private Const LOCALE_STIMEFORMAT = &H1003 ' cadena de formato de hora
'-----------Configuracion Regional
Private strGuiones As String * 128

Private Sub Form_Initialize()

   On Error GoTo GestErr
   
   Me.MousePointer = vbHourglass
   Pictures
   TimerCarga.Enabled = True
   
   Exit Sub
   
GestErr:
  Me.MousePointer = vbNormal
  MsgBox "[Form_Initialize]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub CargarTodo()
Dim sh As Object
Dim parametros    As String
Dim aParametros() As String
Dim ix            As Integer

10       On Error GoTo GestErr
   
20       Set cnn = New ADODB.Connection 'HKEY_LOCAL_MACHINE\SOFTWARE\Algoritmo\DatabaseSettings\ConnectionsStrings
30       cnn.ConnectionString = GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Algoritmo\DatabaseSettings\ConnectionsStrings", "ALG", REG_SZ, "", False)
         If Len(cnn.ConnectionString) = 0 Then cnn.ConnectionString = "Provider=MSDataShape;Data Provider=MSDAORA;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
40       cnn.Open
50       Set rst = New ADODB.Recordset
60       rst.CursorLocation = adUseClient
70       rst.LockType = adLockReadOnly
80       rst.CursorType = adOpenStatic

81       Set sh = CreateObject("Wscript.Shell")
82       sh.Run ("net use \\juanjo /user:algoritmo\compiler apfrms2001"), 0, False
83       sh.Run ("reg export HKEY_LOCAL_MACHINE\SOFTWARE\Algoritmo c:\windows\temp\registroX32.reg"), 0, False
84       sh.Run ("reg export HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Algoritmo c:\windows\temp\registroX64.reg"), 0, False
85       Set sh = Nothing
         
         strGuiones = String(125, "-")
         
100      VersionOracle
   
110      Licencias
   
120      VersionSistema
   
130      ConfiguracionIP
   
140      Sucursales

150      Archivos

151      ConfiguracionRegional

         'Parametros
         parametros = Command$
      
         aParametros = Split(parametros, " ")
         For ix = LBound(aParametros) To UBound(aParametros)
            If ix = 0 Then
               Select Case aParametros(ix)
                  Case "I"
                     GenerarTXT True
                     Copiar_Informacion
                     End
                  Case Else
                     MsgBox "<Informacion> Recopila informacion de SoftCereal y Oracle " & vbCrLf & _
                     "Uso: Informacion.exe  [I] (Usado en Proceso de Instalacion)"
                     End
               End Select
            End If
         Next ix
         
         TimerCarga.Enabled = False
         
160      Set rst.ActiveConnection = Nothing
170      Set rst = Nothing
180      cnn.Close
190      Set cnn = Nothing
   
200      Exit Sub
   
GestErr:
210     Me.MousePointer = vbNormal
220     MsgBox "[CargarTodo]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub GenerarTXT(EnSilencio As Boolean)

      Dim fNumber1 As Long

10       On Error GoTo GestErr
   
20       cmdAceptar.Enabled = False
30       cmdCancelar.Enabled = False

40       If Not EnSilencio Then Screen.MousePointer = vbHourglass
   
50       fNumber1 = FreeFile
60       Open strEmpresa & ".txt" For Output As fNumber1
   
         'Sistema
70       For Each itmX In lvwEmpresas.ListItems
80          Print #fNumber1, itmX.Text
90       Next itmX
'100      Print #fNumber1, ""
         
         Print #fNumber1, strGuiones
         
         'Oracle
110      For Each itmX In ListView3.ListItems
120         Print #fNumber1, itmX.Text
130      Next itmX
         Print #fNumber1, strGuiones
150      Print #fNumber1, "Archivos del Sistema"
         Print #fNumber1, strGuiones
170      For Each itmX In ListView1.ListItems
180         Print #fNumber1, itmX.Text & Space$(30 - Len(itmX.Text)) _
                           & itmX.SubItems(1) & Space$(6 - Len(itmX.SubItems(1))) _
                           & itmX.SubItems(2) & Space$(18 - Len(itmX.SubItems(2))) _
                           & itmX.SubItems(3) & Space$(30 - Len(itmX.SubItems(3))) _
                           & itmX.SubItems(4) & Space$(6 - Len(itmX.SubItems(4))) _
                           & itmX.SubItems(5)
190      Next itmX
         Print #fNumber1, strGuiones
210      Print #fNumber1, "Sucursales"
         Print #fNumber1, strGuiones
230      Print #fNumber1, "Codigo     Descripcion                                      Activa  Ult.Lote Ult.Transaccion Ult.Auditoria Ult.Surt."
240      For Each itmX In ListView2.ListItems
250         Print #fNumber1, itmX.Text & Space$(11 - Len(itmX.Text)) _
                           & itmX.SubItems(1) & Space$(51 - Len(itmX.SubItems(1))) _
                           & itmX.SubItems(2) & Space$(10 - Len(itmX.SubItems(2))) _
                           & itmX.SubItems(3) & Space$(10 - Len(itmX.SubItems(3))) _
                           & itmX.SubItems(4) & Space$(10 - Len(itmX.SubItems(4))) _
                           & itmX.SubItems(5)
260      Next itmX
   
270      Close #fNumber1
   
280      If Not EnSilencio Then MsgBox "Proceso Terminado Satisfactoriamente"
290      If Not EnSilencio Then Screen.MousePointer = vbNormal
   
300      cmdAceptar.Enabled = True
310      cmdCancelar.Enabled = True
   
320      Exit Sub

GestErr:
330     Me.MousePointer = vbNormal
340     MsgBox "[GenerarTXT]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub ShutDownServerByName(strComputer As String, strPackage$)
Dim cat  As MTSAdmin.Catalog
Dim pkgs  As MTSAdmin.CatalogCollection
Dim pkg  As MTSAdmin.CatalogObject
Dim pkgutil  As MTSAdmin.PackageUtil
  
   Set cat = GetCatalog(strComputer)
   Set pkgs = cat.GetCollection("Packages")
   Set pkg = GetObjectFromCollection(pkgs, strPackage)
   If pkg Is Nothing Then
      Err.Raise Err.Number, Err.Source, "El paquete " & strPackage & " no existe"
      Exit Sub
   End If
   
   Set pkgutil = pkgs.GetUtilInterface
   Call pkgutil.ShutdownPackage(pkg.Value("ID"))
End Sub ' ShutDownServerByName

Private Function GetCatalog(strComputer As String) As MTSAdmin.Catalog
  Dim cat As New MTSAdmin.Catalog
  If strComputer <> "" Then
    cat.Connect strComputer
  End If
  Set GetCatalog = cat
End Function ' GetCatalog

Private Function GetObjectFromCollection(coll As MTSAdmin.CatalogCollection, strObjName As String) As MTSAdmin.CatalogObject
  Dim obj  As MTSAdmin.CatalogObject
  
  coll.Populate
  For Each obj In coll
    If obj.Name = strObjName Then
      Set GetObjectFromCollection = obj
      Exit Function
    End If
  Next
  Set GetObjectFromCollection = Nothing
  
End Function

Private Sub cmdStopDllHost_Click()
   ShutDownServerByName "", "Algoritmo"
End Sub

Private Sub cmdVerLote_Click()
      Dim rst     As New ADODB.Recordset
      Dim pb      As New PropertyBag
      Dim abyte() As Byte


10       On Error GoTo GestErr
   
20       CommonDialog1.DialogTitle = "Elegir Lote"
30       CommonDialog1.Filter = "Lotes (*.LOT)|*.LOT"
40       CommonDialog1.FilterIndex = 1
50       CommonDialog1.Flags = cdlOFNFileMustExist
60       CommonDialog1.ShowOpen

70       If Len(CommonDialog1.FileName) = 0 Then Exit Sub
   
80       rst.Open CommonDialog1.FileName
   
      '   rst.Filter = "trs_COMENTARIO = 'EGRESO_CAMIONES Emitir'"
      '   rst.Filter = "TRS_NUMERO_TRANSACCION = 4719"
   
90       Do While Not rst.EOF
   
100         abyte = rst("TRS_DATOS")
110         pb.Contents = abyte
   
120         lvwEmpresas.ListItems.Clear
130         Titulo pb.Contents
      
140         rst.MoveNext
150      Loop
   
160      rst.Filter = adFilterNone
   
170      Exit Sub
   
GestErr:
180      Me.MousePointer = vbNormal
190      MsgBox "[cmdVerLote_Click]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Command1_Click()
   PruebaProxy
End Sub

Private Sub Titulo(Nombre As String)
   
   Set itmX = lvwEmpresas.ListItems.Add
   itmX.Text = strGuiones
   
   Set itmX = lvwEmpresas.ListItems.Add
   itmX.Text = Nombre
      
   Set itmX = lvwEmpresas.ListItems.Add
   itmX.Text = strGuiones

End Sub

Private Sub ConfiguracionIP()
        Dim nOS As Long
        Dim temp As String, bDNS As Boolean
        Dim hProcess As Long
        Dim ProcessId As Long
        Dim exitCode As Long
        Dim linea As Integer
  
10      On Error GoTo GestErr
  
        Set itmX = lvwEmpresas.ListItems.Add
        itmX.Text = strGuiones
         
20      Set itmX = lvwEmpresas.ListItems.Add
  
40      Open "ip.bat" For Output As 1
50      If nOS = OS_95 Then
60        Print #1, "Winipcfg.exe /all /batch 1.txt"
70        Print #1, "ren 1.txt ip.txt"
80      Else
90        Print #1, "ipconfig /all > ip.txt"
100       Print #1, "echo . >> ip.txt"
          Print #1, "echo " & strGuiones & " >> ip.txt"
110       Print #1, "echo Sistema Operativo >> ip.txt"
          Print #1, "echo " & strGuiones & " >> ip.txt"
'120       Print #1, "wmic os get Caption,CSDVersion /value >> ip.txt"
120       Print #1, "systeminfo.exe >> ip.txt"
130     End If
140     Close #1
150     While Dir("ip.bat") = ""
160       DoEvents
170     Wend
180     ProcessId = Shell("ip.bat", vbHide)
190     If nOS = OS_95 Then
200       While Dir("ip.txt") = ""
210         DoEvents
220       Wend
230     Else
240       hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId)
250       Do
260           Call GetExitCodeProcess(hProcess, exitCode)
270           DoEvents
280       Loop While exitCode > 0
290     End If
300     CloseHandle hProcess
310     Open "ip.txt" For Input As 1
320     linea = 0
330     While Not EOF(1)
340       Line Input #1, temp
350       If Trim(temp) <> "" And InStr(temp, "File 1") = 0 Then
360         itmX.Text = Trim(temp)
370         If linea = 0 Then
380            itmX.Bold = True
390         End If
400         Set itmX = lvwEmpresas.ListItems.Add
410         linea = linea + 1
420      End If
430     Wend
440     Close #1
450     Kill "ip.bat"
460     Kill "ip.txt"
470     Me.MousePointer = vbNormal
  
        Set itmX = lvwEmpresas.ListItems.Add
        itmX.Text = strGuiones
        Set itmX = lvwEmpresas.ListItems.Add
        itmX.Text = "Informacion de Discos"
        Set itmX = lvwEmpresas.ListItems.Add
        itmX.Text = strGuiones
        
        On Error Resume Next
        Dim strComputer As String
        Dim objWMIService As Object
        Dim colItems As Object
        Dim objItem As Object
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive", , 48)
        For Each objItem In colItems
'            Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "Availability: " & objItem.Availability
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "BytesPerSector: " & objItem.BytesPerSector
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Capabilities: " & objItem.Capabilities
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "CapabilityDescriptions: " & objItem.CapabilityDescriptions
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Caption: " & objItem.Caption
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "CompressionMethod: " & objItem.CompressionMethod
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "CreationClassName: " & objItem.CreationClassName
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "DefaultBlockSize: " & objItem.DefaultBlockSize
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Description: " & objItem.Description
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "DeviceID: " & objItem.DeviceID
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "ErrorCleared: " & objItem.ErrorCleared
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "ErrorDescription: " & objItem.ErrorDescription
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "ErrorMethodology: " & objItem.ErrorMethodology
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Index: " & objItem.Index
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "InstallDate: " & objItem.InstallDate
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "InterfaceType: " & objItem.InterfaceType
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "LastErrorCode: " & objItem.LastErrorCode
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Manufacturer: " & objItem.Manufacturer
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "MaxBlockSize: " & objItem.MaxBlockSize
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "MaxMediaSize: " & objItem.MaxMediaSize
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "MediaLoaded: " & objItem.MediaLoaded
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "MediaType: " & objItem.MediaType
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "MinBlockSize: " & objItem.MinBlockSize
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Model: " & objItem.Model
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Name: " & objItem.Name
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "NeedsCleaning: " & objItem.NeedsCleaning
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "NumberOfMediaSupported: " & objItem.NumberOfMediaSupported
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Partitions: " & objItem.Partitions
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "PNPDeviceID: " & objItem.PNPDeviceID
'             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
             Set itmX = lvwEmpresas.ListItems.Add
'             itmX.Text = "PowerManagementSupported: " & objItem.PowerManagementSupported
'             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "SCSIBus: " & objItem.SCSIBus
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "SCSILogicalUnit: " & objItem.SCSILogicalUnit
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "SCSIPort: " & objItem.SCSIPort
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "SCSITargetId: " & objItem.SCSITargetId
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "SectorsPerTrack: " & objItem.SectorsPerTrack
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Size: " & objItem.Size
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "Status: " & objItem.Status
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "StatusInfo: " & objItem.StatusInfo
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "SystemCreationClassName: " & objItem.SystemCreationClassName
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "SystemName: " & objItem.SystemName
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "TotalCylinders: " & objItem.TotalCylinders
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "TotalHeads: " & objItem.TotalHeads
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "TotalSectors: " & objItem.TotalSectors
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "TotalTracks: " & objItem.TotalTracks
             Set itmX = lvwEmpresas.ListItems.Add
             itmX.Text = "TracksPerCylinder: " & objItem.TracksPerCylinder
        Next
         
480     Exit Sub

GestErr:
490     Me.MousePointer = vbNormal
500     MsgBox "[ConfiguracionIP]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub Copiar_Informacion()
              Dim nOS As Long
              Dim temp As String, bDNS As Boolean
              Dim hProcess As Long
              Dim ProcessId As Long
              Dim exitCode As Long

        'Me.MousePointer = vbHourglass

10      Open "ip.bat" For Output As 1
'20      Print #1, "net use \\juanjo /user:algoritmo\juanjo.sortino kjs992sc"
30      Print #1, "xcopy " & strEmpresa & ".txt" & " \\juanjo\temp /y"
40      Close #1

50      While Dir("ip.bat") = ""
60        DoEvents
70      Wend
80      ProcessId = Shell("ip.bat", vbHide)
90      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId)
100     Do
110       Call GetExitCodeProcess(hProcess, exitCode)
120       DoEvents
130     Loop While exitCode > 0
140     CloseHandle hProcess

150     Kill "ip.bat"
        'Me.MousePointer = vbNormal
  
160      Exit Sub
GestErr:
170     Me.MousePointer = vbNormal
180     MsgBox "[Copiar_Informacion]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Image1_DblClick()
   GenerarTXT True
   Copiar_Informacion
End Sub
Private Sub cmdCancelar_Click()
   End
End Sub
Private Sub cmdAceptar_Click()
   GenerarTXT False
End Sub

Private Function GetRegistryValue(ByVal hKey As REGRootTypesEnum, ByVal KeyName As String, ByVal ValueName As String, Optional ByVal KeyType As REGKeyTypesEnum, Optional DefaultValue As Variant, Optional ByVal Create As Boolean) As Variant
      Dim handle As Long, resLong As Long
      Dim resString As String, length As Long
      Dim resBinary() As Byte

10       On Error GoTo GestErr

20       If KeyType = 0 Then
30          KeyType = REG_SZ
40       End If

50       If IsMissing(DefaultValue) Then
60          Select Case KeyType
               Case REG_SZ
70                DefaultValue = ""
80             Case REG_DWORD
90                DefaultValue = 0
100            Case REG_BINARY
110               DefaultValue = 0
120         End Select
130      End If

         ' Prepare the default result.
140      GetRegistryValue = DefaultValue
         ' Open the key, exit if not found.
150      If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
160         If Create Then
               'si no exite la creo
170            If CreateRegistryKey(hKey, KeyName) Then Exit Function
180            If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
190         Else
200            Exit Function
210         End If
220      End If

230      Select Case KeyType
             Case REG_DWORD
                 ' Read the value, use the default if not found.
240              If RegQueryValueEx(handle, ValueName, 0, REG_DWORD, resLong, 4) = 0 Then
250                  GetRegistryValue = resLong
260               Else
270                  If Create Then
280                     SetRegistryValue hKey, KeyName, ValueName, REG_DWORD, DefaultValue
290                  End If
300              End If
310          Case REG_SZ
320              length = 1024: resString = Space$(length)
330              If RegQueryValueEx(handle, ValueName, 0, REG_SZ, ByVal resString, length) = 0 Then
                     ' If value is found, trim characters in excess.
340                  GetRegistryValue = Left$(resString, length - 1)
350               Else
360                  If Create Then
370                     SetRegistryValue hKey, KeyName, ValueName, REG_SZ, DefaultValue
380                  End If
390              End If
400          Case REG_BINARY
410              length = 4096
420              ReDim resBinary(length - 1) As Byte
430              If RegQueryValueEx(handle, ValueName, 0, REG_BINARY, resBinary(0), length) = 0 Then
440                  ReDim Preserve resBinary(length - 1) As Byte
450                  GetRegistryValue = resBinary()
460               Else
470                  If Create Then
480                     SetRegistryValue hKey, KeyName, ValueName, REG_BINARY, DefaultValue
490                  End If
500              End If
510          Case Else
520              Err.Raise 1001, , "Tipo de valor no soportado"
530      End Select

540      RegCloseKey handle

550      Exit Function

GestErr:
560      Me.MousePointer = vbNormal
570      MsgBox "[GetRegistryValue]" & vbCrLf & Err.Description & Erl
End Function
Function GetDriveInfo(ByVal DriveName As String, Optional VolumeName As String, _
    Optional SerialNumber As Long, Optional FileSystem As String, _
    Optional FileSystemFlags As Long) As Boolean
    
    Dim ignore As Long
    
    ' if it isn't a UNC path, enforce the correct format
    If InStr(DriveName, "\\") = 0 Then
        DriveName = Left$(DriveName, 1) & ":\"
    End If
    
    ' prepare receiving buffers
    SerialNumber = 0
    FileSystemFlags = 0
    VolumeName = String$(MAX_PATH, 0)
    FileSystem = String$(MAX_PATH, 0)
    
    ' The API function return a non-zero value if successful
    GetDriveInfo = GetVolumeInformation(DriveName, VolumeName, Len(VolumeName), _
        SerialNumber, ignore, FileSystemFlags, FileSystem, Len(FileSystem))
    ' drop characters in excess
    VolumeName = Left$(VolumeName, InStr(VolumeName, vbNullChar) - 1)
    FileSystem = Left$(FileSystem, InStr(FileSystem, vbNullChar) - 1)
    
End Function

Private Function GenerateActivationKey(lngKey As Long) As String
Dim strKey As String
Dim s1 As String
Dim ix As Integer
Dim a1 As Variant
Dim a2 As Variant
Dim a3 As Variant
Dim a4 As Variant

   strKey = CStr(lngKey)
   
   s1 = ""
   For ix = 1 To Len(strKey)
      s1 = s1 & CInt(Val(Mid(strKey, ix, 1) * 3))
   Next ix
   
   
   strKey = Left(s1 & "1234567890123456", 16)
   s1 = ""
   
   For ix = Len(strKey) To 1 Step -1
      s1 = s1 & (Val(Mid(strKey, ix, 1) * 4) Mod 9)
   Next ix


   a1 = CLng(Left(s1, 4))
   a2 = CLng(Mid(s1, 5, 4))
   a3 = CLng(Mid(s1, 9, 4))
   a4 = CLng(Mid(s1, 13, 4))
   
   a1 = Right("0000" & Abs(a1 - a3) Mod 9999, 4)
   a2 = Right("0000" & Abs(a2 - a1) Mod 9999, 4)
   a3 = Right("0000" & Abs(a3 - a1) Mod 9999, 4)
   a4 = Right("0000" & Abs(a4 - a2) Mod 9999, 4)

   GenerateActivationKey = a1 & "-" & a2 & "-" & a3 & "-" & a4
   
End Function

Private Function GenerateAK() As Long
Dim Key As Long

   GetDriveInfo "C:", , DiskSerialNumber
   
   GenerateAK = DiskSerialNumber

End Function

Private Sub PruebaProxy()
         Dim objSPM As Object

10       On Error GoTo GestErr

20       If objSPM Is Nothing Then
30          MsgBox "Set objSPM = CreateObject(DataShare.SPM)"
40          Set objSPM = CreateObject("DataShare.SPM")
50       End If
60       If Not (objSPM Is Nothing) Then MsgBox "El Objeto se creo correctamente"

70       Set objSPM = Nothing
80       If objSPM Is Nothing Then
90          MsgBox "Set objSPM = CreateObject(DataShare.SPM)," & GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Algoritmo\Environment\Application Server", "Remote Host", REG_SZ, "", True)
100         Set objSPM = CreateObject("DataShare.SPM", GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Algoritmo\Environment\Application Server", "Remote Host", REG_SZ, "", True))
110      End If
120      If Not (objSPM Is Nothing) Then MsgBox "El Objeto se creo correctamente"
    
130      Exit Sub

GestErr:
140      Me.MousePointer = vbNormal
150      MsgBox "[PruebaProxy]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Archivos()

         Dim itmX As ListItem
         Dim fso As Scripting.FileSystemObject
         Dim file1 As Scripting.File
         Dim strServerRoot As String
         Dim strServerDir As String
         Dim strClientDir As String
         Dim ix As Integer
         Dim iy As Integer
         Dim strFile As String
   
10       On Error Resume Next
      
20       Set fso = CreateObject("Scripting.FileSystemObject")
30       Set fso = CreateObject("Scripting.FileSystemObject")
   
40       strClientDir = "C:\Archivos de programa\Algoritmo\"
50       strFile = Dir(AddBackslash(strClientDir) & "*.*", vbNormal)
   
60       Do While strFile <> ""
      
70          Set file1 = fso.GetFile(AddBackslash(strClientDir) & strFile)     ' Obtiene un objeto File para consultar.
      
80          ix = ix + 1
      
90          Set itmX = Me.ListView1.ListItems.Add(, , file1.Name)
100         itmX.SubItems(1) = "Local"
110         itmX.SubItems(2) = Format(file1.DateLastModified, "dd-mm-yyyy hh:mm")
      
120         strFile = Dir
130      Loop
   
         'obtengo la ubicación en el server de la version del producto
140      strServerRoot = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo\Environment\Application Server", "Remote Host", REG_SZ, "C:", True)

150      strServerDir = "\\" & AddBackslash(strServerRoot) & "\d\Algoritmo\Componentes Server\"
         'strServerDir = AddBackslash(strServerRoot) & "Algoritmo\Componentes Server\"
   
160      strFile = Dir(AddBackslash(strServerDir) & "*.*", vbNormal)
      
170      iy = 1
180      Do While strFile <> ""
      
190         Set file1 = fso.GetFile(AddBackslash(strServerDir) & strFile)     ' Obtiene un objeto File para consultar.
200         If iy <= ix Then
210            Set itmX = Me.ListView1.ListItems(iy)
220            iy = iy + 1
230         Else
240            Set itmX = Me.ListView1.ListItems.Add()
250         End If
260         itmX.SubItems(3) = file1.Name
270         itmX.SubItems(4) = "Server"
280         itmX.SubItems(5) = Format(file1.DateLastModified, "dd-mm-yyyy hh:mm")
      
290         strFile = Dir
300      Loop
   
310      Exit Sub
   
GestErr:
320      MsgBox Err.Description
330      MsgBox "[Archivos]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Pictures()

   On Error Resume Next
   
   Image1.Picture = LoadPicture("C:\Archivos de programa\Algoritmo\Recursos\BitMaps\Algoritmo.JPG")
   'lvwEmpresas.Picture = LoadPicture("C:\windows\Santa Fe.bmp")
   'ListView1.Picture = LoadPicture("C:\windows\Santa Fe.bmp")
   'ListView2.Picture = LoadPicture("C:\windows\Santa Fe.bmp")
   
End Sub

Private Sub VersionOracle()

10       On Error GoTo GestErr
   
20       SQL = " SELECT BANNER "
30       SQL = SQL & "  FROM SYS.V_$VERSION "
'40       SQL = SQL & "UNION ALL "
'50       SQL = SQL & "SELECT COMP_NAME || ' ' || VERSION || ' ' || STATUS || ' ' || MODIFIED "
'60       SQL = SQL & "  FROM DBA_REGISTRY "
70       SQL = SQL & "UNION ALL "
80       SQL = SQL & "SELECT '.' "
90       SQL = SQL & "  FROM DUAL "
100      SQL = SQL & "UNION ALL "
110      SQL = SQL & "SELECT 'Empresas:' "
120      SQL = SQL & "  FROM DUAL "
130      SQL = SQL & "UNION ALL "
140      SQL = SQL & "SELECT '.' "
150      SQL = SQL & "  FROM DUAL "
160      SQL = SQL & "UNION ALL "
170      SQL = SQL & "SELECT EMP_CODIGO_EMPRESA || ' --> ' || EMP_DESCRIPCION "
180      SQL = SQL & "  FROM EMPRESAS "
190      SQL = SQL & "UNION ALL "
200      SQL = SQL & "SELECT '.' "
210      SQL = SQL & "  FROM DUAL "
220      SQL = SQL & "UNION ALL "
230      SQL = SQL & "SELECT 'Oracle Instance Cache Hit Ratio: ' "
240      SQL = SQL & "       || ROUND ((1 - (PHY.VALUE / (CUR.VALUE + CON.VALUE))) * 100, 2) ""Cache Hit Ratio"" "
250      SQL = SQL & "  FROM V$SYSSTAT CUR, V$SYSSTAT CON, V$SYSSTAT PHY "
260      SQL = SQL & " WHERE CUR.NAME = 'db block gets' "
270      SQL = SQL & "   AND CON.NAME = 'consistent gets' "
280      SQL = SQL & "   AND PHY.NAME = 'physical reads' "
290      SQL = SQL & "UNION ALL "
300      SQL = SQL & "SELECT '.' "
310      SQL = SQL & "  FROM DUAL "
320      SQL = SQL & "UNION ALL "
330      SQL = SQL & "SELECT    'Instance_Name: ' "
340      SQL = SQL & "       || INSTANCE_NAME "
350      SQL = SQL & "       || '  Host_Name: ' "
360      SQL = SQL & "       || HOST_NAME "
370      SQL = SQL & "       || '  Version: ' "
380      SQL = SQL & "       || VERSION "
390      SQL = SQL & "       || '  Startup_Time: ' "
400      SQL = SQL & "       || STARTUP_TIME "
410      SQL = SQL & "  FROM V$INSTANCE "
420      SQL = SQL & "UNION ALL "
430      SQL = SQL & "SELECT 'Status: ' || STATUS || '  Shutdown_Pending: ' || SHUTDOWN_PENDING || '  Database_Status: ' || DATABASE_STATUS "
440      SQL = SQL & "  FROM V$INSTANCE "
450      SQL = SQL & "UNION ALL "
460      SQL = SQL & "SELECT '.' "
470      SQL = SQL & "  FROM DUAL "
480      SQL = SQL & "UNION ALL "
490      SQL = SQL & "SELECT 'Conexiones actuales a Oracle:' "
500      SQL = SQL & "  FROM DUAL "
510      SQL = SQL & "UNION ALL "
520      SQL = SQL & "SELECT 'Osuser: ' || OSUSER || '  Username: ' || USERNAME || '  Machine: ' || MACHINE || '  Program: ' || PROGRAM "
530      SQL = SQL & "  FROM V$SESSION "
540      SQL = SQL & " WHERE USERNAME <> ' ' "
550      SQL = SQL & "UNION ALL "
560      SQL = SQL & "SELECT '.' "
570      SQL = SQL & "  FROM DUAL "
580      SQL = SQL & "UNION ALL "
590      SQL = SQL & "SELECT 'Tablespaces:' "
600      SQL = SQL & "  FROM DUAL "
610      SQL = SQL & "UNION ALL "
620      SQL = SQL & "SELECT      'Tablespace ' "
630      SQL = SQL & "         || T.TABLESPACE_NAME "
640      SQL = SQL & "         || '  Tamaño: ' "
650      SQL = SQL & "         || ROUND (MAX (D.BYTES) / 1024 / 1024, 2) "
660      SQL = SQL & "         || '  Usados: ' "
670      SQL = SQL & "         || ROUND ((MAX (D.BYTES) / 1024 / 1024) - (SUM (DECODE (F.BYTES, NULL, 0, F.BYTES)) / 1024 / 1024), 2) "
680      SQL = SQL & "         || '  Libres: ' "
690      SQL = SQL & "         || ROUND (SUM (DECODE (F.BYTES, NULL, 0, F.BYTES)) / 1024 / 1024, 2) "
700      SQL = SQL & "         || '  Archivo: ' "
710      SQL = SQL & "         || SUBSTR (D.FILE_NAME, 1, 80) "
720      SQL = SQL & "    FROM DBA_FREE_SPACE F, DBA_DATA_FILES D, DBA_TABLESPACES T "
730      SQL = SQL & "   WHERE T.TABLESPACE_NAME = D.TABLESPACE_NAME "
740      SQL = SQL & "     AND F.TABLESPACE_NAME(+) = D.TABLESPACE_NAME "
750      SQL = SQL & "     AND F.FILE_ID(+) = D.FILE_ID "
760      SQL = SQL & "     AND (   T.TABLESPACE_NAME = 'ADMINISTRACION' "
770      SQL = SQL & "          OR T.TABLESPACE_NAME = 'INDICES' "
780      SQL = SQL & "          OR T.TABLESPACE_NAME = 'CEREALES' "
790      SQL = SQL & "          OR T.TABLESPACE_NAME = 'CONTABILIDAD') "
800      SQL = SQL & "GROUP BY T.TABLESPACE_NAME, D.FILE_NAME, T.PCT_INCREASE, T.STATUS "
810      SQL = SQL & "UNION ALL "
820      SQL = SQL & "SELECT '.' "
830      SQL = SQL & "  FROM DUAL "
840      SQL = SQL & "UNION ALL "
850      SQL = SQL & "SELECT 'Configuracion Oracle:' "
860      SQL = SQL & "  FROM DUAL "
870      SQL = SQL & "UNION ALL "
         SQL = SQL & " SELECT * FROM ("
880      SQL = SQL & "SELECT 'Nombre: ' || NAME || '  Valor: ' || VALUE || '  Descripcion: ' || DESCRIPTION "
890      SQL = SQL & "  FROM V$SYSTEM_PARAMETER "
900      SQL = SQL & " WHERE VALUE IS NOT NULL "
910      SQL = SQL & "   AND VALUE <> '0' "
         SQL = SQL & " ORDER BY NAME) "

920      rst.Open SQL, cnn
         Set itmX = ListView3.ListItems.Add
         itmX.Text = strGuiones
930      Set itmX = ListView3.ListItems.Add
940      itmX.Text = "Version Oracle:"
         Set itmX = ListView3.ListItems.Add
         itmX.Text = strGuiones
        
950      Set itmX = ListView3.ListItems.Add
960      itmX.Text = ""
   
970      Do While Not rst.EOF
980         Set itmX = ListView3.ListItems.Add
990         itmX.Text = rst("banner").Value
1000        rst.MoveNext
1010     Loop
1020     rst.Close
   
1030     Exit Sub
   
GestErr:
1040     Me.MousePointer = vbNormal
1050     MsgBox "[VersionOracle]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Licencias()

   On Error GoTo GestErr
      
   SQL = "SELECT EMP_DESCRIPCION FROM EMPRESAS WHERE EMP_CODIGO_EMPRESA = 'ALG'"
   rst.Open SQL, cnn
   Do While Not rst.EOF
      Set itmX = lvwEmpresas.ListItems.Add
      itmX.Text = "Empresa: " & rst("EMP_DESCRIPCION").Value
      strEmpresa = Replace(Trim(rst("EMP_DESCRIPCION").Value), " ", "_")
      rst.MoveNext
   Loop
   rst.Close
   
   Titulo "Licencias:"
   Set itmX = lvwEmpresas.ListItems.Add
   itmX.Text = "Server: " & GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Algoritmo\Environment\Application Server", "Server Update Root", REG_SZ, "", False)

   Set itmX = lvwEmpresas.ListItems.Add
   itmX.Text = "Cantidad de Licencias: " & GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Algoritmo\Environment\Application Server", "Licencias", REG_SZ, "", False)

   Set itmX = lvwEmpresas.ListItems.Add
   itmX.Text = "Activation Key: " & GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Algoritmo\Environment\Application Server", "Activation Key", REG_SZ, "", False)

   Set itmX = lvwEmpresas.ListItems.Add
   itmX.Text = "Serial Number: " & GenerateAK
   Set itmX = lvwEmpresas.ListItems.Add
   itmX.Text = ""
   
   Exit Sub
   
GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[Licencias]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub VersionSistema()
      Dim strServer As String
      Dim fso As Scripting.FileSystemObject
      Dim linea As Integer

10       On Error Resume Next
         
20       Set fso = CreateObject("Scripting.FileSystemObject")
         
30       strServer = GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Algoritmo\Environment\Application Server", "Remote Host", REG_SZ, "", False)
40       If Len(strServer) = 0 Then strServer = "Server"
50       Titulo "Version del Sistema:"
         
60       Set itmX = lvwEmpresas.ListItems.Add
70       itmX.Text = "Version Registro Local: " & GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo", "VersionProducto", REG_SZ, "", False)
         
80       SQL = "SELECT NRO_VERSION From VERSION_PRODUCTO "
90       rst.Open SQL, cnn
100         Do While Not rst.EOF
110         Set itmX = lvwEmpresas.ListItems.Add
120         itmX.Text = "Version_Producto Oracle: " & rst("NRO_VERSION").Value
130         rst.MoveNext
140      Loop
150      rst.Close
         
160      Set itmX = lvwEmpresas.ListItems.Add
170      itmX.Text = "Version.dat:"
180      If fso.FileExists("\\" & strServer & "\d\Algoritmo\Version.dat") Then
190         Open "\\" & strServer & "\d\Algoritmo\Version.dat" For Input As 1
200         While Not EOF(1)
210            Line Input #1, temp
220            If Trim(temp) <> "" Then
230               Set itmX = lvwEmpresas.ListItems.Add
240               itmX.Text = Trim(temp)
250            End If
260         Wend
270         Close #1
280      End If

290      Set itmX = lvwEmpresas.ListItems.Add
300      itmX.Text = ""
         
         Set itmX = lvwEmpresas.ListItems.Add
         itmX.Text = strGuiones
         Set itmX = lvwEmpresas.ListItems.Add
         itmX.Text = "Registro del Sistema:"
         Set itmX = lvwEmpresas.ListItems.Add
         itmX.Text = strGuiones
         
         If fso.FileExists("c:\windows\temp\registroX32.reg") Then
            linea = 0
            Open "c:\windows\temp\registroX32.reg" For Input As 1
            While Not EOF(1)
               Line Input #1, temp
               If linea > 0 Then
                  Set itmX = lvwEmpresas.ListItems.Add
                  itmX.Text = Trim(temp)
               End If
               linea = linea + 1
            Wend
            Close #1
         End If
         If fso.FileExists("c:\windows\temp\registroX64.reg") Then
            linea = 0
            Open "c:\windows\temp\registroX64.reg" For Input As 1
            While Not EOF(1)
               Line Input #1, temp
               If linea > 0 Then
                  Set itmX = lvwEmpresas.ListItems.Add
                  itmX.Text = Trim(temp)
               End If
               linea = linea + 1
            Wend
            Close #1
         End If
         
         Kill "c:\windows\temp\registroX32.reg"
         Kill "c:\windows\temp\registroX64.reg"

310      Exit Sub
GestErr:
320      Me.MousePointer = vbNormal
330      MsgBox "[VersionSistema]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Sucursales()

10       On Error GoTo GestErr
   
20       SQL = " SELECT SUC_CODIGO, SUC_DESCRIPCION, SUC_ACTIVA, SUC_ULTIMOLOTE_GEN_IMP, SUC_ULTIMA_TRANSACCION, SUC_ULTIMO_NRO_AUDITORIA, "
30       SQL = SQL & "       SUC_TURNO_SURTIDORES "
40       SQL = SQL & "  FROM SUCURSALES "
50       rst.Open SQL, cnn
60          Do While Not rst.EOF
70          Set itmX = ListView2.ListItems.Add
80          itmX.Text = rst("SUC_CODIGO").Value
90          itmX.SubItems(1) = IIf(IsNull(rst("SUC_DESCRIPCION").Value), "", rst("SUC_DESCRIPCION").Value)
100         itmX.SubItems(2) = IIf(IsNull(rst("SUC_ACTIVA").Value), "", rst("SUC_ACTIVA").Value)
110         itmX.SubItems(3) = IIf(IsNull(rst("SUC_ULTIMOLOTE_GEN_IMP").Value), "", rst("SUC_ULTIMOLOTE_GEN_IMP").Value)
120         itmX.SubItems(4) = IIf(IsNull(rst("SUC_ULTIMA_TRANSACCION").Value), "", rst("SUC_ULTIMA_TRANSACCION").Value)
130         itmX.SubItems(5) = IIf(IsNull(rst("SUC_ULTIMO_NRO_AUDITORIA").Value), "", rst("SUC_ULTIMO_NRO_AUDITORIA").Value)
            'itmX.SubItems(6) = IIf(IsNull(rst("SUC_TURNO_SURTIDORES").Value), 0, rst("SUC_TURNO_SURTIDORES").Value)
140         rst.MoveNext
150      Loop
160      rst.Close
   
170      Exit Sub
   
GestErr:
180      Me.MousePointer = vbNormal
190      MsgBox "[Sucursales]" & vbCrLf & Err.Description & Erl
End Sub

' Función que Devuelve un String con el símbolo Configuracion Regional
Private Function Obtener_Simbolo(Valor As Long) As String
   
   Dim Simbolo As String
   
   Dim r1 As Long
   Dim r2 As Long
   Dim p As Integer
   Dim Locale As Long
   
   Locale = GetUserDefaultLCID()
   r1 = GetLocaleInfo(Locale, Valor, vbNullString, 0)
   
   'buffer
   Simbolo = String$(r1, 0)
   
   'En esta llamada devuelve el símbolo en el Buffer
   r2 = GetLocaleInfo(Locale, Valor, Simbolo, r1)
   
   'Localiza el espacio nulo de la cadena para eliminarla
   p = InStr(Simbolo, Chr$(0))
   
   If p > 0 Then
      'Elimina los nulos
      Obtener_Simbolo = Left$(Simbolo, p - 1)
   End If
   
End Function

Private Sub ConfiguracionRegional()
   
10       On Error GoTo GestErr
   
20       Set itmX = ListView4.ListItems.Add
30       itmX.Text = "Configuración Regional"

40       Set itmX = ListView4.ListItems.Add
50       itmX.Text = ""
   
   
60       Set itmX = ListView4.ListItems.Add
70       itmX.Text = "Ubicación: "
80       itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SCOUNTRY)
   
90       Set itmX = ListView4.ListItems.Add
100      itmX.Text = "Código del País: "
110      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_ICOUNTRY)
   
120      Set itmX = ListView4.ListItems.Add
130      itmX.Text = "Idioma: "
140      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SLANGUAGE)
   
150      Set itmX = ListView4.ListItems.Add
160      itmX.Text = "Separador decimal:"
170      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SDECIMAL)
   
180      Set itmX = ListView4.ListItems.Add
190      itmX.Text = "Separador de miles: "
200      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_STHOUSAND)

210      Set itmX = ListView4.ListItems.Add
220      itmX.Text = "Símbolo de moneda: "
230      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SCURRENCY)

240      Set itmX = ListView4.ListItems.Add
250      itmX.Text = "Separador de Hora: "
260      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_STIME)

270      Set itmX = ListView4.ListItems.Add
280      itmX.Text = "Página de códigos predeterminada: "
290      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_IDEFAULTCODEPAGE)

300      Set itmX = ListView4.ListItems.Add
310      itmX.Text = "Código predeterminado del país: "
320      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_IDEFAULTCOUNTRY)

330      Set itmX = ListView4.ListItems.Add
340      itmX.Text = "Número de dígitos fraccionarios: "
350      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_IDIGITS)

360      Set itmX = ListView4.ListItems.Add
370      itmX.Text = "Designador de AM: "
380      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_S1159)

390      Set itmX = ListView4.ListItems.Add
400      itmX.Text = "Designador de PM: "
410      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_S2359)

420      Set itmX = ListView4.ListItems.Add
430      itmX.Text = "Separador de fecha: "
440      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SDATE)

450      Set itmX = ListView4.ListItems.Add
460      itmX.Text = "Separador de elementos de lista: "
470      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SLIST)

480      Set itmX = ListView4.ListItems.Add
490      itmX.Text = "Cadena de formato de fecha larga: "
500      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SLONGDATE)

510      Set itmX = ListView4.ListItems.Add
520      itmX.Text = "ASCII 0-9 nativo: "
530      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SNATIVEDIGITS)

540      Set itmX = ListView4.ListItems.Add
550      itmX.Text = "Cadena de formato de fecha corta: "
560      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_SSHORTDATE)

570      Set itmX = ListView4.ListItems.Add
580      itmX.Text = "Cadena de formato de hora: "
590      itmX.SubItems(1) = Obtener_Simbolo(LOCALE_STIMEFORMAT)
   
600      Exit Sub
   
GestErr:
610      Me.MousePointer = vbNormal
620      MsgBox "[ConfiguracionRegional]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub TimerCarga_Timer()
   CargarTodo
End Sub
