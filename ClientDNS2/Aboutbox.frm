VERSION 5.00
Begin VB.Form AboutBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About DNS Client"
   ClientHeight    =   4755
   ClientLeft      =   1365
   ClientTop       =   1425
   ClientWidth     =   5820
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Aboutbox.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4755
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSysInfo 
      Interval        =   1
      Left            =   5280
      Top             =   240
   End
   Begin VB.PictureBox Pic_ApplicationIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   255
      Picture         =   "Aboutbox.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton Cmd_OK 
      Cancel          =   -1  'True
      Caption         =   "EXIT"
      Height          =   360
      Left            =   4920
      TabIndex        =   6
      Top             =   4320
      Width           =   795
   End
   Begin VB.Label lblR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PageFile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label LblMem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "MEMORY Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1680
      TabIndex        =   12
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Shape shpBar 
      BackStyle       =   1  'Opaque
      DrawMode        =   7  'Invert
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   200
      Index           =   3
      Left            =   1200
      Top             =   4455
      Width           =   135
   End
   Begin VB.Label lblResInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "PageFile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   1200
      TabIndex        =   11
      Top             =   4440
      Width           =   3105
   End
   Begin VB.Shape shpFrame 
      Height          =   255
      Index           =   3
      Left            =   1180
      Top             =   4425
      Width           =   3135
   End
   Begin VB.Shape shpBar 
      BackStyle       =   1  'Opaque
      DrawMode        =   7  'Invert
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   200
      Index           =   1
      Left            =   1200
      Top             =   3735
      Width           =   135
   End
   Begin VB.Shape shpBar 
      BackStyle       =   1  'Opaque
      DrawMode        =   7  'Invert
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   2
      Left            =   1200
      Top             =   4095
      Width           =   135
   End
   Begin VB.Label lblResInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      Caption         =   "Physical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   3720
      Width           =   3105
   End
   Begin VB.Label lblResInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "Virtual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   1200
      TabIndex        =   10
      Top             =   4080
      Width           =   3105
   End
   Begin VB.Shape shpFrame 
      Height          =   255
      Index           =   2
      Left            =   1180
      Top             =   4065
      Width           =   3135
   End
   Begin VB.Shape shpFrame 
      Height          =   255
      Index           =   1
      Left            =   1180
      Top             =   3705
      Width           =   3135
   End
   Begin VB.Line lin_HorizontalLine1 
      BorderWidth     =   2
      X1              =   375
      X2              =   4410
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Label Lbl_Manta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DNS CLIENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   990
      TabIndex        =   1
      Top             =   120
      Width           =   3465
   End
   Begin VB.Label Lbl_Version 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Version 2.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1470
   End
   Begin VB.Label Lbl_ComPany 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JAC Computing, Vernon, BC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   390
      TabIndex        =   3
      Top             =   840
      Width           =   4365
   End
   Begin VB.Label Lbl_Info 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2235
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1110
      Width           =   2415
   End
   Begin VB.Label Lbl_InfoValues 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   2760
      TabIndex        =   0
      Top             =   1110
      Width           =   2895
   End
   Begin VB.Label lblR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Physical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Virtual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
Const VER_PLATFORM_WIN32_NT = 2
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const PROCESSOR_ALPHA_21064 = 21064
Const PROCESSOR_INTEL_386 = 386
Const PROCESSOR_INTEL_486 = 486
Const PROCESSOR_INTEL_PENTIUM = 586
Const PROCESSOR_MIPS_R4000 = 4000

'Private Declare Function GetWinFlags Lib "kernel" () As integer
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal N As Integer) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetFreeSpace Lib "kernel32" (ByVal flag As Long) As Long
Private Declare Function GetFreeSystemResources Lib "user32" (ByVal fuSysResource As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Private Sub Cmd_OK_Click()
    Unload AboutBox
End Sub

Private Function DeviceColors(hDC As Long) As Double
Const PLANES = 14
Const BITSPIXEL = 12
    DeviceColors = GetDeviceCaps(hDC, PLANES) * 2 ^ GetDeviceCaps(hDC, BITSPIXEL)
End Function

Private Sub FillSysInfo()
Dim mySys As SYSTEM_INFO
Dim A$, N%, dl&
    Lbl_Version.Caption = "Version " + CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    'Operating System Info.
    Lbl_Info(0) = GetOSVersion
    Lbl_InfoValues = CStr(myVer.dwMajorVersion) + "." + CStr(myVer.dwMinorVersion) + " Build " + CStr(myVer.dwBuildNumber And &HFFFF&)
    GetSystemInfo mySys
    Lbl_Info(0) = Lbl_Info(0) + vbCrLf + "PageSize:" + vbCrLf + "Lowest Memory Address:" + vbCrLf + "Highest Memory Address:"
    Lbl_InfoValues = Lbl_InfoValues + vbCrLf + CStr(mySys.dwPageSize) + vbCrLf + "&H" + Hex$(mySys.lpMinimumApplicationAddress) + vbCrLf + "&H" + Hex$(mySys.lpMaximumApplicationAddress)
    Lbl_Info(0) = Lbl_Info(0) + vbCrLf + "Number of Processors:" + vbCrLf + "Processor:"
    Select Case mySys.dwProcessorType
        Case PROCESSOR_INTEL_386
            A$ = "Intel 386"
        Case PROCESSOR_INTEL_486
            A$ = "Intel 486"
        Case PROCESSOR_INTEL_PENTIUM
            A$ = "Intel Pentium"
        Case PROCESSOR_MIPS_R4000
            A$ = "MIPS R4000"
        Case PROCESSOR_ALPHA_21064
            A$ = "APLHA_21064"
        Case Else
            A$ = "Unknown"
    End Select
    Lbl_InfoValues = Lbl_InfoValues + vbCrLf + CStr(mySys.dwNumberOfProcessors) + vbCrLf + A$
    Lbl_Info(0) = Lbl_Info(0) + vbCrLf + "Video Resolution:" + vbCrLf + "Colors: "
    Lbl_InfoValues = Lbl_InfoValues + vbCrLf + CStr(Screen.Width \ Screen.TwipsPerPixelX) & " x " & CStr(Screen.Height \ Screen.TwipsPerPixelY) + vbCrLf + CStr(DeviceColors(hDC))
    Lbl_Info(0) = Lbl_Info(0) + vbCrLf + "Network:"
    Lbl_InfoValues = Lbl_InfoValues + vbCrLf + GetSysIni("boot.description", "network.drv")
End Sub

Private Sub Form_Load()
    Me.Move (frmDnsCl.Left + 300), (frmDnsCl.Top + 300)
    Call FillSysInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AboutBox = Nothing
End Sub

Private Function GetSysIni(Section$, key$) As String
Dim retval As String, AppName As String, worked As Long
    retval = String$(255, 0)
    worked = GetPrivateProfileString(Section$, key$, "", retval, Len(retval), "System.ini")
    If worked = 0 Then
        GetSysIni = "unknown"
    Else
        GetSysIni = Left(retval, worked)
    End If
End Function

Private Sub tmrSysInfo_Timer()
    Dim YourMemory As MEMORYSTATUS
    Dim lWidth As Integer
    Static FirstFlg As Integer
        
    YourMemory.dwLength = Len(YourMemory)
    GlobalMemoryStatus YourMemory
    If Not FirstFlg Then
        Lbl_Info(0) = Lbl_Info(0) + vbCrLf + "MB Memory:"
        Lbl_InfoValues = Lbl_InfoValues + vbCrLf + CStr(Int((YourMemory.dwTotalPhys / 1048576) + 0.5))
        FirstFlg = True
    End If
    'Update memory info
    'Check width before setting to try and cut down on screen "flashing"
    lWidth = shpFrame(1).Width * (YourMemory.dwAvailPhys / YourMemory.dwTotalPhys)
    If lWidth < 40 Then lWidth = 40
    If lWidth <> shpBar(1).Width Then
        lblResInfo(1).Caption = Int((lWidth / shpFrame(1).Width * 100) + 0.5) & "%"
        shpBar(1).Width = (lWidth)
    End If
    lWidth = shpFrame(2).Width * (YourMemory.dwAvailVirtual / YourMemory.dwTotalVirtual)
    If lWidth < 40 Then lWidth = 40
    If lWidth <> shpBar(2).Width Then
        lblResInfo(2).Caption = Int((lWidth / shpFrame(2).Width * 100) + 0.5) & "%"
        shpBar(2).Width = lWidth
    End If
    lWidth = shpFrame(3).Width * (YourMemory.dwAvailPageFile / YourMemory.dwTotalPageFile)
    If lWidth < 40 Then lWidth = 40
    If lWidth <> shpBar(3).Width Then
        lblResInfo(3).Caption = Int((lWidth / shpFrame(3).Width * 100) + 0.5) & "%"
        shpBar(3).Width = lWidth
    End If
End Sub


