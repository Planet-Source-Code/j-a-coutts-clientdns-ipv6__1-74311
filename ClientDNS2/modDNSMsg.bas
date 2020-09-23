Attribute VB_Name = "modDNSMsg"
Option Explicit

Private Const ERROR_BUFFER_OVERFLOW = 111&
Private Const MAX_HOSTNAME_LEN = 128
Private Const MAX_DOMAIN_NAME_LEN = 128
Private Const MAX_SCOPE_ID_LEN = 256

'    Private Declare Function GetNetworkParams Lib "iphlpapi.dll" (pFixedInfo As Any, pOutBufLen As Long) As Long

Public m_objHeader As DNSHeader
Public m_objQuestion As DNSQuestion
Public m_objRRec As DNSRRec
Public m_objSOARec As SOAData
Public RRArray() As Variant
Private RRPntr As Integer

Private Type IP_ADDR_STRING
    Next As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Private Type FIXED_INFO
    HostName(MAX_HOSTNAME_LEN + 3) As Byte
    DomainName(MAX_DOMAIN_NAME_LEN + 3) As Byte
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId(MAX_SCOPE_ID_LEN + 3) As Byte
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type

Public Type DNSHeader
    MsgID           As Integer
    QR              As Boolean
    OpCode          As Byte
    AA              As Boolean
    TC              As Boolean
    RD              As Boolean
    RA              As Boolean
    Z               As Byte
    RCode           As Byte
    QDCount         As Integer
    ANCount         As Integer
    NSCount         As Integer
    ARCount         As Integer
End Type

Public Type DNSQuestion
    QName           As String
    QType           As Integer
    QClass          As Integer
End Type

Public Type DNSRRec
    RName           As String
    RType           As Integer
    RClass          As Integer
    TTL             As Long
    RDLength        As Integer
    RData           As String
End Type

Public Type SOAData
    MName           As String
    RName           As String
    Serial          As Long
    Refresh         As Long
    Retry           As Long
    Expire          As Long
    Minimum         As Long
End Type

Public Type MXData
    Preference      As Integer
    ExChange           As String
End Type

Public Enum DnsTypeField
    DNS_TYPE_A = 1
    DNS_TYPE_NS
    DNS_TYPE_MD
    DNS_TYPE_MF
    DNS_TYPE_CNAME
    DNS_TYPE_SOA
    DNS_TYPE_MB
    DNS_TYPE_MG
    DNS_TYPE_MR
    DNS_TYPE_NULL
    DNS_TYPE_WKS
    DNS_TYPE_PTR
    DNS_TYPE_HINFO
    DNS_TYPE_MINFO
    DNS_TYPE_MX
    DNS_TYPE_TXT
    DNS_TYPE_AAAA = 28
    DNS_TYPE_SPF = 99
    DNS_TYPE_AXFR = 252
    DNS_TYPE_MAILB = 253
    DNS_TYPE_MAILA = 254
    DNS_TYPE_ALL = 255
End Enum

Public Enum DnsClassField
    DNS_CLASS_IN = 1
    DNS_CLASS_CS
    DNS_CLASS_CH
    DNS_CLASS_HS
End Enum

Public Type OSVERSIONINFO    '148 bytes
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Const REG_SZ As Long = 1
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Public Const REG_DWORD As Long = 4
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const KEY_ALL_ACCESS = &HF003F
Public Const KEY_READ = &H20019
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF
Public Const MIN_SOCKETS_REQD = 1
'Public Const SOCKET_ERROR = -1

Public myVer As OSVERSIONINFO
Public OpSys As Integer

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long


Public Function AddDotSep(strSource As String, intIndex As Integer) As String
    Dim M%, N%
    Dim strTemp As String
    Dim tmpPoint As Integer
    On Error GoTo AddDotSepErr
    tmpPoint = intIndex 'Temporarily store pointer
    'Figure out logical end of string
    Do Until Asc(Mid$(strSource, intIndex)) >= 192 Or Asc(Mid$(strSource, intIndex)) = 0
        intIndex = intIndex + Asc(Mid$(strSource, intIndex)) + 1
    Loop
    If Asc(Mid$(strSource, intIndex)) >= 192 Then
        intIndex = intIndex + 2 'Pointer used
    Else
        intIndex = intIndex + 1
    End If
    strTemp = Mid$(strSource, tmpPoint, intIndex - tmpPoint)
    Debug.Print strTemp
    N% = 1
    M% = Asc(Mid$(strTemp, N%, 1))
    If M% >= 192 Then 'Compression pointer used
        tmpPoint = (M% - 192) * 256 ^ 1 + Asc(Mid$(strTemp, N% + 1, 1)) + 1
        strTemp = Mid$(strSource, tmpPoint, InStr(tmpPoint, strSource, Chr$(0)) - tmpPoint) + Chr$(0)
        M% = Asc(Mid$(strTemp, N%, 1))
    End If
    Do Until M% = 0
        N% = N% + M% + 1
        M% = Asc(Mid$(strTemp, N%, 1))
        'Check if compression pointer used
        If M% >= 192 Then
            tmpPoint = (M% - 192) * 256 ^ 1 + Asc(Mid$(strTemp, N% + 1, 1)) + 1
            strTemp = Left$(strTemp, N% - 1) + Mid$(strSource, tmpPoint, InStr(tmpPoint, strSource, Chr$(0)) - tmpPoint) + Chr$(0)
            M% = Asc(Mid$(strTemp, N%, 1))
        End If
        Mid$(strTemp, N%, 1) = "."
    Loop
    If Len(strTemp) < 2 Then 'Account for one byte records (&H00)
        AddDotSep = ""
    Else
        AddDotSep = Mid$(strTemp, 2, N% - 2) 'Remove first and last char
    End If
    Exit Function
AddDotSepErr:
    AddDotSep = "Error!"
End Function
Public Function ReverseIP(IPaddr As String) As String
    'Return the reverse of the IP sent
    Dim strTemp As String
    Dim IP(4) As Byte
    Dim M%, N%
    On Error GoTo NotValidIP
    strTemp = Replace(IPaddr, ".", "_")
    'Separate number into 4 byte values
    While N% < 4
        IP(N%) = Val(Mid$(strTemp, M% + 1))
        M% = InStr(M% + 1, strTemp, "_")
        If M% = 0 And N% < 3 Then GoTo NotValidIP
        N% = N% + 1
    Wend
    ReverseIP = CStr(IP(3)) + "." + CStr(IP(2)) + "." + CStr(IP(1)) + "." + CStr(IP(0))
    Exit Function
NotValidIP:
    ReverseIP = "Error"
End Function
Public Function GetOSVersion() As String
    Dim dl&
    myVer.dwOSVersionInfoSize = 148
    dl& = GetVersionEx(myVer)
    If myVer.dwPlatformId = 1 Then
        OpSys = 1 'Windows 9x
        GetOSVersion = "Windows 9x"
    ElseIf myVer.dwPlatformId = 2 Then
        If myVer.dwMajorVersion < 5 Then
            OpSys = 2 'NT
            GetOSVersion = "Windows NT"
        ElseIf myVer.dwMinorVersion = 0 Then
            OpSys = 3 '2000
            GetOSVersion = "Windows 2000"
        ElseIf myVer.dwMinorVersion = 1 Then
            OpSys = 4 'XP
            GetOSVersion = "Windows XP"
        Else
            OpSys = 0 'Unknown
            GetOSVersion = ""
        End If
    End If
End Function
'Returns all listed subkeys in a double null terminated string
Public Function RegSubKey(ByVal sKeyBase As Long, ByVal sKeyName As String) As String
    Dim Idx As Long
    Dim lRetVal As Long
    Dim hKey As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    Dim sValueLen As Long
    Dim sClass As String
    Dim sClassLen As Long
    Dim fT As FILETIME
    Dim strTemp As String
    'Open key
    lRetVal = RegOpenKeyEx(sKeyBase, sKeyName, 0, KEY_READ, hKey)
    If lRetVal <> 0 Then
        MsgBox "No Such Key as " + sKeyName
        Exit Function
    End If
    On Error GoTo RegSubKeyErr
    Do
        sValue = String$(2048, 0)
        sValueLen = Len(sValue)
        sClass = String$(2048, 0)
        sClassLen = Len(sClass)
        lRetVal = RegEnumKeyEx(hKey, Idx, sValue, sValueLen, 0&, sClass, sClassLen, fT)
        If lRetVal = 0 Then
            strTemp = strTemp + Left$(sValue, sValueLen + 1)
            Idx = Idx + 1
        End If
    Loop While lRetVal = 0
RegSubKeyExit:
    RegSubKey = strTemp
    RegCloseKey (hKey)
    Exit Function
RegSubKeyErr:
    Resume RegSubKeyExit
End Function
Public Function RegQuery(sKeyBase As Long, sKeyName As String, sValueName As String) As String
    Dim lRetVal As Long
    Dim hKey As Long
    Dim vValue As Variant
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    lRetVal = RegOpenKeyEx(sKeyBase, sKeyName, 0, KEY_READ, hKey)
    If lRetVal <> 0 Then
        MsgBox "No Such Key as " + sKeyName
        Exit Function
    End If
    On Error GoTo RegQueryError
    lrc = RegQueryValueExNULL(hKey, sValueName, 0&, lType, 0&, cch)
    If lrc <> 0 Then Error 5
    Select Case lType
        Case REG_SZ, REG_MULTI_SZ
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(hKey, sValueName, 0&, lType, sValue, cch)
            If lrc = 0 Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        Case REG_DWORD
            lrc = RegQueryValueExLong(hKey, sValueName, 0&, lType, lValue, cch)
            If lrc = 0 Then vValue = lValue
        Case Else
            lrc = -1
    End Select
RegQueryExit:
    RegQuery = vValue
    RegCloseKey (hKey)
    Exit Function
RegQueryError:
    Resume RegQueryExit
End Function


Public Function GetDnsServerAddress() As String
    Dim M%, N%
    Dim strTemp As String
    Dim KeyName As String
    Dim sInterface As String
    Dim sInterfaceKey As String
    Dim DNSServer As String
'    Dim lngRetValue     As Long
'    Dim lngBufferSize   As Long
'    Dim udtNetworkInfo  As FIXED_INFO
    'Get OS Info
    strTemp = GetOSVersion
    Select Case OpSys
        Case 0 'Unknown
            MsgBox "Unknown Operating System"
        Case 1 'Windows 9x
            KeyName = "System\CurrentControlSet\Services\VxD\MSTCP"
            DNSServer = RegQuery(HKEY_LOCAL_MACHINE, KeyName, "NameServer")
            N% = InStr(DNSServer, ",") 'Use first one only
            If N% > 0 Then DNSServer = Left$(DNSServer, N% - 1)
        Case 2 'Win NT
            KeyName = "System\CurrentControlSet\Services\Tcpip\Parameters"
            DNSServer = RegQuery(HKEY_LOCAL_MACHINE, KeyName, "NameServer")
            N% = InStr(DNSServer, " ") 'Use first one only
            If N% > 0 Then DNSServer = Left$(DNSServer, N% - 1)
        Case 3, 4 'Win 2000 or Win XP
            KeyName = "System\CurrentControlSet\Services"
            'Get list of interfaces
            strTemp = RegSubKey(HKEY_LOCAL_MACHINE, KeyName + "\Tcpip\Parameters\Adapters")
            N% = InStr(strTemp, Chr$(0))
            If N% = 0 Then
                MsgBox ("No Ethernet Interfaces Found!")
                End
            End If
            M% = 1
            Do Until N% = 0
                sInterface = Mid$(strTemp, M%, N% - M%)
                sInterfaceKey = RegQuery(HKEY_LOCAL_MACHINE, KeyName + "\Tcpip\Parameters\Adapters\" + sInterface, "IpConfig")
                If frmDnsCl.chkIPv6 Then
                    DNSServer = RegQuery(HKEY_LOCAL_MACHINE, KeyName + "\Tcpip6\Parameters\Interfaces\" + sInterface, "NameServer")
                Else
                    DNSServer = RegQuery(HKEY_LOCAL_MACHINE, KeyName + "\" + sInterfaceKey, "NameServer")
                End If
                'Use first one if more than 1
                If InStr(DNSServer, ",") > 0 Then DNSServer = Left$(DNSServer, InStr(DNSServer, ",") - 1)
                If Len(DNSServer) > 0 Then Exit Do
                M% = N% + 1
                N% = InStr(M%, strTemp, Chr$(0))
            Loop
    End Select
    GetDnsServerAddress = DNSServer
'    If Len(strTemp) > 0 Then
'        lngRetValue = GetNetworkParams(0&, lngBufferSize)
'        If lngRetValue = ERROR_BUFFER_OVERFLOW Then
'            ReDim arrBuffer(lngBufferSize - 1)
'            lngRetValue = GetNetworkParams(arrBuffer(0), lngBufferSize)
'            If lngRetValue = 0 Then
'                CopyMemory udtNetworkInfo, arrBuffer(0), lngBufferSize
'                GetDnsServerAddress = udtNetworkInfo.DnsServerList.IpAddress
'            End If
'        End If
'    Else
'        MsgBox "Unknown Operating System"
'    End If
End Function
Public Function AddToArray(R%, RRec As DNSRRec) As Boolean
    'Expand array as required
    If RRPntr > UBound(RRArray, 2) Then
        ReDim Preserve RRArray(6, RRPntr)
    End If
    With RRec
        RRArray(0, RRPntr) = R% 'Section
        RRArray(1, RRPntr) = .RName
        RRArray(2, RRPntr) = .RType
        RRArray(3, RRPntr) = .RClass
        RRArray(4, RRPntr) = .TTL
        RRArray(5, RRPntr) = .RDLength
        RRArray(6, RRPntr) = .RData
    End With
    RRPntr = RRPntr + 1
End Function

Public Sub ClearArray()
    Dim N% 'Resets the RDLength field in each record
    Do Until N% >= RRPntr
        RRArray(5, N%) = 0
        N% = N% + 1
    Loop
    RRPntr = 0 'Resets the pointer
End Sub


Public Function GetTypeName(lngTypeValue) As String
    Select Case lngTypeValue
        Case DNS_TYPE_A
            GetTypeName = "A"
        Case DNS_TYPE_NS
            GetTypeName = "NS"
        Case DNS_TYPE_MD
            GetTypeName = "MD"
        Case DNS_TYPE_MF
            GetTypeName = "MF"
        Case DNS_TYPE_CNAME
            GetTypeName = "CNAME"
        Case DNS_TYPE_SOA
            GetTypeName = "SOA"
        Case DNS_TYPE_MB
            GetTypeName = "MB"
        Case DNS_TYPE_MG
            GetTypeName = "MG"
        Case DNS_TYPE_MR
            GetTypeName = "MR"
        Case DNS_TYPE_NULL
            GetTypeName = "NULL"
        Case DNS_TYPE_WKS
            GetTypeName = "WKS"
        Case DNS_TYPE_PTR
            GetTypeName = "PTR"
        Case DNS_TYPE_HINFO
            GetTypeName = "HINFO"
        Case DNS_TYPE_MINFO
            GetTypeName = "MINFO"
        Case DNS_TYPE_MX
            GetTypeName = "MX"
        Case DNS_TYPE_TXT
            GetTypeName = "TXT"
        Case DNS_TYPE_SPF
            GetTypeName = "SPF"
        Case DNS_TYPE_AAAA
            GetTypeName = "AAAA"
        Case DNS_TYPE_AXFR
            GetTypeName = "AXFR"
        Case DNS_TYPE_MAILB
            GetTypeName = "MAILB"
        Case DNS_TYPE_MAILA
            GetTypeName = "MAILA"
        Case DNS_TYPE_ALL
            GetTypeName = "*"
        Case Else
            GetTypeName = "UNKNOWN"
    End Select
End Function
