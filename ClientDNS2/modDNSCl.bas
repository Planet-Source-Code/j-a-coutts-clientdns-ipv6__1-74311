Attribute VB_Name = "modDNSCl"
Option Explicit
Public Const gAppName As String = "DNSClient"
Public Const gsDelimiter As String = "|"

Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Function GetTypeInfo(ByVal lngTypeValue As DnsTypeField, ByRef strMnemonicValue As String, ByRef strDescription As String) As Boolean
    Select Case lngTypeValue
        Case DNS_TYPE_A
            strMnemonicValue = "A"
            strDescription = "IPv4 Host address"
        Case DNS_TYPE_NS
            strMnemonicValue = "NS"
            strDescription = "Authoritative name server"
        Case DNS_TYPE_MD
            strMnemonicValue = "MD"
            strDescription = "Mail destination (Obsolete - use MX)"
        Case DNS_TYPE_MF
            strMnemonicValue = "MF"
            strDescription = "Mail forwarder (Obsolete - use MX)"
        Case DNS_TYPE_CNAME
            strMnemonicValue = "CNAME"
            strDescription = "The canonical name for an alias"
        Case DNS_TYPE_SOA
            strMnemonicValue = "SOA"
            strDescription = "Marks the start of a zone of authority"
        Case DNS_TYPE_MB
            strMnemonicValue = "MB"
            strDescription = "Mailbox domain name (EXPERIMENTAL)"
        Case DNS_TYPE_MG
            strMnemonicValue = "MG"
            strDescription = "Mail group member (EXPERIMENTAL)"
        Case DNS_TYPE_MR
            strMnemonicValue = "MR"
            strDescription = "mail rename domain name (EXPERIMENTAL)"
        Case DNS_TYPE_NULL
            strMnemonicValue = "NULL"
            strDescription = "Null RR (EXPERIMENTAL)"
        Case DNS_TYPE_WKS
            strMnemonicValue = "WKS"
            strDescription = "Well known service description"
        Case DNS_TYPE_PTR
            strMnemonicValue = "PTR"
            strDescription = "Domain name pointer"
        Case DNS_TYPE_HINFO
            strMnemonicValue = "HINFO"
            strDescription = "Host information"
        Case DNS_TYPE_MINFO
            strMnemonicValue = "MINFO"
            strDescription = "Mailbox or mail list information"
        Case DNS_TYPE_MX
            strMnemonicValue = "MX"
            strDescription = "Mail exchange"
        Case DNS_TYPE_TXT
            strMnemonicValue = "TXT"
            strDescription = "Text strings"
        Case DNS_TYPE_SPF
            strMnemonicValue = "SPF"
            strDescription = "Sender Policy Framework"
        Case DNS_TYPE_AAAA
            strMnemonicValue = "AAAA"
            strDescription = "IPv6 Host Address"
        Case DNS_TYPE_AXFR
            strMnemonicValue = "AXFR"
            strDescription = "Request for a transfer of an entire zone"
        Case DNS_TYPE_MAILB
            strMnemonicValue = "MAILB"
            strDescription = "Request for mailbox-related records (MB, MG or MR)"
        Case DNS_TYPE_MAILA
            strMnemonicValue = "MAILA"
            strDescription = "Request for mail agent RRs (Obsolete - see MX)"
        Case DNS_TYPE_ALL
            strMnemonicValue = "*"
            strDescription = "Request for all records"
    Case Else
            strMnemonicValue = "UNKNOWN"
            strDescription = "UNKNOWN"
    End Select
End Function

Public Function GetResponseCodeInfo(ByVal lngResponseCode As Long) As String
    Select Case lngResponseCode
        Case 0
            GetResponseCodeInfo = "No error condition"
        Case 1
            GetResponseCodeInfo = "Format error"
        Case 2
            GetResponseCodeInfo = "Server failure"
        Case 3
            GetResponseCodeInfo = "Name Error"
        Case 4
            GetResponseCodeInfo = "Not Implemented"
        Case 5
            GetResponseCodeInfo = "Refused"
        Case Else
            GetResponseCodeInfo = "Unknown response"
    End Select
End Function

Public Sub GetOPCodeInfo(ByVal lngCode As Long, ByRef strMnemonicValue As String, ByRef strDescription As String)
    Select Case lngCode
        Case 0
            strMnemonicValue = "QUERY"
            strDescription = "Standard query"
        Case 1
            strMnemonicValue = "IQUERY"
            strDescription = "Inverse query"
        Case 2
            strMnemonicValue = "STATUS"
            strDescription = "Server status request"
        Case Else
            strMnemonicValue = "Unknown"
            strDescription = "Unknown"
    End Select
End Sub



Public Function GetClassInfo(ByVal lngClassValue As DnsClassField, ByRef strMnemonicValue As String, ByRef strDescription As String) As Boolean
    Select Case lngClassValue
        Case DNS_CLASS_IN
            strMnemonicValue = "IN"
            strDescription = "The Internet"
        Case DNS_CLASS_CS
            strMnemonicValue = "CS"
            strDescription = "the CSNET class (Obsolete - used only for examples in some obsolete RFCs)"
        Case DNS_CLASS_CH
            strMnemonicValue = "CH"
            strDescription = "the CHAOS class"
        Case DNS_CLASS_HS
            strMnemonicValue = "HS"
            strDescription = "Hesiod [Dyer 87]"
    End Select
End Function
