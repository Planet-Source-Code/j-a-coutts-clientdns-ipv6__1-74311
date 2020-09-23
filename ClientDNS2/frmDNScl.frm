VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDnsCl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client DNS"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   Icon            =   "frmDNScl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIPv6 
      Caption         =   "Use IPv6"
      Height          =   240
      Left            =   7200
      TabIndex        =   19
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3600
      Top             =   3120
   End
   Begin VB.CommandButton cmdSendQuery 
      Caption         =   "Send &Query"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdServer 
      Caption         =   "&Add DNS server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdRoot 
      Caption         =   "&Root Servers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame fraResponse 
      Caption         =   "DNS &Response"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   8055
      Begin MSFlexGridLib.MSFlexGrid grdResponse 
         Height          =   3680
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6482
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         MergeCells      =   2
         FormatString    =   "<|<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   4215
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7435
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Header"
               Key             =   "Header"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Questio&n"
               Key             =   "Question"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Answer"
               Key             =   "Answer"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Authorit&y"
               Key             =   "Authority"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Additiona&l"
               Key             =   "Additional"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraQuery 
      Caption         =   "DNS Q&uery"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.ListBox cboDnsServers 
         Height          =   1230
         Left            =   3720
         TabIndex        =   11
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtDomainName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "txtDomainName"
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox cboType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cboClass 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cboOPCode 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkRecursion 
         Caption         =   "R&ecursion Desired"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1410
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Domain name or IP address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Question &type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Question &class:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DNS &Server:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&OPCode:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   645
      End
   End
   Begin VB.Label lblHelp 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   8055
   End
End
Attribute VB_Name = "frmDnsCl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Server As cSocket2
Attribute Server.VB_VarHelpID = -1

Private MyDnsServer As String
Private WithEvents m_objDnsMsg As clsDNSMsg
Attribute m_objDnsMsg.VB_VarHelpID = -1
Private CurrentHost As String
Private Sub cboDnsServers_Click()
    CurrentHost = cboDnsServers.List(cboDnsServers.ListIndex)
    Server.RemoteHost = CurrentHost
End Sub

Private Sub chkIPv6_Click()
    If chkIPv6 Then
        Server.IPv6Flg = 6 'Set IP version
        cboType.ListIndex = 16
    Else
        Server.IPv6Flg = 4
        cboType.ListIndex = 0
    End If
    Server.CloseSck
    MyDnsServer = GetDnsServerAddress
    cboDnsServers.Clear
    cboDnsServers.AddItem MyDnsServer
    CurrentHost = MyDnsServer
    With Server 'Initialize Winsock for UDP
        .RemoteHost = MyDnsServer
        .RemotePort = 53
        .Bind2 .LocalPort
    End With
End Sub

Private Sub cmdRoot_Click()
    chkRecursion.Value = 0
    cboDnsServers.Clear
    cboDnsServers.AddItem MyDnsServer
    cboDnsServers.AddItem "a.root-servers.net"
    cboDnsServers.AddItem "b.root-servers.net"
    cboDnsServers.AddItem "c.root-servers.net"
    cboDnsServers.AddItem "d.root-servers.net"
    cboDnsServers.AddItem "e.root-servers.net"
    cboDnsServers.AddItem "f.root-servers.net"
    cboDnsServers.AddItem "g.root-servers.net"
    cboDnsServers.AddItem "h.root-servers.net"
    cboDnsServers.AddItem "i.root-servers.net"
    cboDnsServers.AddItem "j.root-servers.net"
    cboDnsServers.AddItem "k.root-servers.net"
    cboDnsServers.AddItem "l.root-servers.net"
    cboDnsServers.AddItem "m.root-servers.net"
End Sub


Private Sub cmdSendQuery_Click()
    Dim strName As String
    Dim strSend As String
    Dim DNSid As Long
    Dim QType As Integer
    Dim ErrCount As Integer
    strName = Trim$(txtDomainName.Text)
    QType = cboType.ItemData(cboType.ListIndex)
    If cboType = "PTR" Then
        strName = ReverseIP(strName) + ".in-addr.arpa"
        If Val(strName) = 0 Then
            MsgBox txtDomainName.Text + " Not a valid IP!"
            Exit Sub
        End If
    End If
    strSend = m_objDnsMsg.FormatQuestion(strName, QType, chkRecursion)
    On Error GoTo SendErr 'Added for Win 9x bug on first send
    If Left$(strSend, 5) <> "Error" Then
        Server.RemoteHost = CurrentHost 'Refresh Host
        Server.SendData strSend
        DNSid = Asc(Mid$(strSend, 1, 1)) * 256 + Asc(Mid$(strSend, 2, 1))
    End If
    fraResponse.Caption = "DNS &Response"
    Timer1.Enabled = True
    Exit Sub
SendErr:
    ErrCount = ErrCount + 1
    If ErrCount < 5 Then
        Resume
    Else
        MsgBox "Error" + Str$(err) + " Encountered Sending Data!"
    End If
End Sub

Private Sub cmdServer_Click()
    Dim strServer As String
    strServer = InputBox("Please, type in the address (or domain name) of the DNS server you like to send DNS queries to.", "DNS server address")
    If Len(strServer) > 0 Then
        cboDnsServers.AddItem strServer
    End If
End Sub

Private Sub Form_Activate()
    If GetSettings("IPv6") = "1" Then
        chkIPv6.Value = 1
    End If
End Sub

Private Sub Form_Click()
    AboutBox.Show 1
End Sub

Private Sub Form_Load()
    'Create an instance of the CSocketMaster class
    Set Server = New cSocket2
    Server.Protocol = sckUDPProtocol
    If chkIPv6 Then
        Server.IPv6Flg = 6 'Set IP version
    Else
        Server.IPv6Flg = 4
    End If
    txtDomainName.Text = ""
    With grdResponse
        .ColWidth(0) = 2480
        .ColWidth(1) = 5000
        .ColAlignment(1) = flexAlignLeftCenter
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With
    MyDnsServer = GetDnsServerAddress
    cboDnsServers.AddItem MyDnsServer
    CurrentHost = MyDnsServer
    With Server 'Initialize Winsock for UDP
        .RemoteHost = MyDnsServer
        .RemotePort = 53
        .Bind2 .LocalPort
    End With
    With cboType
        .AddItem "A"
        .ItemData(0) = 1 'DNS_TYPE_A
        .AddItem "NS"
        .ItemData(1) = 2 'DNS_TYPE_NS
        .AddItem "MD"
        .ItemData(2) = 3 'DNS_TYPE_MD
        .AddItem "MF"
        .ItemData(3) = 4 'DNS_TYPE_MF
        .AddItem "CNAME"
        .ItemData(4) = 5 'DNS_TYPE_CNAME
        .AddItem "SOA"
        .ItemData(5) = 6 'DNS_TYPE_SOA
        .AddItem "MB"
        .ItemData(6) = 7 'DNS_TYPE_MB
        .AddItem "MG"
        .ItemData(7) = 8 'DNS_TYPE_MG
        .AddItem "MR"
        .ItemData(8) = 9 'DNS_TYPE_MR
        .AddItem "NULL"
        .ItemData(9) = 10 'DNS_TYPE_NULL
        .AddItem "WKS"
        .ItemData(10) = 11 'DNS_TYPE_WKS
        .AddItem "PTR"
        .ItemData(11) = 12 'DNS_TYPE_PTR
        .AddItem "HINFO"
        .ItemData(12) = 13 'DNS_TYPE_HINFO
        .AddItem "MINFO"
        .ItemData(13) = 14 'DNS_TYPE_MINFO
        .AddItem "MX"
        .ItemData(14) = 15  'DNS_TYPE_MX
        .AddItem "TXT"
        .ItemData(15) = 16 'DNS_TYPE_TXT
        .AddItem "AAAA"
        .ItemData(16) = 28 'DNS_TYPE_AAAA
        .AddItem "SPF"
        .ItemData(17) = 99 'DNS_TYPE_SPF
        .AddItem "*"
        .ItemData(18) = 255 'DNS_TYPE_ALL
    End With
    With cboClass
        .AddItem "IN"
        .ItemData(0) = DNS_CLASS_IN
        .AddItem "CS"
        .ItemData(1) = DNS_CLASS_CS
        .AddItem "CH"
        .ItemData(2) = DNS_CLASS_CH
        .AddItem "HS"
        .ItemData(3) = DNS_CLASS_HS
    End With
    With cboOPCode
        .AddItem "QUERY"
        .AddItem "IQUERY"
        .AddItem "STATUS"
    End With
    cboClass.ListIndex = 0
    cboType.ListIndex = 0
    cboOPCode.ListIndex = 0
    cboDnsServers.ListIndex = 0
    'Create an instance of the clsDNSMsg class
    Set m_objDnsMsg = New clsDNSMsg
    'Initialize array
    ReDim RRArray(6, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSettings("IPv6", CStr(chkIPv6.Value))
End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
    Dim strdata As String
    Dim i       As Integer
    'Turn off the timer
    Timer1.Enabled = False
    Server.GetData strdata
    fraResponse.Caption = "DNS &Response received - " & Len(strdata) & " bytes from " & CurrentHost
    'Parse the message and initialize the object properties
    m_objDnsMsg.ParseData (strdata)
End Sub


Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Str$(Number) & ": " & Description
End Sub


Private Sub grdResponse_Click()
    lblHelp.Caption = "Double-Click cell to copy to Clipboard!"
End Sub

Private Sub grdResponse_DblClick()
    Clipboard.Clear
    DoEvents
    Clipboard.SetText grdResponse.Text
    lblHelp.Caption = grdResponse.Text
End Sub

Private Sub m_objDnsMsg_ParseDone()
    Dim i%
    'Get all the DNS servers from the DNS response message and
    'put them into the DNS servers listbox.
    cboDnsServers.Clear
    cboDnsServers.AddItem MyDnsServer
    For i = 0 To UBound(RRArray, 2)
        If RRArray(2, i%) = 2 Then
            cboDnsServers.AddItem RRArray(6, i%)
        End If
    Next i%
    Call TabStrip1_Click
End Sub

Private Sub TabStrip1_Click()
    Dim i As Integer
    Dim strMnemonicValue As String
    Dim strDescrition As String
''    On Error Resume Next
'    If m_blnResponseReceived Then
        For i = 0 To grdResponse.Rows - 2
            grdResponse.RemoveItem grdResponse.Rows - 1
        Next i
        Select Case TabStrip1.SelectedItem.key
            Case "Header"
                With m_objHeader
                    Call MergeCells(0, "DNS Message Header Fields")
                    grdResponse.AddItem "MessageID" & vbTab & .MsgID
                    grdResponse.AddItem "IsQuery" & vbTab & .QR
                    grdResponse.AddItem "IsResponse" & vbTab & (Not .QR)
                    GetOPCodeInfo .OpCode, strMnemonicValue, strDescrition
                    grdResponse.AddItem "OPCode" & vbTab & strMnemonicValue & " - " & strDescrition
                    grdResponse.AddItem "Authoritative Answer" & vbTab & .AA
                    grdResponse.AddItem "IsTruncated" & vbTab & .TC
                    grdResponse.AddItem "Recursion Desired" & vbTab & .RD
                    grdResponse.AddItem "Recursion Available" & vbTab & .RA
                    grdResponse.AddItem "Reserved" & vbTab & .Z
                    grdResponse.AddItem "Response Code" & vbTab & CStr(.RCode) & _
                          " - " & GetResponseCodeInfo(.RCode)
                    grdResponse.AddItem "QDCount" & vbTab & CStr(.QDCount)
                    grdResponse.AddItem "ANCount" & vbTab & .ANCount
                    grdResponse.AddItem "NSCount" & vbTab & .NSCount
                    grdResponse.AddItem "ARCount" & vbTab & .ARCount
                End With
            Case "Question"
                With m_objQuestion
                    Call MergeCells(0, "DNS Message Question Section")
                    grdResponse.AddItem "Name" & vbTab & .QName
                    GetTypeInfo .QType, strMnemonicValue, strDescrition
                    grdResponse.AddItem "Type" & vbTab & strMnemonicValue & " - " & strDescrition
                    GetClassInfo .QClass, strMnemonicValue, strDescrition
                    grdResponse.AddItem "Class" & vbTab & strMnemonicValue & " - " & strDescrition
                End With
            Case "Answer"
                Call MergeCells(0, "DNS Message Answer Section")
                Call DisplayResRecords(1)
            Case "Authority"
                Call MergeCells(0, "DNS Message Authority Section")
                Call DisplayResRecords(2)
            Case "Additional"
                Call MergeCells(0, "DNS Message Additional Section")
                Call DisplayResRecords(3)
       End Select
'    End If
    If grdResponse.Rows > 15 Then
        grdResponse.ColWidth(1) = 4820
    Else
        grdResponse.ColWidth(1) = 5020
    End If
End Sub


Private Sub DisplayResRecords(Section As Integer)
    Dim i%, M%, N%
    Dim strTypeMnemonic As String
    Dim strTypeDescription As String
    Dim strSection As String
    Dim strTemp As String
    Select Case Section
        Case 1
            strSection = "Answer"
        Case 2
            strSection = "Authority"
        Case 3
            strSection = "Additional"
    End Select
    For i% = 0 To UBound(RRArray, 2)
        If RRArray(5, i%) = 0 Then Exit Sub
        If Section = RRArray(0, i%) Then
            grdResponse.AddItem strSection & " " & i% + 1
            grdResponse.Row = grdResponse.Rows - 1
            grdResponse.Col = 0
            grdResponse.CellFontBold = True
            grdResponse.CellBackColor = vbInactiveCaptionText
            grdResponse.Col = 1
            grdResponse.CellFontBold = True
            grdResponse.CellBackColor = vbInactiveCaptionText
            grdResponse.MergeRow(grdResponse.Rows - 1) = True
            grdResponse.AddItem "Name" & vbTab & RRArray(1, i%)
            GetTypeInfo RRArray(2, i%), strTypeMnemonic, strTypeDescription
            grdResponse.AddItem "Type" & vbTab & strTypeMnemonic & " - " & strTypeDescription
            GetClassInfo RRArray(3, i%), strTypeMnemonic, strTypeDescription
            grdResponse.AddItem "Class" & vbTab & strTypeMnemonic & " - " & strTypeDescription
            grdResponse.AddItem "TTL" & vbTab & RRArray(4, i%) & " seconds"
            grdResponse.AddItem "Data Length" & vbTab & RRArray(5, i%) & " octets"
            Select Case RRArray(2, i%)
                Case DNS_TYPE_MX
                    N% = InStr(RRArray(6, i%), " ")
                    grdResponse.AddItem "Exchange" & vbTab & Mid$(RRArray(6, i%), N% + 1)
                    grdResponse.AddItem "Preference" & vbTab & Left$(RRArray(6, i%), N%)
                Case DNS_TYPE_SOA
                    N% = InStr(RRArray(6, i%), Chr$(0))
                    grdResponse.AddItem "MName" & vbTab & Left$(RRArray(6, i%), N%) 'MName
                    M% = InStr(N% + 1, RRArray(6, i%), Chr$(0))
                    grdResponse.AddItem "RName" & vbTab & Mid$(RRArray(6, i%), N% + 1, M% - N%) 'RName
                    strTemp = Mid$(RRArray(6, i%), M% + 1)
                    grdResponse.AddItem "Serial" & vbTab & CStr(Asc(Mid$(strTemp, 1, 1)) * 256 ^ 3 + Asc(Mid$(strTemp, 2, 1)) * 256 ^ 2 + Asc(Mid$(strTemp, 3, 1)) * 256 ^ 1 + Asc(Mid$(strTemp, 4, 1)))
                    grdResponse.AddItem "Refresh" & vbTab & CStr(Asc(Mid$(strTemp, 5, 1)) * 256 ^ 3 + Asc(Mid$(strTemp, 6, 1)) * 256 ^ 2 + Asc(Mid$(strTemp, 7, 1)) * 256 ^ 1 + Asc(Mid$(strTemp, 8, 1))) & " seconds"
                    grdResponse.AddItem "Retry" & vbTab & CStr(Asc(Mid$(strTemp, 9, 1)) * 256 ^ 3 + Asc(Mid$(strTemp, 10, 1)) * 256 ^ 2 + Asc(Mid$(strTemp, 11, 1)) * 256 ^ 1 + Asc(Mid$(strTemp, 12, 1))) & " seconds"
                    grdResponse.AddItem "Expire" & vbTab & CStr(Asc(Mid$(strTemp, 13, 1)) * 256 ^ 3 + Asc(Mid$(strTemp, 14, 1)) * 256 ^ 2 + Asc(Mid$(strTemp, 15, 1)) * 256 ^ 1 + Asc(Mid$(strTemp, 16, 1))) & " seconds"
                    grdResponse.AddItem "Minimum" & vbTab & CStr(Asc(Mid$(strTemp, 17, 1)) * 256 ^ 3 + Asc(Mid$(strTemp, 18, 1)) * 256 ^ 2 + Asc(Mid$(strTemp, 19, 1)) * 256 ^ 1 + Asc(Mid$(strTemp, 20, 1))) & " seconds"
                Case Else
                    grdResponse.AddItem "Data" & vbTab & RRArray(6, i%)
            End Select
        End If
    Next i%
End Sub

Private Sub MergeCells(ByVal intRow As Integer, strText As String)
    With grdResponse
        .Row = intRow
        .Col = 0
        .Text = Space(50) & strText
        .CellFontBold = True
        .CellBackColor = vbInactiveCaptionText
        .Col = 1
        .Text = Space(50) & strText
        .CellFontBold = True
        .CellBackColor = vbInactiveCaptionText
        .MergeRow(intRow) = True
    End With
End Sub
Private Function GetSettings(sKey As String) As String
    GetSettings = GetSetting(gAppName, "Settings", sKey, "")
End Function
Private Sub SaveSettings(sKey As String, sValue As String)
    SaveSetting gAppName, "Settings", sKey, sValue
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    MsgBox "Timout occurred!", vbExclamation, "Timeout"
End Sub

Private Sub txtDomainName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSendQuery_Click
End Sub





