VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDNSMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************
'clsDnsMsg class
'Copyright 2004 by JAC Computing
'********************************************************************************
'Version: 1.0.0     Last Modified: 08-MAR-2004
'********************************************************************************
Option Explicit

Public Event ParseDone()


Public Function ParseData(strdata As String) As Boolean
    Const strErr As String = "Error!"
    Dim i%, J%, N%, Section%
    Dim Result As Long
    Dim Question As String
    Dim RName As String
    Dim RDLength As Long
    Dim MXPriority As Long
    Dim Answer As String
    Dim tmpPoint As Integer
    Dim tmpDbl As Double
    With m_objHeader
        .MsgID = Asc(Mid$(strdata, 1, 1)) * 256 ^ 1 + Asc(Mid$(strdata, 2, 1)) 'Integer
        .QR = Asc(Mid$(strdata, 3, 1)) And &H80 'Boolean
        .OpCode = Asc(Mid$(strdata, 3, 1)) And &H78 'Byte
        .AA = Asc(Mid$(strdata, 3, 1)) And &H4 'Boolean
        .TC = Asc(Mid$(strdata, 3, 1)) And &H2 'Boolean
        .RD = Asc(Mid$(strdata, 3, 1)) And &H1 'Boolean
        .RA = Asc(Mid$(strdata, 4, 1)) And &H80 'Boolean
        .Z = Asc(Mid$(strdata, 4, 1)) And &H70 'Byte
        .RCode = Asc(Mid$(strdata, 4, 1)) And &HF 'Byte
        .QDCount = Asc(Mid$(strdata, 5, 1)) * 256 ^ 1 + Asc(Mid$(strdata, 6, 1)) 'Integer
        .ANCount = Asc(Mid$(strdata, 7, 1)) * 256 ^ 1 + Asc(Mid$(strdata, 8, 1)) 'Integer
        .NSCount = Asc(Mid$(strdata, 9, 1)) * 256 ^ 1 + Asc(Mid$(strdata, 10, 1)) 'Integer
        .ARCount = Asc(Mid$(strdata, 11, 1)) * 256 ^ 1 + Asc(Mid$(strdata, 12, 1)) 'Integer
    End With
    N% = 13
    Question = AddDotSep(strdata, N%)
    'N% should now point to QType
    If Question = strErr Then GoTo ParseDataErr
    With m_objQuestion
        .QName = Question 'String
        .QType = Asc(Mid$(strdata, N%, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 1, 1))  'Integer
        .QClass = Asc(Mid$(strdata, N% + 2, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 3, 1)) 'Integer
    End With
    N% = N% + 4 'Adjust to start of Resourse Record
    ClearArray
    For J% = 1 To 3
        Select Case J%
            Case 1
                Section% = m_objHeader.ANCount
            Case 2
                Section% = m_objHeader.NSCount
            Case 3
                Section% = m_objHeader.ARCount
        End Select
        For i% = 1 To Section%
            RName = AddDotSep(strdata, N%)
            Debug.Print RName
            RDLength = Asc(Mid$(strdata, N% + 8, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 9, 1))
            With m_objRRec
                .RName = RName 'String
                .RType = Asc(Mid$(strdata, N%, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 1, 1)) 'Integer
                .RClass = Asc(Mid$(strdata, N% + 2, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 3, 1)) 'Integer
                .TTL = Asc(Mid$(strdata, N% + 4, 1)) * 256 ^ 3 + Asc(Mid$(strdata, N% + 5, 1)) * 256 ^ 2 + Asc(Mid$(strdata, N% + 6, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 7, 1)) 'Long
                .RDLength = RDLength 'Integer
            End With
            If RName = strErr Then GoTo ParseDataErr
            N% = N% + 10 'Adjust pointer to point to RData
            Select Case m_objRRec.RType
                Case DNS_TYPE_A
                    Answer = CStr(Asc(Mid$(strdata, N%, 1))) + "." + CStr(Asc(Mid$(strdata, N% + 1, 1))) + "." + CStr(Asc(Mid$(strdata, N% + 2, 1))) + "." + CStr(Asc(Mid$(strdata, N% + 3, 1)))
                    Debug.Print Answer
                    m_objRRec.RData = Answer
                    N% = N% + 4
                    If Answer = strErr Then GoTo ParseDataErr
                Case DNS_TYPE_AAAA
                    Answer = Hex$(Asc(Mid$(strdata, N%))) + Right$("0" + Hex$(Asc(Mid$(strdata, N% + 1))), 2) _
                      + ":" + Hex$(Asc(Mid$(strdata, N% + 2))) + Right$("0" + Hex$(Asc(Mid$(strdata, N% + 3))), 2) _
                      + ":" + Hex$(Asc(Mid$(strdata, N% + 4))) + Right$("0" + Hex$(Asc(Mid$(strdata, N% + 5))), 2) _
                      + ":" + Hex$(Asc(Mid$(strdata, N% + 6))) + Right$("0" + Hex$(Asc(Mid$(strdata, N% + 7))), 2) _
                      + ":" + Hex$(Asc(Mid$(strdata, N% + 8))) + Right$("0" + Hex$(Asc(Mid$(strdata, N% + 9))), 2) _
                      + ":" + Hex$(Asc(Mid$(strdata, N% + 10))) + Right$("0" + Hex$(Asc(Mid$(strdata, N% + 11))), 2) _
                      + ":" + Hex$(Asc(Mid$(strdata, N% + 12))) + Right$("0" + Hex$(Asc(Mid$(strdata, N% + 13))), 2) _
                      + ":" + Hex$(Asc(Mid$(strdata, N% + 14))) + Right$("0" + Hex$(Asc(Mid$(strdata, N% + 15))), 2)
                    Debug.Print Answer
                    m_objRRec.RData = Answer
                    N% = N% + 16
                    If Answer = strErr Then GoTo ParseDataErr
                Case DNS_TYPE_CNAME, DNS_TYPE_NS, DNS_TYPE_MB, DNS_TYPE_MG, DNS_TYPE_MR, DNS_TYPE_PTR
                    Answer = AddDotSep(strdata, N%)
                    Debug.Print Answer
                    m_objRRec.RData = Answer
                    If Answer = strErr Then GoTo ParseDataErr
                Case DNS_TYPE_MX
                    MXPriority = Asc(Mid$(strdata, N%, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 1, 1))
                    N% = N% + 2
                    Answer = AddDotSep(strdata, N%)
                    Debug.Print CStr(MXPriority) + " " + Answer
                    m_objRRec.RData = CStr(MXPriority) + " " + Answer
                    If Answer = strErr Then GoTo ParseDataErr
                Case DNS_TYPE_SOA
                    'Recover Host Name
                    Answer = AddDotSep(strdata, N%)
                    Debug.Print Answer
                    m_objSOARec.MName = Answer 'String
                    If Answer = strErr Then GoTo ParseDataErr
                    'Recover Host Admin
                    Answer = AddDotSep(strdata, N%)
                    Debug.Print Answer
                    With m_objSOARec
                        .RName = Answer 'String
                        .Serial = Asc(Mid$(strdata, N%, 1)) * 256 ^ 3 + Asc(Mid$(strdata, N% + 1, 1)) * 256 ^ 2 + Asc(Mid$(strdata, N% + 2, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 3, 1)) 'Long
                        .Refresh = Asc(Mid$(strdata, N% + 4, 1)) * 256 ^ 3 + Asc(Mid$(strdata, N% + 5, 1)) * 256 ^ 2 + Asc(Mid$(strdata, N% + 6, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 7, 1)) 'Long
                        .Retry = Asc(Mid$(strdata, N% + 8, 1)) * 256 ^ 3 + Asc(Mid$(strdata, N% + 9, 1)) * 256 ^ 2 + Asc(Mid$(strdata, N% + 10, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 15, 1)) 'Long
                        .Minimum = Asc(Mid$(strdata, tmpPoint + 16, 1)) * 256 ^ 3 + Asc(Mid$(strdata, N% + 17, 1)) * 256 ^ 2 + Asc(Mid$(strdata, N% + 18, 1)) * 256 ^ 1 + Asc(Mid$(strdata, N% + 19, 1)) 'Long
                    End With
                    m_objRRec.RData = m_objSOARec.MName + Chr$(0) + m_objSOARec.RName + Chr$(0) + Mid$(strdata, N%, 20)
                    If Answer = strErr Then GoTo ParseDataErr
                    N% = N% + 20
                Case DNS_TYPE_TXT
                    Answer = Mid$(strdata, N% + 1, RDLength - 1)
                    Debug.Print Answer
                    m_objRRec.RData = Answer
                    If Answer = strErr Then GoTo ParseDataErr
                Case DNS_TYPE_SPF
                    Answer = Mid$(strdata, N% + 1, RDLength - 1)
                    Debug.Print Answer
                    m_objRRec.RData = Answer
                    If Answer = strErr Then GoTo ParseDataErr
            End Select
            Result = AddToArray(J%, m_objRRec)
        Next i%
    Next J%
    RaiseEvent ParseDone
    Exit Function
ParseDataErr:
    Result = AddToArray(J%, m_objRRec)
    RaiseEvent ParseDone
End Function

Public Function FormatQuestion(QName As String, QType As Integer, QRecurse As Byte) As String
    Dim strSend As String
    Dim strTemp As String
    Const sPair00 As String = vbNullChar + vbNullChar
    Const sPair01 As String = vbNullChar + "" ' is Chr$(1)
    Const sPair10 As String = "" + vbNullChar
    Const sPair11 As String = "" + ""
    Dim M%, M1%, N%
    Dim strArray(10) As String
    Dim strControl As String
    Static lByte As Byte
    Static hByte As Byte
    On Error GoTo FormatQuestionErr
    If lByte = 255 Then
        hByte = hByte + 1
        lByte = 0
    Else
        lByte = lByte + 1
    End If
    'Replace "."
    strTemp = "|" + Replace(QName, ".", "|") + "|"
    'Separate name into string values
    M% = 1
    M1% = InStr(2, strTemp, "|")
    While M1% > 0
        strArray(N%) = Mid$(strTemp, M% + 1, M1% - M% - 1)
        M% = M1%
        M1% = InStr(M% + 1, strTemp, "|")
        N% = N% + 1
    Wend
    'Reassmble with length
    N% = N% - 1
    M% = Len(strTemp)
    Mid$(strTemp, M%) = Chr$(0)
    M% = M% - Len(strArray(N%)) - 1
    While N% > 0
        Mid$(strTemp, M%) = Chr$(Len(strArray(N%)))
        N% = N% - 1
        M% = M% - Len(strArray(N%)) - 1
    Wend
    Mid$(strTemp, 1) = Chr$(Len(strArray(N%)))
    If QRecurse Then
        strControl = sPair10
    Else
        strControl = sPair00
    End If
    strSend = Chr$(hByte) + Chr$(lByte) + strControl _
      + sPair01 + sPair00 + sPair00 + sPair00 _
      + strTemp + Chr$(0) + Chr$(QType) + sPair01
    FormatQuestion = strSend
    Exit Function
FormatQuestionErr:
    FormatQuestion = "Error" + Error$(err)
End Function

