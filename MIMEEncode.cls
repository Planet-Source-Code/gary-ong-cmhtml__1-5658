VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMHTML"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Author: Gary Ong
' Acknowledgements: Sebastian, Luis Cantero for EncodeBase64
'                   Brian Anderson for the SMTP Winsock code
'

Option Explicit
Option Compare Text

Public Event StatusUpdate(vsText As String, vlEventType As StatusEventType)

Public Enum StatusEventType
    glSETInfo = 0
    glSETWarning = 1
    glSETError = 2
End Enum

Private Const BOUNDARY_ID As String = "NextMimePart"

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Response            As String       ' Used to store server response code
Private frmWS               As frmWinsock
Private WithEvents Winsock1 As Winsock
Attribute Winsock1.VB_VarHelpID = -1
Private mdicAttachments     As Dictionary   ' store attachment filename and CID
Private mnTimeout           As Integer      ' timeout period in seconds
Private mbOmitHeader        As Boolean      ' debug property if true will omit all mail headers
Private mbSaveSession       As Boolean      ' debug property if true will write session details to app.path \ session.txt
Private miSessionFile       As Integer      ' file handler for saving session text

Private Sub Class_Initialize()

    Set frmWS = New frmWinsock
    Set Winsock1 = frmWS.Winsock1
    Set mdicAttachments = New Dictionary
    mnTimeout = 25 'default timeout (sec)
    
End Sub

Private Sub Class_Terminate()
    
    Set Winsock1 = Nothing
    Unload frmWS
    Set frmWS = Nothing
    Set mdicAttachments = Nothing

End Sub

Public Property Let ResponseTimeout(ByVal vnTimeoutInSeconds As Integer)
    mnTimeout = vnTimeoutInSeconds
End Property

Public Property Get ResponseTimeout() As Integer
    ResponseTimeout = mnTimeout
End Property

Public Property Let OmitHeader(ByVal vbMode As Boolean)
    mbOmitHeader = vbMode
End Property

Public Property Let SaveSession(ByVal vbMode As Boolean)
    mbSaveSession = vbMode
End Property

Public Function AddAttachment(ByVal vsFilename As String, Optional ByVal vsContentID As String) As Boolean
    On Error GoTo ErrAddAttachment
    
    AddAttachment = False
    
    If Dir$(vsFilename) <> vbNullString Then
        
        ' Add attachments to dictionary for encoding later
        If vsContentID = vbNullString Then
            mdicAttachments.Add vsFilename, vsFilename 'use the full pathname as the key
        Else
            mdicAttachments.Add vsContentID, vsFilename 'use the provided ContentID
        End If
        AddAttachment = True
    
    End If
    Exit Function
    
ErrAddAttachment:

End Function

Public Function SendEmail(ByVal vsMailServerName As String, ByVal vsFromName As String, _
        ByVal vsFromEmailAddress As String, ByVal vsToEmailAddress As String, _
        ByVal vsEmailSubject As String, ByVal vsEmailBodyOfMessage As String) As Boolean
    
    Dim sDateNow    As String
    Dim sFrom       As String
    Dim sTo         As String
    Dim sDate       As String
    Dim sFromDetail As String
    Dim sToDetail   As String
    Dim sSubject    As String
    Dim sBody       As String
    Dim sBodyHeader As String
    Dim sMailerName As String
    Dim sMIMEHeader As String
    Dim sDomain     As String
    Dim iPos        As Integer
    Dim iRecipCount As Integer
    Dim vRecipients As Variant
    Dim i           As Integer
    
    SendEmail = False
    
    ' Sanity check parameters
    If vsMailServerName = vbNullString Or vsFromName = vbNullString Or _
        vsFromEmailAddress = vbNullString Or vsToEmailAddress = vbNullString Then
        RaiseEvent StatusUpdate("Missing parameter in call to SendEmail", glSETError)
        Exit Function
    End If
    
    ' Open the debug output file if specified
    If mbSaveSession Then
        miSessionFile = FreeFile
        Open App.Path & "\session.txt" For Output As #miSessionFile
    End If
    
    ' Convert comma/semicolon delimited string into an array
    iRecipCount = GetRecipients(vsToEmailAddress, vRecipients)
    
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
    If Winsock1.State = sckClosed Then ' Check to see if socet is closed
        
        If Not mbOmitHeader Then
        
            sDateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss")
            sFrom = "mail from:" & " <" + vsFromEmailAddress + ">" & vbCrLf  ' Get who's sending E-Mail address
            
            sToDetail = "To: undisclosed recipients" & vbCrLf   ' Hide who its going to
            sDate = "Date:" + Chr(32) + sDateNow + vbCrLf ' Date when being sent
            sFromDetail = "From: " & Chr(34) & vsFromName & Chr(34) & " <" + vsFromEmailAddress + ">" + vbCrLf  ' Who's Sending
            sSubject = "Subject:" + Chr(32) + vsEmailSubject + vbCrLf ' Subject of E-Mail
            sBody = vsEmailBodyOfMessage + vbCrLf ' E-mail message body
            sMailerName = "X-Mailer: GOMail" + vbCrLf ' What program sent the e-mail, customize this
            
            sMIMEHeader = GetMIMEHeader(BOUNDARY_ID)
            
            sBodyHeader = sFromDetail + sDate + sMailerName + sToDetail + sSubject & sMIMEHeader  ' Combine for proper SMTP sending
        
        Else
        
            ' This is good for debugging as you can eliminate possible problems
            ' in the email header. The email message content will be only the
            ' vsBodyOfMessage parameter.
            
            sFrom = "mail from:" & " <" + vsFromEmailAddress + ">" & vbCrLf  ' Get who's sending E-Mail address
            sBody = vsEmailBodyOfMessage + vbCrLf
        
        End If
        
        ' Derive domain name from senders email address
        sDomain = vsFromEmailAddress
        iPos = Len(sDomain) - InStr(sDomain, "@")
        sDomain = Right$(sDomain, iPos)
        
        Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
        Winsock1.RemoteHost = vsMailServerName ' Set the server address
        Winsock1.RemotePort = 25 ' Set the SMTP Port
        Winsock1.Connect ' Start connection
        
        WaitFor ("220")
        RaiseEvent StatusUpdate("Connecting...", glSETInfo)
        
        Call SendData("HELO " & sDomain & vbCrLf)
        WaitFor ("250")
        RaiseEvent StatusUpdate("Connected", glSETInfo)
    
        Call SendData(sFrom)
        RaiseEvent StatusUpdate("Sending message...", glSETInfo)
        WaitFor ("250")
    
        ' List the email recipients
        For i = 0 To iRecipCount - 1
            If Trim$(vRecipients(i)) <> "" Then
                Call SendData("RCPT TO: " & "<" & Trim$(vRecipients(i)) & ">" & vbCrLf)
                WaitFor ("250")
            End If
        Next
    
        ' This the email body
        Call SendData("data" + vbCrLf)
        WaitFor ("354")
        
        If Not mbOmitHeader Then
        
            Call SendData(sBodyHeader + vbCrLf)
            Call SendData(sBody)
            Call SendData(EncodeAttachments())
            
        Else
        
            Call SendData(sBody)
            
        End If
        
        ' This is how we end the SMTP session
        Call SendData(vbCrLf & "." & vbCrLf)
        WaitFor ("250")
    
        Call SendData("quit" + vbCrLf)
        RaiseEvent StatusUpdate("Disconnecting...", glSETInfo)
        WaitFor ("221")
    
        Winsock1.Close
        
        RaiseEvent StatusUpdate("Mail sent", glSETInfo)
        SendEmail = True
        
    Else
        RaiseEvent StatusUpdate("Internal error - winsock control in unexpected state", glSETError)
    End If
       
    ' Remove all attachments after send
    mdicAttachments.RemoveAll
    
    ' Close the debug output file
    If mbSaveSession Then
        Close #miSessionFile
    End If
    
End Function

Private Sub SendData(ByVal vsData As String)

    Winsock1.SendData vsData
    Debug.Print vsData
    If mbSaveSession Then
        Print #miSessionFile, vsData
    End If
    
End Sub

Private Function GetRecipients(ByVal vsList As String, ByRef vvList As Variant) As Integer

    ' Check for multiple recipients delimited by commas or semi-colons
    If InStr(vsList, ",") <> 0 Then
        vvList = Split(vsList, ",")
        GetRecipients = UBound(vvList) + 1
    Else
        If InStr(vsList, ";") <> 0 Then
            vvList = Split(vsList, ";")
            GetRecipients = UBound(vvList) + 1
        Else
            ' Only one recipient
            GetRecipients = 1
            vvList = Array(vsList)
        End If
    End If

End Function

Private Sub WaitFor(ResponseCode As String)
    Dim Start               As Single
    Dim Tmr                 As Single
    
    Start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = Start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > mnTimeout Then ' Time in seconds to wait
            RaiseEvent StatusUpdate("SMTP service error, timed out while waiting for response", glSETError)
            Exit Sub
        End If
        Sleep 1  ' to prevent hogging of the CPU
    Wend
    
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            RaiseEvent StatusUpdate("SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, glSETError)
            Exit Sub
        End If
        Sleep 1  ' to prevent hogging of the CPU
    Wend
    Response = "" ' Sent response code to blank **IMPORTANT**
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Winsock1.GetData Response ' Check for incoming response *IMPORTANT*

End Sub

Public Function GetMIMEHeader(ByVal vsBoundaryID As String) As String
    
    GetMIMEHeader = "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: multipart/related; boundary=" & _
        Chr(34) & vsBoundaryID & Chr(34) & "; type=" & Chr(34) & _
        "text/html" & Chr(34) & vbCrLf & _
        "Text displayed only to non-MIME-compliant mailers" & vbCrLf & _
        "--" & vsBoundaryID & vbCrLf & _
        "Content-Type: text/html; charset=us-ascii" & vbCrLf & _
        "Content-Transfer-Encoding: 7bit" & vbCrLf
        
End Function

Private Function EncodeAttachments() As String
    Dim sResult As String
    Dim n As Integer
    Dim k As Variant
    Dim i As Variant
    
    k = mdicAttachments.Keys
    i = mdicAttachments.Items
    
    If mdicAttachments.Count > 0 Then
        For n = 1 To mdicAttachments.Count
            sResult = sResult & EncodeFile(i(n - 1), k(n - 1), BOUNDARY_ID)
        Next
    End If
    
    sResult = sResult & vbCrLf & "--" & BOUNDARY_ID & "--" & vbCrLf
    EncodeAttachments = sResult

End Function

Public Function EncodeFile(ByVal vsFullPathname As String, ByVal vsCID As String, ByVal vsBoundaryID As String) As String
    Dim sResult As String
    Dim sFileName As String
    
    sFileName = GetFilename(vsFullPathname)
    
    'Preparing the Mime Header
    sResult = vbCrLf & "--" & vsBoundaryID & vbNewLine
    sResult = sResult & "Content-Type: application/octet-stream; name=" & Chr(34) & sFileName & Chr(34) & vbNewLine
    sResult = sResult & "Content-ID: <" & vsCID & ">" & vbNewLine
    sResult = sResult & "Content-Transfer-Encoding: base64" & vbNewLine
    sResult = sResult & "Content-Disposition: attachment; filename=" & Chr(34) & sFileName & Chr(34) & vbNewLine
    
    sResult = sResult & EncodeBase64(vsFullPathname)

    EncodeFile = sResult
    
End Function

Private Function GetFilename(ByVal vsFullPathname As String, Optional ByVal vbOmitExtension As Boolean = False) As String
    Dim iBackslashPos As Integer
    Dim iExtensionPos As Integer
    Dim i As Integer
    
    For i = Len(vsFullPathname) To 1 Step -1
        iBackslashPos = InStr(i, vsFullPathname, "\")
        If iBackslashPos > 0 Then Exit For
    Next
    
    If Not vbOmitExtension Then
        GetFilename = Mid(vsFullPathname, iBackslashPos + 1)
    Else
    
        For i = Len(vsFullPathname) To 1 Step -1
            iExtensionPos = InStr(i, vsFullPathname, ".")
            If iExtensionPos > 0 Then Exit For
        Next
        
        GetFilename = Mid(vsFullPathname, iBackslashPos + 1, iExtensionPos - iBackslashPos - 1)
    
    End If
    
End Function

Public Function EncodeBase64(ByVal vsFullPathname As String) As String
    'For Encoding BASE64
    Dim b           As Integer
    Dim Base64Tab   As Variant
    Dim bin(3)      As Byte
    Dim s           As String
    Dim l           As Long
    Dim i           As Long
    Dim FileIn      As Long
    Dim sResult     As String
    Dim n           As Long
    
    'Base64Tab=>tabla de tabulación
    Base64Tab = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")
    
    Erase bin
    l = 0: i = 0: FileIn = 0: b = 0:
    s = ""
    
    'Gets the next free filenumber
    FileIn = FreeFile
    
    'Open Base64 Input File
    Open vsFullPathname For Binary As FileIn
    
    sResult = s & vbCrLf
    s = ""
    
    l = LOF(FileIn) - (LOF(FileIn) Mod 3)
    
    For i = 1 To l Step 3

        'Read three bytes
        Get FileIn, , bin(0)
        Get FileIn, , bin(1)
        Get FileIn, , bin(2)
        
        'Always wait until there're more then 64 characters
        If Len(s) > 64 Then

            s = s & vbCrLf
            sResult = sResult & s
            s = ""

        End If

        'Calc Base64-encoded char
        b = (bin(n) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
        s = s & Base64Tab(b) 'the character s holds the encoded chars
        
        b = ((bin(n) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
        s = s & Base64Tab(b)
        
        b = ((bin(n + 1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
        s = s & Base64Tab(b)
        
        b = bin(n + 2) And &H3F
        s = s & Base64Tab(b)
        
    Next i

    'Now, you need to check if there is something left
    If Not (LOF(FileIn) Mod 3 = 0) Then

        'Reads the number of bytes left
        For i = 1 To (LOF(FileIn) Mod 3)
            Get FileIn, , bin(i - 1)
        Next i
    
        'If there are only 2 chars left
        If (LOF(FileIn) Mod 3) = 2 Then
            b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
            s = s & Base64Tab(b)
            
            b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
            s = s & Base64Tab(b)
            
            b = ((bin(1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
            s = s & Base64Tab(b)
            
            s = s & "="
        
        Else 'If there is only one char left
            b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
            s = s & Base64Tab(b)
            
            b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
            s = s & Base64Tab(b)
            
            s = s & "=="
        End If
    End If

    'Send the characters left
    If s <> "" Then
        s = s & vbCrLf
        sResult = sResult & s
    End If
    
    'Send the last part of the MIME Body
    s = ""
    
    Close FileIn
    EncodeBase64 = sResult
    
End Function
