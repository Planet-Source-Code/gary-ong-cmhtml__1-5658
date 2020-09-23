Attribute VB_Name = "modMain"
Option Explicit


Private Function SendMHTML(ByVal vsSubject As String, ByVal vsText As String, ByVal vsEmailFile As String) As Boolean
    On Error GoTo ErrSendMHTML
    Dim sEmail              As String
    Dim sEmails()           As String
    Dim sData               As String
    Dim lResult             As Long
    Dim i                   As Integer
    Dim objMHTML            As CMHTML
    Dim sAttachments        As String
    Dim saAttachments()     As String
    Dim iPos                As Integer
    Dim sServer             As String
    Dim sFromName           As String
    Dim sFromEmail          As String
    
    sData = String(32000, " ")

    ' Get the recipients from the email.txt file in the app.path
    lResult = GetPrivateProfileSection("Email Addresses", sData, Len(sData), vsEmailFile)
    sEmail = Left(sData, lResult)
    sEmails = Split(sEmail, Chr(0))
    sEmail = Join(sEmails, ";")
    
    ' Get the list of attachments from the email.txt file
    lResult = GetPrivateProfileSection("Attachments", sData, Len(sData), vsEmailFile)
    sAttachments = Left(sData, lResult)
    saAttachments = Split(sAttachments, Chr(0))
    
    ' Instantiate the MHTML object that sends the email
    Set objMHTML = New CMHTML
        
    ' Add the attachments as specified in the email.txt file -
    ' all attachments must be located in the app.path
    For i = 0 To UBound(saAttachments) - 1
        iPos = InStr(saAttachments(i), "=")
        objMHTML.AddAttachment App.Path & "\" & Right(saAttachments(i), Len(saAttachments(i)) - iPos), Left(saAttachments(i), iPos - 1)
    Next
    
    ' Grab the email server name from email.txt
    lResult = GetPrivateProfileString("General", "Server", "mail.labyrinth.net.au", sData, Len(sData), vsEmailFile)
    sServer = Left$(sData, lResult)
    
    ' Grab the from name
    lResult = GetPrivateProfileString("General", "FromName", "X1", sData, Len(sData), vsEmailFile)
    sFromName = Left$(sData, lResult)
    
    ' Grab the from email address
    lResult = GetPrivateProfileString("General", "FromEmail", "X1@labyrinth.net.au", sData, Len(sData), vsEmailFile)
    sFromEmail = Left$(sData, lResult)
    
    objMHTML.OmitHeader = False
    objMHTML.SaveSession = True
 
    ' Send it off
    objMHTML.SendEmail sServer, sFromName, sFromEmail, _
        sEmail, vsSubject, vsText
    
    ' Clean up
    Set objMHTML = Nothing

    SendMHTML = True
    
    Exit Function
    
ErrSendMHTML:

End Function
