VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CFTPLink Test"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Check File Types"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Type your password in here"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Enter username to login to server"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Enter FTP server name here"
      Top             =   840
      Width           =   3375
   End
   Begin VB.ListBox List2 
      DragMode        =   1  'Automatic
      Height          =   645
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      ToolTipText     =   "Drag and drop files here"
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Clear"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   4320
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Debug Information"
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Upload Files"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Files to send"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "FTP Server"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
DefInt A-Z

Private WithEvents ftp As CFTPLink
Attribute ftp.VB_VarHelpID = -1

Private Sub Command1_Click()
    Dim i As Integer
    
    With ftp
        .Server = Text1(0).Text
        .Username = Text1(1).Text
        .Password = Text1(2).Text
        
        For i = 0 To List2.ListCount - 1
            .AddFileToSend List2.List(i), GetFilename(List2.List(i))
        Next
        .SendFiles
    End With
    
End Sub

Private Sub Command2_Click()
    List1.Clear
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    
    For i = 0 To List2.ListCount - 1
        If ftp.IsBinaryFile(List2.List(i)) Then
            MsgBox List2.List(i) & " is a binary file"
        Else
            MsgBox List2.List(i) & " is a text file"
        End If
    Next
    
End Sub

Private Sub Form_Load()
    Set ftp = New CFTPLink
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ftp = Nothing
End Sub

Private Sub ftp_StatusUpdate(vsText As String, vlEventType As StatusEventType)
    Debug.Print "StatusUpdate [" & vlEventType & "] : " & vsText
    List1.AddItem "StatusUpdate [" & vlEventType & "] : " & vsText
End Sub

Private Sub List2_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 1 To data.Files.Count
        If (GetAttr(data.Files.Item(i)) And vbDirectory) = 0 Then
            List2.AddItem data.Files.Item(i)
        End If
    Next
End Sub

Private Function GetFilename(ByVal vsFullPath As String) As String
    Dim v As Variant
    
    If InStr(vsFullPath, "/") <> 0 Then
        v = Split(vsFullPath, "/")
        GetFilename = v(UBound(v))
    
    ElseIf InStr(vsFullPath, "\") <> 0 Then
        v = Split(vsFullPath, "\")
        GetFilename = v(UBound(v))
    
    Else
        GetFilename = Replace(vsFullPath, ":", "_")
        
    End If
    
End Function
