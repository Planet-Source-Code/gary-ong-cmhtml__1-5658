VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CFTPLink Test"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   6840
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Drag & drop files to attach using MIME"
      Top             =   6000
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   1845
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "MHTMLTest.frx":0000
      ToolTipText     =   "Enter the HTML email body"
      Top             =   3600
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1320
      TabIndex        =   5
      Text            =   "The Planet Rocks"
      ToolTipText     =   "Enter the email subject"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   645
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      ToolTipText     =   "Enter recipients email address(es) - use commas or semicolons to delimit"
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Enter your email address"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Enter your name"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Enter Mail server name here"
      Top             =   120
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      ToolTipText     =   "Drag & drop files to attach using MIME"
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send Email"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Debug Info"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "HTML Body"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Subject"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Recipients"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Files to attach"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From email"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "From name"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Mail Server"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Author: Gary Ong
' Notes : There's one trick that can catch you out in this code. The cid
'         tag in your html is case sensitive depending on your email client
'         To be safe - always use lowercase both in your email and when you
'         call AddAttachment and specify the cid. Be warned!
'

Option Explicit
Option Compare Text
DefInt A-Z

Private WithEvents objMHTML As CMHTML
Attribute objMHTML.VB_VarHelpID = -1

Private Sub Command1_Click()
    Dim i As Integer
    
    With objMHTML
        
        For i = 0 To List1.ListCount - 1
            ' By default I make the cid the filename without the path
            ' This means you can't have 2 attachments with the same filename
            ' for this code example
            .AddAttachment List1.List(i), LCase$(GetFilename(List1.List(i)))
        Next
        
        ' This is the important call - note that the ToEmailAddress parameter
        ' is a comma or semicolon delimited string.  The body text must be
        ' html otherwise ... I've never actually tried ...
        
        .SendEmail Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text, Text1(5).Text
    
    End With
    
    Beep
    
End Sub

Private Sub Form_Load()
    Set objMHTML = New CMHTML
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objMHTML = Nothing
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

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 1 To Data.Files.Count
        If (GetAttr(Data.Files.Item(i)) And vbDirectory) = 0 Then
            List1.AddItem Data.Files.Item(i)
        End If
    Next

End Sub

Private Sub objMHTML_StatusUpdate(vsText As String, vlEventType As StatusEventType)
    List2.AddItem vsText
End Sub

Private Sub Command2_Click()
    List2.Clear
End Sub

