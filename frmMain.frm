VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Multiple Connection Winsock Server"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdSendAll 
      Caption         =   "Send &all"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdSendData 
      Caption         =   "&Send"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtSendData 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Text            =   "Send this text to a client ..."
      Top             =   4440
      Width           =   4335
   End
   Begin VB.ComboBox cmbClients 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock sckAccept 
      Index           =   0
      Left            =   6600
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   6120
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4500
   End
   Begin VB.TextBox txtDebug 
      Height          =   4215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' a Multiconnection Winsock TCP Server by over
'
' If you have got any questions about this example contact me:
' overkillpage@gmx.net
'
' You can change the basic server settings like port and maximum connections
' in the modServer.bas module.
'
' Greetings fly out to AbsoluteB
'

Option Explicit




Private Sub cmdSendAll_Click()

    Dim i As Integer
    
    'loop through all sockets
    For i = 1 To MaxCon
        
        'Check if the "current" socket is conencted, if so ...
        If sckAccept(i).State = 7 Then
        
            '... send the message
            sckAccept(i).SendData txtSendData.Text
            DoEvents
        
        End If
        
    Next i

End Sub

' starting our server
Private Sub Form_Load()
    
    ' Set sckListen to listen mode. It will accept all incoming connection requests.
    sckListen.LocalPort = ServerPort
    sckListen.Listen

    ' create some accept sockets
    If Not InitAcceptSockets Then
        MsgBox "ERROR Can't create accept sockets!", vbCritical, "Error"
        End
    End If
    
    ' debug some server information
    DebugText "Server is listening on port " & ServerPort & " for incoming connections ..."

End Sub



' a new client connected
Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)

    Dim aFreeSocket As Integer
    
    ' Request the number of an unused socket
    aFreeSocket = GetFreeSocket
    
    If aFreeSocket = 0 Then
        
        ' Tell the new client that the server is full and close the connection
        DebugText "Number of maximum clients reached ! A connection had to be refused!"
        sckAccept(0).Accept requestID
        DoEvents
        sckAccept(0).SendData "Sorry, server is full!"
        DoEvents
        sckAccept(0).Close
        
    Else
        
        ' accept the connection on a free socket. set status of this socket to true(used)
        bSocketStatus(aFreeSocket) = True
        sckAccept(aFreeSocket).Accept requestID
        DoEvents
        DebugText "A new Client with ID " & aFreeSocket & " and IP " & sckAccept(aFreeSocket).RemoteHostIP & " connected!"
        ' Send a welcome message to the new client
        sckAccept(aFreeSocket).SendData "Connection Accepted. Have a lot of fun."
        ' Refresh the combobox -> add our new client
        RefreshComboBox
        
    End If
    
End Sub



' One of the connected clients sent some data ...
' Add login function and additional commands here ...
Private Sub sckAccept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim sData As String
    sckAccept(Index).GetData sData
    
    ' output the incoming data to the debug textbox
    DebugText "Client " & Index & ": " & sData
    
End Sub



' a client disconnected
Private Sub sckAccept_Close(Index As Integer)
    
    ' Free the used socket.
    bSocketStatus(Index) = False
    DebugText "Client " & Index & " (" & sckAccept(Index).RemoteHostIP & ") disconnected"
    sckAccept(Index).Close
    
End Sub



' Sends data to a selected client
Private Sub cmdSendData_Click()

    If cmbClients.Text = "" Then
        MsgBox "Select a client first!", vbInformation, "No client selected"
        Exit Sub
    End If
    
    ' Get the client id from the Combobox
    Dim iClientId As Integer
    iClientId = Split(cmbClients.Text, " ")(1)
    
    ' Send data to client
    sckAccept(iClientId).SendData txtSendData.Text
    
    ' clear textbox
    txtSendData.Text = ""
    
End Sub

