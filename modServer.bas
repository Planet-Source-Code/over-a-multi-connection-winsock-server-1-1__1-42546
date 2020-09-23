Attribute VB_Name = "modServer"
Option Explicit

'Basic Server Settings

Public Const ServerPort = 4500                                  'Which port should our server listen to ?
Public Const MaxCon = 20                                         'Maximum Number of Connections. Increase this if necessary.

Public bSocketStatus(MaxCon) As Boolean                         'Stores which Sockets are already used



' Creates several Accept sockets
' Sockets 1 to MaxCon are used to accept connections.
' Socket 0 is used to tell a client that the server is already full.
Public Function InitAcceptSockets() As Boolean
    
    On Error GoTo err
    
    Dim i As Integer
    
    For i = 1 To MaxCon
    
        ' creates a copy of frmMain.sckAccept(0) during runtime
        Load frmMain.sckAccept(i)
        
    Next i
    
    ' Everything went fine
    InitAcceptSockets = True
    Exit Function
    
err:
    
    InitAcceptSockets = False
    
End Function



' Returns the Number of a unused Socket.
' if no free sockets are left then 0 is returned
Public Function GetFreeSocket() As Integer
    
    Dim i As Integer
    
    For i = 1 To MaxCon
        
        If bSocketStatus(i) = False Then
            ' socket i is unused.
            GetFreeSocket = i
            Exit Function
        End If
        
    Next i
    
    ' No free sockets left!
    GetFreeSocket = 0
    
End Function



' Add a line of text to the debug textbox and scroll down.
Public Sub DebugText(sText As String)
    
    With frmMain
        .txtDebug.text = .txtDebug.text & sText & vbNewLine
        .txtDebug.SelStart = Len(.txtDebug) - 2
    End With
    
End Sub



' Insert all connected clients into the combobox
Public Sub RefreshComboBox()
    
    ' Clear the combobox
    frmMain.cmbClients.Clear
    
    Dim i As Integer
    
    For i = 1 To MaxCon
    
        ' A client is connected on socket i
        If bSocketStatus(i) = True Then frmMain.cmbClients.AddItem "Client " & i & " (" & frmMain.sckAccept(i).RemoteHostIP & ")"
    
    Next i
    
End Sub
