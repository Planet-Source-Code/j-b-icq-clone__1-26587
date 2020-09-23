Please note that every command sent from either program has a "/" in front of it.

From the server:

    '<<B>> Recieved Data Buffer <<B>>
    UserCommand = UserInfo(Index).Data & UserCommand
    If InStr(1, UserCommand, "`") Then
        Do While InStr(StrtPos + 1, UserCommand, "`") > 0
            StrtPos = InStr(StrtPos + 1, UserCommand, "`")
        Loop
        UserInfo(Index).Data = Mid(UserCommand, StrtPos + 3, (Len(UserCommand) - StrtPos) - 2)
        UserCommand = Mid(UserCommand, 1, StrtPos - 1)
        UserCommand = Replace(UserCommand, "`", "")
    Else
        UserInfo(Index).Data = UserInfo(Index).Data & UserCommand
        Exit Sub
    End If
    '<<E>> Recieved Data Buffer <<E>>

From the client:

'<<B>> Recieved Data Buffer <<B>>
ServerCommand = UserInfo.Data & ServerCommand
If Len(ServerCommand) > 32000 Then
    ServerCommand = ""
    UserInfo.Data = ""
    Exit Sub
End If
If InStr(1, ServerCommand, "`") Then
    Do While InStr(StrtPos + 1, ServerCommand, "`") > 0
        StrtPos = InStr(StrtPos + 1, ServerCommand, "`")
    Loop
    UserInfo.Data = Mid(ServerCommand, StrtPos + 3, (Len(ServerCommand) - StrtPos) - 2)
    ServerCommand = Mid(ServerCommand, 1, StrtPos - 1)
    ServerCommand = Replace(ServerCommand, "`", "")
Else
    UserInfo.Data = UserInfo.Data & ServerCommand
    Exit Sub
End If
'<<E>> Recieved Data Buffer <<E>>

From both:

Public Function Send(Index As Integer, Data As String)
On Error Resume Next
frmMain.sckMain(Index).SendData Data & "`"
End Function