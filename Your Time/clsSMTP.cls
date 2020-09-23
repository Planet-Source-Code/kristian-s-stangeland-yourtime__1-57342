VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Dim WithEvents Sock As clsSocket
Attribute Sock.VB_VarHelpID = -1

Dim m_Server As String      ' mail server host
Dim m_Port As String        ' mail server port
Dim m_MailFrom As String    ' from address
Dim m_MailTo As String      ' to address
Dim m_BCC As String         ' blind carbon copy addresses
Dim m_CCC As String         ' carbon copy addresses
Dim m_Subject As String     ' email subject
Dim m_NameFrom As String    ' from name
Dim m_NameTo As String      ' to name
Dim m_Body As String        ' email body
Dim m_Log As String         ' log of transaction

' private state variables
Dim LastResponse As String

Public Property Get Log() As String
    Log = m_Log
End Property

Public Property Get Name() As String
    Name = "SMTP"
End Property

Public Property Get Server() As String
    Server = m_Server
End Property

Public Property Let Server(ByVal Data As String)
    m_Server = Data
End Property

Public Property Get Port() As String
    Port = m_Port
End Property

Public Property Let Port(ByVal Data As String)
    m_Port = Data
End Property

Public Property Get MailFrom() As String
    MailFrom = m_MailFrom
End Property

Public Property Let MailFrom(ByVal Data As String)
    m_MailFrom = Data
End Property

Public Property Get MailTo() As String
    MailTo = m_MailTo
End Property

Public Property Let MailTo(ByVal Data As String)
    m_MailTo = Data
End Property

Public Property Get BCC() As String
    BCC = m_BCC
End Property

Public Property Let BCC(ByVal Data As String)
    m_BCC = Data
End Property

Public Property Get CCC() As String
    CCC = m_CCC
End Property

Public Property Let CCC(ByVal Data As String)
    m_CCC = Data
End Property

Public Property Get Subject() As String
    Subject = m_Subject
End Property

Public Property Let Subject(ByVal Data As String)
    m_Subject = Data
End Property

Public Property Get NameTo() As String
    NameTo = m_NameTo
End Property

Public Property Let NameTo(ByVal Data As String)
    m_NameTo = Data
End Property

Public Property Get NameFrom() As String
    NameFrom = m_NameFrom
End Property

Public Property Let NameFrom(ByVal Data As String)
    m_NameFrom = Data
End Property

Public Property Get Body() As String
    Body = m_Body
End Property

Public Property Let Body(ByVal Data As String)
    m_Body = Data
End Property

Private Function ConnectToServer() As Boolean

Sock.CloseSocket
Sock.RemoteHost = Me.Server
Sock.RemotePort = Me.Port
Sock.Connect

Do While Sock.State <> sckConnected
    Sleep 10
    DoEvents
    
    If Sock.State = sckError Then
        Exit Function
    End If
Loop

ConnectToServer = True

End Function

Public Function SendMail() As Long

Dim SMTPCommands(0 To 10) As String
Dim SMTPResponses(0 To 10) As String

Dim Success As Boolean
Dim i As Integer

' construct an array of commands
SMTPCommands(0) = "HELO " & Me.Server
SMTPCommands(1) = "MAIL FROM:" & Me.MailFrom
SMTPCommands(2) = "RCPT TO:" & Me.MailTo
SMTPCommands(3) = "DATA"
SMTPCommands(4) = "BCC:" & Me.BCC
SMTPCommands(5) = "CCC:" & Me.CCC
SMTPCommands(6) = "SUBJECT:" & Me.Subject
SMTPCommands(7) = "TO:" & Me.NameTo
SMTPCommands(8) = "FROM:" & Me.NameFrom & vbCrLf ' extra vbCrLf
SMTPCommands(9) = Me.Body & vbCrLf & "."
SMTPCommands(10) = "QUIT"

SMTPResponses(0) = "250"
SMTPResponses(1) = "250"
SMTPResponses(2) = "250"
SMTPResponses(3) = "354"
SMTPResponses(4) = ""
SMTPResponses(5) = ""
SMTPResponses(6) = ""
SMTPResponses(7) = ""
SMTPResponses(8) = ""
SMTPResponses(9) = "250"
SMTPResponses(10) = "221"

' Reset the log
m_Log = ""

' connect to the server
If ConnectToServer = False Then
    MsgBox "Error: " & "Couldn't connect to server", vbCritical
    Exit Function
Else
    ' wait for the welcome message to be received
    WaitForResponse "220"
End If

' send each command, waiting for a response
For i = 0 To 10

    ' send the command
    SMTPSend SMTPCommands(i)
    
    ' wait for the response
    Success = WaitForResponse(SMTPResponses(i))
    
    ' check if we were successful
    If Success = False Then
        m_Log = m_Log & "ERROR: Request timed out" & vbNewLine
        SendMail = -1
        Exit Function
    End If
Next

End Function

Private Sub SMTPSend(ByVal Data As String)

' Log the data
m_Log = m_Log & Data & vbNewLine

If Sock.State = sckConnected Then
    Sock.SendData Data & vbCrLf
    DoEvents
End If

End Sub

Private Function WaitForResponse(ByVal Response As String) As Boolean

' Log the data
m_Log = m_Log & Response & vbNewLine

' if we're not waiting for a response then exit
If Response = "" Then
    WaitForResponse = True
    Exit Function
Else

    Do While LastResponse = ""
        Sleep 10
        DoEvents
    Loop

    WaitForResponse = CBool(Response = LastResponse)
End If

LastResponse = ""

End Function

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)

Dim Data As Variant

Sock.GetData Data
LastResponse = Mid$(Data, 1, 3)

End Sub

Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

ErrorIn "clsSMTP.Error(Number, Description, Scode, Source, HelpFile, HelpContext)", Array(Number, Description, Scode, Source, HelpFile, HelpContext), EA_NORERAISE
HandleError

m_Log = m_Log & "ERROR (" & Number & ") :" & Description & vbNewLine

End Sub

Public Function Clone() As clsSMTP

    Set Clone = New clsSMTP
    
    Clone.Server = m_Server
    Clone.Port = m_Port
    Clone.MailFrom = m_MailFrom
    Clone.MailTo = m_MailTo
    Clone.BCC = m_BCC
    Clone.CCC = m_CCC
    Clone.Subject = m_Subject
    Clone.NameFrom = m_NameFrom
    Clone.NameTo = m_NameTo
    Clone.Body = m_Body
    
End Function
