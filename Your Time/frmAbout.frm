VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5760
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4290
      TabIndex        =   2
      Top             =   4320
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4275
      TabIndex        =   1
      Top             =   3900
      Width           =   1260
   End
   Begin VB.PictureBox picAnimation 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      HasDC           =   0   'False
      Height          =   3375
      Left            =   120
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   367
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Dette programmet er til helt fri bruk, men under beskyttelse av GPL (General Public License) . Les mer om dette i vedlagt fil."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   120
      TabIndex        =   3
      Top             =   3900
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   135
      X2              =   5670
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5685
      Y1              =   3720
      Y2              =   3720
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim Stjerner() As Stjerner
Dim Looping As Boolean
Dim Buffer As Buffer
Dim Roller() As Line
Dim lCount As Long

Private Sub cmdOK_Click()

UnLoadBuffer
Looping = False
frmAbout.Hide

End Sub

Private Sub cmdSysInfo_Click()

StartSysInfo

End Sub

Sub LoadStjerner(ByVal Amout As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".LoadStjerner(Amout)", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim a&

ReDim Stjerner(Amout)

For a = LBound(Stjerner) To UBound(Stjerner)
    Stjerner(a).X = Rnd * picAnimation.ScaleWidth
    Stjerner(a).Y = Rnd * picAnimation.ScaleHeight
    Stjerner(a).r = Rnd * 255
    Stjerner(a).Color = RGB(Stjerner(a).r, Stjerner(a).r, Stjerner(a).r)
Next

End Sub

Sub LoadBuffer()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".LoadBuffer", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

UnLoadBuffer
Buffer.hdc = CreateCompatibleDC(frmAbout.hdc)
Buffer.Picture = CreateCompatibleBitmap(frmAbout.hdc, picAnimation.ScaleWidth, picAnimation.ScaleHeight)
SelectObject Buffer.hdc, Buffer.Picture

Buffer.ScaleWidth = picAnimation.ScaleWidth
Buffer.ScaleHeight = picAnimation.ScaleHeight

End Sub

Sub UnLoadBuffer()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".UnLoadBuffer", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

DeleteDC Buffer.hdc
DeleteObject Buffer.Picture

End Sub

Sub MainLoop()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".MainLoop", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim a&, Y#, Rect As Rect, Speed&

Show

Speed = Script.AnimationSpeed
Looping = True

LoadStjerner Script.StarAmout
LoadBuffer

SetBkMode Buffer.hdc, 1
SetTextColor Buffer.hdc, &HC0C0&

Y = picAnimation.ScaleHeight

Do Until Looping = False
    BitBlt Buffer.hdc, 0, 0, Buffer.ScaleWidth, Buffer.ScaleHeight, 0, 0, 0, vbWhiteness
    
    Y = Y - 0.5
    
    For a = LBound(Stjerner) To UBound(Stjerner)
        Stjerner(a).Y = Stjerner(a).Y - (Stjerner(a).r / 100)
        If Stjerner(a).Y < 0 Then Stjerner(a).Y = Buffer.ScaleHeight
        SetPixelV Buffer.hdc, Stjerner(a).X, Stjerner(a).Y, Stjerner(a).Color
    Next
    
    ClearRollerText
    AddRollerLine "Utvikleren:", FW_BOLD
    AddRollerLine Arthor, FW_NORMAL
    AddRollerLine "", 0
    AddRollerLine "Beta-testere", FW_BOLD
    AddRollerLine Arthor, FW_NORMAL
    AddRollerLine "Christopher Williams", FW_NORMAL
    AddRollerLine "Kristian Stangeland", FW_NORMAL
    AddRollerLine "Christer Nilsen", FW_NORMAL
    AddRollerLine "Torbjørn Tobsen", FW_NORMAL
    AddRollerLine "Tommy Johnsen", FW_NORMAL
    AddRollerLine "", 0
    AddRollerLine "Psykologisk hjelp:", FW_BOLD
    AddRollerLine "Christopher Williams", FW_NORMAL
    AddRollerLine "Kristian Stangeland", FW_NORMAL
    AddRollerLine "", 0
    AddRollerLine "Utviklet i året 2004", FW_NORMAL
    AddRollerLine "", 0
    AddRollerLine "Jeg fraskriver meg enhver ansvar", FW_NORMAL
    AddRollerLine "for skade påført av dette program!", FW_NORMAL
    AddRollerLine "", 0
    AddRollerLine "Takk til URFIN JUS (www.urfinjus.net) for HuntERR", FW_NORMAL
    AddRollerLine "", 0
    AddRollerLine "Versjon: " & Script.Version, FW_NORMAL
    
    For a = 1 To lCount
        Buffer.Font = CreateFont(16, 0, 0, 0, Roller(a).Style, False, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_MASK, DEFAULT_QUALITY, 34, picAnimation.FontName)
        SelectObject Buffer.hdc, Buffer.Font
        SetRect Rect, 0, Y + ((a - 1) * 16), picAnimation.ScaleWidth, Y + (a * 16)
        DrawTextA Buffer.hdc, Roller(a).Text, Len(Roller(a).Text), Rect, DT_CENTER
        DeleteObject Buffer.Font
    Next
    
    If Y < -(lCount * 16) Then Y = Buffer.ScaleHeight
    
    BitBlt picAnimation.hdc, 0, 0, Buffer.ScaleWidth, Buffer.ScaleHeight, Buffer.hdc, 0, 0, vbSrcCopy
    
    If Speed > 0 Then Sleep Speed
    DoEvents
Loop

UnLoadBuffer

End Sub

Public Sub ClearRollerText()

Erase Roller
lCount = 0

End Sub

Public Sub AddRollerLine(Text As String, ByVal Style As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".AddRollerLine(Text, Style)", Array(Text, Style), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

lCount = lCount + 1

ReDim Preserve Roller(1 To lCount)

Roller(lCount).Text = Text
Roller(lCount).Style = Style

End Sub

Public Sub StopAnimation()

UnLoadBuffer
Looping = False

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hwnd

StopAnimation

End Sub

Public Sub StartSysInfo()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".StartSysInfo", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim rc As Long
Dim SysInfoPath As String

If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then

    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
    Else
        Err.Number = 53
        Err.Description = Error(53)
        GoTo errHandler
    End If

Else
    Err.Number = 1000
    Err.Description = "System information is unavailable at this time"
    GoTo errHandler
End If

Shell SysInfoPath, vbNormalFocus

End Sub
