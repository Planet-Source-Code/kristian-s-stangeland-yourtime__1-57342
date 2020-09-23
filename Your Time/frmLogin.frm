VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Logg på bruker"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4950
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNewUser 
      Caption         =   "&Ny bruker"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cmbUserName 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblUsername 
      Caption         =   "Brukernavn:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Passord:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogin"
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

Private Sub cmdCancel_Click()

Me.Top = "cancel"
Me.Hide

End Sub

Private Sub cmdOK_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdOK_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim TmpID&

TmpID = Script.FindUser(cmbUserName.Text)

If TmpID < 1 Then
    MsgBox "Feil: Kan ikke finne bruker ved navn " & cmbUserName.Text, vbCritical, "Feil"
    Exit Sub
End If

MD5.InBuffer = txtPassword.Text
MD5.Calculate

If MD5.OutBuffer = Users(TmpID).Password Or Users(TmpID).Password = "" Then

    UserID = TmpID

    Me.Tag = "success"
    Me.Hide
Else
    MsgBox "Feil: Passordet er ikke korrekt!", vbCritical, "Feil"
    txtPassword.SetFocus
End If

End Sub

Private Sub cmdNewUser_Click()

frmSettings.AddNewUser
UpdateUsers

End Sub

Private Sub Form_Load()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".Form_Load", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

' Common code that should be executet in all forms
FormLoad Me

UpdateUsers

End Sub

Public Sub UpdateUsers()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".UpdateUsers", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell As Long

cmbUserName.Clear

For Tell = LBound(Users) To UBound(Users)
    If Users(Tell).UserName <> "" Then
        cmbUserName.AddItem Users(Tell).UserName
    End If
Next

cmbUserName.ListIndex = UserID - 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hwnd

End Sub
