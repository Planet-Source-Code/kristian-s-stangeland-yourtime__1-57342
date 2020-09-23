VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sikkerhetskopi"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6930
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameGetbackup 
      Caption         =   "&Gjennopprett all data fra fil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6615
      Begin VB.CommandButton cmdGetBrowse 
         Caption         =   "F&inn"
         Height          =   280
         Left            =   4080
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "&Hent"
         Height          =   280
         Left            =   5280
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtGetBackup 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Text            =   "C:\"
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Frame frameBackup 
      Caption         =   "K&opier all data til fil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin VB.CommandButton cmdCopyBrowse 
         Caption         =   "&Finn"
         Height          =   280
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdMake 
         Caption         =   "&Kopier"
         Height          =   280
         Left            =   5280
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtCopyPath 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "C:\"
         Top             =   480
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmBackup"
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

Private Sub cmdCopyBrowse_Click()

txtCopyPath.Text = Common.ShowSave(Me.hWnd)

End Sub

Private Sub cmdGet_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdGet_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim BuffProp As New PropertyBag, Tmp, Tell&, File$

BuffProp.Contents = Script.LoadFile(txtGetBackup.Text)
Tmp = Split(BuffProp.ReadProperty("FileNames", ""), ";")

File = Dir(ValidPath(App.Path) & "Data\")

Do Until File = ""

    If GetExtention(File) = "dat" Then
        Kill File
    End If

    File = Dir
Loop

For Tell = LBound(Tmp) To UBound(Tmp)
    If Tmp(Tell) <> "" Then
        Script.SaveFile ValidPath(App.Path) & "Data\" & Tmp(Tell), BuffProp.ReadProperty(Tmp(Tell), ""), False
    End If
Next

End Sub

Private Sub cmdGetBrowse_Click()

txtGetBackup.Text = Common.ShowOpen(Me.hWnd)

End Sub

Private Sub cmdMake_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdMake_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim BuffProp As New PropertyBag, File$, FileNames$

File = Dir(ValidPath(App.Path) & "Data\")

Do Until File = ""

    If GetExtention(File) = "dat" Then
        BuffProp.WriteProperty GetFileName(File), Script.LoadFile(File), ""
        FileNames = FileNames & GetFileName(File) & ";"
    End If
    
    File = Dir
Loop

BuffProp.WriteProperty "FileNames", FileNames, ""
Script.SaveFile txtCopyPath.Text, BuffProp.Contents, False

Set BuffProp = Nothing

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hWnd

End Sub
