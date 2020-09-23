VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bruker"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4890
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   750
      Width           =   2895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1140
      Width           =   2895
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblPassword 
      Caption         =   "Passord:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   780
      Width           =   1575
   End
   Begin VB.Label lblOldPassword 
      Caption         =   "Gammelt passord:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblUserName 
      Caption         =   "Brukernavn:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmUser"
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

Me.Tag = "cancel"
Me.Hide

End Sub

Private Sub cmdOK_Click()

Me.Tag = "success"
Me.Hide

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hWnd

End Sub
