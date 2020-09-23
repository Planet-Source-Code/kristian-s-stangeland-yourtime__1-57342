VERSION 5.00
Begin VB.Form frmErrorInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ekstra feilmeldingsinformasjon"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEvents 
      Height          =   1125
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblEvents 
      Caption         =   "Hendelser:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblEmail 
      Caption         =   "E-post:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblName 
      Caption         =   "Navn:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmErrorInfo"
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
UnHookForm Me.hwnd

End Sub
