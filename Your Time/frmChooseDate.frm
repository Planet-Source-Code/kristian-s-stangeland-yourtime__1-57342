VERSION 5.00
Begin VB.Form frmChooseDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Velg dato"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEndDate 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Dato (dd.mm.åååå)"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtStartDate 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      ToolTipText     =   "Dato (dd.mm.åååå)"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   795
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblEndDate 
      Caption         =   "Slutt-dato:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblStartDate 
      Caption         =   "Start-dato:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmChooseDate"
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
