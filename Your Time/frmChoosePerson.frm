VERSION 5.00
Begin VB.Form frmChoosePerson 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Velg person"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox cmbChoose 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      ToolTipText     =   "Velg ønsket person her."
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "frmChoosePerson"
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

Public Sub AddPersons()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".AddPersons", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell As Long

cmbChoose.Clear

For Tell = Abs(SafeLBound(VarPtrArray(Peoples))) To SafeUBound(VarPtrArray(Peoples))
    cmbChoose.AddItem Peoples(Tell).Name
Next

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
' Unsubclass the form
UnHookForm Me.hwnd

End Sub
