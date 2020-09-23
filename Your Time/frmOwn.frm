VERSION 5.00
Begin VB.Form frmOwn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Egne merkedager"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5625
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtExpression 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   3330
      TabIndex        =   10
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   690
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1170
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.VScrollBar VSSOwn 
      Height          =   3495
      Left            =   5280
      Max             =   285
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtMonth 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   405
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "0"
      Top             =   360
      Width           =   300
   End
   Begin VB.TextBox txtDay 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   300
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblExpression 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uttrykk:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3330
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblText 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tekst:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1170
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   368
      X2              =   8
      Y1              =   249
      Y2              =   249
   End
   Begin VB.Label lblDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dd  mm  åååå"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   368
      X2              =   8
      Y1              =   248
      Y2              =   248
   End
End
Attribute VB_Name = "frmOwn"
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

Dim ObjLoaded As Boolean
Dim TmpData() As Remember

Private Sub cmdCancel_Click()

Me.Hide

End Sub

Private Sub cmdOK_Click()

Save VSSOwn.Tag
CopyDatabase Users(UserID).DataOwn, TmpData

Me.Hide
frmMain.uscCalender.Redraw

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

LoadControls
ObjLoaded = True

VSSOwn.Tag = 0
VSSOwn.Value = 0
Update 0

End Sub

Public Sub LoadDatabase()

CopyDatabase TmpData, Users(UserID).DataOwn

End Sub

Public Sub LoadControls()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".LoadControls", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, Cnt&

If ObjLoaded = False Then

    SetNumber txtYear(0), True
    SetNumber txtMonth(0), True
    SetNumber txtDay(0), True

    Cnt = (VSSOwn.Height - lblText.Height) / txtText(0).Height

    For Tell = 1 To Cnt
        Load txtYear(Tell)
        Load txtMonth(Tell)
        Load txtDay(Tell)
        Load txtText(Tell)
        Load txtExpression(Tell)
        
        SetNumber txtYear(Tell), True
        txtYear(Tell).Visible = True
        txtYear(Tell).Top = txtYear(Tell - 1).Top + txtYear(Tell).Height - 1
        
        SetNumber txtMonth(Tell), True
        txtMonth(Tell).Visible = True
        txtMonth(Tell).Top = txtMonth(Tell - 1).Top + txtMonth(Tell).Height - 1
        
        SetNumber txtDay(Tell), True
        txtDay(Tell).Visible = True
        txtDay(Tell).Top = txtDay(Tell - 1).Top + txtDay(Tell).Height - 1
        
        txtText(Tell).Visible = True
        txtText(Tell).Top = txtText(Tell - 1).Top + txtText(Tell).Height - 1
        
        txtExpression(Tell).Visible = True
        txtExpression(Tell).Top = txtExpression(Tell - 1).Top + txtExpression(Tell).Height - 1
    Next

End If

End Sub

Public Sub Update(AddValue As Integer)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".Update(AddValue)", Array(AddValue), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

VSSOwn.Max = UBound(TmpData) - txtText.count - 1

For Tell = 0 To txtText.count - 1
    txtYear(Tell).Text = Year(TmpData(Tell + AddValue).RemDate)
    txtMonth(Tell).Text = Month(TmpData(Tell + AddValue).RemDate)
    txtDay(Tell).Text = Day(TmpData(Tell + AddValue).RemDate)
    txtExpression(Tell).Text = TmpData(Tell + AddValue).ExtraData
    txtText(Tell).Text = TmpData(Tell + AddValue).Text
Next

End Sub

Public Sub Save(AddValue As Integer)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".Save(AddValue)", Array(AddValue), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

For Tell = 0 To txtText.count - 1
    TmpData(Tell + AddValue).RemDate = DateSerial(txtYear(Tell).Text, txtMonth(Tell).Text, txtDay(Tell).Text)
    TmpData(Tell + AddValue).ExtraData = txtExpression(Tell).Text
    TmpData(Tell + AddValue).Text = txtText(Tell).Text
    TmpData(Tell + AddValue).Enabled = CBool(txtText(Tell).Text <> "")
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hWnd

End Sub

Private Sub VSSOwn_Change()

Save Val(VSSOwn.Tag)
Update VSSOwn.Value

VSSOwn.Tag = VSSOwn.Value

End Sub
