VERSION 5.00
Begin VB.Form frmPhonebook 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Telefonkatalog"
   ClientHeight    =   4875
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7815
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin YourTime.SheetTabs SheetTabs1 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   450
      TabBackColor    =   12632256
      ExtraWidth      =   0
      FontSize        =   8
   End
   Begin VB.VScrollBar vscPhone 
      Height          =   3855
      Left            =   7500
      Max             =   500
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox picBook 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   3855
      Left            =   120
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   3
      Top             =   600
      Width           =   7335
      Begin VB.Label lblAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   7
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   6
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblFirm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Lukk"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   375
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPhonebook.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdNew 
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Rediger"
      Begin VB.Menu mnuNewName 
         Caption         =   "Nytt navn"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetName 
         Caption         =   "Hent navn"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Søk"
      Begin VB.Menu mnuFind 
         Caption         =   "Finn"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Finn etter"
         Begin VB.Menu mnuFindBy 
            Caption         =   "Navn"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuFindBy 
            Caption         =   "Firma"
            Index           =   1
         End
         Begin VB.Menu mnuFindBy 
            Caption         =   "Addresse"
            Index           =   2
         End
         Begin VB.Menu mnuFindBy 
            Caption         =   "Telefonnummer"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "frmPhonebook"
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

Dim lpArray() As People
Dim lSelected As Long

Private Sub cmdClose_Click()

Me.Hide

End Sub

Private Sub cmdNew_Click()

mnuNewName_Click

End Sub

Private Sub cmdSearch_Click()

mnuFind_Click

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

lSelected = -1

ProcessControls "sheettabs"
ProcessControls "allocatecontrols"
ProcessControls "update"
ProcessControls "refreshselected"

' Icon's
Set cmdNew.Picture = frmMain.cmdAny(4).Picture

End Sub

Public Sub ProcessControls(strCommand As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".ProcessControls(strCommand)", Array(strCommand), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, lIndex&, uCtr&

Select Case LCase(strCommand)
Case "allocatecontrols"

    If lblName.count < picBook.ScaleHeight / lblName(0).Height Then

        For Tell = 1 To picBook.ScaleHeight / lblName(0).Height
        
            Load lblName(Tell)
            Load lblFirm(Tell)
            Load lblAddress(Tell)
            Load lblPhone(Tell)
        
            lblName(Tell).Visible = True
            lblFirm(Tell).Visible = True
            lblAddress(Tell).Visible = True
            lblPhone(Tell).Visible = True
            
            lblName(Tell).Top = lblName(Tell - 1).Top + lblName(Tell).Height - 1
            lblFirm(Tell).Top = lblName(Tell).Top
            lblAddress(Tell).Top = lblName(Tell).Top
            lblPhone(Tell).Top = lblName(Tell).Top
        Next

    End If

Case "update"

    Erase lpArray
    lpArray = FindPerson(SheetTabs1.Tabs(SheetTabs1.Selected), 0)

    ProcessControls "refreshvariables"

Case "refreshvariables"

    uCtr = SafeUBound(VarPtrArray(lpArray))

    For Tell = 0 To lblName.count - 1
        lIndex = Tell + vscPhone.Value
    
        If lIndex <= uCtr Then
            lblName(Tell).Tag = lIndex
            lblName(Tell).Caption = " " & lpArray(lIndex).Name
            lblFirm(Tell).Caption = " " & lpArray(lIndex).Firm
            lblAddress(Tell).Caption = " " & lpArray(lIndex).Address
            lblPhone(Tell).Caption = " " & Choose(lpArray(lIndex).VisibleNum + 1, "Pr.", "Fir.", "Fax.", "Mob.") & " " & Choose(lpArray(lIndex).VisibleNum + 1, lpArray(lIndex).PhoneNum, lpArray(lIndex).FirmNum, lpArray(lIndex).Fax, lpArray(lIndex).MobNum)
        Else
            lblName(Tell).Tag = -1
            lblName(Tell).Caption = ""
            lblFirm(Tell).Caption = ""
            lblAddress(Tell).Caption = ""
            lblPhone(Tell).Caption = ""
        End If
    Next

Case "sheettabs"

    For Tell = 1 To 26
        SheetTabs1.AddTab Chr(vbKeyA + Tell - 1)
    Next

    SheetTabs1.AddTab "(S)"
    SheetTabs1.ExtraWidth = 15
    SheetTabs1.Redraw
    
Case "refreshselected"

    For Tell = 0 To lblName.count - 1
        SetColor Tell, vbWindowBackground, vbWindowText
    Next

    If lSelected - vscPhone.Value >= 0 And lSelected - vscPhone.Value <= lblName.count - 1 Then
        SetColor lSelected - vscPhone.Value, vbHighlight, vbHighlightText
    End If

End Select

End Sub

Public Sub SetColor(Index As Long, BackColor As Long, ForeColor As Long)

lblName(Index).BackColor = BackColor
lblName(Index).ForeColor = ForeColor
lblFirm(Index).BackColor = BackColor
lblFirm(Index).ForeColor = ForeColor
lblAddress(Index).BackColor = BackColor
lblAddress(Index).ForeColor = ForeColor
lblPhone(Index).BackColor = BackColor
lblPhone(Index).ForeColor = ForeColor

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hwnd

End Sub

Private Sub lblAddress_Click(Index As Integer)

lSelected = Index + vscPhone.Value
ProcessControls "refreshselected"

End Sub

Private Sub lblFirm_Click(Index As Integer)

lSelected = Index + vscPhone.Value
ProcessControls "refreshselected"

End Sub

Private Sub lblName_Click(Index As Integer)

lSelected = Index + vscPhone.Value
ProcessControls "refreshselected"

End Sub

Private Sub lblPhone_Click(Index As Integer)

lSelected = Index + vscPhone.Value
ProcessControls "refreshselected"

End Sub

Private Sub mnuFind_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".mnuFind_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim strSearchText As String

strSearchText = InputBox("Søketekst:", "Søk etter")
If strSearchText = "" Then Exit Sub

SheetTabs1.Selected = SheetTabs1.TabCount

lpArray = FindPerson(strSearchText, Script.FindSelected(mnuFindBy), True)
ProcessControls "refreshvariables"

End Sub

Private Sub mnuFindBy_Click(Index As Integer)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".mnuFindBy_Click(Index)", Array(Index), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

For Tell = 0 To mnuFindBy.count - 1
    mnuFindBy(Tell).Checked = False
Next

mnuFindBy(Index).Checked = True

End Sub

Private Sub mnuNewName_Click()

frmPerson.Show

End Sub

Private Sub SheetTabs1_TabChanged(LastIndex As Long)

ProcessControls "update"
ProcessControls "refreshselected"

End Sub

Private Sub vscPhone_Change()

ProcessControls "update"
ProcessControls "refreshselected"

End Sub
