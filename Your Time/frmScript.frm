VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScript 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Skript"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5715
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImgList 
      Left            =   5040
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstData 
      Height          =   1695
      Left            =   960
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2990
      View            =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
      _Version        =   393217
      HideSelection   =   0   'False
      TextRTF         =   $"frmScript.frx":0725
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Lagre"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Åpne"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Kjør"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Innstilninger:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   5175
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         ItemData        =   "frmScript.frx":07A7
         Left            =   2400
         List            =   "frmScript.frx":07B4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblScriptType 
         Caption         =   "Skript språk:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmScript"
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

Private Sub cmdClose_Click()

Me.Hide

End Sub

Private Sub cmdOpen_Click()

Dim strFile As String

strFile = Common.ShowOpen(Me.hWnd)

If strFile <> "" And Dir(strFile) <> "" Then
    txtCode.Text = Script.LoadFile(strFile)
End If

End Sub

Private Sub cmdRun_Click()

Script.Language = cmbLanguage.ListIndex
Script.Run txtCode.Text

Me.Show

End Sub

Private Sub cmdSave_Click()

Dim strFile As String

strFile = Common.ShowSave(Me.hWnd)
Script.SaveFile strFile, txtCode.Text, False

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

cmbLanguage.ListIndex = Script.Language

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hWnd

End Sub

Private Sub lstData_ItemClick(ByVal Item As MSComctlLib.ListItem)

Item.Selected = True
lstData_KeyDown 13, 0

End Sub

Private Sub lstData_KeyDown(KeyCode As Integer, Shift As Integer)

Dim sTmp$, lTmp&, lastSel&

If KeyCode = 13 Or KeyCode = vbKeySpace Then
    sTmp = txtCode.Text
    lastSel = txtCode.SelStart + 1
    lTmp = InStrRev(sTmp, ".", lastSel)
    
    If lTmp > 0 Then
        txtCode.Text = Mid(sTmp, lTmp, lastSel - lTmp) & lstData.SelectedItem.Text & Mid(sTmp, lastSel)
    Else
        txtCode.Text = sTmp & lstData.SelectedItem.Text
    End If
    
    txtCode.SetFocus
    txtCode.SelStart = lastSel - 1 + Len(lstData.SelectedItem.Text)
    
    ' Hide the listbox
    lstData.Visible = False
    lstData.Enabled = False
End If

End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)

Dim curSel&, aPt As POINTAPI, ret&, aDC&, textHeight, aObjects, strLastWord$

aObjects = Script.GetScriptObjects

' Trigger the pop-up when the user write a 'puntum'
If KeyCode = 190 Then
    
    aDC = GetWindowDC(txtCode.hWnd)
    GetTextExtentPoint32 aDC, "qD", 2, aPt
    
    textHeight = aPt.Y
    
    curSel = txtCode.SelStart
    SendMessage txtCode.hWnd, EM_POSFROMCHAR, VarPtr(aPt), ByVal curSel + 1
    
    ' Get the last word
    strLastWord = LastWord(txtCode, curSel)
    
    If Script.InArray(Script.GetScriptObjects, strLastWord) >= 0 Then
    
        Script.EnumerateMethods lstData, Script.ObjectPtr(strLastWord)
    
        lstData.Left = aPt.X + txtCode.Left
        lstData.Top = aPt.Y + txtCode.Top + textHeight
        lstData.Visible = True
        lstData.Enabled = True
        lstData.SetFocus
        
    End If
Else
    lstData.Visible = False
End If

End Sub

Public Function LastWord(RichTextBox As RichTextBox, lStart As Long) As String

Dim lTmp&, strText$

strText = Mid(RichTextBox.Text, Script.NotOver(InStrRev(RichTextBox.Text, Chr(13), lStart), Flag_NotUnderOne))
lTmp = InStrRev(strText, " ", lStart, vbTextCompare)

If lTmp = 0 Then
    
    If lStart > 0 Then
        LastWord = Replace(Mid(strText, 1, lStart), vbNewLine, "")
    End If
    
    Exit Function
End If

LastWord = Replace(Mid(strText, lTmp, lStart - lTmp), vbNewLine, "")

End Function

Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

lstData.Visible = False

End Sub
