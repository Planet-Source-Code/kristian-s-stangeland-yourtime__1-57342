VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDiary 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dagbok"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8385
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   559
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLast 
      Caption         =   "Forrige"
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Neste"
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmDiary.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   375
      Index           =   1
      Left            =   480
      Picture         =   "frmDiary.frx":0376
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   375
      Index           =   2
      Left            =   840
      Picture         =   "frmDiary.frx":06F8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      Picture         =   "frmDiary.frx":0A76
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1680
      Picture         =   "frmDiary.frx":0DEA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2040
      Picture         =   "frmDiary.frx":1171
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3720
      Picture         =   "frmDiary.frx":1508
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   4080
      Picture         =   "frmDiary.frx":1583
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   4440
      Picture         =   "frmDiary.frx":15FE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   4920
      Picture         =   "frmDiary.frx":1679
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   5280
      Picture         =   "frmDiary.frx":1A10
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   5640
      Picture         =   "frmDiary.frx":1A98
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      Left            =   6120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox RichTextBox 
      Height          =   4455
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7858
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmDiary.frx":1DFD
   End
   Begin RichTextLib.RichTextBox RichBuffer 
      Height          =   735
      Left            =   6960
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmDiary.frx":1E80
   End
   Begin VB.Label lblStat 
      Caption         =   "Dag: 25.06.2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5220
      Width           =   5775
   End
End
Attribute VB_Name = "frmDiary"
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

Dim CurrRecord As Long

Private Sub cmdEdit_Click(Index As Integer)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdEdit_Click(Index)", Array(Index), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Select Case Index
Case 0
    RichTextBox.TextRTF = ""

Case 1
    Printer.Print ""
    RichTextBox.SelStart = 1
    RichTextBox.SelLength = Len(RichTextBox.Text)
    RichTextBox.SelPrint Printer.hdc
    Printer.EndDoc

Case 2
    CurrRecord = SaveRecord(Users(UserID).DataDiary, RichTextBox.TextRTF, CurrentDate.Contents, RichTextBox.Text)
    UpdateButtons
    
Case 3, 4
    Clipboard.Clear
    Clipboard.SetText RichTextBox.SelText
    
    If Index = 3 Then
        RichTextBox.SelText = ""
    End If

Case 5
    RichTextBox.SelText = Clipboard.GetText

Case 6
    RichTextBox.SelBold = Not RichTextBox.SelBold

Case 7
    RichTextBox.SelItalic = Not RichTextBox.SelItalic

Case 8
    RichTextBox.SelUnderline = Not RichTextBox.SelUnderline

Case 9, 10, 11
    RichTextBox.SelAlignment = Choose(Index - 8, 0, 2, 1)

Case 12
    RichTextBox.SelColor = Common.ShowColor(Me.hWnd)

Case 13
    Users(UserID).DataDiary(CurrRecord).Text = ""
    Users(UserID).DataDiary(CurrRecord).Enabled = False
    RichTextBox.TextRTF = ""

Case 14
    SendMessage RichTextBox.hWnd, EM_UNDO, 0, 0

End Select

RichTextBox.SetFocus

End Sub

Private Sub cmbFont_Change()

RichTextBox.SelFontName = cmbFont.Text

End Sub

Private Sub cmbFont_Click()

cmbFont_Change

End Sub

Private Sub cmdLast_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdLast_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

CurrRecord = CurrRecord - 1

frmMain.ProcessControls "saveall"
CurrentDate.Contents = Users(UserID).DataDiary(CurrRecord).RemDate
frmMain.ProcessControls "update"

LoadCurrentDate

End Sub

Private Sub cmdNext_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdNext_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

CurrRecord = CurrRecord + 1

frmMain.ProcessControls "saveall"
CurrentDate.Contents = Users(UserID).DataDiary(CurrRecord).RemDate
frmMain.ProcessControls "update"

LoadCurrentDate

End Sub

Public Sub LoadCurrentDate()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".LoadCurrentDate", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

RichTextBox.TextRTF = Script.LoadDiary(CurrentDate.Contents, CurrRecord)
lblStat.Caption = Language.Constant("Day") & ": " & CurrentDate.Contents
UpdateButtons

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

' Fill the combobox with all the fonts
FillComboWithFonts cmbFont

LoadCurrentDate
RichTextBox_Change

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hWnd

cmdEdit_Click 2

End Sub

Private Sub RichTextBox_Change()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".RichTextBox_Change", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, FontName$

cmdEdit(5).Enabled = Not SendMessage(RichTextBox.hWnd, EM_CANPASTE, 0, 0) = 0
cmdEdit(14).Enabled = Not SendMessage(RichTextBox.hWnd, EM_CANUNDO, 0, 0) = 0

FontName = RichTextBox.SelFontName

If LCase(cmbFont.Text) <> LCase(FontName) Then

    For Tell = 0 To cmbFont.ListCount - 1
        If LCase(cmbFont.List(Tell)) = LCase(FontName) Then
            cmbFont.ListIndex = Tell
        End If
    Next
End If

End Sub

Public Sub UpdateButtons()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".UpdateButtons", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

cmdLast.Enabled = CBool(CurrRecord > 0)
cmdNext.Enabled = CBool(CurrRecord < SafeUBound(VarPtrArray(Users(UserID).DataDiary)))
frmMain.ProcessControls "updatediary"

End Sub
