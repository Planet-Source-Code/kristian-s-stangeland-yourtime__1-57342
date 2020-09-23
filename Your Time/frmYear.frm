VERSION 5.00
Begin VB.Form frmYear 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Årskalender - 2004"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Lukk"
      Height          =   495
      Left            =   7800
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ListBox lstInfo 
      BackColor       =   &H8000000F&
      Height          =   2790
      Left            =   7080
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.OptionButton chkOptions 
      Caption         =   "Dager med oppgaver"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   2775
   End
   Begin VB.OptionButton chkOptions 
      Caption         =   "Dager med avtaler"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   2775
   End
   Begin VB.OptionButton chkOptions 
      Caption         =   "Helligdager og merkedager"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Value           =   -1  'True
      Width           =   2775
   End
   Begin YourTime.YearReview YearReview 
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   4948
   End
   Begin VB.Label lblStat 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmYear"
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

Private Sub chkOptions_Click(Index As Integer)

YearReview.Redraw

End Sub

Private Sub cmdClose_Click()

Me.Hide

End Sub

Private Sub cmdPrint_Click()

PrintClass.PrintForm Me

End Sub

Private Sub YearReview_MouseDown(BlockDate As Date, Button As Integer, Shift As Integer)

If Button = 1 Then
    Me.Hide
    
    frmMain.ProcessControls "saveall"
    CurrentDate.Contents = BlockDate
    frmMain.ProcessControls "update"
End If

End Sub

Private Sub YearReview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblStat.Caption = ""

End Sub

Private Sub YearReview_MouseOver(BlockDate As Date, Shift As Integer)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".YearReview_MouseOver(BlockDate, Shift)", Array(BlockDate, Shift), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim strText As String, hDate&, hDateTwo&, cTmp As New clsDate

lblStat.ForeColor = vbBlack
cTmp.Contents = BlockDate

strText = CapitalizeFirstLetter(WeekdayName(Weekday(BlockDate), , vbUseSystemDayOfWeek)) & _
" den. " & Day(BlockDate) & " " & CapitalizeFirstLetter(MonthName(Month(BlockDate))) & " " & Year(BlockDate) & " - "

If chkOptions(0).Value Then
    hDate = Search(StaticDays, BlockDate, 3)
    hDateTwo = Search(Users(UserID).DataOwn, BlockDate, 3)
    
    If hDate >= 0 Then
        strText = strText & Script.RemoveTags(StaticDays(hDate).Text) & " - "
        If InStr(1, StaticDays(hDate).Text, "(!CH)", vbTextCompare) <> 0 Then lblStat.ForeColor = vbRed
    End If
    
    If hDateTwo >= 0 Then
        strText = strText & Script.RemoveTags(Users(UserID).DataOwn(hDateTwo).Text) & " - "
        If InStr(1, Users(UserID).DataOwn(hDateTwo).Text, "(!CH)", vbTextCompare) <> 0 Then lblStat.ForeColor = vbRed
    End If
    
Else
    lstInfo.Clear
    Script.AddRem lstInfo, BlockDate
    Script.AddTasks lstInfo, BlockDate

End If

' Legg til ukenummeret
strText = strText & "Uke: " & cTmp.cWeekNum

lblStat.Caption = strText

End Sub

Private Sub YearReview_Redrawing(BlockDescription() As String, BlockBackcolor() As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".YearReview_Redrawing(BlockDescription, BlockBackcolor)", Array(VarPtrArray(BlockDescription), VarPtrArray(BlockBackcolor)), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim lineX&, lineY&, Num&, hDate&, lDate As Date

For lineX = 0 To 30
    For lineY = 0 To 11

        If Script.IsDateSerial(CurrentDate.cYear, lineY + 1, lineX + 1) Then
        
            lDate = DateSerial(CurrentDate.cYear, lineY + 1, lineX + 1)
            BlockBackcolor(lineX, lineY) = vbYellow
            
            If chkOptions(0).Value = True Then
            
                ' Helligdager og merkedager
                hDate = Search(StaticDays, lDate, 3)
            
                If hDate >= 0 Then
                    If InStr(1, StaticDays(hDate).Text, "(!CH)") <> 0 Then
                        BlockDescription(lineX, lineY) = "H"
                        BlockBackcolor(lineX, lineY) = vbBlue
                    Else
                        BlockDescription(lineX, lineY) = "M"
                        BlockBackcolor(lineX, lineY) = vbCyan
                    End If
                End If
                
                ' Egne merkedager
                hDate = Search(Users(UserID).DataOwn, lDate, 3)
            
                If hDate >= 0 Then
                    BlockDescription(lineX, lineY) = "E"
                    BlockBackcolor(lineX, lineY) = RGB(128, 128, 128)
                End If
                
            ElseIf chkOptions(1).Value = True Then
            
                ' Dager med avtaler
                
                hDate = Search(Users(UserID).DataRem, lDate, 1)
            
                If hDate >= 0 Then
                    BlockDescription(lineX, lineY) = "*"
                    BlockBackcolor(lineX, lineY) = vbGreen
                End If
            
            Else
            
                ' Dager med oppgaver
            
                hDate = Search(Users(UserID).DataTasks, lDate, 1)
            
                If hDate >= 0 Then
                    BlockDescription(lineX, lineY) = "#"
                    BlockBackcolor(lineX, lineY) = vbRed
                End If
            
            End If
            
        Else
            BlockBackcolor(lineX, lineY) = vbWhite
        End If

    Next
Next

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hwnd

End Sub

