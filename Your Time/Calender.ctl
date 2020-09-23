VERSION 5.00
Begin VB.UserControl Calender 
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   ToolboxBitmap   =   "Calender.ctx":0000
   Begin VB.PictureBox imgMoon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008000FF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   4
      Left            =   1560
      Picture         =   "Calender.ctx":0312
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox imgMoon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008000FF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   3
      Left            =   1320
      Picture         =   "Calender.ctx":0669
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox imgMoon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008000FF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   1080
      Picture         =   "Calender.ctx":09D1
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox imgMoon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008000FF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   840
      Picture         =   "Calender.ctx":0D38
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox imgMoon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008000FF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   600
      Picture         =   "Calender.ctx":10A0
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picCalender 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   311
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   343
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
   Begin VB.Image imgFlag 
      Height          =   210
      Left            =   1920
      Picture         =   "Calender.ctx":13F7
      Top             =   5520
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "Calender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event DateChanging()
Event DateChanged(LastDate As Date)
Event Redrawing(ItemCount As Long, lpTextArray() As String, MoonVal() As Long, ExtFlag() As Long)

Dim lDate As Date, lSelector As Boolean, TextArray() As String

Private Sub picCalender_KeyDown(KeyCode As Integer, Shift As Integer)

RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next
lDate = PropBag.ReadProperty("CalenderDate", Date)
lSelector = PropBag.ReadProperty("Selector", True)
picCalender.FontSize = PropBag.ReadProperty("FontSize", 8)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

On Error Resume Next
PropBag.WriteProperty "CalenderDate", lDate, Date
PropBag.WriteProperty "Selector", lDate, Date
PropBag.WriteProperty "FontSize", picCalender.FontSize, 8

End Sub

Private Sub UserControl_Resize()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "Calender.Resize", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

lblHeader.Width = UserControl.ScaleWidth
picCalender.Width = UserControl.ScaleWidth
picCalender.Height = UserControl.ScaleHeight - lblHeader.Height

End Sub

Public Sub Redraw()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "Calender.Redraw", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, MoonVal() As Long, ExtFlag() As Long, TmpDate As New clsDate, Y#, Tmp$, BH#, Rect As Rect, ItemCount&

TmpDate.Contents = lDate
ItemCount = TmpDate.GetMonthLenght(TmpDate.cMonth)

ReDim TextArray(1 To ItemCount)
ReDim MoonVal(1 To ItemCount)
ReDim ExtFlag(1 To ItemCount)
RaiseEvent Redrawing(ItemCount, TextArray, MoonVal, ExtFlag)

BH = (picCalender.Height - 7) / ItemCount
picCalender.Cls

For Tell = 1 To ItemCount
    TmpDate.cDay = Tell
    Tmp = UCase(Mid(TmpDate.cDayName, 1, 1))
    
    If TmpDate.cWeekDay = vbSunday Or InStr(1, TextArray(Tell), "(!CH)") > 0 Then picCalender.Line (0, (Tell - 1) * BH + 1)-(picCalender.Width - 13, Tell * BH), RGB(0, 255, 255), BF
    If InStr(1, TextArray(Tell), "(!CF)") > 0 Then picCalender.PaintPicture imgFlag.Picture, picCalender.Width - 73, ((Tell - 1) * BH) + (BH / 2) - (imgFlag.Height / 2)
    
    picCalender.Line (0, Tell * BH)-(picCalender.Width, Tell * BH), &HC0C0C0
    picCalender.Line (40, (Tell - 1) * BH)-(40, Tell * BH), &HC0C0C0
    picCalender.Line (picCalender.Width - 12, (Tell - 1) * BH)-(picCalender.Width - 12, Tell * BH), &HC0C0C0
    picCalender.Line (picCalender.Width - 7, (Tell - 1) * BH)-(picCalender.Width - 7, Tell * BH), &HC0C0C0
    
    If (ExtFlag(Tell) And 1) = 1 Then
        picCalender.Line (picCalender.Width - 11, (Tell - 1) * BH + 1)-(picCalender.Width - 8, Tell * BH - 1), &HFF00&, BF
    End If
    
    If (ExtFlag(Tell) And 2) = 2 Then
        picCalender.Line (picCalender.Width - 6, (Tell - 1) * BH + 1)-(picCalender.Width, Tell * BH - 1), &H8000&, BF
    End If
    
    picCalender.ForeColor = vbBlack
    
    If Tell = Day(lDate) Then
    
        With Rect
        .Left = 0
        .Top = (Tell - 1) * BH + 1
        .Right = 40
        .Bottom = Tell * BH
        End With
        
        picCalender.Line (0, Rect.Top - 1)-(Rect.Right - 1, Rect.Bottom - 1), vbBlue, BF
        If GetFocus = picCalender.hWnd And lSelector Then DrawFocusRect picCalender.hdc, Rect
        picCalender.ForeColor = vbWhite
    End If

    Y = ((Tell - 1) * BH) + (BH / 2) - (picCalender.textHeight("A") / 2)

    picCalender.CurrentX = 1
    picCalender.CurrentY = Y
    picCalender.Print Tmp & Space(1) & TmpDate.cDay
    
    picCalender.ForeColor = vbBlack
    picCalender.FontName = "Times New Roman"
    Script.DrawTextEx picCalender, TextArray(Tell), 46, Y
    
    picCalender.FontName = "Courier New"
    
    If Tmp = "M" Then
        Tmp = TmpDate.cWeekNum
        picCalender.CurrentX = UserControl.ScaleWidth - UserControl.TextWidth(Tmp) - 28
        picCalender.CurrentY = Y
        picCalender.Print Tmp
    End If

    If MoonVal(Tell) Mod 90 < 15 And MoonVal(Script.NotOver(Tell - 1, Flag_NotUnderOne)) <> 0 Then
        TransparentBlt picCalender.hdc, UserControl.ScaleWidth - 37, Y + 3, imgMoon(0).Width, imgMoon(0).Height, imgMoon(CInt(MoonVal(Tell) / 90)).hdc, 0, 0, imgMoon(0).Width, imgMoon(0).Height, RGB(255, 0, 128)
        MoonVal(Tell) = 0
    End If
Next

End Sub

Public Property Get Text() As String
Text = TextArray(Day(lDate))
End Property

Public Property Get Header() As String
Header = lblHeader.Caption
End Property

Public Property Let Header(ByVal vNewValue As String)
lblHeader.Caption = vNewValue
End Property

Public Property Get CalenderDate() As Date
CalenderDate = lDate
End Property

Public Property Let CalenderDate(ByVal vNewValue As Date)

Dim TmpDate As Date

TmpDate = lDate

If lDate <> vNewValue Then
    RaiseEvent DateChanging
    lDate = vNewValue
    RaiseEvent DateChanged(TmpDate)
End If

Redraw

End Property

Public Property Get Selector() As Boolean
Selector = lSelector
End Property

Public Property Let Selector(ByVal vNewValue As Boolean)

lSelector = vNewValue
Redraw

End Property

Private Sub picCalender_GotFocus()

Redraw

End Sub

Private Sub picCalender_LostFocus()

Redraw

End Sub

Private Sub picCalender_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "Calender.picCalender_MouseDown(Button, Shift, X, Y)", Array(Button, Shift, X, Y), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim BH&, TmpDate As New clsDate, StopParent As Boolean, CheckDate As Date

RaiseEvent MouseDown(Button, Shift, X, Y)

TmpDate.Contents = lDate

BH = Val(Y / ((picCalender.Height - 7) / TmpDate.GetMonthLenght(TmpDate.cMonth)))
CheckDate = DateSerial(Year(lDate), Month(lDate), BH + 1)

If X < 39 And Not CheckDate = TmpDate.Contents Then
    CalenderDate = CheckDate
End If

End Sub

Public Property Get FontSize() As Long
    FontSize = picCalender.FontSize
End Property

Public Property Let FontSize(ByVal vNewValue As Long)
    picCalender.FontSize = vNewValue
End Property
