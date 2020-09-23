VERSION 5.00
Begin VB.UserControl YearReview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7080
   ScaleHeight     =   223
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
End
Attribute VB_Name = "YearReview"
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

Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOver(BlockDate As Date, Shift As Integer)
Event MouseDown(BlockDate As Date, Button As Integer, Shift As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Redrawing(BlockDescription() As String, BlockBackcolor() As Long)
Event Redrawed()

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Public Sub Redraw()

Dim BlockDescription(30, 11) As String, BlockBackcolor(30, 11) As Long
Dim lineX As Long, lineY As Long, Text As String

' Get input
RaiseEvent Redrawing(BlockDescription, BlockBackcolor)

' The first line is special
UserControl.Line (20, 0)-(20, UserControl.ScaleHeight), vbButtonShadow

' Draw all the X-lines and description
For lineX = 1 To 31
    UserControl.Line (((lineX - 1) * 14) + 35, 0)-(((lineX - 1) * 14) + 35, UserControl.ScaleHeight), vbButtonShadow
    UserControl.CurrentX = ((lineX - 1) * 14) + 27.5 - (UserControl.TextWidth(CStr(lineX)) / 2)
    UserControl.CurrentY = 3
    UserControl.Print CStr(lineX)
Next

' Draw all the Y-lines and description
For lineY = 0 To 11
    Text = CapitalizeFirstLetter(Mid$(MonthName(lineY + 1), 1, 3))

    UserControl.Line (0, 15 + (lineY * 14))-(UserControl.ScaleWidth, 15 + (lineY * 14)), vbButtonShadow
    UserControl.CurrentX = 9.5 - UserControl.TextWidth(Text) / 2
    UserControl.CurrentY = 23.5 + (lineY * 14) - (UserControl.TextWidth(Text) / 2)
    UserControl.Print Text
Next

' Draw all the boxes
For lineX = 0 To 30
    For lineY = 0 To 11
        UserControl.Line (((lineX - 1) * 14) + 36, 16 + (lineY * 14))-((lineX * 14) + 34, 14 + ((lineY + 1) * 14)), BlockBackcolor(lineX, lineY), BF
        UserControl.CurrentX = ((lineX - 1) * 14) + 34 + (UserControl.TextWidth(BlockDescription(lineX, lineY)) / 2)
        UserControl.CurrentY = 22.5 + (lineY * 14) - (UserControl.textHeight(BlockDescription(lineX, lineY)) / 2)
        UserControl.Print BlockDescription(lineX, lineY)
    Next
Next

' Finish
RaiseEvent Redrawed

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ret

ret = GetDateFromPosition(X, Y)

If IsDate(ret) Then
    RaiseEvent MouseDown(CDate(ret), Button, Shift)
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ret

ret = GetDateFromPosition(X, Y)

If IsDate(ret) Then
    RaiseEvent MouseOver(CDate(ret), Shift)
Else
    RaiseEvent MouseMove(Button, Shift, X, Y)
End If

End Sub

Public Function GetDateFromPosition(X As Single, Y As Single) As Variant

Dim lineX&, lineY&

lineX = ((X - 20) \ 14)
lineY = ((Y - 16) \ 14)

If Not Script.IsDateSerial(CurrentDate.cYear, lineY + 1, lineX + 1) Then
    'GetDateFromPosition -1
    Exit Function
End If

GetDateFromPosition = DateSerial(CurrentDate.cYear, lineY + 1, lineX + 1)

End Function
