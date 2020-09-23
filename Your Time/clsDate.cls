VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private CurrDate As Date

' Returnes a number between 0-360, which indicates the current moon phase.
Public Property Get MoonPhase() As Integer
    MoonPhase = Script.GetMoonPhase(CurrDate)
End Property

Public Property Get DaylightSaving() As Boolean
    DaylightSaving = CBool(CurrDate >= LastWeekday(vbSunday, 3) And CurrDate <= LastWeekday(vbSunday, 10))
End Property

Public Property Get cDay() As Long
    cDay = Day(CurrDate)
End Property

Public Property Let cDay(ByVal vNewValue As Long)
    CurrDate = DateSerial(Year(CurrDate), Month(CurrDate), vNewValue)
End Property

Public Property Get cYear() As Long
    cYear = Year(CurrDate)
End Property

Public Property Let cYear(ByVal vNewValue As Long)
    CurrDate = DateSerial(vNewValue, Month(CurrDate), Day(CurrDate))
End Property

Public Property Get cMonth() As Long
    cMonth = Month(CurrDate)
End Property

Public Property Let cMonth(ByVal vNewValue As Long)
    CurrDate = DateSerial(Year(CurrDate), vNewValue, Day(CurrDate))
End Property

Public Property Get cMonthName() As String
    cMonthName = MonthName(Month(CurrDate))
End Property

Public Property Get cDayName() As String
    cDayName = WeekdayName(cWeekDay, , vbUseSystemDayOfWeek)
End Property

Public Property Get cWeekDay() As Long
    cWeekDay = Weekday(CurrDate, vbUseSystemDayOfWeek)
End Property

Public Property Get cWeekNum() As Long
    cWeekNum = DateDiff("ww", DateSerial(cYear, 1, 1), CurrDate) + 1
End Property

Public Property Get cTotalDays() As Long
    cTotalDays = DateDiff("s", DateSerial(cYear, 1, 1), CurrDate) / 86400
End Property

Public Property Get cLeapYear() As Long
    If (cYear Mod 4 = 0) And (cYear Mod 100 <> 0) Or (cYear Mod 400 = 0) Then
        cLeapYear = 29
    Else
        cLeapYear = 28
    End If
End Property

Public Property Get Contents() As Date
    Contents = CurrDate
End Property

Public Property Let Contents(ByVal vNewValue As Date)
    CurrDate = vNewValue
End Property

Public Function Clone() As clsDate

Set Clone = New clsDate

Clone.Contents = Me.Contents

End Function

Private Sub Class_Initialize()
    CurrDate = Date
End Sub

Public Function Add(Interval As String, Number As Double)

CurrDate = DateAdd(Interval, Number, CurrDate)

End Function

Public Function GetMonthLenght(Month As Long) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsDate.GetMontLenght(Month)", Array(Month), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tmp As Long

Select Case Month
Case 1: Tmp = 31
Case 2: Tmp = cLeapYear
Case 3: Tmp = 31
Case 4: Tmp = 30
Case 5: Tmp = 31
Case 6: Tmp = 30
Case 7: Tmp = 31
Case 8: Tmp = 31
Case 9: Tmp = 30
Case 10: Tmp = 31
Case 11: Tmp = 30
Case 12: Tmp = 31
End Select

GetMonthLenght = Tmp

End Function

Public Function LastWeekday(lWeekday As Long, lMonth As Long, Optional dayNum As Long = 1) As Date

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsDate.LastWeekday(lWeekday, lMonth, [dayNum])", Array(lWeekday, lMonth, dayNum), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim lDate As Date, Tell&

For Tell = GetMonthLenght(lMonth) To 1 Step -1

    lDate = DateSerial(cYear, lMonth, Tell)

    If Weekday(lDate) = lWeekday Then
        
        dayNum = dayNum - 1
        
        If dayNum <= 0 Then
            LastWeekday = lDate
            Exit Function
        End If
        
    End If

Next

End Function

Public Function FirstWeekday(lWeekday As Long, lMonth As Long, Optional dayNum As Long = 1) As Date

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsDate.FirstWeekday(lWeekday, lMonth, [dayNum])", Array(lWeekday, lMonth, dayNum), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim lDate As Date, Tell&

For Tell = 1 To GetMonthLenght(lMonth)

    lDate = DateSerial(cYear, lMonth, Tell)

    If Weekday(lDate) = lWeekday Then
    
        dayNum = dayNum - 1
    
        If dayNum <= 0 Then
            FirstWeekday = lDate
            Exit Function
        End If
        
    End If

Next

End Function

Public Function WeekdaysFromDate(ByVal lDate As Date, lWeekday As Long, lDirection As Long, dayNum As Long) As Date

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsDate.WeekdaysFromDate(lDate, lWeekday, lDirection, dayNum)", Array(lDate, lWeekday, lDirection, dayNum), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

If lDirection = 0 Then
    WeekdaysFromDate = lDate
    Exit Function
End If

Do
   lDate = lDate + lDirection

    If Weekday(lDate) = lWeekday Then
    
        dayNum = dayNum - 1
        
        If dayNum <= 0 Then
            WeekdaysFromDate = lDate
            Exit Function
        End If

    End If
Loop

End Function

Public Property Get EasterDay() As Date

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsDate.EasterDay", , EA_NORERAISE: HandleError: Exit Property
End If
' *** BEGIN CODE ***

    Dim FirstDig, Remain19, temp
    Dim tA, tB, tC, tD, tE, d, m
    
    FirstDig = cYear \ 100
    Remain19 = cYear Mod 19
    
    temp = (FirstDig - 15) \ 2 + 202 - 11 * Remain19
    
    Select Case FirstDig
    Case 21, 24, 25, 27 To 32, 34, 35, 38
        temp = temp - 1
    Case 33, 36, 37, 39, 40
        temp = temp - 2
    End Select
    
    temp = temp Mod 30
    
    tA = temp + 21
    If temp = 29 Then tA = tA - 1
    If (temp = 28 And Remain19 > 10) Then tA = tA - 1
    
    ' Find the next Sunday
    tB = (tA - 19) Mod 7
    
    tC = (40 - FirstDig) Mod 4
    If tC = 3 Then tC = tC + 1
    If tC > 1 Then tC = tC + 1
    
    temp = cYear Mod 100
    tD = (temp + temp \ 4) Mod 7
    
    tE = ((20 - tB - tC - tD) Mod 7) + 1
    d = tA + tE
    
    If d > 61 Then
        d = d - 61
        m = 5
    ElseIf d > 31 Then
        d = d - 31
        m = 4
    Else
        m = 3
    End If

    EasterDay = DateSerial(cYear, m, d)
    
End Property

Public Property Get Name() As String
    Name = "Date"
End Property
