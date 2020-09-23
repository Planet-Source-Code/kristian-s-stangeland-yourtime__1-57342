VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSunTime"
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

Const pi = 3.14159265358979
Const RADEG = 180 / pi
Const DEGRAD = pi / 180

Dim lYear As Long
Dim lMonth As Long
Dim lDay As Long

Dim lSunSet As Date
Dim lSunRise As Date
Dim lChanged As Boolean
Dim lLongitude As Double
Dim lLatitude As Double
Dim lTZ As Double
Dim lIsDst As Double
Dim TmpDate As New clsDate

Public Property Get Name() As String
    Name = "SunTime"
End Property

Public Property Get SunSet() As Date
    TmpDate.Contents = DateSerial(lYear, lMonth, lDay)
    SunSet = lSunSet + IIf(Sommertid = 1, IIf(TmpDate.DaylightSaving, TimeSerial(1, 0, 0), 0), 0)
End Property

Public Property Get SunRise() As Date
    TmpDate.Contents = DateSerial(lYear, lMonth, lDay)
    SunRise = lSunRise + IIf(Sommertid = 1, IIf(TmpDate.DaylightSaving, TimeSerial(1, 0, 0), 0), 0)
End Property

Public Property Get Changed() As Long
    Changed = lChanged
End Property

Public Property Get cYear() As Long
    cYear = lYear
End Property

Public Property Let cYear(ByVal vNewValue As Long)
    lChanged = lChanged Or CBool(lYear <> vNewValue)
    lYear = vNewValue
End Property

Public Property Get cMonth() As Long
    cMonth = lMonth
End Property

Public Property Let cMonth(ByVal vNewValue As Long)
    lChanged = lChanged Or CBool(lMonth <> vNewValue)
    lMonth = vNewValue
End Property

Public Property Get cDay() As Long
    cDay = lDay
End Property

Public Property Let cDay(ByVal vNewValue As Long)
    lChanged = lChanged Or CBool(lDay <> vNewValue)
    lDay = vNewValue
End Property

Public Property Get cCurrentDate() As Date
    cCurrentDate = DateSerial(lYear, lMonth, lDay)
End Property

Public Property Let cCurrentDate(ByVal vNewValue As Date)
    
    cYear = Year(vNewValue)
    cMonth = Month(vNewValue)
    cDay = Day(vNewValue)
    
End Property

Public Property Get Longitude() As Double
    Longitude = lLongitude
End Property

Public Property Let Longitude(ByVal vNewValue As Double)
    lChanged = lChanged Or CBool(lLongitude <> vNewValue)
    lLongitude = vNewValue
End Property

Public Property Get Latitude() As Double
    Latitude = lLatitude
End Property

Public Property Let Latitude(ByVal vNewValue As Double)
    lChanged = lChanged Or CBool(lLatitude <> vNewValue)
    lLatitude = vNewValue
End Property

Public Property Get TimeZone() As Double
    TimeZone = lTZ
End Property

Public Property Let TimeZone(ByVal vNewValue As Double)
    lChanged = lChanged Or CBool(lTZ <> vNewValue)
    lTZ = vNewValue
End Property

Public Property Get Sommertid() As Double
    Sommertid = lIsDst
End Property

Public Property Let Sommertid(ByVal vNewValue As Double)
    lChanged = lChanged Or CBool(lIsDst <> vNewValue)
    lIsDst = vNewValue
End Property

Public Function Clone() As clsSunTime

Set Clone = New clsSunTime

Clone.cCurrentDate = Me.cCurrentDate
Clone.Sommertid = Me.Sommertid
Clone.TimeZone = Me.TimeZone
Clone.Latitude = Me.Latitude
Clone.Longitude = Me.Longitude

End Function

Public Sub Calculate()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsSunTime.Calculate", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim d#, n#, i#, w#, m#, L#, E#, e1#, a#, xv#, yv#, V#, XS#, YS#, xe#, ecl#, lonsun#, ye#, ze#
Dim GMST0#, UT_Sun_in_south#, LHA#, r#, ra#, dec#, h#, hour_rise#, hour_set#, min_rise#, min_set#

d = (367 * (lYear) - Int((7 * ((lYear) + (((lMonth) + 9) / 12))) / 4) + Int((275 * (lMonth)) / 9) + (lDay) - 730530)

' Orbital elements of the Sun:
n = 0#
i = 0#
w = 282.9404 + 0.0000470935 * d

a = 1#
E = 0.016709 - 0.000000001151 * d

m = 356.047 + 0.9856002585 * d
m = Rev(m)

ecl = 23.4393 - 0.0000003563 * d
L = w + m

If (L < 0 Or L > 360) Then
    L = Rev(L)
End If

' position of the Sun
e1 = m + E * (180 / pi) * Sind(m) * (1# + E * Cosd(m))
xv = Cosd(e1) - E
yv = Sqrt(1# - E * E) * Sind(e1)

V = Atan2d(yv, xv)
r = Sqrt(xv * xv + yv * yv)
lonsun = V + w

If (lonsun < 0 Or lonsun > 360) Then
    lonsun = Rev(lonsun)
End If

XS = r * Cosd(lonsun)
YS = r * Sind(lonsun)
xe = XS
ye = YS * Cosd(ecl)
ze = YS * Sind(ecl)
ra = Atan2d(ye, xe)
dec = Atan2d(ze, (Sqrt((xe * xe) + (ye * ye))))
h = -0.833

GMST0 = L + 180

If (GMST0 < 0 Or GMST0 > 360) Then
    GMST0 = Rev(GMST0)
End If

UT_Sun_in_south = (ra - GMST0 - lLongitude) / 15#

If (UT_Sun_in_south < 0) Then
    UT_Sun_in_south = UT_Sun_in_south + 24
End If

LHA = (Sind(h) - (Sind(lLatitude) * Sind(dec))) / (Cosd(lLatitude) * Cosd(dec))
LHA = Acosd(LHA) / 15

hour_rise = UT_Sun_in_south - LHA
hour_set = UT_Sun_in_south + LHA
min_rise = Int((hour_rise - Int(hour_rise)) * 60)
min_set = Int((hour_set - Int(hour_set)) * 60)

hour_rise = (Int(hour_rise) + (lTZ + lIsDst))
hour_set = (Int(hour_set) + (lTZ + lIsDst))

lChanged = False
lSunRise = TimeSerial(hour_rise, min_rise, 0)
lSunSet = TimeSerial(hour_set, min_set, 0)

End Sub

' Support Functions
Private Function Sind(qqq As Double)

Sind = Sin((qqq) * DEGRAD)

End Function

Private Function Cosd(qqq As Double)

Cosd = Cos((qqq) * DEGRAD)

End Function

Private Function Tand(qqq As Double)

Tand = Tan((qqq) * DEGRAD)

End Function

Private Function Acosd(qqq As Double)

Acosd = (RADEG * Acos(qqq))

End Function

Private Function Atan2d(qqq As Double, qqq1 As Double)

Atan2d = (RADEG * Atan2(qqq, qqq1))

End Function

Private Function Rev(qqq As Double)

Dim X As Double

X = (qqq - Int(qqq / 360#) * 360#)

If (X <= 0) Then
    X = X + 360
End If

Rev = X
 
End Function

Private Function Atan2(YS As Double, XS As Double)

Dim Theta As Double

If XS <> 0 Then

    Theta = Atn(YS / XS)
    
    If XS < 0 Then
        Theta = Theta + pi
    End If
Else

    If YS < 0 Then
        Theta = 3 * pi / 2 '90
    Else
        Theta = pi / 2 '270
    End If
    
End If

Atan2 = Theta

End Function

Private Function Acos(X As Double)

Acos = Atn(-X / Sqrt(-X * X + 1)) + 2 * Atn(1)
 
End Function

Private Function Sqrt(X As Double)

If X > 0 Then
    Sqrt = Sqr(X)
Else
    Sqrt = 0
End If

End Function

Private Sub Class_Initialize()

cCurrentDate = Date

End Sub
