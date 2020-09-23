VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
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

Private Type GUIDs
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As Long, rclsid As GUIDs) As Long
Private Declare Function CoDisconnectObject Lib "ole32.dll" (ByVal pUnk As IUnknown, pvReserved As Long) As Long
Private Declare Function RegisterActiveObject Lib "oleaut32.dll" (ByVal pUnk As IUnknown, rclsid As GUIDs, ByVal dwFlags As Long, pdwRegister As Long) As Long
Private Declare Function RevokeActiveObject Lib "oleaut32.dll" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long

'Declares needed to register object in the ROT
Private Const ACTIVEOBJECT_STRONG = 0
Private Const ACTIVEOBJECT_WEAK = 1

Enum ScriptLang
    Lang_Simple = 0
    Lang_VB = 1
    Lang_JS = 2
End Enum

Enum Flag_NotOver
    Flag_NotNegative = 1
    Flag_NotPositive = 2
    Flag_NotUnderOne = 3
    Flag_Procent = 4
End Enum

Enum InvokeKinds
    INVOKE_UNKNOWN = 0
    INVOKE_FUNC = 1
    INVOKE_PROPERTYGET = 2
    INVOKE_PROPERTYPUT = 4
    INVOKE_PROPERTYPUTREF = 8
    INVOKE_EVENTFUNC = 16
    INVOKE_CONST = 32
End Enum

Dim OLEInstance As Long
Dim TLIApp As Object
Dim objScriptEn As Object
Dim strCommandLine As String
Dim lLanguage As ScriptLang
Dim bRunning As Boolean
Dim bIsIDE As Boolean
Dim lTag As Variant

Public Property Get WindowsVersion() As String

    Dim OSInfo As OSVERSIONINFO, PId As String, ret As Long

    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    ret = GetVersionEx(OSInfo)
    
    If ret = 0 Then
        Exit Property
    End If
    
    Select Case OSInfo.dwPlatformId
        Case 0
            PId = "Windows 32s "
        Case 1
            PId = "Windows 95/98"
        Case 2
            PId = "Windows NT "
    End Select

    WindowsVersion = "Microsoft " & PId & "(" & Str$(OSInfo.dwMajorVersion) & "." + LTrim(Str(OSInfo.dwMinorVersion)) & Str(OSInfo.dwBuildNumber) & ")"

End Property

Public Property Get GlobalPeoplesCount() As Long
    GlobalPeoplesCount = SafeUBound(VarPtrArray(Peoples))
End Property

Public Property Get PeoplesData(ByVal Index, ByVal Variable) As Variant

    Select Case LCase(Variable)
    Case "address": PeoplesData = Peoples(Index).Address
    Case "birthday": PeoplesData = Peoples(Index).Birthday
    Case "country": PeoplesData = Peoples(Index).Country
    Case "email": PeoplesData = Peoples(Index).Email
    Case "enabled": PeoplesData = Peoples(Index).Enabled
    Case "fax": PeoplesData = Peoples(Index).Fax
    Case "firm": PeoplesData = Peoples(Index).Firm
    Case "firmnum": PeoplesData = Peoples(Index).FirmNum
    Case "homepage": PeoplesData = Peoples(Index).Homepage
    Case "information": PeoplesData = Peoples(Index).Information
    Case "mobnum": PeoplesData = Peoples(Index).MobNum
    Case "name": PeoplesData = Peoples(Index).Name
    Case "phonenum": PeoplesData = Peoples(Index).PhoneNum
    Case "postcity": PeoplesData = Peoples(Index).PostCity
    Case "visiblenum": PeoplesData = Peoples(Index).VisibleNum
    End Select

End Property

Public Property Let PeoplesData(ByVal Index, ByVal Variable, ByVal vNewValue)

    Select Case LCase(Variable)
    Case "address": Peoples(Index).Address = vNewValue
    Case "birthday":  Peoples(Index).Birthday = vNewValue
    Case "country": Peoples(Index).Country = vNewValue
    Case "email":  Peoples(Index).Email = vNewValue
    Case "enabled":  Peoples(Index).Enabled = vNewValue
    Case "fax": Peoples(Index).Fax = vNewValue = vNewValue
    Case "firm":  Peoples(Index).Firm = vNewValue
    Case "firmnum":  Peoples(Index).FirmNum = vNewValue
    Case "homepage":  Peoples(Index).Homepage = vNewValue
    Case "information":  Peoples(Index).Information = vNewValue
    Case "mobnum":  Peoples(Index).MobNum = vNewValue
    Case "name":  Peoples(Index).Name = vNewValue
    Case "phonenum":  Peoples(Index).PhoneNum = vNewValue
    Case "postcity":  Peoples(Index).PostCity = vNewValue
    Case "visiblenum":  Peoples(Index).VisibleNum = vNewValue
    End Select
    
End Property

Public Property Get UserVariables(ByVal Variable) As Variant

    Select Case LCase(Variable)
    Case "changed": UserVariables = Users(UserID).Changed
    Case "created": UserVariables = Users(UserID).Created
    Case "lastpath": UserVariables = Users(UserID).LastPath
    Case "loggedon": UserVariables = Users(UserID).LoggedOn
    Case "password": UserVariables = Users(UserID).Password
    Case "username": UserVariables = Users(UserID).UserName
    End Select

End Property

Public Property Let UserVariables(ByVal Variable, ByVal vNewValue)
    
    Select Case LCase(Variable)
    Case "changed": Users(UserID).Changed = vNewValue
    Case "created": Users(UserID).Created = vNewValue
    Case "lastpath": Users(UserID).LastPath = vNewValue
    Case "loggedon": Users(UserID).LoggedOn = vNewValue
    Case "password": Users(UserID).Password = vNewValue
    Case "username": Users(UserID).UserName = vNewValue
    End Select
    
End Property

Public Property Get LanguagePack() As String
    LanguagePack = GetSetting("YourTime", "Global", "LanguagePack", "")
End Property

Public Property Let LanguagePack(vNewValue As String)
    SaveSetting "YourTime", "Global", "LanguagePack", vNewValue
End Property

Public Property Get CurrentUser() As Long
    CurrentUser = UserID
End Property

Public Property Let CurrentUser(ByVal vNewValue As Long)
    UserID = vNewValue
End Property

Public Property Get Name() As String
    Name = "Script"
End Property

Public Property Get Running() As Boolean
    Running = bRunning
End Property

Public Property Let Running(ByVal vNewValue As Boolean)
    bRunning = vNewValue
End Property

Public Property Get FindPlugins() As Boolean
    FindPlugins = GetSetting("YourTime", "Global", "FindPlugins", True)
End Property

Public Property Let FindPlugins(vNewValue As Boolean)
    SaveSetting "YourTime", "Global", "FindPlugins", vNewValue
End Property

Public Property Get EnableTags(Index As String) As Long
    EnableTags = GetSetting("YourTime", "Global", "Tag" & Index, 1)
End Property

Public Property Let EnableTags(Index As String, vNewValue As Long)
    SaveSetting "YourTime", "Global", "Tag" & Index, vNewValue
End Property

Public Property Get CGlobalObjects() As Collection
    Set CGlobalObjects = GlobalObjects
End Property

Public Property Get CPlugins() As Collection
    Set CPlugins = Plugins
End Property

Public Property Get CForms() As Object
    Set CForms = Forms
End Property

Public Property Get LagreFont() As Boolean
    LagreFont = GetSetting("YourTime", "Global", "LagreFont", True)
End Property

Public Property Let LagreFont(ByVal vNewValue As Boolean)
    SaveSetting "YourTime", "Global", "LagreFont", vNewValue
End Property

Public Property Get RegScript() As String
    RegScript = GetSetting("YourTime", "Global", "RegScript", "")
End Property

Public Property Let RegScript(ByVal vNewValue As String)
    SaveSetting "YourTime", "Global", "RegScript", vNewValue
End Property

Public Property Get RunScript() As Boolean
    RunScript = GetSetting("YourTime", "Global", "RunScript", True)
End Property

Public Property Let RunScript(ByVal vNewValue As Boolean)
    SaveSetting "YourTime", "Global", "RunScript", vNewValue
End Property

Public Property Get TransparentKey() As Long
    TransparentKey = GetSetting("YourTime", "Global", "TransparentKey", 180)
End Property

Public Property Let TransparentKey(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Global", "TransparentKey", vNewValue
End Property

Public Property Get EnableTransparent() As Boolean
    EnableTransparent = GetSetting("YourTime", "Global", "EnableTransparent", True)
End Property

Public Property Let EnableTransparent(ByVal vNewValue As Boolean)
    SaveSetting "YourTime", "Global", "EnableTransparent", vNewValue
End Property

Public Property Get TransparentDefault() As Boolean
    TransparentDefault = GetSetting("YourTime", "Global", "TransparentDefault", False)
End Property

Public Property Let TransparentDefault(ByVal vNewValue As Boolean)
    SaveSetting "YourTime", "Global", "TransparentDefault", vNewValue
End Property

Public Property Get LogToEventLog() As Boolean
    LogToEventLog = GetSetting("YourTime", "Global", "LogToEventLog", False)
End Property

Public Property Let LogToEventLog(ByVal vNewValue As Boolean)
    SaveSetting "YourTime", "Global", "LogToEventLog", vNewValue
End Property

Public Property Get DefaultLanguage() As Long
    DefaultLanguage = GetSetting("YourTime", "Global", "DefaultLanguage", Lang_VB)
End Property

Public Property Let DefaultLanguage(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Global", "DefaultLanguage", vNewValue
End Property

Public Property Get AnimationSpeed() As Long
    AnimationSpeed = GetSetting("YourTime", "Constants", "AnimationSpeed", 10)
End Property

Public Property Let AnimationSpeed(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Constants", "AnimationSpeed", vNewValue
End Property

Public Property Get StarAmout() As Long
    StarAmout = GetSetting("YourTime", "Constants", "StarAmout", 300)
End Property

Public Property Let StarAmout(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Constants", "StarAmout", vNewValue
End Property

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get ShowMoon() As Long
    ShowMoon = GetSetting("YourTime", "Global", "ShowMoon", 1)
End Property

Public Property Let ShowMoon(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Global", "ShowMoon", vNewValue
End Property

Public Property Get ShowDayLenght() As Long
    ShowDayLenght = GetSetting("YourTime", "Global", "ShowDayLenght", 1)
End Property

Public Property Let ShowDayLenght(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Global", "ShowDayLenght", vNewValue
End Property

Public Property Get ShowSunUpDown() As Long
    ShowSunUpDown = GetSetting("YourTime", "Global", "ShowSunUpDown", 1)
End Property

Public Property Let ShowSunUpDown(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Global", "ShowSunUpDown", vNewValue
End Property

Public Property Get ShowFlag() As Long
    ShowFlag = GetSetting("YourTime", "Global", "ShowFlag", 1)
End Property

Public Property Let ShowFlag(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Global", "ShowFlag", vNewValue
End Property

Public Property Get ShowOwnDays() As Long
    ShowOwnDays = GetSetting("YourTime", "Global", "ShowOwnDays", 1)
End Property

Public Property Let ShowOwnDays(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Global", "ShowOwnDays", vNewValue
End Property

Public Property Get ShowHoliday() As Long
    ShowHoliday = GetSetting("YourTime", "Global", "ShowHoliday", 1)
End Property

Public Property Let ShowHoliday(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Global", "ShowHoliday", vNewValue
End Property

Public Property Get NumOfRecords() As Long
    NumOfRecords = GetSetting("YourTime", "Constants", "NumOfRecords", 150)
End Property

Public Property Let NumOfRecords(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Constants", "NumOfRecords", vNewValue
End Property

Public Property Get DefaultUser() As String
    DefaultUser = GetSetting("YourTime", "Constants", "DefaultUser", "Ukjent")
End Property

Public Property Let DefaultUser(ByVal vNewValue As String)
    SaveSetting "YourTime", "Constants", "DefaultUser", vNewValue
End Property

Public Property Get StartupUser() As Long
    StartupUser = GetSetting("YourTime", "Constants", "StartupUser", 1)
End Property

Public Property Let StartupUser(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Constants", "StartupUser", vNewValue
End Property

Public Property Get Tag() As Variant
    Tag = lTag
End Property

Public Property Let Tag(ByVal vNewValue As Variant)
    lTag = vNewValue
End Property

Public Property Get Language() As ScriptLang
    Language = lLanguage
End Property

Public Property Let Language(ByVal vNewValue As ScriptLang)
    lLanguage = vNewValue
End Property

Public Property Get Latitude() As Double
    Latitude = GetSetting("YourTime", "SunTime", "Latitude", 60.29)
End Property

Public Property Let Latitude(ByVal vNewValue As Double)
    SaveSetting "YourTime", "SunTime", "Latitude", vNewValue
End Property

Public Property Get Longitude() As Double
    Longitude = GetSetting("YourTime", "SunTime", "Longitude", 5.22)
End Property

Public Property Let Longitude(ByVal vNewValue As Double)
    SaveSetting "YourTime", "SunTime", "Longitude", vNewValue
End Property

Public Property Get Sommertid() As Long
    Sommertid = GetSetting("YourTime", "SunTime", "Sommertid", 1)
End Property

Public Property Let Sommertid(ByVal vNewValue As Long)
    SaveSetting "YourTime", "SunTime", "Sommertid", vNewValue
End Property

Public Property Get TimeZone() As Long
    TimeZone = GetSetting("YourTime", "SunTime", "TimeZone", 0)
End Property

Public Property Let TimeZone(ByVal vNewValue As Long)
    SaveSetting "YourTime", "SunTime", "TimeZone", vNewValue
End Property

Public Property Get MinWidth() As Long
    MinWidth = GetSetting("YourTime", "Constants", "MinWidth", 800)
End Property

Public Property Let MinWidth(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Constants", "MinWidth", vNewValue
End Property

Public Property Get MinHeight() As Long
    MinHeight = GetSetting("YourTime", "Constants", "MinHeight", 600)
End Property

Public Property Let MinHeight(ByVal vNewValue As Long)
    SaveSetting "YourTime", "Constants", "MinHeight", vNewValue
End Property

Public Property Get CommandLine() As String
    CommandLine = strCommandLine
End Property

Public Property Get ProcessCL() As Boolean
    ProcessCL = GetSetting("YourTime", "Constants", "ProcessCL", True)
End Property

Public Property Let ProcessCL(ByVal vNewValue As Boolean)
    SaveSetting "YourTime", "Constants", "ProcessCL", vNewValue
End Property

Public Property Get UseHooking() As Boolean
    UseHooking = GetSetting("YourTime", "Constants", "UseHooking", True)
End Property

Public Property Let UseHooking(ByVal vNewValue As Boolean)
    SaveSetting "YourTime", "Constants", "UseHooking", vNewValue
End Property

Public Property Get DumpError() As Boolean
    DumpError = GetSetting("YourTime", "Constants", "DumpError", True)
End Property

Public Property Let DumpError(ByVal vNewValue As Boolean)
    SaveSetting "YourTime", "Constants", "DumpError", vNewValue
End Property

Public Property Get UserPeopleCount() As Long
    UserPeopleCount = SafeUBound(VarPtrArray(Users(UserID).DataPeoples))
End Property

Public Property Get PeopleCount() As Long
    PeopleCount = SafeUBound(VarPtrArray(Peoples))
End Property

Public Property Get NoUsers() As Boolean
    
' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.NoUsers", , EA_NORERAISE: HandleError: Exit Property
End If
' *** BEGIN CODE ***
    
    Dim Tell&
    
    For Tell = LBound(Users) To UBound(Users)
        If Users(Tell).UserName <> "" Then
            NoUsers = False
            Exit Property
        End If
    Next
    
    NoUsers = True
    
End Property

Public Function Clone() As clsScript

Set Clone = New clsScript

Clone.Language = Me.Language
Clone.Tag = Me.Tag

End Function

Private Sub Class_Initialize()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.Class_Initialize", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Comm As String

Set TLIApp = CreateObject("TLI.TLIApplication")
Set objScriptEn = CreateObject("MSScriptControl.ScriptControl")

objScriptEn.Language = IIf(lLanguage = Lang_VB, "VBScript", "JScript")
lLanguage = DefaultLanguage

Comm = GetCommLine
strCommandLine = Mid$(Comm, InStrRev(Comm, Chr(34)) + 1)

End Sub

Public Function RegisterObject(lpObject As Object, strName As String) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.RegisterObject", , EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim mGUID As GUIDs
Dim lp As Long

' This code is responsible for creating the entry for "YourTime.Main" in the rot.
' Without this, ActiveX EXE can't use GetObject to access the main object of this app.

OLEInstance = 0
lp = CLSIDFromProgID(StrPtr(strName), mGUID)
    
If lp = 0 Then
    lp = RegisterActiveObject(lpObject, mGUID, ACTIVEOBJECT_WEAK, OLEInstance)
    RegisterObject = OLEInstance
End If

End Function

Private Function GetCommLine() As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.GetCommLine", , EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim RetStr&, SLen&, Buffer$

RetStr = GetCommandLine
SLen = lstrlen(RetStr)

If SLen > 0 Then
    GetCommLine = Space$(SLen)
    CopyMemory ByVal GetCommLine, ByVal RetStr, SLen
End If

End Function

Private Function InStrX(ByVal Begin As Integer, Str As Variant, Optional SearchFor As String) As Integer

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.InStrX(Begin, Str, [SearchFor]", Array(Begin, Str, SearchFor), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim a&, Buff$, OneChar$, DontLook%

If SearchFor = "" Then SearchFor = " "

For a = Begin To Len(Str)
    Buff = Mid(Str, a, Len(SearchFor))
    OneChar = Mid(Buff, 1, 1)
    If OneChar = "(" And DontLook >= 0 Then DontLook = DontLook + 1
    If OneChar = ")" And DontLook > 0 Then DontLook = DontLook - 1
    If OneChar = Chr(34) And DontLook < 1 Then DontLook = IIf(DontLook < 0, 0, -1)
    If DontLook = 0 And Buff = SearchFor Then InStrX = a: Exit Function
Next

End Function

Public Function SplitX(Text As String, Delimiter As String) As Variant

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.SplitX(Text, Delimiter)", Array(Text, Delimiter), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&, Last&, Arr() As String, Cnt&

If Text = "" Then Exit Function
Last = 1

Do Until Tell >= Len(Text)
    Tell = InStrX(Last, Text, Delimiter)
    If Tell = 0 Then Tell = Len(Text) + 1
    
    ReDim Preserve Arr(Cnt)
    Arr(Cnt) = Mid$(Text, Last, Tell - Last)
    Cnt = Cnt + 1
    
    Last = Tell + 1
Loop

SplitX = Arr

End Function

Public Function LoadDiary(Optional lpDate As Date, Optional RecordNum As Long) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.LoadDiary([lpDate], [RecordNum])", Array(lpDate, RecordNum), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim lNum As Long

If IsMissing(lpDate) Then
    lpDate = CurrentDate.Contents
End If

lNum = Search(Users(UserID).DataDiary, lpDate, 1)
If lNum < 0 Then Exit Function

RecordNum = lNum
LoadDiary = Users(UserID).DataDiary(lNum).Text

End Function

Public Function NotOver(ByVal Value As Long, ByVal Flag As Flag_NotOver) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.NotOver(Value, Flag)", Array(Value, Flag), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Select Case Flag
Case Flag_NotNegative
    If Value < 0 Then Value = 0

Case Flag_NotPositive
    If Value > 0 Then Value = 0

Case Flag_NotUnderOne
    If Value < 1 Then Value = 1

Case Flag_Procent
    If Value < 0 Then Value = 0
    If Value > 100 Then Value = 255
    
End Select

NotOver = Value

End Function

Public Sub DrawTextEx(PictureBox As Object, Text As String, ByVal X As Long, ByVal Y As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.DrawTextEx(PictureBox, Text, X, Y)", Array(PictureBox.Parent.Name & "." & PictureBox.Name, Text, X, Y), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, LastW&, Tmp$, NoCh&, Val(4) As Boolean

LastW = 1

' Cache the settings to optimize speed
Val(0) = (Script.EnableTags("B") = 1)
Val(1) = (Script.EnableTags("S") = 1)
Val(2) = (Script.EnableTags("U") = 1)
Val(3) = (Script.EnableTags("I") = 1)
Val(4) = (Script.EnableTags("C") = 1)

For Tell = 1 To Len(Text)

    If Mid$(Text, Tell, 1) = "(" Or Tell = Len(Text) Then

        PictureBox.CurrentX = X
        PictureBox.CurrentY = Y
        PictureBox.Print Mid$(Text, LastW, Tell - LastW + IIf(Tell = Len(Text), 1, 0))
        
        Tmp = Mid$(Text, Tell + 1, NotOver(InStr(Tell + 1, Text, ")") - Tell - 1, Flag_NotNegative))

        Select Case UCase(Tmp)
        Case "!CF", "!CH", "!YA"
            ' These are constants for the calender, don't include them in the drawing
        Case "!FB"
            If Val(0) Then PictureBox.FontBold = Not PictureBox.FontBold Else Tmp = " ": NoCh = 1
        Case "!FI"
            If Val(3) Then PictureBox.FontItalic = Not PictureBox.FontItalic Else Tmp = " ": NoCh = 1
        Case "!FU"
            If Val(2) Then PictureBox.FontUnderline = Not PictureBox.FontUnderline Else Tmp = " ": NoCh = 1
        Case "!FS"
            If Val(1) Then PictureBox.FontStrikethru = Not PictureBox.FontStrikethru Else Tmp = " ": NoCh = 1
        Case "!FR"
            ResetFont PictureBox.Font
            PictureBox.ForeColor = vbBlack
        Case Else
        
            If UCase$(Mid$(Tmp, 1, 3)) = "!FC" And Val(4) Then
                PictureBox.ForeColor = Script.ToColor(Mid$(Tmp, 4))
            Else
                Tmp = " "
                NoCh = 1
            End If
        
        End Select
        
        X = X + PictureBox.TextWidth(Mid$(Text, LastW, NotOver(Tell - LastW + IIf(Tell = Len(Text), 1, 0), Flag_NotNegative)))
        Tell = Tell + Len(Tmp)
        
        LastW = Tell + IIf(NoCh <> 1, 2, -1)
        NoCh = 0
    End If

Next

ResetFont PictureBox.Font
PictureBox.ForeColor = vbBlack

End Sub

Public Function ToColor(Expression As String) As Long

On Error Resume Next
Dim Tmp As Variant

Select Case LCase(Expression)
Case "white": ToColor = vbWhite
Case "blue": ToColor = vbBlue
Case "red": ToColor = vbRed
Case "green": ToColor = vbGreen
Case "yellow": ToColor = vbYellow
Case "pink": ToColor = vbMagenta
Case "black": ToColor = vbBlack
Case "cyan": ToColor = vbCyan
Case Else

    Select Case LCase$(Mid$(Expression, 1, 1))
    Case "h": ToColor = Val("&H" & Mid(Expression, 2))
    Case "d": ToColor = Val(Mid(Expression, 2))
    Case "r"
    
        Tmp = Split(Mid$(Expression, 2), ",")
        ToColor = RGB(Tmp(0), Tmp(1), Tmp(2))
    
    Case Else
    
        ToColor = Val(Expression)
    End Select
    
End Select

End Function

Public Function ToDate(Expression As String) As Date

On Error Resume Next
Dim Tmp As Variant

Tmp = Split(Expression)

If UBound(Tmp) > 1 Then
    ToDate = DateSerial(CInt(Val(Tmp(0))), CInt(Val(Tmp(1))), CInt(Val(Tmp(2))))
End If

End Function

Public Sub ResetFont(Font As IFontDisp)

On Error Resume Next

Font.Bold = False
Font.Italic = False
Font.Strikethrough = False
Font.Underline = False

End Sub

Public Function IsDateEqual(Date1, Date2) As Boolean

On Error Resume Next
IsDateEqual = CBool(DateSerial(Year(Date1), Month(Date1), Day(Date1)) = DateSerial(Year(Date2), Month(Date2), Day(Date2)))

End Function

Public Function IsTimeEqual(Date1, Date2) As Boolean

On Error Resume Next
IsTimeEqual = CBool(TimeSerial(Hour(Date1), Minute(Date1), Second(Date1)) = TimeSerial(Hour(Date2), Minute(Date2), Second(Date2)))

End Function

Public Function Run(strCode As String) As Variant

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
bRunning = False
errHandler: ErrorIn "clsScript.Run(strCode)", Array(strCode), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Lines, Line, Elements
Dim Paramenters, Par, Obj, cClone

If bRunning Then
    Set cClone = Me.Clone

    cClone.Run strCode

    Set cClone = Nothing
    Exit Function
End If

bRunning = True

Select Case lLanguage
Case Lang_Simple

    Lines = Split(strCode, vbNewLine)

    For Each Line In Lines
    
        Elements = Split(Line, " ", 2)
    
        If Mid(Elements(1), 1, 1) = "=" Then
            ' It's a property
            CallByName Me, Elements(0), VbLet, Replace(Mid(Elements(1), 2), Chr(34), " ")
        Else
            ' It's a sub/function
        
            Paramenters = SplitX(CStr(Elements(1)), ",")
            If Elements(1) = "" Then Exit Function
        
            For Each Par In Paramenters
                Par = Replace(Par, Chr(34), " ")
            Next
        
            CallByName Me, Elements(0), VbMethod, Paramenters
        End If
    Next

Case Lang_VB, Lang_JS
    
    objScriptEn.Language = IIf(lLanguage = Lang_VB, "VBScript", "JScript")
    objScriptEn.Reset
    
    For Each Obj In GetScriptObjects
        If InStr(1, strCode, Obj, vbTextCompare) <> 0 Then objScriptEn.AddObject Obj, ObjectPtr(CStr(Obj))
    Next
    
    objScriptEn.AddCode strCode

End Select

bRunning = False

End Function

Public Function LoadFile(Path As String) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.LoadFile(Path)", Array(Path), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Free&, Buff$

Free = FreeFile

Open Path For Binary As Free
    Buff = Space(LOF(Free))
    Get #Free, , Buff
Close Free

LoadFile = Buff

End Function

Public Function SimpleRegExp(ByVal Par1 As Variant, ByVal Par2 As Variant, Expression As String) As Boolean

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.SimpleRegExp(Par1, Par2, Expression)", Array(Par1, Par2, Expression), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&, Value As Boolean

Value = True
Par1 = CStr(Par1)
Par2 = CStr(Par2)

If LCase$(Mid$(Expression, 1, 7)) = "script;" Then
    lTag = False
    Script.Run "Const SavedData = " & Chr(34) & Par1 & Chr(34) & vbNewLine & "Const CompareData = " & Chr(34) & Par2 & Chr(34) & vbNewLine & "Main.Tag = " & Mid$(Expression, 8)
    SimpleRegExp = lTag
    Exit Function
End If

For Tell = 1 To Len(Par1)
    Select Case Mid$(Expression, Tell, 1)
    Case "-"
        Value = Value And CBool(Mid(Par1, Tell, 1) = Mid(Par2, Tell, 1))
        
    Case "+"
        Value = Value Or CBool(Mid(Par1, Tell, 1) = Mid(Par2, Tell, 1))
    
    Case "?"
        Value = Value Xor CBool(Mid(Par1, Tell, 1) = Mid(Par2, Tell, 1))
    
    Case "^"
        Value = Value And CBool(Mid(Par1, Tell, 1) <> Mid(Par2, Tell, 1))
        
    Case "*"
        ' Ignore
    End Select
Next

SimpleRegExp = Value

End Function

Public Sub SaveFile(Path As String, Data As String, Append As Boolean)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.SaveFile(Path, Data, Append)", Array(Path, Data, Append), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Free As Long

If Not Append And Dir(Path) <> "" Then
    Kill Path
End If

Free = FreeFile

Open Path For Append As Free
    Print #Free, Data
Close Free

End Sub

Public Function GetMoonPhase(lDate As Date) As Integer

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.GetMoonPhase(lDate)", Array(lDate), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim lD As Long
Dim dD As Double

lD = DateDiff("d", "01.01.2001", lDate)
dD = 0.20439731 + lD * 0.03386319269
lD = Int(dD)
dD = dD - lD

lD = 360 * dD
If lD < 0 Then lD = lD + 360

GetMoonPhase = CInt(lD / 2) * 2

End Function

Public Sub DeleteUserData(User As Long)

On Error Resume Next
Erase Users(User).DataDiary
Erase Users(User).DataPeoples
Erase Users(User).DataRem
Erase Users(User).DataTasks

ReDim Users(User).DataOwn(Script.NumOfRecords)

End Sub

Public Function CreateForm(Caption As String, Optional Width As Long, Optional Height As Long) As Object

Set CreateForm = New frmDummy

CreateForm.Caption = Caption
CreateForm.Width = IIf(Width <= 0, 400, Width)
CreateForm.Height = IIf(Height <= 0, 200, Height)

End Function

Public Function HexToString(Hex As String) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.HexToString(Hex)", Array(Hex), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&, Tmp$, Buff$

For Tell = 1 To Len(Hex) Step 3
    Tmp = Mid$(Hex, Tell, 3)
    
    If Mid$(Tmp, 1, 1) = "%" Then
        Buff = Buff & Chr(Val("&H" & Mid$(Tmp, 2, 2)))
    End If
Next

HexToString = Buff

End Function

Public Function StringToHex(Text As String) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.StringToHex(Text)", Array(Text), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&, Buff$

Buff = Space(Len(Text) * 3)

For Tell = 1 To Len(Text)
    Mid(Buff, (Tell - 1) * 3 + 1, 3) = "%" & Hex(Asc(Mid$(Text, Tell, 1)))
Next

StringToHex = Buff

End Function

Public Function AddControl(Class As String, Name As String, Parent As Object) As Object

On Error Resume Next
Set AddControl = Parent.Controls.Add(Class, Name)

End Function

Public Function AddMenu(strCaption As String) As Object

Dim Tell As Long

For Tell = 0 To frmMain.mnuPlugInMenu.count - 1

    If frmMain.mnuPlugInMenu(Tell).Enabled = False Then
    
        ' Yeah, I know you shouldn't use GoTo too much
        GoTo SetData
        
    End If

Next

Tell = frmMain.mnuPlugInMenu.count
Load frmMain.mnuPlugInMenu(Tell)

SetData:

frmMain.mnuPlugInMenu(Tell).Visible = True
frmMain.mnuPlugInMenu(Tell).Enabled = True
frmMain.mnuPlugInMenu(Tell).Caption = strCaption

Set AddMenu = frmMain.mnuPlugInMenu(Tell)


End Function

Public Sub ReleaseMenu(lpMenu As Object)

lpMenu.Enabled = False
lpMenu.Visible = False

End Sub

Public Function LoadSetting(Section As String, Key As String, Optional Default As String) As String

On Error Resume Next
LoadSetting = GetSetting("YourTime", Section, Key, Default)

End Function

Public Sub SetSetting(Section As String, Key As String, Value As String)

On Error Resume Next
SaveSetting "YourTime", Section, Key, Value

End Sub

Public Sub ShowAbout()

frmAbout.MainLoop

End Sub

Public Function FindSelected(lpMenu As Object) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.FindSelected(lpMenu)", Array(VarPtr(lpMenu)), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&

For Tell = 0 To lpMenu.count - 1
    If lpMenu.Item(Tell).Checked Then
        FindSelected = Tell
        Exit Function
    End If
Next

FindSelected = -1

End Function

Public Sub SetControls(Form As Object, BeginWith As String, Property As String, Value As Variant)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.SetControls(Form, BeginWith, Property, Value)", Array(Form.Name, BeginWith, Property, Value), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Control As Control

For Each Control In Form.Controls
    If Mid(Control.Name, 1, Len(BeginWith)) = BeginWith Then
        CallByName Control, Property, VbLet, Value
    End If
Next

End Sub

Public Function FindUser(strName As String) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.FindUser(strName)", Array(strName), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&

For Tell = LBound(Users) To UBound(Users)
    If strName = Users(Tell).UserName Then
        FindUser = Tell
        Exit Function
    End If
Next

End Function

Public Function NewUser(strCaption As String, Optional strName As String, Optional strPassword As String, Optional OldPass As Variant = False) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.NewUser(strCaption, [strName], [strPassword])", Array(strCaption, strName, strPassword), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

frmUser.txtUserName.Text = strName
frmUser.txtPassword.Text = strPassword
frmUser.Caption = strCaption
frmUser.lblOldPassword.Visible = (OldPass = True)
frmUser.txtOldPassword.Visible = (OldPass = True)
frmUser.Tag = ""
frmUser.Show

Do Until frmUser.Tag <> ""
    Sleep 10
    DoEvents
Loop

If frmUser.Tag = "success" Then
    OldPass = frmUser.txtOldPassword
    strName = frmUser.txtUserName.Text
    strPassword = frmUser.txtPassword.Text
    Exit Function
End If

NewUser = -1

End Function

Public Function ValidateUser() As Boolean

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.ValidateUser", , EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

If Users(UserID).Password <> "" Then

    frmMain.Enabled = False
    frmLogin.Tag = "waiting"
    frmLogin.Show

    Do Until frmLogin.Tag <> "waiting"
        DoEvents
        Sleep 25
    Loop
    
    frmMain.Enabled = True
    
    If frmLogin.Tag <> "success" Then
        ValidateUser = False
        Exit Function
    End If
    
End If

Users(UserID).LoggedOn = Now
ValidateUser = True

End Function

Public Function ProcessTags(ByVal Expression As String, lDate As Date, RefDate As Date) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.ProcessTags(Expression, lDate, RefDate)", Array(Expression, lDate, RefDate), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Val As Long

' Look after tags
If InStr(1, Expression, "(!YA)", vbTextCompare) > 0 Then

    ' Calculate the years
    Val = DateDiff("yyyy", RefDate, lDate)
    
    If Val >= 0 And Script.EnableTags("YA") = 1 Then
        Expression = Replace(Expression, "!YA", Val, , , vbTextCompare)
    Else
        Expression = Replace(Expression, "(!YA)", "", , , vbTextCompare)
    End If

End If

If Script.EnableTags("YA") = 0 Then
    Expression = Replace(Expression, "(!CH)", "", , , vbTextCompare)
End If

ProcessTags = Expression

End Function

Public Function ProcessVariables(ByVal strText As String) As String

' Seven different variables
strText = Replace(strText, "%year%", CurrentDate.cYear, , , vbTextCompare)
strText = Replace(strText, "%day%", CurrentDate.cDay, , , vbTextCompare)
strText = Replace(strText, "%month%", CurrentDate.cMonth, , , vbTextCompare)
strText = Replace(strText, "%monthname%", CurrentDate.cMonthName, , , vbTextCompare)
strText = Replace(strText, "%totaldays%", CurrentDate.cTotalDays, , , vbTextCompare)
strText = Replace(strText, "%dayname%", CurrentDate.cDayName, , , vbTextCompare)
strText = Replace(strText, "%weekday%", CurrentDate.cWeekDay, , , vbTextCompare)

ProcessVariables = strText

End Function

Public Function RemoveTags(Expression As String) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.RemoveTags(Expression)", Array(Expression), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell As Long, JumpOver As Boolean, sChar As String

For Tell = 1 To Len(Expression)

    sChar = Mid(Expression, Tell, 1)

    If sChar = "(" Then
        If Tell = Len(Expression) Then
            JumpOver = True
        Else
            JumpOver = (Mid(Expression, Tell + 1, 1) = "!")
        End If
    End If
    
    If JumpOver = False Then RemoveTags = RemoveTags & sChar
    If sChar = ")" Then JumpOver = False
Next

End Function

Public Function ShowInformation(Text As String) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.ShowInformation(Text)", Array(Text), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

frmInformation.txtText.Text = Text
frmInformation.Tag = ""
frmInformation.Show

Do Until frmInformation.Tag <> ""
    Sleep 10
    DoEvents
Loop

If LCase(frmInformation.Tag) = "success" Then
    Text = frmInformation.txtText.Text
    ShowInformation = 1
Else
    ShowInformation = 0
End If

End Function

Public Function ChoosePerson() As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.ChoosePerson", , EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

frmChoosePerson.AddPersons
frmChoosePerson.Tag = ""
frmChoosePerson.Show

Do Until frmChoosePerson.Tag <> ""
    Sleep 10
    DoEvents
Loop

If LCase(frmChoosePerson.Tag) = "success" Then
    ChoosePerson = frmChoosePerson.cmbChoose.ListIndex
Else
    ChoosePerson = -1
End If

End Function

Public Function RemoveTime(lDate As Date) As Date
    
' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.RemoveTime(lDate)", Array(lDate), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

RemoveTime = DateSerial(Year(lDate), Month(lDate), Day(lDate))

End Function

Public Function ConvertToNumeric(Text As String) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.ConvertToNumeric(Text)", Array(Text), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell As Long, LocalDesimal$, Char$, Buff$

LocalDesimal = LocalInfo.DecimalSeparator

For Tell = 1 To Len(Text)

    Char = Mid(Text, Tell, 1)

    If IsNumeric(Char) Or Char <> "." Or Char <> LocalDesimal Then
        
        If Char = "." Then
            Char = LocalDesimal
        End If
        
        Buff = Buff & Char
    End If

Next

ConvertToNumeric = Buff

End Function

Public Sub AddRem(ListBox As Object, lDate As Date)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.AddRem(ListBox, lDate)", Array(ListBox.Parent.Name & "." & ListBox.Name, lDate), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, h&, m&, Num&

ListBox.Clear

For Tell = 1 To (24 * 60) / RemStep

    Num = Search(Users(UserID).DataRem, lDate + TimeSerial(h, m, 0), 0)

    If Num >= 0 Then
        ListBox.AddItem FillOut(CStr(h), 2) & "." & FillOut(CStr(m), 2) & "  " & Users(UserID).DataRem(Num).Text
    End If
    
    m = m + RemStep
    
    If m >= 60 Then
        m = m - 60
        h = h + 1
    End If
    
    If h > 23 Then Exit For
Next

End Sub

Public Sub AddTasks(ListBox As Object, lDate As Date)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.AddTasks(ListBox, lDate)", Array(ListBox.Parent.Name & "." & ListBox.Name, lDate), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, Num&

ListBox.AddItem ""
ListBox.AddItem "OPPGAVER:"

For Tell = 0 To 50
    
    Num = Search(Users(UserID).DataTasks, lDate + TimeSerial(0, Tell, 0), 0)
    
    If Num >= 0 Then
        ListBox.AddItem Users(UserID).DataTasks(Num).Text
    End If

Next

If ListBox.ListCount <= 2 Then
    ListBox.Clear
End If

End Sub

Public Function ChooseDate(lStartDate As Date, lEndDate As Date) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.ChooseDate(lStartDate, lEndDate)", Array(lStartDate, lEndDate), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

frmChooseDate.Tag = ""
frmChooseDate.Show

Do Until frmChooseDate.Tag <> ""
    Sleep 10
    DoEvents
Loop

If LCase(frmChooseDate.Tag) = "success" Then
    lStartDate = Script.ToDate(frmChooseDate.txtStartDate)
    lEndDate = Script.ToDate(frmChooseDate.txtEndDate)
Else
    ChooseDate = -1
End If

End Function

Public Sub ResetSettings()

DeleteSetting "YourTime"

End Sub

Public Sub EventFormLoad(Form As Object)

FormLoad Form

End Sub

Public Function SubClass(hwnd As Long) As Long

If UseHooking Then
    HookForm hwnd
Else
    SubClass = -1
End If

End Function

Public Function UnSubClass(hwnd As Long) As Long

If UseHooking Then
    UnHookForm hwnd
Else
    UnSubClass = -1
End If

End Function

Public Function GetScriptObjects() As Variant

GetScriptObjects = Array("Main", "Print", "CurrentDate", "Sun", "Common", "SMTP", "Socket", "LocalInfo", "frmAbout", _
"frmBackup", "frmChooseDate", "frmchoosePerson", "frmDiary", "frmInformation", "frmLogin", "frmMain", _
"frmOwn", "frmPerson", "frmPhonebook", "frmScript", "frmSettings", "frmShowError", "frmUser", "frmYear", "frmerrorinfo")

End Function

Public Function ObjectPtr(ObjectName As String) As Object

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.ObjectPtr(ObjectName)", Array(ObjectName), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Select Case LCase(ObjectName)
Case "main": Set ObjectPtr = Me
Case "print": Set ObjectPtr = PrintClass
Case "currentDate": Set ObjectPtr = CurrentDate
Case "sun": Set ObjectPtr = Sun
Case "common": Set ObjectPtr = Common
Case "localinfo": Set ObjectPtr = LocalInfo
Case "smtp": Set ObjectPtr = SMTP
Case "socket": Set ObjectPtr = New clsSocket
Case "frmabout": Set ObjectPtr = frmAbout
Case "frmbackup": Set ObjectPtr = frmBackup
Case "frmchoosedate": Set ObjectPtr = frmChooseDate
Case "frmchooseperson": Set ObjectPtr = frmChoosePerson
Case "frmdiary": Set ObjectPtr = frmDiary
Case "frminformation": Set ObjectPtr = frmInformation
Case "frmlogin": Set ObjectPtr = frmLogin
Case "frmmain": Set ObjectPtr = frmMain
Case "frmown": Set ObjectPtr = frmOwn
Case "frmperson": Set ObjectPtr = frmPerson
Case "frmphonebook": Set ObjectPtr = frmPhonebook
Case "frmscript": Set ObjectPtr = frmScript
Case "frmsettings": Set ObjectPtr = frmSettings
Case "frmshowerror": Set ObjectPtr = frmShowError
Case "frmuser": Set ObjectPtr = frmUser
Case "frmyear": Set ObjectPtr = frmYear
Case "frmerrorinfo": Set ObjectPtr = frmErrorInfo
End Select

End Function

Public Sub EnumerateMethods(ListView As Object, lpObject As Object)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.EnumerateMethods(ListView, lpObject)", Array(ListView.Name, ObjPtr(lpObject)), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim ClassInfo As Object, Tell&, InvKind&

' Get interface information about the object
Set ClassInfo = TLIApp.InterfaceInfoFromObject(lpObject)

' Clear all data
ListView.ListItems.Clear

For Tell = 1 To ClassInfo.Members.count

    InvKind = ClassInfo.Members(Tell).InvokeKind

    If Not (InvKind = INVOKE_PROPERTYPUT Or InvKind = INVOKE_PROPERTYPUTREF) Then
        ListView.ListItems.Add , , ClassInfo.Members(Tell).Name, , IIf(InvKind = INVOKE_FUNC Or INVOKE_UNKNOWN, 1, 2)
    End If

Next

End Sub

Public Function GetFolderList(Path As String, ByVal Extention As Variant) As Collection

Dim sFile As String, AllowAll As Boolean

Set GetFolderList = New Collection

sFile = Dir(Path)
Extention = Split(Extention, ";")
AllowAll = InArray(Extention, "*") >= 0

Do Until sFile = ""

    If InArray(Extention, GetExtention(sFile)) >= 0 Or AllowAll Then
        GetFolderList.Add sFile
    End If

    sFile = Dir
Loop

End Function

Public Sub RemoveFolder(sPath As String)

Dim sFile As Variant

For Each sFile In GetFolderList(sPath, "*")
    Kill ValidPath(sPath) & sFile
Next

RmDir sPath

End Sub

Public Sub Uninstall()

Dim sFile As Variant

For Each sFile In GetFolderList(App.Path & "\Plugins\", "dll")
    ' Unregister ActiveX-dll
   UnregisterPlugin CStr(sFile)
Next

' Remove folders
RemoveFolder App.Path & "\Plugins\"
RemoveFolder App.Path & "\Data\"

' Delete all data in registry
ResetSettings

End Sub

Public Sub UnregisterPlugin(sFile As String)

Shell "regsvr32 /u " & sFile

End Sub

Public Function InArray(lpArray As Variant, Find As String) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.InArray(lpArray, Find)", Array(lpArray, Find), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell As Long

If IsArray(lpArray) Then

    For Tell = LBound(lpArray) To UBound(lpArray)
        If LCase(lpArray(Tell)) = LCase(Find) Then
            InArray = Tell
            Exit Function
        End If
    Next

End If

InArray = -1

End Function

Public Function JoinArray(Array1, Array2)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsScript.JoinArray(Array1, Array2)", Array(Array1, Array2), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell As Long, aTmp()

ReDim Preserve aTmp(LBound(Array1) To UBound(Array1) + UBound(Array1) + 1)

For Tell = LBound(Array1) To UBound(Array1)
    aTmp(Tell) = Array1(Tell)
Next

For Tell = UBound(Array1) + 1 To UBound(Array2) + UBound(Array1) + 1
    aTmp(Tell) = Array2(Tell - UBound(Array1) - 1)
Next

Let JoinArray = aTmp

End Function

Public Function OpenRecordset(Name As String) As clsRecordset

Set OpenRecordset = New clsRecordset

Select Case LCase(Name)
Case "diary": OpenRecordset.Connection Users(UserID).DataDiary
Case "own": OpenRecordset.Connection Users(UserID).DataOwn
Case "rem": OpenRecordset.Connection Users(UserID).DataRem
Case "tasks": OpenRecordset.Connection Users(UserID).DataTasks
Case "staticdays": OpenRecordset.Connection StaticDays
Case Else: Set OpenRecordset = Nothing
End Select

End Function

Public Function SendError(strErrorMessage, strEmail, strName, strEvents)

Dim sBuff As Variant

' Header
sBuff = sBuff & "This message is auto generatet, and sended because " & _
        "there has occured an error in the program 'YourTime'. Following is" & _
        " a description of the problem:" & vbCrLf

sBuff = sBuff & "- User information:" & vbNewLine & vbCrLf & _
                "  Email:   " & strEmail & vbCrLf & _
                "  Name:    " & strName & vbCrLf & _
                "  Version: " & Version & vbCrLf & _
                "  Windows: " & WindowsVersion & vbCrLf & vbCrLf & _
                "- Software: " & vbCrLf & "- NOT AVALIBLE -" & vbCrLf & _
                "- Events: " & vbCrLf & strEvents & vbCrLf & vbCrLf & _
                "- Error message: " & vbCrLf & strErrorMessage & vbCrLf

SMTP.Port = 25
SMTP.Server = "mail.online.no" ' The hotmail server isn't reliable enought
SMTP.MailFrom = strEmail
SMTP.MailTo = "kr_stang@hotmail.com"
SMTP.NameFrom = strName
SMTP.NameTo = "Kristian"
SMTP.Subject = "Error report; Your Time"
SMTP.Body = sBuff

' Send the email
SMTP.SendMail

End Function

Public Sub InvokeError(ErrNum As Long, Optional Description As String)

On Error Resume Next
Err.Number = ErrNum
Err.Description = Description

ErrorIn "clsScript.InvokeError(ErrNum, Description)", Array(ErrNum, Description), EA_NORERAISE
HandleError

End Sub

Public Function IsDateSerial(Year As Integer, Month As Integer, Day As Integer) As Boolean

Dim remYear As Long

remYear = CurrentDate.cYear

If Year < 0 Or Year > 9999 Then
    Exit Function
ElseIf Month < 1 Or Month > 12 Then
    Exit Function
ElseIf Day < 1 Or Day > CurrentDate.GetMonthLenght(CLng(Month)) Then
    Exit Function
End If

CurrentDate.cYear = remYear
IsDateSerial = True

End Function

Public Function RemoveDuplicans(sText As String, sLettersToRemove As String, Optional Compare As VbCompareMethod = vbTextCompare)

Dim sBuff As String, Tell&, Cnt&, Char$

sBuff = Space(Len(sText))

For Tell = 1 To Len(sText)

    Char = Mid(sText, Tell, 1)

    If Not (InStr(1, sLettersToRemove, Char, Compare) > 0 And InStr(1, sBuff, Char, Compare) > 0) Then
        Cnt = Cnt + 1
        Mid(sBuff, Cnt, 1) = Char
    End If
Next

RemoveDuplicans = Left(sBuff, Cnt)

End Function

Public Sub UnregisterObject(lpObject As Object, Optional lOLEInstance As Variant)

On Error Resume Next
'Once we are done with the main program, lets clean up the rot
'by removing the entry for our ActiveX Server

If IsMissing(OLEInstance) Then
    lOLEInstance = OLEInstance
End If

If OLEInstance <> 0 Then
    RevokeActiveObject lOLEInstance, 0
End If

CoDisconnectObject lpObject, 0
    
End Sub

Public Sub Quit()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "Script.Quit", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Form As Form

' Save all neccesary data
frmMain.ProcessControls "saveall"

' Remove propeties
RemoveProp frmMain.hwnd, "SecProc"

' Unsubclass frmMain
UnHookForm frmMain.hwnd

' Save all data
SaveData ValidPath(App.Path) & "Data\"

' Unregister ActiveX-EXE object
Script.UnregisterObject Script

For Each Form In Forms
    If LCase(Form.Name) <> "frmmain" Then Unload Form
Next

ErrSysHandlerRelease

End Sub

Public Function IsIDE() As Boolean

bIsIDE = False
Debug.Assert CheckIDE

IsIDE = bIsIDE

End Function

Private Function CheckIDE() As Boolean

bIsIDE = True
CheckIDE = True

End Function

Private Sub Class_Terminate()

On Error Resume Next

DefaultLanguage = lLanguage
RevokeActiveObject OLEInstance, 0

End Sub
