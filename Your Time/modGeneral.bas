Attribute VB_Name = "modGeneral"
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

' Memory methods and thread control
Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Lots of stuff
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function GetFocus Lib "user32" () As Long

' To obtain the command line
Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

' Varius uses, but it's mainly implementet to support numeric textboxes
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' To draw transparent moon phases
Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

' To get info from the system-register
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

' Gdi32 api-cals (Graphics)
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal cp As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetRect Lib "user32.dll" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

' Windows-version
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' For printing
Public Const EM_FORMATRANGE As Long = WM_USER + 57
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Public Enum ENUM_ERRORS_DEMO
    ERRD_FIRST = ERRMAP_APP_FIRST
    ERRD_API
    ERRD_XMLLIB_NOTAVAILABLE
End Enum

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Type InternalClipboard
    DataType As Long
    Data As Remember
End Type

Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type Buffer
    hdc As Long
    Picture As Long
    Brush As Long
    Font As Long
    Pen As Long
    Rect As Rect
    ScaleWidth As Integer
    ScaleHeight As Integer
End Type

Type CharRange
    cpMin As Long
    cpMax As Long
End Type

Type FormatRange
    hdc As Long
    hdcTarget As Long
    rc As Rect
    rcPage As Rect
    chrg As CharRange
End Type

Type Stjerner
    X As Double
    Y As Double
    r As Byte
    Color As Long
End Type

Type Line
    Text As String
    Style As Long
End Type

Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
  ntmFlags As Long
  ntmSizeEM As Long
  ntmCellHeight As Long
  ntmAveWidth As Long
End Type

' For the createfont
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const DEFAULT_CHARSET = 1
Public Const OEM_CHARSET = 255
Public Const OEM_FIXED_FONT = 10
Public Const OUT_DEFAULT_PRECIS = 0
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_MASK = &HF
Public Const TRANSPARENT = 1
Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const PROOF_QUALITY = 2

' For the DrawTextA api
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_CHARSTREAM = 4
Public Const DT_DISPFILE = 6
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5
Public Const DT_NOPREFIX = &H800
Public Const DT_NOCLIP = &H100
Public Const DT_PLOTTER = 0
Public Const DT_RASCAMERA = 3
Public Const DT_RASDISPLAY = 1
Public Const DT_RASPRINTER = 2
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

' Getting information from the system registry
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const StrRun = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Public Const gREGVALSYSINFOLOC = "MSINFO"
Public Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Public Const gREGVALSYSINFO = "PATH"

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

' tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4
Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

' EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

' Just to remove duplicate letters in text
Const strAlphabet = "abcdefghijklmnopqrstuvwxyz"

Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_CANPASTE = (WM_USER + 50)

Public Const Arthor = "Kristian S.Stangeland"
Public Const RemStep = 15

Public UserID As Long

' Global classes
Public Sun As clsSunTime
Public Common As New clsCommon
Public Script As New clsScript
Public CurrentDate As New clsDate
Public PrintClass As New clsPrint
Public LocalInfo As New clsLocalInfo
Public Language As New clsLanguage
Public MD5 As New clsMD5
Public SMTP As New clsSMTP

' Plugins
Public GlobalObjects As New Collection
Public Plugins As New Collection

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean

Dim hKey As Long
Dim i As Long
Dim rc As Long
Dim KeyValType As Long
Dim tmpVal As String
Dim KeyValSize As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler:
    
    ErrorIn "modGeneral.UpdateButtons", , EA_NORERAISE
    HandleError
    
    KeyVal = ""
    GetKeyValue = False
    rc = RegCloseKey(hKey)
    
    Exit Function
End If
' *** BEGIN CODE ***

rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)

If (rc <> ERROR_SUCCESS) Then GoTo errHandler

tmpVal = String$(1024, 0)
KeyValSize = 1024

rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)

If (rc <> ERROR_SUCCESS) Then GoTo errHandler

If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then
    tmpVal = Left$(tmpVal, KeyValSize - 1)
Else
    tmpVal = Left$(tmpVal, KeyValSize)
End If

Select Case KeyValType
Case REG_SZ

    KeyVal = tmpVal
    
Case REG_DWORD

    For i = Len(tmpVal) To 1 Step -1
        KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))
    Next
    
    KeyVal = Format$("&h" + KeyVal)
    
End Select

GetKeyValue = True
rc = RegCloseKey(hKey)

End Function

Public Sub SetNumber(NumberText As TextBox, Flag As Boolean)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.SetNumber(NumberText, Flag)", Array(NumberText.Parent.Name & "." & NumberText.Name, Flag), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim curstyle As Long, newstyle As Long

curstyle = GetWindowLong(NumberText.hwnd, GWL_STYLE)

If Flag Then
   curstyle = curstyle Or ES_NUMBER
Else
   curstyle = curstyle And (Not ES_NUMBER)
End If

newstyle = SetWindowLong(NumberText.hwnd, GWL_STYLE, curstyle)
NumberText.Refresh
    
End Sub

Sub Main()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.Main", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim RegScript As String, Folder As Collection

ErrSysHandlerSet
UserID = Val(GetSetting("YourTime", "Constants", "StartupUser", 1)) ' Can't call to the script engine until after we have done this

If UserID < 1 Or UserID > 20 Then
    UserID = 1
End If

' Load in the sun-module
Set Sun = New clsSunTime

' Update the sun-module
UpdateSun

' Initialize globalobjects
GlobalObjects.Add Script, Script.Name
GlobalObjects.Add CurrentDate, CurrentDate.Name
GlobalObjects.Add Sun, Sun.Name
GlobalObjects.Add PrintClass, PrintClass.Name
GlobalObjects.Add Common, Common.Name
GlobalObjects.Add LocalInfo, LocalInfo.Name
GlobalObjects.Add Language, Language.Name
GlobalObjects.Add MD5, MD5.Name
GlobalObjects.Add SMTP, SMTP.Name

' Register the object
Script.RegisterObject Script, "YourTime.clsScript"

' Load the language packs
If LenB(Script.LanguagePack) = 0 Then
    ' If there's no language pack specified, try to get one from the application folder
    Set Folder = Script.GetFolderList(ValidPath(App.Path), "lpk")
    
    If Folder.count > 0 Then
        ' Use the first file (in alphabetical order)
        Script.LanguagePack = Folder.Item(1)
    End If
End If

' Call the language function
Language.LoadLanguagePack ValidPath(App.Path) & Script.LanguagePack

' Run script from the command line
If Script.ProcessCL Then
    Script.Run Script.CommandLine
End If

' Run script-file as defined in the registry
If Script.RunScript Then
    RegScript = Script.RegScript
    If RegScript <> "" Then Script.Run RegScript
End If

' Disable hooking (We're working in IDE, so using hooking will prevent us from debuging)
If Script.IsIDE Then
    Script.UseHooking = False
    SaveSetting "YourTime", "Constants", "HookingOn", True
Else
    If GetSetting("YourTime", "Constants", "HookingOn", False) Then
        Script.UseHooking = True
        DeleteSetting "YourTime", "Constants", "HookingOn"
    End If
End If

SetProperties

LoadObjects
LoadPlugins
LoadData ValidPath(App.Path) & "Data\"
InitializeVariables

' Show the main-form
frmMain.ProcessControls "updatediary"
frmMain.UpdateAll
frmMain.Show

' Validate the user
If Script.ValidateUser = False Then
    Unload frmMain
    End
End If

End Sub

Public Sub SetProperties()

' If hooking is on

If Script.UseHooking Then
    Check SetProp(frmMain.hwnd, "SecProc", AddressOf WindowProc) <> 0, ERRD_API, "SetProp returned 0."
    Check SetProp(frmMain.hwnd, "MinWidth", Script.MinWidth) <> 0, ERRD_API, "SetProp returned 0."
    Check SetProp(frmMain.hwnd, "MinHeight", Script.MinHeight) <> 0, ERRD_API, "SetProp returned 0."
    Check SetProp(frmScript.hwnd, "DeTop", True) <> 0, ERRD_API, "SetProp returned 0."
    Check SetProp(frmShowError.hwnd, "DeTop", True) <> 0, ERRD_API, "SetProp returned 0."
    Check SetProp(frmMain.hwnd, "HandleMinMax", 1) <> 0, ERRD_API, "SetProp returned 0."
    Check SetProp(frmAbout.hwnd, "TransDisable", 1) <> 0, ERRD_API, "SetProp returned 0."
End If

End Sub

Public Sub UpdateSun()

Sun.Latitude = Script.Latitude
Sun.Longitude = Script.Longitude
Sun.Sommertid = Script.Sommertid
Sun.TimeZone = Script.TimeZone

End Sub

Public Sub InitializeVariables()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.InitializeVariables", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim lTmp As Long

Erase StaticDays
AddRecord StaticDays, DateSerial(2004, 1, 21), "Prinsesse Ingrid Alexandra (!YA)(!CF)", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 1, 1), "Nyttårsdag", "--.--.****"
AddRecord StaticDays, CurrentDate.FirstWeekday(vbSunday, 1, 1), "Kristi åpenbaringssøndag", "--.--.****"
AddRecord StaticDays, CurrentDate.FirstWeekday(vbSunday, 2, 2), "Morsdag", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 2, 14), "Valentinsdag", "--.--.****"
AddRecord StaticDays, DateSerial(1937, 2, 21), "Kong Harald V's fødelsdag (!YA)(!CF)", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 3, 20) + IIf(CurrentDate.cLeapYear = 29, 1, 0), "Vårjevndøgn", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 4, 1), "Aprilsnarr", "--.--.****"

lTmp = AddRecord(StaticDays, CurrentDate.EasterDay, "Påskedag", "--.--.****")
AddRecord StaticDays, StaticDays(lTmp).RemDate + 1, "2. Påskedag (!CH)", "--.--.****"
AddRecord StaticDays, StaticDays(lTmp).RemDate - 1, "Påskeaften (!CH)", "--.--.****"
AddRecord StaticDays, StaticDays(lTmp).RemDate - 2, "Langfredag (!CH)", "--.--.****"
AddRecord StaticDays, StaticDays(lTmp).RemDate - 3, "Skjærtorsdag (!CH)", "--.--.****"
AddRecord StaticDays, StaticDays(lTmp).RemDate - 7, "Palmesøndag", "--.--.****"
AddRecord StaticDays, StaticDays(lTmp).RemDate + 39, "Kristi himmelfartsdag (!CH)", "--.--.****"
AddRecord StaticDays, StaticDays(lTmp).RemDate + 49, "Pinsedag (!CH) (!CF)", "--.--.****"
AddRecord StaticDays, StaticDays(lTmp).RemDate + 50, "2. Pinsedag (!CH)", "--.--.****"

AddRecord StaticDays, DateSerial(2004, 5, 1), "Off. høytidsdag (!CH) (!CF)", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 5, 8), "Frigjøringsdag (!CF)", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 5, 17), "Grunnlovsdag (!CH)", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 6, 23), "St.Hansaften", "--.--.****"
AddRecord StaticDays, DateSerial(1937, 7, 4), "Dronning Sonjas fødselsdag (!YA)(!CF)", "--.--.****"
AddRecord StaticDays, DateSerial(1973, 7, 20), "Kronprins Haakons fødselsdag (!YA)(!CF)", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 7, 23), "Første hundedag", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 7, 29), "Oslok", "--.--.****"
AddRecord StaticDays, DateSerial(1973, 8, 19), "Mette-Marits fødselsdag (!YA)(!CF)", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 9, 22) + IIf(CurrentDate.cLeapYear = 29, 1, 0), "Høstjevndøgn", "--.--.****"
AddRecord StaticDays, DateSerial(1971, 9, 22), "Märtha Louises fødselsdag (!YA)(!CF)", "--.--.****"
AddRecord StaticDays, CurrentDate.FirstWeekday(vbSunday, 11), "Allehelgensdag", "--.--.****"
AddRecord StaticDays, CurrentDate.FirstWeekday(vbSunday, 11) - 7, "Bots- og bededag", "--.--.****"
AddRecord StaticDays, CurrentDate.FirstWeekday(vbSunday, 11, 2), "Farsdag", "--.--.****"
AddRecord StaticDays, CurrentDate.WeekdaysFromDate(DateSerial(CurrentDate.cYear, 12, 24), vbSunday, -1, 4), "1. Advent", "--.--.****"
AddRecord StaticDays, CurrentDate.WeekdaysFromDate(DateSerial(CurrentDate.cYear, 12, 24), vbSunday, -1, 3), "2. Advent", "--.--.****"
AddRecord StaticDays, CurrentDate.WeekdaysFromDate(DateSerial(CurrentDate.cYear, 12, 24), vbSunday, -1, 2), "3. Advent", "--.--.****"
AddRecord StaticDays, CurrentDate.WeekdaysFromDate(DateSerial(CurrentDate.cYear, 12, 24), vbSunday, -1, 1), "4. Advent", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 12, 25), "Juledag (!CH)", "--.--.****"
AddRecord StaticDays, DateSerial(2004, 12, 26), "2. Juledag (!CH)", "--.--.****"

End Sub

Public Sub LoadObjects()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.LoadObjects", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

With frmMain
    For Tell = 1 To 19
        Load .cmdUsers(Tell)
        .cmdUsers(Tell).Visible = True
    Next
End With

End Sub

Public Sub LoadPlugins()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.LoadPlugins", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim sFile As Variant, Plugin As Object, strClassName As String

If Script.FindPlugins = False Then
    Exit Sub
End If

For Each sFile In Script.GetFolderList(App.Path & "\Plugins\", "dll;exe")

    Select Case GetExtention(CStr(sFile))
    Case "dll"

        If GetSetting("YourTime", "Plugin", GetFileBase(CStr(sFile)), True) = True Then
        
            Err.Clear
            
            ' How the class is registered in the registry
            strClassName = GetFileBase(CStr(sFile)) & ".PluginMain"
            
            ' Try to create the object
            Set Plugin = CreateObject(strClassName)
    
            If Err = 429 Then ' ERROR: ActiveX component can't create object
                ' Try to register the object
                Shell "regsvr32 " & Chr(34) & ValidPath(App.Path) & "Plugins\" & CStr(sFile) & Chr(34)
                
                ' Load the plugin again
                Set Plugin = CreateObject(strClassName)
            End If
    
            ' Add plugin
            Plugins.Add Plugin, Plugin.Name
    
            ' Initialize plugin
            Plugin.Initialize GlobalObjects
        End If
    
    Case "exe"
    
        ' Shell the ActiveX-DLL
        Shell ValidPath(App.Path) & "Plugin\" & CStr(sFile)
    
    End Select
    
Next

End Sub

Public Function CapitalizeFirstLetter(ByVal Text As String) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.CapitalizeFirstLetter(Text)", Array(Text), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

If Text = "" Then Exit Function

Mid$(Text, 1, 1) = UCase(Mid$(Text, 1, 1))
CapitalizeFirstLetter = Text

End Function

Public Function FillOut(Text As String, Lenght As Long) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.FillOut(Text, Lenght)", Array(Text, Lenght), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

FillOut = String(Lenght - Len(Text), "0") & Text

End Function

Public Function GetExtention(File As String) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.GetExtention(File)", Array(File), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

GetExtention = Right(File, Len(File) - InStrRev(File, "."))

End Function

Public Function GetFileName(File As String) As String

On Error Resume Next
GetFileName = Right(File, Len(File) - InStrRev(File, "\"))

End Function

Public Function GetFileBase(File As String) As String

On Error Resume Next
Dim Buff As String

Buff = GetFileName(File)
GetFileBase = Left(Buff, InStr(1, Buff, ".") - 1)

End Function

Public Function ValidPath(Path As String) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.StartHooking(Path)", Array(Path), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

ValidPath = Path & IIf(Right(Path, 1) = "\", "", "\")

End Function

Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, lParam As ListBox) As Long

On Error Resume Next
Dim FaceName$, FullName$, Buff$

FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
Buff = Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
lParam.AddItem Buff
EnumFontFamProc = 1

End Function

Public Sub FillComboWithFonts(CB As ComboBox)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modGeneral.FillComboWithFonts(CB)", Array(CB.Parent.Name & "." & CB.Name), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim hdc As Long

CB.Clear
hdc = GetDC(CB.hwnd)
EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, CB
ReleaseDC CB.hwnd, hdc

End Sub

Public Sub FormLoad(Form As Form)

' Subclass the form
HookForm Form.hwnd

' Set the language
Language.SetLanguageInForm Form

End Sub

Public Function SafeUBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long

On Error Resume Next
Dim lAddress&, cElements&, lLbound&, cDims%

If Dimension < 1 Then
    SafeUBound = -1
    Exit Function
End If

CopyMemory lAddress, ByVal lpArray, 4

If lAddress = 0 Then
    ' The array isn't initilized
    SafeUBound = -1
    Exit Function
End If

' Calculate the dimenstions
CopyMemory cDims, ByVal lAddress, 2
Dimension = cDims - Dimension + 1

' Obtain the needed data
CopyMemory cElements, ByVal (lAddress + 16 + ((Dimension - 1) * 8)), 4
CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4

SafeUBound = cElements + lLbound - 1

End Function

Public Function SafeLBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long

On Error Resume Next
Dim lAddress&, cElements&, lLbound&, cDims%

If Dimension < 1 Then
    SafeLBound = -1
    Exit Function
End If

CopyMemory lAddress, ByVal lpArray, 4

If lAddress = 0 Then
    ' The array isn't initilized
    SafeLBound = -1
    Exit Function
End If

' Calculate the dimenstions
CopyMemory cDims, ByVal lAddress, 2
Dimension = cDims - Dimension + 1

' Obtain the needed data
CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4

SafeLBound = lLbound

End Function
