Attribute VB_Name = "modHooking"
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

Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Type POINTAPI
    X As Long
    Y As Long
End Type

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10

Public Const MFS_DISABLED = &H2
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_ENABLED = &H0
Public Const MFS_CHECKED = &H8

Public Const WM_USER = &H400
Public Const WM_SYSCOMMAND = &H112
Public Const WM_INITMENU = &H116
Public Const WM_NCDESTROY = &H82
Public Const WM_YOURTIME = WM_USER + &H200
Public Const WM_DESTROY = &H2
Public Const WM_GETMINMAXINFO = &H24

Public Const WM_GETFONT = &H31
Public Const EM_GETSEL = &HB0
Public Const EM_POSFROMCHAR = (WM_USER + 38)

Public Const PAGE_EXECUTE_READWRITE = &H40
Public Const MEM_COMMIT = &H1000
Public Const MEM_DECOMMIT = &H4000
Public Const MEM_RELEASE = &H8000

Public Const GWL_WNDPROC = (-4)

Public Sub HookForm(hwnd As Long)

If Script.UseHooking = True Then
    If GetProp(hwnd, "PrevProc") = 0 Then
        Check SetProp(hwnd, "PrevProc", SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)) <> 0, ERRD_API, "SetProp returned 0."
        Check SetProp(hwnd, "OldStyle", GetWindowLong(hwnd, GWL_EXSTYLE)) <> 0, ERRD_API, "Couldn't save EXSTYLE for window. Transparent effect will not work"
    End If
End If

End Sub

Public Sub UnHookForm(hwnd As Long)

Dim PrevProc As Long

PrevProc = GetProp(hwnd, "PrevProc")

If PrevProc <> 0 Then
    SetWindowLong hwnd, GWL_WNDPROC, PrevProc
    RemoveProp hwnd, "OnTop"
    RemoveProp hwnd, "Transparent"
    RemoveProp hwnd, "TransDisable"
    RemoveProp hwnd, "MinWidth"
    RemoveProp hwnd, "MinHeight"
    RemoveProp hwnd, "MaxWidth"
    RemoveProp hwnd, "MaxHeight"
    RemoveProp hwnd, "HandleMinMax"
    RemoveProp hwnd, "DeTop"
    RemoveProp hwnd, "ItemAdded"
    RemoveProp hwnd, "OldStyle"
End If

End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next ' Det nytter ikke med en error-handler i slike prosedyrer

Dim PrevProc As Long, MM As MINMAXINFO, ret&, onTop As Boolean
Dim mData(3) As Long, hSysMenu As Long, mii As MENUITEMINFO

PrevProc = GetProp(hwnd, "PrevProc")

Select Case uMsg
Case WM_GETMINMAXINFO

    If GetProp(hwnd, "HandleMinMax") = 1 Then
        CopyMemory MM, ByVal lParam, Len(MM)
        
        mData(0) = GetProp(hwnd, "MinWidth")
        mData(1) = GetProp(hwnd, "MinHeight")
        mData(2) = GetProp(hwnd, "MaxWidth")
        mData(3) = GetProp(hwnd, "MaxHeight")
            
        If mData(0) <> 0 Then MM.ptMinTrackSize.X = mData(0)
        If mData(1) <> 0 Then MM.ptMinTrackSize.Y = mData(1)
        If mData(2) <> 0 Then MM.ptMaxTrackSize.X = mData(2)
        If mData(3) <> 0 Then MM.ptMaxTrackSize.Y = mData(3)
        
        CopyMemory ByVal lParam, MM, Len(MM)
        Exit Function
    End If

Case WM_INITMENU

    If GetProp(hwnd, "ItemAdded") = 0 Then
        AddMenuItems hwnd
    End If

    hSysMenu = GetSystemMenu(hwnd, 0)
    
    With mii
        .cbSize = Len(mii)
        .fMask = MIIM_STATE
        .fState = MFS_ENABLED Or IIf(GetProp(hwnd, "OnTop"), MFS_CHECKED, 0)
    End With
    
    ret = SetMenuItemInfo(hSysMenu, 1, 0, mii)
    
    mii.fState = IIf(GetProp(hwnd, "TransDisable") = 1, MFS_DISABLED, MFS_ENABLED) Or IIf(GetProp(hwnd, "Transparent"), MFS_CHECKED, 0)
    ret = SetMenuItemInfo(hSysMenu, 2, 0, mii)
    
    WindowProc = 0
    Exit Function
    
Case WM_SYSCOMMAND

    Select Case wParam
    Case 1
    
        onTop = Not CBool(GetProp(hwnd, "OnTop"))
        
        SetProp hwnd, "OnTop", onTop
        SetWindowPos hwnd, IIf(onTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        
        WindowProc = 0
        Exit Function
        
    Case 2
    
        If GetProp(hwnd, "TransDisable") = 1 Or Script.EnableTransparent = False Then
            WindowProc = 0
            Exit Function
        End If
    
        onTop = Not CBool(GetProp(hwnd, "Transparent"))
             
        ret = GetWindowLong(hwnd, GWL_EXSTYLE)
        SetProp hwnd, "Transparent", onTop
        
        If onTop Then
            ret = ret Or WS_EX_LAYERED
        Else
            ret = GetProp(hwnd, "OldStyle")
        End If
        
        SetWindowLong hwnd, GWL_EXSTYLE, ret
        SetLayeredWindowAttributes hwnd, 0, IIf(onTop, Script.TransparentKey, 255), LWA_ALPHA
        
        WindowProc = 0
        Exit Function
    End Select

Case WM_YOURTIME  ' Is outwar window

    WindowProc = WM_YOURTIME
    Exit Function

Case WM_YOURTIME + 1 ' Run command

    Dim Buff As String

    CopyMemory ByVal VarPtr(Buff), ByVal lParam, 4
    Script.Run Buff
    ZeroMemory ByVal VarPtr(Buff), 4
    
    WindowProc = 0
    Exit Function
    
Case WM_YOURTIME + 2 ' Allocate memory

    WindowProc = VirtualAlloc(ByVal 0&, wParam, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Exit Function
    
Case WM_YOURTIME + 3 ' Free memory

    VirtualFree wParam, lParam, MEM_DECOMMIT
    WindowProc = VirtualFree(wParam, lParam, MEM_RELEASE)
    Exit Function

Case WM_DESTROY, WM_NCDESTROY
    UnHookForm hwnd

End Select

WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)

End Function

Public Sub AddMenuItems(hwnd As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modHooking.AddMenuItems(hWnd)", Array(hwnd), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim hSysMenu As Long
Dim count As Long
Dim mii As MENUITEMINFO
Dim retval As Long

hSysMenu = GetSystemMenu(hwnd, 0)
count = GetMenuItemCount(hSysMenu)

With mii
    .cbSize = Len(mii)
    .fMask = MIIM_ID Or MIIM_TYPE
    .fType = MFT_SEPARATOR
    .wID = 0
End With

retval = InsertMenuItem(hSysMenu, count, 1, mii)

With mii
    .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
    .fType = MFT_STRING
    .fState = MFS_ENABLED
    .wID = 1
    .dwTypeData = "&Always On Top"
    .cch = Len(.dwTypeData)
End With

retval = InsertMenuItem(hSysMenu, count + 1, 1, mii)

' Insert transparent menu item
mii.wID = 2
mii.dwTypeData = "&Transparent"

retval = InsertMenuItem(hSysMenu, count + 2, 1, mii)

SetProp hwnd, "OnTop", GetProp(hwnd, "DeTop")
SetProp hwnd, "Transparent", Script.TransparentDefault
SetProp hwnd, "ItemAdded", 1

End Sub
