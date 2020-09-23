VERSION 5.00
Begin VB.UserControl SheetTabs 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
   ToolboxBitmap   =   "SheetTabs.ctx":0000
End
Attribute VB_Name = "SheetTabs"
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

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Event Click()
Event TabChanging()
Event TabChanged(LastIndex As Long)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Type COORD
    X As Long
    Y As Long
End Type

Private Type SheetTab
    Text As String
    PositionX As Long
    Width As Long
End Type

Const ALTERNATE = 1
Const WINDING = 2

Dim lTabCount As Long
Dim lTabPos As Long
Dim lTabs() As SheetTab
Dim lTabsWidth As Long
Dim lExtraWidth As Long
Dim lBorderColor As Long
Dim lTabBackColor As Long
Dim lSelectedColor As Long
Dim lSelected As Long
Dim lBackColor As Long

Public Property Get ExtraWidth() As Long
    ExtraWidth = lExtraWidth
End Property

Public Property Let ExtraWidth(ByVal vNewValue As Long)
    lExtraWidth = vNewValue
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal vNewValue As Boolean)
    UserControl.FontUnderline = vNewValue
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal vNewValue As Boolean)
    UserControl.FontStrikethru = vNewValue
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal vNewValue As Boolean)
    UserControl.FontItalic = vNewValue
End Property

Public Property Get FontBold() As Boolean
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
    UserControl.FontBold = vNewValue
End Property

Public Property Get FontName() As String
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal vNewValue As String)
    UserControl.FontName = vNewValue
End Property

Public Property Get FontSize() As Long
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal vNewValue As Long)
    UserControl.FontSize = vNewValue
End Property

Public Property Get TabCount() As Variant
    TabCount = lTabCount
End Property

Public Property Get Selected() As Long
    Selected = lSelected
End Property

Public Property Let Selected(ByVal vNewValue As Long)
    
    Dim Tmp&
    
    If vNewValue < 0 Or vNewValue > lTabCount Then
        Exit Property
    End If
    
    If lSelected = vNewValue Then
        Exit Property
    End If
    
    RaiseEvent TabChanging
    
    Tmp = lSelected
    lSelected = vNewValue
    
    RaiseEvent TabChanged(Tmp)
    Redraw

End Property

Public Property Get Tabs(ByVal Index As Long) As String
    If Index > SafeUBound(VarPtrArray(lTabs)) Then Exit Property
    Tabs = lTabs(Index).Text
End Property

Public Property Let Tabs(ByVal Index As Long, ByVal vNewValue As String)
    lTabs(Index).Text = vNewValue
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = lBorderColor
End Property

Public Property Let BorderColor(ByVal vNewValue As OLE_COLOR)
    lBorderColor = vNewValue
End Property

Public Property Get SelectedColor() As OLE_COLOR
    SelectedColor = lSelectedColor
End Property

Public Property Let SelectedColor(ByVal vNewValue As OLE_COLOR)
    lSelectedColor = vNewValue
End Property

Public Property Get TabBackColor() As OLE_COLOR
    TabBackColor = lTabBackColor
End Property

Public Property Let TabBackColor(ByVal vNewValue As OLE_COLOR)
    lTabBackColor = vNewValue
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    UserControl.BackColor = vNewValue
End Property

Public Sub AddTab(Text As String)

lTabCount = lTabCount + 1

ReDim Preserve lTabs(lTabCount)
lTabs(lTabCount).Text = Text

End Sub

Public Sub RemoveTab(Index As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "SheetTabs.RemoveTab(Index)", Array(Index), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

lTabCount = lTabCount - 1

For Tell = Index To lTabCount
    lTabs(Tell).Text = lTabs(Tell + 1).Text
Next

ReDim Preserve lTabs(lTabCount)

End Sub

Private Sub DrawPolygon(X As Long, Width As Long, lColor As Long, Text As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "SheetTabs.DrawPolygon(X, Width, lColor, Text)", Array(X, Width, lColor, Text), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim NumCoords As Long, hBrush As Long, hRgn As Long, poly(1 To 4) As COORD

NumCoords = 4

poly(1).X = X
poly(1).Y = 2
poly(2).X = X + 7
poly(2).Y = UserControl.ScaleHeight - 1
poly(3).X = X + Width - 7
poly(3).Y = UserControl.ScaleHeight - 1
poly(4).X = X + Width
poly(4).Y = 2

hBrush = CreateSolidBrush(lColor)
hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
If hRgn Then FillRgn UserControl.hdc, hRgn, hBrush

DeleteObject hRgn
DeleteObject hBrush

Polygon UserControl.hdc, poly(1), NumCoords

UserControl.CurrentX = X + 7
UserControl.CurrentY = (UserControl.ScaleHeight / 2) - (UserControl.textHeight(Text) / 2)
UserControl.Print Text

End Sub

Public Sub Redraw()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "SheetTabs.Redraw", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, X&

If lTabCount < 0 Then Exit Sub

UserControl.Cls
UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), vbButtonShadow
UserControl.Line (0, 1)-(UserControl.ScaleWidth, 1), vbWhite

For Tell = 0 To lTabCount
    lTabs(Tell).PositionX = X
    lTabs(Tell).Width = UserControl.TextWidth(lTabs(Tell).Text) + lExtraWidth

    X = X + lTabs(Tell).Width - 4
Next

For Tell = lTabCount To 0 Step -1
    If Tell <> lSelected Then
        DrawPolygon lTabs(Tell).PositionX + ((UserControl.ScaleWidth - X) / 2), lTabs(Tell).Width, lTabBackColor, lTabs(Tell).Text
    End If
Next

DrawPolygon lTabs(lSelected).PositionX + ((UserControl.ScaleWidth - X) / 2), lTabs(lSelected).Width, lSelectedColor, lTabs(lSelected).Text
lTabsWidth = X

End Sub

Private Sub UserControl_Initialize()

UserControl.ForeColor = vbButtonText
lTabCount = -1

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

RaiseEvent KeyDown(KeyCode, Shift)

Select Case KeyCode
Case vbKeyLeft
    Selected = lSelected - 1
Case vbKeyRight
    Selected = lSelected + 1
End Select

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "SheetTabs.MouseDown(Button, Shift, X, Y)", Array(Button, Shift, X, Y), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

RaiseEvent MouseDown(Button, Shift, X, Y)

For Tell = 0 To lTabCount
    lTabPos = lTabs(Tell).PositionX + ((UserControl.ScaleWidth - lTabsWidth) / 2)
    
    If X > lTabPos And X < lTabPos + lTabs(Tell).Width Then
        Selected = Tell
        Exit Sub
    End If
Next

RaiseEvent Click

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next
lBorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
lTabBackColor = PropBag.ReadProperty("TabBackColor", vbCyan)
lSelectedColor = PropBag.ReadProperty("SelectedColor", vbYellow)
lExtraWidth = PropBag.ReadProperty("ExtraWidth", 32)
lSelected = PropBag.ReadProperty("Selected", 0)

UserControl.FontSize = PropBag.ReadProperty("FontSize", 10)
UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", False)
UserControl.FontItalic = PropBag.ReadProperty("FontItalic", False)
UserControl.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
UserControl.FontBold = PropBag.ReadProperty("FontBold", False)

End Sub

Private Sub UserControl_Resize()

Redraw

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

On Error Resume Next
PropBag.WriteProperty "BorderColor", lBorderColor, vbBlack
PropBag.WriteProperty "TabBackColor", lTabBackColor, vbCyan
PropBag.WriteProperty "SelectedColor", lSelectedColor, vbYellow
PropBag.WriteProperty "ExtraWidth", lExtraWidth, 32
PropBag.WriteProperty "Selected", lSelected, 0

PropBag.WriteProperty "FontSize", UserControl.FontSize, 10
PropBag.WriteProperty "FontStrikethru", UserControl.FontStrikethru, False
PropBag.WriteProperty "FontUnderline", UserControl.FontUnderline, False
PropBag.WriteProperty "FontItalic", UserControl.FontItalic, False
PropBag.WriteProperty "FontName", UserControl.FontName, "MS Sans Serif"
PropBag.WriteProperty "FontBold", UserControl.FontBold, False

End Sub
