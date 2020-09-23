VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
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

Dim bInvokeDialog As Boolean

Public Property Get InvokeDialog() As Boolean
    InvokeDialog = bInvokeDialog
End Property

Public Property Let InvokeDialog(ByVal vNewValue As Boolean)
    bInvokeDialog = vNewValue
End Property

Public Function Clone() As clsPrint

Set Clone = New clsPrint

Clone.InvokeDialog = Me.InvokeDialog

End Function

Private Sub Class_Initialize()

bInvokeDialog = True

End Sub

Public Sub PrintForm(Form As Object)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsPrint.PrintForm(Form)", Array(Form.Name), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

If bInvokeDialog Then
    Common.ShowPrinter Form
End If

Form.PrintForm

End Sub

Public Sub PrintHeader(lpObject As Object, strText As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsPrint.PrintHeader(lpObject)", Array(ObjPtr(lpObject)), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

lpObject.FontSize = 12
lpObject.CurrentY = 35
lpObject.CurrentX = (lpObject.ScaleWidth / 2) - (lpObject.TextWidth(strText) / 2)
lpObject.FontBold = True
lpObject.Print strText

End Sub

Public Sub PrintRemTasks(lStartDate As Date, lEndDate As Date, PrinterObject As Object)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsPrint.PrintRemTasks(lStartDate, lEndDate, PrinterObject)", Array(lStartDate, lEndDate, TypeName(PrinterObject)), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

PrinterObject.ScaleMode = vbPixels
PrintHeader PrinterObject, "Avtaler og oppgaver i tiden mellom " & lStartDate & " og " & lEndDate

' Skriv ut all data
DrawDatabase Users(UserID).DataRem, lStartDate, lEndDate, "Avtaler:", Printer
DrawDatabase Users(UserID).DataTasks, lStartDate, lEndDate, "Oppgaver:", Printer

' Kun hvis vi har med selve printer-objektet
If TypeName(PrinterObject) = "Printer" Then
    ' Fullfør jobben
    PrinterObject.EndDoc
End If

End Sub

Friend Sub DrawDatabase(Database() As Remember, lStartDate As Date, lEndDate As Date, strHeaderText As String, PrinterObject As Object)

Dim Tell As Long

If strHeaderText <> "" Then
    ' Tegn linjen
    PrinterObject.Line (25, PrinterObject.CurrentY + 10)-(PrinterObject.ScaleWidth - 25, PrinterObject.CurrentY + 10)
    PrinterObject.CurrentY = PrinterObject.CurrentY + 15
    PrinterObject.CurrentX = (PrinterObject.ScaleWidth / 2) - (PrinterObject.TextWidth(strHeaderText) / 2)
    PrinterObject.Print strHeaderText
    PrinterObject.CurrentY = PrinterObject.CurrentY + 15
End If

For Tell = LBound(Database) To UBound(Database)

    If Database(Tell).RemDate >= lStartDate And Database(Tell).RemDate <= lEndDate Then

        PrinterObject.CurrentX = 10
        PrinterObject.FontBold = False
        PrinterObject.Print "[" & Database(Tell).RemDate & "] " & Database(Tell).Text
    
    End If
    
Next

End Sub

Public Sub PrintDiary(Form As Object, lStartDate As Date, lEndDate As Date, PrinterObject As Object)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsPrint.PrintDiary(Form, lStartDate, lEndDate, PrinterObject)", Array(Form.Name, lStartDate, lEndDate, TypeName(PrinterObject)), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell As Long, r As Long
Dim LeftOffset As Long, TopOffset As Long
Dim LeftMargin As Long, TopMargin As Long
Dim RightMargin As Long, BottomMargin As Long
Dim fr As FormatRange
Dim rcDrawTo As Rect
Dim rcPage As Rect
Dim TextLength As Long
Dim NextCharPosition As Long, LastPos As Long

If bInvokeDialog Then
    Common.ShowPrinter Form
End If

PrinterObject.ScaleMode = vbPixels
PrintHeader PrinterObject, "Dagbok for perioden " & Date & " til " & Date + 5

' Get the offsett to the printable area on the page in twips
LeftOffset = GetDeviceCaps(PrinterObject.hdc, PHYSICALOFFSETX)
TopOffset = GetDeviceCaps(PrinterObject.hdc, PHYSICALOFFSETY)

' Calculate the Left, Top, Right, and Bottom margins
LeftMargin = LeftOffset - 25
TopMargin = TopOffset - 20
RightMargin = PrinterObject.Width - LeftOffset
BottomMargin = PrinterObject.Height - TopOffset

' Set printable area rect
rcPage.Left = 0
rcPage.Top = 0
rcPage.Right = PrinterObject.ScaleWidth
rcPage.Bottom = PrinterObject.ScaleHeight

' Set rect in which to print (relative to printable area)
rcDrawTo.Right = RightMargin
rcDrawTo.Bottom = BottomMargin
rcDrawTo.Left = LeftMargin

For Tell = LBound(Users(UserID).DataDiary) To UBound(Users(UserID).DataDiary)

    If Users(UserID).DataDiary(Tell).RemDate >= lStartDate And Users(UserID).DataDiary(Tell).RemDate <= lEndDate Then

        PrinterObject.Line (25, PrinterObject.CurrentY + 10)-(PrinterObject.ScaleWidth - 25, PrinterObject.CurrentY + 10)
        PrinterObject.CurrentY = PrinterObject.CurrentY + 15
        
        PrinterObject.CurrentX = 30
        PrinterObject.FontSize = 9
        PrinterObject.FontBold = True
        PrinterObject.Print "Dag - " & Date + Tell
        PrinterObject.FontBold = False
        
        PrinterObject.CurrentY = PrinterObject.CurrentY + 10
        
        ' Get length of text in RTF
        frmDiary.RichBuffer.Text = Users(UserID).DataDiary(Tell).Text
        TextLength = Len(frmDiary.RichBuffer.Text)

        ' Loop printing each page until done
        Do
           rcDrawTo.Top = PrinterObject.CurrentY
           
           ' Print the page by sending EM_FORMATRANGE message
           NextCharPosition = SendMessage(frmDiary.RichBuffer.hwnd, EM_FORMATRANGE, True, fr)
           If NextCharPosition >= TextLength Then Exit Do  'If done then exit
                  
           fr.chrg.cpMin = NextCharPosition
           PrinterObject.NewPage
           PrinterObject.Print Space(1)
           fr.hdc = PrinterObject.hdc
           fr.hdcTarget = PrinterObject.hdc
        Loop
        
        ' Allow the RTF to free up memory
        r = SendMessage(frmDiary.RichBuffer.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    End If

Next

' Kun hvis vi har med selve printer-objektet
If TypeName(PrinterObject) = "Printer" Then
    ' Fullfør jobben
    PrinterObject.EndDoc
End If

End Sub

Public Property Get Name() As String
    Name = "Print"
End Property
