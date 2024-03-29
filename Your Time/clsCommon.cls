VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommon"
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

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long

Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As Rect
    rtMargin As Rect
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type LOGFONT
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
        lfFaceName As String * 31
End Type

Private Type ChooseFont
        lStructSize As Long
        hwndOwner As Long
        hdc As Long
        lpLogFont As Long
        iPointSize As Long
        flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
End Type

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Dim OFName As OPENFILENAME
Dim CustomColors(15) As Byte

Public Function Clone() As clsCommon

Set Clone = New clsCommon

End Function

Public Function ShowColor(hwnd As Long) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsCommon.ShowColor(hWnd)", Array(hwnd), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    cc.lStructSize = Len(cc)
    cc.hwndOwner = hwnd
    cc.hInstance = App.hInstance
    cc.lpCustColors = VarPtr(CustomColors(0))
    cc.flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CopyMemory CustomColors(0), ByVal cc.lpCustColors, 16
    Else
        ShowColor = -1
    End If
    
End Function

Public Function ShowOpen(hwnd As Long) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsCommon.ShowOpen(hWnd)", Array(hwnd), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hwndOwner = hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the dialog title
    OFName.lpstrTitle = "Open File"
    'no extra flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
    
End Function

Public Function ShowFont(hwnd As Long) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsCommon.ShowFont(hWnd)", Array(hwnd), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

    Dim cf As ChooseFont, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim FontName As String, retval As Long
    lfont.lfHeight = 0  ' determine default height
    lfont.lfWidth = 0  ' determine default width
    lfont.lfEscapement = 0  ' angle between baseline and escapement vector
    lfont.lfOrientation = 0  ' angle between baseline and orientation vector
    lfont.lfWeight = FW_NORMAL  ' normal weight i.e. not bold
    lfont.lfCharSet = DEFAULT_CHARSET  ' use default character set
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
    lfont.lfQuality = DEFAULT_QUALITY  ' default quality setting
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN  ' default pitch, proportional with serifs
    lfont.lfFaceName = "Times New Roman" & vbNullChar  ' string must be null-terminated
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)  ' lock and get pointer
    CopyMemory ByVal pMem, lfont, Len(lfont)  ' copy structure's contents into block
    ' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
    cf.lStructSize = Len(cf)  ' size of structure
    cf.hwndOwner = hwnd
    cf.hdc = Printer.hdc  ' device context of default printer (using VB's mechanism)
    cf.lpLogFont = pMem   ' pointer to LOGFONT memory block buffer
    cf.iPointSize = 120  ' 12 point font (in units of 1/10 point)
    cf.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    cf.rgbColors = RGB(0, 0, 0)  ' black
    cf.nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
    cf.nSizeMin = 10  ' minimum point size
    cf.nSizeMax = 72  ' maximum point size
    ' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
    ' and then print out the attributes we mentioned earlier that the user selected.
    retval = ChooseFont(cf)  ' open the dialog box
    If retval <> 0 Then  ' success
        CopyMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
        ' Now make the fixed-length string holding the font name into a "normal" string.
        ShowFont = Left$(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        Debug.Print  ' end the line
    End If
    ' Deallocate the memory block we created earlier.  Note that this must
    ' be done whether the function succeeded or not.
    retval = GlobalUnlock(hMem)  ' destroy pointer, unlock block
    retval = GlobalFree(hMem)  ' free the allocated memory
    
End Function

Public Function ShowSave(hwnd) As String

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsCommon.ShowSave(hWnd)", Array(hwnd), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hwndOwner = hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the dialog title
    OFName.lpstrTitle = "Save File"
    'no extra flags
    OFName.flags = 0

    'Show the 'Save File'-dialog
    If GetSaveFileName(OFName) Then
        ShowSave = Trim$(OFName.lpstrFile)
    Else
        ShowSave = ""
    End If
    
End Function

Public Function ShowPageSetupDlg(hwnd As Long) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsCommon.ShowPageSetupDlg(hWnd)", Array(hwnd), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

    Dim m_PSD As PAGESETUPDLG
    'Set the structure size
    m_PSD.lStructSize = Len(m_PSD)
    'Set the owner window
    m_PSD.hwndOwner = hwnd
    'Set the application instance
    m_PSD.hInstance = App.hInstance
    'no extra flags
    m_PSD.flags = 0

    'Show the pagesetup dialog
    If PAGESETUPDLG(m_PSD) Then
        ShowPageSetupDlg = 0
    Else
        ShowPageSetupDlg = -1
    End If
    
End Function

Public Sub ShowPrinter(frmOwner As Object, Optional PrintFlags As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsCommon.ShowPrinter(frmOwner, [PrintFlags])", Array(frmOwner.Name, PrintFlags), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

    '-> Code by Donald Grover
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = frmOwner.hwnd

    PrintDlg.flags = PrintFlags
    On Error Resume Next
    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left$(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
            Next
        End If

        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        On Error GoTo 0
    End If

End Sub

Public Property Get Name() As String
    Name = "Common"
End Property
