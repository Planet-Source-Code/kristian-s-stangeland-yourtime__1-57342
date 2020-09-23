Attribute VB_Name = "modUsers"
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

Type Remember
    Enabled As Boolean
    Text As String
    ExtraData As String
    ExLong As Long
    RemDate As Date
End Type

Type People
    Enabled As Boolean
    Birthday As Date
    Name As String
    PostCity As String
    PhoneNum As Double
    FirmNum As Double
    MobNum As Double
    Fax As Double
    Email As String
    Homepage As String
    Address As String
    Country As String
    Firm As String
    Information As String
    VisibleNum As Byte
End Type

Type User
    Created As Date
    Changed As Boolean
    Password As String
    LastPath As String
    LoggedOn As Date
    UserName As String
    DataRem() As Remember
    DataTasks() As Remember
    DataDiary() As Remember
    DataOwn() As Remember
    DataPeoples() As People
End Type

Const Databases = 5

Public StaticDays() As Remember
Public Peoples() As People
Public Users(1 To 20) As User
Public Index(1 To 20) As String

Public Function AddRecord(lpArray() As Remember, lDate As Date, Text As String, ExtraData As String, Optional ExLong As Long = -1) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.AddRecord(lpArray, lDate, Text, ExtraData)", Array(VarPtrArray(lpArray), lDate, Text, ExtraData), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&, Cnt&

Cnt = SafeUBound(VarPtrArray(lpArray))

For Tell = Abs(SafeLBound(VarPtrArray(lpArray))) To Cnt
    If lpArray(Tell).Enabled = False Or lpArray(Tell).Text = "" Then
        Users(UserID).Changed = True
        lpArray(Tell).Text = Text
        lpArray(Tell).RemDate = lDate
        lpArray(Tell).Enabled = True
        lpArray(Tell).ExtraData = ExtraData
        lpArray(Tell).ExLong = ExLong
        AddRecord = Tell
        Exit Function
    End If
Next

' I tilfelle ingen ledige elementer
ReDim Preserve lpArray(Cnt + 1)

Users(UserID).Changed = True
lpArray(Cnt + 1).Text = Text
lpArray(Cnt + 1).RemDate = lDate
lpArray(Cnt + 1).Enabled = True
lpArray(Cnt + 1).ExtraData = ExtraData
lpArray(Cnt + 1).ExLong = ExLong
AddRecord = Tell

End Function

Public Function FindPerson(strFirstLetter As String, SearchString As Long, Optional ByEntireSting As Boolean = False) As People()

On Error Resume Next
SearchInPeople strFirstLetter, FindPerson, Peoples, SearchString, ByEntireSting
SearchInPeople strFirstLetter, FindPerson, Users(UserID).DataPeoples, SearchString, ByEntireSting

End Function

Public Sub SearchInPeople(strFirstLetter As String, lpArray() As People, lpDestArray() As People, SearchString As Long, ByEntireSting As Boolean)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.SearchInPeople(strFirstLetter, lpArray, lpDestArray, SearchString, ByEntireString)", Array(strFirstLetter, VarPtrArray(lpArray), VarPtrArray(lpDestArray), SearchString, ByEntireSting), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, Tmp&, lCount&, lArrayCount&, strDest$

lCount = SafeUBound(VarPtrArray(lpDestArray))
lArrayCount = SafeUBound(VarPtrArray(lpArray))

For Tell = SafeLBound(VarPtrArray(lpDestArray)) To lCount

    If Tell < 0 Then
        Exit For
    End If

    Select Case SearchString
    Case 0: strDest = lpDestArray(Tell).Name
    Case 1: strDest = lpDestArray(Tell).Firm
    Case 2: strDest = lpDestArray(Tell).Address
    Case 3: strDest = Choose(lpDestArray(Tell).VisibleNum + 1, lpDestArray(Tell).PhoneNum, lpDestArray(Tell).FirmNum, lpDestArray(Tell).Fax, lpDestArray(Tell).MobNum)
    End Select

    If IIf(ByEntireSting, InStr(1, strDest, strFirstLetter) > 0, Mid(lpDestArray(Tell).Name, 1, 1) = strFirstLetter) Then

        lArrayCount = lArrayCount + 1
        
        ReDim Preserve lpArray(lArrayCount)
    
        LSet lpArray(lArrayCount) = lpDestArray(Tell)
    
    End If
    
Next

End Sub

Public Function Search(Database() As Remember, lDate As Variant, Flag As Long, Optional Start As Long = -1) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.Search(Database, lDate, Flag)", Array(VarPtrArray(Database), lDate, Flag), EA_NORERAISE: HandleError: Search = -1: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&, lStart&

If Start < 0 Then
    lStart = Abs(SafeLBound(VarPtrArray(Database)))
Else
    lStart = Start
End If

For Tell = lStart To SafeUBound(VarPtrArray(Database))

    If Database(Tell).Enabled = True Then

        Select Case Flag
        Case 0
            If lDate = Database(Tell).RemDate Then
                Search = Tell
                Exit Function
            End If
        
        Case 1
            If Script.IsDateEqual(lDate, Database(Tell).RemDate) Then
                Search = Tell
                Exit Function
            End If
        
        Case 2
            If Script.IsTimeEqual(lDate, Database(Tell).RemDate) Then
                Search = Tell
                Exit Function
            End If
            
        Case 3
            
            If Script.SimpleRegExp(lDate, Database(Tell).RemDate, Database(Tell).ExtraData) Then
                Search = Tell
                Exit Function
            End If
        
        Case 4
        
            If LCase(lDate) = LCase(Database(Tell).Text) Then
                Search = Tell
                Exit Function
            End If
        
        Case 5
        
            If InStr(1, Database(Tell).Text, lDate, vbTextCompare) <> 0 Or InStr(1, Database(Tell).ExtraData, lDate, vbTextCompare) <> 0 Then
                Search = Tell
                Exit Function
            End If
        
        End Select
    
    End If
Next

Search = -1

End Function

Public Sub CreateNewUser(Index As Long, Name As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.CreateNewUser(Index, Name)", Array(Index, Name), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Users(Index).Created = Now
Users(Index).UserName = Name
ReDim Users(Index).DataOwn(Script.NumOfRecords)

End Sub

Public Sub LoadData(Path As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.LoadData(Path)", Array(Path), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tmp&, Free&

If Dir(Path & "Index.dat") <> "" Then
    
    Free = FreeFile
    
    Open Path & "Index.dat" For Binary As Free
        Get #Free, , Index
    Close Free
    
    For Tmp = LBound(Index) To UBound(Index)
        If Index(Tmp) <> "" And Dir(Path & Index(Tmp)) <> "" Then
            LoadUser Users(Tmp), Path & Index(Tmp)
        End If
    Next
End If

If Script.NoUsers Then
    CreateNewUser 1, Script.DefaultUser
End If

If Users(UserID).UserName = "" Then
    UserID = 1
End If

If Dir(Path & "Peoples.dat") <> "" Then

    Free = FreeFile
    
    Open Path & "Peoples.dat" For Binary As Free
        Get #Free, , Tmp
        ReDim Peoples(Tmp)
        
        Get #Free, , Peoples
    Close Free
End If

End Sub

Public Sub SaveData(Path As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.SaveData(Path)", Array(Path), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tmp&, Free&, Index(1 To 20) As String, ExtraData() As Long
On Error Resume Next

If GetAttr(Path) <> vbDirectory Then
    MkDir Path
End If

If Dir(Path & "Index.dat") <> "" Then
    Kill Path & "Index.dat"
End If

ReDim ExtraData(Databases - 1)

For Tmp = 1 To 20
    If (Users(Tmp).LastPath <> Users(Tmp).UserName & ".dat" Or Users(Tmp).Changed = True) And Not Users(Tmp).UserName = "" Then
        If Dir(Users(Tmp).LastPath) <> "" And Users(Tmp).LastPath <> "" Then Kill Users(Tmp).LastPath
        
        Free = FreeFile
        Index(Tmp) = Users(Tmp).UserName & ".dat"
        Users(Tmp).LastPath = ""
        
        ExtraData(0) = SafeUBound(VarPtrArray(Users(Tmp).DataRem))
        ExtraData(1) = SafeUBound(VarPtrArray(Users(Tmp).DataTasks))
        ExtraData(2) = SafeUBound(VarPtrArray(Users(Tmp).DataDiary))
        ExtraData(3) = SafeUBound(VarPtrArray(Users(Tmp).DataOwn))
        ExtraData(4) = SafeUBound(VarPtrArray(Users(Tmp).DataPeoples))

        Open Path & Users(Tmp).UserName & ".dat" For Binary As Free
            Put #Free, , ExtraData
            Put #Free, , Users(Tmp)
        Close Free
        
        Index(Tmp) = Users(Tmp).UserName & ".dat"
    End If
Next

Free = FreeFile

Open Path & "Index.dat" For Binary As Free
    Put #Free, , Index
Close Free

Tmp = Script.PeopleCount

If Dir(Path & "Peoples.dat") <> "" Then Kill Path & "Peoples.dat"

If Tmp >= 0 Then
    Free = FreeFile

    Open Path & "Peoples.dat" For Binary As Free
        Put #1, , Tmp
        Put #1, , Peoples
    Close Free
End If

End Sub

Public Sub LoadUser(lpUser As User, FilePath As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.LoadUser(lpUser, FilePath)", Array(lpUser.UserName, FilePath), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim ExtraData() As Long, Free&

ReDim ExtraData(Databases - 1)
Free = FreeFile

Open FilePath For Binary As Free
    Get #Free, , ExtraData
    
    If ExtraData(0) >= 0 Then ReDim lpUser.DataRem(ExtraData(0))
    If ExtraData(1) >= 0 Then ReDim lpUser.DataTasks(ExtraData(1))
    If ExtraData(2) >= 0 Then ReDim lpUser.DataDiary(ExtraData(2))
    If ExtraData(3) >= 0 Then ReDim lpUser.DataOwn(ExtraData(3))
    If ExtraData(4) >= 0 Then ReDim lpUser.DataPeoples(ExtraData(4))
    
    Get #Free, , lpUser
Close Free

lpUser.LastPath = FilePath

End Sub

Public Function SaveRecord(lpDatabase() As Remember, Text As String, lDate As Date, Optional Expression As Variant, Optional ExtraData As String) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.SaveRecord(lpDatabase, Text, lData, [Expression], [ExtraData])", Array(VarPtrArray(lpDatabase), Text, lDate, Expression, ExtraData), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Num&

If IsMissing(Expression) Then
    Expression = Text
End If

Num = Search(lpDatabase, lDate, 0)

If Num >= 0 Then
    Users(UserID).Changed = True
    lpDatabase(Num).Enabled = Not CBool(Expression = "")
    lpDatabase(Num).Text = Text
    lpDatabase(Num).RemDate = lDate
    lpDatabase(Num).ExtraData = ExtraData
Else
    If Expression <> "" Then
        AddRecord lpDatabase, lDate, Text, ExtraData
    End If
End If

SaveRecord = Num

End Function

Public Sub CopyDatabase(lpDest() As Remember, lpSource() As Remember)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.CopyDatabase(lpDest, lpSource)", Array(VarPtrArray(lpDest), VarPtrArray(lpSource)), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

ReDim lpDest(LBound(lpSource) To UBound(lpSource))

For Tell = LBound(lpDest) To UBound(lpDest)
    lpDest(Tell).Enabled = lpSource(Tell).Enabled
    lpDest(Tell).ExtraData = lpSource(Tell).ExtraData
    lpDest(Tell).RemDate = lpSource(Tell).RemDate
    lpDest(Tell).Text = lpSource(Tell).Text
Next

End Sub

Public Function AddPerson(lpDest() As People) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.AddPerson(lpDest)", Array(VarPtrArray(lpDest)), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell As Long

For Tell = Abs(SafeLBound(VarPtrArray(lpDest))) To SafeUBound(VarPtrArray(lpDest))
    If lpDest(Tell).Enabled = False Or (lpDest(Tell).Name = "" And lpDest(Tell).Firm = "") Then
        Users(UserID).Changed = True
        lpDest(Tell).Enabled = True
        AddPerson = Tell
        Exit Function
    End If
Next

ReDim Preserve lpDest(SafeUBound(VarPtrArray(lpDest)) + 1)

Users(UserID).Changed = True
lpDest(UBound(lpDest)).Enabled = True
AddPerson = UBound(lpDest)

End Function

Public Sub CopyRecord(lpSource() As People, lpRecord As People, ByVal Index As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "modUsers.CopyRecord(lpSource, lpRecord, Index)", Array(VarPtrArray(lpSource), lpRecord.Name, Index), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

LSet lpRecord = lpSource(Index)

End Sub
