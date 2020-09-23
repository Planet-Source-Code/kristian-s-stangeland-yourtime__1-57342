VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
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

Dim lpArray() As Remember
Dim lCurrentRecord As Long

Public Property Get Fields(ByVal Item) As Variant

Select Case LCase(Item)
Case "enabled", 0: Fields = lpArray(lCurrentRecord).Enabled
Case "text", 1: Fields = lpArray(lCurrentRecord).Text
Case "extradata", 2: Fields = lpArray(lCurrentRecord).ExtraData
Case "exlong", 3: Fields = lpArray(lCurrentRecord).ExLong
Case "remdate", 4: Fields = lpArray(lCurrentRecord).RemDate
End Select

End Property

Public Property Let Fields(ByVal Item, ByVal vNewValue As Variant)

Select Case LCase(Item)
Case "enabled", 0: lpArray(lCurrentRecord).Enabled = vNewValue
Case "text", 1:  lpArray(lCurrentRecord).Text = vNewValue
Case "extradata", 2:  lpArray(lCurrentRecord).ExtraData = vNewValue
Case "exlong", 3: lpArray(lCurrentRecord).ExLong = vNewValue
Case "remdate", 4: lpArray(lCurrentRecord).RemDate = vNewValue
End Select

End Property

Public Property Get Name() As String
    Name = "Recordset"
End Property

Public Property Get BOF() As Boolean
    BOF = CBool(lCurrentRecord < 0)
End Property

Public Property Get EOF() As Boolean
    EOF = CBool(lCurrentRecord > SafeUBound(VarPtrArray(lpArray)))
End Property

Private Sub Class_Terminate()
    CopyMemory ByVal VarPtrArray(lpArray), ByVal 0&, 4
End Sub

Public Function AddNew(Optional sData As Variant)

Dim lIndex As Long, Tell&

lIndex = SafeUBound(VarPtrArray(lpArray)) + 1
ReDim Preserve lpArray(lIndex)

If Not IsMissing(sData) Then

    Select Case varType(sData)
    Case vbArray
    
        For Tell = LBound(sData) To UBound(sData)
        
            Select Case Tell - LBound(sData)
            Case 1: lpArray(lIndex).Text = sData(Tell)
            Case 2: lpArray(lIndex).ExtraData = sData(Tell)
            Case 3: lpArray(lIndex).ExLong = sData(Tell)
            Case 4: lpArray(lIndex).RemDate = sData(Tell)
            End Select
        
        Next
    
    Case Else
    
        lpArray(lIndex).Text = sData
    
    End Select

End If

lpArray(lIndex).Enabled = True

End Function

Public Sub Find(Criteria As String, Optional SearchDirection As Long, Optional Start)

Dim Tell As Long, strCode As String, objScriptEn As Object

' Load the script object
Set objScriptEn = CreateObject("MSScriptControl.ScriptControl")

strCode = "Function Eval(Enabled, Text, ExtraData, ExLong, RemDate)" & vbNewLine
strCode = strCode & "Eval = CBool(" & Criteria & ")" & vbNewLine
strCode = strCode & "End Function"

objScriptEn.Language = "VBScript"
objScriptEn.AddCode strCode

For Tell = Start To SafeUBound(VarPtrArray(lpArray)) Step SearchDirection

    If objScriptEn.Run("Eval", Array(lpArray(Tell).Enabled, lpArray(Tell).Text, lpArray(Tell).ExtraData, lpArray(Tell).ExLong, lpArray(Tell).RemDate)) Then
        lCurrentRecord = Tell
        Exit Sub
    End If

Next

End Sub

Public Function Clone() As Object
    
Set Clone = New clsRecordset
    
End Function

Public Sub MoveFirst()

lCurrentRecord = 0

End Sub

Public Sub MoveNext()

lCurrentRecord = lCurrentRecord + 1

End Sub

Public Sub MovePrevious()

lCurrentRecord = lCurrentRecord - 1

End Sub

Public Sub MoveLast()

lCurrentRecord = SafeUBound(VarPtrArray(lpArray))

End Sub

Public Sub Delete()

Dim lTell As Long, lCnt As Long

lCnt = SafeUBound(VarPtrArray(lpArray)) - 1

For lTell = lCurrentRecord To lCnt
    LSet lpArray(lTell) = lpArray(lTell + 1)
Next

ReDim Preserve lpArray(lCnt)

End Sub

Friend Sub Connection(destArray() As Remember)

CopyMemory ByVal VarPtrArray(lpArray), ByVal VarPtrArray(destArray), 4

End Sub
