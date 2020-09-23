VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

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

Private Type ObjectProperties
    Name As String
    Properties As Collection
End Type

Private Type FormProperties
    Name As String
    Caption As String
    ObjectCount As Integer
    ObjProp() As ObjectProperties
End Type

' Private data about the language
Private FormProperties() As FormProperties
Private lConstantIndex As Long

Public Property Get Constant(ByVal sName As String) As String

On Error Resume Next
Dim Tell&, Tmp&, Prop, aItems

If lConstantIndex < 0 Then
    Exit Property
End If

For Tell = 0 To FormProperties(lConstantIndex).ObjectCount

    For Each Prop In FormProperties(lConstantIndex).ObjProp(Tell).Properties
    
        aItems = Split(Prop, ":", 3)
    
        If aItems(0) = sName Then
            Constant = Script.ProcessVariables(aItems(2))
            Exit Property
        End If
    
    Next

Next

End Property

Public Property Get Name() As String
    Name = "Language"
End Property

Private Function ParseText(sText As String) As String

On Error Resume Next
Dim Tmp As Long

Tmp = InStr(1, sText, Chr(34))

If Tmp <> 0 Then
    ' It's a string. Return the text between the "'s
    ParseText = Mid(sText, Tmp + 1, InStr(Tmp + 1, sText, Chr(34)) - Tmp - 1)
Else
    ParseText = Script.ConvertToNumeric(sText)
End If

End Function

Public Sub LoadLanguagePack(strFile As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler:

    If Err.Number = 75 Or Err.Number = 76 Then ' File/Path Error
        ' The file does not exist
        Exit Sub
    End If
    
    ErrorIn "clsLanguage.LoadLanguagePack(strFile)", Array(strFile), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Free&, Tmp&, Tell&, sLine$, sProperty$, sObjectName$, Scode, sInnerLine, lIndex&

ClearLanguagePack
Free = FreeFile

Open strFile For Input As Free

    Do Until EOF(Free)

        Line Input #Free, sLine
    
        Select Case Mid(sLine, 1, 1)
        Case "["
    
            lIndex = lIndex + 1
    
            ReDim Preserve FormProperties(lIndex)
    
            FormProperties(lIndex).Name = Mid(sLine, 2, Len(sLine) - 2)
            FormProperties(lIndex).ObjectCount = -1
            
            If LCase(FormProperties(lIndex).Name) = "constants" Then
                lConstantIndex = lIndex
            End If
            
        Case ";"
            ' Jump over, since this is a comment
        
        Case Else ' See if anything should be done
        
            Scode = Split(sLine, "=")
        
            If UBound(Scode) > 0 Then
                If LCase(Mid(Scode(0), 1, 8)) = "caption " Then
            
                    FormProperties(lIndex).Caption = ParseText(CStr(Scode(1)))
            
                Else
                
                    sInnerLine = Split(Scode(0), ".")
                    sObjectName = Replace(sInnerLine(0), " ", "")
                
                    If UBound(sInnerLine) >= 0 Then
                        sProperty = "caption"
                    Else
                        sProperty = Replace(LCase(sInnerLine(1)), " ", "")
                    End If
                    
                    Tmp = -1
                    
                    ' Check if the object already exists
                    For Tell = 0 To FormProperties(lIndex).ObjectCount
                        If sObjectName = FormProperties(lIndex).ObjProp(Tell).Name Then
                            Tmp = Tell
                            Exit For
                        End If
                    Next
                    
                    If Tmp < 0 Then
                        FormProperties(lIndex).ObjectCount = FormProperties(lIndex).ObjectCount + 1
                        ReDim Preserve FormProperties(lIndex).ObjProp(FormProperties(lIndex).ObjectCount)
                        
                        Tmp = FormProperties(lIndex).ObjectCount
                    End If
                    
                    Set FormProperties(lIndex).ObjProp(Tmp).Properties = New Collection
                    
                    FormProperties(lIndex).ObjProp(Tmp).Name = sObjectName
                    FormProperties(lIndex).ObjProp(Tmp).Properties.Add sObjectName & ":" & sProperty & ":" & ParseText(CStr(Scode(1))), sObjectName
    
                End If
            End If
            
        End Select
    
    Loop
    
Close Free

End Sub

Public Function SetLanguageInForm(Form As Object) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn "clsLanguage.SetLanguageInForm(Form)", Array(Form.Name), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Tell&, lIndex&, Tmp&, aItems, Prop, lpObject As Object

For Tell = 0 To SafeUBound(VarPtrArray(FormProperties))
    If FormProperties(Tell).Name = Form.Name Then
        lIndex = Tell
        Exit For
    End If
Next

If lIndex > SafeUBound(VarPtrArray(FormProperties)) Then
    ' Couln't find form
    SetLanguageInForm = -1
    Exit Function
End If

Form.Caption = Script.ProcessVariables(FormProperties(lIndex).Caption)

For Tell = 0 To FormProperties(lIndex).ObjectCount

    For Each Prop In FormProperties(lIndex).ObjProp(Tell).Properties
        
        aItems = Split(Prop, ":", 3)
        Tmp = InStr(1, CStr(aItems(0)), "(")
        
        If Tmp <> 0 Then
            Set lpObject = Form.Controls(Mid(aItems(0), 1, Tmp - 1)).Item(Val(Script.ConvertToNumeric(Mid(CStr(aItems(0)), Tmp + 1))))
        Else
            Set lpObject = Form.Controls(FormProperties(lIndex).ObjProp(Tell).Name)
        End If
        
        CallByName lpObject, CStr(aItems(1)), VbLet, Script.ProcessVariables(CStr(aItems(2)))
    Next

Next

End Function

Public Function Clone() As clsLanguage

Set Clone = New clsLanguage

End Function

Public Function EnumLanguagePacks(ByVal sFolder As String, ByVal sExtension As String) As Collection

sFolder = ValidPath(sFolder)

If Dir(sFolder) = "" Then
    Exit Function
End If

Set EnumLanguagePacks = Script.GetFolderList(sFolder, sExtension)

End Function

Public Sub ClearLanguagePack()

Erase FormProperties

End Sub

Private Sub Class_Initialize()

lConstantIndex = -1

End Sub
