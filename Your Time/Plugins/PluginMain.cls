VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PluginMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The plug-ins name. This is the string that will identify your class for the application and other plugins.
Public Property Get Name() As String
    Name = "Test"
End Property

' Gets the description of the plug-in
Public Property Get Description() As String
    Description = "Just a test plugin to show how to make such things"
End Property

' If this plug-in is configurable or not
Public Property Get Configurable() As Boolean
    Configurable = True
End Property

' Is invoked at start. gObjects is a collection of the global classes
Public Sub Initialize(gObjects As Collection)

Dim Script As Object

' Retrive the script-object
Set Script = gObjects("Script")

If MsgBox("This is the test plugin. Do you want to terminate the program now?", vbQuestion + vbYesNo, "The plugin") = vbYes Then
    ' Call the quit procedure in the script-object
    Script.Quit
End If

End Sub

' Called when the user has asked to configure the plug-in
Public Sub Configure()
    MsgBox "No yet implemented"
End Sub

' Used for the plug-ins to communticate with each other
Public Function Message(uMsgCode As Long, wData As Long, lData As Variant)
    ' I'm not using it for anything, but other might do it.
End Function

'Not always necessary, but can be useful when other plug-ins or the application wants to clone the object.
Public Function Clone() As Object
    ' As stated above, this is not necessary.
    Set Clone = New PluginMain
End Function

