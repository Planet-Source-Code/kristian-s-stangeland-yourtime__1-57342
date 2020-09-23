VERSION 5.00
Begin VB.Form frmShowError 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Error Report"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   8460
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   564
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNoMore 
      Caption         =   "Ikke vis denne feilmeldingen igjen"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton cmdSendError 
      Caption         =   "&Send rapport"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtError 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8460
   End
End
Attribute VB_Name = "frmShowError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
'                               HuntERR
'                   Error Handling and Reporting Library
'                    from URFIN JUS (www.urfinjus.net)
'                  Copyright 2001-2002. All rights reserved.
'version 3.1, 04/25/2002
'Simple window to show error report
'=========================================================================================

Option Explicit

Public Property Let ErrorReport(ByVal AReport As String)
    txtError.Text = AReport
    If Not Visible Then ShowSelf
    On Error Resume Next
    Me.SetFocus
End Property

'Try to show non-modal. If there is already modal window in application,
'then it will fail, and we'll show as modal.
Private Sub ShowSelf()
    On Error Resume Next
    Show
    If Err.Number <> 0 Then Show vbModal
End Sub

Private Sub cmdOK_Click()

If chkNoMore.Value = 1 Then
    ErrAddToList
    chkNoMore.Value = 0
End If

Me.Hide
    
End Sub

Private Sub cmdSendError_Click()

On Error Resume Next

frmErrorInfo.Tag = ""
frmErrorInfo.Show

Do Until frmErrorInfo.Tag <> ""
    Sleep 10
    DoEvents
Loop

If frmErrorInfo.Tag = "success" Then
    Script.SendError txtError.Text, frmErrorInfo.txtEmail, frmErrorInfo.txtName, frmErrorInfo.txtEvents
End If

End Sub

Private Sub Form_Activate()

If Script.UseHooking = False Then
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End If

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

End Sub

Private Sub Form_Resize()

txtError.Move 0, 0, ScaleWidth, ScaleHeight - cmdSendError.Height - 16
cmdSendError.Left = ScaleWidth - cmdSendError.Width - 8
cmdSendError.Top = ScaleHeight - cmdSendError.Height - 8
cmdOK.Left = cmdSendError.Left - cmdOK.Width - 8
cmdOK.Top = cmdSendError.Top
chkNoMore.Top = cmdSendError.Top + 8

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hwnd

End Sub
