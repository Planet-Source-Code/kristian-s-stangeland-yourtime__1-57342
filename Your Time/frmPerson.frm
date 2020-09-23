VERSION 5.00
Begin VB.Form frmPerson 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7245
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Lagre"
      Height          =   375
      Left            =   5640
      TabIndex        =   35
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Ny post"
      Height          =   375
      Left            =   5640
      TabIndex        =   39
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame frameDatabase 
      Caption         =   "Database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   32
      Top             =   1920
      Width           =   1455
      Begin VB.OptionButton optDatabase 
         Caption         =   "&Egen"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optDatabase 
         Caption         =   "&Felles"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame frameVisible 
      Caption         =   "&Synlig num:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5640
      TabIndex        =   25
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optVisibleNum 
         Caption         =   "Mob"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optVisibleNum 
         Caption         =   "Fax"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optVisibleNum 
         Caption         =   "Firma"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optVisibleNum 
         Caption         =   "Privat"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame frameExtra 
      Caption         =   "&Ekstra informasjon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   23
      Top             =   5160
      Width           =   6975
      Begin VB.TextBox txtExtra 
         Height          =   1245
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame framePerson 
      Caption         =   "&Person/Firma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtMobNum 
         Height          =   285
         Left            =   1800
         TabIndex        =   31
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox txtFaxNum 
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   3720
         Width           =   3375
      End
      Begin VB.TextBox txtFirmNum 
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtPrivateNum 
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txtHomepage 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtCountry 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtPostnum 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtFirm 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblMobNum 
         Caption         =   "&Mobilnummer:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label lblAge 
         Caption         =   "&Fødselsdag:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lblFaxNum 
         Caption         =   "Fa&xnummer:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label lblPrivateNum 
         Caption         =   "Pri&vatnummer:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblFirmNum 
         Caption         =   "F&irmanummer:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblHomepage 
         Caption         =   "&Hjemmeside:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblEmail 
         Caption         =   "&E-mail:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblCountry 
         Caption         =   "&Land:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblPostcity 
         Caption         =   "&Postnum/By:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblAddress 
         Caption         =   "A&ddresse:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblFrim 
         Caption         =   "&Firma:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblName 
         Caption         =   "N&avn:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Neste post"
      Height          =   375
      Left            =   5640
      TabIndex        =   36
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdPast 
      Caption         =   "Forrige post"
      Height          =   375
      Left            =   5640
      TabIndex        =   37
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Slett post"
      Height          =   375
      Left            =   5640
      TabIndex        =   38
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim CurrPost As Long

Public Sub Update()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".Update", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim TmpData As People, Tmp&

cmdDelete.Enabled = CBool(CurrPost >= 0 Or CurrPost <= IIf(optDatabase(0).Value, Script.PeopleCount, Script.UserPeopleCount))
cmdPast.Enabled = CBool(CurrPost > 0)
cmdNext.Enabled = CBool(CurrPost < IIf(optDatabase(0).Value, Script.PeopleCount, Script.UserPeopleCount))

Select Case optDatabase(0).Value
Case True

    If Script.PeopleCount < 0 Then
        cmdSave.Enabled = False
        Script.SetControls Me, "txt", "Enabled", False
        Exit Sub
    End If

    CopyRecord Peoples, TmpData, CurrPost
Case False

    If Script.UserPeopleCount < 0 Then
        cmdSave.Enabled = False
        Script.SetControls Me, "txt", "Enabled", True
        Exit Sub
    End If

    CopyRecord Users(UserID).DataPeoples, TmpData, CurrPost
End Select

If TmpData.Enabled = False Then

    If CurrPost <> 0 Then
        CurrPost = 0
        Update
        Exit Sub
    Else
        Script.SetControls Me, "txt", "Enabled", False
        Exit Sub
    End If
    
End If

cmdSave.Enabled = True
Script.SetControls Me, "txt", "Enabled", True

txtName.Text = TmpData.Name
txtAge.Text = TmpData.Birthday
txtFirm.Text = TmpData.Firm
txtAddress.Text = TmpData.Address
txtPostnum.Text = TmpData.PostCity
txtCountry.Text = TmpData.Country
txtEmail.Text = TmpData.Email
txtHomepage.Text = TmpData.Homepage
txtPrivateNum.Text = TmpData.PhoneNum
txtFirmNum.Text = TmpData.FirmNum
txtFaxNum.Text = TmpData.Fax
txtMobNum.Text = TmpData.MobNum
txtExtra.Text = TmpData.Information

For Tmp = 0 To optVisibleNum.count - 1
    optVisibleNum(Tmp).Value = False
Next

optVisibleNum(TmpData.VisibleNum).Value = True

End Sub

Private Sub cmdNew_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdNew_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Select Case optDatabase(0).Value
Case True
    CurrPost = AddPerson(Peoples)
Case False
    CurrPost = AddPerson(Users(UserID).DataPeoples)
End Select

Update

End Sub

Private Sub cmdNext_Click()

CurrPost = CurrPost + 1
Update

End Sub

Private Sub cmdPast_Click()

CurrPost = CurrPost - 1
Update

End Sub

Public Sub SaveCurrent()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".SaveCurrent", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim TmpData As People, Tmp&

If CurrPost < 0 Or CurrPost > IIf(optDatabase(0).Value, Script.PeopleCount, Script.UserPeopleCount) Then
    Exit Sub
End If

Users(UserID).Changed = True
TmpData.Enabled = True
TmpData.Name = txtName.Text
TmpData.Birthday = txtAge.Text
TmpData.Firm = txtFirm.Text
TmpData.Address = txtAddress.Text
TmpData.PostCity = txtPostnum.Text
TmpData.Country = txtCountry.Text
TmpData.Email = txtEmail.Text
TmpData.Homepage = txtHomepage.Text
TmpData.PhoneNum = Val(txtPrivateNum.Text)
TmpData.FirmNum = Val(txtFirmNum.Text)
TmpData.Fax = Val(txtFaxNum.Text)
TmpData.MobNum = Val(txtMobNum.Text)
TmpData.Information = txtExtra.Text

For Tmp = 0 To optVisibleNum.count - 1
    If optVisibleNum(Tmp).Value Then
        TmpData.VisibleNum = Tmp
        Exit For
    End If
Next

Select Case optDatabase(0).Value
Case True
    LSet Peoples(CurrPost) = TmpData
Case False
    LSet Users(UserID).DataPeoples(CurrPost) = TmpData
End Select

End Sub

Private Sub cmdSave_Click()

SaveCurrent

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

SetNumber txtPrivateNum, True
SetNumber txtFirmNum, True
SetNumber txtMobNum, True
SetNumber txtFaxNum, True
Update

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hWnd

End Sub

Private Sub optDatabase_Click(Index As Integer)

Update

End Sub

Public Property Get CurrentPost() As Long
    CurrentPost = CurrPost
End Property

Public Property Let CurrentPost(ByVal vNewValue As Long)
    CurrPost = vNewValue
End Property
