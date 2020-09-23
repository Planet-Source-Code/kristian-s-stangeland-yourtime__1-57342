VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Arternativer"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5355
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   518
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUse 
      Caption         =   "&Bruk"
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "Generelt"
      TabPicture(0)   =   "frmSettings.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frameCalender"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameTimeZone"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameTags"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frameLanguage"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Brukere"
      TabPicture(1)   =   "frmSettings.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameUsers"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameLoad"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Avansert"
      TabPicture(2)   =   "frmSettings.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameAdvance"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frameScript"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Plugins"
      TabPicture(3)   =   "frmSettings.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "frameAdmin"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "framePluginSet"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame frameLanguage 
         Caption         =   "Språk:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74760
         TabIndex        =   56
         Top             =   5640
         Width           =   4575
         Begin VB.ComboBox cmdLanguagePack 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblChooseLanguage 
            Caption         =   "Velg språkpakke:"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame framePluginSet 
         Caption         =   "Plugins:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   54
         Top             =   3960
         Width           =   4575
         Begin VB.CheckBox chkFindPlugins 
            Caption         =   "Automatisk søk etter plugins ved oppstart"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   480
            Width           =   3495
         End
      End
      Begin VB.Frame frameAdmin 
         Caption         =   "Administrer plugins:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   240
         TabIndex        =   51
         Top             =   600
         Width           =   4575
         Begin VB.CommandButton cmdConfigure 
            Caption         =   "&Konfigurer"
            Height          =   375
            Left            =   240
            TabIndex        =   52
            Top             =   2520
            Width           =   1455
         End
         Begin MSComctlLib.ListView lstPlugins 
            Height          =   2055
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   3625
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame frameLoad 
         Caption         =   "Oppstart:"
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
         Left            =   -74760
         TabIndex        =   48
         Top             =   3740
         Width           =   4575
         Begin VB.TextBox txtUserLoad 
            Height          =   285
            Left            =   2865
            TabIndex        =   50
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblStart 
            Caption         =   "Start følgende bruker:"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.Frame frameTags 
         Caption         =   "Bruk følgende koder:"
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
         Left            =   -74760
         TabIndex        =   40
         Top             =   2540
         Width           =   4575
         Begin VB.CheckBox chkTag 
            Caption         =   "YA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   6
            Left            =   3840
            TabIndex        =   47
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "CH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Index           =   5
            Left            =   3180
            TabIndex        =   46
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   45
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   44
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   43
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   42
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkTag 
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   41
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame frameTimeZone 
         Caption         =   "Posisjon:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74760
         TabIndex        =   32
         Top             =   3740
         Width           =   4575
         Begin VB.TextBox txtTimeZone 
            Height          =   285
            Left            =   2040
            TabIndex        =   38
            Top             =   1170
            Width           =   2295
         End
         Begin VB.TextBox txtLongitude 
            Height          =   285
            Left            =   2040
            TabIndex        =   37
            Top             =   820
            Width           =   2295
         End
         Begin VB.TextBox txtLatitude 
            Height          =   285
            Left            =   2040
            TabIndex        =   36
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblTimeZone 
            Caption         =   "Tidssone:"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblLatitude 
            Caption         =   "Breddegrader:"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblLongitude 
            Caption         =   "Lengdegrader:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame frameUsers 
         Caption         =   "Administrer brukere:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -74760
         TabIndex        =   23
         Top             =   620
         Width           =   4575
         Begin VB.ListBox lstUsers 
            Height          =   1620
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   4095
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "&Endre"
            Height          =   375
            Left            =   3000
            TabIndex        =   26
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdDeleteUser 
            Caption         =   "&Slett"
            Height          =   375
            Left            =   1560
            TabIndex        =   25
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdNewUser 
            Caption         =   "&Ny bruker"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   2280
            Width           =   1215
         End
      End
      Begin VB.Frame frameScript 
         Caption         =   "Skript:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74760
         TabIndex        =   18
         Top             =   620
         Width           =   4575
         Begin VB.CheckBox chkProcessCL 
            Caption         =   "&Kjør kommandolinjen"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   1440
            Width           =   2655
         End
         Begin VB.ComboBox cmbLanguage 
            Height          =   315
            ItemData        =   "frmSettings.frx":0070
            Left            =   2160
            List            =   "frmSettings.frx":007D
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox chkRunScript 
            Caption         =   "&Kjør skript fra registeret i oppstart"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label lblScriptLang 
            Caption         =   "Skript språk:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame frameAdvance 
         Caption         =   "Avansert:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74760
         TabIndex        =   11
         Top             =   2780
         Width           =   4575
         Begin VB.TextBox txtTransparent 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   31
            Top             =   800
            Width           =   2175
         End
         Begin VB.CheckBox chkEnableTransparent 
            Caption         =   "&Muliggjør gjennomsiktige vinduer"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   2880
            Width           =   3015
         End
         Begin VB.CheckBox chkDefaultTransparent 
            Caption         =   "&Gjennomsiktige vinduer som standard"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   2520
            Width           =   3375
         End
         Begin VB.CheckBox chkDumpError 
            Caption         =   "Logg feilmeldinger"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtDefaultUser 
            Height          =   285
            Left            =   2160
            TabIndex        =   14
            Top             =   435
            Width           =   2175
         End
         Begin VB.CheckBox chkUseHooking 
            Caption         =   "Bruk 'hooking'"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CheckBox chkLogErrorInApp 
            Caption         =   "Logg feilmeldinger i hendelsesloggen"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1800
            Width           =   3135
         End
         Begin VB.Label lblTransparent 
            Caption         =   "Prosent gjennomsiktig:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblDefaultUser 
            Caption         =   "Standard bruker:"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame frameCalender 
         Caption         =   "Kalender:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74760
         TabIndex        =   3
         Top             =   620
         Width           =   4575
         Begin VB.CheckBox chkLagreFont 
            Caption         =   "Stor skrift"
            Height          =   255
            Left            =   2520
            TabIndex        =   39
            Top             =   960
            Width           =   2030
         End
         Begin VB.CheckBox chkShowHolidays 
            Caption         =   "Vis heligdager"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox chkOwnDays 
            Caption         =   "Vis egne merkedager"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   2295
         End
         Begin VB.CheckBox chkShowFlag 
            Caption         =   "Vis flag"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox chkSommertid 
            Caption         =   "Sommertid"
            Height          =   255
            Left            =   2520
            TabIndex        =   7
            Top             =   1200
            Width           =   2030
         End
         Begin VB.CheckBox chkShowSun 
            Caption         =   "Vis sol opp og ned"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox chkShowDayLenght 
            Caption         =   "Vis dag nr. / lengde"
            Height          =   255
            Left            =   2520
            TabIndex        =   5
            Top             =   480
            Width           =   2030
         End
         Begin VB.CheckBox chkShowMoon 
            Caption         =   "Vis måne"
            Height          =   255
            Left            =   2520
            TabIndex        =   4
            Top             =   720
            Width           =   2030
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Public Sub Update()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".Update", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell As Long, Plugin, sFile, ListItem As ListItem

' Update users
lstUsers.Clear

For Tell = LBound(Users) To UBound(Users)
    If Users(Tell).UserName <> "" Then
        lstUsers.AddItem Users(Tell).UserName
    End If
Next

' Update plugins
lstPlugins.ListItems.Clear

For Each Plugin In Plugins
    Set ListItem = lstPlugins.ListItems.Add(, , Plugin.Name)
    
    ListItem.Checked = GetSetting("YourTime", "Plugin", Plugin.Name, True)
    ListItem.SubItems(1) = Plugin.Description
Next

' Update language pack
cmdLanguagePack.Clear

For Each sFile In Language.EnumLanguagePacks(App.Path, "lpk")
    cmdLanguagePack.AddItem sFile
Next

' Set the selected item to the current language pack
For Tell = 0 To cmdLanguagePack.ListCount
    If cmdLanguagePack.List(Tell) = Script.LanguagePack And cmdLanguagePack.List(Tell) <> "" Then
        cmdLanguagePack.ListIndex = Tell
    End If
Next

End Sub

Public Sub LoadData()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".LoadData", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell As Long

txtTransparent.Text = Script.TransparentKey
txtLatitude.Text = Script.Latitude
txtLongitude.Text = Script.Longitude
txtTimeZone.Text = Script.TimeZone
txtUserLoad.Text = Script.StartupUser

chkShowHolidays.Value = Script.ShowHoliday
chkOwnDays.Value = Script.ShowOwnDays
chkShowFlag.Value = Script.ShowFlag
chkShowSun.Value = Script.ShowSunUpDown
chkShowDayLenght.Value = Script.ShowDayLenght
chkShowMoon.Value = Script.ShowMoon
chkSommertid.Value = Script.Sommertid

cmbLanguage.ListIndex = Script.DefaultLanguage
txtDefaultUser.Text = Script.DefaultUser
chkDumpError.Value = IIf(Script.DumpError, 1, 0)
chkUseHooking.Value = IIf(Script.UseHooking, 1, 0)
chkLogErrorInApp.Value = IIf(Script.LogToEventLog, 1, 0)
chkRunScript.Value = IIf(Script.RunScript, 1, 0)
chkProcessCL.Value = IIf(Script.ProcessCL, 1, 0)
chkEnableTransparent.Value = IIf(Script.EnableTransparent, 1, 0)
chkDefaultTransparent.Value = IIf(Script.TransparentDefault, 1, 0)
chkLagreFont.Value = IIf(Script.LagreFont, 1, 0)
chkFindPlugins.Value = IIf(Script.FindPlugins, 1, 0)

For Tell = 0 To chkTag.count - 1
    chkTag(Tell).Value = Script.EnableTags(chkTag(Tell).Caption)
Next

Update

End Sub

Public Sub SaveData()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".SaveData", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell As Long, Form As Form

Script.ShowHoliday = chkShowHolidays.Value
Script.ShowOwnDays = chkOwnDays.Value
Script.ShowFlag = chkShowFlag.Value
Script.ShowSunUpDown = chkShowSun.Value
Script.ShowDayLenght = chkShowDayLenght.Value
Script.ShowMoon = chkShowMoon.Value
Script.Sommertid = chkSommertid.Value

Script.DefaultLanguage = cmbLanguage.ListIndex
Script.DefaultUser = txtDefaultUser.Text
Script.DumpError = (chkDumpError.Value = 1)
Script.UseHooking = (chkUseHooking.Value = 1)
Script.LogToEventLog = (chkLogErrorInApp.Value = 1)
Script.RunScript = (chkRunScript.Value = 1)
Script.ProcessCL = (chkProcessCL.Value = 1)
Script.EnableTransparent = (chkEnableTransparent.Value = 1)
Script.TransparentDefault = (chkDefaultTransparent.Value = 1)
Script.LagreFont = (chkLagreFont.Value = 1)
Script.FindPlugins = (chkFindPlugins.Value = 1)
Script.TransparentKey = txtTransparent.Text

Script.Latitude = txtLatitude.Text
Script.Longitude = txtLongitude.Text
Script.TimeZone = txtTimeZone.Text
Script.StartupUser = txtUserLoad.Text
Script.LanguagePack = cmdLanguagePack.Text

For Tell = 0 To chkTag.count - 1
    Script.EnableTags(chkTag(Tell).Caption) = chkTag(Tell).Value
Next

For Tell = 1 To lstPlugins.ListItems.count
    SaveSetting "YourTime", "Plugin", Plugins(Tell).Name, lstPlugins.ListItems(Tell).Checked
Next

' Update all forms
For Each Form In Forms
    Language.SetLanguageInForm Form
Next

modGeneral.UpdateSun
frmMain.uscCalender.Redraw

End Sub

Public Function AddNewUser() As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".AddNewUser", , EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Pass$, User$, Tell&

If Script.NewUser("Ny bruker", User, Pass) = 0 Then

    If Script.FindUser(User) = 0 Then
        
        Tell = Script.FindUser("")
        
        If Tell <= 0 Then
            MsgBox "Ingen flere ledige brukere", vbCritical, "Feil"
            AddNewUser = -1
            Exit Function
        End If
        
        MD5.InBuffer = Pass
        MD5.Calculate
        
        Users(Tell).Created = Now
        Users(Tell).UserName = User
        Users(Tell).Password = MD5.OutBuffer
        ReDim Users(Tell).DataOwn(Script.NumOfRecords)
        
        AddNewUser = Tell
        Exit Function
    Else
        MsgBox "Brukernavnet eksisterer allerede", vbCritical, "Feil"
    End If
End If

AddNewUser = -1

End Function

Private Sub cmdCancel_Click()

Me.Hide

End Sub

Private Sub cmdChange_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdChange_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim User$, Pass$, OldPass, Tell&

If lstUsers.ListIndex >= 0 Then

    Tell = Script.FindUser(lstUsers.Text)

    If Tell <= 0 Then
        MsgBox "Finner ikke bruker", vbCritical, "Feil"
        Exit Sub
    End If

    User = Users(Tell).UserName
    OldPass = True

    If Script.NewUser("Endre bruker", User, Pass, OldPass) = 0 Then
        
        MD5.InBuffer = OldPass
        MD5.Calculate

        If MD5.OutBuffer = Users(Tell).Password Or Users(Tell).Password = "" Then
            'Change user
            MD5.InBuffer = Pass
            MD5.Calculate
            
            Users(Tell).Password = MD5.OutBuffer
            Users(Tell).UserName = User
            Users(Tell).Changed = True
            
            If Pass = "" Then
                Users(Tell).Password = ""
            End If
            
            Update
        Else
            MsgBox "Passord ikke korrekt", vbCritical, "Feil"
            Exit Sub
        End If
        
    End If

End If

End Sub

Private Sub cmdConfigure_Click()

On Error Resume Next
Plugins(lstPlugins.SelectedItem.Index).Configure

End Sub

Private Sub cmdDeleteUser_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdDeleteUser_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

If lstUsers.ListIndex >= 0 Then

    Tell = Script.FindUser(lstUsers.Text)
    
    If Tell <= 0 Then
        MsgBox "Finner ikke bruker", vbCritical, "Feil"
        Exit Sub
    End If
    
    If Users(Tell).Password <> "" Then
        If InputBox("Skriv inn passord for brukeren") <> Users(Tell).Password Then
            MsgBox "Passord ikke korrekt", vbCritical, "Feil"
            Exit Sub
        End If
    Else
        If MsgBox("Vil du virkelig slette denne brukeren?", vbQuestion + vbYesNo, "Bekreft sletting") = vbNo Then
            Exit Sub
        End If
    End If
    
    Users(Tell).UserName = ""
    
    Update
    MsgBox "Bruker slettet!", vbInformation, "Slett"
End If

End Sub

Private Sub cmdNewUser_Click()

AddNewUser
Update

End Sub

Private Sub cmdOK_Click()

SaveData
Me.Hide

End Sub

Private Sub cmdUse_Click()

SaveData

End Sub

Private Sub Form_Load()

' Common code that should be executet in all forms
FormLoad Me

SetNumber txtTransparent, True
SetNumber txtUserLoad, True

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Unsubclass the form
UnHookForm Me.hwnd

End Sub

Private Sub lstPlugins_Click()

On Error Resume Next

If lstPlugins.SelectedItem.Index < Plugins.count And lstPlugins.SelectedItem.Index >= 0 Then
    cmdConfigure.Enabled = Plugins(lstPlugins.SelectedItem.Index)
End If

End Sub

Private Sub lstPlugins_KeyDown(KeyCode As Integer, Shift As Integer)

lstPlugins_Click

End Sub

Private Sub txtLatitude_Change()

txtLatitude.Text = Script.ConvertToNumeric(txtLatitude.Text)

End Sub

Private Sub txtLongitude_Change()

txtLongitude.Text = Script.ConvertToNumeric(txtLongitude.Text)

End Sub

Private Sub txtTimeZone_Change()

txtTimeZone.Text = Script.ConvertToNumeric(txtTimeZone.Text)

End Sub
