VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Your time"
   ClientHeight    =   9060
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   15270
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   604
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   3  'Windows Default
   Begin YourTime.SheetTabs SheetTabs1 
      Height          =   375
      Left            =   0
      TabIndex        =   40
      Top             =   8640
      Width           =   15255
      _extentx        =   26908
      _extenty        =   661
      extrawidth      =   29
      fontsize        =   12
   End
   Begin YourTime.Calender uscCalender 
      Height          =   7935
      Left            =   60
      TabIndex        =   39
      Top             =   60
      Width           =   5055
      _extentx        =   8916
      _extenty        =   13996
      fontsize        =   8.25
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   20
      Left            =   12000
      Top             =   5400
   End
   Begin VB.PictureBox picPanel 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   8175
      Left            =   14520
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   19
      Top             =   60
      Width           =   735
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   0
         Picture         =   "frmMain.frx":02EA
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Avslutt"
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   360
         Picture         =   "frmMain.frx":065D
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vis data med linjer"
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   0
         Picture         =   "frmMain.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Vis data med linjer"
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   0
         Picture         =   "frmMain.frx":0D4B
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Årsoversikt"
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Søk etter"
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   0
         Picture         =   "frmMain.frx":10ED
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Fjerner eller legger til dagens oppgaver"
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   0
         Picture         =   "frmMain.frx":149F
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Telefonbok"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   360
         Picture         =   "frmMain.frx":1843
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Gå til dagsdato"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   0
         Picture         =   "frmMain.frx":1BCF
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Dagbok"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   0
         Picture         =   "frmMain.frx":1F54
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Persondatabase"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         Picture         =   "frmMain.frx":22DC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Dag tilbake"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         Picture         =   "frmMain.frx":2373
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Dag fram"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         Caption         =   "ð"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   20
         ToolTipText     =   "År fram"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdAny 
         Caption         =   "ï"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   21
         ToolTipText     =   "År tilbake"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   1980
      Left            =   5355
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   575
      TabIndex        =   9
      Top             =   60
      Width           =   8655
      Begin VB.PictureBox picUsers 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   0
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   577
         TabIndex        =   23
         Top             =   1440
         Width           =   8655
         Begin VB.CommandButton cmdUsers 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Label lblDayDescription 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1320
         TabIndex        =   32
         Top             =   120
         Width           =   3975
      End
      Begin VB.Image imgDiary 
         Height          =   255
         Left            =   4440
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label lblFooter 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Bruker:   Ukjent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   -15
         TabIndex        =   22
         Top             =   1080
         Width           =   8655
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   57.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   -75
         Width           =   1095
      End
      Begin VB.Label lblMonth 
         BackStyle       =   0  'Transparent
         Caption         =   "Jan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   16
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lblYear 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   2400
         TabIndex        =   15
         Top             =   690
         Width           =   840
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "12:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   7320
         TabIndex        =   14
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label lblWeekday 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label lblWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "UKE1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label lblDays 
         BackStyle       =   0  'Transparent
         Caption         =   "100 / 231 dager."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblSunstat 
         BackStyle       =   0  'Transparent
         Caption         =   "Sol opp: 8.37   ned: 12:23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.PictureBox picRem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   3855
      Left            =   5355
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   4
      Top             =   2025
      Width           =   4335
      Begin VB.VScrollBar vscRem 
         Height          =   3855
         Left            =   4080
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picIndicator 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   585
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   7
         Top             =   -15
         Width           =   615
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1185
         TabIndex        =   6
         Top             =   -15
         Width           =   2775
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00.00"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   -15
         TabIndex        =   8
         Top             =   -15
         Width           =   615
      End
   End
   Begin VB.PictureBox picTasks 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   1965
      Left            =   5355
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   0
      Top             =   6000
      Width           =   4335
      Begin VB.PictureBox picTaskIndicator 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -15
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   3
         Top             =   -15
         Width           =   615
      End
      Begin VB.TextBox txtTask 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   585
         TabIndex        =   2
         Top             =   -15
         Width           =   3495
      End
      Begin VB.VScrollBar vscTasks 
         Height          =   1935
         Left            =   4080
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.ListBox lstData 
      Height          =   840
      Left            =   10560
      TabIndex        =   36
      Top             =   2010
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblTasks 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Oppgaver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5355
      TabIndex        =   18
      Top             =   5880
      Width           =   4335
   End
   Begin VB.Menu mnuFil 
      Caption         =   "&Fil"
      Begin VB.Menu mnuRunScript 
         Caption         =   "Kjør skript"
      End
      Begin VB.Menu mnuPlugins 
         Caption         =   "Plug-ins"
         Begin VB.Menu mnuPlugInMenu 
            Caption         =   "Empty"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Sikkerhetskopi"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Avslutt"
      End
   End
   Begin VB.Menu mnuBruker 
      Caption         =   "&Your Time"
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Slett all data"
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "Slett brukerdata"
      End
      Begin VB.Menu mnuDeleteUserDaylig 
         Caption         =   "Slett brukers daglige tekster"
      End
      Begin VB.Menu mnuDeleteUserDB 
         Caption         =   "Slett persondatabasen"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOwn 
         Caption         =   "Egne merkedager"
      End
      Begin VB.Menu mnuPersonDB 
         Caption         =   "Persondatabase"
      End
      Begin VB.Menu mnuDiary 
         Caption         =   "Dagbok"
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoToCurrentDay 
         Caption         =   "Gå til dagsdato"
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Arternativer..."
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Skriv ut"
      Begin VB.Menu mnuScreenshot 
         Caption         =   "Skjermbildet"
      End
      Begin VB.Menu mnuPrintDiary 
         Caption         =   "Dagbok"
      End
      Begin VB.Menu mnuPrintRem 
         Caption         =   "Avtaler og oppgaver"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Hjelp"
      Begin VB.Menu mnuAbout 
         Caption         =   "Om"
      End
   End
   Begin VB.Menu mnuRemember 
      Caption         =   "Remember"
      Visible         =   0   'False
      Begin VB.Menu mnuMove 
         Caption         =   "Flytt"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Kopier"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Sett inn"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Slett"
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "Informasjon"
      End
      Begin VB.Menu mnuRemoveInformation 
         Caption         =   "Fjern informasjon"
      End
      Begin VB.Menu mnuGetFromDatabase 
         Caption         =   "Hent fra persondatabase"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Dim udtClipboard As InternalClipboard
Dim SelectedObject As Object

Private Sub cmdAny_Click(Index As Integer)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdAny_Click(Index)", Array(Index), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Num&, LastNum&, LastDate As Date, strSearch$, Tell&, h, m

ProcessControls "saveall"

Select Case Index
Case 0: CurrentDate.cYear = CurrentDate.cYear - 1
Case 1: CurrentDate.cYear = CurrentDate.cYear + 1
Case 2: CurrentDate.cDay = CurrentDate.cDay - 1
Case 3: CurrentDate.cDay = CurrentDate.cDay + 1
Case 4
    frmPerson.Show

Case 5
    frmDiary.LoadCurrentDate
    frmDiary.Show
    
Case 6
    CurrentDate.Contents = Date

Case 7
    ' Phonebook
    frmPhonebook.Show
    
Case 8
    ' Exit program
    Unload Me

Case 9
    ' Fjerner eller setter inn dagens oppgaver
    lblTasks.Visible = Not lblTasks.Visible
    picTasks.Visible = Not picTasks.Visible
    picRem.Visible = True
    lstData.Visible = False

    ' Oppdaterer formen
    Form_Resize

Case 10

    ' Søke-prosedyren
    
    strSearch = InputBox("Skriv inn søketeksten", "Søk")
    If strSearch = "" Then Exit Sub

    LastNum = Search(Users(UserID).DataRem, CurrentDate.Contents, 1)
    Num = Search(Users(UserID).DataRem, strSearch, 5, LastNum)
    
    If Num >= 0 Then
        CurrentDate.Contents = Script.RemoveTime(Users(UserID).DataRem(Num).RemDate)
    Else
    
        LastNum = Search(Users(UserID).DataTasks, CurrentDate.Contents, 1)
        Num = Search(Users(UserID).DataTasks, strSearch, 5, LastNum)
    
        If Num >= 0 Then
            CurrentDate.Contents = Script.RemoveTime(Users(UserID).DataTasks(Num).RemDate)
        Else
            Exit Sub
        End If
    
    End If
    
Case 11
    
    ' På og av med listbox
    lstData.Visible = Not lstData.Visible
    lblTasks.Visible = Not lstData.Visible
    picTasks.Visible = Not lstData.Visible
    picRem.Visible = Not lstData.Visible
    
    ' Oppdaterer formen
    Form_Resize
    
Case 12

    ' Gå til neste avtale
    
    For Tell = vscRem.Value + 1 To (24 * 60) / RemStep
    
        ' Beregn tiden
        h = Tell / (60 / RemStep)
        m = (h - Val(h)) * 60
        h = Val(h)
    
        Num = Search(Users(UserID).DataRem, CurrentDate.Contents + TimeSerial(h, m, 0), 0)
        
        If Num >= 0 Then
            vscRem.Value = IIf(Tell > vscRem.Max, vscRem.Max, Tell)
            Exit For
        End If
    
    Next

Case 13

    ' Årsoversikt
    frmYear.YearReview.Redraw
    frmYear.Show

End Select

ProcessControls "update"

End Sub

Private Sub cmdUsers_Click(Index As Integer)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".cmdUsers_Click(Index)", Array(Index), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

If cmdUsers(Index).Caption = "" Then Exit Sub

ProcessControls "saveall"
UserID = Index + 1
ProcessControls "update"

End Sub

Private Sub Form_Load()

Dim Tell&

' Common code that should be executet in all forms
FormLoad Me

For Tell = 1 To 12
    SheetTabs1.AddTab CapitalizeFirstLetter(MonthName(Tell))
Next

Set cmdAny(10).Picture = frmPhonebook.cmdSearch.Picture

End Sub

Private Sub Form_Resize()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".Form_Resize", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

If Me.WindowState = 1 Then Exit Sub

If Not Script.UseHooking Then
    If Me.ScaleWidth < Script.MinWidth Then Me.Width = Me.ScaleX(Script.MinWidth, vbPixels, vbTwips)
    If Me.ScaleHeight < Script.MinHeight Then Me.Height = Me.ScaleY(Script.MinHeight, vbPixels, vbTwips)
End If

uscCalender.Height = Me.ScaleHeight - 25
uscCalender.Width = Me.ScaleWidth / 3

picHeader.Width = Me.ScaleWidth - uscCalender.Width - 62
picHeader.Width = picHeader.Width
picHeader.Left = uscCalender.Width + uscCalender.Left + 4
picUsers.Width = picHeader.Width

If lstData.Visible = True Then
    lstData.Height = uscCalender.Height - (lstData.Top - uscCalender.Top)
    lstData.Width = picHeader.Width
    lstData.Left = picHeader.Left
End If

If picRem.Visible = True Then
    picRem.Left = picHeader.Left
    picRem.Height = uscCalender.Height - (picRem.Top - uscCalender.Top) - IIf(lblTasks.Visible, 156, 0) - 6
    picRem.Width = picHeader.Width

    If picRem.Height > txtRem(0).Height * (24 * (60 / RemStep)) - 5 Then picRem.Height = txtRem(0).Height * (24 * (60 / RemStep)) - 5
End If

If lblTasks.Visible Then
    lblTasks.Top = picRem.Height + picRem.Top - 1
    lblTasks.Left = picRem.Left
    lblTasks.Width = picRem.Width

    picTasks.Top = lblTasks.Top + lblTasks.Height - 1
    picTasks.Height = uscCalender.Height - lblTasks.Top - 25
    picTasks.Left = lblTasks.Left
    picTasks.Width = picRem.Width
End If

picPanel.Left = picHeader.Left + picHeader.Width + 4
lblFooter.Width = picHeader.Width

vscRem.Height = picRem.Height
vscRem.Left = picRem.Width - vscRem.Width - 1

vscTasks.Height = picTasks.Height
vscTasks.Left = picTasks.Width - vscTasks.Width - 1

lblSunstat.Left = picHeader.Width - lblSunstat.Width
lblTime.Left = picHeader.Width - lblTime.Width
lblDays.Left = lblSunstat.Left
lblWeek.Left = lblSunstat.Left

SheetTabs1.Width = Me.ScaleWidth
SheetTabs1.Top = Me.ScaleHeight - SheetTabs1.Height

ProcessControls "cmdusers"
ProcessControls "txtrem"
ProcessControls "txttask"
ProcessControls "lbltimer"

uscCalender.CalenderDate = CurrentDate.Contents

End Sub

Public Sub ProcessControls(ControlName As String)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".ProcessControls(ControlName)", Array(ControlName), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, Num&, h&, m&, Tmp, TmpDate As Date

Select Case LCase(ControlName)
Case "saveall"
    ProcessControls "savetxtrem"
    ProcessControls "savetxttasks"

Case "update"

    ProcessControls "lbltimer"
    ProcessControls "txttask"
    ProcessControls "updatediary"

    uscCalender.CalenderDate = CurrentDate.Contents
    SheetTabs1.Selected = CurrentDate.cMonth - 1
    
Case "updatediary"

    If Search(Users(UserID).DataDiary, CurrentDate.Contents, 0) >= 0 Then
        Set imgDiary = cmdAny(5).Picture
    Else
        Set imgDiary = Nothing
    End If

Case "savetxttasks"

    For Tell = 0 To txtTask.count - 1
        TmpDate = CurrentDate.Contents + TimeSerial(0, Tell + Val(vscTasks.Tag), 0)
        SaveRecord Users(UserID).DataTasks, txtTask(Tell).Text, TmpDate
    Next

Case "savetxtrem"

    For Tell = 0 To txtRem.count - 1
        Tmp = Split(lblTimer(Tell).Caption, ".")
        TmpDate = CurrentDate.Contents + TimeSerial(CInt(Tmp(0)), CInt(Tmp(1)), 0)
        SaveRecord Users(UserID).DataRem, txtRem(Tell).Text, TmpDate
    Next

Case "cmdusers"

    For Tell = 0 To cmdUsers.count - 1
        cmdUsers(Tell).Width = picHeader.Width / 10
        cmdUsers(Tell).Caption = Users(Tell + 1).UserName
        
        If Tell > 0 Then
            cmdUsers(Tell).Left = cmdUsers(Tell - 1).Left + cmdUsers(Tell).Width
            cmdUsers(Tell).Top = cmdUsers(Tell - 1).Top
        End If
        
        If cmdUsers(Tell).Left + cmdUsers(Tell).Width > picHeader.Width + 10 Then
            cmdUsers(Tell).Top = cmdUsers(Tell).Top + cmdUsers(Tell).Height
            cmdUsers(Tell).Left = 0
        End If
    Next

Case "lbltimer"

    Select Case lstData.Visible
    Case False

        ' Hvis vi bruker kontroller for å vise all data

        For Tell = 1 To vscRem.Value
            m = m + RemStep
            
            If m >= 60 Then
                m = m - 60
                h = h + 1
            End If
        Next
    
        For Tell = 0 To lblTimer.count - 1
            lblTimer(Tell).Caption = FillOut(CStr(h), 2) & "." & FillOut(CStr(m), 2)
            
            Num = Search(Users(UserID).DataRem, CurrentDate.Contents + TimeSerial(h, m, 0), 0)
        
            If Num >= 0 Then
                txtRem(Tell).Text = Users(UserID).DataRem(Num).Text
            Else
                txtRem(Tell).Text = ""
            End If
        
            m = m + RemStep
            
            If m >= 60 Then
                m = m - 60
                h = h + 1
            End If
        Next

    Case True
    
        ' Koden for å legge til data i listbox-en
        Script.AddRem lstData, CurrentDate.Contents

    End Select
    
Case "txtrem"

    Num = picRem.Height / txtRem(0).Height
    vscRem.Max = (24 * (60 / RemStep)) - Num - 1
    vscRem.Value = IIf((Hour(Time) * 4) + (Minute(Time) \ 15) < vscRem.Max, (Hour(Time) * 4) + (Minute(Time) \ 15), vscRem.Max)

    If Num <> txtRem.count Then
    
        For Tell = 1 To Num + 1
        
            If Tell > txtRem.count - 1 Then
                Load txtRem(Tell)
                Load picIndicator(Tell)
                Load lblTimer(Tell)
                
                txtRem(Tell).Top = txtRem(Tell - 1).Top + txtRem(Tell).Height - 1
                txtRem(Tell).Visible = True
                
                picIndicator(Tell).Visible = True
                picIndicator(Tell).Top = txtRem(Tell).Top
                
                lblTimer(Tell).Visible = True
                lblTimer(Tell).Top = txtRem(Tell).Top
            End If
        Next
    
    End If
    
    For Tell = 0 To txtRem.count - 1
        txtRem(Tell).Width = picRem.Width - txtRem(Tell).Left - vscRem.Width
    Next
    
Case "txttask"

    Num = picTasks.Height / txtTask(0).Height
    vscTasks.Max = 50 - Num

    If Num <> txtTask.count Then
    
        For Tell = 1 To Num + 1
        
            If Tell > txtTask.count - 1 Then
                Load txtTask(Tell)
                Load picTaskIndicator(Tell)
                
                txtTask(Tell).Top = txtTask(Tell - 1).Top + txtTask(Tell).Height - 1
                txtTask(Tell).Visible = True
                
                picTaskIndicator(Tell).Visible = True
                picTaskIndicator(Tell).Top = txtTask(Tell).Top
            End If
        Next
    
    End If
    
    ' Oppdater visiningen av data
    
    Select Case lstData.Visible
    Case False
    
        ' Om vi bruker kontroller for å vise data
    
        For Tell = 0 To txtTask.count - 1
            txtTask(Tell).Width = picTasks.Width - txtTask(Tell).Left - vscTasks.Width
            
            Num = Search(Users(UserID).DataTasks, CurrentDate.Contents + TimeSerial(0, Tell + vscTasks.Value, 0), 0)
        
            If Num >= 0 Then
                txtTask(Tell).Text = Users(UserID).DataTasks(Num).Text
            Else
                txtTask(Tell).Text = ""
            End If
        Next
    
    Case True
    
        ' Hvis vi bruker en listbox for å vise data
        Script.AddTasks lstData, CurrentDate.Contents
        
    End Select

End Select

' Process other messages
DoEvents

End Sub

Public Sub UpdateAll()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".UpdateAll", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&

For Tell = 0 To cmdUsers.count - 1
    If cmdUsers(Tell).Caption <> Users(Tell + 1).UserName Then
        cmdUsers(Tell).Caption = Users(Tell + 1).UserName
    End If
Next

If Hour(Time) = 0 And Minute(Time) = 0 And Second(Time) = 0 Then
    CurrentDate.Add "d", 1
End If

uscCalender.Header = Space(2) & CapitalizeFirstLetter(CurrentDate.cMonthName) & Space(9) & CurrentDate.cYear
SheetTabs1.Selected = CurrentDate.cMonth - 1

lblTime.Caption = Time
lblDay.Caption = CurrentDate.cDay
lblDayDescription.Caption = Script.RemoveTags(uscCalender.Text)
lblMonth.Caption = CapitalizeFirstLetter(Left$(CurrentDate.cMonthName, 3))
lblYear.Caption = CurrentDate.cYear
lblWeekday.Caption = CapitalizeFirstLetter(CurrentDate.cDayName)
lblWeek.Caption = Language.Constant("Week") & ": " & CurrentDate.cWeekNum
lblFooter.Caption = Space(3) & Language.Constant("User") & ":" & Space(3) & Users(UserID).UserName

If Script.ShowDayLenght = 1 Then
    lblDays.Caption = CurrentDate.cTotalDays & " \ " & CurrentDate.cLeapYear + 337
    lblDays.Visible = True
Else
    lblDays.Visible = False
End If

If Script.ShowSunUpDown = 1 Then
    Sun.cCurrentDate = CurrentDate.Contents
    
    If Sun.Changed Then
        Sun.Calculate
        lblSunstat.Visible = True
        lblSunstat.Caption = Language.Constant("SunUp") & ": " & Format(Sun.SunRise, "hh:mm") & Space(3) & Language.Constant("Down") & ": " & Format(Sun.SunSet, "hh:mm")
    End If
Else
    lblSunstat.Visible = False
End If

' Process other messages
DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)

Script.Quit

End Sub

Private Sub mnuAbout_Click()

frmAbout.MainLoop

End Sub

Private Sub mnuBackup_Click()

frmBackup.Show

End Sub

Private Sub mnuCopy_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".mnuCopy_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Num As Long

Select Case LCase(SelectedObject.Name)
Case "txtrem"

    Num = GetObjectIndex(SelectedObject)

    If Num >= 0 Then
        udtClipboard.DataType = 1
        Let udtClipboard.Data = Users(UserID).DataRem(Num)
    End If

Case "txttask"

    Num = GetObjectIndex(SelectedObject)

    If Num >= 0 Then
        udtClipboard.DataType = 2
        Let udtClipboard.Data = Users(UserID).DataTasks(Num)
    End If

End Select

End Sub

Private Sub mnuDelete_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".mnuDelete_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

SelectedObject.Text = ""
ProcessControls "saveall"
ProcessControls "update"

End Sub

Private Sub mnuDeleteAll_Click()

Dim Tell As Long

For Tell = 1 To 20
    Script.DeleteUserData Tell
Next

ProcessControls "update"
MsgBox Language.Constant("DeleteAllConfirmed"), vbInformation, "Slett"

End Sub

Private Sub mnuDeleteUser_Click()

Script.DeleteUserData UserID

ProcessControls "update"
MsgBox Language.Constant("DeleteUserConfirmed"), vbInformation, "Slett"

End Sub

Private Sub mnuDeleteUserDaylig_Click()

Erase Users(UserID).DataRem
Erase Users(UserID).DataTasks

ProcessControls "update"
MsgBox Language.Constant("DeleteUserDayligConfirmed"), vbInformation, "Slett"

End Sub

Private Sub mnuDeleteUserDB_Click()

Erase Users(UserID).DataPeoples

MsgBox Language.Constant("UserDBDeletetConfirmed"), vbInformation, "Slett"

End Sub

Private Sub mnuDiary_Click()

cmdAny_Click 5

End Sub

Private Sub mnuExit_Click()

Unload Me

End Sub

Private Sub mnuGetFromDatabase_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".mnuGetFromDatabase_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Num As Long

Num = GetObjectIndex(SelectedObject)

Select Case LCase(SelectedObject.Name)
Case "txtrem": Users(UserID).DataRem(Num).ExLong = Script.ChoosePerson
Case "txttask": Users(UserID).DataTasks(Num).ExLong = Script.ChoosePerson
End Select

End Sub

Private Sub mnuGoToCurrentDay_Click()

cmdAny_Click 6

End Sub

Public Function GetObjectIndex(lpObject As Object) As Long

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".GetObjectIndex(lpObject)", Array(ObjPtr(lpObject)), EA_NORERAISE: HandleError: Exit Function
End If
' *** BEGIN CODE ***

Dim Val As Variant

Select Case LCase(lpObject.Name)
Case "txtrem"
    
    Val = Split(lblTimer(lpObject.Index).Caption, ".")
    
    If UBound(Val) > 0 Then
        GetObjectIndex = Search(Users(UserID).DataRem, CurrentDate.Contents + TimeSerial(Val(0), Val(1), 0), 0)
    Else
        GetObjectIndex = -1
    End If

Case "txttask"

    GetObjectIndex = Search(Users(UserID).DataTasks, CurrentDate.Contents + TimeSerial(0, lpObject.Index + vscTasks.Value, 0), 0)

Case Else
    
    GetObjectIndex = -1
End Select

End Function

Private Sub mnuInformation_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".mnuInformation_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim ret&, Num&, sText$

Select Case LCase(SelectedObject.Name)
Case "txtrem"
    
    Num = GetObjectIndex(SelectedObject)

    If Num >= 0 Then
        
        sText = Users(UserID).DataRem(Num).ExtraData
        ret = Script.ShowInformation(sText)
        
        If ret > 0 Then
            Users(UserID).DataRem(Num).ExtraData = sText
        End If
        
    End If

Case "txttask"

    Num = GetObjectIndex(SelectedObject)
    
    If Num >= 0 Then
    
        sText = Users(UserID).DataTasks(Num).ExtraData
        ret = Script.ShowInformation(sText)
        
        If ret > 0 Then
            Users(UserID).DataTasks(Num).ExtraData = sText
        End If
    
    End If

End Select

End Sub

Private Sub mnuInsert_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".mnuInsert_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Num As Long, Val

ProcessControls "saveall"

Select Case LCase(SelectedObject.Name)
Case "txtrem"

    If udtClipboard.DataType = 1 Then
        Num = GetObjectIndex(SelectedObject)
    
        If Num >= 0 Then
            Users(UserID).DataRem(Num).Text = udtClipboard.Data.Text
            Users(UserID).DataRem(Num).ExtraData = udtClipboard.Data.ExtraData
            Users(UserID).DataRem(Num).ExLong = udtClipboard.Data.ExLong
        Else
            Val = Split(lblTimer(SelectedObject.Index).Caption, ".")
            AddRecord Users(UserID).DataRem, CurrentDate.Contents + TimeSerial(Val(0), Val(1), 0), udtClipboard.Data.Text, udtClipboard.Data.ExtraData, udtClipboard.Data.ExLong
        End If
    End If

Case "txttask"

    If udtClipboard.DataType = 2 Then
        Num = GetObjectIndex(SelectedObject)
    
        If Num >= 0 Then
            Users(UserID).DataTasks(Num).Text = udtClipboard.Data.Text
            Users(UserID).DataTasks(Num).ExtraData = udtClipboard.Data.ExtraData
            Users(UserID).DataTasks(Num).ExLong = udtClipboard.Data.ExLong
        Else
            AddRecord Users(UserID).DataTasks, CurrentDate.Contents + TimeSerial(0, SelectedObject.Index + vscTasks.Value, 0), udtClipboard.Data.Text, udtClipboard.Data.ExtraData, udtClipboard.Data.ExLong
        End If
    End If

End Select

ProcessControls "update"

End Sub

Private Sub mnuMove_Click()

' Dette er ganske enkelt, ettersom å flytte er jo det samme som å kopiere og slette samtidig.
mnuCopy_Click
mnuDelete_Click

End Sub

Private Sub mnuOptions_Click()

frmSettings.LoadData
frmSettings.Show

End Sub

Private Sub mnuOwn_Click()

frmOwn.LoadDatabase
frmOwn.Show

End Sub

Private Sub mnuPersonDB_Click()

frmPerson.Show

End Sub

Private Sub mnuPrintDiary_Click()

Dim lStart As Date, lEnd As Date, ret&

ret = Script.ChooseDate(lStart, lEnd)
If ret < 0 Then Exit Sub

PrintClass.PrintDiary Me, lStart, lEnd, Printer

End Sub

Private Sub mnuPrintRem_Click()

Dim lStart As Date, lEnd As Date, ret&

ret = Script.ChooseDate(lStart, lEnd)
If ret < 0 Then Exit Sub

PrintClass.PrintRemTasks lStart, lEnd, Printer

End Sub

Private Sub mnuRemoveInformation_Click()

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".mnuRemoveInformation_Click", , EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Num As Long

Select Case LCase(SelectedObject.Name)
Case "txtrem"
    
    Num = GetObjectIndex(SelectedObject)

    If Num >= 0 Then
        Users(UserID).DataRem(Num).ExtraData = ""
    End If

Case "txttask"

    Num = GetObjectIndex(SelectedObject)
    
    If Num >= 0 Then
        Users(UserID).DataTasks(Num).ExtraData = ""
    End If

End Select

End Sub

Private Sub mnuRunScript_Click()

frmScript.Show

End Sub

Private Sub mnuScreenshot_Click()

PrintClass.PrintForm Me

End Sub

Private Sub picIndicator_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".picIndicator_MouseDown(Index, Button, Shift, X, Y)", Array(Index, Button, Shift, X, Y), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Num As Long, Val

If Button = 2 Then
    mnuInsert.Enabled = (udtClipboard.DataType = 1)
    mnuDelete.Enabled = (txtRem(Index).Text <> "")
    mnuInformation.Enabled = mnuDelete.Enabled
    mnuMove.Enabled = mnuDelete.Enabled
    mnuCopy.Enabled = mnuDelete.Enabled
    
    Val = Split(lblTimer(Index).Caption, ".")
    
    If UBound(Val) > 0 Then
        Num = Search(Users(UserID).DataRem, CurrentDate.Contents + TimeSerial(Val(0), Val(1), 0), 0)
    Else
        Num = -1
    End If
    
    If Num >= 0 Then
        mnuRemoveInformation.Enabled = (Users(UserID).DataRem(Num).ExtraData <> "")
    Else
        mnuRemoveInformation.Enabled = False
    End If

    Set SelectedObject = txtRem(Index)
    Me.PopupMenu mnuRemember
End If

End Sub

Private Sub picTaskIndicator_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".picTaskIndicator_MouseDown(Index, Button, Shift, X, Y)", Array(Index, Button, Shift, X, Y), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Num As Long

If Button = 2 Then
    mnuInsert.Enabled = (udtClipboard.DataType = 2)
    mnuDelete.Enabled = (txtTask(Index).Text <> "")
    mnuInformation.Enabled = mnuDelete.Enabled
    mnuMove.Enabled = mnuDelete.Enabled
    mnuCopy.Enabled = mnuDelete.Enabled

    Num = Search(Users(UserID).DataTasks, CurrentDate.Contents + TimeSerial(0, Index + vscTasks.Value, 0), 0)
    
    If Num >= 0 Then
        mnuRemoveInformation.Enabled = (Users(UserID).DataTasks(Num).ExtraData <> "")
    Else
        mnuRemoveInformation.Enabled = False
    End If

    Set SelectedObject = txtTask(Index)
    Me.PopupMenu mnuRemember
End If

End Sub

Private Sub SheetTabs1_TabChanged(LastIndex As Long)

CurrentDate.cMonth = SheetTabs1.Selected + 1
ProcessControls "update"

End Sub

Private Sub SheetTabs1_TabChanging()

ProcessControls "saveall"

End Sub

Private Sub tmrUpdate_Timer()

UpdateAll

End Sub

Private Sub txtRem_Change(Index As Integer)

picIndicator(Index).BackColor = IIf(Len(txtRem(Index).Text) > 0, vbGreen, vbYellow)

End Sub

Private Sub txtRem_LostFocus(Index As Integer)

ProcessControls "savetxtrem"

End Sub

Private Sub txtTask_LostFocus(Index As Integer)

ProcessControls "savetxttasks"

End Sub

Private Sub uscCalender_DateChanged(LastDate As Date)

CurrentDate.Contents = uscCalender.CalenderDate

If CurrentDate.cYear <> Year(LastDate) Then
    ' Change static days
    InitializeVariables
End If

ProcessControls "update"

End Sub

Private Sub uscCalender_DateChanging()

ProcessControls "saveall"

End Sub

Private Sub uscCalender_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case vbKeyUp
    ' Day back
    cmdAny_Click 2

Case vbKeyDown
    ' Day forward
    cmdAny_Click 3
End Select

End Sub

Private Sub uscCalender_Redrawing(ItemCount As Long, lpTextArray() As String, MoonVal() As Long, ExtFlag() As Long)

' *** START ERROR HANDLER ***
On Error GoTo errHandler
If Err.Number <> 0 Then
errHandler: ErrorIn Me.Name & ".uscCalender_Redrawing(ItemCount)", Array(ItemCount), EA_NORERAISE: HandleError: Exit Sub
End If
' *** BEGIN CODE ***

Dim Tell&, hData&, lDate As Date, Settings(2) As Boolean

Settings(0) = CBool(Script.ShowOwnDays = 1)
Settings(1) = CBool(Script.ShowHoliday = 1)
Settings(2) = CBool(Script.ShowMoon = 1)

For Tell = 1 To ItemCount
    lDate = DateSerial(CurrentDate.cYear, CurrentDate.cMonth, Tell)
    
    ' Own days
    If Settings(0) Then
        hData = Search(Users(UserID).DataOwn, lDate, 3)
    
        If hData >= 0 Then
            lpTextArray(Tell) = Script.ProcessTags(Users(UserID).DataOwn(hData).Text, lDate, Users(UserID).DataOwn(hData).RemDate) & "  "
        End If
    End If
    
    ' Holidays
    If Settings(1) Then
        hData = Search(StaticDays, lDate, 3)

        If hData >= 0 Then
            lpTextArray(Tell) = lpTextArray(Tell) & Script.ProcessTags(StaticDays(hData).Text, lDate, StaticDays(hData).RemDate)
        End If
    End If
    
    If Settings(2) Then
        MoonVal(Tell) = Script.GetMoonPhase(lDate)
    End If

    ExtFlag(Tell) = IIf(Search(Users(UserID).DataRem, lDate, 1) >= 0, 1, 0) Or IIf(Search(Users(UserID).DataTasks, lDate, 1) >= 0, 2, 0)
Next

' Set extra variables
uscCalender.FontSize = IIf(Script.LagreFont, 11, 8)

End Sub

Private Sub vscRem_Change()

ProcessControls "savetxtrem"
ProcessControls "lbltimer"

vscRem.Tag = vscRem.Value

End Sub

Private Sub vscTasks_Change()

ProcessControls "savetxttasks"
ProcessControls "txttask"

vscTasks.Tag = vscTasks.Value

End Sub
