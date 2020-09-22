VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmGuard 
   Caption         =   "Directory Guard"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   Icon            =   "frmGuard20.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Caption         =   " All thats hidden "
      Height          =   3615
      Left            =   0
      TabIndex        =   31
      Top             =   4560
      Visible         =   0   'False
      Width           =   8415
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   2880
      End
      Begin VB.FileListBox HiddenFilelist 
         Height          =   2040
         Hidden          =   -1  'True
         Left            =   120
         System          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.ListBox lstDateTime 
         Height          =   2010
         Left            =   2280
         TabIndex        =   35
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox lstAttribute 
         Height          =   2010
         Left            =   4560
         TabIndex        =   34
         Top             =   600
         Width           =   495
      End
      Begin VB.ListBox lstFileSize 
         Height          =   2010
         Left            =   3600
         TabIndex        =   33
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox lstDirs 
         Height          =   2010
         Left            =   5160
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Timer1"
         Height          =   255
         Left            =   840
         TabIndex        =   38
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Files in directorie                        Filedate                FileSize         Attrib     Directorys"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   6375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.Frame Frame7 
         Caption         =   " Logfile "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4215
         Left            =   5640
         TabIndex        =   17
         Top             =   120
         Width           =   5055
         Begin VB.CommandButton cmdInsert 
            Caption         =   "Add"
            Height          =   255
            Left            =   1920
            TabIndex        =   30
            ToolTipText     =   "Add a commentline to logfile"
            Top             =   3840
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   255
            Left            =   1320
            TabIndex        =   27
            ToolTipText     =   "Save the logfile"
            Top             =   3840
            Width           =   495
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Exit"
            Height          =   255
            Left            =   3960
            TabIndex        =   23
            ToolTipText     =   "xit program"
            Top             =   3840
            Width           =   975
         End
         Begin VB.CommandButton cmdClearLoG 
            Caption         =   "Clear"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Clear the logfile"
            Top             =   3840
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            Height          =   255
            Left            =   720
            TabIndex        =   21
            ToolTipText     =   "Print the logfile"
            Top             =   3840
            Width           =   495
         End
         Begin RichTextLib.RichTextBox rtbChangedfiles 
            Height          =   3495
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "DirGuard Logfile"
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   6165
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmGuard20.frx":030A
         End
         Begin VB.Label Label5 
            Caption         =   "Changes"
            Height          =   255
            Left            =   2880
            TabIndex        =   26
            Top             =   3840
            Width           =   855
         End
         Begin VB.Label lblChanges 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   2400
            TabIndex        =   25
            Top             =   3840
            Width           =   375
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Guarding "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   2880
         TabIndex        =   16
         Top             =   3600
         Width           =   2655
         Begin VB.CommandButton cmdGuardStop 
            Caption         =   "&Stop"
            Height          =   375
            Left            =   1440
            TabIndex        =   20
            ToolTipText     =   "Stop Guarding"
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "&Start"
            Height          =   360
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "Start Guarding"
            Top             =   260
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Result + Status "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   2880
         TabIndex        =   11
         Top             =   2160
         Width           =   2655
         Begin VB.TextBox txtPasses 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtChanges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "No changes"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Idle"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Runs"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   980
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Change"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   620
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Status"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   260
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Settings "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   2880
         TabIndex        =   6
         Top             =   960
         Width           =   2655
         Begin VB.TextBox txtTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "5 Sec"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtFiles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Update time"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Checked files #"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Refresh time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   2880
         TabIndex        =   4
         Top             =   120
         Width           =   2655
         Begin MSComctlLib.Slider Slider1 
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   260
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   10
            SmallChange     =   5
            Min             =   5
            Max             =   30
            SelStart        =   5
            Value           =   5
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Navigation "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2655
         Begin VB.FileListBox lstFiles 
            Appearance      =   0  'Flat
            Height          =   1395
            Hidden          =   -1  'True
            Left            =   120
            System          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2640
            Width           =   2415
         End
         Begin VB.DirListBox lstMap 
            Appearance      =   0  'Flat
            Height          =   1665
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   2415
         End
         Begin VB.DriveListBox Drivestation 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "frmGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoChange As Boolean                     ' Signal changes
Dim GuardStart As Boolean                   ' Signals Guard-start


' because the code got to complex the refresh / check part of the code is
' completely rewritten.
' I have tested the program on a Win98SE and WinNT PC and it worked just fine but i
' can not guarantee that it always will (usage at own risc !)
' In the rootdir of WinNT there is a system (.sys) file that gives a lot of problems
' but the name of this file depends on the version / servicepack number or language
' and its even configurable by the installer so NT-users dont monitor the ROOT !
' You can include a check to not monitor fileextentions SYS or just the 1 file to
' solve this problem.

' The program is written in seperate sub's so a beginner can read it more easely.
' Yes...some (or all) subroutines can be made into 1 or even in a module but this
' would becom unreadable for the beginner.

' Finaly, this is the last version, i have learned a lot making it and its up to you
' to make your own versions.
' I am starting on a new version using a datagrid or an multi-dimentional array for
' intermediates (using a .BAS file for all the checks).
' All who rated this program and send me comments THANX !!!

' 8 Feb 2002    Jay (Jan)  www.cybsoft.nl / cybsoft@bigfoot.com

Private Sub Form_Activate()
Me.Caption = " Directory Guard V" & App.Major & "." & App.Minor
txtFiles.Text = lstFiles.ListCount
GuardStart = False
NoChange = True
End Sub

Private Sub Drivestation_Change()                   ' we change drive
Call StopTimer                                      ' FIRST stop the timer
txtStatus.Text = "Guard stopped"                    ' change the status
txtStatus.Refresh
    On Error GoTo error                             ' in case a disk is not available
    lstMap.Path = Drivestation.Drive                ' set the directory for the map-list
    Exit Sub
error:                                              ' disk was not available
    Dim answer As Integer
    answer = MsgBox(Err.Description, 5, "Device error !")
    If Annswer = 4 Then Resume                      ' they pressed ok
End Sub



Private Sub lstMap_Change()                         ' the map-information
txtStatus.Text = "Guard stopped"                    ' change the label
txtStatus.Refresh
Call StopTimer                                      ' if not stopped...stop it now
lstMap.Refresh                                      ' refresh it
lstFiles.Path = lstMap.Path                         ' set the map for the filelisting
HiddenFilelist.Path = lstMap.Path                   ' needed to detect what file is moved
txtFiles.Text = lstFiles.ListCount                  ' put number of files on the form
End Sub
Private Sub cmdStart_Click()
GuardStart = True                                   ' Signal start
lstMap.Enabled = False                              ' Dont navigate now
Drivestation.Enabled = False                        ' ,,  ,,  ,,  ,,
cmdClearLoG.Enabled = False                         ' no log-operations now
cmdPrint.Enabled = False
cmdSave.Enabled = False
cmdStart.Enabled = False                            ' no need to start twice
txtStatus.Text = "Guard started"                    ' what are we doing
txtPasses.Text = "0"                                ' reset counter
txtTime.Text = Slider1.Value & " Sec"               ' Adjust value on form
txtTime.Refresh
Timer1.Interval = Slider1.Value * 1000              ' the delaytime
Call LogHeader                                      ' make a log-header
Call CollectData                                    ' collect the data and store information

End Sub
Private Sub cmdGuardStop_Click()
txtStatus = "Wait for run-end"                      ' takes some time
txtStatus.Refresh
Call StopTimer                                      ' stop timer
cmdStart.Enabled = True
lstMap.Enabled = True                               ' Navigation OK
Drivestation.Enabled = True
cmdClearLoG.Enabled = True                         ' log-operations OK
cmdPrint.Enabled = True
cmdSave.Enabled = True
Call CollectData                                    ' collect the latest data
txtStatus.Text = "Guard stopped"                    ' adjust the label
Call LogFooter
Me.Refresh
End Sub

Private Sub Timer1_Timer()
txtStatus.Text = "Checking"                         ' what are we doing
txtStatus.Refresh                                   '
If Drivestation.Enabled = True Then                 ' prevent last timerevent when
    Exit Sub                                        ' Guard-Stop is pressed
End If                                              ' If not last event is finished before
                                                    ' Timer1.interval = 0
NoChange = True                                     ' needed to signal change
txtChanges = "No Changes found"                     ' We start "blank"
lstFiles.Refresh
Call CheckAdditions                                 ' is a file added ?
Call CheckDeletions                                 ' is a file deleted ?
Call CheckSize                                      ' is a file renamed ?
Call CheckDateTime                                  ' have date/time changed ?
Call CheckDirectory                                 ' SubDirs change ?

If NoChange = False Then
txtChanges.Text = "Changes found"
End If
Me.Refresh

Call CollectData                                    ' gather fresh info
txtStatus.Text = "Waiting"                          ' what are we doing
txtStatus.Refresh                                   '
txtPasses.Text = Val(txtPasses.Text) + 1

End Sub
Private Sub CheckAdditions()
If HiddenFilelist.ListCount < lstFiles.ListCount Then
    For Counter = 0 To HiddenFilelist.ListCount - 1 'Get the filenames stored
        Checkitem = HiddenFilelist.List(Counter)
        StoredFiles = StoredFiles & " " & Checkitem
    Next Counter
    
    For Counter = 0 To lstFiles.ListCount - 1
        Checkitem = lstFiles.List(Counter)
        If InStr(StoredFiles, Checkitem) = 0 Then   'this is the addition

            rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  " & Checkitem & "  " & "was Added" & vbCrLf
            lblChanges.Caption = Val(lblChanges.Caption) + 1
        End If
    Next Counter
    NoChange = False                                'there where changes
End If
End Sub

Private Sub CheckDeletions()
If HiddenFilelist.ListCount > lstFiles.ListCount Then
    For Counter = 0 To lstFiles.ListCount - 1 'Get the filenames stored
        Checkitem = lstFiles.List(Counter)
        StoredFiles = StoredFiles & " " & Checkitem
    Next Counter
    
    For Counter = 0 To HiddenFilelist.ListCount - 1
        Checkitem = HiddenFilelist.List(Counter)
        If InStr(StoredFiles, Checkitem) = 0 Then   'this is the Deletion

            rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  " & Checkitem & "  " & "was deleted" & vbCrLf
            lblChanges.Caption = Val(lblChanges.Caption) + 1
        End If
    Next Counter
    NoChange = False                                'there where changes
End If
End Sub
Private Sub CheckDateTime()
If NoChange = False Then Exit Sub                       'no need until next run

For Counter = 0 To lstDateTime.ListCount - 1
    StoreDT = lstDateTime.List(Counter)                'get stored DateTime
    StoreFile = HiddenFilelist.List(Counter)           'and the file it belongs to
    
    
    ToLookFor = lstFiles.Path
    If Right(ToLookFor, 1) = "\" Then
        ToLookFor = ToLookFor & StoreFile              ' the files 1 by 1
        Else
        ToLookFor = ToLookFor & "\" & StoreFile
    End If
    
    LetsFind = Dir(ToLookFor, vbDirectory)            ' first check if file still there !

        If Len(LetsFind) = 0 Then                     ' yep...DELETED !while checking
            GoTo WasMissing                             ' = next counter
        End If
        
    ActualDT = (FileDateTime(ToLookFor))

    If StoreDT = ActualDT Then GoTo WasMissing          'Identical = next counter
    
    If CStr(StoreDT) <> CStr(ActualDT) Then

        rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  " & StoreFile & " Date/Time changed" & vbCrLf
        lblChanges.Caption = Val(lblChanges.Caption) + 1
        NoChange = False
    End If

WasMissing:
Next Counter
End Sub

Private Sub CheckSize()
If NoChange = False Then Exit Sub                       'no need until next run

For Counter = 0 To lstFileSize.ListCount - 1
    StoreSize = Val(lstFileSize.List(Counter))          'get stored size
    StoreFile = HiddenFilelist.List(Counter)            'and the file it belongs to
    
    
    ToLookFor = lstFiles.Path
    If Right(ToLookFor, 1) = "\" Then
        ToLookFor = ToLookFor & StoreFile              ' the files 1 by 1
        Else
        ToLookFor = ToLookFor & "\" & StoreFile
    End If


    LetsFind = Dir(ToLookFor, vbDirectory)              ' first check if file still there !

        If Len(LetsFind) = 0 Then                       ' yep...DELETED !while checking
            GoTo WasMissing                             '= next counter
        End If
    

    
    ActualSize = Val((FileLen(ToLookFor)))
    If StoreSize = ActualSize Then GoTo WasMissing      'Identical = next counter
    

    
    If StoreSize <> ActualSize Then
        Change = ActualSize - StoreSize
            If Change > 0 Then
                Change = "+" & CStr(Change) & " Bits"
            Else
                Change = "-" & CStr(Change) & " Bits"
            End If
            
        rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  " & StoreFile & " Sizechange : " & Change & vbCrLf
        lblChanges.Caption = Val(lblChanges.Caption) + 1
        NoChange = False
    End If

WasMissing:
Next Counter
End Sub
Private Sub CheckDirectory()
lstMap.Refresh                                          ' Get fresh data

If lstMap.ListCount > lstDirs.ListCount Then            ' we have additions
        
For Counter = 0 To lstDirs.ListCount - 1                ' quick and dirty
    Buffer = Buffer & lstDirs.List(Counter)             ' store in string
Next Counter
        
For Counter = 0 To lstMap.ListCount - 1
    NewDirInfo = Trim(lstMap.List(Counter))             ' fresh info
    If InStr(Buffer, NewDirInfo) = 0 Then               ' its him
        Position = InStrRev(NewDirInfo, "\")            ' find last slash
        Checkitem = Right(NewDirInfo, Len(NewDirInfo) - Position)
                                                        ' we now have only the name
        rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  Directory :" & Checkitem & "  " & "Added" & vbCrLf
        lblChanges.Caption = Val(lblChanges.Caption) + 1
        NoChange = False
        GoTo WayOut
    End If
Next Counter
End If

If lstMap.ListCount < lstDirs.ListCount Then            ' we have deletions

For Counter = 0 To lstMap.ListCount - 1                ' quick and dirty
    Buffer = Buffer & lstMap.List(Counter)             ' store in string
Next Counter
        
For Counter = 0 To lstDirs.ListCount - 1
    NewDirInfo = Trim(lstDirs.List(Counter))            ' fresh info
    If InStr(Buffer, NewDirInfo) = 0 Then               ' its him
        Position = InStrRev(NewDirInfo, "\")            ' find last slash
        Checkitem = Right(NewDirInfo, Len(NewDirInfo) - Position)
                                                        ' we now have only the name
        rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  Directory :" & Checkitem & "  " & "Removed" & vbCrLf
        lblChanges.Caption = Val(lblChanges.Caption) + 1
        NoChange = False
        GoTo WayOut
    End If
Next Counter
End If

If lstMap.ListCount = lstDirs.ListCount Then                ' Do we have renames ?

    For Counter = 0 To lstDirs.ListCount - 1
        StoredDir = Trim(lstDirs.List(Counter))             ' why not trim ?
        NewDirInfo = Trim(lstMap.List(Counter))
        
        If NewDirInfo <> StoredDir Then                     ' Yep...renamed
            Position = InStrRev(NewDirInfo, "\")            ' find last slash
            Checkitem = Right(NewDirInfo, Len(NewDirInfo) - Position)
            
            Position = InStrRev(StoredDir, "\")            ' find last slash
            ResultItem = Right(StoredDir, Len(StoredDir) - Position)
        
        rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  Directory :" & ResultItem & "  " & " Renamed in " & Checkitem & vbCrLf
        lblChanges.Caption = Val(lblChanges.Caption) + 1
        NoChange = False
        GoTo WayOut
    End If
Next Counter
End If

WayOut:
If NoChange = False Then txtChanges.Text = "Changes found"  ' Change signal
lstDirs.Clear                                               ' until new data

End Sub

Private Sub CollectData()
txtStatus.Text = "Collecting"                           ' what are we doing
txtStatus.Refresh                                       '
lstDateTime.Clear                                       ' first clear them lists
lstFileSize.Clear
lstAttribute.Clear
txtFiles.Text = lstFiles.ListCount
lstMap.Refresh                                          ' this is our new data
HiddenFilelist.Path = lstFiles.Path                       ' now in the hidden box
HiddenFilelist.Refresh                                  ' also refreshed

FilePath = HiddenFilelist.Path & "\"                    ' again..the filepath

If Right(FilePath, 2) = "\\" Then                       ' prevent \\ on rootdir
    FilePath = Left(FilePath, (Len(FilePath) - 1))
End If


For Counter = 0 To HiddenFilelist.ListCount - 1
        CurrentItem = Trim(HiddenFilelist.List(Counter))
        If CurrentItem = "" Then GoTo WayOut            ' if during operations a file
                                                        ' is added but not yet in filelists
                                                        ' chance is SLIM but..who knows
        
    ToLookFor = FilePath & CurrentItem                  ' the files 1 by 1

        
        lstDateTime.AddItem (FileDateTime(ToLookFor))  ' Store date and time
        lstFileSize.AddItem (FileLen(ToLookFor))       ' Store size
        lstAttribute.AddItem (GetAttr(ToLookFor))       ' store attribute
        
WasMissing:                                             ' if deleted during collecting
 Next Counter                                           ' next file
 
 
' ------------------------------------------------------------------------------------
WayOut:
                                                        ' later added store directorys
lstMap.Refresh                                          ' Get fresh info
lstDirs.Clear                                           ' Clear old info

For Counter = 0 To lstMap.ListCount - 1                 ' Count directorys
    storeDir = lstMap.List(Counter)                     ' get items
    lstDirs.AddItem (Trim(storeDir))                    ' store it
Next Counter                                            ' complete directorypaths are stored
End Sub

Private Sub Slider1_Change()                ' change the update-time for the timer
Dim updatetime As Integer                   ' there is no need for all declarations
                                            ' its best to get used to it and use them
updatetime = Slider1.Value                  ' whats the value of the slider ?

If GuardStart = True Then                   ' unless we guard . . .  !!
    Timer1.Interval = (updatetime * 1000)   ' make seconds from milliseconds
    txtTime = CStr(updatetime) & " Sec"     ' put it in a string
    rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  Refreshtime changed now : " & txtTime.Text & vbCrLf
End If

End Sub
Private Sub LogHeader()
Message1 = "Directory Guard V" & App.Major & "." & App.Minor & vbCrLf
Message2 = "GuardReport for : " & lstMap.Path & vbCrLf
Message3 = "Guarding started on : " & Date & " At : " & Time & " Refresh-Time : " & txtTime.Text & vbCrLf
Message4 = "==================================================" & vbCrLf
rtbChangedfiles.Text = rtbChangedfiles.Text & Message1 & Message2 & Message3 & Message4
End Sub
Private Sub LogFooter()
Message1 = "==================================================" & vbCrLf
Message2 = "Guarding stopped on : " & Date & " At : " & Time & vbCrLf
Message3 = "==================================================" & vbCrLf
rtbChangedfiles.Text = rtbChangedfiles.Text & Message1 & Message2 & Message3
End Sub
Private Sub cmdClearLoG_Click()
rtbChangedfiles.Text = ""
lblChanges.Caption = ""
cmdStart.SetFocus           ' the buttons are so small, the caption is deformed
End Sub                     ' seting the focus on another button resolves this
Private Sub cmdPrint_Click()
    rtbChangedfiles.SelPrint (Printer.hDC)  ' could be better but it works
    Printer.EndDoc                          ' just print the lof and eject page
    cmdStart.SetFocus                       ' Buttoncaption restored
End Sub
Private Sub cmdInsert_Click()
FrmNote.Show
cmdGuardStop.SetFocus                       'small button
End Sub
Private Sub StopTimer()     ' this sub is seperate because its used on several occasions
Timer1.Interval = 0         ' Making the interval = 0 stops the timer
End Sub
Private Sub cmdSave_Click()
frmSave.Show                ' show the form containing a dirty savefile
End Sub

Private Sub cmdExit_Click()
Call StopTimer
Unload FrmNote
Unload frmSave
Unload frmGuard
End
End Sub

Private Sub txtTime_Change()
If txtStatus.Text = "Wait for refresh" Or txtStatus.Text = "Updating" Then
' only while running / after first run
rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  Interval Change now : " & txtTime.Text & vbCrLf
End If
End Sub
