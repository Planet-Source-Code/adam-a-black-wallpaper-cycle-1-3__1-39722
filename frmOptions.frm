VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6435
   ClientLeft      =   1320
   ClientTop       =   435
   ClientWidth     =   8520
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabOptions 
      Height          =   5535
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Wallpaper Changing"
      TabPicture(0)   =   "frmOptions.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraChanging"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Startup"
      TabPicture(1)   =   "frmOptions.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRem"
      Tab(1).Control(1)=   "cmdAdd"
      Tab(1).Control(2)=   "lblStartup"
      Tab(1).Control(3)=   "lblInfo"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Display"
      TabPicture(2)   =   "frmOptions.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraMode"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Settings"
      TabPicture(3)   =   "frmOptions.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   -74880
         TabIndex        =   36
         Top             =   420
         Width           =   7935
         Begin VB.CheckBox chkHideTray 
            Caption         =   "Hide Wallpaper Cycle tray icon."
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2880
            Width           =   3135
         End
         Begin VB.ComboBox cmbAction 
            Height          =   315
            ItemData        =   "frmOptions.frx":04B2
            Left            =   120
            List            =   "frmOptions.frx":04C5
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   2200
            Width           =   2895
         End
         Begin VB.CheckBox chkWallCheck 
            Caption         =   "Enable Wallpaper Check"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label lblTrayIcon 
            BackStyle       =   0  'Transparent
            Caption         =   "When I double-click on the Wallpaper Cycle tray icon, I want it to:"
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   3735
         End
         Begin VB.Label lblWallCheck 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":0530
            Height          =   855
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.Frame fraMode 
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   -74880
         TabIndex        =   30
         Top             =   420
         Width           =   7935
         Begin MSComctlLib.ProgressBar pbCreateThumbs 
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   4320
            Visible         =   0   'False
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.PictureBox picPreview 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   5520
            ScaleHeight     =   1335
            ScaleWidth      =   1935
            TabIndex        =   31
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optIcon 
            Caption         =   "Display as icons"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Width           =   2415
         End
         Begin VB.OptionButton optThumbnail 
            Caption         =   "Display as thumbnails"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label lblCreateThumbs 
            BackStyle       =   0  'Transparent
            Caption         =   "Creating Thumbnails... Please Wait"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   4080
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.Label lblPreview 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            Height          =   195
            Left            =   5520
            TabIndex        =   33
            Top             =   120
            Width           =   570
         End
         Begin VB.Label lblDisplayInfo 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":05D2
            Height          =   1275
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   4800
         End
      End
      Begin VB.Frame fraChanging 
         BorderStyle     =   0  'None
         Height          =   4980
         Left            =   120
         TabIndex        =   26
         Top             =   420
         Width           =   8055
         Begin VB.CommandButton cmdDefault 
            Caption         =   "&Default"
            Height          =   315
            Left            =   2720
            TabIndex        =   5
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cmbAMPM 
            Height          =   315
            ItemData        =   "frmOptions.frx":06F4
            Left            =   1860
            List            =   "frmOptions.frx":06FE
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1080
            Width           =   735
         End
         Begin VB.ComboBox cmbMM 
            Height          =   315
            ItemData        =   "frmOptions.frx":070A
            Left            =   1020
            List            =   "frmOptions.frx":0732
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   735
         End
         Begin VB.ComboBox cmbHH 
            Height          =   315
            ItemData        =   "frmOptions.frx":0766
            Left            =   120
            List            =   "frmOptions.frx":078E
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   680
         End
         Begin VB.Frame fraSmartsize 
            Caption         =   "Configure Smart Size"
            Height          =   4740
            Left            =   3840
            TabIndex        =   39
            Top             =   120
            Visible         =   0   'False
            Width           =   4095
            Begin VB.CheckBox chkAutoSmart 
               Caption         =   "Use automatic Smart size configuration."
               Height          =   375
               Left            =   120
               TabIndex        =   10
               Top             =   280
               Width           =   3735
            End
            Begin VB.TextBox txtRatio 
               Height          =   285
               Left            =   90
               MaxLength       =   3
               TabIndex        =   11
               Top             =   1440
               Width           =   615
            End
            Begin VB.ComboBox cmbScenario1 
               Height          =   315
               ItemData        =   "frmOptions.frx":07B9
               Left            =   120
               List            =   "frmOptions.frx":07C6
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   3120
               Width           =   1335
            End
            Begin VB.TextBox txtSize 
               Height          =   285
               Left            =   120
               MaxLength       =   3
               TabIndex        =   12
               Top             =   2400
               Width           =   615
            End
            Begin VB.ComboBox cmbScenario2 
               Height          =   315
               ItemData        =   "frmOptions.frx":07E6
               Left            =   120
               List            =   "frmOptions.frx":07F3
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   4200
               Width           =   1335
            End
            Begin VB.Label lblScenario1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Scenario 1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   720
               Width           =   930
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "The image's aspect ratio is within what percentage of variance of the screen aspect ratio?"
               Height          =   495
               Left            =   120
               TabIndex        =   46
               Top             =   960
               Width           =   3855
            End
            Begin VB.Label lblRatioPercent 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "percent"
               Height          =   195
               Left            =   840
               TabIndex        =   45
               Top             =   1485
               Width           =   540
            End
            Begin VB.Label lblSize 
               BackStyle       =   0  'Transparent
               Caption         =   "The resolution of the image is at least what percentage of the screen resolution?"
               Height          =   495
               Left            =   120
               TabIndex        =   44
               Top             =   1920
               Width           =   3855
            End
            Begin VB.Label lblIfScenario1 
               BackStyle       =   0  'Transparent
               Caption         =   "If Scenario 1 occurs, set the wallpaper as:"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   2880
               Width           =   3855
            End
            Begin VB.Label lblSizePercent 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "percent"
               Height          =   195
               Left            =   840
               TabIndex        =   42
               Top             =   2445
               Width           =   540
            End
            Begin VB.Label lblScenario2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Scenario 2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   3720
               Width           =   930
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "If Scenario 1 doesn't occur, set the wallpaper as:"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   3960
               Width           =   3855
            End
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Randomly"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   2100
            Width           =   3255
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "In order from the first file to the last"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Top             =   1860
            Width           =   3255
         End
         Begin VB.ComboBox cmbChange 
            Height          =   315
            ItemData        =   "frmOptions.frx":0813
            Left            =   840
            List            =   "frmOptions.frx":0826
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   1215
         End
         Begin VB.TextBox txtChange 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   0
            Text            =   "1"
            Top             =   420
            Width           =   495
         End
         Begin VB.CheckBox chkManual 
            Caption         =   "DON'T change my wallpaper automatically"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   2520
            Width           =   3615
         End
         Begin VB.ComboBox cmbMode 
            Height          =   315
            ItemData        =   "frmOptions.frx":0858
            Left            =   120
            List            =   "frmOptions.frx":0868
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label lblColon 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   720
            TabIndex        =   49
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblAt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "At the specific time:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label lblSequence 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I want the wallpapers to change:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1575
            Width           =   2310
         End
         Begin VB.Label lblWhenToChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change my wallpaper every:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   1995
         End
         Begin VB.Label lblMode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wallpaper position"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   3000
            Width           =   1305
         End
      End
      Begin VB.CommandButton cmdRem 
         Caption         =   "Delete from startup"
         Height          =   375
         Left            =   -72960
         TabIndex        =   19
         Top             =   1500
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add to startup"
         Height          =   375
         Left            =   -74760
         TabIndex        =   18
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label lblStartup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wallpaper Cycle is:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   50
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":0894
         Height          =   675
         Left            =   -74760
         TabIndex        =   25
         Top             =   540
         Width           =   7665
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   5880
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAutoSmart_Click()
    If chkAutoSmart.Value = Checked Then
        txtSize.Enabled = False
        txtRatio.Enabled = False
        cmbScenario1.Enabled = False
        cmbScenario2.Enabled = False
    Else
        txtRatio.Enabled = True
        txtSize.Enabled = True
        cmbScenario1.Enabled = True
        cmbScenario2.Enabled = True
    End If
End Sub

Private Sub chkHideTray_Click()
    If chkHideTray.Value = Checked Then
        cmbAction.Enabled = False
    Else
        cmbAction.Enabled = True
    End If
End Sub

Private Sub chkManual_Click()
    If chkManual.Value = Checked Then
        optOrder(0).Enabled = False
        optOrder(1).Enabled = False
        cmbChange.Enabled = False
        txtChange.Enabled = False
        cmbHH.Enabled = False
        cmbMM.Enabled = False
        cmbAMPM.Enabled = False
        cmdDefault.Enabled = False
    Else
        optOrder(0).Enabled = True
        optOrder(1).Enabled = True
        cmbChange.Enabled = True
        txtChange.Enabled = True
        cmbHH.Enabled = True
        cmbMM.Enabled = True
        cmbAMPM.Enabled = True
        cmdDefault.Enabled = True
    End If
End Sub

Private Sub cmbChange_Click()
    If cmbChange.ListIndex = 4 Then
        txtChange.Visible = False
        chkManual.Enabled = False
        cmbHH.Visible = False
        cmbMM.Visible = False
        cmbAMPM.Visible = False
        lblAt.Visible = False
        lblColon.Visible = False
        cmdDefault.Visible = False
    Else
        If cmbChange.ListIndex = 2 Or cmbChange.ListIndex = 3 Then
            cmbHH.Visible = True
            cmbMM.Visible = True
            cmbAMPM.Visible = True
            lblAt.Visible = True
            lblColon.Visible = True
            cmdDefault.Visible = True
        Else
            cmbHH.Visible = False
            cmbMM.Visible = False
            cmbAMPM.Visible = False
            lblAt.Visible = False
            lblColon.Visible = False
            cmdDefault.Visible = False
        End If
        
        txtChange.Visible = True
        chkManual.Enabled = True
    End If
End Sub

Private Sub cmbMode_Click()
    If cmbMode.ListIndex = 3 Then
        fraSmartsize.Visible = True
    Else
        fraSmartsize.Visible = False
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo StopAdd
    Dim Ans As Byte
    AddedToSTartup = True
    
    AddToStartup
    
    Ans = MsgBox("You have added Wallpaper Cycle to the StartUp. Do you want to launch it now?", vbYesNo + vbQuestion)
    
    If OnStartup = True Then
        lblStartup.Caption = "Wallpaper Cycle is on the startup"
        lblStartup.ForeColor = vbBlue
    Else
        lblStartup.Caption = "Wallpaper Cycle is not on the startup"
        lblStartup.ForeColor = vbRed
    End If
    
    If Ans = vbYes Then
        Shell AppPath & "Change Wallpaper.exe /allow"
    End If
    
Exit Sub
StopAdd:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdCancel_Click()
    Unload frmOptions
End Sub

Private Sub cmdDefault_Click()
    cmbHH.ListIndex = 11
    cmbMM.ListIndex = 0
    cmbAMPM.ListIndex = 0
End Sub

Private Sub cmdRem_Click()
    RemoveFromStartup
    AddedToSTartup = False
    
    If OnStartup = True Then
        lblStartup.Caption = "Wallpaper Cycle is on the startup"
        lblStartup.ForeColor = vbBlue
    Else
        lblStartup.Caption = "Wallpaper Cycle is not on the startup"
        lblStartup.ForeColor = vbRed
    End If
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    Dim Ans As Byte
    Dim NextChange As String
    Dim Order As Byte
    Dim i As Long
    Dim SaveThumbs As Byte
    Dim Restart As Boolean
    Dim Interval As Long
    Dim ChangeTime As Date
        
    If cmbChange.ListIndex = 2 Or cmbChange.ListIndex = 3 Then
        ChangeTime = cmbHH.List(cmbHH.ListIndex) & ":" & cmbMM.List(cmbMM.ListIndex) & " " & cmbAMPM.List(cmbAMPM.ListIndex)
        SaveSetting AppName, "Options", "ChangeTime", ChangeTime
        SaveSetting AppName, "Options", "cmbHH", cmbHH.ListIndex
        SaveSetting AppName, "Options", "cmbMM", cmbMM.ListIndex
        SaveSetting AppName, "Options", "cmbAMPM", cmbAMPM.ListIndex
    End If
   
    If txtChange.Text = 0 Then
        MsgBox "You have specified an invalid interval between wallpaper changes. Please change this value.", vbExclamation
        Exit Sub
    End If
        
    If cmbChange.ListIndex = 0 Then
        NextChange = CDate(Now + TimeSerial(0, txtChange.Text, 0))
        Interval = txtChange.Text
    End If
        
    If cmbChange.ListIndex = 1 Then
        NextChange = CDate(Now + TimeSerial(txtChange.Text, 0, 0))
        Interval = txtChange.Text
    End If
    
    If cmbChange.ListIndex = 2 Then
        NextChange = CDate((Date + txtChange) & " " & ChangeTime)
        Interval = txtChange.Text
    End If
    
    If cmbChange.ListIndex = 3 Then
        NextChange = CDate((Date + (txtChange) * 7) & " " & ChangeTime)
        Interval = txtChange.Text * 7
    End If
    
    If cmbChange.ListIndex = 4 Then
        NextChange = "Startup"
        Interval = 0
    End If
    
    For i = 0 To 1
        If optOrder(i).Value = True Then Order = i
    Next
    
    If Not optThumbnail.Value = GetSetting(AppName, "Options", "Thumbnail", False) Then
        If optThumbnail.Value = True Then
            MsgBox "Thumbnails must be created for the images in your list. This may take a while depending on how many images you have in your list. If your computer's hard disk stops loading for a long period of time end task to the program.", vbInformation
            lblCreateThumbs.Visible = True
            pbCreateThumbs.Visible = True
            fraMode.Refresh
            tabOptions.Enabled = False
            cmdSave.Enabled = False
            SaveThumbs = frmMain.SaveThumbnails(frmOptions, pbCreateThumbs, True, True)
            If SaveThumbs = 12 Then
                Kill ThumbPath & "*.*"
                Exit Sub
            End If
        Else
            Ans = MsgBox("Do you wish to delete the thumbnails created by Wallpaper Cycle when the Wallpaper list was in thumbnail mode?", vbQuestion + vbYesNo)
            
            If Ans = vbYes Then
                Kill ThumbPath & "*.*"
            End If
        End If
        
        Restart = True
    Else
        Restart = False
    End If
        
    If NextChange <> GetSetting(AppName, "Options", "ChangeDateTime", -1) Then
        SaveSetting AppName, "Options", "ChangeDateTime", NextChange
    End If
    
    If GetSetting(AppName, "Options", "Hide Tray", Unchecked) <> chkHideTray.Value Then
        MsgBox "No changes will occur to the tray icon until the next reboot.", vbInformation
    End If
    
    SaveSetting AppName, "Options", "cmbChange", cmbChange.ListIndex
    SaveSetting AppName, "Options", "txtChange", Val(txtChange)
    SaveSetting AppName, "Options", "Interval", Interval
    SaveSetting AppName, "Options", "Thumbnail", optThumbnail.Value
    SaveSetting AppName, "Options", "Mode", cmbMode.ListIndex
    SaveSetting AppName, "Options", "Manual", chkManual.Value
    SaveSetting AppName, "Options", "WallCheck", chkWallCheck.Value
    SaveSetting AppName, "Options", "Action", cmbAction.ListIndex
    SaveSetting AppName, "Options", "Hide Tray", chkHideTray.Value
    SaveSetting AppName, "Smart Size", "Auto Config", chkAutoSmart.Value
    SaveSetting AppName, "Smart Size", "Ratio", Val(txtRatio.Text)
    SaveSetting AppName, "Smart Size", "Resolution", Val(txtSize.Text)
    SaveSetting AppName, "Smart Size", "Scenario 1", cmbScenario1.ListIndex
    SaveSetting AppName, "Smart Size", "Scenario 2", cmbScenario2.ListIndex
    
    If GetSetting(AppName, "Options", "Order", 0) = 1 And optOrder(0).Value = True Then
        i = GetSetting(AppName, "Options", "CurrentIndex", 0)
        SaveSetting AppName, "Options", "CurrentIndex", i + 1
    End If
    
    If GetSetting(AppName, "Options", "Order", 1) = 0 And optOrder(1).Value = True Then
        i = GetSetting(AppName, "Options", "CurrentIndex", 2)
        SaveSetting AppName, "Options", "CurrentIndex", i - 1
    End If
    
    SaveSetting AppName, "Options", "Order", Order
    
    If Restart = True Then
        RestartApp
    End If
    
    Unload frmOptions
End Sub

Private Sub Form_Load()
    Dim i As Byte
    txtChange.Text = GetSetting(AppName, "Options", "txtChange", 1)
    cmbChange.ListIndex = GetSetting(AppName, "Options", "cmbChange", 2)
    cmbHH.ListIndex = GetSetting(AppName, "Options", "cmbHH", 11)
    cmbMM.ListIndex = GetSetting(AppName, "Options", "cmbMM", 0)
    cmbAMPM.ListIndex = GetSetting(AppName, "Options", "cmbAMPM", 0)
    optOrder(GetSetting(AppName, "Options", "Order", 0)).Value = True
    optThumbnail.Value = GetSetting(AppName, "Options", "Thumbnail", False)
    If optThumbnail.Value = False Then optIcon.Value = True
    cmbMode.ListIndex = GetSetting(AppName, "Options", "Mode", 0)
    chkManual.Value = GetSetting(AppName, "Options", "Manual", Unchecked)
    chkWallCheck.Value = GetSetting(AppName, "Options", "WallCheck", Unchecked)
    cmbAction.ListIndex = GetSetting(AppName, "Options", "Action", 0)
    chkHideTray.Value = GetSetting(AppName, "Options", "Hide Tray", Unchecked)
    chkAutoSmart.Value = GetSetting(AppName, "Smart Size", "Auto Config", Checked)
    txtRatio.Text = GetSetting(AppName, "Smart Size", "Ratio", 10)
    txtSize.Text = GetSetting(AppName, "Smart Size", "Resolution", 60)
    cmbScenario1.ListIndex = GetSetting(AppName, "Smart Size", "Scenario 1", 0)
    cmbScenario2.ListIndex = GetSetting(AppName, "Smart Size", "Scenario 2", 2)
    
    If OnStartup = True Then
        lblStartup.Caption = "Wallpaper Cycle is on the startup"
        lblStartup.ForeColor = vbBlue
    Else
        lblStartup.Caption = "Wallpaper Cycle is not on the startup"
        lblStartup.ForeColor = vbRed
    End If
    
    SaveSetting AppName, "Options", "RunOnce", True
End Sub

Private Sub optIcon_Click()
    If optIcon.Value = True Then
        picPreview.Picture = LoadPicture(AppPath & "picIcons.bmp")
    End If
End Sub

Private Sub optThumbnail_Click()
    If optThumbnail.Value = True Then
        picPreview.Picture = LoadPicture(AppPath & "picThumbs.bmp")
    End If
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    'validation. Only numbers, or backspace is accepted.
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtRatio_KeyPress(KeyAscii As Integer)
    'validation. Only numbers, or backspace is accepted.
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
    'validation. Only numbers, or backspace is accepted.
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
