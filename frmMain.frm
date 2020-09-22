VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Wallpaper Cycle"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPrev 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1152
      Left            =   9750
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1536
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "File Selection"
      TabPicture(0)   =   "frmMain.frx":2892
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picFrame"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Wallpaper List"
      TabPicture(1)   =   "frmMain.frx":28AE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraWL"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraWL 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6210
         Left            =   -74880
         TabIndex        =   9
         Top             =   600
         Width           =   10950
         Begin VB.CommandButton cmdReorder 
            Caption         =   "Re-order List"
            Height          =   375
            Left            =   9615
            TabIndex        =   16
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton cmdSetWall 
            Caption         =   "Set Wallpaper"
            Height          =   375
            Left            =   9615
            TabIndex        =   12
            ToolTipText     =   "Double clicking on an image in the wallpaper list will also set it as the wallpaper."
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   9615
            TabIndex        =   11
            Top             =   0
            Width           =   1335
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Remove All"
            Height          =   375
            Left            =   9615
            TabIndex        =   10
            Top             =   480
            Width           =   1335
         End
         Begin MSComctlLib.ListView lstFiles 
            Height          =   6210
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   9420
            _ExtentX        =   16616
            _ExtentY        =   10954
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            OLEDropMode     =   1
            _Version        =   393217
            Icons           =   "imgListW"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            OLEDropMode     =   1
            NumItems        =   0
         End
      End
      Begin VB.PictureBox picFrame 
         AutoRedraw      =   -1  'True
         Height          =   2055
         Left            =   9480
         ScaleHeight     =   133
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
         TabIndex        =   1
         Top             =   4680
         Visible         =   0   'False
         Width           =   1695
         Begin VB.PictureBox picThumb 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1152
            Left            =   0
            ScaleHeight     =   77
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   102
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   1536
         End
         Begin VB.PictureBox picSrc 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   0
            ScaleHeight     =   33
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   41
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Timer tmrLoadThumbnails 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   1200
            Top             =   1440
         End
         Begin MSComctlLib.ImageList imgListW 
            Left            =   0
            Top             =   1320
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imgList 
            Left            =   600
            Top             =   1320
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
      End
      Begin VB.Frame fraFS 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6210
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   10935
         Begin VB.PictureBox picSplitter 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            FillColor       =   &H00808080&
            Height          =   240
            Left            =   9480
            ScaleHeight     =   104.506
            ScaleMode       =   0  'User
            ScaleWidth      =   780
            TabIndex        =   15
            Top             =   2400
            Visible         =   0   'False
            Width           =   72
         End
         Begin WallpaperCycle.Treefolder dlb 
            Height          =   6210
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   10954
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmMain.frx":28CA
            LabelEdit       =   1
            Indentation     =   150.236
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add All"
            Height          =   375
            Left            =   9600
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   9600
            TabIndex        =   6
            Top             =   0
            Width           =   1335
         End
         Begin MSComctlLib.ListView flb 
            Height          =   6210
            Left            =   2640
            TabIndex        =   8
            Top             =   0
            Width           =   6780
            _ExtentX        =   11959
            _ExtentY        =   10954
            Arrange         =   2
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            Icons           =   "imgList"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Image imgSplitter 
            Height          =   6210
            Left            =   2580
            MousePointer    =   9  'Size W E
            Top             =   0
            Width           =   60
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuThumbnails 
         Caption         =   "&Thumbnails"
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "&Icons"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuTasks 
      Caption         =   "&Tasks"
      Begin VB.Menu mnuRecreateThumbs 
         Caption         =   "&Recreate Thumbs"
      End
      Begin VB.Menu mnuWallCheck 
         Caption         =   "&Wallpaper Check"
      End
   End
   Begin VB.Menu mnuHelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260

Private Type FILETIME
       dwLowDateTime As Long
       dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
       dwFileAttributes As Long
       ftCreationTime As FILETIME
       ftLastAccessTime As FILETIME
       ftLastWriteTime As FILETIME
       nFileSizeHigh As Long
       nFileSizeLow As Long
       dwReserved0 As Long
       dwReserved1 As Long
       cFileName As String * MAX_PATH
       cAlternate As String * 14
End Type

Dim EnablePreview As Boolean
Dim Filename As String
Dim INIPath As String
Dim flbList As New Collection
Dim Files As New Collection
Dim Icons() As Long
Dim FirstVisible As Long

Private Sub cmdAdd_Click()
    On Error Resume Next
    Dim i As Long, j As Long
    Dim Path As String
    Dim LCount As Long

    Path = dlb.Path
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    For i = 1 To flb.ListItems.Count
        If flb.ListItems(i).Selected = True Then LCount = LCount + 1
    Next
    
    If LCount = 0 Then Exit Sub
    
    frmProgress.Show
    frmProgress.Refresh
    
    picPrev.BackColor = vbWhite
    
    For i = 1 To flb.ListItems.Count
        If flb.ListItems(i).Selected = True Then
            For j = 1 To lstFiles.ListItems.Count
                If LCase$(flb.ListItems(i)) = LCase$(lstFiles.ListItems(j)) Then
                    GoTo NextFile
                End If
            Next
            
            Files.Add Path & flb.ListItems(i)

            If GetSetting(AppName, "Options", "Thumbnail", False) = True Then
                If GetAttr(ThumbPath & flb.ListItems(i)) = vbError Then
                    If mnuThumbnails.Checked = True Then
                        If Icons(i) = 0 Then GoTo IconMode
                        imgListW.ListImages.Add , , imgList.ListImages(Icons(i)).Picture
                        If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                        lstFiles.ListItems.Add , , flb.ListItems(i), imgListW.ListImages.Count
                        SavePicture imgListW.ListImages(imgListW.ListImages.Count).Picture, ThumbPath & flb.ListItems(i)
                    Else
IconMode:
                        DrawPreview i, flbList
                        imgListW.ListImages.Add , , picPrev.Image
                        If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                        lstFiles.ListItems.Add , , flb.ListItems(i), imgListW.ListImages.Count
                        SavePicture picPrev.Image, ThumbPath & flb.ListItems(i)
                    End If
                
                Else
                    picPrev.Picture = LoadPicture(ThumbPath & flb.ListItems(i))
                    imgListW.ListImages.Add , , picPrev.Picture
                    If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                    lstFiles.ListItems.Add , , flb.ListItems(i), imgListW.ListImages.Count
                End If
            Else
                If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                lstFiles.ListItems.Add , , flb.ListItems(i), 1
            End If

NextFile:
            frmProgress.pb = i / LCount * 100
            frmProgress.Refresh
        End If
    Next
    
    picPrev.BackColor = vbButtonFace
    
    WriteINI
    
    Unload frmProgress
    
    lstFiles.Arrange = lvwAutoTop
    lstFiles.Refresh
End Sub

Private Sub cmdAddAll_Click()
    On Error Resume Next
    
    Dim i As Long, j As Long
    Dim Path As String
    
    If flb.ListItems.Count = 0 Then Exit Sub
        
    Path = dlb.Path
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    frmProgress.Show
    frmProgress.Refresh
    
    picPrev.BackColor = vbWhite
    
    For i = 1 To flb.ListItems.Count
        
        For j = 1 To lstFiles.ListItems.Count
            If LCase$(flb.ListItems(i)) = LCase$(lstFiles.ListItems(j)) Then
                GoTo NextFile
            End If
        Next

        Files.Add Path & flb.ListItems(i)
        
        If GetSetting(AppName, "Options", "Thumbnail", False) = True Then
            If GetAttr(ThumbPath & flb.ListItems(i)) = vbError Then
                If mnuThumbnails.Checked = True Then
                    If Icons(i) = 0 Then GoTo IconMode
                    imgListW.ListImages.Add , , imgList.ListImages(Icons(i)).Picture
                    If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                    lstFiles.ListItems.Add , , flb.ListItems(i), imgListW.ListImages.Count
                    SavePicture imgListW.ListImages(imgListW.ListImages.Count).Picture, ThumbPath & flb.ListItems(i)
                Else
IconMode:
                    DrawPreview i, flbList
                    imgListW.ListImages.Add , , picPrev.Image
                    If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                    lstFiles.ListItems.Add , , flb.ListItems(i), imgListW.ListImages.Count
                    SavePicture picPrev.Image, ThumbPath & flb.ListItems(i)
                End If
            
            Else
                picPrev.Picture = LoadPicture(ThumbPath & flb.ListItems(i))
                imgListW.ListImages.Add , , picPrev.Picture
                If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                lstFiles.ListItems.Add , , flb.ListItems(i), imgListW.ListImages.Count
            End If
        
        Else
            If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
            lstFiles.ListItems.Add , , flb.ListItems(i), 1
        End If
        
NextFile:
        
        frmProgress.pb = i / flb.ListItems.Count * 100
        frmProgress.Refresh
        
    Next
    
    picPrev.BackColor = vbButtonFace
    
    WriteINI
    
    Unload frmProgress
    
    lstFiles.Arrange = lvwAutoTop
    lstFiles.Refresh
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    Dim i As Long
    For i = lstFiles.ListItems.Count To 1 Step -1
        If lstFiles.ListItems(i).Selected = True Then
            Kill ThumbPath & lstFiles.ListItems(i)
            lstFiles.ListItems.Remove (i)
            Files.Remove (i)
                        
            Set lstFiles.Icons = Nothing
                        
            If GetSetting(AppName, "Options", "Thumbnail", False) = True Then
                If imgListW.ListImages.Count > 0 Then
                    imgListW.ListImages.Remove (i)
                End If
            End If
            
            lstFiles.Refresh
        End If
    Next
    
    Set lstFiles.Icons = imgListW
    
    If GetSetting(AppName, "Options", "Thumbnail", False) = True Then
        For i = 1 To lstFiles.ListItems.Count
            lstFiles.ListItems(i).Icon = i
        Next
    Else
        For i = 1 To lstFiles.ListItems.Count
            lstFiles.ListItems(i).Icon = 1
        Next
    End If
    
    WriteINI
    
    lstFiles.Arrange = lvwAutoTop
    lstFiles.Refresh
End Sub

Private Sub cmdRemoveAll_Click()
    On Error Resume Next
    Dim i As Long
    Dim Ans As Integer
    
    Ans = MsgBox("Are you sure you want to remove the files from your wallpaper list. Removing items from the wallpaper list will NOT affect the actual images.", vbInformation + vbYesNo)
    
    If Ans = vbNo Then Exit Sub
    
    lstFiles.ListItems.Clear
    Set lstFiles.Icons = Nothing
        
    If GetSetting(AppName, "Options", "Thumbnail", False) = True Then
        imgListW.ListImages.Clear
    End If
    
    Kill ThumbPath & "*.*"
    
    For i = Files.Count To 1 Step -1
        Files.Remove (i)
    Next
    
    WriteINI
    
End Sub

Private Function GetVisibleCount(lstView As ListView) As Integer
    Dim lstItem As New Collection
    Dim x As Single, y As Single
    Dim i As Integer
    Dim curLI As ListItem
    
    For x = 0 To flb.Width Step 500
        For y = 0 To flb.Height Step 250
            Set curLI = lstView.HitTest(x, y)
            
            If Not curLI Is Nothing Then
                For i = 1 To lstItem.Count
                    If lstItem(i) = curLI.Index Then GoTo NextItem
                Next
                
                If lstView.GetFirstVisible.Index > curLI.Index Then GoTo NextItem
                
                lstItem.Add curLI.Index
            End If
NextItem:
        Next
    Next
    
    GetVisibleCount = lstItem.Count
End Function

Private Sub cmdReorder_Click()
    Dim i As Long
    Set frmReorder.tmpFiles = Files
    Load frmReorder
    frmReorder.Show
End Sub

Private Sub cmdSetWall_Click()
    On Error Resume Next
    If IsError(LoadPicture(Files(lstFiles.SelectedItem.Index))) = True Then Exit Sub
    SavePicture LoadPicture(Files(lstFiles.SelectedItem.Index)), AppPath & "Wallpaper.bmp"
    
    Call SetWallMode(AppPath & "Wallpaper.bmp")
    
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, AppPath & "Wallpaper.bmp", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
    
    If GetSetting(AppName, "Options", "Order", 1) = 0 Then
        SaveSetting AppName, "Options", "CurrentIndex", lstFiles.SelectedItem.Index + 1
    ElseIf GetSetting(AppName, "Options", "Order", 0) = 1 Then
        SaveSetting AppName, "Options", "CurrentIndex", lstFiles.SelectedItem.Index
    End If
    
    SaveSetting AppName, "Options", "CurrentWallpaper", Files(lstFiles.SelectedItem.Index)
End Sub

Private Sub dlb_PathChanged()
    Set picPrev.Picture = Nothing
    If mnuThumbnails.Checked = True Then
        Call LoadThumbnails
    Else
        Call LoadIcons
    End If
End Sub

Private Sub flb_DblClick()
    Dim pt As POINTAPI
    Dim li As ListItem
    
    Call GetCursorPos(pt)
    Call ScreenToClient(flb.hWnd, pt)
    
    Set li = flb.HitTest(CSng(pt.x * Screen.TwipsPerPixelX), CSng(pt.y * Screen.TwipsPerPixelY))
    If Not li Is Nothing Then
        ShellExecute frmMain.hWnd, "OPEN", flbList(li.Index), 0, "", vbMaximizedFocus
    End If
    
    Set li = Nothing
End Sub

Private Sub flb_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = 0 Then
        If mnuPreview.Checked = True Then
            Call DrawPreview(flb.SelectedItem.Index, flbList)
        End If
    End If
End Sub

Private Sub flb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If mnuPreview.Checked = True Then
        Call DrawPreview(flb.SelectedItem.Index, flbList)
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim Ans As Byte
    Dim i As Long
    If App.PrevInstance = True Then End
    
    AppPath = App.Path
    If Right$(AppPath, 1) <> "\" Then
        AppPath = AppPath & "\"
    End If
    
    ThumbPath = AppPath & "Thumbnails\"
    
    If Command$ = "/setwall" Then
        tabMain.Tab = 1
    End If
    
    'checks to see if the thumbnail directory exists and if not create it.
    If GetAttr(ThumbPath) = vbError Then
        MkDir ThumbPath
    End If
    
    INIPath = AppPath & "Files.ini"
    
    picSplitter.Left = GetSetting(AppName, "Options", "Splitter Left", imgSplitter.Left)
    frmMain.Width = GetSetting(AppName, "Options", "Form Width", 11670)
    frmMain.Height = GetSetting(AppName, "Options", "Form Height", 8025)
    mnuPreview.Checked = GetSetting(AppName, "Options", "Show Preview", True)
    mnuThumbnails.Checked = GetSetting(AppName, "Options", "Thumbnail Viewstyle", False)
    mnuIcons.Checked = GetSetting(AppName, "Options", "Icon Viewstyle", True)
    
    If mnuPreview.Checked = True Then picPrev.Visible = True
    
    dlb.SetPath GetSetting(AppName, "Path", "LastDir", "")
    
    GetEntries INIPath
    
    If GetSetting(AppName, "Options", "Order", 0) = 0 Then
        i = GetSetting(AppName, "Options", "CurrentIndex", 2) - 2
        CurrentFile = lstFiles.ListItems(i)
    End If
    
    EnablePreview = True
    
    If Command$ = "/opts" Then
        frmOptions.Show vbModal
        Unload frmMain
        Unload frmAbout
        Unload frmOptions
        Unload frmProgress
        End
    End If
    
    If frmMain.Visible = False Then frmMain.Visible = True
    Refresh
End Sub

Private Sub GetEntries(File As String)
    On Error Resume Next
    Dim i As Long
    Dim FDT As String 'FileDateTime
    Dim Bin As String
    Dim SPos As Long
    Dim EPos As Long
    Dim curfile As String
    
    Open File For Binary As 1
        Bin = Space(LOF(1))
        Get 1, , Bin
    Close
    
    For i = Files.Count To 1 Step -1
        Files.Remove (i)
    Next
    
    SPos = 1
    EPos = 0
    
    If GetSetting(AppName, "Options", "Thumbnail", False) = False Then
        imgListW.ListImages.Add , , LoadPicture(AppPath & "imgfile.bmp")
        If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
    End If
    
    'retrieve items from the file.
    Do
        EPos = EPos + 1
        EPos = InStr(EPos, Bin, ";")
        If EPos > 0 Then
            curfile = Mid$(Bin, SPos, EPos - SPos)
            FDT = vbNullString
            FDT = FileDateTime(curfile)
            
            If GetSetting(AppName, "Options", "WallCheck", Unchecked) = Checked And FDT = vbNullString Then
                Kill ThumbPath & FileNameFromPath(curfile) 'File doesn't exist so delete thumbnail.
                GoTo SkipFile
            End If
            
            Files.Add curfile
            lstFiles.ListItems.Add , , FileNameFromPath(curfile)
        End If
SkipFile:
        SPos = EPos + 3
    Loop Until EPos = 0
        
    If Files.Count = 0 Then Exit Sub
    
    picPrev.BackColor = vbWhite
    
    frmProgress.Show
    frmProgress.Refresh
    
    For i = 1 To Files.Count
        If GetSetting(AppName, "Options", "Thumbnail", False) = True Then
            If GetAttr(ThumbPath & lstFiles.ListItems(i)) = vbError Then
                Call DrawPreview(i, Files)
                imgListW.ListImages.Add , , picPrev.Image
                If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                lstFiles.ListItems(i).Icon = i
                SavePicture picPrev.Image, ThumbPath & lstFiles.ListItems(i)
            Else
                picPrev.Picture = LoadPicture(ThumbPath & lstFiles.ListItems(i))
                imgListW.ListImages.Add , , picPrev.Picture
                If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                lstFiles.ListItems(i).Icon = i
            End If
        Else
            lstFiles.ListItems(i).Icon = 1
        End If
        
        frmProgress.pb.Value = Int(i / lstFiles.ListItems.Count * 100)
    Next
    
    Unload frmProgress
    picPrev.BackColor = vbButtonFace
    Set picPrev.Picture = Nothing
    
    lstFiles.Arrange = lvwAutoTop
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState = vbNormal Then
        If frmMain.Height < 6000 Then frmMain.Height = 6000
        If frmMain.Width < 8500 Then frmMain.Width = 8500
        tabMain.Width = frmMain.Width - 375
        tabMain.Height = frmMain.Height - 1050
        fraFS.Width = frmMain.Width - 735
        fraFS.Height = frmMain.Height - 1815
        fraWL.Width = frmMain.Width - 735
        fraWL.Height = frmMain.Height - 1815
        flb.Left = picSplitter.Left + 60
        flb.Width = fraFS.Width - 1515 - flb.Left
        flb.Height = fraFS.Height
        dlb.Width = picSplitter.Left
        dlb.Height = fraFS.Height
        cmdAdd.Left = fraFS.Width - 1335
        cmdAddAll.Left = fraFS.Width - 1335
        cmdRemove.Left = fraWL.Width - 1335
        cmdRemoveAll.Left = fraWL.Width - 1335
        cmdSetWall.Left = fraWL.Width - 1335
        cmdReorder.Left = fraWL.Width - 1335
        lstFiles.Width = fraWL.Width - 1515
        lstFiles.Height = fraWL.Height
        picPrev.Left = frmMain.Width - 1920
        imgSplitter.Height = fraFS.Height
        imgSplitter.Left = picSplitter.Left
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
    SaveSetting AppName, "Path", "LastDir", dlb.Path
    SaveSetting AppName, "Options", "Icon Viewstyle", mnuIcons.Checked
    SaveSetting AppName, "Options", "Thumbnail Viewstyle", mnuThumbnails.Checked
    SaveSetting AppName, "Options", "Show Preview", mnuPreview.Checked
    SaveSetting AppName, "Options", "Form Width", frmMain.Width
    SaveSetting AppName, "Options", "Form Height", frmMain.Height
    SaveSetting AppName, "Options", "Splitter Left", imgSplitter.Left
    
    WriteINI
    
    If GetSetting(AppName, "Options", "RunOnce", "X") = "X" Then
        SaveSetting AppName, "Options", "cmbChange", 1
        SaveSetting AppName, "Options", "ChangeDateTime", Date + 1
        SaveSetting AppName, "Options", "Interval", 1
        
        MsgBox "You haven't configured your options. You should do that first before proceeding."
        frmOptions.Show vbModal
    
        If AddedToSTartup = False Then
            Ans = MsgBox("Do you want Wallpaper Cycle to start when your computer does? If you choose no your wallpaper will not be changed by this program.", vbYesNo + vbQuestion)
            If Ans = vbYes Then
                AddToStartup
                Shell AppPath & "Change Wallpaper.exe /allow"
            Else
                RemoveFromStartup
            End If
        End If
    End If
End Sub

Private Sub WriteINI()
    On Error Resume Next
    Dim Bin As String
    Dim i As Long
    
    Kill INIPath
    
    For i = 1 To Files.Count
        Bin = Bin & Files(i)
        If i < Files.Count Then
            Bin = Bin & ";" & vbNewLine
            Else
            Bin = Bin & ";"
        End If
    Next
    
    Open INIPath For Binary As 1
        Put 1, , Bin
    Close
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Button = 1 Then Exit Sub
    
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width, .Height - 20
    End With
    picSplitter.Visible = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    
    If Button = 1 Then
        sglPos = x + imgSplitter.Left
        If sglPos < 500 Then
            picSplitter.Left = 500
        ElseIf sglPos > flb.Left + flb.Width - 500 Then
            picSplitter.Left = flb.Left + flb.Width - 500
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSplitter.Visible = False
    Call Form_Resize
End Sub

Private Sub lstFiles_DblClick()
    Dim pt As POINTAPI
    Dim li As ListItem
    
    Call GetCursorPos(pt)
    Call ScreenToClient(lstFiles.hWnd, pt)
    
    Set li = lstFiles.HitTest(CSng(pt.x * Screen.TwipsPerPixelX), CSng(pt.y * Screen.TwipsPerPixelY))
    If Not li Is Nothing Then
        Call cmdSetWall_Click
    End If
    
    Set li = Nothing
End Sub

Private Sub lstFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If mnuPreview.Checked = True Then
            Call DrawPreview(lstFiles.SelectedItem.Index, Files)
        End If
    End If
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If mnuPreview.Checked = True Then
        Call DrawPreview(lstFiles.SelectedItem.Index, Files)
    End If
End Sub

Private Sub lstFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim colFiles As New Collection
    Dim fExt As String
    Dim i As Long, j As Long
    Dim cfname As String
    
    frmProgress.Show
    frmProgress.Refresh
    
    For i = 1 To Data.Files.Count
        'validation which doesn't allow directories.
        If Len(Dir(Data.Files.Item(i))) <> 0 Then
            fExt = GetFileExtension(Data.Files.Item(i))
            
            'don't allow invalid files
            If (fExt = "jpg") Or (fExt = "jpeg") Or (fExt = "jpe") Or (fExt = "jfif") Or (fExt = "jif") Or (fExt = "gif") Or (fExt = "bmp") Or (fExt = "dib") Then
                colFiles.Add Data.Files.Item(i)
            End If
        End If
    Next i
    
    picPrev.BackColor = vbWhite
    
    For i = 1 To colFiles.Count
        cfname = FileNameFromPath(colFiles(i))
        
        For j = 1 To lstFiles.ListItems.Count
            If LCase$(cfname) = LCase$(lstFiles.ListItems(j)) Then
                GoTo NextFile
            End If
        Next
        
        Files.Add colFiles(i)

        If GetSetting(AppName, "Options", "Thumbnail", False) = True Then
            If GetAttr(ThumbPath & cfname) = vbError Then
                DrawPreview i, colFiles
                imgListW.ListImages.Add , , picPrev.Image
                If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                lstFiles.ListItems.Add , , cfname, imgListW.ListImages.Count
                SavePicture picPrev.Image, ThumbPath & cfname
            Else
                picPrev.Picture = LoadPicture(ThumbPath & cfname)
                imgListW.ListImages.Add , , picPrev.Picture
                If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
                lstFiles.ListItems.Add , , cfname, imgListW.ListImages.Count
            End If
        Else
            If lstFiles.Icons Is Nothing Then lstFiles.Icons = imgListW
            lstFiles.ListItems.Add , , cfname, 1
        End If

NextFile:
        frmProgress.pb = i / colFiles.Count * 100
        frmProgress.Refresh
    Next
    
    picPrev.BackColor = vbButtonFace
    
    WriteINI
    
    Unload frmProgress
    
    lstFiles.Arrange = lvwAutoTop
    lstFiles.Refresh

Exit Sub
CancelDrop:
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload frmMain
End Sub

Private Sub mnuHelp_Click()
    ShellExecute frmMain.hWnd, "", AppPath & "Help.doc", 1, AppPath, 3
End Sub

Private Sub mnuIcons_Click()
    If mnuIcons.Checked = True Then Exit Sub
    
    mnuIcons.Checked = True
    mnuThumbnails.Checked = False
    tmrLoadThumbnails.Enabled = False
        
    mnuRefresh_Click
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuPreview_Click()
    If mnuPreview.Checked = True Then
        mnuPreview.Checked = False
        picPrev.Visible = False
    Else
        mnuPreview.Checked = True
        picPrev.Visible = True
    End If
End Sub

Private Sub mnuRecreateThumbs_Click()
    On Error Resume Next
    If MsgBox("This will recreate the thumbnails for all the images in your wallpaper list. Do you want to do this?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    frmProgress.Show
    frmProgress.lblInfo.Caption = "Recreating thumbnails......."
    frmProgress.Refresh
    Kill ThumbPath & "*.*"
    Call SaveThumbnails(frmProgress, frmProgress.pb, True, True)
    RestartApp
End Sub

Private Sub mnuRefresh_Click()
    Call dlb_PathChanged
End Sub

Private Sub GetFiles(Path As String)
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long, fPath As String, fName As String
   Dim colFiles As Collection
   Dim varFile As Variant
   
   fPath = AddBackslash(Path)
   fName = fPath & "*.*"
   Set colFiles = New Collection
   
   hFile = FindFirstFile(fName, WFD)
   If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
       colFiles.Add fPath & StripNulls(WFD.cFileName)
   End If
   
   While FindNextFile(hFile, WFD)
       If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
           colFiles.Add fPath & StripNulls(WFD.cFileName)
       End If
   Wend
   
   FindClose hFile
   
   For Each varFile In colFiles
       flbList.Add varFile
   Next
   Set colFiles = Nothing
End Sub

Private Function StripNulls(f As String) As String
   StripNulls = Left$(f, InStr(1, f, Chr$(0)) - 1)
End Function

Private Function AddBackslash(S As String) As String
   If Len(S) Then
      If Right$(S, 1) <> "\" Then
         AddBackslash = S & "\"
      Else
         AddBackslash = S
      End If
   Else
      AddBackslash = "\"
   End If
End Function

Function GetFileExtension(FilePath As String) As String
    Dim i As Long
    
    For i = Len(FilePath) To 1 Step -1
        If Mid$(FilePath, i, 1) = "." Then
            GetFileExtension = LCase$(Right$(FilePath, Len(FilePath) - i))
            Exit Function
        End If
    Next
    
    GetFileExtension = vbNullString
End Function

Private Sub mnuThumbnails_Click()
    If mnuThumbnails.Checked = True Then Exit Sub
    
    If GetSetting(AppName, "Options", "Memory Warning", False) = False Then
        MsgBox "Warning, Displaying the images as thumbnails may be slow if your system doesn't have much memory and has a slow processor.", vbInformation
        SaveSetting AppName, "Options", "Memory Warning", True
    End If
    
    mnuThumbnails.Checked = True
    mnuIcons.Checked = False
        
    mnuRefresh_Click
End Sub

Private Sub LoadThumbnails()
    Dim i As Long
    Dim fExt As String
    Dim hHeight As Double, hWidth As Double
    
    For i = flbList.Count To 1 Step -1
        flbList.Remove (i)
    Next
    
    flb.Icons = Nothing
    imgList.ListImages.Clear
    
    flb.ListItems.Clear
    flb.Refresh
    
    imgList.ListImages.Add , , LoadPicture(AppPath & "blank.bmp")
    Set flb.Icons = imgList
    
    GetFiles dlb.Path
    
    For i = flbList.Count To 1 Step -1
        fExt = GetFileExtension(flbList.Item(i))
        If fExt <> "jpg" And fExt <> "jpeg" And fExt <> "jpe" And fExt <> "jfif" And fExt <> "jif" And fExt <> "gif" And fExt <> "bmp" And fExt <> "dib" Then
            flbList.Remove (i)
        End If
    Next
    
    If flbList.Count = 0 Then Exit Sub
    
    For i = 1 To flbList.Count
        flb.ListItems.Add , , FileNameFromPath(flbList(i)), 1
    Next
    
    flb.ListItems(1).EnsureVisible
    
    ReDim Icons(1 To flbList.Count)
    
    For i = 1 To UBound(Icons)
        Icons(i) = 0
    Next
    
    FirstVisible = -1
    
    Refresh
    
    tmrLoadThumbnails.Enabled = True
    
    If flb.ListItems.Count > 2 Then flb.ListItems(3).EnsureVisible
    
    flb.Arrange = lvwAutoTop
    flb.Refresh
End Sub

Private Sub LoadIcons()
    Dim i As Long
    For i = flbList.Count To 1 Step -1
        flbList.Remove (i)
    Next
    
    flb.Icons = Nothing
    imgList.ListImages.Clear
    
    flb.ListItems.Clear
    flb.Refresh
    
    GetFiles dlb.Path
    
    For i = flbList.Count To 1 Step -1
        fExt = GetFileExtension(flbList.Item(i))
        If fExt <> "jpg" And fExt <> "jpeg" And fExt <> "jpe" And fExt <> "jfif" And fExt <> "jif" And fExt <> "gif" And fExt <> "bmp" And fExt <> "dib" Then
            flbList.Remove (i)
        End If
    Next

    imgList.ListImages.Add , , LoadPicture(AppPath & "imgfile.bmp")
    
    flb.Icons = imgList
    
    For i = 1 To flbList.Count
        flb.ListItems.Add , , FileNameFromPath(flbList(i)), 1
    Next
    
    flb.Arrange = lvwAutoTop
    flb.Refresh
End Sub

Private Sub DrawPreview(hIndex As Long, hList As Collection)
    On Error GoTo DrawPreviewError

    Dim hWidth As Long, hHeight As Long
    
    If hIndex = 0 Then Exit Sub
    
    picPrev.Cls
    If Not picPrev.Picture = 0 Then
        Set picPrev.Picture = Nothing
    End If
    
    picSrc.Picture = LoadPicture(hList(hIndex))
    
    hWidth = picSrc.Width
    hHeight = picSrc.Height
    
    If hHeight > 76.8 Then
        hWidth = 76.8 * picSrc.Width / picSrc.Height
        hHeight = 76.8
    End If
    
    If hWidth > 102.4 Then
        hHeight = 102.4 * picSrc.Height / picSrc.Width
        hWidth = 102.4
    End If
    
    picPrev.PaintPicture picSrc, (picThumb.Width - hWidth) / 2, (picThumb.Height - hHeight) / 2, hWidth, hHeight
    
Exit Sub
DrawPreviewError:
picSrc.Picture = LoadPicture(AppPath & "nopreview.bmp")
picPrev.PaintPicture picSrc, 0, 0
End Sub

Function SaveThumbnails(Optional hForm As Form, Optional pb As ProgressBar, Optional DoEventsOn As Boolean, Optional StopOnCancel As Boolean) As Byte
    On Error Resume Next
    Dim i As Long, j As Long
    Dim Path As String
    Dim LCount As Long
    
    Path = dlb.Path
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    LCount = lstFiles.ListItems.Count
    
    picPrev.BackColor = vbWhite
    
    For i = 1 To lstFiles.ListItems.Count
        'If the user has cancelled
        If hForm.Visible = False And StopOnCancel = True Then
            SaveThumbnails = 12
            Exit For
        End If
        
        'If the thumbnail doesn't exit (hasn't been created yet
        If GetAttr(ThumbPath & lstFiles.ListItems(i)) = vbError Then
            DrawPreview i, Files
            SavePicture picPrev.Image, ThumbPath & lstFiles.ListItems(i)
        End If
            
        pb = i / LCount * 100
        pb.Refresh
        
        If DoEventsOn = True Then DoEvents
    Next
    
    picPrev.BackColor = vbButtonFace
End Function


Private Sub mnuWallCheck_Click()
    Dim i As Long
    On Error Resume Next
    
    If MsgBox("This will check the wallpaper list for images that no longer exist and remove them from the wallpaper list. Do you want to do this?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    For i = Files.Count To 1 Step -1
        If GetAttr(Files(i)) = vbError Then
            Kill ThumbPath & lstFiles.ListItems(i)
            Files.Remove (i)
            lstFiles.ListItems.Remove (i)
        End If
    Next
    
    RestartApp
End Sub

Private Sub tmrLoadThumbnails_Timer()
    On Error Resume Next
    Dim i As Long
    Dim lstVisCount As Integer
    Dim HadFocus As Boolean
    
    If tabMain.Tab = 1 Then Exit Sub
    
    If flb.ListItems.Count = 0 Then
        tmrLoadThumbnails.Enabled = False
        Exit Sub
    End If
    
    lstVisCount = flb.GetFirstVisible.Index + GetVisibleCount(flb) - 1
    
    FirstVisible = flb.GetFirstVisible.Index
    
    HadFocus = IIf(ActiveControl Is flb, True, False)
    
StartLoop:
    'from first visible item to the next visible items
    For i = FirstVisible To lstVisCount
        If Icons(i) > 0 Then GoTo LoadNextThumbnail
        
        Err.Number = 0
        
        picSrc.Picture = LoadPicture(flbList(i))
        
        If Err.Number = 481 Then picSrc.Picture = LoadPicture(AppPath & "nopreview.bmp")
        
        hWidth = picSrc.Width
        hHeight = picSrc.Height
    
        If hHeight > 76.8 Then
            hWidth = 76.8 * picSrc.Width / picSrc.Height
            hHeight = 76.8
        End If
    
        If hWidth > 102.4 Then
            hHeight = 102.4 * picSrc.Height / picSrc.Width
            hWidth = 102.4
        End If
    
        picThumb.PaintPicture picSrc, (picThumb.Width - hWidth) / 2, (picThumb.Height - hHeight) / 2, hWidth, hHeight
        imgList.ListImages.Add , , picThumb.Image
        Icons(i) = imgList.ListImages.Count
        
        HookListview flb
        flb.ListItems(i).Icon = Icons(i)
        UnhookListview
        flb.Visible = False
        flb.Visible = True
        
        picThumb.Cls
        
        flb.Refresh
LoadNextThumbnail:
    Next
    
    If HadFocus = True Then flb.SetFocus
    
    'Exit Sub
End Sub
