VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Wallpaper Cycle"
   ClientHeight    =   2205
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   2790
   ControlBox      =   0   'False
   Icon            =   "frmCycle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin ChangeWallpaper.ShellIcon shIcon 
      Left            =   1320
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "frmCycle.frx":014A
      SysMenu         =   0   'False
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   500
      Left            =   600
      Top             =   840
   End
   Begin VB.PictureBox picWallpaper 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer tmrCheck 
      Interval        =   10000
      Left            =   120
      Top             =   840
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuNewWallpaper 
         Caption         =   "&Change"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "&Quick Change"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SENDWININICHANGE = &H2
    
Dim Arg As String
Dim Files As New Collection
Dim CurFile As Long
Dim AppPath As String
Dim INIPath As String
Dim QuickChange As Boolean
Dim Random As New Collection

Private Sub GetEntries(File As String, Optional DoNotEnd As Boolean)
    Dim Bin As String
    Dim SPos As Long
    Dim EPos As Long
    Dim i As Long
    
    Open File For Binary As 1
        Bin = Space(LOF(1))
        Get 1, , Bin
    Close
    
    SPos = 1
    EPos = 0
    
    For i = Files.Count To 1 Step -1
        Files.Remove i
    Next
    
    Do
        EPos = EPos + 1
        EPos = InStr(EPos, Bin, ";")
        If EPos > 0 Then
            Files.Add Mid(Bin, SPos, EPos - SPos)
        End If
        
        SPos = EPos + 3
        
    Loop Until EPos = 0
    
    If Files.Count < 2 And DoNotEnd = False Then End
End Sub

Private Sub Form_Load()
    On Error GoTo ErrFormLoad
    'don't make it visible to the task manager
    App.TaskVisible = False
    
    'Don't let more than one copy run
    If App.PrevInstance = True Then End
    
    AppPath = App.Path
    
    If Right$(AppPath, 1) <> "\" Then
        AppPath = AppPath & "\"
    End If
        
    INIPath = AppPath & "Files.ini"
    
    Arg = LCase$(Command)
    
    'If not run under proper circumstances exit.
    If Arg <> "/allow" And Arg <> "/startup" Then
        Shell "Wallpaper Cycle.exe", vbNormalFocus
        End
    End If
    
    QuickChange = False
    
    If Arg = "/startup" And GetSetting(AppName, "Options", "cmbChange", 2) = 4 Then
        QuickChange = True
    End If
    
    Call tmrCheck_Timer
    
    If GetSetting(AppName, "Options", "Hide Tray", Unchecked) = Unchecked Then
        shIcon.Visible = True
    End If

Exit Sub
ErrFormLoad:
If Err.Number = 53 Then
    MsgBox "Wallpaper Cycle.exe could not be found. Please make sure it is in the same directory as Change Wallpaper.exe", vbExclamation
Else
    MsgBox Err.Description, vbExclamation, "Error Number " & Err.Number
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    shIcon.Visible = False
    End
End Sub

Private Sub mnuChange_Click()
    QuickChange = True
    Call tmrCheck_Timer
End Sub

Private Sub mnuExit_Click()
    Unload frmMain
End Sub

Private Sub mnuInfo_Click()
    On Error Resume Next
    Dim NextChage As String
    Dim NextFile As String
    Dim CurrentFile As String
    Dim FileIndex As Long
    Dim Order As Byte
    Dim ChangeDateTime As String
    
    FileIndex = GetSetting(AppName, "Options", "CurrentIndex", 1)
    Order = GetSetting(AppName, "Options", "Order", 0)
    
    GetEntries INIPath, True
    
    If Order = 0 And FileIndex > 0 And FileIndex - 1 <= Files.Count Then
        If FileIndex > Files.Count Then FileIndex = 1
        NextFile = Files(FileIndex)
    End If
    
    If FileIndex > Files.Count And Order = 0 Then
        NextFile = Files(1)
    End If
    
    If Order = 1 Then
        NextFile = "Random"
    End If
    
    If Files.Count < 2 Then
        NextFile = "N/A"
    End If
    
    CurrentFile = GetSetting(AppName, "Options", "CurrentWallpaper", "Unknown")
    
    ChangeDateTime = GetSetting(AppName, "Options", "ChangeDateTime", 0)
    
    If GetSetting(AppName, "Options", "Manual", Unchecked) = Checked Then
        ChangeDateTime = "N/A"
    End If
    
    MsgBox "The wallpaper will be cycled on: " & ChangeDateTime & vbNewLine & "The next wallpaper will be: " & FileNameFromPath(NextFile) & vbNewLine & "Wallpaper Cycle's current wallpaper is: " & FileNameFromPath(CurrentFile)
End Sub

Private Sub mnuNewWallpaper_Click()
    Shell AppPath & "Wallpaper Cycle.exe /setwall", vbNormalFocus
End Sub

Private Sub mnuOptions_Click()
    Shell AppPath & "Wallpaper Cycle.exe /opts", vbNormalFocus
End Sub

Private Sub shIcon_Click(Button As Integer)
    If Button = 2 Then
        SetForegroundWindow hWnd
        PopupMenu mnuPopup
    End If
End Sub

Private Sub shIcon_DblClick(Button As Integer)
    Dim hResult As Byte
    hResult = GetSetting(AppName, "Options", "Action", 0)
    'choose action according to options.
    If hResult = 0 Then
        Shell AppPath & "Wallpaper Cycle.exe", vbNormalFocus
    ElseIf hResult = 1 Then
        Shell AppPath & "Wallpaper Cycle.exe /setwall", vbNormalFocus
    ElseIf hResult = 2 Then
        Call mnuChange_Click
    ElseIf hResult = 3 Then
        Shell AppPath & "Wallpaper Cycle.exe /opts", vbNormalFocus
    ElseIf hResult = 4 Then
        Call mnuInfo_Click
    End If
End Sub

Private Sub tmrCheck_Timer()
    On Error Resume Next
    Dim Order As Byte
    Dim NewDate As Date
    Dim LoopCount As Long
    Dim i As Long
    
    If GetSetting(AppName, "Options", "Manual", Unchecked) = Checked And QuickChange = False Then Exit Sub
    If GetSetting(AppName, "Options", "ChangeDateTime", 0) = "Startup" And QuickChange = False Then Exit Sub
    
    If Now >= CDate(GetSetting(AppName, "Options", "ChangeDateTime", 0)) Or QuickChange = True Then
        Randomize Timer
        
        For i = Random.Count To 1 Step -1
            Random.Remove (i)
        Next
        
        LoopCount = 0
        
        GetEntries INIPath, True
        
        Order = GetSetting(AppName, "Options", "Order", 0)
        
        If Order = 0 Then
            CurFile = GetSetting(AppName, "Options", "CurrentIndex", 1)
            If CurFile > Files.Count Then CurFile = 1
            SaveSetting AppName, "Options", "CurrentIndex", CurFile + 1
        ElseIf Order = 1 Then
            Do
                CurFile = Int(Rnd * Files.Count) + 1
            Loop Until CurFile <> GetSetting(AppName, "Options", "CurrentIndex", Int(Rnd * Files.Count) + 1)
            Random.Add CurFile
        End If
        
        'conventional Do/Loop loops do not work in this situation. Not sure why
        If IsError(LoadPicture(Files(CurFile))) = True Then
Top:
            LoopCount = LoopCount + 1
            If Order = 0 Then
                CurFile = CurFile + 1
                If CurFile > Files.Count Then CurFile = 1
                SaveSetting AppName, "Options", "CurrentIndex", CurFile + 1
                If LoopCount > Files.Count Then Exit Sub
            ElseIf Order = 1 Then
                Do
                    If Random.Count = Files.Count Then Exit Sub
                    
                    CurFile = Int(Rnd * Files.Count) + 1
                    For i = 1 To Random.Count
                        If Random(i) = CurFile Then GoTo IsDoubled
                    Next
                    
                    Random.Add CurFile
IsDoubled:
                Loop Until CurFile <> GetSetting(AppName, "Options", "CurrentIndex", Int(Rnd * Files.Count) + 1)
            End If
            If IsError(LoadPicture(Files(CurFile))) = True Then GoTo Top
        End If
        
        If Order = 1 Then
            SaveSetting AppName, "Options", "CurrentIndex", CurFile
        End If
        
        picWallpaper.Picture = LoadPicture(Files(CurFile))
        
        SavePicture picWallpaper.Picture, AppPath & "Wallpaper.bmp"
        
        SetWallMode AppPath & "Wallpaper.bmp"
        
        SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, AppPath & "Wallpaper.bmp", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
        
        SaveSetting AppName, "Options", "CurrentWallpaper", Files(CurFile)
        
        If QuickChange = False Then
            If GetSetting(AppName, "Options", "cmbChange", 2) = 0 Then
                NewDate = Now + TimeSerial(0, GetSetting(AppName, "Options", "Interval", 1), 0)
            ElseIf GetSetting(AppName, "Options", "cmbChange", 2) = 1 Then
                NewDate = Now + TimeSerial(GetSetting(AppName, "Options", "Interval", 1), 0, 0)
            ElseIf GetSetting(AppName, "Options", "cmbChange", 2) = 4 Then
                NewDate = "Startup"
            Else
                NewDate = (Date + GetSetting(AppName, "Options", "Interval", 1)) & " " & GetSetting(AppName, "Options", "ChangeTime", "12:00:00 AM")
            End If
            
            SaveSetting AppName, "Options", "ChangeDateTime", NewDate
        Else
            QuickChange = False
        End If
    End If
End Sub

Private Sub tmrUpdate_Timer()
    If GetSetting(AppName, "Options", "Manual", Unchecked) = Checked Then
        shIcon.ToolTipText = "There will be no wallpaper change"
    ElseIf GetSetting(AppName, "Options", "Manual", Unchecked) = Unchecked Then
        shIcon.ToolTipText = "Wallpaper change on " & GetSetting(AppName, "Options", "ChangeDateTime", 0)
    ElseIf GetSetting(AppName, "Options", "cmbChange", 2) = 4 Then
        shIcon.ToolTipText = "The wallpaper will be changed next startup"
    End If
End Sub
 
