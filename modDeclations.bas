Attribute VB_Name = "modDeclations"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const WM_STYLECHANGED = &H7D
Public Const GWL_WNDPROC = (-4)
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
Public AppPath As String
Public ThumbPath As String
Public AddedToSTartup As Boolean
Public gPrevWndProc As Long
Public gHW As Long
Public gLV As ListView

Public Type POINTAPI
   x  As Long
   y  As Long
End Type

Sub AddToStartup()
    Dim wshShell As Object, oShellLink As Object
    Dim strStartUp As String
    Set wshShell = CreateObject("WScript.Shell")
    strStartUp = wshShell.SpecialFolders("StartUp")
    Set oShellLink = wshShell.CreateShortcut(strStartUp & "\Change Wallpaper.lnk")
    
    oShellLink.TargetPath = AppPath & "Change Wallpaper.exe"
    oShellLink.Arguments = "/startup"
    oShellLink.WindowStyle = 1
    oShellLink.IconLocation = AppPath & "Change Wallpaper.exe, 0"
    oShellLink.Description = "Change Wallpaper"
    oShellLink.WorkingDirectory = AppPath
    oShellLink.Save
    Set oShellLink = Nothing
    Set wshShell = Nothing
End Sub

Sub RemoveFromStartup()
    On Error Resume Next
    Dim wshShell As Object
    Dim StartUp As String
    
    Set wshShell = CreateObject("WScript.Shell")
    
    StartUp = wshShell.SpecialFolders("StartUp")
    
    Set oShellLink = Nothing
    Set wshShell = Nothing
    
    Kill StartUp & "\Change Wallpaper.lnk"
End Sub

Function OnStartup() As Boolean
    On Error Resume Next
    Dim wshShell As Object
    Dim StartUp As String
    
    OnStartup = True
    Set wshShell = CreateObject("WScript.Shell")
    
    StartUp = wshShell.SpecialFolders("StartUp")
    
    Set oShellLink = Nothing
    Set wshShell = Nothing
    
    If GetAttr(StartUp & "\Change Wallpaper.lnk") = vbError Then OnStartup = False
End Function

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
   ''Just add other messages if they bug you
   'If uMsg = WM_VSCROLL Or uMsg = WM_HSCROLL Then
   '    'Stop processing, no scrolling allowed
   '    WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
   'Else
   '    'Call prev. window procedure, the original!
   '    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
   'End If
End Function

Public Sub HookListview(lv As ListView)
   If Not lv Is Nothing Then
       If gHW <> 0 Then
           UnhookListview
       End If
       Set gLV = lv
       gHW = gLV.hWnd
       gPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
   End If
End Sub

Public Sub UnhookListview()
   If gHW <> 0 And gPrevWndProc <> 0 Then
       SetWindowLong gHW, GWL_WNDPROC, gPrevWndProc
       gHW = 0
       Set gLV = Nothing
   End If
End Sub

Public Sub RestartApp()
    Unload frmAbout
    Unload frmOptions
    Unload frmProgress
    Unload frmMain
    Load frmMain
End Sub
