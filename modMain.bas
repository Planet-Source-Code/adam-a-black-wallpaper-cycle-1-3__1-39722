Attribute VB_Name = "modMain"
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Const HKEY_CURRENT_USER = &H80000001
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const AppName = "Wallpaper Cycle"

Public x As Long

Public Sub SaveString(hKey As Long, strpath As String, strValue As String, strdata As String)
   Dim keyhand As Long
   Dim r As Long
   x = RegCreateKey(hKey, strpath, keyhand)
   x = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
   x = RegCloseKey(keyhand)
End Sub

Public Sub SetWallMode(WPath As String)
    Dim Mode As Byte
    Dim ScreenHeight As Integer, ScreenWidth As Integer
    Dim RatioPercent As Single, ResPercent As Single
    Dim ImgRatio As Single, ScreenRatio As Single
    Dim Scenario1 As Byte, Scenario2 As Byte
    Dim MinSRatio As Single, MaxSRatio As Single
    Dim MinImgHeight As Single, MinImgWidth As Single
    
    Mode = GetSetting(AppName, "Options", "Mode", 0)
    
    If Mode = 3 Then
        ReadImageInfo WPath
        ImgRatio = ImageWidth / ImageHeight
        'Load the settings.
        If GetSetting(AppName, "Smart Size", "Auto Config", Checked) = Checked Then
            ResPercent = 60
            RatioPercent = 10
            Scenario1 = 0
            Scenario2 = 2
        Else
            ResPercent = GetSetting(AppName, "Smart Size", "Resolution", 60)
            RatioPercent = GetSetting(AppName, "Smart Size", "Ratio", 10)
            Scenario1 = GetSetting(AppName, "Smart Size", "Scenario 1", 0)
            Scenario2 = GetSetting(AppName, "Smart Size", "Scenario 2", 2)
        End If
        
        'Get monitor display settings.
        ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
        ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
        ScreenRatio = Screen.Width / Screen.Height
        
        'Evaluate how the wallpaper should be set.
        MinSRatio = ((100 - RatioPercent) / 100) * ScreenRatio
        MaxSRatio = ((RatioPercent + 100) / 100) * ScreenRatio
        MinImgWidth = ScreenWidth * (ResPercent / 100)
        MinImgHeight = ScreenHeight * (ResPercent / 100)
        
        If ImgRatio >= MinSRatio And ImgRatio <= MaxSRatio And ImageWidth >= MinImgWidth And ImageHeight >= MinImgHeight Then
            Mode = Scenario1
        Else
            Mode = Scenario2
        End If
    End If
    
    If Mode = 0 Then
        'Stretch wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0")
    ElseIf Mode = 1 Then
        'Tile wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "1")
    ElseIf Mode = 2 Then
        'Center wallpaper
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0")
        Call SaveString(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0")
    End If
End Sub

Public Function FileNameFromPath(Filename As String) As String
    Dim SPos As Integer
    For I = Len(Filename) To 1 Step -1
        If Mid(Filename, I, 1) = "\" Then
            FileNameFromPath = Mid(Filename, I + 1)
            Exit Function
        End If
    Next
    
    If FileNameFromPath = vbNullString Then FileNameFromPath = Filename
End Function
