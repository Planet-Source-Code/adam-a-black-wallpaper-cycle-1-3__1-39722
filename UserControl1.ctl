VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Treefolder 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   ScaleHeight     =   2880
   ScaleWidth      =   3930
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.DirListBox Dir1 
      Height          =   288
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.DirListBox CheckForChildDir 
      Height          =   288
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.DirListBox DummyDir 
      Height          =   288
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   852
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2172
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   3836
      _Version        =   393217
      Indentation     =   176
      LineStyle       =   1
      Style           =   7
      ImageList       =   "Default"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList Default 
      Left            =   2760
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0432
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0552
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":09A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0BE6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Treefolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type SECURITY_ATTRIBUTES
nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Dim IsV As Boolean
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event NodeCheck(ByVal Node As MSComctlLib.Node)
Event Collapse(ByVal Node As MSComctlLib.Node)
Event Expand(ByVal Node As MSComctlLib.Node)
Event NodeClick(ByVal Node As MSComctlLib.Node)
Event PathChanged()
Event AfterLabelEdit(Cancel As Integer, NewString As String)
Event BeforeLabelEdit(Cancel As Integer)
Event DriveNotReadyError()
Public Sub CreateNewDirectory(NewDirectory As String)
Dim sDirTest As String
Dim SecAttrib As SECURITY_ATTRIBUTES
Dim bSuccess As Boolean
Dim sPath As String
Dim iCounter As Integer, LastiCounter As Integer
Dim sTempDir As String, Last As String
iFlag = 0
sPath = NewDirectory
If Right(sPath, Len(sPath)) <> "\" Then
    sPath = sPath & "\"
End If
iCounter = 1
Do Until InStr(iCounter, sPath, "\") = 0
    LastiCounter = iCounter
    iCounter = InStr(iCounter, sPath, "\")
    Last = sTempDir
    sTempDir = Left(sPath, iCounter)
    sDirTest = Dir(sTempDir)
    iCounter = iCounter + 1
    'create directory
    SecAttrib.lpSecurityDescriptor = &O0
    SecAttrib.bInheritHandle = False
    SecAttrib.nLength = Len(SecAttrib)
    bSuccess = CreateDirectory(sTempDir, SecAttrib)
    sTempDir = Left(sTempDir, Len(sTempDir) - 1)
    If bSuccess Then
     FolderName = Mid(sTempDir, LastiCounter)
     On Error Resume Next
     TreeView1.Nodes.Add Last, tvwChild, sTempDir, FolderName, 4
    End If
Loop
End Sub
Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
RaiseEvent BeforeLabelEdit(Cancel)
End Sub

Private Sub TreeView1_Click()
RaiseEvent Click
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
 For I = 1 To Node.Children
    TreeView1.Nodes.Remove Node.Child.Index
 Next I
 
 TreeView1.Nodes.Add Node.Key, tvwChild, ""
  
RaiseEvent Collapse(Node)
End Sub

Private Sub TreeView1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
Dim CurrentPath As String, FolderName As String
On Error GoTo ErrorTreeView
DummyDir.Path = Node.Key
DummyDir.Refresh
TreeView1.Nodes(Node.Key).EnsureVisible
If Node.Child.Text = "" Then
 TreeView1.Nodes.Remove Node.Child.Index
 For I = 0 To DummyDir.ListCount - 1
  FolderName = Mid(DummyDir.List(I), Len(DummyDir.Path) + 2)
  If Len(DummyDir.Path) = 3 Then FolderName = Mid(DummyDir.List(I), Len(DummyDir.Path) + 1)
  TreeView1.Nodes.Add DummyDir.Path, tvwChild, LCase(DummyDir.List(I)), FolderName, 4
  CheckForChildDir.Path = DummyDir.List(I) 'checking for childs
  If CheckForChildDir.ListCount > 0 Then
   TreeView1.Nodes.Add LCase(DummyDir.List(I)), tvwChild, ""
   TreeView1.Nodes(LCase(DummyDir.List(I))).ExpandedImage = 5
  End If
 Next I
End If
ErrorTreeView:
If Err.Number = 68 Then
 TreeView1.Nodes(Node.Index).Expanded = False
 RaiseEvent DriveNotReadyError
End If
RaiseEvent Expand(Node)
End Sub
Public Function Path() As String
Path = Dir1.Path
End Function
Private Sub BuildDriveList()
Dim I As Integer
Dim TreePath As String
Dim TreeIcon As Integer
TreeView1.Nodes.Clear
For I = 0 To Drive1.ListCount - 1
TreePath = Left(Drive1.List(I), 1) & ":\"
drivetype = GetDriveType(Drive1.List(I))
Select Case drivetype
 Case 2: TreeIcon = 1
 Case 1, 3: TreeIcon = 2
 Case Else: TreeIcon = 3
End Select
TreeView1.Nodes.Add , , TreePath, UCase(Drive1.List(I)), TreeIcon
TreeView1.Nodes.Add TreePath, tvwChild, ""
Next
End Sub
Public Function SetPath(MyPath As String)
On Error GoTo ErrorTreeView
Dim SubDirNum As Integer, Dummy As Integer, NextSlash As Integer
Dim MyFolder(0 To 20) As String
I = 0
NextSlash = 1
Do
MyFolder(I) = Left(MyPath, InStr(NextSlash, MyPath, "\", 0) - 1)
I = I + 1
NextSlash = InStr(NextSlash, MyPath, "\", 0) + 1
Loop Until InStr(NextSlash, MyPath, "\", 0) = 0
MyFolder(0) = MyFolder(0) & "\"
MyFolder(I) = MyPath
SubDirNum = I
For Dummy = 0 To SubDirNum - 1
 TreeView1.Nodes(MyFolder(Dummy)).Expanded = True
Next Dummy
Dir1.Path = MyFolder(SubDirNum)
TreeView1.Nodes(Dir1.Path).Selected = True
TreeView1.Nodes(Dir1.Path).SelectedImage = 5 - IsV
RaiseEvent PathChanged
ErrorTreeView:
End Function
Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
RaiseEvent NodeCheck(Node)
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo NodeClickErrorHandler
If Right(Node.Key, 1) <> "\" Then TreeView1.Nodes(Node.Key).SelectedImage = 5 - IsV
Dir1.Path = Node.Key
RaiseEvent NodeClick(Node)
RaiseEvent PathChanged
Exit Sub
NodeClickErrorHandler:
If Err.Number = 68 Then RaiseEvent DriveNotReadyError
RaiseEvent NodeClick(Node)
End Sub

Private Sub UserControl_Initialize()
BuildDriveList
TreeView1.Top = 0
TreeView1.Left = 0
Dir1.Path = "c:\"
End Sub

Private Sub UserControl_Resize()
TreeView1.Width = UserControl.Width ' - 100 remove commet to see half tooltip
TreeView1.Height = UserControl.Height ' - 100
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 TreeView1.Enabled = PropBag.ReadProperty("Enabled", True)
 IsV = PropBag.ReadProperty("V_When_Selected", False)
 Set Font = PropBag.ReadProperty("Font", Ambient.Font)
 TreeView1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
 TreeView1.LabelEdit = PropBag.ReadProperty("LabelEdit", 1)
 TreeView1.HotTracking = PropBag.ReadProperty("HotTracking", False)
 TreeView1.Appearance = PropBag.ReadProperty("Appearance", 1)
 TreeView1.HideSelection = PropBag.ReadProperty("HideSelection", True)
 TreeView1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
 Set MouseIcon = PropBag.ReadProperty("MouseIcon", "")
    TreeView1.Style = PropBag.ReadProperty("Style", 7)
    TreeView1.Indentation = PropBag.ReadProperty("Indentation", 99.77953)
End Sub
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = TreeView1.MousePointer
End Property
Public Property Get LabelEdit() As LabelEditConstants
    LabelEdit = TreeView1.LabelEdit
End Property
Public Property Let LabelEdit(ByVal New_LabelEdit As LabelEditConstants)
    TreeView1.LabelEdit() = New_LabelEdit
    PropertyChanged "LabelEdit"
End Property
Public Property Get HideSelection() As Boolean
    HideSelection = TreeView1.HideSelection
End Property
Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    TreeView1.HideSelection = New_HideSelection
    PropertyChanged "HideSelection"
End Property
Public Property Get HotTracking() As Boolean
    HotTracking = TreeView1.HotTracking
End Property
Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    TreeView1.HotTracking = New_HotTracking
    PropertyChanged "HotTracking"
End Property

Public Property Get BorderStyle() As Integer
 BorderStyle = TreeView1.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    TreeView1.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
MsgBox "This control made by Alon gal. If you got any questions or suggestions E-mail me to cuinl@hotmail.com", , "About this control"
End Sub


Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    TreeView1.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property
Public Property Get Appearance() As AppearanceConstants
    Appearance = TreeView1.Appearance
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    TreeView1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = TreeView1.MouseIcon
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set TreeView1.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Font() As Font
    Set Font = TreeView1.Font
End Property
Public Property Get Enabled() As Boolean
    Enabled = TreeView1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    TreeView1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get V_When_Selected() As Boolean
    V_When_Selected = IsV
End Property

Public Property Let V_When_Selected(ByVal New_V_When_Selected As Boolean)
    IsV = New_V_When_Selected
    PropertyChanged "V_When_Selected"
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set TreeView1.MouseIcon() = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Enabled", TreeView1.Enabled, True)
 Call PropBag.WriteProperty("V_When_Selected", IsV, False)
 Call PropBag.WriteProperty("Font", Font, Ambient.Font)
 Call PropBag.WriteProperty("MouseIcon", TreeView1.MouseIcon, "")
 Call PropBag.WriteProperty("MousePointer", TreeView1.MousePointer, 0)
 Call PropBag.WriteProperty("LabelEdit", TreeView1.LabelEdit, 0)
 Call PropBag.WriteProperty("HideSelection", TreeView1.HideSelection, True)
 Call PropBag.WriteProperty("HotTracking", TreeView1.HotTracking, False)
 Call PropBag.WriteProperty("BorderStyle", TreeView1.BorderStyle, 0)
 Call PropBag.WriteProperty("Appearance", TreeView1.Appearance, 1)
    Call PropBag.WriteProperty("Style", TreeView1.Style, 7)
    Call PropBag.WriteProperty("Indentation", TreeView1.Indentation, 99.77953)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Style
Public Property Get Style() As TreeStyleConstants
Attribute Style.VB_Description = "Displays a hierarchical list of Node objects, each of which consists of a label and an optional bitmap."
    Style = TreeView1.Style
End Property

Public Property Let Style(ByVal New_Style As TreeStyleConstants)
    TreeView1.Style() = New_Style
    PropertyChanged "Style"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TreeView1,TreeView1,-1,Indentation
Public Property Get Indentation() As Single
Attribute Indentation.VB_Description = "Returns/sets the width of the indentation for a TreeView control."
    Indentation = TreeView1.Indentation
End Property

Public Property Let Indentation(ByVal New_Indentation As Single)
    TreeView1.Indentation() = New_Indentation
    PropertyChanged "Indentation"
End Property

