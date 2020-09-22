VERSION 5.00
Begin VB.Form frmReorder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-order Wallpaper List"
   ClientHeight    =   3240
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7515
   Icon            =   "frmReorder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7515
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Picture         =   "frmReorder.frx":2892
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Move selection up on the list by 1"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Picture         =   "frmReorder.frx":2D65
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Move selection down on the list by 1"
      Top             =   1800
      Width           =   855
   End
   Begin VB.ListBox lsbfiles 
      Height          =   2985
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmReorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tmpFiles As Collection
Dim Reordered As Boolean

Private Sub CancelButton_Click()
    Unload frmReorder
End Sub

Private Sub cmdMoveDown_Click()
    Dim Item() As String
    Dim i As Long
    
    If lsbfiles.ListCount = 0 Then Exit Sub
    If lsbfiles.Selected(lsbfiles.ListCount - 1) = True Then Exit Sub
    
    ReDim Item(0 To lsbfiles.ListCount - 1)
    
    For i = lsbfiles.ListCount - 1 To 0 Step -1
        If lsbfiles.Selected(i) = True Then
            Item(i) = tmpFiles(i + 1)
            lsbfiles.RemoveItem (i)
            tmpFiles.Remove (i + 1)
        End If
    Next
    
    For i = 0 To UBound(Item)
        If Item(i) <> vbNullString Then
            lsbfiles.AddItem FileNameFromPath(Item(i)), i + 1
            tmpFiles.Add Item(i), , , i + 1
            lsbfiles.Selected(i + 1) = True
        End If
    Next
End Sub

Private Sub cmdMoveUp_Click()
    Dim Item() As String
    Dim i As Long
    
    If lsbfiles.ListCount = 0 Then Exit Sub
    If lsbfiles.Selected(0) = True Then Exit Sub
        
    ReDim Item(0 To lsbfiles.ListCount - 1)
    
    For i = lsbfiles.ListCount - 1 To 0 Step -1
        If lsbfiles.Selected(i) = True Then
            Item(i) = tmpFiles(i + 1)
            lsbfiles.RemoveItem (i)
            tmpFiles.Remove (i + 1)
        End If
    Next
    
    For i = 0 To UBound(Item)
        If Item(i) <> vbNullString Then
            lsbfiles.AddItem FileNameFromPath(Item(i)), i - 1
            tmpFiles.Add Item(i), , i
            lsbfiles.Selected(i - 1) = True
        End If
    Next
End Sub

Private Sub Form_Load()
    Dim i As Long
    lsbfiles.Clear
    Reordered = False
    
    For i = 1 To tmpFiles.Count
        frmReorder.lsbfiles.AddItem FileNameFromPath(tmpFiles(i))
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tmpFiles = Nothing
    
    If Reordered Then Call RestartApp
End Sub

Private Sub OKButton_Click()
    On Error Resume Next
    Dim Bin As String
    Dim i As Long
    
    Kill AppPath & "Files.ini"
    
    For i = 1 To tmpFiles.Count
        Bin = Bin & tmpFiles(i)
        If i < tmpFiles.Count Then
            Bin = Bin & ";" & vbNewLine
            Else
            Bin = Bin & ";"
        End If
    Next
    
    Open AppPath & "Files.ini" For Binary As 1
        Put 1, , Bin
        Reordered = True
    Close
    
    Unload frmReorder
End Sub
