VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
