VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Preview"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPreview 
      AutoSize        =   -1  'True
      Height          =   2895
      Left            =   0
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
