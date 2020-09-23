VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Browse Files On Your Computer"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
Shell Dir1.Path & File1.FileName
End Sub
