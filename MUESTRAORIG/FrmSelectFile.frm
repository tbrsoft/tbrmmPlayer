VERSION 5.00
Begin VB.Form FrmSelectFile 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select file"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5805
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   1620
      TabIndex        =   4
      Top             =   4290
      Width           =   975
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   3060
      TabIndex        =   3
      Top             =   4290
      Width           =   975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5565
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   150
      TabIndex        =   1
      Top             =   510
      Width           =   2445
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   2730
      Pattern         =   $"FrmSelectFile.frx":0000
      System          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2955
   End
End
Attribute VB_Name = "FrmSelectFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dim StoreDirive As Variant
    StoreDirive = Dir1.Path
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    If Err = 68 Then Drive1.Drive = StoreDirive
End Sub
Public Function GetFile() As String
Dim File As String
    If File1.filename = "" Then
        MsgBox "Please select a file", vbCritical, "Select file"
        GetFile = "error"
        Exit Function
    End If
    'CHECK IF THE FILE IN ROOT DIR
    If Len(File1.Path) > 3 Then
        File = File1.Path & "\" & File1.filename
    Else
        File = File1.Path & File1.filename
    End If
                
GetFile = File
End Function

Private Sub CmdOk_Click()
    Dim filename As String
    filename = GetFile 'call function to get the file and check if it gave any space
    If filename = "error" Then Exit Sub
    FrmMain.LbFileName(FrmMain.Tag) = filename
    Unload Me
End Sub

Private Sub Form_Load()
    FrmMain.Enabled = False
    Drive1.Drive = "D:\"
    Dir1.Path = "D:\musica\cuartetazo\anastacia\"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.Enabled = True
End Sub
