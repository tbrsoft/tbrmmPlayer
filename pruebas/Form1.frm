VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   945
      Left            =   9360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "Form1.frx":0000
      Top             =   60
      Width           =   3585
   End
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   495
      Left            =   3540
      TabIndex        =   8
      Top             =   60
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Left            =   3510
      TabIndex        =   5
      Top             =   630
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   630
      Width           =   3315
   End
   Begin VB.CommandButton Command3 
      Caption         =   "cerrar"
      Height          =   525
      Left            =   7440
      TabIndex        =   3
      Top             =   60
      Width           =   1605
   End
   Begin VB.CommandButton Command2 
      Caption         =   "play"
      Height          =   525
      Left            =   5700
      TabIndex        =   2
      Top             =   60
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "open"
      Height          =   495
      Left            =   3810
      TabIndex        =   1
      Top             =   60
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   7365
      Left            =   180
      ScaleHeight     =   7305
      ScaleWidth      =   10245
      TabIndex        =   0
      Top             =   1050
      Width           =   10305
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   945
         Left            =   240
         TabIndex        =   6
         Top             =   6150
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C0C0&
         Height          =   585
         Left            =   5010
         Shape           =   3  'Circle
         Top             =   2550
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   945
         Left            =   300
         TabIndex        =   7
         Top             =   6210
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   225
      Left            =   5250
      TabIndex        =   11
      Top             =   690
      Width           =   1425
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   555
      Left            =   60
      TabIndex        =   9
      Top             =   30
      Width           =   3435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPath As String

Dim WithEvents MM As tbrPlayer02.MainPlayer
Attribute MM.VB_VarHelpID = -1

Dim RET As Long

Private Sub Command1_Click()
    Dim F As String
    F = sPath + List1
    If LCase(Right(F, 3)) = "mn0" Then
        MM.DoOpenKar F, Picture1, Shape1
    End If
    
    If LCase(Right(F, 3)) = "mpg" Then 'es un video
        MM.FileName(0) = F
        RET = MM.DoOpenVideo("child", Picture1.hWnd, 0, 0, Picture1.Width / 15, Picture1.Height / 15, 0)
        Text1.Text = "Open: " + CStr(RET)
    End If
    
    If LCase(Right(F, 3)) = "mp3" Then 'es un video
        MM.FileName(0) = F
        RET = MM.DoOpen(0)
        Text1.Text = "Open: " + CStr(RET)
        MM.Volumen(0) = 50
    End If
End Sub

Private Sub Command2_Click()
    If LCase(Right(sPath, 3)) = "mn0" Then
        MM.DoPlayKar
    Else
        RET = MM.DoPlay(0)
        Text1.Text = "Play: " + CStr(RET)
    End If
End Sub

Private Sub Command3_Click()
    If LCase(Right(sPath, 3)) = "mn0" Then
        MM.DoStopKar
    Else
        MM.DoStop 0
    End If
    Command4_Click
End Sub

Private Sub Command4_Click()
    If List1.Height = 645 Then
        List1.Height = Picture1.Height
    Else
        List1.Height = 645
    End If
End Sub

Private Sub Command5_Click()
    Dim CM As New CommonDialog
    CM.InitDir = "d:\"
    CM.ShowFolder
    

    sPath = CM.InitDir
    If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
    
    Label3.Caption = sPath
    
    
    Dim A As String
    A = Dir(sPath + "*.*")
    List1.Clear
    Do While A <> ""
        List1.AddItem A
        A = Dir
    Loop
End Sub

Private Sub Form_Load()
    Set MM = New tbrPlayer02.MainPlayer
End Sub

Private Sub Form_Resize()
    Picture1.Left = 0
    Picture1.Width = Me.Width
    Picture1.Height = Me.Height - Picture1.Top - 300
End Sub

Private Sub List1_Click()
    Command4_Click
End Sub

Private Sub MM_FaltaNextEvKAR(dMiliSec As Double)
    Label1 = Format(dMiliSec, "00")
    Label2 = Label1

    Label1.Visible = (dMiliSec > 0)
    Label2.Visible = Label1.Visible
End Sub

Private Sub MM_mmError(txtMasHist As String)
    Text1 = Text1 + txtMasHist + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf
    
End Sub

Private Sub MM_Played(SecondsPlayed As Long, iAlias As Long, MS As Long)
    Label4.Caption = CStr(SecondsPlayed)
    Label4.Refresh
End Sub
