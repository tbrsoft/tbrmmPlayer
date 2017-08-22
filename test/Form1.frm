VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox VU2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   2250
      ScaleHeight     =   1005
      ScaleWidth      =   600
      TabIndex        =   3
      Top             =   30
      Width           =   600
   End
   Begin VB.PictureBox VU1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   600
      TabIndex        =   2
      Top             =   30
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   660
      TabIndex        =   0
      Top             =   540
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1110
      TabIndex        =   1
      Top             =   270
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents TM As tbrPlayer.MainPlayer
Attribute TM.VB_VarHelpID = -1
Dim VU As tbrSoftVumetro.tbrDrawVUM
Attribute VU.VB_VarHelpID = -1

Private Sub Command1_Click()
    TM.FileName(0) = "D:\musica\Cuartetazo\Jaf - Aire\01 - Jaf - Agua Lenta.Mp3"
    TM.DoOpen 0
    
    TM.DoPlay 0
End Sub

Private Sub Form_Load()
    Set TM = New tbrPlayer.MainPlayer
    
    Set VU = New tbrSoftVumetro.tbrDrawVUM
    
    VU.Enabled = True
    VU.DefinePictureBox VU1
    VU.DefinePictureBox2 VU2
    VU.CantCuadros = 30
    VU.ModoVumetro = TresColoresEstereo
    VU.Enabled = True
    VU.FramePorSeg = 15
    VU.ColorBase = vbRed
    VU.Empezar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TM.DoClose 0
    VU.DoPause False
    VU.Terminar
    Set TM = Nothing
    Set VU = Nothing
End Sub

Private Sub TM_Played(SecondsPlayed As Long, iAlias As Long)
    Label1.Caption = SecondsPlayed
    Label1.Refresh
End Sub
