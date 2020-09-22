VERSION 5.00
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1485
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   1800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ActiveScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "ActiveScreen.frx":030A
   PaletteMode     =   2  'Custom
   ScaleHeight     =   1485
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   0
      Left            =   45
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   -30
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   480
      Top             =   15
      Width           =   855
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public W As Long, H As Long, SW As Long, SH As Long
Public ID As Byte, MaxPic As Byte
Public Dx As Integer, Dy As Integer
Public Delay As Byte

Private Sub Form_Activate()
    SetFormTopMost Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHdl
    SW = Screen.Width
    SH = Screen.Height
    Image2.Move 0, 0
    Image2.Stretch = False
    Picture1(0).Move 0, 0
    Picture1(0).AutoSize = True
    Picture1(0).Visible = False
    Picture1(0).ScaleMode = vbPixels
    Image2.Visible = True
    Me.Visible = False
    Me.ScaleMode = vbPixels
    If Forms.Count = 2 Then
        Delay = 4
        Dx = 150: Dy = 150
    Else
        Delay = 1 + Rnd * 5
        Dx = 60 + Rnd * 300: Dy = 60 + Rnd * 300
    End If
    LoadPictureFile Me
    Me.Visible = True
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbCritical, "Error ..."
    Err.Clear
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Shift
        Case 1: LoadPictureFile Me
        Case 2: If Dx = 0 Then Dx = 150: LoadPictureFile Me Else Dx = 0: Dy = 0: Me.Move SW - Me.Width, 200
        Case 4: MsgBox ToolTipTxt
        Case Else: Image2.ToolTipText = Now
    End Select
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE, 0
    Else
        LoadMenu
        LoadDriveMenu
        Set Tao = Me
        PopupMenu FrmOption.mnuMain
    End If
End Sub
