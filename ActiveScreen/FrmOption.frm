VERSION 5.00
Begin VB.Form FrmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   1755
   ClientLeft      =   2760
   ClientTop       =   4035
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1395
      Top             =   420
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuUser 
         Caption         =   "&User"
         Begin VB.Menu mnuUsers 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuSystem 
         Caption         =   "&System"
         Begin VB.Menu mnuSystems 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "&Desk top"
         Begin VB.Menu mnuDesktops 
            Caption         =   "Hide All Windows"
            Index           =   0
         End
         Begin VB.Menu mnuDesktops 
            Caption         =   "Undo Hide All"
            Index           =   1
         End
         Begin VB.Menu mnuDesktops 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuDesktops 
            Caption         =   "Minimize All Windows"
            Index           =   3
         End
         Begin VB.Menu mnuDesktops 
            Caption         =   "Undo Minimize All"
            Index           =   4
         End
         Begin VB.Menu mnuDesktops 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuDesktops 
            Caption         =   "New Shortcut"
            Index           =   6
         End
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "D&rive"
         Begin VB.Menu mnuDrives 
            Caption         =   "Free [Alt+Mosemove]"
            Index           =   0
         End
         Begin VB.Menu mnuDrives 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuDrives 
            Caption         =   ""
            Index           =   2
         End
      End
      Begin VB.Menu S3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "Sh&ut Down"
         Begin VB.Menu mnuExitWindows 
            Caption         =   "Shut down"
            Index           =   0
         End
         Begin VB.Menu mnuExitWindows 
            Caption         =   "Restart"
            Index           =   1
         End
         Begin VB.Menu mnuExitWindows 
            Caption         =   "Restart in MS-DOS mode"
            Index           =   2
         End
         Begin VB.Menu mnuExitWindows 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuExitWindows 
            Caption         =   "Terminate All"
            Index           =   4
         End
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "&Option"
         Begin VB.Menu mnuSpeed 
            Caption         =   "Speed"
         End
         Begin VB.Menu mnuStep 
            Caption         =   "Step"
         End
         Begin VB.Menu mnuCount 
            Caption         =   "Count"
         End
         Begin VB.Menu mnuOptionS0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStep0 
            Caption         =   "Step = 0 [Ctrl+MouseMove]"
         End
         Begin VB.Menu mnuOptionS1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEdit 
            Caption         =   "Menu edit"
         End
      End
      Begin VB.Menu S0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Windows() As Long
Private Times As Integer

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2

Private Sub Form_Load()
    Check
    ReDim Windows(0) As Long
    Timer1.Enabled = True
    FrmMain.Show
End Sub

Private Sub mnuAbout_Click()
    Dim hwnd As Long, hdc As Long, Temp As Long, Txt As String, L As Integer
    Dim i As Integer, Char As String
    Txt = "§Õn th­îng ®Õ còng thÊy rÔ dïng !!!" & vbCrLf & _
    "ChØ cÇn mét có click lµ b¹n sÏ cã mét thùc ®¬n ®Çy c¸c mãn ¨n mµ Windows cung cÊp cho b¹n !" & vbCrLf & _
    "H·y nÕm c¸c mãn ¨n nµy nÕu b¹n muèn trë thµnh mét ng­êi sµnh ®iÖu..." & vbCrLf & _
    "Råi b¹n sÏ thÊy Windows kú thó ®Õn møc nµo ?" & vbCrLf & _
    "H·y t×m tßi nghiªn cøu, file Menu.Mnu còng cho b¹n thÊy mét sè c¸ch gÆm nhÊm c¸c hµm trong Windows" & vbCrLf & _
    "M©m c¬m nµy ®­îc ®¨ng ký b¶n quyÒn ë 'Kh«ng ®©u c¶' vµ lµ b¶n quyÒn cña 'Kh«ng ai c¶'" & vbCrLf & _
    "Nã ®­îc b¸n víi gi¸ 50 VN§, nh­ng lo¹i tiÒn nµy hiÖn nay rÊt khã kiÕm nªn nã ®­îc cho kh«ng !" & vbCrLf & _
    "" & vbCrLf & _
    "DDT - FIT - HUT - 3/2K"
    L = Len(Txt)
    hwnd = GetDesktopWindow: hdc = GetWindowDC(hwnd)
    Temp = SetBkMode(hdc, TRANSPARENT)
    Temp = SetTextAlign(hdc, TA_LEFT)
    Dim X As Long, Y As Long, TextWidthHeight As Size
    X = 50
    Y = 50
    For i = 1 To L Step 1
        Char = Mid(Txt, i, 1)
        If InStr(1, vbCrLf, Char) Then
            Temp = GetTextExtentPoint32(hdc, "H", 1, TextWidthHeight)
            X = 50
            Y = Y + TextWidthHeight.cy
            Delay (10)
        Else
            Temp = SetTextColor(hdc, Rnd * &H1000000)
            Temp = GetTextExtentPoint32(hdc, Char, 1, TextWidthHeight)
            Temp = TextOut(hdc, X, Y, Char, 1)
            X = X + TextWidthHeight.cx
            Delay (2)
        End If
    Next i
    Temp = ReleaseDC(hwnd, hdc)
    Temp = DeleteDC(hdc)
End Sub

Private Sub mnuCount_Click()
    Dim n As Byte, i As Integer
    Dim F As Form
    n = 0
    While (n <= 0) Or (n > 10)
        n = Val(InputBox("Number of ico: (1 to 10)", , Forms.Count - 1))
        DoEvents
    Wend
    For i = Forms.Count To n Step 1
        Set F = New FrmMain
        Load F
        F.Show
        DoEvents
    Next i
    For i = Forms.Count - 1 To n + 1 Step -1
        Unload Forms(i)
        DoEvents
    Next i
End Sub

Private Sub mnuDesktops_Click(Index As Integer)
    On Error GoTo ErrHdl
    Dim Temp As Long, Reserved As Long
    Select Case Index
        Case 0, 3, 4
            Dim CurrWnd As Long, OldW As Long, Count As Long
            Dim Length As Long, TaskName As String
            CurrWnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
            OldW = CurrWnd
            Count = 0
            If Index = 0 Then mnuDesktops_Click (1)
            While OldW <> 0
                DoEvents
                Length = GetWindowTextLength(OldW)
                TaskName = Space$(Length + 1)
                Length = GetWindowText(OldW, TaskName, Length + 1)
                TaskName = Trim(Left$(TaskName, Len(TaskName) - 1))
                If (Trim(TaskName) <> Trim(App.Title)) And (Trim(TaskName) <> "") And (IsWindow(OldW)) Then
                    If (IsWindowVisible(OldW)) Then
                        Select Case Index
                            Case 0: Temp = SW_HIDE
                            Case 3: Temp = SW_SHOWMINNOACTIVE
                            Case 4: Temp = SW_RESTORE
                        End Select
                        Temp = ShowWindow(OldW, Temp)
                        Count = Count + 1
                        ReDim Preserve Windows(Count) As Long
                        Windows(Count) = OldW
                    End If
                End If
                OldW = CurrWnd
                CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
            Wend
        Case 1
            For Count = UBound(Windows) To 1 Step -1
                If (IsWindow(Windows(Count))) And (Not IsWindowVisible(Windows(Count))) Then
                    Temp = ShowWindow(Windows(Count), SW_SHOW)
                End If
            Next Count
        Case 6
            Dim FileName As String, Fn As Integer
            Fn = FreeFile
            FileName = Space(256)
            Temp = GetWindowsDirectory(FileName, Len(FileName))
            FileName = Trim(FileName)
            Mid(FileName, Len(FileName), 1) = "\"
            FileName = FileName + "Desktop\Tao ë ®©y.lnk"
            Open FileName For Binary Access Write As Fn
            Close Fn
            Shell "Rundll32.exe AppWiz.Cpl,NewLinkHere " + FileName
    End Select
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub mnuDrives_Click(Index As Integer)
    On Error GoTo ErrHdl
    If Index = 0 Then MsgBox ToolTipTxt: Exit Sub
    Shell "Explorer.exe " + mnuDrives(Index).Caption, vbNormalFocus
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub mnuEdit_Click()
    On Error Resume Next
    ChDir App.Path
    Shell "notepad.exe Menu.mnu", vbNormalFocus
    If Err Then MsgBox Err.Description, vbCritical, "Running: " + "Notepad.exe Menu.txt": Err.Clear
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuExitWindows_Click(Index As Integer)
    On Error GoTo ErrHdl
    Dim Temp As Long, Reserved As Long
    Select Case Index
        Case 0
            Temp = ExitWindowsEx(EWX_SHUTDOWN, Reserved)
        Case 1
            Temp = ExitWindowsEx(EWX_RESET, Reserved)
        Case 2
            Shell "Exit To Dos.pif", vbNormalFocus
        Case 4
            If MsgBox("Do you wish to terminate all programs now and lose any unsaved information in the programs?", vbExclamation + vbYesNo) = vbYes Then
                Temp = ExitWindowsEx(EWX_FORCE, Reserved)
            End If
    End Select
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub mnuSpeed_Click()
    Dim n As Integer
    n = 0
    While (n <= 0) Or (n > 10)
        n = Val(InputBox("Time delay: (1 to 10)", , Tao.Delay))
    Wend
    Tao.Delay = n
End Sub

Private Sub mnuStep_Click()
    Dim n As Integer
    n = -1
    While (n < 0) Or (n > 100)
        n = Val(InputBox("Step: (0 to 100)", , Tao.Dx \ 15))
    Wend
    Tao.Dx = n * 15
End Sub

Private Sub mnuStep0_Click()
    Tao.Dx = 0
    Tao.Dy = 0
End Sub

Private Sub mnuSystems_Click(Index As Integer)
    On Error Resume Next
    ChDir App.Path
    Shell SysMenus(Index).MnCommand, vbNormalFocus
    If Err Then MsgBox Err.Description, vbCritical, "Running: " + SysMenus(Index).MnCommand:: Err.Clear
End Sub

Private Sub mnuUsers_Click(Index As Integer)
    On Error Resume Next
    ChDir App.Path
    Shell Menus(Index).MnCommand, vbNormalFocus
    If Err Then MsgBox Err.Description, vbCritical, "Running: " + Menus(Index).MnCommand:: Err.Clear
End Sub

Sub Check()
    If App.PrevInstance Then MsgBox "Active screen already nunning !", vbInformation: End
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    DoEvents
    Dim i As Integer, hwnd As Long, hdc As Long
    Dim L As Long, T As Long, Temp As Long
    Dim M As FrmMain
    Times = (Times + 1) Mod 30000
    For i = 1 To Forms.Count - 1 Step 1
        Set M = Forms(i)
        If (Times Mod M.Delay) = 0 Then
            L = M.Left
            T = M.Top
            M.Width = 0
            DoEvents
            hwnd = GetDesktopWindow: hdc = GetWindowDC(hwnd)
            Temp = BitBlt(M.hdc, 0, 0, M.W / 15, M.H / 15, hdc, (L + M.Dx) / 15, (T + M.Dy) / 15, SRCCOPY)
            Temp = ReleaseDC(hwnd, hdc)
            Temp = DeleteDC(hdc)
            M.Width = M.W
            L = L + M.Dx
            T = T + M.Dy
            If L > M.SW Then L = -M.W: LoadPictureFile M
            If T > M.SH Then T = -M.H: LoadPictureFile M
            If T < -M.H Then T = M.SH: LoadPictureFile M
            M.Move L, T
            M.ID = (M.ID + 1) Mod M.MaxPic
            'M.Image2.Picture = LoadPicture()
            M.Image2.Picture = M.Picture1(M.ID).Picture
            
            If (Times Mod 1000 = 0) Then LoadPictureFile M: SetFormTopMost M
            If Rnd < 0.05 Then M.Dy = 2 * Rnd * M.Dx - M.Dx
        End If
    Next i
End Sub
