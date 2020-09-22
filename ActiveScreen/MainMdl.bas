Attribute VB_Name = "MainMdl"
'DDT - FIT - HUT - 3/2K
Option Compare Text
Option Explicit

Type Mn
    MnCation As String
    MnCommand As String
End Type
Public Type Size
        cx As Long
        cy As Long
End Type
Public Type DiskInformation
    lpSectorsPerCluster As Long
    lpBytesPerSector As Long
    lpNumberOfFreeClusters As Long
    lpTotalNumberOfClusters As Long
End Type



Const KEY_ALL_ACCESS = &H2003F
Const ERROR_SUCCESS = 0
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Const REG_DWORD = 4
Const gREGKEYDDT = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Const gREGVALDDT = "ActiveScreen"
Public Const SWP_NOSIDE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SRCCOPY = &HCC0020
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF012
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const EWX_POWEROFF = 8
Public Const EWX_RESET = EWX_LOGOFF + EWX_REBOOT 'EWX_FORCE
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const TRANSPARENT = 1
Public Const TA_LEFT = 0


Public Menus() As Mn
Public SysMenus() As Mn
Public Tao As FrmMain

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Boolean
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Boolean
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Boolean
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Boolean
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Sub SetFormTopMost(F As Form)
    Dim Temp As Long
    Temp = SetWindowPos(F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIDE)
End Sub

Public Sub LoadPictureFile(M As FrmMain)
    On Error GoTo ErrHdl
    Dim FileName As String, Direc As String, ex As String
    Dim Direcs() As String, AppPath As String
    Dim Count As Byte, i As Integer
ReLoad:
    FileName = ""
    Count = 0
    AppPath = App.Path + "\"
    Direc = Dir(AppPath, vbDirectory)
    Do While Direc <> ""
        If Direc <> "." And Direc <> ".." Then
            If (GetAttr(AppPath & Direc) And vbDirectory) = vbDirectory Then
                ReDim Preserve Direcs(Count) As String
                Direcs(Count) = Direc
                Count = Count + 1
            End If
        End If
        Direc = Dir
    Loop
    If Count = 0 Then MsgBox "No picture directory found ?", vbCritical, "Error load picture files !": End
    Direc = AppPath + Direcs((Second(Now) / 60) * (Count - 1)) + "\"
    ChDir Direc
    FileName = Dir("*.*")
    M.MaxPic = 0
    While FileName <> ""
        If M.Picture1.UBound < M.MaxPic Then Load M.Picture1(M.MaxPic)
        'M.Picture1(M.MaxPic).Picture = LoadPicture()
        M.Picture1(M.MaxPic).Picture = LoadPicture(FileName)
        M.MaxPic = M.MaxPic + 1
        FileName = Dir
    Wend
    Reset
    M.ID = 0
    M.W = M.Picture1(0).Width * 15
    M.H = M.Picture1(0).Height * 15
    If M.Dx <> 0 Then M.Move -2 * M.W, M.SH * Rnd, M.W, M.H Else M.Move M.Left, M.Top, M.W, M.H
    If M.MaxPic = 0 Then GoTo ReLoad
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbCritical, "Error load picture files ..."
    Err.Clear
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Public Sub Delay(Tenms As Integer)
    Dim Value As Long
    Value = Timer * 100 + Tenms
    While Timer * 100 < Value
        DoEvents
    Wend
End Sub

Public Function FreeSpace(Drive As String) As String
    On Error GoTo ErrHdl
       Dim info As DiskInformation
       Dim lAnswer As Long
       Dim lpRootPathName As String
       Dim lpSectorsPerCluster As Long
       Dim lpBytesPerSector As Long
       Dim lpNumberOfFreeClusters As Long
       Dim lpTotalNumberOfClusters As Long
       Dim lBytesPerCluster As Long
       Dim lNumFreeBytes As Double
       Dim sString As String

       lpRootPathName = Drive
       lAnswer = GetDiskFreeSpace(lpRootPathName, lpSectorsPerCluster, _
       lpBytesPerSector, lpNumberOfFreeClusters, lpTotalNumberOfClusters)
        If lAnswer Then
            lBytesPerCluster = lpSectorsPerCluster * lpBytesPerSector
            lNumFreeBytes = lBytesPerCluster
            lNumFreeBytes = lNumFreeBytes * lpNumberOfFreeClusters
            sString = Format(((lNumFreeBytes / 1024) / 1024), "0.00") & "MB/"
            lNumFreeBytes = lBytesPerCluster
            lNumFreeBytes = lNumFreeBytes * lpTotalNumberOfClusters
            sString = sString + Format(((lNumFreeBytes / 1024) / 1024), "0.00") & " (" & Format(100 * lpNumberOfFreeClusters / lpTotalNumberOfClusters, "0.00") & "%)"
        End If
       FreeSpace = sString
       Exit Function
ErrHdl:
       MsgBox Err.Description, vbCritical, "FrreSpace..."
       Err.Clear
End Function



Public Function DriveList() As String
    On Error GoTo ErrHdl
    Dim BufLen As Long, Pos As Integer
    Dim BufString As String * 256
    BufLen = 256
    Call GetLogicalDriveStrings(BufLen, BufString)
    Pos = InStr(1, BufString, "\")
    While Pos > 0
        Mid(BufString, Pos + 1, 1) = " "
        Pos = InStr(Pos + 1, BufString, "\")
    Wend
    DriveList = BufString
    Exit Function
ErrHdl:
    MsgBox Err.Description, vbCritical, "FrreSpace..."
    Err.Clear
End Function

Public Sub LoadMenu()
    On Error GoTo ErrHdl
    ChDir App.Path
    Dim Fn As Integer, St As String, MnCount As Byte, MnCount1 As Byte, Pos As Integer, Kind As Byte
    Fn = FreeFile
    MnCount = 0
    MnCount1 = 0
    Kind = 0
    Open "Menu.mnu" For Input As Fn
    Do While Not EOF(Fn)
        St = ""
        Line Input #Fn, St
        Select Case Trim(St)
            Case "[User]"
                Kind = 1
            Case "[System]"
                Kind = 2
            Case Is <> ""
            If Left(St, 1) <> ";" Then
                If Kind = 1 Then
                    ReDim Preserve Menus(MnCount) As Mn
                    Pos = InStr(1, St, ":")
                    Menus(MnCount).MnCation = Trim(Left(St, Pos - 1))
                    Menus(MnCount).MnCommand = Trim(Right(St, Len(St) - Pos))
                    MnCount = MnCount + 1
                End If
                If Kind = 2 Then
                    ReDim Preserve SysMenus(MnCount1) As Mn
                    Pos = InStr(1, St, ":")
                    SysMenus(MnCount1).MnCation = Trim(Left(St, Pos - 1))
                    SysMenus(MnCount1).MnCommand = Trim(Right(St, Len(St) - Pos))
                    MnCount1 = MnCount1 + 1
                End If
            End If
        End Select
    Loop
    Close Fn
    Dim i As Integer
    For i = 0 To MnCount - 1 Step 1
        If FrmOption.mnuUsers.UBound < i Then Load FrmOption.mnuUsers(i)
        FrmOption.mnuUsers(i).Caption = Menus(i).MnCation
        FrmOption.mnuUsers(i).Visible = True
    Next i
    For i = FrmOption.mnuUsers.UBound To MnCount Step -1
        FrmOption.mnuUsers(i).Visible = False
    Next i
    For i = 0 To MnCount1 - 1 Step 1
        If FrmOption.mnuSystems.UBound < i Then Load FrmOption.mnuSystems(i)
        FrmOption.mnuSystems(i).Caption = SysMenus(i).MnCation
        FrmOption.mnuSystems(i).Visible = True
    Next i
    For i = FrmOption.mnuSystems.UBound To MnCount1 Step -1
        FrmOption.mnuSystems(i).Visible = False
    Next i
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbCritical, "Load menu : Menu.Mnu"
    Err.Clear
End Sub

Public Sub LoadDriveMenu()
    Dim i As Integer, Drives As String, Pos As Integer
    Drives = DriveList
    Pos = InStr(1, Drives, "\")
    i = 2
    While Pos > 0
        If FrmOption.mnuDrives.UBound < i Then Load FrmOption.mnuDrives(i)
        FrmOption.mnuDrives(i).Caption = UCase(Mid(Drives, Pos - 2, 3)) ' & "  Free: " & FreeSpace(Mid(Drives, Pos - 2, 3))
        FrmOption.mnuDrives(i).Visible = True
        Pos = InStr(Pos + 1, Drives, "\")
        i = i + 1
    Wend
End Sub

Public Function ToolTipTxt() As String
    On Error Resume Next
    Dim Drives As String, Pos As Integer, St As String, Free As String
    Drives = DriveList
    Pos = InStr(1, Drives, "\")
     St = ""
    While Pos > 0
        Free = Trim(FreeSpace(Mid(Drives, Pos - 2, 3)))
        If Free = "" Then Free = "  Not ready    " Else Free = "  Free:" + Free + "    "
        St = St + UCase(Mid(Drives, Pos - 2, 3)) & Free & vbCrLf
        Pos = InStr(Pos + 1, Drives, "\")
    Wend
    ToolTipTxt = St & vbCrLf & vbCrLf & Now
End Function
