VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Utils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MyComputer As WinInfo
Public MyLocale   As Locale

Public Enum EShellShowConstants
     essSW_HIDE = 0
     essSW_MAXIMIZE = 3
     essSW_MINIMIZE = 6
     essSW_SHOWMAXIMIZED = 3
     essSW_SHOWMINIMIZED = 2
     essSW_SHOWNORMAL = 1
     essSW_SHOWNOACTIVATE = 4
     essSW_SHOWNA = 8
     essSW_SHOWMINNOACTIVE = 7
     essSW_SHOWDEFAULT = 10
     essSW_RESTORE = 9
     essSW_SHOW = 5
End Enum

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                     '  file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                     '  path not found
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_SHARE = 26


Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260

Private Const INTERNET_CONNECTION_MODEM = 1
Private Const INTERNET_CONNECTION_LAN = 2
Private Const INTERNET_CONNECTION_PROXY = 4
Private Const INTERNET_CONNECTION_MODEM_BUSY = 8

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Const FLASHW_TRAY = 2

Private Const WM_USER    As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)

Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7

Public Enum FilePartTypes
    FileExtOnly
    FileNameOnly
    FileNameAndExt
    FilePathOnly
End Enum

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Public Enum GoodDateTypes
    AnyDate         'Allows any valid date to be entered
    FutureDate      'Only allows future dates (after today) to be entered
    FutureTime      'Only allows today or future date to be entered
    PastDate        'Only allows past dates (before today) to be entered
    PastTime        'Only allows today or a previous day to be entered
End Enum

Private Type Rectangle
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type

'Flash Window
Private Type FLASHWINFO
  cbSize    As Long
  hWnd      As Long
  dwFlags   As Long
  uCount    As Long
  dwTimeout As Long
End Type


Private Type GUID
    Data1    As Long
    Data2    As Long
    Data3    As Long
    Data4(8) As Byte
End Type

Public Enum MonitorState
    MonitorOn = -1
    MonitorOff = 2
    MonitorStandby = 1
End Enum
                        
'API Subs
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub Sleep Lib "kernel32" (ByVal MilliSeconds As Long)
'API Functions
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFilename As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, ByVal lpInputName As String, ByVal lpOutputName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Function AsPercent(MaxVal As Long, ValNow As Long, Optional UpToMax As Boolean) As Integer
    AsPercent = Round(ValNow / MaxVal * 100, 0)
    If AsPercent > 100 And UpToMax = True Then AsPercent = 100
End Function

Public Function DLLBaseAddress() As String
    Randomize       'Initialise Rnd (using system time as seed)
    DLLBaseAddress = "&H" & LCase(Hex(CLng((32768 - 256 + 1) * Rnd + 256) * 65536))
End Function

Public Function Extract(ByVal TextIN As String, Optional StartTag As String = " ", Optional ByVal EndTag As String = " ") As String

    On Error GoTo LocalError
    ' Extracts Text from string using start and end "tags"
    'NB:    If EndTag is ommitted the entire string from:
    '       StartTag to EndOfString is returned...
    Dim lArray As Variant
    Extract = ""
    lArray = Split(TextIN, StartTag)
    If IsArray(lArray) Then
        Extract = lArray(1)
        lArray = Split(Extract, EndTag)
        If IsArray(lArray) Then
            Extract = lArray(0)
        Else
            Extract = ""
        End If
    End If
Exit Function

LocalError:
    Extract = ""
End Function

'File Operations
Public Function FileAttributes(ByVal Filename$) As String

    Dim lngFileAttributes As Long
    Dim strFileAttributeFlags As String

    On Error Resume Next

    If Not FileExists(Filename) Then Exit Function

    lngFileAttributes = GetFileAttributes(Filename)
    If lngFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then FileAttributes = FileAttributes + "D"
    If lngFileAttributes And FILE_ATTRIBUTE_ARCHIVE Then FileAttributes = FileAttributes + "A"
    If lngFileAttributes And FILE_ATTRIBUTE_SYSTEM Then FileAttributes = FileAttributes + "S"
    If lngFileAttributes And FILE_ATTRIBUTE_HIDDEN Then FileAttributes = FileAttributes + "H"
    If lngFileAttributes And FILE_ATTRIBUTE_READONLY Then FileAttributes = FileAttributes + "R"

End Function

Public Function FileCopy(SourceFile$, TargetFile$, Optional ErrMsg$ = "") As Boolean

    Dim FSO As Variant
    Dim Src As Variant
    Dim TRG As Variant

    On Error GoTo LocalError

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FileExists(SourceFile) Then
        If FileExists(TargetFile) Then
            Kill TargetFile
        End If
        Set Src = FSO.GetFile(SourceFile)
        Src.Copy TargetFile
        If FileExists(TargetFile) Then FileCopy = True
    End If
Exit Function

LocalError:
    ErrMsg = Err.Number & " - " & Err.Description
    FileCopy = False
End Function

Public Function FileCopyAPI(Src$, Dest$, Optional FailIfDestExists As Boolean) As Boolean

    'PARAMETERS:    src:    Source File (FullPath)
    '               dest:   Destination File (FullPath)
    '               FailIfDestExists (Optional):
    Dim lRet As Long
    lRet = CopyFile(Src, Dest, FailIfDestExists)
    FileCopyAPI = (lRet > 0)

End Function

Public Function FileCreated(ByVal Filename$) As Date

    Dim lngFileHandle As Long
    Dim udtSystemTime As SYSTEMTIME
    Dim udtWinFindData As WIN32_FIND_DATA

    On Error Resume Next

    If Not FileExists(Filename) Then Exit Function
    lngFileHandle = FindFirstFile(Filename, udtWinFindData)
    Call FileTimeToSystemTime(udtWinFindData.ftCreationTime, udtSystemTime)
    FileCreated = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour + FileAdjustTime, udtSystemTime.wMinute, udtSystemTime.wSecond)
    Call FindClose(lngFileHandle)

End Function

Public Sub FileDirectories(ByVal StartPath$, ByVal colFilePaths As Collection)
    
    Dim Filename As String
    Dim lngReturnCode As Long
    Dim udtWinFindData As WIN32_FIND_DATA
    Dim blnFoundMatch As Boolean
    Dim strPath As String

    StartPath = PathCheck(StartPath)
    lngReturnCode = FindFirstFile(StartPath + "*", udtWinFindData)
    blnFoundMatch = (lngReturnCode > 0)

    Do While blnFoundMatch
        Filename = udtWinFindData.cFileName
        If InStr(Filename, Chr$(0)) > 0 Then
            Filename = Left$(Filename, InStr(Filename, Chr$(0)) - 1)
        End If
        If Filename <> "." And Filename <> ".." Then
            If udtWinFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                strPath = StartPath + Filename
                If Right$(strPath, 1) <> "\" Then
                    strPath = strPath + "\"
                End If
                colFilePaths.Add strPath, strPath
                FileDirectories strPath, colFilePaths
            End If
        End If
        blnFoundMatch = FindNextFile(lngReturnCode, udtWinFindData)
    Loop

    FindClose lngReturnCode

End Sub

Public Function FileExists(ByVal Filename$) As Boolean

    Dim lngFileHandle As Long
    Dim udtWinFindData As WIN32_FIND_DATA

    On Error Resume Next

    If ((Len(Filename) > 3) And (Right$(Filename, 1) = "\")) Then
        Filename = Left$(Filename, Len(Filename) - 1)
    End If
    lngFileHandle = FindFirstFile(Filename, udtWinFindData)
    FileExists = lngFileHandle <> INVALID_HANDLE_VALUE
    Call FindClose(lngFileHandle)

End Function

Public Function FileFind(RootPath$, Filename$) As String
    
    Dim lNullPos As Long
    Dim lResult As Long
    Dim sBuffer As String

    Const MAX_PATH = 260

    On Error GoTo LocalError

    'Allocate buffer
    sBuffer = Space(MAX_PATH * 2)
    'Find the file
    lResult = SearchTreeForFile(RootPath, Filename, sBuffer)

    'Trim null, if exists
    If lResult Then
        lNullPos = InStr(sBuffer, vbNullChar)
        If Not lNullPos Then
            sBuffer = Left$(sBuffer, lNullPos - 1)
        End If
        'Return filename
        FileFind = sBuffer
    Else
        'Nothing found
        FileFind = vbNullString
    End If
Exit Function

LocalError:
    FileFind = vbNullString
End Function

Public Function FileKill(FileMask$, Optional OlderThan As Variant, Optional Prompt As Boolean) As Boolean

    On Error GoTo LocalError

    If Not IsMissing(OlderThan) Then
        If IsDate(OlderThan) Then
            Dim NextFile As String
            OlderThan = CDate(OlderThan)
            NextFile = Dir(FileMask)
            If Prompt And Len(NextFile) > 0 Then
                Dim lResponse As VbMsgBoxResult
                lResponse = MsgBox("Delete file(s) " & FilePart(FileMask, FileNameAndExt), vbYesNo + vbExclamation)
                If lResponse = vbNo Then Exit Function
            End If
            While Len(NextFile) > 0
                If FileLastAccessed(FileMask) < OlderThan Then
                    Kill NextFile   'Delete the file
                End If
                NextFile = Dir
            Wend
        End If
    Else    'Just do it
        Kill FileMask
    End If
    FileKill = True
Exit Function

LocalError:
    If Err.Number = 53 Then
        'file(s) was not found - continue
        FileKill = True
    Else
        FileKill = False
    End If
End Function

Public Function FileLastAccessed(ByVal Filename$) As Date

    Dim datFileCreationDate As Date
    Dim lngFileHandle As Long
    Dim udtSystemTime As SYSTEMTIME
    Dim udtWinFindData As WIN32_FIND_DATA

    On Error Resume Next

    If Not FileExists(Filename) Then Exit Function

    lngFileHandle = FindFirstFile(Filename, udtWinFindData)
    Call FileTimeToSystemTime(udtWinFindData.ftLastAccessTime, udtSystemTime)
    datFileCreationDate = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour + FileAdjustTime, udtSystemTime.wMinute, udtSystemTime.wSecond)
    FileLastAccessed = datFileCreationDate
    Call FindClose(lngFileHandle)

End Function

Public Function FileLastModified(ByVal Filename$) As Date

    Dim datFileCreationDate As Date
    Dim lngFileHandle As Long
    Dim udtSystemTime As SYSTEMTIME
    Dim udtWinFindData As WIN32_FIND_DATA

    On Error Resume Next

    If Not FileExists(Filename) Then Exit Function
    lngFileHandle = FindFirstFile(Filename, udtWinFindData)
    Call FileTimeToSystemTime(udtWinFindData.ftLastWriteTime, udtSystemTime)
    datFileCreationDate = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour + FileAdjustTime, udtSystemTime.wMinute, udtSystemTime.wSecond)
    FileLastModified = datFileCreationDate
    Call FindClose(lngFileHandle)

End Function

Public Function FilePart(FullPath As String, Optional WhichPart As FilePartTypes = FileNameOnly) As String

    If Len(FullPath) = 0 Then Exit Function

    Dim lArray As Variant
    Dim lSeperator As String

    lSeperator = "\"
    If InStr(FullPath, "/") > 0 Then lSeperator = "/"

    Select Case WhichPart
        Case FileExtOnly
            If InStr(FullPath, ".") Then
                lArray = Split(FullPath, ".")
                FilePart = lArray(UBound(lArray))
            End If
        Case FileNameOnly, FileNameAndExt
            lArray = Split(FullPath, lSeperator)
            FilePart = lArray(UBound(lArray))
            If WhichPart = FileNameOnly Then
                lArray = Split(FilePart, ".")
                FilePart = lArray(LBound(lArray))
            End If
        Case FilePathOnly
            Dim lFileName As String
            lFileName = FilePart(FullPath, FileNameAndExt)
            FilePart = Replace(FullPath, lFileName, "")
    End Select

End Function

Public Function FileRead(ByVal Filename$) As String

    Dim lngFileHandle As Long

    On Error Resume Next

    If FileExists(Filename) Then
        If Not InStr(FileAttributes(Filename), "D") Then
            lngFileHandle = FreeFile
            Open Filename For Binary As #lngFileHandle
            FileRead = Space(FileLen(Filename))
            Get #lngFileHandle, , FileRead
            Close #lngFileHandle
        End If
    End If

End Function

Public Function FileShortPath(ByVal Filename$) As String

    Dim strBuffer As String * 255
    Dim lngReturnCode As Long

    lngReturnCode = GetShortPathName(Filename, strBuffer, 255)
    FileShortPath = Left$(strBuffer, lngReturnCode)

End Function

Public Function FileSize(ByVal Filename As String) As Long

    'Get the file Size
    FileSize = FileLen(Filename) \ 1024

End Function

Public Function FileWrite(ByVal Filename$, ByVal FileContents$) As Boolean

    Dim lngFileHandle As Long

    On Error Resume Next

    If FileExists(Filename) Then
        If InStr(FileAttributes(Filename), "D") Then
            Exit Function
        Else
            Kill Filename
        End If
    End If

    lngFileHandle = FreeFile
    Open Filename For Binary As #lngFileHandle
    Put #lngFileHandle, , FileContents
    Close #lngFileHandle
    FileWrite = True

End Function

Private Function FileAdjustTime() As Long

    Dim datSystemDate As Date
    Dim udtSystemTime As SYSTEMTIME

    On Error Resume Next

    Call GetSystemTime(udtSystemTime)
    datSystemDate = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour, udtSystemTime.wMinute, udtSystemTime.wSecond)
    FileAdjustTime = DateDiff("h", datSystemDate, Now)

End Function
'End of File Operations

Public Sub FormOnTop(hWnd As Long, OnTop As Boolean)

    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos hWnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, TOPMOST_FLAGS

End Sub

Public Function Gender(GenderX As Variant) As Variant

    If IsNumeric(GenderX) Then
        'Output Conversion required
        Select Case GenderX
            Case 1:     Gender = "Female"
            Case 2:     Gender = "Male"
            Case Else:  Gender = "Unknown"
        End Select
    Else
        'Input conversion required
        Select Case UCase(Left$(GenderX, 1))
            Case "F":  Gender = 1
            Case "M":  Gender = 2
            Case Else: Gender = 0
        End Select
    End If

End Function

Public Function GUIDx(Optional Formatted As Boolean = False) As String

    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim Retval As Long

    lLen = 40
    bGUID = String(lLen, 0)
    CoCreateGuid uGUID

    'Convert the structure into a displayable string
    Retval = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    sGUID = bGUID

    If (Asc(Mid$(sGUID, Retval, 1)) = 0) Then Retval = Retval - 1
    GUIDx = Left$(sGUID, Retval)

    If Not Formatted Then
        GUIDx = Replace(GUIDx, "{", "")
        GUIDx = Replace(GUIDx, "}", "")
        GUIDx = Replace(GUIDx, "-", "")
    End If

End Function

Public Function OLEColor(ByVal Color As OLE_COLOR, Optional Palette As Long = 0) As OLE_COLOR

    Const CLR_INVALID As Long = &HFFFFFFFF
    OLEColor = IIf(OleTranslateColor(Color, Palette, OLEColor) <> 0, CLR_INVALID, OLEColor)

End Function

Public Function PI() As Double

    'returns the value of PI to 14 decimal places
    PI = (4 * Atn(1))

End Function

Public Sub PlaySound(FromSoundFile$, Optional Continuous As Boolean = False)

    Dim lSoundFile As String
    Dim lSysMedia   As String
    Dim lAppMedia   As String

    Const SND_ASYNC = &H1
    Const SND_LOOP = &H8
    Const SND_NODEFAULT = &H2
    Const SND_SYNC = &H0
    Const SND_NOSTOP = &H10
    Const SND_MEMORY = &H4

    lSoundFile = FromSoundFile
    
    If FileExists(FromSoundFile) Then
        lSoundFile = FromSoundFile
    ElseIf FileExists(PathCheck(App.Path) & FilePart(FromSoundFile, FileNameAndExt)) Then
    End If

    If FileExists(FromSoundFile) Or FileExists(MyComputer.Directory(dirWINDOWS) & "\Media\" & FilePart(FromSoundFile, FileNameAndExt)) Then
        sndPlaySound FromSoundFile, IIf(Continuous, SND_ASYNC Or SND_LOOP, SND_ASYNC)
    End If

End Sub

Public Function RandomNumber(Lower As Integer, Upper As Integer) As Integer

    On Error GoTo LocalError

    'Generates a Random Number BETWEEN then LOWER and UPPER values
    Randomize
    RandomNumber = Int((Upper - Lower + 1) * Rnd + Lower)
Exit Function

LocalError:
    RandomNumber = 1
End Function

Public Function ShellEx(ByVal sFile As String, _
               Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
               Optional ByVal sParameters As String = "", _
               Optional ByVal sDefaultDir As String = "", _
               Optional sOperation As String = "open", _
               Optional Owner As Long = 0) As Boolean

    Dim lR As Long
    Dim lErr As Long, sErr As Long

    On Error Resume Next

    If (InStr(UCase$(sFile), ".EXE") <> 0) Then eShowCmd = 0

    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
    End If

    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
            Case 0: lErr = 7: sErr = "Out of memory"
            Case ERROR_FILE_NOT_FOUND:   lErr = 53: sErr = "File not found"
            Case ERROR_PATH_NOT_FOUND:   lErr = 76: sErr = "Path not found"
            Case ERROR_BAD_FORMAT:       sErr = "The executable file is invalid or corrupt"
            Case SE_ERR_ACCESSDENIED:    lErr = 75: sErr = "Path/file access error"
            Case SE_ERR_ASSOCINCOMPLETE: sErr = "This file type does not have a valid file association."
            Case SE_ERR_DDEBUSY:         lErr = 285: sErr = "The file could not be opened because the target application is busy.  Please try again in a moment."
            Case SE_ERR_DDEFAIL:         lErr = 285: sErr = "The file could not be opened because the DDE transaction failed.  Please try again in a moment."
            Case SE_ERR_DDETIMEOUT:      lErr = 286: sErr = "The file could not be opened due to time out.  Please try again in a moment."
            Case SE_ERR_DLLNOTFOUND:     lErr = 48: sErr = "The specified dynamic-link library was not found."
            Case SE_ERR_FNF:             lErr = 53: sErr = "File not found"
            Case SE_ERR_NOASSOC:         sErr = "No application is associated with this file type."
            Case SE_ERR_OOM:             lErr = 7: sErr = "Out of memory"
            Case SE_ERR_PNF:             lErr = 76: sErr = "Path not found"
            Case SE_ERR_SHARE:           lErr = 75: sErr = "A sharing violation occurred."
            Case Else:                   sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select
        Err.Raise lErr, , App.EXEName & ".GShell", sErr
        ShellEx = False
    End If

End Function

Public Sub ShortCut()

    On Error Resume Next

    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")

    Dim MyShortcut, _
        MyDesktop, _
        DesktopPath

    DesktopPath = WSHShell.specialfolders("AllUsersDesktop")

    Set MyShortcut = WSHShell.CreateShortcut(DesktopPath & "\" & ProjectName & ".lnk")

    With MyShortcut
        .targetpath = WSHShell.ExpandEnvironmentStrings(PathCheck(App.Path)) & ProjectName & ".exe"
        .WorkingDirectory = WSHShell.ExpandEnvironmentStrings(PathCheck(App.Path))
        .WindowStyle = 4
        .IconLocation = WSHShell.ExpandEnvironmentStrings(PathCheck(App.Path) & ProjectName & ".exe") & ", 0"
        .Save
    End With

    Err.Clear

End Sub

Public Function SoundX(ByVal pWord As String, Optional pAccuracy As Byte = 4) As String

    On Error GoTo LocalError
    '  Char importance "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    Dim x As Integer
    Dim CChar As String

    If pAccuracy > 10 Then
        pAccuracy = 10                     ' maximum accuracy allowed
    ElseIf pAccuracy < 4 Then
        pAccuracy = 4                      ' minimum accuracy allowed
    End If

    ' account for the first character
    pAccuracy = pAccuracy - 1
    pWord = UCase(pWord)

    ' strip all invalid characters
    For x = 1 To Len(pWord)
        If Asc(Mid(pWord, x, 1)) < 65 Or _
            Asc(Mid(pWord, x, 1)) > 90 Then
            Mid(pWord, x, 1) = "@"           ' assign a catchable value
        End If
    Next x

    pWord = Trim(pWord)

    SoundX = pWord
    ' assign values to the string
    SoundX = Replace(SoundX, "A", "0")
    SoundX = Replace(SoundX, "E", "0")
    SoundX = Replace(SoundX, "I", "0")
    SoundX = Replace(SoundX, "O", "0")
    SoundX = Replace(SoundX, "U", "0")
    SoundX = Replace(SoundX, "Y", "0")
    SoundX = Replace(SoundX, "H", "0")
    SoundX = Replace(SoundX, "W", "0")
    SoundX = Replace(SoundX, "B", "1")
    SoundX = Replace(SoundX, "P", "1")
    SoundX = Replace(SoundX, "F", "1")
    SoundX = Replace(SoundX, "V", "1")
    SoundX = Replace(SoundX, "C", "2")
    SoundX = Replace(SoundX, "S", "2")
    SoundX = Replace(SoundX, "G", "2")
    SoundX = Replace(SoundX, "J", "2")
    SoundX = Replace(SoundX, "K", "2")
    SoundX = Replace(SoundX, "Q", "2")
    SoundX = Replace(SoundX, "X", "2")
    SoundX = Replace(SoundX, "Z", "2")
    SoundX = Replace(SoundX, "D", "3")
    SoundX = Replace(SoundX, "T", "3")
    SoundX = Replace(SoundX, "L", "4")
    SoundX = Replace(SoundX, "M", "5")
    SoundX = Replace(SoundX, "N", "5")
    SoundX = Replace(SoundX, "R", "6")

    CChar = Left(SoundX, 1)

    For x = 2 To Len(SoundX)
        If Mid(SoundX, x, 1) = CChar Then
            Mid(SoundX, x, 1) = "@"
        Else
            CChar = Mid(SoundX, x, 1)
        End If
    Next x

    SoundX = Replace(SoundX, "@", "")
    SoundX = Mid(SoundX, 2)
    SoundX = Replace(SoundX, "0", "")
    SoundX = SoundX & String(pAccuracy, "0")
    SoundX = Left(pWord, 1) & Left(SoundX, pAccuracy)

Exit Function

LocalError:
End Function

Function SoundsLike(ByVal WordIN$, Optional Accuracy As Integer = 5, Optional AsNumeric As Boolean = False) As String

    On Error GoTo LocalError

    If Len(WordIN) < 2 Then Exit Function

    Dim x       As Long
    Dim ThisVal As Integer
    Dim PrevVal As Integer
    Dim OldVal  As Integer

    Const Decifer = "01230120022455012623010202"
    'Char. Values    ABCDEFGHIJKLMNOPQRSTUVWXYZ

    WordIN = UCase$(TrimALL(WordIN))                        ' Soundex is case-INSENSITIVE

    For x = 1 To Len(WordIN)
        ThisVal = Asc(Mid$(WordIN, x, 1)) - 64              ' obtain CURRENT Character Value
        If ThisVal >= 1 And ThisVal <= 26 Then              ' discard NON-Alphabetic chars
            If OldVal = 0 Then
                OldVal = Asc(Mid$(Decifer, Asc(Mid(WordIN, x, 1)) - 64))
                If Not AsNumeric Then SoundsLike = Mid$(WordIN, x, 1)    ' the FIRST letter is copied in the result
            End If
            PrevVal = Asc(Mid$(Decifer, ThisVal, 1))        ' convert to a DIGIT
            If PrevVal <> 48 And PrevVal <> OldVal Then     ' don't insert REPEATED digits
                SoundsLike = SoundsLike & Chr$(PrevVal)     ' concatenate the PREVIOUS Character Value
            End If
            OldVal = PrevVal
        End If
        If Len(SoundsLike) >= Accuracy Then Exit For        ' do not continue if ACCURACY is reached
    Next x

    If Not AsNumeric Then
        SoundsLike = SoundsLike & String(Accuracy, "0")
        SoundsLike = Left$(SoundsLike, Accuracy)
    End If

Exit Function

LocalError:
    SoundsLike = "*****"
End Function

Public Function Wait(ByVal HowLong As Single, Optional AsMilliSecs As Boolean = False) As Boolean

    Call Sleep(Int(HowLong * IIf(AsMilliSecs, 1, 1000#)))
    Wait = True  'Notify caller when finished

End Function

'+-------------------+
'|DATE TIME Functions|
'+-------------------+
Public Function BusinessDays(ByVal Date1 As Variant, Optional ByVal Date2 As Variant, Optional WeekEndsOnly As Boolean = False) As Long

    If Not IsDate(Date1) Then Exit Function

    If Not IsDate(Date2) Then Date2 = Date

    Date1 = CDate(Date1)
    Date2 = CDate(Date2)

    If Date1 > Date2 Then
        'Swap values
        Dim lTemp As Date
        lTemp = Date2
        Date2 = Date1
        Date1 = lTemp
    End If

    BusinessDays = 0

    If WeekEndsOnly Then    'ONLY count the Days on the WeekEnd(s)
        While Date1 <= Date2
            If IsWeekend(Date1) Then BusinessDays = BusinessDays + 1
            Date1 = Date1 + 1
        Wend
    Else                    'ONLY count the Days during the Week
        While Date1 <= Date2
            If Not IsWeekend(Date1) Then BusinessDays = BusinessDays + 1
            Date1 = Date1 + 1
        Wend
    End If

End Function

Function BusinessDaysAdd(ByVal SomeDate As Date, ByVal Days As Long, Optional ByVal SaturdayIsHoliday As Boolean = True) As Date

    Do While Days
        SomeDate = SomeDate + Sgn(Days)   ' increment or decrement the date
        ' check that it is a week day
        If Weekday(SomeDate) <> vbSunday And (Weekday(SomeDate) <> vbSaturday Or Not SaturdayIsHoliday) Then
            ' days becomes closer to zero
            Days = Days - Sgn(Days)
        End If
    Loop

    BusinessDaysAdd = SomeDate

End Function

Public Function DaysInMonth(Optional ByVal SomeDate As Variant) As Integer

    '!!!Warning!!!
    '   If Somedate is NOT a date then Current year is assumed
    '   Ambiguous dates are assumed to be mm/dd/yyyy

    Dim lMonth As Integer
    Dim lYear  As Integer

    If IsMissing(SomeDate) Then SomeDate = Date ' Assume Today

    If IsDate(SomeDate) Then
        lYear = Year(SomeDate)
        lMonth = Month(SomeDate)
    ElseIf IsNumeric(SomeDate) Then
        lYear = Year(Date)
        lMonth = IIf(SomeDate > 0 And SomeDate < 13, CInt(SomeDate), 0)
    ElseIf VarType(SomeDate) = vbString Then
        lYear = Year(Date)
        Select Case UCase(Left$(SomeDate, 3))
            Case "FEB":                                             lMonth = 2
            Case "JAN", "MAR", "MAY", "JUL", "AUG", "OCT", "DEC":   lMonth = 1
            Case "APR", "JUN", "SEP", "NOV":                        lMonth = 4
        End Select
    End If

    Select Case lMonth
        Case 2:                     DaysInMonth = IIf(IsLeapYear(SomeDate), 29, 28)
        Case 1, 3, 5, 7, 8, 10, 12: DaysInMonth = 31
        Case 4, 6, 9, 11:           DaysInMonth = 30
    End Select

End Function

Public Function DaysInYear(ByVal SomeValue As Variant) As Integer
    If IsDate(SomeValue) Or IsNumeric(SomeValue) Then DaysInYear = IIf(IsLeapYear(SomeValue), 366, 365)
End Function

Public Function DayPart(Optional vTime As Variant = "", Optional Greeting As String) As String

    If vTime = "" Then vTime = Time

    If IsDate(vTime) Then vTime = FormatDateTime(vTime, vbShortTime)

    If (vTime >= #12:00:00 AM#) And (vTime < #12:00:00 PM#) Then
        DayPart = "morning"
    ElseIf (vTime > #12:00:00 AM#) And (vTime < #5:00:00 PM#) Then
        DayPart = "afternoon"
    Else
        DayPart = "evening"
    End If

    If Len(Greeting) > 0 Then DayPart = TrimALL(Greeting & " " & DayPart)

End Function

Public Function EndOfMonth(SomeDate As Variant) As Date

    If IsDate(SomeDate) Then
        EndOfMonth = DateAdd("m", 1, SomeDate)
        EndOfMonth = DateSerial(Year(EndOfMonth), Month(EndOfMonth), 1)
        EndOfMonth = DateAdd("d", -1, EndOfMonth)
    End If

End Function

Function EndOfWeek(ByVal SomeDate As Date) As Date

    If IsDate(SomeDate) Then
        EndOfWeek = FormatDateTime(SomeDate - Weekday(SomeDate) + 7, vbGeneralDate)
    End If

End Function

Public Function ExactAge(StartDate As Variant, Optional YearsOnly As Boolean) As String

    If Not IsDate(StartDate) Then Exit Function

    Dim yer As Integer, mon As Integer, d As Integer
    Dim dt As Date
    Dim sAns  As String

    dt = CDate(StartDate)

    If dt > Now Then Exit Function

    yer = Year(dt)
    mon = Month(dt)
    d = Day(dt)
    yer = Year(Date) - yer
    mon = Month(Date) - mon
    d = Day(Date) - d

    If Sgn(d) = -1 Then
        d = 30 - Abs(d)
        mon = mon - 1
    End If

    If Sgn(mon) = -1 Then
        mon = 12 - Abs(mon)
        yer = yer - 1
    End If

    sAns = yer & IIf(YearsOnly, "", " year(s) " & mon & " month(s) " & d & " day(s)")
    ExactAge = sAns

End Function

Function HowOld(varBirthDate As Variant) As Integer

   If Not IsDate(varBirthDate) Then Exit Function

   Dim varAge As Variant

   varAge = DateDiff("yyyy", varBirthDate, Now)

   If Date < DateSerial(Year(Now), Month(varBirthDate), Day(varBirthDate)) Then
      varAge = varAge - 1
   End If

   HowOld = CInt(varAge)

End Function

Function IsGoodDate(DateIN As Variant, Optional CompDate As Variant, Optional ByVal DateType As GoodDateTypes = AnyDate, Optional ByVal IsRequired As Boolean = True) As Boolean

    On Error GoTo LocalError

    Dim lDateIN As Double
    Dim lDateXX As Double

    If IsMissing(CompDate) Then
        CompDate = Now
    End If

    If IsDate(DateIN) And IsDate(CompDate) Then
        lDateIN = CDec(CDate(DateIN))
        lDateXX = CDate(CompDate)
        Select Case DateType
            Case FutureDate:    IsGoodDate = CLng(lDateIN) > CLng(lDateXX)
            Case FutureTime:    IsGoodDate = lDateIN > lDateXX
            Case PastDate:      IsGoodDate = CLng(lDateIN) < CLng(lDateXX)
            Case PastTime:      IsGoodDate = lDateIN < lDateXX
            Case Else:          IsGoodDate = True
        End Select
    End If
Exit Function

LocalError:
    IsGoodDate = False
End Function

Public Function IsLeapYear(ByVal SomeValue As Variant) As Boolean

    On Error GoTo LocalError

    Dim intYear As Integer

    'The 3 Golden rules are:
    '1. True if it is divisible by 4
    '2. False if it is divisible by 100
    '3. TRUE if it is divisble by 400
    If IsDate(SomeValue) Then intYear = Year(SomeValue) Else intYear = CInt(SomeValue)

    If TypeName(intYear) = "Integer" Then
        'Using DateSerial Function
        IsLeapYear = Day(DateSerial(intYear, 3, 0)) = 29
        'IsLeapYear = Day(DateSerial(intYear, 2, 29)) = 29
        'Using Calculations
        'IsLeapYear = ((intYear Mod 4 = 0) And (intYear Mod 100 <> 0) Or (intYear Mod 400 = 0))
    End If
Exit Function

LocalError:
End Function

Public Function IsWeekend(ByVal SomeDate As Variant) As Boolean

    If IsDate(SomeDate) Then If (Weekday(SomeDate) = 1) Or (Weekday(SomeDate) = 7) Then IsWeekend = True

End Function

Public Function NextDate(ByVal d As Date, Optional ByVal WhatDay As VbDayOfWeek = vbSaturday, Optional GetNext As Boolean = True) As Date

    NextDate = (((d - WhatDay + GetNext) \ 7) - GetNext) * 7 + WhatDay

End Function

Public Function WeekNo(Optional SomeDate As Variant) As Integer

    WeekNo = DatePart("ww", IIf(IsDate(SomeDate), SomeDate, Date))

End Function
Public Function Today() As Date:        Today = Date:                           End Function
Public Function Tomorrow() As Date:     Tomorrow = DateAdd("d", 1, Date):       End Function
Public Function Yesterday() As Date:    Yesterday = DateAdd("d", -1, Date):     End Function

Private Sub Class_Initialize()
    Set MyComputer = New WinInfo
    Set MyLocale = New Locale
End Sub

Private Sub Class_Terminate()
    Set MyComputer = Nothing
    Set MyLocale = Nothing
End Sub
