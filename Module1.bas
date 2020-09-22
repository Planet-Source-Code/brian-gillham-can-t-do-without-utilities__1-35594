Attribute VB_Name = "Common"
Option Explicit


#If False Then  'Keep Case
    Private LoadSettings
    Private SaveSettings
#End If

Public Enum FormState_
    FormState_LoadSettings
    FormState_SaveSettings
End Enum

'Variables
Public gTEMP        As Variant                  ' Use for ANYTHING

'Objects used by Application
Public oUtils       As New Utils                ' API and General Utilities
Public oRegistry    As New Registry                 ' Registry Class

Public Sub FormSettings(ByVal Setting As FormState_, ByVal oForm As Form, Optional lTop As Long, Optional lLeft As Long, Optional lWidth As Long, Optional lHeight As Long)

    Dim fName As String

    With oForm
        fName = IIf(Len(.Tag) > 0, .Tag, .Name)
        If Setting = FormState_LoadSettings Then
            .Top = IIf(lTop > 0, lTop, CLng(oRegistry.GetSetting(fName, "Top", "0")))
            .Left = IIf(lLeft > 0, lLeft, CLng(oRegistry.GetSetting(fName, "Left", "0")))

            If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
                .Width = IIf(lWidth > 0, lWidth, CLng(oRegistry.GetSetting(fName, "Width", Screen.Width \ 2)))
                .Height = IIf(lHeight > 0, lHeight, CLng(oRegistry.GetSetting(fName, "Height", Screen.Height \ 2)))
            End If

            If .Top <= 0 Or .Left <= 0 Then
                .Width = Screen.Width \ 2
                .Height = Screen.Height \ 2
                .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
            End If

            If CBool(oRegistry.GetSetting(fName, "WasMaximized", False)) = True Then
                .WindowState = vbMaximized
            End If

            ' Ensure that the window is still visible on the screen
            '   if the screen size is smaller
            If .Left + .Width > Screen.Width Then .Left = Max(0, Screen.Width - .Width)
            .Width = Min(Screen.Width, .Width)
            If .Top + .Height > Screen.Height Then .Top = Max(0, Screen.Height - .Height)
            .Height = Min(Screen.Height, .Height)
            .Left = Max(0, .Left)
            .Top = Max(0, .Top)
        ElseIf Setting = FormState_SaveSettings Then
            oRegistry.SaveSetting fName, "WasMaximized", .WindowState = vbMaximized
            If .WindowState <> vbMinimized And .WindowState <> vbMaximized Then
                oRegistry.SaveSetting fName, "Top", .Top
                oRegistry.SaveSetting fName, "Left", .Left
                oRegistry.SaveSetting fName, "Width", .Width
                oRegistry.SaveSetting fName, "Height", .Height
            End If
        End If
    End With

End Sub

Public Function IsInArray(FindValue As Variant, arrSearch As Variant) As Boolean

    On Error GoTo LocalError

    If Not IsArray(arrSearch) Then Exit Function
    IsInArray = InStr(1, vbNullChar & Join(arrSearch, vbNullChar) & vbNullChar, vbNullChar & FindValue & vbNullChar) > 0

LocalError: 'Justin (just in case)

End Function

Public Function Min(a As Variant, b As Variant) As Variant
    ' JM: Return min value
    If a < b Then Min = a Else Min = b
End Function

Public Function Max(a As Variant, b As Variant) As Variant
    ' JM: Return max value
    If a > b Then Max = a Else Max = b
End Function

Public Function PathCheck(ByVal PathName$, Optional AltDelimiter$ = "") As String

    Dim Delimiter As String

    Delimiter = IIf(InStr(PathName, "/"), "/", "\")
    PathCheck = IIf(Right$(PathName, 1) = Delimiter, PathName, PathName & Delimiter)
    PathCheck = IIf(Len(AltDelimiter) = 0, PathCheck, Replace(PathCheck, Delimiter, AltDelimiter))

End Function

Public Function ProjectName() As String

    Static Result As String

    If Len(Result) = 0 Then
        On Error Resume Next
        ' cause a dummy, harmless error
        Err.Raise 999
        ' retrieve the project name from the Err.Source property
        Result = Err.Source
        On Error GoTo 0
    End If

    ProjectName = Result

End Function

Public Function StripText(ByRef TextIN$, Optional Unwanted$)

    Dim currLoc As Integer
    Dim tmpChar As String

    If Len(Unwanted) = 0 Then Unwanted = "~`!@#$%^&*{}[]()_+-=|\?/.>,<" & Chr(34)

    For currLoc = 1 To Len(TextIN)
        tmpChar = Mid$(TextIN, currLoc, 1)
        If InStr(Unwanted, tmpChar) Then
            tmpChar = " " 'replace with a space
        End If
        StripText = StripText & tmpChar
    Next

    StripText = TrimALL(StripText)

End Function

Public Function TrimALL(ByVal TextIN As String) As String

    TrimALL = Trim(TextIN)

    While InStr(TrimALL, String(2, " ")) > 0
        TrimALL = Replace(TrimALL, String(2, " "), " ")
    Wend

End Function

Public Function TrimNull(InString As String) As String

    Dim Pos As Long

    Pos = InStr(InString, Chr$(0))
    TrimNull = IIf(Pos > 0, Left$(InString, Pos - 1), InString)
End Function

Public Sub UnloadForms(oForm As Form)

    Dim lForm As Form
    
    For Each lForm In Forms
        If lForm.Name <> oForm.Name Then
            Unload lForm
            Set lForm = Nothing
        End If
    Next lForm
    Set oForm = Nothing

End Sub

Public Function Version(Optional FullVersion As Boolean = False) As String

    Dim lVersion As String

    lVersion = App.Major & "." & ZeroPad(App.Minor, 2) & "." & ZeroPad(App.Revision, 4)
    If FullVersion Then
        Version = ProjectName & " v" & lVersion
        If App.Major < 1 And App.Minor < 1 Then
            Version = Version & " (beta)"
        ElseIf App.Major > 0 And App.Minor < 1 Then
            Version = Version & " (preview)"
        End If
    Else
        Version = lVersion
    End If
    
End Function

Public Function ZeroPad(pNumber As Variant, pLength As Integer) As String

    If IsNumeric(pNumber) And IsNumeric(pLength) And (pLength > Len(pNumber)) Then
        Dim Padding As String
        Padding = String(pLength, "0")
        ZeroPad = Padding & CStr(pNumber)
        ZeroPad = Right$(ZeroPad, pLength)
    Else
        ZeroPad = pNumber
    End If

End Function

