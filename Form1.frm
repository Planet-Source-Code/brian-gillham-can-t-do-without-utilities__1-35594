VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "My System"
      Height          =   345
      Left            =   1170
      TabIndex        =   2
      Top             =   60
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "My Display"
      Height          =   345
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   450
      Width           =   3945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Author: Brian Gillham - FailSafe Systems
' Legal:  Do with this what you wish. Just do not claim credit for it.
'         Small ask for a lot of utilities huh!
' This app utilises the Utilities Class which has too many
' functions to explain through the medium of this Demo.
' Basically it expose many functions which are used day to day.
' 1. Utilities  - General Utilities
' 2. Wininfo    - Information about your Computer
' 3. Local      - Information about your your Region / Locale
' Just work through the Properties and Functions of each Utility
' and the usage will becom quite apparent.
' With the execption of the FileKill function

Private Sub Command1_Click()
    oUtils.MyComputer.ControlPanel cplDisplay
End Sub

Private Sub Command2_Click()
    oUtils.MyComputer.ControlPanel cplSystem
End Sub

Private Sub Form_Load()

    Dim lText As String

    With oRegistry
        .AppName = "MyApp"
        .Company = "FailSafe"
    End With

    ' get Saved Form Settings
    Common.FormSettings FormState_LoadSettings, Me

    lText = "Information about you" & vbCrLf
    lText = lText & "=========================" & vbCrLf
    lText = lText & "------------" & vbCrLf
    lText = lText & "Your Locale" & vbCrLf
    lText = lText & "------------" & vbCrLf
    lText = lText & "Your Country: " & oUtils.MyLocale.Country & vbCrLf
    lText = lText & "Your Currency: " & oUtils.MyLocale.CurrencySpecifier & vbCrLf
    lText = lText & "Your Date Fromat: " & oUtils.MyLocale.DateFormat & vbCrLf
    lText = lText & "Your Language: " & oUtils.MyLocale.LanguageName & vbCrLf
    lText = lText & "------------" & vbCrLf
    lText = lText & "Your Computer" & vbCrLf
    lText = lText & "------------" & vbCrLf
    lText = lText & "Your UserName: " & oUtils.MyComputer.CurrentUser & vbCrLf
    lText = lText & "Computer Name: " & oUtils.MyComputer.Name & vbCrLf
    lText = lText & "Computer Win Version: " & oUtils.MyComputer.OSVersion & vbCrLf
    lText = lText & "Computer Started: " & oUtils.MyComputer.Started & vbCrLf
    lText = lText & "Computer running: " & oUtils.MyComputer.Running & vbCrLf
    lText = lText & "Computer has Sound Card: " & oUtils.MyComputer.SoundCard & vbCrLf

    Text1.Text = lText

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    With Me
        Text1.Move .ScaleLeft, .Command1.Top + .Command1.Height, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Remember Form Settings
    Common.FormSettings FormState_SaveSettings, Me
    Set oRegistry = Nothing
    Set oUtils = Nothing

End Sub
