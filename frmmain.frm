VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{69734A7F-9DD3-11D3-A30A-000001165224}#2.0#0"; "System Tray Icon.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   Caption         =   "Desktop Server 1.6.5"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   720
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   480
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin SystemTrayIcon.SysIcon SysIcon1 
      Left            =   1440
      Top             =   4200
      _ExtentX        =   2355
      _ExtentY        =   2143
      NormalPicture   =   "frmmain.frx":0442
      AnimPicture     =   "frmmain.frx":0894
      IconText        =   "Desktop Server"
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "To Tray"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings:"
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox Check4 
         Caption         =   "Enable Webcam Tag"
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Sound When Making ScreenShot"
         Height          =   495
         Left            =   2040
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hide Desktop/Cam Mode"
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "To Tray On Start"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         Caption         =   "Connections"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1695
         Begin VB.Label current 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Custom Message (Use HTML IF YOU WANT)"
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   3975
         Begin VB.TextBox Text2 
            Height          =   1935
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Text            =   "frmmain.frx":0CE6
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Port"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Text            =   "2001"
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Checking Updates Please Wait...."
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   720
      TabIndex        =   11
      Top             =   960
      Width           =   4095
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Checking For Updates on Novaslp.Net...."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim hits As Integer


Private Sub Check4_Click()
MsgBox "Requires restart for settings to take effect."
End Sub

Private Sub Command1_Click()
savesettings

    '// Disable all callbacks
    capSetCallbackOnError lwndC, vbNull
    capSetCallbackOnStatus lwndC, vbNull
    capSetCallbackOnYield lwndC, vbNull
    capSetCallbackOnFrame lwndC, vbNull
    capSetCallbackOnVideoStream lwndC, vbNull
    capSetCallbackOnWaveStream lwndC, vbNull
    capSetCallbackOnCapControl lwndC, vbNull
    

End
End Sub

Private Sub Command3_Click()
Me.Visible = False
End Sub

Private Sub Command4_Click()
frmabout.Visible = True
End Sub

Private Sub Command5_Click()
If Command5.Caption = "Stop" Then
frmmain.Winsock1.Close
Command5.Caption = "Start"
Else
frmmain.Winsock1.Close
Winsock1.LocalPort = Text1.text
frmmain.Winsock1.Listen
Command5.Caption = "Stop"
End If
End Sub

Private Sub Form_Load()
loadsetTings
frmmain.SysIcon1.ShowNormalIcon
Winsock1.LocalPort = Text1.text
Winsock1.Listen
For i = 1 To 200
Load Winsock2(i)
Next i
If Check1.Value = 1 Then
    Me.Visible = False
Else
    Me.Visible = True
End If
update

'start cam stuff

    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim Caps As CAPDRIVERCAPS
        
    '//Create Capture Window
    capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
   ' lwndC = capCreateCaptureWindowA(lpszName, WS_CAPTION Or WS_THICKFRAME Or WS_VISIBLE Or WS_CHILD, 0, 0, 160, 120, , 0)

    '// Set title of window to name of driver
    'SetWindowText lwndC, lpszName
    
    '// Set the video stream callback function
    capSetCallbackOnStatus lwndC, AddressOf MyStatusCallback
    capSetCallbackOnError lwndC, AddressOf MyErrorCallback
    
    '// Connect the capture window to the driver
    If capDriverConnect(lwndC, 0) Then
        '/////
        '// Only do the following if the connect was successful.
        '// if it fails, the error will be reported in the call
        '// back function.
        '/////
        '// Get the capabilities of the capture driver
        capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
        
        '// If the capture driver does not support a dialog, grey it out
        '// in the menu bar.
        If Caps.fHasDlgVideoSource = 0 Then mnuSource.Enabled = False
        If Caps.fHasDlgVideoFormat = 0 Then mnuFormat.Enabled = False
        If Caps.fHasDlgVideoDisplay = 0 Then mnuDisplay.Enabled = False
        
        '// Turn Scale on
        capPreviewScale lwndC, True
            
        '// Set the preview rate in milliseconds
        capPreviewRate lwndC, 66
        
        '// Start previewing the image from the camera
        capPreview lwndC, True
            
        '// Resize the capture window to show the whole image
        ResizeCaptureWindow lwndC

    End If

End Sub

Private Sub Form_Terminate()
savesettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
savesettings
End Sub

Private Sub SysIcon1_IconLeftDouble()
Me.Visible = True
End Sub

Private Sub update()
Dim version As String
Dim currentver As String
On Error GoTo damn
currentver = Inet1.OpenURL("http://www.novaslp.net/update/desktopserver.ver")
version = App.Major & "." & App.Minor & "." & App.Revision
If currentver = version Then
 Frame6.Visible = False
 Exit Sub
Else
 MsgBox "Your version is out of date. Please Update download the newest version from. Www.Novaslp.Net"
 End
End If
Exit Sub
damn:
MsgBox "The server failed to return the current version. Please try again later or check your internet connection."
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer
For i = 0 To 200
If Winsock2(i).State = sckClosed Then
Winsock2(i).Close
Winsock2(i).Accept (requestID)
current.Caption = current.Caption + 1
Exit Sub
End If
Next i
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim strGet As String
Dim spc2 As Long
Dim page As String
Winsock2(Index).GetData strData
If Mid(strData, 1, 3) = "GET" Then
strGet = InStr(strData, "GET ")
spc2 = InStr(strGet + 5, strData, " ")
page = Trim(Mid(strData, strGet + 5, spc2 - (strGet + 4)))
If Right(page, 1) = "/" Then page = Left(page, Len(page) - 1)
On Error Resume Next
Winsock2(Index).SendData LoadFile(App.Path & "\" & page)
End If
End Sub
Public Function LoadFile(page As String) As String

End If
If page = App.Path & "\" & "screenie.jpg" Then
    If Check2.Value = 1 Then
     Open "hidden.jpg" For Binary As #1
     LoadFile = Input(FileLen(App.Path & "\" & "hidden.jpg"), #1)
     Close #1
    Else
     If Check3.Value = 1 Then
     playsound (App.Path & "\" & "default.wav")
     End If
     getpicture
     On Error GoTo hell
     Open page For Binary As #1
     LoadFile = Input(FileLen(page), #1)
     Close #1
    End If
ElseIf page = App.Path & "\screenie2.jpg" Then
 LoadFile = takewebcamshot
hell:
  If Err.Number = 76 Then LoadFile = "Cant find file! Oh no!"
Else
 hits = hits + 1
 LoadFile = tagify(Text2.text)
 End If
End Function

Private Sub Winsock2_SendComplete(Index As Integer)
Winsock2(Index).Close
current.Caption = current.Caption - 1
End Sub

Public Sub getpicture()
GetImage App.Path & "\" & "screenie.bmp"
End Sub

Public Function GetImage(OutputBitmap As String)
'Note that wherever it says Me.[anything],
'ME can be changed to FORMNAME.[anything].
'Me is a shortcut to the current form.
'Note that the form's AutoRedraw property must be true.

'This pauses the computer to give it time to hide
'the screenshot program. Without it, frmSShot would
'appear in the screenshot.
Sleep 100
DoEvents 'This refreshes after the delay

'Declare variables
Dim wHand As Long
Dim wDC As Long
Dim nHeight As Long, nWidth As Long

wHand = GetDesktopWindow 'Get the desktop's hWnd
wDC = GetDC(wHand) 'Convert hWnd to hDC

'Get screen resolution
nHeight = Screen.Height / Screen.TwipsPerPixelY
nWidth = Screen.Width / Screen.TwipsPerPixelX

'Take snapshot
BitBlt Me.hdc, 0, 0, nWidth, nHeight, wDC, 0, 0, vbSrcCopy

'Save to file
SavePicture Me.Image, OutputBitmap
SavePictureToJPG App.Path & "\" & "screenie.bmp"
'Clear form
Me.Cls
End Function

Public Function tagify(text) As String
    text = "<HTML>" & text
    text = Replace(text, "<time>", Time)
    text = Replace(text, "<date>", date)
    text = Replace(text, "<memorytot>", MemoryTotal())
    text = Replace(text, "<memoryavail>", MemoryAvailable())
    text = Replace(text, "<memoryused>", MemoryUsed())
    text = Replace(text, "<os>", WindowsVer(1))
    text = Replace(text, "<osmajor>", WindowsVer(2))
    text = Replace(text, "<osminor>", WindowsVer(3))
    text = Replace(text, "<osbuild>", WindowsVer(4))
    text = Replace(text, "<processor>", processorvars(2))
    text = Replace(text, "<processornum>", processorvars(1))
    text = Replace(text, "<uptimem>", Uptime("m"))
    text = Replace(text, "<uptimemm>", Uptime("mm"))
    text = Replace(text, "<uptimed>", Uptime("d"))
    text = Replace(text, "<uptimedd>", Uptime("dd"))
    text = Replace(text, "<uptimes>", Uptime("s"))
    text = Replace(text, "<uptimess>", Uptime("ss"))
    text = Replace(text, "<uptimeh>", Uptime("h"))
    text = Replace(text, "<uptimehh>", Uptime("h"))
    text = Replace(text, "<activewindow>", GetActiveWindow())
    text = Replace(text, "<cpuuse>", GetCPUUsage())
    text = Replace(text, "<appinfo>", applicationinfo())
    text = Replace(text, "<hits>", hits)
    text = Replace(text, "<picture>", "<p><img border='0' src='screenie.jpg'></p>")
    If Check4.Value = 1 Then
    text = Replace(text, "<webpicture>", "<p><img border='0' src='screenie2.jpg'></p>")
    Else
    text = Replace(text, "<webpicture>", "<p>Webcam Screenshot is currently disabled.</p>")
    End If
    tagify = text
End Function

Public Sub loadsetTings()
Text1.text = GetSetting(appname:="Desktopserver", section:="general", Key:="port", Default:="2001")
Text2.text = GetSetting(appname:="Desktopserver", section:="general", Key:="text", Default:="<html><head></head><body>This Is My Desktop It's Leet :)It has recieved: <hits>And the picture:<picture><appinfo>")
hits = GetSetting(appname:="Desktopserver", section:="general", Key:="hits", Default:=0)
Check1.Value = GetSetting(appname:="Desktopserver", section:="general", Key:="trayonstart", Default:=0)
Check2.Value = GetSetting(appname:="Desktopserver", section:="general", Key:="hidedesktop", Default:=0)
Check3.Value = GetSetting(appname:="Desktopserver", section:="general", Key:="sound", Default:=0)
Check4.Value = GetSetting(appname:="Desktopserver", section:="general", Key:="webcam", Default:=0)
End Sub
Public Sub savesettings()
SaveSetting "Desktopserver", "general", "port", Text1.text
SaveSetting "Desktopserver", "general", "text", Text2.text
SaveSetting "Desktopserver", "general", "hits", hits
SaveSetting "Desktopserver", "general", "trayonstart", Check1.Value
SaveSetting "Desktopserver", "general", "hidedesktop", Check2.Value
SaveSetting "Desktopserver", "general", "sound", Check3.Value
SaveSetting "Desktopserver", "general", "webcam", Check4.Value
End Sub

Public Sub playsound(filename As String)
MMControl1.Command = "Close"
On Error GoTo last
MMControl1.DeviceType = "WaveAudio"
On Error GoTo last
MMControl1.filename = filename
On Error GoTo last
MMControl1.Command = "Open"
On Error GoTo last
MMControl1.Command = "Play"
Exit Sub
last:
MMControl1.Command = "Close"
End Sub

Public Function takewebcamshot()

End Function
