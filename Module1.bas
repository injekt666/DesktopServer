Attribute VB_Name = "Module1"
' This OCX control is made by M. Schermer from the Netherlands
' This OCX can be used to capture the fullscreen or the active
' window. I'm take no care about damage on your computer but
' that is impossible.
'
' To use this control compile it to an OCX and start an new project
' Goto COMPONENTS and add the compiled OCX file
' Now you can use it to add the control to your form and for
' example:

'   Dim FileToOpen as String
'
'   FileToOpen = CaptureScreen1.CaptureActiveScreen("C:\MyScreen.BMP", BMP, True)
'   Image1.Picture = LoadPictures(FileToOpen)
'
' PS: If you want to make a loop (for example: refresh picture every 1 second)
' check the CaptureScreen1.ReadyState
' it returns True if the process is ready
' it returns False if the process is not ready
'
'                           Have fun using this control

    Public ReadyState As Boolean
    Public strOutputFile As String
    'Public rectactive As RECT

Public Function CaptureScreen() As Picture
    Set CaptureScreen = CaptureWindow(GetDesktopWindow, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function

'Public Function CaptureActiveWindow() As Picture
 '   Call GetWindowRect(GetForegroundWindow, rectactive)
'    Set CaptureActiveWindow = CaptureWindow(GetForegroundWindow, False, 0, 0, rectactive.Right - rectactive.Left, rectactive.Bottom - rectactive.Top)
'End Function

'Public Function CaptureWindowHWND(WindowHWND As Long) As Picture
 '   Call GetWindowRect(WindowHWND, rectactive)
'    Set CaptureWindowHWND = CaptureWindow(WindowHWND, False, 0, 0, rectactive.Right - rectactive.Left, rectactive.Bottom - rectactive.Top)
'End Function


Public Sub SavePictureToBMP(strFileName As String)
    Dim CheckFile As Boolean
        
        CheckFile = FileExists(strFileName)
        If CheckFile = True Then
            Kill strFileName
        End If
    
    DoEvents
    Call SavePicture(Form1.Picture, strFileName)
    DoEvents
    strOutputFile = strFileName
End Sub

Public Sub SavePictureToJPG(strFileName As String)
    Dim c As New cDIBSection
    Dim CheckFile As Boolean
    Set c = New cDIBSection
    DoEvents
'    Call SavePicture(frmmain.Picture, strFileName)
'    DoEvents
    c.CreateFromPicture LoadPicture(strFileName)
    
    DoEvents
    Call SaveJPG(c, App.Path & "\" & "screenie.jpg")
    DoEvents
    strOutputFile = strFileName

End Sub


Private Function FileExists(filename) As Boolean
  On Error GoTo ErrorHandler
   FileExists = (Dir(filename) <> "")
Exit Function

ErrorHandler:
    FileExists = False
End Function

Private Sub CreateDirectory(strDirToCreate As String)
  On Error Resume Next
    Dim FindBeginPos As Integer
    Dim FindEndPos As Integer
    Dim DirIsCreated As Boolean
    Dim CreatePath As String
    
    FindBeginPos = InStr(strDirToCreate, ":")
    If FindBeginPos <> 0 Then
        CreatePath = Mid(strDirToCreate, FindBeginPos - 1, FindBeginPos) & "\"
            Do Until DirIsCreated = True
                FindBeginPos = InStr(strDirToCreate, "\")
                    If FindBeginPos <> 0 Then
                        FindEndPos = InStr(Mid(strDirToCreate, FindBeginPos + 1), "\")
                            If FindEndPos <> 0 Then
                                FindEndPos = (FindEndPos + FindBeginPos) - 2
                                CreatePath = CreatePath & Mid(strDirToCreate, FindBeginPos + 1, FindEndPos - 2) & "\"
                                MkDir CreatePath
                                strDirToCreate = Mid(strDirToCreate, FindEndPos)
                            Else
                                DirIsCreated = True
                            End If
                    End If
            Loop
    End If
End Sub

