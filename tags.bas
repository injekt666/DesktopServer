Attribute VB_Name = "tags"
Public Function Punct(ByVal text As String) As String
text = text & " "

text = Replace(text, "im ", "I'm ", 1, 10)
text = Replace(text, "dont", "don't", 1, 10)
text = Replace(text, "hes", "he's", 1, 10)
text = Replace(text, "shes", "she's", 1, 10)
text = Replace(text, "thats", "that's", 1, 10)
text = Replace(text, " ive ", " I've ", 1, 10)
text = Replace(text, "cant", "can't", 1, 10)
text = Replace(text, "shouldnt", "shouldn't", 1, 10)
text = Replace(text, "havent", "haven't", 1, 10)
text = Replace(text, "youre", "you're", 1, 10)
text = Replace(text, "couldnt", "couldn't", 1, 10)
text = Replace(text, "wouldnt", "wouldn't", 1, 10)
text = Replace(text, "wont", "won't", 1, 10)
text = Replace(text, "isnt", "isn't", 1, 10)
text = Replace(text, "arent", "aren't", 1, 10)
text = Replace(text, "thats", "that's", 1, 10)
text = Replace(text, " i ", " I ", 1, 10)
If Left(text, 1) = LCase(Left(text, 1)) Then Mid(text, 1, 1) = UCase(Left(text, 1))
If Right(text, 1) <> "." And Right(text, 1) <> "?" And Right(text, 1) <> "!" And Right(text, 1) <> "," And Right(text, 1) <> "*" And Right(text, 1) <> UCase(Right(text, 1)) Then text = text & "."
Punct = Trim(text)
End Function
Public Function promode(ByVal text As String) As String
    
    If UCase(Mid(text, 1, 1)) <> Mid(text, 1, 1) Then
        Mid(text, 1, 1) = UCase(Mid(text, 1, 1))
    End If
    
    If Right(text, 1) <> "." And Right(text, 1) <> "!" And Right(text, 1) <> "?" Then
        text = text & "."
    End If
    promode = Punct(text)
End Function
Public Function tagify(text) As String
    text = Replace(text, "<time>", Time)
    text = Replace(text, "<date>", date)
    text = Replace(text, "<time>", Time)
    text = Replace(text, "<memorytot>", MemoryTotal())
    text = Replace(text, "<memoryavail>", MemoryAvailable())
    text = Replace(text, "<memoryused>", MemoryUsed())
    text = Replace(text, "<os>", WindowsVer(1))
    text = Replace(text, "<osmajor>", WindowsVer(2))
    text = Replace(text, "<osminor>", WindowsVer(3))
    text = Replace(text, "<osbuild>", WindowsVer(4))
    text = Replace(text, "<processor>", processorvars(1))
    text = Replace(text, "<processornum>", processorvars(2))
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
    tagify = text
End Function
