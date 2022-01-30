Function YTDLP()
    Dim GetClipboard
    GetClipboard = CreateObject("htmlfile").ParentWindow.ClipboardData.getData("text") & ""

    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = "sm[0-9]+"
        .IgnoreCase = False
        .Global = True
    End With

    If reg.Test(GetClipboard) Then
        GetClipboard = "yt-dlp -i " + GetClipboard + " --config-location ni-config.ini"
        WScript.CreateObject("WScript.Shell").Run GetClipboard, 4, False
        Exit Function
    End If

    With reg
        .Pattern = "\/watch\?v="
        .IgnoreCase = False
        .Global = True
    End With

    If reg.Test(GetClipboard) Then
        GetClipboard = "yt-dlp -i " + GetClipboard + " --config-location yt-config.ini"
        WScript.CreateObject("WScript.Shell").Run GetClipboard, 4, False
        Exit Function
    End If

    With reg
        .Pattern = "http"
        .IgnoreCase = False
        .Global = True
    End With

    If reg.Test(GetClipboard) Then
        GetClipboard = "yt-dlp -i " + GetClipboard + " --config-location other-config.ini"
        WScript.CreateObject("WScript.Shell").Run GetClipboard, 4, False
        Exit Function
    End If
    
ENd Function

YTDLP()
