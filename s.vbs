' Progress-Bar

Option Explicit

Dim ScreenWidth, ScreenHeight, Title_ProgressDisplay, ProgressBarWidth, ProgressBarHeight

Dim objIE ' objExplorer needy for Function "ProgressDisplay"

GetMonitorProperties

Title_ProgressDisplay = "Google Utility Tool"
ProgressBarWidth = 400
ProgressBarHeight = 150

' Program

Main

' End of Program

' Procedures

Sub Main: Dim Progress
    ProgressDisplay "Open",""
    For Progress = 0 To 100 ' Get Progress from program
        ShowProgress(Progress)
        WScript.Sleep 20
    Next:   ProgressDisplay "Close",""
End Sub

Sub ShowProgress(Progress0to100): Dim Text, k: k = (ProgressBarWidth - 2*19-21)
        Text = "<p align=""center"">Repairing " & CStr(Progress0to100) & " %</p>" & _
            "<table border=""0"" cellpadding=""0"" cellspacing=""0""><tr><td width=""" & _
            CStr(k*Progress0to100/100) & _
            """ height=""15"" bgcolor=""#0000FF"">&nbsp;</td></tr></table>"
        ProgressDisplay "Display",Text: WScript.Sleep 20
End Sub

Sub ProgressDisplay (Mode, AnyText): Dim String1, String2, colItems, objItem
    ' Mode = Open, Display, Close
    ' AnyText only used in Display-Mode
    Mode = UCase(Left(Mode,1)) & LCase(Right(Mode,Len(Mode)-1))
    Select Case Mode
        Case "Open"
            Set objIE = CreateObject("InternetExplorer.Application")
            With objIE
                .Navigate "about:blank"
                .ToolBar = False: .StatusBar = False
                .Width = ProgressBarWidth: .Height = ProgressBarHeight
                .Left = (ScreenWidth - ProgressBarWidth) \ 2
                .Top = (ScreenHeight - ProgressBarHeight) \ 2
                .Visible = True
                With .Document
                    .title = Title_ProgressDisplay
                    .ParentWindow.focus()
                    With .Body.Style
                        .backgroundcolor = "#F0F7FE"
                        .color = "#0060FF"
                        .Font = "11pt 'Calibri'"
                    End With
                End With: While .Busy: Wend
            End With
        Case "Display"
            On Error Resume Next ' for clicking away the bar while running
            If Err.Number = 0 Then
                With objIE.Document
                    .Body.InnerHTML = AnyText: WScript.Sleep 200: .ParentWindow.focus()
                End With
            End If
        Case "Close": WScript.Sleep 100: objIE.Quit
    End Select
End Sub

Sub GetMonitorProperties
    Dim strComputer, objWMIService, objItem, colItems, VMD: strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController")
    For Each objItem In colItems: VMD = objItem.VideoModeDescription: Next
    ' VMD = 1280 x 1024 x 4294967296 Farben
    VMD = Split(VMD, " x "): ScreenWidth = Eval(VMD(0)): ScreenHeight = Eval(VMD(1))
End Sub