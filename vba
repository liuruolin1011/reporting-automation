Option Explicit

' 发送 Clock In 邮件
Sub SendClockInEmail()
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = "manager@example.com" ' 替换为你的经理邮箱
        .Subject = "Ruolin Liu - Clock In/Clock Out"
        .Body = "Clock In - 9:00 AM" & vbCrLf & "Best regards," & vbCrLf & "Ruolin Liu"
        .Send
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

    MsgBox "Clock In - 9:00 AM 邮件已发送！", vbInformation
End Sub

' 查找并回复最后一封打卡邮件
Sub ReplyToLastEmail(replyText As String)
    Dim OutApp As Object
    Dim OutNamespace As Object
    Dim Inbox As Object
    Dim Items As Object
    Dim Mail As Object
    Dim ReplyMail As Object
    Dim i As Integer
    Dim Found As Boolean

    Set OutApp = CreateObject("Outlook.Application")
    Set OutNamespace = OutApp.GetNamespace("MAPI")
    Set Inbox = OutNamespace.GetDefaultFolder(5) ' 5 = 已发送邮件
    Set Items = Inbox.Items

    ' 按时间排序
    Items.Sort "[SentOn]", True

    ' 查找最近的 Clock In/Clock Out 邮件
    Found = False
    For i = 1 To Items.Count
        Set Mail = Items(i)
        If InStr(1, Mail.Subject, "Ruolin Liu - Clock In/Clock Out", vbTextCompare) > 0 Then
            Set ReplyMail = Mail.Reply
            ReplyMail.Body = replyText & vbCrLf & "Best regards," & vbCrLf & "Ruolin Liu" & vbCrLf & vbCrLf & ReplyMail.Body
            ReplyMail.Send
            MsgBox replyText & " 邮件已发送！", vbInformation
            Found = True
            Exit For
        End If
    Next i

    If Not Found Then
        MsgBox "未找到打卡邮件，无法回复！", vbExclamation
    End If

    Set Mail = Nothing
    Set ReplyMail = Nothing
    Set Items = Nothing
    Set Inbox = Nothing
    Set OutNamespace = Nothing
    Set OutApp = Nothing
End Sub

' 预设的打卡邮件计划
Sub AutoClockInOut()
    ' 发送 9:00 AM Clock In
    Application.OnTime TimeValue("09:00:00"), "SendClockInEmail"

    ' 1:00 PM 发送 Clock Out
    Application.OnTime TimeValue("13:00:00"), "ClockOut_1PM"

    ' 1:30 PM 发送 Clock In
    Application.OnTime TimeValue("13:30:00"), "ClockIn_1_30PM"

    ' 6:30 PM 发送 Clock Out
    Application.OnTime TimeValue("18:30:00"), "ClockOut_6_30PM"

    MsgBox "打卡计划已设定！", vbInformation
End Sub

' 具体的回复任务
Sub ClockOut_1PM()
    ReplyToLastEmail "Clock Out - 1:00 PM"
End Sub

Sub ClockIn_1_30PM()
    ReplyToLastEmail "Clock In - 1:30 PM"
End Sub

Sub ClockOut_6_30PM()
    ReplyToLastEmail "Clock Out - 6:30 PM"
End Sub