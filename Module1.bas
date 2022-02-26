Attribute VB_Name = "Module1"
Global address
Sub spam()
end_row = Cells(1048575, 1).End(xlUp).Row
If Cells(6, 3).Value = "from_first_row" Then
    start_row = 2
Else
    start_row = Cells(6, 4)
End If
For x = start_row To end_row
address = Cells(x, 1)
Call Send_Mail
Next
End Sub
Sub Send_Mail()
    Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sSubject As String, sBody As String, sAttachment As String
 
    Application.ScreenUpdating = False
    On Error Resume Next
    'пробуем подключиться к Outlook, если он уже открыт
    Set objOutlookApp = GetObject(, "Outlook.Application")
    Err.Clear 'Outlook закрыт, очищаем ошибку
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    'раскомментировать строку, если в Outlook несколько учетных записей и нужно подключиться к конкретной(только если Outlook закрыть)
    '   [параметры]: Session.Logon "имя профиля","пароль",[показывать окно выбора профиля], [запускать в новой сессии]
    'objOutlookApp.Session.Logon "profile","1234",False, True
    Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
    'если не получилось создать приложение или экземпляр сообщения - выходим
    If Err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = address    'Кому(можно заменить значением из ячейки - sTo = Range("A1").Value)
    sSubject = Cells(2, 8)   'Тема письма(можно заменить значением из ячейки - sSubject = Range("A2").Value)
    sBody = Cells(2, 8)    'Текст письма(можно заменить значением из ячейки - sBody = Range("A3").Value)
    If Cells(2, 7).Value <> "" Then
        sAttachment = ThisWorkbook.Path & "\" & Cells(2, 7).Value   'Вложение(полный путь к файлу. Можно заменить значением из ячейки - sAttachment = Range("A4").Value)
    Else
        sAttachment = ""
    End If
 
    'создаем сообщение
    With objMail
        .to = sTo 'адрес получателя
        .CC = "" 'адрес для копии
        .BCC = "" 'адрес для скрытой копии
        .Subject = sSubject 'тема сообщения
        .Body = sBody 'текст сообщения
        '.HTMLBody = sBody 'если необходим форматированные текст сообщения(различные шрифты, цвет шрифта и т.п.)
        'добавляем вложение, если файл по указанному пути существует(dir проверяет это)
        If sAttachment <> "" Then
            If Dir(sAttachment, 16) <> "" Then
                .Attachments.Add sAttachment 'просто вложение
                'чтобы отправить активную книгу вместо sAttachment указать ActiveWorkbook.FullName
            End If
        End If
        .Send 'Display,Send если необходимо просмотреть сообщение, а не отправлять без просмотра
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub
