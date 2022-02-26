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
    '������� ������������ � Outlook, ���� �� ��� ������
    Set objOutlookApp = GetObject(, "Outlook.Application")
    Err.Clear 'Outlook ������, ������� ������
    If objOutlookApp Is Nothing Then
        Set objOutlookApp = CreateObject("Outlook.Application")
    End If
    '����������������� ������, ���� � Outlook ��������� ������� ������� � ����� ������������ � ����������(������ ���� Outlook �������)
    '   [���������]: Session.Logon "��� �������","������",[���������� ���� ������ �������], [��������� � ����� ������]
    'objOutlookApp.Session.Logon "profile","1234",False, True
    Set objMail = objOutlookApp.CreateItem(0)   '������� ����� ���������
    '���� �� ���������� ������� ���������� ��� ��������� ��������� - �������
    If Err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
 
    sTo = address    '����(����� �������� ��������� �� ������ - sTo = Range("A1").Value)
    sSubject = Cells(2, 8)   '���� ������(����� �������� ��������� �� ������ - sSubject = Range("A2").Value)
    sBody = Cells(2, 8)    '����� ������(����� �������� ��������� �� ������ - sBody = Range("A3").Value)
    If Cells(2, 7).Value <> "" Then
        sAttachment = ThisWorkbook.Path & "\" & Cells(2, 7).Value   '��������(������ ���� � �����. ����� �������� ��������� �� ������ - sAttachment = Range("A4").Value)
    Else
        sAttachment = ""
    End If
 
    '������� ���������
    With objMail
        .to = sTo '����� ����������
        .CC = "" '����� ��� �����
        .BCC = "" '����� ��� ������� �����
        .Subject = sSubject '���� ���������
        .Body = sBody '����� ���������
        '.HTMLBody = sBody '���� ��������� ��������������� ����� ���������(��������� ������, ���� ������ � �.�.)
        '��������� ��������, ���� ���� �� ���������� ���� ����������(dir ��������� ���)
        If sAttachment <> "" Then
            If Dir(sAttachment, 16) <> "" Then
                .Attachments.Add sAttachment '������ ��������
                '����� ��������� �������� ����� ������ sAttachment ������� ActiveWorkbook.FullName
            End If
        End If
        .Send 'Display,Send ���� ���������� ����������� ���������, � �� ���������� ��� ���������
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
End Sub
