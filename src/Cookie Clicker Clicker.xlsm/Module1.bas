Attribute VB_Name = "Module1"
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public bQuit As Boolean

Sub main()
    
    On Error Resume Next
    
    If ThisWorkbook.Name = "Debug.xlsm" Then
        Stop
    End If
    
    Dim i As Long
    i = 0
    bQuit = False
    
    Application.Visible = False
    
    UserForm1.Show
    UserForm1.Label1.Caption = "������..."
    
    Dim objIE As Object
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://orteil.dashnet.org/cookieclicker/"
    
    Do While objIE.Busy = True Or objIE.readyState <> 4
        DoEvents
    Loop
    
    '�����҂� �ی�
    Sleep 5000
    
    '�N�b�L�[���N���b�N����
    Do While (True)
        
        objIE.Document.getElementById("bigCookie").Click
        
        If i < 2147483647 Then
            i = i + 1
            UserForm1.Label1.Caption = i & "��N���b�N���܂���..."
        Else
            UserForm1.Label1.Caption = "2147483647��ȏ�N���b�N���܂���..."
        End If
        
        'Quit���������ƏI��
        If bQuit = True Then
            Exit Do
        End If
        
        DoEvents
        
    Loop
    
'    objIE.Quit
'    Set objIE = Nothing
    
    Application.Quit
    
End Sub
