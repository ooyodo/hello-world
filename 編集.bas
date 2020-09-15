Attribute VB_Name = "�ҏW"
Option Explicit

'�ҏW����
Public Sub ValidatData(ByVal Target As Range)
    Dim rep As Object
    Dim i As Long
    
    Application.EnableEvents = False
    Set rep = CreateObject("VBScript.RegExp")
    
    '���͋K���ɔ�����ꍇ�Ԃ��Ԋ|�����s�Ȃ��Ē��ӂ𑣂�
    For i = 1 To Target.Count: Do
        Target(i).Interior.ColorIndex = xlNone
        If Target(i).text = "" Then Exit Do
        
        '�w�b�_���珈����I��
        Select Case Cells(1, Target(i).Column)
        Case "�Ԏ���"
            '���p�p��1��+���p����2��
            rep.Pattern = "[A-Za-z][0-9]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "���i����"
            '���p����4��
            rep.Pattern = "^[0-9]{4}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "���i���ގ}��"
            '���p����4��
            rep.Pattern = "^[0-9]{4}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�C�����@����"
            '���p�p��1-3��
            rep.Pattern = "^[A-Za-z]{1,3}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�N������"
            '���p����1��
            rep.Pattern = "^[0-9]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "���ި�`����"
            '���p����2��
            rep.Pattern = "^[0-9]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "��ڰ�޺���"
            '���p�p��1-5��
            rep.Pattern = "^[A-Za-z]{1,5}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "VA"
            '���p�p��1-2��
            rep.Pattern = "^[A-Za-z]{1,2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�ް����o�ԍ�"
            '���p�p����1-18��
            rep.Pattern = "^[\-0-9A-Za-z]{1,18}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�J�n"
            '���p����1-12��
            rep.Pattern = "^[0-9]{1,12}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�I��"
            '���p����1-12��
            rep.Pattern = "^[0-9]{1,12}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "���E"
            '"L" or "R"
            rep.Pattern = "^[LlRr]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�O��"
            '"F" or "R"
            rep.Pattern = "^[FfRr]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "���i����"
            '���p�p���J�i1-18��
            rep.Pattern = "^[0-9A-Za-z\uFF61-\uFF9F]{1,18}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "���i�ԍ�"
            '���p�p����1-17��
            rep.Pattern = "^[\-0-9A-Za-z]{1,17}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�\����ٰ��"
            '���p�p����2��
            rep.Pattern = "^[0-9A-Za-z]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�֘A��ƺ���"
            '10��
            If LenMbcs(Target(i).text) > 10 Then Call SetInputErr(Target(i)): Exit Do
            
        Case "���l"
            '20��
            If LenMbcs(Target(i).text) > 20 Then Call SetInputErr(Target(i)): Exit Do
            
        Case "�H���敪"
            '1
            If Target(i).text <> "1" Then Call SetInputErr(Target(i)): Exit Do
            
        Case "��ٰ�ߺ���"
            '* or #
            rep.Pattern = "^[\*#]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case ""
            '�w�b�_���󔒈ȍ~�̓f�[�^���Ȃ�
            Exit Do
        End Select
        
        '�啶���ϊ�
        Target(i).Value = UCase(Target(i).text)
    Loop Until 1: Next
    
    Set rep = Nothing
    Application.EnableEvents = True
End Sub

Private Sub SetInputErr(ByVal Target As Range, Optional desc As String = "")
    Target.Interior.ColorIndex = 38
    If desc <> "" Then MsgBox desc
End Sub

