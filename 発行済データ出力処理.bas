Attribute VB_Name = "���s�σf�[�^�o�͏���"
Option Explicit

Public Sub OutputReleaseData()
    Dim ws As Worksheet
    Dim obj As Object
    Dim buf As String
    Dim str As String
    Dim l As Long
    Dim o As Long
    Dim n As Long
    
    '���s�N������
    buf = Application.InputBox("���s�N������͂��Ă�������", , , , , , , 2)
    If buf = "false" Then Exit Sub
    
    '�t�@�C���I�[�v��
    Set ws = Sheets("�J���Ԏ���ꗗ")
    With Workbooks.Add
        o = 7
        n = 1
        Do: Do
            Set obj = ws.Rows(o)
            
            '�f�[�^���Ȃ��Ȃ�����I��
            If Application.WorksheetFunction.CountA(obj) = 0 Then Exit Sub
            
            '���s�N�������͂����l�Ɠ����̂ݏ���
            If StrConv(obj.Cells(5).Value, vbNarrow) <> StrConv(buf, vbNarrow) Then Exit Do
            
            '�f�[�^�擾
            obj.Copy Destination:=.Worksheets(1).Range(n & ":" & n)
            '���́w���s�\��N���x�w�S���ҁx�w�ύX���t�x�w�ύX���e�x�w�������s�L���x��������
            obj.Cells(5).Value = ""
            obj.Cells(6).Value = ""
            obj.Cells(7).Value = ""
            obj.Cells(8).Value = ""
            obj.Cells(9).Value = ""
            
            n = n + 1
            Application.CutCopyMode = False
            
        Loop Until 1: o = o + 1: Loop
        
    End With
End Sub
