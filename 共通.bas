Attribute VB_Name = "����"
Option Explicit

'�w�肵��������́AShift_JIS�x�[�X�̃o�C�g����Ԃ�
Public Function LenMbcs(ByVal str As String)
    LenMbcs = LenB(StrConv(str, vbFromUnicode))
End Function

'��������́AShift_JIS�x�[�X�̃o�C�g���Ŏw�肵���ʒu�̕�����Ԃ�
Public Function MidMbcs(ByVal str As String, start, length)
    MidMbcs = StrConv(MidB(StrConv(str, vbFromUnicode), start, length), vbUnicode)
End Function

Public Sub �`���~()
    'Application���\�b�h��ύX
    With Application
        'Change�C�x���g���~�߂�
        .EnableEvents = False
        '��ʕ`����~�߂�
        .ScreenUpdating = False
        '�X�e�[�^�X�o�[�X�V
        .StatusBar = "��ʍX�V���ł��c"
    End With
End Sub

Public Sub �`��J�n()
    'Application���\�b�h�����ɖ߂�
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub


