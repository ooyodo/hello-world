Attribute VB_Name = "�t�H���_�I���_�C�A���O�p"
Option Explicit

Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                                    (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                                    (lpBrowseInfo As BROWSEINFO) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                    (ByVal hWnd As Long, ByVal wMsg As Long, _
                                     ByVal wParam As Long, lParam As Any) As Long

Public Const WM_USER = &H400
Public Const BFFM_SETSELECTIONA = (WM_USER + 102)
Public Const BFFM_INITIALIZED = 1

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As String    ''�����t�H���_���w�肷��Ƃ���String�ɂ���
    iImage As Long
End Type


Public Function GetDirectory(Optional Msg, Optional UserPath) As String
    Dim bInfo As BROWSEINFO, pPath As String
    Dim r As Long, X As Long, pos As Integer
    With bInfo
        .pidlRoot = &H0
        If IsMissing(Msg) Then
            .lpszTitle = "�t�H���_�̑I��..."
        Else
            .lpszTitle = Msg
        End If
        .ulFlags = &H40
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)
        If IsMissing(UserPath) Then
            .lParam = CurDir & Chr(0)   ''�܂���vbNullChar
        Else
            .lParam = UserPath & Chr(0)
        End If
    End With
    X = SHBrowseForFolder(bInfo)
    pPath = Space$(512)
    r = SHGetPathFromIDList(ByVal X, ByVal pPath)
    CoTaskMemFree X
    If r Then
        pos = InStr(pPath, Chr(0))
        GetDirectory = Left(pPath, pos - 1)
    Else
        GetDirectory = ""
    End If
End Function

Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal _
                                lParam As Long, ByVal lpData As Long) As Long  ''�R�[���o�b�N�֐�
    If uMsg = BFFM_INITIALIZED Then
          SendMessage hWnd, BFFM_SETSELECTIONA, 1, ByVal lpData
    End If
End Function

Public Function FARPROC(pfn As Long) As Long    ''AddressOf���Z�q�̖߂�l��߂��֐�
    FARPROC = pfn
End Function


'�g����
'--------------------------------
'        With Application.FileDialog(msoFileDialogFolderPicker)
'            .InitialFileName = ����\���t�H���_
'            .Title = "�f�[�^�i�[�t�H���_��I�����Ă��������B"
'
'            If .Show = False Then Exit Sub
'
'            Path = .SelectedItems(1)
'        End With
'--------------------------------

