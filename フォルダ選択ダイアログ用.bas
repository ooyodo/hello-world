Attribute VB_Name = "フォルダ選択ダイアログ用"
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
    lParam As String    ''初期フォルダを指定するときはStringにする
    iImage As Long
End Type


Public Function GetDirectory(Optional Msg, Optional UserPath) As String
    Dim bInfo As BROWSEINFO, pPath As String
    Dim r As Long, X As Long, pos As Integer
    With bInfo
        .pidlRoot = &H0
        If IsMissing(Msg) Then
            .lpszTitle = "フォルダの選択..."
        Else
            .lpszTitle = Msg
        End If
        .ulFlags = &H40
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)
        If IsMissing(UserPath) Then
            .lParam = CurDir & Chr(0)   ''またはvbNullChar
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
                                lParam As Long, ByVal lpData As Long) As Long  ''コールバック関数
    If uMsg = BFFM_INITIALIZED Then
          SendMessage hWnd, BFFM_SETSELECTIONA, 1, ByVal lpData
    End If
End Function

Public Function FARPROC(pfn As Long) As Long    ''AddressOf演算子の戻り値を戻す関数
    FARPROC = pfn
End Function


'使い方
'--------------------------------
'        With Application.FileDialog(msoFileDialogFolderPicker)
'            .InitialFileName = 初回表示フォルダ
'            .Title = "データ格納フォルダを選択してください。"
'
'            If .Show = False Then Exit Sub
'
'            Path = .SelectedItems(1)
'        End With
'--------------------------------

