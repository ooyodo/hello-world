Attribute VB_Name = "共通"
Option Explicit

'指定した文字列の、Shift_JISベースのバイト数を返す
Public Function LenMbcs(ByVal str As String)
    LenMbcs = LenB(StrConv(str, vbFromUnicode))
End Function

'文字列内の、Shift_JISベースのバイト数で指定した位置の文字を返す
Public Function MidMbcs(ByVal str As String, start, length)
    MidMbcs = StrConv(MidB(StrConv(str, vbFromUnicode), start, length), vbUnicode)
End Function

Public Sub 描画停止()
    'Applicationメソッドを変更
    With Application
        'Changeイベントを止める
        .EnableEvents = False
        '画面描画を止める
        .ScreenUpdating = False
        'ステータスバー更新
        .StatusBar = "画面更新中です…"
    End With
End Sub

Public Sub 描画開始()
    'Applicationメソッドを元に戻す
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub


