Attribute VB_Name = "編集"
Option Explicit

'編集処理
Public Sub ValidatData(ByVal Target As Range)
    Dim rep As Object
    Dim i As Long
    
    Application.EnableEvents = False
    Set rep = CreateObject("VBScript.RegExp")
    
    '入力規則に反する場合赤く網掛けを行なって注意を促す
    For i = 1 To Target.Count: Do
        Target(i).Interior.ColorIndex = xlNone
        If Target(i).text = "" Then Exit Do
        
        'ヘッダから処理を選択
        Select Case Cells(1, Target(i).Column)
        Case "車種ｺｰﾄﾞ"
            '半角英字1桁+半角数字2桁
            rep.Pattern = "[A-Za-z][0-9]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "部品ｺｰﾄﾞ"
            '半角数字4桁
            rep.Pattern = "^[0-9]{4}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "部品ｺｰﾄﾞ枝番"
            '半角数字4桁
            rep.Pattern = "^[0-9]{4}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "修理方法ｺｰﾄﾞ"
            '半角英字1-3桁
            rep.Pattern = "^[A-Za-z]{1,3}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "年式ｺｰﾄﾞ"
            '半角数字1桁
            rep.Pattern = "^[0-9]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "ﾎﾞﾃﾞｨ形状ｺｰﾄﾞ"
            '半角数字2桁
            rep.Pattern = "^[0-9]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "ｸﾞﾚｰﾄﾞｺｰﾄﾞ"
            '半角英字1-5桁
            rep.Pattern = "^[A-Za-z]{1,5}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "VA"
            '半角英字1-2桁
            rep.Pattern = "^[A-Za-z]{1,2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "ﾃﾞｰﾀ抽出番号"
            '半角英数字1-18桁
            rep.Pattern = "^[\-0-9A-Za-z]{1,18}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "開始"
            '半角数字1-12桁
            rep.Pattern = "^[0-9]{1,12}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "終了"
            '半角数字1-12桁
            rep.Pattern = "^[0-9]{1,12}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "左右"
            '"L" or "R"
            rep.Pattern = "^[LlRr]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "前後"
            '"F" or "R"
            rep.Pattern = "^[FfRr]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "部品名称"
            '半角英数カナ1-18桁
            rep.Pattern = "^[0-9A-Za-z\uFF61-\uFF9F]{1,18}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "部品番号"
            '半角英数字1-17桁
            rep.Pattern = "^[\-0-9A-Za-z]{1,17}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "構成ｸﾞﾙｰﾌﾟ"
            '半角英数字2桁
            rep.Pattern = "^[0-9A-Za-z]{2}$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case "関連作業ｺｰﾄﾞ"
            '10桁
            If LenMbcs(Target(i).text) > 10 Then Call SetInputErr(Target(i)): Exit Do
            
        Case "備考"
            '20桁
            If LenMbcs(Target(i).text) > 20 Then Call SetInputErr(Target(i)): Exit Do
            
        Case "工賃区分"
            '1
            If Target(i).text <> "1" Then Call SetInputErr(Target(i)): Exit Do
            
        Case "ｸﾞﾙｰﾌﾟｺｰﾄﾞ"
            '* or #
            rep.Pattern = "^[\*#]$"
            rep.Global = True
            If Not rep.Test(Target(i).text) Then Call SetInputErr(Target(i)): Exit Do
            
        Case ""
            'ヘッダが空白以降はデータがない
            Exit Do
        End Select
        
        '大文字変換
        Target(i).Value = UCase(Target(i).text)
    Loop Until 1: Next
    
    Set rep = Nothing
    Application.EnableEvents = True
End Sub

Private Sub SetInputErr(ByVal Target As Range, Optional desc As String = "")
    Target.Interior.ColorIndex = 38
    If desc <> "" Then MsgBox desc
End Sub

