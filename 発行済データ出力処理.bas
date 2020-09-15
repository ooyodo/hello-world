Attribute VB_Name = "発行済データ出力処理"
Option Explicit

Public Sub OutputReleaseData()
    Dim ws As Worksheet
    Dim obj As Object
    Dim buf As String
    Dim str As String
    Dim l As Long
    Dim o As Long
    Dim n As Long
    
    '発行年月入力
    buf = Application.InputBox("発行年月を入力してください", , , , , , , 2)
    If buf = "false" Then Exit Sub
    
    'ファイルオープン
    Set ws = Sheets("開発車種情報一覧")
    With Workbooks.Add
        o = 7
        n = 1
        Do: Do
            Set obj = ws.Rows(o)
            
            'データがなくなったら終了
            If Application.WorksheetFunction.CountA(obj) = 0 Then Exit Sub
            
            '発行年月が入力した値と同等のみ処理
            If StrConv(obj.Cells(5).Value, vbNarrow) <> StrConv(buf, vbNarrow) Then Exit Do
            
            'データ取得
            obj.Copy Destination:=.Worksheets(1).Range(n & ":" & n)
            '元の『発行予定年月』『担当者』『変更日付』『変更内容』『資料発行有無』を初期化
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
