Attribute VB_Name = "DB関係"
Option Explicit

'変数
'office2007用
Private Const strProvider = "Provider=Microsoft.ACE.OLEDB.12.0;"
'DBファイルパス
Private strDbPath As String

'関数
'DBファイルパス設定
Public Sub SetDbPath(ByVal buf As String)
    strDbPath = "Data Source=" & buf & ";"
End Sub

'クエリ発行
'登録, 更新
Public Sub QueryExecute(ByVal strSQL As String)
    Dim connect As New ADODB.Connection
    
    'エラー制御ON(DEBUG用)
    On Error GoTo SQLERROR
    
    'DB接続
    connect.Open strProvider & strDbPath
    
    'クエリ発行
    Call connect.Execute(strSQL)
    
    'エラー制御OFF(DEBUG用)
    On Error GoTo 0
    
    'terminate
    connect.Close
    Set connect = Nothing
    
    Exit Sub
    
SQLERROR:
    Call ErrorInfo(connect, strSQL)
End Sub

'検索
Public Sub QuerySelect(ByVal strSQL As String, ByVal rngTarget As Range)
    Dim connect As New ADODB.Connection
    Dim recordset As New ADODB.recordset
    
    'エラー制御ON(DEBUG用)
    On Error GoTo SQLERROR
    
    'DB接続
    connect.Open strProvider & strDbPath
    
    'クエリ発行
    recordset.Open strSQL, connect, adLockReadOnly
    
    'データ格納
    rngTarget.CopyFromRecordset recordset
    
    'エラー制御OFF(DEBUG用)
    On Error GoTo 0
    
    'terminate
    recordset.Close
    Set recordset = Nothing
    connect.Close
    Set connect = Nothing
    
    Exit Sub
    
SQLERROR:
    Call ErrorInfo(connect, strSQL)
End Sub


'エラー詳細の表示
Private Sub ErrorInfo(ByVal connect As ADODB.Connection, ByVal strSQL As String)
    
    Debug.Print "=== ERROR ==="
    
    'イミディエイトウィンドウへエラー詳細を出力
    With connect.Errors.Item(0)
        Debug.Print " Description=" & .Description
        Debug.Print " HelpContext=" & .HelpContext
        Debug.Print " HelpFile=" & .HelpFile
        Debug.Print " NativeError=" & .NativeError
        Debug.Print " Number=" & .Number
        Debug.Print " Source=" & .Source
        Debug.Print " SQLState=" & .SqlState
    End With
    
    Debug.Print " SQL=" & strSQL
    
    Debug.Print "=== ERROR ==="
End Sub
