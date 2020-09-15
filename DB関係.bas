Attribute VB_Name = "DB�֌W"
Option Explicit

'�ϐ�
'office2007�p
Private Const strProvider = "Provider=Microsoft.ACE.OLEDB.12.0;"
'DB�t�@�C���p�X
Private strDbPath As String

'�֐�
'DB�t�@�C���p�X�ݒ�
Public Sub SetDbPath(ByVal buf As String)
    strDbPath = "Data Source=" & buf & ";"
End Sub

'�N�G�����s
'�o�^, �X�V
Public Sub QueryExecute(ByVal strSQL As String)
    Dim connect As New ADODB.Connection
    
    '�G���[����ON(DEBUG�p)
    On Error GoTo SQLERROR
    
    'DB�ڑ�
    connect.Open strProvider & strDbPath
    
    '�N�G�����s
    Call connect.Execute(strSQL)
    
    '�G���[����OFF(DEBUG�p)
    On Error GoTo 0
    
    'terminate
    connect.Close
    Set connect = Nothing
    
    Exit Sub
    
SQLERROR:
    Call ErrorInfo(connect, strSQL)
End Sub

'����
Public Sub QuerySelect(ByVal strSQL As String, ByVal rngTarget As Range)
    Dim connect As New ADODB.Connection
    Dim recordset As New ADODB.recordset
    
    '�G���[����ON(DEBUG�p)
    On Error GoTo SQLERROR
    
    'DB�ڑ�
    connect.Open strProvider & strDbPath
    
    '�N�G�����s
    recordset.Open strSQL, connect, adLockReadOnly
    
    '�f�[�^�i�[
    rngTarget.CopyFromRecordset recordset
    
    '�G���[����OFF(DEBUG�p)
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


'�G���[�ڍׂ̕\��
Private Sub ErrorInfo(ByVal connect As ADODB.Connection, ByVal strSQL As String)
    
    Debug.Print "=== ERROR ==="
    
    '�C�~�f�B�G�C�g�E�B���h�E�փG���[�ڍׂ��o��
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
