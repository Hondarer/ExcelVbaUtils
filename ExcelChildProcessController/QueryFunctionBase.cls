VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryFunctionBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------------------------
' ExcelChildProcessController
' https://github.com/Hondarer/ExcelChildProcessController
' -----------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2017 Tetsuo Honda
' t-honda@hondarer-soft.com
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
' 以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル
' （以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを
' 無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、
' 変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、
' およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
'
' 上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な
' 部分に記載するものとします。
'
' ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、
' 何らの保証もなく提供されます。
' ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての
' 保証も含みますが、それに限定されるものではありません。
' 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、
' ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の
' 扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わない
' ものとします。
'
' -----------------------------------------------------------------------------
'
' データベースからの値取得サンプルの基底クラス
'
' -----------------------------------------------------------------------------

Implements IProcessControlCallbacks

Private Const DEFAULT_COLUMN_DELIMITER = ","

Private Const QUERY_TIMEOUT = 60& * 1000& ' [ms]

Public controller As New ProcessController

Public isError As Boolean
Public errorDescription As String

' データ受信中かどうかを保持します。
Dim dataReceiving As Boolean

Dim records As Long
Dim columns As Long
Dim columnNames() As String

' 結果を保持します。
Dim result() As Variant

Dim username_ As String
Dim password_ As String
Dim tlsname_ As String
Dim sql_ As String
Dim columnDelimiter_ As String

Dim queryTimeout_ As Long

' -----------------------------------------------------------------------------
' このオブジェクトの IProcessControlCallbacks インターフェースを取得します。
' -----------------------------------------------------------------------------
Property Get IProcessControlCallbacks_() As IProcessControlCallbacks
    Set IProcessControlCallbacks_ = Me
End Property

' -----------------------------------------------------------------------------
' ユーザー名を取得します。
' -----------------------------------------------------------------------------
Property Get Username() As String
    Username = username_
End Property
' -----------------------------------------------------------------------------
' ユーザー名を設定します。
' -----------------------------------------------------------------------------
Property Let Username(username__ As String)
    username_ = username__
End Property

' -----------------------------------------------------------------------------
' パスワードを設定します。
' -----------------------------------------------------------------------------
Property Let Password(password__ As String)
    password_ = password__
End Property

' -----------------------------------------------------------------------------
' ネットサービス名を取得します。
' -----------------------------------------------------------------------------
Property Get Tlsname() As String
    Tlsname = tlsname_
End Property
' -----------------------------------------------------------------------------
' ネットサービス名を設定します。
' -----------------------------------------------------------------------------
Property Let Tlsname(tlsname__ As String)
    tlsname_ = tlsname__
End Property

' -----------------------------------------------------------------------------
' SQL 文字列を取得します。
' -----------------------------------------------------------------------------
Property Get sql() As String
    sql = sql_
End Property
' -----------------------------------------------------------------------------
' SQL 文字列を設定します。
' -----------------------------------------------------------------------------
Property Let sql(sql__ As String)
    sql_ = sql__
End Property

' -----------------------------------------------------------------------------
' 列区切り文字列を取得します。
' -----------------------------------------------------------------------------
Property Get ColumnDelimiter() As String
    ColumnDelimiter = columnDelimiter_
End Property
' -----------------------------------------------------------------------------
' 列区切り文字列を設定します。
' -----------------------------------------------------------------------------
Property Let ColumnDelimiter(columnDelimiter__ As String)
    columnDelimiter_ = columnDelimiter__
End Property

' -----------------------------------------------------------------------------
' 列数を取得します。
' -----------------------------------------------------------------------------
Property Get ColumnCount() As Long
    ColumnCount = columns
End Property

' -----------------------------------------------------------------------------
' 問い合わせタイムアウト[ms]を取得します。
' -----------------------------------------------------------------------------
Property Get QueryTimeoutMilliSeconds() As Long
    QueryTimeoutMilliSeconds = queryTimeout_
End Property
' -----------------------------------------------------------------------------
' 問い合わせタイムアウト[ms]を設定します。
' -----------------------------------------------------------------------------
Property Let QueryTimeoutMilliSeconds(queryTimeout__ As Long)
    queryTimeout_ = queryTimeout__
End Property

' -----------------------------------------------------------------------------
' 新しいインスタンスを初期化します。
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()

    ' 列の初期化
    Call ClearColumns
    
    ' タイムアウトのデフォルト値を設定
    queryTimeout_ = QUERY_TIMEOUT
    
    ' デフォルトの列区切りを設定
    columnDelimiter_ = DEFAULT_COLUMN_DELIMITER
    
End Sub

' -----------------------------------------------------------------------------
' 列を追加します。
' -----------------------------------------------------------------------------
Public Sub AddColumns(columnName As String)

    If columns = 0 Then
        columnNames(0) = columnName
    Else
        ReDim Preserve columnNames(columns)
        columnNames(columns) = columnName
    End If
    
    columns = columns + 1

End Sub

' -----------------------------------------------------------------------------
' 列を初期化します。
' -----------------------------------------------------------------------------
Public Sub ClearColumns()

    columns = 0
    ReDim columnNames(0)

End Sub

' -----------------------------------------------------------------------------
' 列名を取得します。
' -----------------------------------------------------------------------------
Public Function GetColumnName(column As Long) As String
    GetColumnName = columnNames(column)
End Function

' -----------------------------------------------------------------------------
' 機能を実行します。
' -----------------------------------------------------------------------------
Public Function Execute() As Long
    Execute = ExecuteWithCallbackObject(Me)
End Function

' -----------------------------------------------------------------------------
' コールバック対象のオブジェクトを指定して機能を実行します。
' -----------------------------------------------------------------------------
Public Function ExecuteWithCallbackObject(callbackObject As IProcessControlCallbacks) As Long
    
    ' 大量データを参照する際には、デバッグ出力による性能低下が無視できないので
    ' デバッグ出力を無効にする
    controller.DisableIoTrace = True
    
    ' [username]/[passowrd]@[tlsname]
    ExecuteWithCallbackObject = controller.ExecuteProcess("sqlplus.exe " & username_ & "/" & password_ & "@" & tlsname_, callbackObject)
    
End Function

' -----------------------------------------------------------------------------
' 値を返します。
' -----------------------------------------------------------------------------
Public Function GetResult(record As Long, column As Long) As Variant
    GetResult = result(column, record)
End Function

' -----------------------------------------------------------------------------
' 行数を返します。列名を示す行は含まれません。
' -----------------------------------------------------------------------------
Public Function GetRecordsCount() As Long
    GetRecordsCount = records
End Function

' -----------------------------------------------------------------------------
' 列数を返します。
' -----------------------------------------------------------------------------
Public Function GetColumnsCount() As Long
    GetColumnsCount = columns
End Function

' -----------------------------------------------------------------------------
' 対象のプロセスと接続された際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnConnected(sender As ProcessController)
    
    ' 操作の予約
    sender.SendLineToStdin "set linesize 32767"
    sender.SendLineToStdin "set head off"
    sender.SendLineToStdin "set newpage 0"
    sender.SendLineToStdin "set feedback off"
    
    Debug.Print "Info SimpleQueryFunction: " & sql
    
    sender.SendLineToStdin sql
    sender.SendLineToStdin "exit"
    
    ' 列名を結果に注入する
    Dim column As Long
    ReDim Preserve result(columns - 1, 0)
    For column = 0 To columns - 1
        result(column, 0) = columnNames(column)
    Next

    ' タイムアウトを設定
    Call sender.SetDeepIdleTimeoutMilliseconds(queryTimeout_)

End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスとの通信が長い時間行われなかった際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnDeepIdle(sender As ProcessController)
    Call sender.RequestTerminate(EXIT_FAILURE)
    Call MsgBox("Oracle 操作がタイムアウトしました。", Buttons:=vbCritical, Title:="エラー")
End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスが終了した際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnDisconnected(sender As ProcessController)
End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスが短時間通信をしなかった際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnIdle(sender As ProcessController, lineString As String)
End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスの標準エラー出力に 1 行出力された際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnRecieveLineFromStderr(sender As ProcessController, recvErrString As String)
End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスの標準出力に 1 行出力された際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnRecieveLineFromStdout(sender As ProcessController, recvOutString As String)

    ' エラーが発生した場合はその内容を保持して、プログラムを終了させる
    If (recvOutString = "ERROR:") Or (Right(recvOutString, 12) = "エラーが発生しました。:") Then
        isError = True
        sender.RequestTerminate EXIT_FAILURE
    ElseIf isError = True And errorDescription = "" Then
        errorDescription = recvOutString
        Debug.Print "Eror SimpleQueryFunction: " & errorDescription
        Call MsgBox("Oracle にてエラーが発生しました。" & vbCrLf & errorDescription, Buttons:=vbCritical, Title:="エラー")
        
        ' データを削除する
        ReDim Preserve result(columns - 1, 0)
        records = 0
        
        Exit Sub
    ElseIf isError = True Then
        Exit Sub
    End If

    ' プロンプトからプロンプトまでを、データ部として扱う判定処理
    If Left(recvOutString, 5) = "SQL> " Then
        If dataReceiving = False Then
            dataReceiving = True
        Else
            dataReceiving = False
        End If
    End If
    
    ' データ受信中の処理
    If dataReceiving = True Then
    
        Dim column As Long
        Dim data() As String
            
        ' "set newpage 0" により付加される、改ページの削除
        ' &hC(Form Feed(改ページ))を検出してそこまで取り除く
        ' 改ページはデータの開始判定に利用しているため、意識的に付加している
        If InStr(recvOutString, Chr(&HC)) > 0 Then
            recvOutString = Mid(recvOutString, InStr(recvOutString, Chr(&HC)) + 1)
        End If
        
        data = Split(recvOutString, columnDelimiter_)
        
        ' エラー文字列などではなく、データ部であること
        ' (区切られた要素の個数が予想される列数と一致していること)
        If UBound(data) = (columns - 1) Then
        
'           Debug.Print "> " & data(0) & " / " & data(1)

            records = records + 1
            ReDim Preserve result(columns - 1, records)
            
            For column = 0 To columns - 1
                result(column, records) = data(column)
            Next
        
        End If
        
    End If
    
End Sub



