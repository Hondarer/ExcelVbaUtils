VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryFunction"
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
' データベースからの値取得サンプル
'
' VBA には継承の概念が無いため、内包して等価の処理を行う。
'
' -----------------------------------------------------------------------------

Implements IProcessControlCallbacks

Dim base As QueryFunctionBase

Dim from_ As String
Dim where_ As String
Dim orderby_ As String

' -----------------------------------------------------------------------------
' ユーザー名を取得します。
' -----------------------------------------------------------------------------
Property Get Username() As String
    Username = base.Username
End Property
' -----------------------------------------------------------------------------
' ユーザー名を設定します。
' -----------------------------------------------------------------------------
Property Let Username(username__ As String)
    base.Username = username__
End Property

' -----------------------------------------------------------------------------
' パスワードを設定します。
' -----------------------------------------------------------------------------
Property Let Password(password__ As String)
    base.Password = password__
End Property

' -----------------------------------------------------------------------------
' ネットサービス名を取得します。
' -----------------------------------------------------------------------------
Property Get Tlsname() As String
    Tlsname = base.Tlsname
End Property
' -----------------------------------------------------------------------------
' ネットサービス名を設定します。
' -----------------------------------------------------------------------------
Property Let Tlsname(tlsname__ As String)
    base.Tlsname = tlsname__
End Property

' -----------------------------------------------------------------------------
' from 句を取得します。
' -----------------------------------------------------------------------------
Property Get From() As String
    From = from_
End Property
' -----------------------------------------------------------------------------
' from 句を設定します。
' -----------------------------------------------------------------------------
Property Let From(from__ As String)
    from_ = from__
End Property

' -----------------------------------------------------------------------------
' where 句を取得します。
' -----------------------------------------------------------------------------
Property Get Where() As String
    Where = where_
End Property
' -----------------------------------------------------------------------------
' where 句を設定します。
' -----------------------------------------------------------------------------
Property Let Where(where__ As String)
    where_ = where__
End Property

' -----------------------------------------------------------------------------
' order by 句を取得します。
' -----------------------------------------------------------------------------
Property Get Orderby() As String
    Orderby = orderby_
End Property
' -----------------------------------------------------------------------------
' order by 句を設定します。
' -----------------------------------------------------------------------------
Property Let Orderby(orderby__ As String)
    orderby_ = orderby__
End Property

' -----------------------------------------------------------------------------
' 列数を取得します。
' -----------------------------------------------------------------------------
Property Get ColumnCount() As Long
    ColumnCount = base.ColumnCount()
End Property

' -----------------------------------------------------------------------------
' 問い合わせタイムアウト[ms]を取得します。
' -----------------------------------------------------------------------------
Property Get QueryTimeoutMilliSeconds() As Long
    QueryTimeoutMilliSeconds = base.QueryTimeoutMilliSeconds
End Property
' -----------------------------------------------------------------------------
' 問い合わせタイムアウト[ms]を設定します。
' -----------------------------------------------------------------------------
Property Let QueryTimeoutMilliSeconds(queryTimeout__ As Long)
    base.QueryTimeoutMilliSeconds = queryTimeout__
End Property

' -----------------------------------------------------------------------------
' 新しいインスタンスを初期化します。
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set base = New QueryFunctionBase
End Sub
' -----------------------------------------------------------------------------
' 列を追加します。
' -----------------------------------------------------------------------------
Public Sub AddColumns(columnName As String)
    Call base.AddColumns(columnName)
End Sub

' -----------------------------------------------------------------------------
' 列を初期化します。
' -----------------------------------------------------------------------------
Public Sub ClearColumns()
    Call base.ClearColumns
End Sub

' -----------------------------------------------------------------------------
' 列名を取得します。
' -----------------------------------------------------------------------------
Public Function GetColumnName(column As Long) As String
    GetColumnName = base.GetColumnName(column)
End Function

' -----------------------------------------------------------------------------
' 機能を実行します。
' -----------------------------------------------------------------------------
Public Function Execute() As Long
    Execute = base.ExecuteWithCallbackObject(Me)
End Function

' -----------------------------------------------------------------------------
' 値を返します。
' -----------------------------------------------------------------------------
Public Function GetResult(record As Long, column As Long) As Variant
    GetResult = base.GetResult(record, column)
End Function

' -----------------------------------------------------------------------------
' 行数を返します。列名を示す行は含まれません。
' -----------------------------------------------------------------------------
Public Function GetRecordsCount() As Long
    GetRecordsCount = base.GetRecordsCount
End Function

' -----------------------------------------------------------------------------
' 列数を返します。
' -----------------------------------------------------------------------------
Public Function GetColumnsCount() As Long
    GetColumnsCount = base.GetColumnsCount
End Function

' -----------------------------------------------------------------------------
' SQL 文字列を生成します。
' -----------------------------------------------------------------------------
Private Function GetSqlSelectString() As String
    
    Dim column As Long
    
    GetSqlSelectString = "select "
    
    For column = 0 To base.GetColumnsCount - 1
    
        GetSqlSelectString = GetSqlSelectString & base.GetColumnName(column)
        
        If column < (base.GetColumnsCount - 1) Then
            GetSqlSelectString = GetSqlSelectString & " || ',' || "
        End If
        
    Next

    GetSqlSelectString = GetSqlSelectString & " from " & from_
        
    If where_ <> "" Then
        GetSqlSelectString = GetSqlSelectString & " where " & where_
    End If

    If orderby_ <> "" Then
        GetSqlSelectString = GetSqlSelectString & " order by " & orderby_
    Else
        GetSqlSelectString = GetSqlSelectString & " order by ROWID"
    End If
    
    GetSqlSelectString = GetSqlSelectString & ";"
    
End Function

' -----------------------------------------------------------------------------
' 対象のプロセスと接続された際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnConnected(sender As ProcessController)
    
    base.sql = GetSqlSelectString()
    
    Call base.IProcessControlCallbacks_.OnConnected(sender)

End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスとの通信が長い時間行われなかった際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnDeepIdle(sender As ProcessController)
    Call base.IProcessControlCallbacks_.OnDeepIdle(sender)
End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスが終了した際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnDisconnected(sender As ProcessController)
    Call base.IProcessControlCallbacks_.OnDisconnected(sender)
End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスが短時間通信をしなかった際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnIdle(sender As ProcessController, lineString As String)
    Call base.IProcessControlCallbacks_.OnIdle(sender, lineString)
End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスの標準エラー出力に 1 行出力された際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnRecieveLineFromStderr(sender As ProcessController, recvErrString As String)
    Call base.IProcessControlCallbacks_.OnRecieveLineFromStderr(sender, recvErrString)
End Sub

' -----------------------------------------------------------------------------
' 対象のプロセスの標準出力に 1 行出力された際に発生します。
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnRecieveLineFromStdout(sender As ProcessController, recvOutString As String)
    Call base.IProcessControlCallbacks_.OnRecieveLineFromStdout(sender, recvOutString)
End Sub




