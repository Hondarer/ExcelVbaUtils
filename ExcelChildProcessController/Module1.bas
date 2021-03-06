Attribute VB_Name = "Module1"
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

' -----------------------------------------------------------------------------
' コマンドプロンプトのサンプル 1 を実行します。
' -----------------------------------------------------------------------------
Public Sub Sample1()

    Dim func As New SampleFunction1
    
    Debug.Print "戻り値:" & func.Execute
    
End Sub

' -----------------------------------------------------------------------------
' コマンドプロンプトのサンプル 2 を実行します。
' -----------------------------------------------------------------------------
Public Sub Sample2()

    Dim func As New SampleFunction2
    
    Debug.Print "戻り値:" & func.Execute
    
End Sub

' -----------------------------------------------------------------------------
' コマンドプロンプトのサンプル 3 を実行します。
' -----------------------------------------------------------------------------
Public Sub Sample3()

    Dim controller As New ProcessController
    
    ' コールバックが必要ない場合の直接起動
    Debug.Print "戻り値:" & controller.ExecuteProcess("cmd.exe /k ver & dir & fff & exit", Nothing)
    
End Sub

' -----------------------------------------------------------------------------
' コマンドプロンプトのサンプル 4 を実行します。
' -----------------------------------------------------------------------------
Public Sub Sample4()

    Dim controller As New ProcessController
    
    ' コンソール ウインドウでの制御を行う
    ' 残置した標準入出力のコールバックは実施されない
    controller.ShowConsoleWindow = True
    controller.LeaveStdin = True
    controller.LeaveStdout = True
    controller.LeaveStderr = True
    
    ' タイムアウト監視の無効化
    Call controller.DisableDeepIdleTimeoutMilliseconds
    
    Debug.Print "戻り値:" & controller.ExecuteProcess("cmd.exe /k ver & dir & fff & pause & exit", Nothing)
    
End Sub

' -----------------------------------------------------------------------------
' データベースからの値取得サンプルを実行します。
' -----------------------------------------------------------------------------
Public Sub QueryFunctionSample()

    Dim func As New QueryFunction

    ' 設定値
    ' シートから取るなり定数で与えるなりお好きに
    func.Username = "hr"
    func.Password = "tiger"
    func.Tlsname = "XE"

    ' 取得する列
    '    Call func.AddColumns("ROWID")
    Call func.AddColumns("REGION_ID")
    Call func.AddColumns("REGION_NAME")

    ' from 句
    ' テーブル関数の場合は、引数まで与える
    func.From = "REGIONS"

    '    ' where 句
    '    func.Where = "REGION_ID = 4"

    ' order by 句
    ' asc, desc が指定必要な場合は、ここで与える
    ' 明示的に指定しない場合は、フレームワークにより ROWID で order される。
    ' テーブル関数など ROWID が存在しない場合は必ず指定すること。
    func.Orderby = "REGION_ID"

    '    ' タイムアウトの設定
    '    ' 明示的に指定しない場合は、フレームワークにより
    '    ' デフォルトのタイムアウト(60 秒)が指定される。単位は[ms]。
    '    func.QueryTimeoutMilliSeconds = 60& * 1000&

    ' 問い合わせ実行
    Debug.Print "戻り値:" & func.Execute

    ' セルに書き出すには以下のようにする
    ' ヘッダ行があるため、行の数はデータ + 1 となる、注意 !

    Dim record As Long
    Dim column As Long

    For record = 0 To func.GetRecordsCount
        For column = 0 To func.GetColumnsCount - 1
            ThisWorkbook.Worksheets("Sheet1").Cells(record + 1, column + 1).Value = func.GetResult(record, column)
        Next
    Next

End Sub
