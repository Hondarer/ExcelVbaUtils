Attribute VB_Name = "LogManager"
Option Explicit
' -----------------------------------------------------------------------------
' ExcelVbaUtils
' https://github.com/Hondarer/ExcelVbaUtils
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

' Dependency: None

#Const ENABLE_TEST_METHODS = 1

' フォールバック用の Logger を保持します。
Private fallbackLogger_ As ILog

' デフォルトの Logger を保持します。
Private defaultLogger_ As ILog

' カテゴリ別の Logger を保持します。
Private loggers As Object

' -----------------------------------------------------------------------------
' デフォルトの Logger を取得します｡
' <OUT> ILog デフォルトの Logger。
' -----------------------------------------------------------------------------
Public Function GetDefaultLogger() As ILog
    
    If defaultLogger_ Is Nothing Then
        If fallbackLogger_ Is Nothing Then
            Set fallbackLogger_ = New FallbackLogger
        Else
            Set GetDefaultLogger = fallbackLogger_
        End If
        Exit Function
    End If
    
    Set GetDefaultLogger = defaultLogger_

End Function

' -----------------------------------------------------------------------------
' デフォルトの Logger を設定します｡
' -----------------------------------------------------------------------------
Public Sub SetDefaultLogger(defaultLogger__ As ILog)
    Set defaultLogger_ = defaultLogger__
End Sub

' -----------------------------------------------------------------------------
' カテゴリ別の Logger のイニシャルチェックを行います。
' -----------------------------------------------------------------------------
Private Sub InitLoggers()
    If loggers Is Nothing Then
        Set loggers = CreateObject("Scripting.Dictionary")
    End If
End Sub

' -----------------------------------------------------------------------------
' カテゴリ別の Logger を初期化します。
' -----------------------------------------------------------------------------
Public Sub ClearLoggers()
    Call InitLoggers
    Call loggers.RemoveAll
End Sub

' -----------------------------------------------------------------------------
' カテゴリ別の Logger を登録します。
' -----------------------------------------------------------------------------
Public Sub RegistLogger(key As String, logger As ILog)
    Call InitLoggers
    Call loggers.Add(key, logger)
End Sub

' -----------------------------------------------------------------------------
' カテゴリ別の Logger を取り出します。
' -----------------------------------------------------------------------------
Public Function GetLogger(key As String) As ILog
    Call InitLoggers
    If Not loggers.Exists(key) Then
        Call GetDefaultLogger.LogFatal("[GetLogger] キー '" & key & "' が見つかりません。")
        Set GetLogger = GetDefaultLogger
        Exit Function
    End If
    Set GetLogger = loggers.Item(key)
End Function

#If ENABLE_TEST_METHODS = 1 Then

Public Function Test()
    Call GetDefaultLogger.ClearAppenders
    Call GetDefaultLogger.RegistAppender(New DebugPrintAppender)
    Call GetDefaultLogger.RegistAppender(New OutputDebugStringAppender)
    Call GetDefaultLogger.RegistAppender(New TextFileAppender)
    GetDefaultLogger.LogDebug "test"
    GetDefaultLogger.LogFatal "test"
End Function

Public Function FallbackTest()
    GetDefaultLogger.LogDebug "test1"
    GetLogger("Dummy").LogDebug "test2"
End Function

#End If
