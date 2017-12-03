VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' -----------------------------------------------------------------------------
' 開発用のデバッグメッセージを出力します。
' -----------------------------------------------------------------------------
Public Sub LogDebug(message As String)
End Sub

' -----------------------------------------------------------------------------
' 操作ログなどの情報を出力します。
' -----------------------------------------------------------------------------
Public Sub LogInfo(message As String)
End Sub

' -----------------------------------------------------------------------------
' 障害ではない注意警告を出力します。
' -----------------------------------------------------------------------------
Public Sub LogWarn(message As String)
End Sub

' -----------------------------------------------------------------------------
' システム停止はしないが、問題となる障害を出力します。
' -----------------------------------------------------------------------------
Public Sub LogError(message As String)
End Sub

' -----------------------------------------------------------------------------
' システム停止するような致命的な障害を出力します。
' -----------------------------------------------------------------------------
Public Sub LogFatal(message As String)
End Sub

' -----------------------------------------------------------------------------
' 開発用のデバッグメッセージの出力が有効かどうかを取得します。
' -----------------------------------------------------------------------------
Public Property Get IsDebugEnabled() As Boolean
End Property
' -----------------------------------------------------------------------------
' 開発用のデバッグメッセージの出力が有効かどうかを設定します。
' -----------------------------------------------------------------------------
Public Property Let IsDebugEnabled(isDebugEnabled_ As Boolean)
End Property

' -----------------------------------------------------------------------------
' 操作ログなどの情報の出力が有効かどうかを取得します。
' -----------------------------------------------------------------------------
Public Property Get IsInfoEnabled() As Boolean
End Property
' -----------------------------------------------------------------------------
' 操作ログなどの情報の出力が有効かどうかを設定します。
' -----------------------------------------------------------------------------
Public Property Let IsInfoEnabled(isInfoEnabled_ As Boolean)
End Property

' -----------------------------------------------------------------------------
' 障害ではない注意警告の出力が有効かどうかを取得します。
' -----------------------------------------------------------------------------
Public Property Get IsWarnEnabled() As Boolean
End Property
' -----------------------------------------------------------------------------
' 障害ではない注意警告の出力が有効かどうかを設定します。
' -----------------------------------------------------------------------------
Public Property Let IsWarnEnabled(isWarnEnabled_ As Boolean)
End Property

' -----------------------------------------------------------------------------
' システム停止はしないが、問題となる障害の出力が有効かどうかを取得します。
' -----------------------------------------------------------------------------
Public Property Get IsErrorEnabled() As Boolean
End Property
' -----------------------------------------------------------------------------
' システム停止はしないが、問題となる障害の出力が有効かどうかを設定します。
' -----------------------------------------------------------------------------
Public Property Let IsErrorEnabled(isErrorEnabled_ As Boolean)
End Property

' -----------------------------------------------------------------------------
' システム停止するような致命的な障害の出力が有効かどうかを取得します。
' -----------------------------------------------------------------------------
Public Property Get IsFatalEnabled() As Boolean
End Property
' -----------------------------------------------------------------------------
' システム停止するような致命的な障害の出力が有効かどうかを設定します。
' -----------------------------------------------------------------------------
Public Property Let IsFatalEnabled(isFatalEnabled_ As Boolean)
End Property

' -----------------------------------------------------------------------------
' 登録されている IAppender をクリアします。
' -----------------------------------------------------------------------------
Public Sub ClearAppenders()
End Sub

' -----------------------------------------------------------------------------
' IAppender を登録します。
' -----------------------------------------------------------------------------
Public Sub RegistAppender(appender As IAppender)
End Sub

' -----------------------------------------------------------------------------
' 登録されている子をクリアします。
' -----------------------------------------------------------------------------
Public Sub ClearChildren()
End Sub

' -----------------------------------------------------------------------------
' 子を登録します。
' -----------------------------------------------------------------------------
Public Sub RegistChild(child As ILog)
End Sub
