VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileAppender"
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

' Dependency: IAppender, CLoggingEvent, LogLevelEnum, FileUtility

Implements IAppender

' -----------------------------------------------------------------------------
' ログファイルのローテーション種別を定義します。
' -----------------------------------------------------------------------------
Public Enum TextFileAppender_LotationKind
    None = 0
    Monthly = 2
End Enum

' このオブジェクトの名称を保持します。
Private name_ As String

' この Appender の出力先パスを保持します。
Private filePath_ As String

' ローテーション種別を保持します。
Private lotationKind_ As TextFileAppender_LotationKind

' -----------------------------------------------------------------------------
' DebugPrintAppender クラスの新しいインスタンスを初期化します。
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    name_ = "TextFileAppender"
    filePath_ = ThisWorkbook.Path & "\" & RemoveExtension(ThisWorkbook.Name) & ".log"
    lotationKind_ = Monthly
End Sub

' -----------------------------------------------------------------------------
' Appender を閉じます。
' -----------------------------------------------------------------------------
Private Sub IAppender_AppenderClose()
    ' NOP
End Sub

' -----------------------------------------------------------------------------
' CLoggingEvent を追加します。
' -----------------------------------------------------------------------------
Private Sub IAppender_DoAppend(loggingEvent As CLoggingEvent)
    
    Dim appendFilePath As String
    Dim fileno As Long
    fileno = FreeFile
    
    Select Case lotationKind_
        Case None
            appendFilePath = filePath_
        Case Monthly
            appendFilePath = RemoveExtension(filePath_) & Format(Now, "_yyyymm") & GetExtension(filePath_)
        Case Else
            appendFilePath = filePath_
    End Select
    
    On Error GoTo errHandler
        Open appendFilePath For Append As fileno
            Print #fileno, loggingEvent.TimeStamp & vbTab & _
                           GetLogLevelString(loggingEvent.LogLevel) & vbTab & _
                           loggingEvent.MachineName & vbTab & _
                           loggingEvent.ClientName & vbTab & _
                           loggingEvent.UserDomainAndName & vbTab & _
                           loggingEvent.message
        Close fileno
    On Error GoTo 0
    
    Exit Sub
    
errHandler:
    On Error Resume Next
        Close fileno
    On Error GoTo 0
    
End Sub

' -----------------------------------------------------------------------------
' このオブジェクトの名称を取得します。
' -----------------------------------------------------------------------------
Public Property Get IAppender_Name() As String
    IAppender_Name = name_
End Property
' -----------------------------------------------------------------------------
' このオブジェクトの名称を設定します。
' -----------------------------------------------------------------------------
Public Property Let Name(name__ As String)
    name_ = name__
End Property

' -----------------------------------------------------------------------------
' 出力先ファイルパスを取得します。
' -----------------------------------------------------------------------------
Public Property Get filePath() As String
    filePath = filePath_
End Property
' -----------------------------------------------------------------------------
' 出力先ファイルパスを設定します。
' -----------------------------------------------------------------------------
Public Property Let filePath(filePath__ As String)
    filePath_ = filePath__
End Property


