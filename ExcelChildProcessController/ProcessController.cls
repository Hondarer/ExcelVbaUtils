VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessController"
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
' Excel プロセス制御ライブラリ
'
' コンソール アプリケーションを任意の引数で起動し、
' 標準入出力を VBA から扱います。
'
' http://qiita.com/asterisk9101/items/4f0e8da6d1e2e9df8e14
' http://d.hatena.ne.jp/takoashi_net/20090305/1236239906
' http://qiita.com/shela/items/4340c5f234911db7cf9c
' https://www.moug.net/tech/exvba/0150034.html
'
' -----------------------------------------------------------------------------

Dim udtProcInfo As PROCESS_INFORMATION

Dim hOutRead As Long
Dim hOutWrite As Long
Dim hInRead As Long
Dim hInWrite As Long
Dim hErrRead As Long
Dim hErrWrite As Long

Dim ov1 As OVERLAPPED
Dim ov2 As OVERLAPPED
Dim ov3 As OVERLAPPED

Private Const POLLING_MILLISEC = 50& ' [ms]
Private Const DEFAULT_IDLE_TIMEOUTMILLISEC = 500& ' [ms]
Private Const DEFAULT_DEEPIDLE_TIMEOUTMILLISEC = 60& * 1000& ' [ms]

' パイプ通信のバッファサイズを表します。
Private Const BUFFER_LIMIT_LENGTH = 4096&

Dim idleCount As Long
Dim idleTimeoutCount As Long
Dim deepIdleCount As Long
Dim deepIdleTimeoutCount As Long

Dim recvBuffer() As Byte ' 生バッファ
Dim recvBuffer1() As Byte ' 1 次バッファ(シフト JIS のアライン未確定)
Dim recvBuffer2() As Byte ' 2 次バッファ(ラインバッファ)

Dim recvBuffer1Empty As Boolean
Dim recvBuffer2Empty As Boolean

Dim errBuffer() As Byte ' 生バッファ
Dim errBuffer1() As Byte ' 1 次バッファ(シフト JIS のアライン未確定)
Dim errBuffer2() As Byte ' 2 次バッファ(ラインバッファ)

Dim errBuffer1Empty As Boolean
Dim errBuffer2Empty As Boolean

Dim disableIoTrace_ As Boolean

Dim leaveStdin_ As Boolean
Dim leaveStdout_ As Boolean
Dim leaveStderr_ As Boolean

Dim showConsoleWindow_ As Boolean

' -----------------------------------------------------------------------------
' イミディエイト ウインドウへのデバッグ出力無効設定を取得します。
' -----------------------------------------------------------------------------
Property Get DisableIoTrace() As Boolean
    DisableIoTrace = disableIoTrace_
End Property
' -----------------------------------------------------------------------------
' イミディエイト ウインドウへのデバッグ出力無効設定を設定します。
' -----------------------------------------------------------------------------
Property Let DisableIoTrace(disableIoTrace__ As Boolean)
    disableIoTrace_ = disableIoTrace__
End Property

' -----------------------------------------------------------------------------
' 標準入力をコンソール ウインドウに残すかどうかを取得します。
' -----------------------------------------------------------------------------
Property Get LeaveStdin() As Boolean
    LeaveStdin = leaveStdin_
End Property
' -----------------------------------------------------------------------------
' 標準入力をコンソール ウインドウに残すかどうかを設定します。
' -----------------------------------------------------------------------------
Property Let LeaveStdin(leaveStdin__ As Boolean)
    leaveStdin_ = leaveStdin__
End Property

' -----------------------------------------------------------------------------
' 標準出力をコンソール ウインドウに残すかどうかを取得します。
' -----------------------------------------------------------------------------
Property Get LeaveStdout() As Boolean
    LeaveStdout = leaveStdout_
End Property
' -----------------------------------------------------------------------------
' 標準出力をコンソール ウインドウに残すかどうかを設定します。
' -----------------------------------------------------------------------------
Property Let LeaveStdout(leaveStdout__ As Boolean)
    leaveStdout_ = leaveStdout__
End Property

' -----------------------------------------------------------------------------
' 標準エラー出力をコンソール ウインドウに残すかどうかを取得します。
' -----------------------------------------------------------------------------
Property Get LeaveStderr() As Boolean
    LeaveStderr = leaveStderr_
End Property
' -----------------------------------------------------------------------------
' 標準エラー出力をコンソール ウインドウに残すかどうかを設定します。
' -----------------------------------------------------------------------------
Property Let LeaveStderr(leaveStderr__ As Boolean)
    leaveStderr_ = leaveStderr__
End Property

' -----------------------------------------------------------------------------
' コンソール ウインドウを表示するかどうかを取得します。
' -----------------------------------------------------------------------------
Property Get ShowConsoleWindow() As Boolean
    ShowConsoleWindow = showConsoleWindow_
End Property
' -----------------------------------------------------------------------------
' コンソール ウインドウを表示するかどうかを設定します。
' -----------------------------------------------------------------------------
Property Let ShowConsoleWindow(showConsoleWindow__ As Boolean)
    showConsoleWindow_ = showConsoleWindow__
End Property

' -----------------------------------------------------------------------------
' ターゲットの標準入力に文字列を与えます。
' -----------------------------------------------------------------------------
Public Function SendTextToStdin(sendstring As String) As Boolean

    Dim dwWritten As Long
    Dim writeBuffer() As Byte
    
    If (hInWrite = 0) Or (leaveStdin_ = True) Then
        SendTextToStdin = False
        Exit Function
    End If
    
    If disableIoTrace_ <> True Then
        Debug.Print "StdI " & sendstring; ' 改行しない
    End If
    
    writeBuffer = StrConv(sendstring, vbFromUnicode)
    If WriteFile(hInWrite, writeBuffer(0), UBound(writeBuffer) + 1, dwWritten, ov1) = 0 Then
        SendTextToStdin = False
        Exit Function
    End If
    
    SendTextToStdin = True

End Function

' -----------------------------------------------------------------------------
' ターゲットの標準入力に文字列と改行コードを与えます。
' -----------------------------------------------------------------------------
Public Function SendLineToStdin(sendstring As String) As Boolean

    SendLineToStdin = SendTextToStdin(sendstring & vbCrLf)
    
End Function

' -----------------------------------------------------------------------------
' アイドル検出時間をミリ秒で設定します。
' -----------------------------------------------------------------------------
Public Sub SetIdleTimeoutMilliseconds(timeout As Long)
    idleCount = 0
    idleTimeoutCount = timeout / POLLING_MILLISEC
End Sub

' -----------------------------------------------------------------------------
' アイドル検出時間をリセットします。
' -----------------------------------------------------------------------------
Public Sub ResetIdleTimeoutMilliseconds()
    Call SetIdleTimeoutMilliseconds(DEFAULT_IDLE_TIMEOUTMILLISEC)
End Sub

' -----------------------------------------------------------------------------
' アイドル検出時間を無効にします。
' -----------------------------------------------------------------------------
Public Sub DisableIdleTimeoutMilliseconds()
    Call SetIdleTimeoutMilliseconds(0)
End Sub

' -----------------------------------------------------------------------------
' ディープアイドル検出時間をミリ秒で設定します。
' -----------------------------------------------------------------------------
Public Sub SetDeepIdleTimeoutMilliseconds(timeout As Long)
    deepIdleCount = 0
    deepIdleTimeoutCount = timeout / POLLING_MILLISEC
End Sub

' -----------------------------------------------------------------------------
' ディープアイドル検出時間をリセットします。
' -----------------------------------------------------------------------------
Public Sub ResetDeepIdleTimeoutMilliseconds()
    SetDeepIdleTimeoutMilliseconds (DEFAULT_DEEPIDLE_TIMEOUTMILLISEC)
End Sub

' -----------------------------------------------------------------------------
' ディープアイドル検出時間を無効にします。
' -----------------------------------------------------------------------------
Public Sub DisableDeepIdleTimeoutMilliseconds()
    SetDeepIdleTimeoutMilliseconds (0)
End Sub

' -----------------------------------------------------------------------------
' 標準出力のラインバッファをクリアします。
' -----------------------------------------------------------------------------
Public Sub ClearStdoutLineBuffer()
    ReDim recvBuffer2(0)
    recvBuffer2Empty = True
End Sub

' -----------------------------------------------------------------------------
' ターゲットのプロセスの終了を要求します。
' -----------------------------------------------------------------------------
Public Function RequestTerminate(uExitCode As Long) As Boolean
    
    ' 使用禁止の終了コードはエラーとする
    If uExitCode = STILL_ACTIVE Then
        RequestTerminate = False
        Exit Function
    End If
    
    Debug.Print "Info ProcessController: Child process(pid=" & udtProcInfo.dwProcessId & ") request stop with code=" & uExitCode & "."
    
    If TerminateProcess(udtProcInfo.hProcess, uExitCode) = 0 Then
        Debug.Print "Warn ProcessController: Child process(pid=" & udtProcInfo.dwProcessId & ") stop failed(in general it has already terminated)."
        RequestTerminate = False
        Exit Function
    End If
    
    RequestTerminate = True
    
End Function

' -----------------------------------------------------------------------------
' プロセスを実行します。
' -----------------------------------------------------------------------------
Public Function ExecuteProcess(commandLine As String, func As IProcessControlCallbacks) As Long

    Dim sa As SECURITY_ATTRIBUTES

    Dim udtStartInfo As STARTUPINFO

    Dim procExitCode As Long
    
    Dim dwRead As Long
    
    Dim strBuffer As String
    
    ReDim recvBuffer1(0)
    ReDim recvBuffer2(0)
    
    recvBuffer1Empty = True
    recvBuffer2Empty = True
    
    ReDim errBuffer1(0)
    ReDim errBuffer2(0)
    
    errBuffer1Empty = True
    errBuffer2Empty = True
    
    Dim nextRecv1Index As Long
    Dim sjisFirstByte As Byte
    Dim processRecv1Index As Long

    ' セキュリティディスクリプタの設定
    Call ZeroMemory(sa, Len(sa))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&

    ' 構造体の初期化
    udtStartInfo.cb = LenB(udtStartInfo)

    If leaveStdin_ <> True Then
        ' 入力パイプの作成
        If CreatePipe(hInRead, hInWrite, sa, 0) = 0 Then
            Debug.Print "入力パイプの作成に失敗しました。エラーコード:" & Err.LastDllError
            GoTo Error
        End If
    
        ' 入力書き込みハンドルを子プロセスに継承できなくする
        If SetHandleInformation(hInWrite, HANDLE_FLAG_INHERIT, 0) = 0 Then
            Debug.Print "入力パイプの書き込みハンドルの継承設定に失敗しました。エラーコード:" & Err.LastDllError
            GoTo Error
        End If
        
        udtStartInfo.hStdInput = hInRead
    Else
        udtStartInfo.hStdInput = GetStdHandle(STD_INPUT_HANDLE)
    End If

    If leaveStdout_ <> True Then
        ' 出力パイプの作成
        If CreatePipe(hOutRead, hOutWrite, sa, 0) = 0 Then
            Debug.Print "出力パイプの作成に失敗しました。エラーコード:" & Err.LastDllError
            GoTo Error
        End If
    
        ' 出力読み込みハンドルを子プロセスに継承できなくする
        If SetHandleInformation(hOutRead, HANDLE_FLAG_INHERIT, 0) = 0 Then
            Debug.Print "出力パイプの読み込みハンドルの継承設定に失敗しました。エラーコード:" & Err.LastDllError
            GoTo Error
        End If
    
        udtStartInfo.hStdOutput = hOutWrite
    Else
        udtStartInfo.hStdOutput = GetStdHandle(STD_OUTPUT_HANDLE)
    End If

    If leaveStderr_ <> True Then
        ' エラー出力パイプの作成
        If CreatePipe(hErrRead, hErrWrite, sa, 0) = 0 Then
            Debug.Print "エラー出力パイプの作成に失敗しました。エラーコード:" & Err.LastDllError
            GoTo Error
        End If
    
        ' エラー出力読み込みハンドルを子プロセスに継承できなくする
        If SetHandleInformation(hErrRead, HANDLE_FLAG_INHERIT, 0) = 0 Then
            Debug.Print "エラー出力パイプの読み込みハンドルの継承設定に失敗しました。エラーコード:" & Err.LastDllError
            GoTo Error
        End If
    
        udtStartInfo.hStdError = hErrWrite
    Else
        udtStartInfo.hStdError = GetStdHandle(STD_ERROR_HANDLE)
    End If
    
    udtStartInfo.dwFlags = STARTF_USESTDHANDLES
    
    ' ウインドウを消去する
    If showConsoleWindow_ <> True Then
        udtStartInfo.wShowWindow = SW_HIDE
        udtStartInfo.dwFlags = udtStartInfo.dwFlags Or STARTF_USESHOWWINDOW
    End If

    ' プロセス起動
    If CreateProcess(vbNullString, commandLine, ByVal 0&, ByVal 0&, 1&, _
        NORMAL_PRIORITY_CLASS, ByVal 0&, ThisWorkbook.Path, udtStartInfo, udtProcInfo) = 0 Then
        Call MsgBox("プロセスの生成に失敗しました。エラーコード:" & Err.LastDllError & vbCrLf & commandLine, Buttons:=vbCritical, Title:="エラー")
        GoTo Error
    End If

    Debug.Print "Info ProcessController: Child process(" & commandLine & ", pid=" & udtProcInfo.dwProcessId & ") started up."
   
    ' 接続完了コールバックを呼び出す
    If Not func Is Nothing Then
        Call func.OnConnected(Me)
    End If
    
    ' メインループ
    
    Application.Cursor = xlWait
    
    Do
    
        Dim hasData As Boolean
        hasData = False
        
        Dim hasRemainData As Boolean
        hasRemainData = False
    
        If GetExitCodeProcess(udtProcInfo.hProcess, procExitCode) = 0 Then
            ' 終了コード取得異常
            GoTo Error
        End If
        
        If leaveStdout_ <> True Then
            ' 標準出力からデータが読めることを確認する
            If PeekNamedPipe(hOutRead, 0&, 0, 0&, dwRead, ByVal 0) <> 0 Then
                If (dwRead > 0) Then
                    ' データがある
                    
                    ' 未読み取りデータが多い場合は、クリップしながら処理する
                    If dwRead > BUFFER_LIMIT_LENGTH Then
                        dwRead = BUFFER_LIMIT_LENGTH
                        hasRemainData = True
                    End If
                    
                    ' 標準出力からデータを読み取る
                    ReDim recvBuffer(dwRead - 1) As Byte
                    If ReadFile(hOutRead, recvBuffer(0), UBound(recvBuffer) + 1, dwRead, ov2) <> 0 Then
                    
                        ' 受信バッファの処理
                        ' 生バッファはまず、1 次バッファに積む
                        nextRecv1Index = 0
                        If recvBuffer1Empty = True Then
                            ReDim recvBuffer1(dwRead - 1)
                            recvBuffer1Empty = False
                        Else
                            nextRecv1Index = UBound(recvBuffer1) + 1
                            ReDim Preserve recvBuffer1(UBound(recvBuffer1) + dwRead)
                        End If
                        
                        Call CopyMemory(recvBuffer1(nextRecv1Index), recvBuffer(0), dwRead)
                    
                        ' 1 次バッファから順次、2 次バッファに詰めなおす
                        ' CR および LF は、2 次バッファには詰めない
                        ' LF が出現したら、行の区切りとし、コールバックを行う
                        ' シフト JIS の上位バイトのみが余った場合は、1 次バッファに残しておく
                        
                        sjisFirstByte = 0
                        
                        If ((recvBuffer1(UBound(recvBuffer1)) >= &H80) And (recvBuffer1(UBound(recvBuffer1)) <= &H9F)) Or _
                           ((recvBuffer1(UBound(recvBuffer1)) >= &HE0) And (recvBuffer1(UBound(recvBuffer1)) <= &HFF)) Then
                           sjisFirstByte = recvBuffer1(UBound(recvBuffer1))
                           ReDim Preserve recvBuffer1(UBound(recvBuffer1) - 1)
                        End If
                        
                        For processRecv1Index = 0 To UBound(recvBuffer1)
                            If recvBuffer1(processRecv1Index) = Asc(vbCr) Then
                                ' NOP
                            ElseIf recvBuffer1(processRecv1Index) = Asc(vbLf) Then
                                ' コールバック実施
                                If recvBuffer2Empty = True Then
                                    If disableIoTrace_ <> True Then
                                        Debug.Print "StdO "
                                    End If
                                    If Not func Is Nothing Then
                                        Call func.OnRecieveLineFromStdout(Me, "")
                                    End If
                                Else
                                    strBuffer = StrConv(recvBuffer2, vbUnicode)
                                    If disableIoTrace_ <> True Then
                                        Debug.Print "StdO " & strBuffer
                                    End If
                                    If Not func Is Nothing Then
                                        Call func.OnRecieveLineFromStdout(Me, strBuffer)
                                    End If
                                End If
                                ReDim recvBuffer2(0)
                                recvBuffer2Empty = True
                            Else
                                ' 積み上げ
                                If recvBuffer2Empty = True Then
                                    recvBuffer2Empty = False
                                Else
                                    ReDim Preserve recvBuffer2(UBound(recvBuffer2) + 1)
                                End If
                                
                                recvBuffer2(UBound(recvBuffer2)) = recvBuffer1(processRecv1Index)
                            End If
                        Next
                        
                        ReDim recvBuffer1(0)
                        If sjisFirstByte <> 0 Then
                            recvBuffer1(0) = sjisFirstByte
                        Else
                            recvBuffer1Empty = True
                        End If
                    
                        idleCount = 0
                        deepIdleCount = 0
                        hasData = True
                    
                    Else
                        ' プロセスの終了コードによりループを抜けるため、エラーは無視する。
    '                    ' API 呼出異常
    '                    GoTo Error
                    End If
                End If
            Else
                ' プロセスの終了コードによりループを抜けるため、エラーは無視する。
    '            ' API 呼出異常
    '            GoTo Error
            End If
        End If

        If leaveStderr_ <> True Then
            ' 標準エラー出力からデータが読めることを確認する
            If PeekNamedPipe(hErrRead, 0&, 0, 0&, dwRead, ByVal 0) <> 0 Then
                If (dwRead > 0) Then
                    ' データがある
                    
                    ' 未読み取りデータが多い場合は、クリップしながら処理する
                    If dwRead > BUFFER_LIMIT_LENGTH Then
                        dwRead = BUFFER_LIMIT_LENGTH
                        hasRemainData = True
                    End If
                    
                    ' 標準エラー出力からデータを読み取る
                    ReDim errBuffer(dwRead - 1) As Byte
                    If ReadFile(hErrRead, errBuffer(0), UBound(errBuffer) + 1, dwRead, ov2) <> 0 Then
                    
                        ' 受信バッファの処理
                        ' 生バッファはまず、1 次バッファに積む
                        nextRecv1Index = 0
                        If errBuffer1Empty = True Then
                            ReDim errBuffer1(dwRead - 1)
                            errBuffer1Empty = False
                        Else
                            nextRecv1Index = UBound(errBuffer1) + 1
                            ReDim Preserve errBuffer1(UBound(errBuffer1) + dwRead)
                        End If
                        
                        CopyMemory errBuffer1(nextRecv1Index), errBuffer(0), dwRead
                    
                        ' 1 次バッファから順次、2 次バッファに詰めなおす
                        ' CR および LF は、2 次バッファには詰めない
                        ' LF が出現したら、行の区切りとし、コールバックを行う
                        ' シフト JIS の上位バイトのみが余った場合は、1 次バッファに残しておく
                        
                        sjisFirstByte = 0
                        
                        If ((errBuffer1(UBound(errBuffer1)) >= &H80) And (errBuffer1(UBound(errBuffer1)) <= &H9F)) Or _
                           ((errBuffer1(UBound(errBuffer1)) >= &HE0) And (errBuffer1(UBound(errBuffer1)) <= &HFF)) Then
                           sjisFirstByte = errBuffer1(UBound(errBuffer1))
                           ReDim Preserve errBuffer1(UBound(errBuffer1) - 1)
                        End If
                        
                        For processRecv1Index = 0 To UBound(errBuffer1)
                            If errBuffer1(processRecv1Index) = Asc(vbCr) Then
                                ' NOP
                            ElseIf errBuffer1(processRecv1Index) = Asc(vbLf) Then
                                ' コールバック実施
                                If errBuffer2Empty = True Then
                                    If disableIoTrace_ <> True Then
                                        Debug.Print "StdE "
                                    End If
                                    If Not func Is Nothing Then
                                        Call func.OnRecieveLineFromStderr(Me, "")
                                    End If
                                Else
                                    strBuffer = StrConv(errBuffer2, vbUnicode)
                                    If disableIoTrace_ <> True Then
                                        Debug.Print "StdE " & strBuffer
                                    End If
                                    If Not func Is Nothing Then
                                        Call func.OnRecieveLineFromStderr(Me, strBuffer)
                                    End If
                                End If
                                ReDim errBuffer2(0)
                                errBuffer2Empty = True
                            Else
                                ' 積み上げ
                                If errBuffer2Empty = True Then
                                    errBuffer2Empty = False
                                Else
                                    ReDim Preserve errBuffer2(UBound(errBuffer2) + 1)
                                End If
                                
                                errBuffer2(UBound(errBuffer2)) = errBuffer1(processRecv1Index)
                            End If
                        Next
                        
                        ReDim errBuffer1(0)
                        If sjisFirstByte <> 0 Then
                            errBuffer1(0) = sjisFirstByte
                        Else
                            errBuffer1Empty = True
                        End If
                    
                        idleCount = 0
                        deepIdleCount = 0
                        hasData = True
                    
                    Else
                        ' プロセスの終了コードによりループを抜けるため、エラーは無視する。
    '                    ' API 呼出異常
    '                    GoTo Error
                    End If
                End If
            Else
                ' プロセスの終了コードによりループを抜けるため、エラーは無視する。
    '            ' API 呼出異常
    '            GoTo Error
            End If
        End If

        ' 読みだすべきデータがパイプから存在せず、かつ
        ' プロセスの終了コードが正しく設定されていたら、ループを終了させる
        If (hasData = False) And (procExitCode <> STILL_ACTIVE) Then
        
            ' データがあれば出力
            If recvBuffer2Empty <> True Then
                strBuffer = StrConv(recvBuffer2, vbUnicode)
                If disableIoTrace_ <> True Then
                    Debug.Print "StdO " & strBuffer
                End If
                If Not func Is Nothing Then
                    Call func.OnRecieveLineFromStdout(Me, strBuffer)
                End If
            End If
            ReDim recvBuffer2(0)
            recvBuffer2Empty = True
            
            If errBuffer2Empty <> True Then
                strBuffer = StrConv(errBuffer2, vbUnicode)
                If disableIoTrace_ <> True Then
                    Debug.Print "StdE " & strBuffer
                End If
                If Not func Is Nothing Then
                    Call func.OnRecieveLineFromStderr(Me, strBuffer)
                End If
            End If
            ReDim errBuffer2(0)
            errBuffer2Empty = True
      
            Exit Do
            
        ElseIf hasData = False Then
        
            ' アイドルコールバックの判定
            If idleTimeoutCount > 0 Then
                If idleCount > idleTimeoutCount Then
                    If Not func Is Nothing Then
                        If recvBuffer2Empty = True Then
                            Call func.OnIdle(Me, "")
                        Else
                            Call func.OnIdle(Me, StrConv(recvBuffer2, vbUnicode))
                        End If
                    End If
                    
                    ' 多重でコールバックしないように抑止する
                    idleCount = -1
                ElseIf idleCount >= 0 Then
                    idleCount = idleCount + 1
                End If
            End If
            
            ' ディープアイドルコールバックの判定
            If deepIdleTimeoutCount > 0 Then
                If deepIdleCount > deepIdleTimeoutCount Then
                    If Not func Is Nothing Then
                        Call func.OnDeepIdle(Me)
                    Else
                        Call RequestTerminate(EXIT_FAILURE)
                    End If
                    
                    ' 多重でコールバックしないように抑止する
                    deepIdleCount = -1
                ElseIf deepIdleCount >= 0 Then
                    deepIdleCount = deepIdleCount + 1
                End If
            End If
        
        End If
    
        ' 待ち
        If hasRemainData = False Then
            Sleep POLLING_MILLISEC
        End If
        
        ' イベント処理(マクロの多重起動は呼出し元で対処しておくこと)
        DoEvents
    
    Loop
          
    ' 切断完了コールバックを呼び出す
    If Not func Is Nothing Then
        Call func.OnDisconnected(Me)
    End If
        
    Application.Cursor = xlDefault
        
    ' ハンドルを閉じる
    CloseHandle udtProcInfo.hThread
    CloseHandle udtProcInfo.hProcess
    udtProcInfo.hProcess = 0
    CloseHandle hOutWrite
    CloseHandle hOutRead
    CloseHandle hInWrite
    CloseHandle hInRead
    CloseHandle hErrWrite
    CloseHandle hErrRead

    ' 戻り値の設定
    ExecuteProcess = procExitCode
    
    Debug.Print "Info ProcessController: Child process(" & commandLine & ", pid=" & udtProcInfo.dwProcessId & ") has exited. code=" & procExitCode
    
    Exit Function

' エラー処理
Error:

    Debug.Print "Warn ProcessController: Child process(" & commandLine & ") could not be executed with error."
    
    ' ハンドルを閉じる
    CloseHandle udtProcInfo.hThread
    CloseHandle udtProcInfo.hProcess
    If udtProcInfo.hProcess <> 0 Then
        Call RequestTerminate(EXIT_FAILURE)
    End If
    udtProcInfo.hProcess = 0
    CloseHandle hOutWrite
    CloseHandle hOutRead
    CloseHandle hInWrite
    CloseHandle hInRead
    CloseHandle hErrWrite
    CloseHandle hErrRead
    
    ' 戻り値の設定
    ExecuteProcess = EXIT_FAILURE

End Function

' -----------------------------------------------------------------------------
' 新しいインスタンスを初期化します。
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    
    ' タイムアウトの設定
    Call ResetIdleTimeoutMilliseconds
    Call ResetDeepIdleTimeoutMilliseconds

End Sub

' -----------------------------------------------------------------------------
' インスタンスが破棄される際の処理をします。
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()

'    Debug.Print "Info ProcessController: terminate."

    ' ハンドルを閉じる
    CloseHandle udtProcInfo.hThread
    CloseHandle udtProcInfo.hProcess
    If udtProcInfo.hProcess <> 0 Then
        Call RequestTerminate(EXIT_FAILURE)
    End If
    udtProcInfo.hProcess = 0
    CloseHandle hOutWrite
    CloseHandle hOutRead
    CloseHandle hInWrite
    CloseHandle hInRead
    CloseHandle hErrWrite
    CloseHandle hErrRead
    
End Sub
