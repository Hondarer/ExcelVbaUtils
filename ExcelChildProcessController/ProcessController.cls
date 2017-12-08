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
' �ȉ��ɒ�߂�����ɏ]���A�{�\�t�g�E�F�A����ъ֘A�����̃t�@�C��
' �i�ȉ��u�\�t�g�E�F�A�v�j�̕������擾���邷�ׂĂ̐l�ɑ΂��A�\�t�g�E�F�A��
' �������Ɉ������Ƃ𖳏��ŋ����܂��B����ɂ́A�\�t�g�E�F�A�̕������g�p�A���ʁA
' �ύX�A�����A�f�ځA�Еz�A�T�u���C�Z���X�A�����/�܂��͔̔����錠���A
' ����у\�t�g�E�F�A��񋟂��鑊��ɓ������Ƃ������錠�����������Ɋ܂܂�܂��B
'
' ��L�̒��쌠�\������і{�����\�����A�\�t�g�E�F�A�̂��ׂĂ̕����܂��͏d�v��
' �����ɋL�ڂ�����̂Ƃ��܂��B
'
' �\�t�g�E�F�A�́u����̂܂܁v�ŁA�����ł��邩�Öقł��邩���킸�A
' ����̕ۏ؂��Ȃ��񋟂���܂��B
' �����ł����ۏ؂Ƃ́A���i���A����̖ړI�ւ̓K�����A����ь�����N�Q�ɂ��Ă�
' �ۏ؂��܂݂܂����A����Ɍ��肳�����̂ł͂���܂���B
' ��҂܂��͒��쌠�҂́A�_��s�ׁA�s�@�s�ׁA�܂��͂���ȊO�ł��낤�ƁA
' �\�t�g�E�F�A�ɋN���܂��͊֘A���A���邢�̓\�t�g�E�F�A�̎g�p�܂��͂��̑���
' �����ɂ���Đ������؂̐����A���Q�A���̑��̋`���ɂ��ĉ���̐ӔC������Ȃ�
' ���̂Ƃ��܂��B
'
' -----------------------------------------------------------------------------
'
' Excel �v���Z�X���䃉�C�u����
'
' �R���\�[�� �A�v���P�[�V������C�ӂ̈����ŋN�����A
' �W�����o�͂� VBA ���爵���܂��B
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

' �p�C�v�ʐM�̑��M�o�b�t�@�P�ʂ�\���܂��B
Private Const BUFFER_LIMIT_LENGTH = 1024&

Dim idleCount As Long
Dim idleTimeoutCount As Long
Dim deepIdleCount As Long
Dim deepIdleTimeoutCount As Long

Dim recvBuffer1() As Byte ' 1 ���o�b�t�@(�V�t�g JIS �̃A���C�����m��)
Dim recvBuffer2() As Byte ' 2 ���o�b�t�@(���C���o�b�t�@)

Dim recvBuffer1Empty As Boolean
Dim recvBuffer2Empty As Boolean

Dim errBuffer1() As Byte ' 1 ���o�b�t�@(�V�t�g JIS �̃A���C�����m��)
Dim errBuffer2() As Byte ' 2 ���o�b�t�@(���C���o�b�t�@)

Dim errBuffer1Empty As Boolean
Dim errBuffer2Empty As Boolean

Dim disableIoTrace_ As Boolean

Dim leaveStdin_ As Boolean
Dim leaveStdout_ As Boolean
Dim leaveStderr_ As Boolean

Dim showConsoleWindow_ As Boolean

' -----------------------------------------------------------------------------
' �C�~�f�B�G�C�g �E�C���h�E�ւ̃f�o�b�O�o�͖����ݒ���擾���܂��B
' -----------------------------------------------------------------------------
Property Get DisableIoTrace() As Boolean
    DisableIoTrace = disableIoTrace_
End Property
' -----------------------------------------------------------------------------
' �C�~�f�B�G�C�g �E�C���h�E�ւ̃f�o�b�O�o�͖����ݒ��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let DisableIoTrace(disableIoTrace__ As Boolean)
    disableIoTrace_ = disableIoTrace__
End Property

' -----------------------------------------------------------------------------
' �W�����͂��R���\�[�� �E�C���h�E�Ɏc�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Property Get LeaveStdin() As Boolean
    LeaveStdin = leaveStdin_
End Property
' -----------------------------------------------------------------------------
' �W�����͂��R���\�[�� �E�C���h�E�Ɏc�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let LeaveStdin(leaveStdin__ As Boolean)
    leaveStdin_ = leaveStdin__
End Property

' -----------------------------------------------------------------------------
' �W���o�͂��R���\�[�� �E�C���h�E�Ɏc�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Property Get LeaveStdout() As Boolean
    LeaveStdout = leaveStdout_
End Property
' -----------------------------------------------------------------------------
' �W���o�͂��R���\�[�� �E�C���h�E�Ɏc�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let LeaveStdout(leaveStdout__ As Boolean)
    leaveStdout_ = leaveStdout__
End Property

' -----------------------------------------------------------------------------
' �W���G���[�o�͂��R���\�[�� �E�C���h�E�Ɏc�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Property Get LeaveStderr() As Boolean
    LeaveStderr = leaveStderr_
End Property
' -----------------------------------------------------------------------------
' �W���G���[�o�͂��R���\�[�� �E�C���h�E�Ɏc�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let LeaveStderr(leaveStderr__ As Boolean)
    leaveStderr_ = leaveStderr__
End Property

' -----------------------------------------------------------------------------
' �R���\�[�� �E�C���h�E��\�����邩�ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Property Get ShowConsoleWindow() As Boolean
    ShowConsoleWindow = showConsoleWindow_
End Property
' -----------------------------------------------------------------------------
' �R���\�[�� �E�C���h�E��\�����邩�ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let ShowConsoleWindow(showConsoleWindow__ As Boolean)
    showConsoleWindow_ = showConsoleWindow__
End Property

' -----------------------------------------------------------------------------
' �^�[�Q�b�g�̕W�����͂ɕ������^���܂��B
' -----------------------------------------------------------------------------
Public Function SendTextToStdin(sendstring As String) As Boolean

    Dim dwWritten As Long
    Dim writeBuffer() As Byte
    
    If (hInWrite = 0) Or (leaveStdin_ = True) Then
        SendTextToStdin = False
        Exit Function
    End If
    
    If disableIoTrace_ <> True Then
        GetDefaultLogger.LogDebug "[ProcessController][SendTextToStdin] StdI " & sendstring
    End If
    
    writeBuffer = StrConv(sendstring, vbFromUnicode)
    
    Dim remain As Long
    remain = UBound(writeBuffer) + 1
    
    While remain > 0
        
        ' �o�̓o�b�t�@���t�Ńv���Z�X���f�b�h���b�N���邱�Ƃ�h�����߁A
        ' �o�͂�ǂݎ���ăo�b�t�@���m�ۂ��Ă���
        Call SuckRecvBuffer
    
        If remain < BUFFER_LIMIT_LENGTH Then
            If WriteFile(hInWrite, writeBuffer(UBound(writeBuffer) - remain + 1), remain, dwWritten, ov1) = 0 Then
                SendTextToStdin = False
                Exit Function
            End If
        Else
            If WriteFile(hInWrite, writeBuffer(UBound(writeBuffer) - remain + 1), BUFFER_LIMIT_LENGTH, dwWritten, ov1) = 0 Then
                SendTextToStdin = False
                Exit Function
            End If
        End If
        
        remain = remain - dwWritten
        
    Wend
    
    SendTextToStdin = True

End Function

' -----------------------------------------------------------------------------
' �^�[�Q�b�g�̕W�����͂ɕ�����Ɖ��s�R�[�h��^���܂��B
' -----------------------------------------------------------------------------
Public Function SendLineToStdin(sendstring As String) As Boolean

    SendLineToStdin = SendTextToStdin(sendstring & vbCrLf)
    
End Function

' -----------------------------------------------------------------------------
' �A�C�h�����o���Ԃ��~���b�Őݒ肵�܂��B
' -----------------------------------------------------------------------------
Public Sub SetIdleTimeoutMilliseconds(timeout As Long)
    idleCount = 0
    idleTimeoutCount = timeout / POLLING_MILLISEC
End Sub

' -----------------------------------------------------------------------------
' �A�C�h�����o���Ԃ����Z�b�g���܂��B
' -----------------------------------------------------------------------------
Public Sub ResetIdleTimeoutMilliseconds()
    Call SetIdleTimeoutMilliseconds(DEFAULT_IDLE_TIMEOUTMILLISEC)
End Sub

' -----------------------------------------------------------------------------
' �A�C�h�����o���Ԃ𖳌��ɂ��܂��B
' -----------------------------------------------------------------------------
Public Sub DisableIdleTimeoutMilliseconds()
    Call SetIdleTimeoutMilliseconds(0)
End Sub

' -----------------------------------------------------------------------------
' �f�B�[�v�A�C�h�����o���Ԃ��~���b�Őݒ肵�܂��B
' -----------------------------------------------------------------------------
Public Sub SetDeepIdleTimeoutMilliseconds(timeout As Long)
    deepIdleCount = 0
    deepIdleTimeoutCount = timeout / POLLING_MILLISEC
End Sub

' -----------------------------------------------------------------------------
' �f�B�[�v�A�C�h�����o���Ԃ����Z�b�g���܂��B
' -----------------------------------------------------------------------------
Public Sub ResetDeepIdleTimeoutMilliseconds()
    SetDeepIdleTimeoutMilliseconds (DEFAULT_DEEPIDLE_TIMEOUTMILLISEC)
End Sub

' -----------------------------------------------------------------------------
' �f�B�[�v�A�C�h�����o���Ԃ𖳌��ɂ��܂��B
' -----------------------------------------------------------------------------
Public Sub DisableDeepIdleTimeoutMilliseconds()
    SetDeepIdleTimeoutMilliseconds (0)
End Sub

' -----------------------------------------------------------------------------
' �W���o�͂̃��C���o�b�t�@���N���A���܂��B
' -----------------------------------------------------------------------------
Public Sub ClearStdoutLineBuffer()
    ReDim recvBuffer2(0)
    recvBuffer2Empty = True
End Sub

' -----------------------------------------------------------------------------
' �^�[�Q�b�g�̃v���Z�X�̏I����v�����܂��B
' -----------------------------------------------------------------------------
Public Function RequestTerminate(uExitCode As Long) As Boolean
    
    ' �g�p�֎~�̏I���R�[�h�̓G���[�Ƃ���
    If uExitCode = STILL_ACTIVE Then
        RequestTerminate = False
        Exit Function
    End If
    
    GetDefaultLogger.LogInfo "[ProcessController][RequestTerminate] Child process(pid=" & udtProcInfo.dwProcessId & ") request stop with code=" & uExitCode & "."
    
    If TerminateProcess(udtProcInfo.hProcess, uExitCode) = 0 Then
        GetDefaultLogger.LogWarn "[ProcessController][RequestTerminate] Child process(pid=" & udtProcInfo.dwProcessId & ") stop failed(in general it has already terminated)."
        RequestTerminate = False
        Exit Function
    End If
    
    RequestTerminate = True
    
End Function

' -----------------------------------------------------------------------------
' �v���Z�X�����s���܂��B
' -----------------------------------------------------------------------------
Public Function ExecuteProcess(commandLine As String, func As IProcessControlCallbacks) As Long

    Dim sa As SECURITY_ATTRIBUTES

    Dim udtStartInfo As STARTUPINFO

    Dim procExitCode As Long
    
    Dim strBuffer As String
    
    ReDim recvBuffer1(0)
    ReDim recvBuffer2(0)
    
    recvBuffer1Empty = True
    recvBuffer2Empty = True
    
    ReDim errBuffer1(0)
    ReDim errBuffer2(0)
    
    errBuffer1Empty = True
    errBuffer2Empty = True
    
    Dim sjisFirstByte As Byte
    Dim processRecv1Index As Long
    
    ' ���O�Ǘ��@�\�̏�����
    Call LogManagerHelper_LocalInit

    GetDefaultLogger.LogInfo "[ProcessController][ExecuteProcess] �J�n commandLine=" & commandLine

    ' �Z�L�����e�B�f�B�X�N���v�^�̐ݒ�
    Call ZeroMemory(sa, Len(sa))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&

    ' �\���̂̏�����
    udtStartInfo.cb = LenB(udtStartInfo)

    If leaveStdin_ <> True Then
        ' ���̓p�C�v�̍쐬
        If CreatePipe(hInRead, hInWrite, sa, 0) = 0 Then
            GetDefaultLogger.LogFatal "[ProcessController][ExecuteProcess] ���̓p�C�v�̍쐬�Ɏ��s���܂��� Err.LastDllError=" & Err.LastDllError
            GoTo Error
        End If
    
        ' ���͏������݃n���h�����q�v���Z�X�Ɍp���ł��Ȃ�����
        If SetHandleInformation(hInWrite, HANDLE_FLAG_INHERIT, 0) = 0 Then
            GetDefaultLogger.LogFatal "[ProcessController][ExecuteProcess] ���̓p�C�v�̏������݃n���h���̌p���ݒ�Ɏ��s���܂��� Err.LastDllError=" & Err.LastDllError
            GoTo Error
        End If
        
        udtStartInfo.hStdInput = hInRead
    Else
        udtStartInfo.hStdInput = GetStdHandle(STD_INPUT_HANDLE)
    End If

    If leaveStdout_ <> True Then
        ' �o�̓p�C�v�̍쐬
        If CreatePipe(hOutRead, hOutWrite, sa, 0) = 0 Then
            GetDefaultLogger.LogFatal "[ProcessController][ExecuteProcess] �o�̓p�C�v�̍쐬�Ɏ��s���܂��� Err.LastDllError=" & Err.LastDllError
            GoTo Error
        End If
    
        ' �o�͓ǂݍ��݃n���h�����q�v���Z�X�Ɍp���ł��Ȃ�����
        If SetHandleInformation(hOutRead, HANDLE_FLAG_INHERIT, 0) = 0 Then
            GetDefaultLogger.LogFatal "[ProcessController][ExecuteProcess] �o�̓p�C�v�̓ǂݍ��݃n���h���̌p���ݒ�Ɏ��s���܂��� Err.LastDllError=" & Err.LastDllError
            GoTo Error
        End If
    
        udtStartInfo.hStdOutput = hOutWrite
    Else
        udtStartInfo.hStdOutput = GetStdHandle(STD_OUTPUT_HANDLE)
    End If

    If leaveStderr_ <> True Then
        ' �G���[�o�̓p�C�v�̍쐬
        If CreatePipe(hErrRead, hErrWrite, sa, 0) = 0 Then
            GetDefaultLogger.LogFatal "[ProcessController][ExecuteProcess] �G���[�o�̓p�C�v�̍쐬�Ɏ��s���܂��� Err.LastDllError=" & Err.LastDllError
            GoTo Error
        End If
    
        ' �G���[�o�͓ǂݍ��݃n���h�����q�v���Z�X�Ɍp���ł��Ȃ�����
        If SetHandleInformation(hErrRead, HANDLE_FLAG_INHERIT, 0) = 0 Then
            GetDefaultLogger.LogFatal "[ProcessController][ExecuteProcess] �G���[�o�̓p�C�v�̓ǂݍ��݃n���h���̌p���ݒ�Ɏ��s���܂��� Err.LastDllError=" & Err.LastDllError
            GoTo Error
        End If
    
        udtStartInfo.hStdError = hErrWrite
    Else
        udtStartInfo.hStdError = GetStdHandle(STD_ERROR_HANDLE)
    End If
    
    udtStartInfo.dwFlags = STARTF_USESTDHANDLES
    
    ' �E�C���h�E����������
    If showConsoleWindow_ <> True Then
        udtStartInfo.wShowWindow = SW_HIDE
        udtStartInfo.dwFlags = udtStartInfo.dwFlags Or STARTF_USESHOWWINDOW
    End If

    ' �v���Z�X�N��
    If CreateProcess(vbNullString, commandLine, ByVal 0&, ByVal 0&, 1&, _
        NORMAL_PRIORITY_CLASS, ByVal 0&, ThisWorkbook.Path, udtStartInfo, udtProcInfo) = 0 Then
        Call MsgBox("�v���Z�X�̐����Ɏ��s���܂����B�G���[�R�[�h:" & Err.LastDllError & vbCrLf & commandLine, Buttons:=vbCritical, Title:="�G���[")
        GetDefaultLogger.LogFatal "[ProcessController][ExecuteProcess] �v���Z�X�̐����Ɏ��s���܂��� Err.LastDllError=" & Err.LastDllError
        GoTo Error
    End If

    GetDefaultLogger.LogInfo "[ProcessController][ExecuteProcess] Child process(" & commandLine & ", pid=" & udtProcInfo.dwProcessId & ") started up."
   
    ' �ڑ������R�[���o�b�N���Ăяo��
    If Not func Is Nothing Then
        Call func.OnConnected(Me)
    End If
    
    ' ���C�����[�v
    
    Application.Cursor = xlWait
    
    Do
    
        Dim hasData As Boolean
        hasData = False
    
        If GetExitCodeProcess(udtProcInfo.hProcess, procExitCode) = 0 Then
            ' �I���R�[�h�擾�ُ�
            GoTo Error
        End If
        
        Call SuckRecvBuffer
        
        ' �W���o�͂���f�[�^��ǂݎ��
        If recvBuffer1Empty = False Then
        
            ' 1 ���o�b�t�@���珇���A2 ���o�b�t�@�ɋl�߂Ȃ���
            ' CR ����� LF �́A2 ���o�b�t�@�ɂ͋l�߂Ȃ�
            ' LF ���o��������A�s�̋�؂�Ƃ��A�R�[���o�b�N���s��
            ' �V�t�g JIS �̏�ʃo�C�g�݂̂��]�����ꍇ�́A1 ���o�b�t�@�Ɏc���Ă���
            
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
                    ' �R�[���o�b�N���{
                    If recvBuffer2Empty = True Then
                        If disableIoTrace_ <> True Then
                            GetDefaultLogger.LogDebug "[ProcessController][ExecuteProcess] StdO "
                        End If
                        If Not func Is Nothing Then
                            Call func.OnRecieveLineFromStdout(Me, "")
                        End If
                    Else
                        strBuffer = StrConv(recvBuffer2, vbUnicode)
                        If disableIoTrace_ <> True Then
                            GetDefaultLogger.LogDebug "[ProcessController][ExecuteProcess] StdO " & strBuffer
                        End If
                        If Not func Is Nothing Then
                            Call func.OnRecieveLineFromStdout(Me, strBuffer)
                        End If
                    End If
                    ReDim recvBuffer2(0)
                    recvBuffer2Empty = True
                Else
                    ' �ςݏグ
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
            
        End If
        
        ' �W���G���[�o�͂���f�[�^��ǂݎ��
        If errBuffer1Empty = False Then
        
            ' 1 ���o�b�t�@���珇���A2 ���o�b�t�@�ɋl�߂Ȃ���
            ' CR ����� LF �́A2 ���o�b�t�@�ɂ͋l�߂Ȃ�
            ' LF ���o��������A�s�̋�؂�Ƃ��A�R�[���o�b�N���s��
            ' �V�t�g JIS �̏�ʃo�C�g�݂̂��]�����ꍇ�́A1 ���o�b�t�@�Ɏc���Ă���
            
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
                    ' �R�[���o�b�N���{
                    If errBuffer2Empty = True Then
                        If disableIoTrace_ <> True Then
                            GetDefaultLogger.LogDebug "[ProcessController][ExecuteProcess] StdE "
                        End If
                        If Not func Is Nothing Then
                            Call func.OnRecieveLineFromStderr(Me, "")
                        End If
                    Else
                        strBuffer = StrConv(errBuffer2, vbUnicode)
                        If disableIoTrace_ <> True Then
                            GetDefaultLogger.LogDebug "[ProcessController][ExecuteProcess] StdE " & strBuffer
                        End If
                        If Not func Is Nothing Then
                            Call func.OnRecieveLineFromStderr(Me, strBuffer)
                        End If
                    End If
                    ReDim errBuffer2(0)
                    errBuffer2Empty = True
                Else
                    ' �ςݏグ
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
            
        End If

        ' �ǂ݂����ׂ��f�[�^���p�C�v���瑶�݂����A����
        ' �v���Z�X�̏I���R�[�h���������ݒ肳��Ă�����A���[�v���I��������
        If (hasData = False) And (procExitCode <> STILL_ACTIVE) Then
        
            ' �f�[�^������Ώo��
            If recvBuffer2Empty <> True Then
                strBuffer = StrConv(recvBuffer2, vbUnicode)
                If disableIoTrace_ <> True Then
                    GetDefaultLogger.LogDebug "[ProcessController][ExecuteProcess] StdO " & strBuffer
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
                    GetDefaultLogger.LogDebug "[ProcessController][ExecuteProcess] StdE " & strBuffer
                End If
                If Not func Is Nothing Then
                    Call func.OnRecieveLineFromStderr(Me, strBuffer)
                End If
            End If
            ReDim errBuffer2(0)
            errBuffer2Empty = True
      
            Exit Do
            
        ElseIf hasData = False Then
        
            ' �A�C�h���R�[���o�b�N�̔���
            If idleTimeoutCount > 0 Then
                If idleCount > idleTimeoutCount Then
                    If Not func Is Nothing Then
                        If recvBuffer2Empty = True Then
                            Call func.OnIdle(Me, "")
                        Else
                            Call func.OnIdle(Me, StrConv(recvBuffer2, vbUnicode))
                        End If
                    End If
                    
                    ' ���d�ŃR�[���o�b�N���Ȃ��悤�ɗ}�~����
                    idleCount = -1
                ElseIf idleCount >= 0 Then
                    idleCount = idleCount + 1
                End If
            End If
            
            ' �f�B�[�v�A�C�h���R�[���o�b�N�̔���
            If deepIdleTimeoutCount > 0 Then
                If deepIdleCount > deepIdleTimeoutCount Then
                    If Not func Is Nothing Then
                        Call func.OnDeepIdle(Me)
                    Else
                        Call RequestTerminate(EXIT_FAILURE)
                    End If
                    
                    ' ���d�ŃR�[���o�b�N���Ȃ��悤�ɗ}�~����
                    deepIdleCount = -1
                ElseIf deepIdleCount >= 0 Then
                    deepIdleCount = deepIdleCount + 1
                End If
            End If
        
            ' �҂�
            Sleep POLLING_MILLISEC
            
        End If
        
        ' �C�x���g����(�}�N���̑��d�N���͌ďo�����őΏ����Ă�������)
        DoEvents
    
    Loop
          
    ' �ؒf�����R�[���o�b�N���Ăяo��
    If Not func Is Nothing Then
        Call func.OnDisconnected(Me)
    End If
        
    Application.Cursor = xlDefault
        
    ' �n���h�������
    CloseHandle udtProcInfo.hThread
    CloseHandle udtProcInfo.hProcess
    udtProcInfo.hProcess = 0
    CloseHandle hOutWrite
    CloseHandle hOutRead
    CloseHandle hInWrite
    CloseHandle hInRead
    CloseHandle hErrWrite
    CloseHandle hErrRead

    ' �߂�l�̐ݒ�
    ExecuteProcess = procExitCode
    
    GetDefaultLogger.LogInfo "[ProcessController][ExecuteProcess] �I�� Child process(" & commandLine & ", pid=" & udtProcInfo.dwProcessId & ") has exited. code=" & procExitCode
    
    Exit Function

' �G���[����
Error:

    
    Application.Cursor = xlDefault
    
    ' �n���h�������
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
    
    ' �߂�l�̐ݒ�
    ExecuteProcess = EXIT_FAILURE
    
    GetDefaultLogger.LogWarn "[ProcessController][ExecuteProcess] �I�� Child process(" & commandLine & ") could not be executed with error."

End Function

' -----------------------------------------------------------------------------
' �o�͂���уG���[�o�͂���f�[�^��ǂݎ��A1 ���o�b�t�@�ɒ~�ς��܂��B
' -----------------------------------------------------------------------------
Private Sub SuckRecvBuffer()

    ' �N�������v���Z�X�̏o�̓o�b�t�@�������ς��ɂȂ�ƁA�v���Z�X�͈ꎞ��~����B
    ' ���̂��߁A�N�������v���Z�X�����Ƀf�[�^���z���グ�Ȃ��ƁA
    ' �����炩��f�[�^��^���悤�Ƃ����Ƃ��Ƀf�b�h���b�N�Ɋׂ�B
    ' �����ɂ́A�}���`�X���b�h�łȂ��Ɩ��̖{���ɑΏ��ł��Ȃ����A
    ' ���M���u���b�N�����āA�p�ɂɃf�[�^���z���グ�邱�ƂőΏ�����B
    
    Dim dwRemain As Long
    Dim nextRecv1Index As Long
    Dim dwRead As Long
    
    If leaveStdout_ <> True Then
        ' �W���o�͂���f�[�^���ǂ߂邱�Ƃ��m�F����
        If PeekNamedPipe(hOutRead, 0&, 0, 0&, dwRemain, ByVal 0) <> 0 Then
            If (dwRemain > 0) Then
                ' �f�[�^������
                
                ' �W���o�͂���f�[�^��ǂݎ��
                nextRecv1Index = 0
                If recvBuffer1Empty = True Then
                    ReDim recvBuffer1(dwRemain - 1)
                    recvBuffer1Empty = False
                Else
                    nextRecv1Index = UBound(recvBuffer1) + 1
                    ReDim Preserve recvBuffer1(UBound(recvBuffer1) + dwRemain)
                End If
                
                If ReadFile(hOutRead, recvBuffer1(nextRecv1Index), dwRemain, dwRead, ov2) = 0 Then
                    ' �v���Z�X�̏I���R�[�h�ɂ�胋�[�v�𔲂��邽�߁A�G���[�͖�������B
                End If
            End If
        Else
            ' �v���Z�X�̏I���R�[�h�ɂ�胋�[�v�𔲂��邽�߁A�G���[�͖�������B
        End If
    End If
    
    If leaveStderr_ <> True Then
        ' �W���G���[�o�͂���f�[�^���ǂ߂邱�Ƃ��m�F����
        If PeekNamedPipe(hErrRead, 0&, 0, 0&, dwRemain, ByVal 0) <> 0 Then
            If (dwRemain > 0) Then
                ' �f�[�^������
                
                ' �W���G���[�o�͂���f�[�^��ǂݎ��
                nextRecv1Index = 0
                If errBuffer1Empty = True Then
                    ReDim errBuffer1(dwRemain - 1)
                    errBuffer1Empty = False
                Else
                    nextRecv1Index = UBound(errBuffer1) + 1
                    ReDim Preserve errBuffer1(UBound(errBuffer1) + dwRemain)
                End If
                
                If ReadFile(hErrRead, errBuffer1(nextRecv1Index), dwRemain, dwRead, ov2) = 0 Then
                    ' �v���Z�X�̏I���R�[�h�ɂ�胋�[�v�𔲂��邽�߁A�G���[�͖�������B
                End If
            End If
        Else
            ' �v���Z�X�̏I���R�[�h�ɂ�胋�[�v�𔲂��邽�߁A�G���[�͖�������B
        End If
    End If

End Sub

' -----------------------------------------------------------------------------
' �V�����C���X�^���X�����������܂��B
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    
    ' �^�C���A�E�g�̐ݒ�
    Call ResetIdleTimeoutMilliseconds
    Call ResetDeepIdleTimeoutMilliseconds

End Sub

' -----------------------------------------------------------------------------
' �C���X�^���X���j�������ۂ̏��������܂��B
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()

    ' �n���h�������
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