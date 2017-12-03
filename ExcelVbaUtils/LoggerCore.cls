VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerCore"
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

' Dependency: IAppender, CLoggingEvent, ILog

Implements ILog

' �q�� Logger �̃R���N�V������ێ����܂��B
Private children As Collection

' ���� Logger �̏o�͐�� Appender ��ێ����܂��B
Private appenders As Collection

' �J���p�̃f�o�b�O���b�Z�[�W�̏o�͂��L�����ǂ�����ێ����܂��B
Private isDebugEnabled_ As Boolean
' �V�X�e����~�͂��Ȃ����A���ƂȂ��Q�̏o�͂��L�����ǂ�����ێ����܂��B
Private isErrorEnabled_ As Boolean
' �V�X�e����~����悤�Ȓv���I�ȏ�Q�̏o�͂��L�����ǂ�����ێ����܂��B
Private isFatalEnabled_ As Boolean
' ���샍�O�Ȃǂ̏��̏o�͂��L�����ǂ�����ێ����܂��B
Private isInfoEnabled_ As Boolean
' ��Q�ł͂Ȃ����ӌx���̏o�͂��L�����ǂ�����ێ����܂��B
Private isWarnEnabled_ As Boolean

' -----------------------------------------------------------------------------
' LoggerCore �N���X�̐V�����C���X�^���X�����������܂��B
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()

    Set appenders = New Collection
    
    isDebugEnabled_ = False
    isErrorEnabled_ = True
    isFatalEnabled_ = True
    isInfoEnabled_ = True
    isWarnEnabled_ = True

End Sub

' -----------------------------------------------------------------------------
' �q�� Logger �̃R���N�V�������N���A���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_ClearChildren()
    Set children = Nothing
End Sub

' -----------------------------------------------------------------------------
' �q�� Logger ��o�^���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_RegistChild(child As ILog)
    If children Is Nothing Then
        Set children = New Collection
    End If
    Call children.Add(child)
End Sub

' -----------------------------------------------------------------------------
' ���� Logger �̏o�͐�� Appender ���N���A���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_ClearAppenders()
    Set appenders = New Collection
End Sub

' -----------------------------------------------------------------------------
' ���� Logger �̏o�͐�� Appender ��o�^���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_RegistAppender(appender As IAppender)
    Call appenders.Add(appender)
End Sub

' -----------------------------------------------------------------------------
' �J���p�̃f�o�b�O���b�Z�[�W�̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsDebugEnabled() As Boolean

    Dim result As Boolean
    result = isDebugEnabled_
    
    If result = False And Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsDebugEnabled = True Then
                result = True
            End If
        Next
    End If
    
    ILog_IsDebugEnabled = result
    
End Property
' -----------------------------------------------------------------------------
' �J���p�̃f�o�b�O���b�Z�[�W�̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsDebugEnabled(isDebugEnabled__ As Boolean)
    isDebugEnabled_ = isDebugEnabled__
End Property

' -----------------------------------------------------------------------------
' �V�X�e����~�͂��Ȃ����A���ƂȂ��Q�̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsErrorEnabled() As Boolean

    Dim result As Boolean
    result = isErrorEnabled_
    
    If result = False And Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsErrorEnabled = True Then
                result = True
            End If
        Next
    End If

    ILog_IsErrorEnabled = result

End Property
' -----------------------------------------------------------------------------
' �V�X�e����~�͂��Ȃ����A���ƂȂ��Q�̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsErrorEnabled(isErrorEnabled__ As Boolean)
    isErrorEnabled_ = isErrorEnabled__
End Property

' -----------------------------------------------------------------------------
' �V�X�e����~����悤�Ȓv���I�ȏ�Q�̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsFatalEnabled() As Boolean

    Dim result As Boolean
    result = isFatalEnabled_
    
    If result = False And Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsFatalEnabled = True Then
                result = True
            End If
        Next
    End If

    ILog_IsFatalEnabled = result

End Property
' -----------------------------------------------------------------------------
' �V�X�e����~����悤�Ȓv���I�ȏ�Q�̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsFatalEnabled(isFatalEnabled__ As Boolean)
    isFatalEnabled_ = isFatalEnabled__
End Property

' -----------------------------------------------------------------------------
' ���샍�O�Ȃǂ̏��̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsInfoEnabled() As Boolean

    Dim result As Boolean
    result = isInfoEnabled_
    
    If result = False And Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsInfoEnabled = True Then
                result = True
            End If
        Next
    End If

    ILog_IsInfoEnabled = result

End Property
' -----------------------------------------------------------------------------
' ���샍�O�Ȃǂ̏��̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsInfoEnabled(isInfoEnabled__ As Boolean)
    isInfoEnabled_ = isInfoEnabled__
End Property

' -----------------------------------------------------------------------------
' ��Q�ł͂Ȃ����ӌx���̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsWarnEnabled() As Boolean

    Dim result As Boolean
    result = isWarnEnabled_
    
    If result = False And Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsWarnEnabled = True Then
                result = True
            End If
        Next
    End If

    ILog_IsWarnEnabled = result

End Property
' -----------------------------------------------------------------------------
' ��Q�ł͂Ȃ����ӌx���̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsWarnEnabled(isWarnEnabled__ As Boolean)
    isWarnEnabled_ = isWarnEnabled__
End Property

' -----------------------------------------------------------------------------
' �J���p�̃f�o�b�O���b�Z�[�W�̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogDebug(message_ As String)

    If isDebugEnabled_ = True Then
        Call LogCore(LvlDebug, message_)
    End If
    
    If Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsDebugEnabled = True Then
                Call child.LogDebug(message_)
            End If
        Next
    End If
    
End Sub

' -----------------------------------------------------------------------------
' �V�X�e����~�͂��Ȃ����A���ƂȂ��Q�̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogError(message_ As String)

    If isErrorEnabled_ = True Then
        Call LogCore(LvlError, message_)
    End If
    
    If Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsErrorEnabled = True Then
                Call child.LogError(message_)
            End If
        Next
    End If
    
End Sub

' -----------------------------------------------------------------------------
' �V�X�e����~����悤�Ȓv���I�ȏ�Q�̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogFatal(message_ As String)

    If isFatalEnabled_ = True Then
        Call LogCore(LvlFatal, message_)
    End If
    
    If Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsFatalEnabled = True Then
                Call child.LogFatal(message_)
            End If
        Next
    End If
    
End Sub

' -----------------------------------------------------------------------------
' ���샍�O�Ȃǂ̏��̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogInfo(message_ As String)

    If isInfoEnabled_ = True Then
        Call LogCore(LvlInfo, message_)
    End If
    
    If Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsInfoEnabled = True Then
                Call child.LogInfo(message_)
            End If
        Next
    End If
    
End Sub

' -----------------------------------------------------------------------------
' ��Q�ł͂Ȃ����ӌx���̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogWarn(message_ As String)

    If isWarnEnabled_ = True Then
        Call LogCore(LvlWarn, message_)
    End If
    
    If Not (children Is Nothing) Then
        Dim child As ILog
        For Each child In children
            If child.IsWarnEnabled = True Then
                Call child.LogWarn(message_)
            End If
        Next
    End If
    
End Sub

' -----------------------------------------------------------------------------
' ���O�̎��������s���܂��B
' -----------------------------------------------------------------------------
Private Sub LogCore(loglevel_ As LogLevel, message_ As String)

    Dim loggingEvent As CLoggingEvent
    Set loggingEvent = New CLoggingEvent
    
    loggingEvent.LogLevel = loglevel_
    loggingEvent.message = message_
    
    Dim appender As IAppender
    If Not (appenders Is Nothing) Then
        For Each appender In appenders
            Call appender.DoAppend(loggingEvent)
        Next
    End If
    
End Sub