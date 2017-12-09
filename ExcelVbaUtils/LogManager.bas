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

' Dependency: None

#Const ENABLE_TEST_METHODS = 1

' �t�H�[���o�b�N�p�� Logger ��ێ����܂��B
Private fallbackLogger_ As ILog

' �f�t�H���g�� Logger ��ێ����܂��B
Private defaultLogger_ As ILog

' �J�e�S���ʂ� Logger ��ێ����܂��B
Private loggers As Object

' -----------------------------------------------------------------------------
' �f�t�H���g�� Logger ���擾���܂��
' <OUT> ILog �f�t�H���g�� Logger�B
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
' �f�t�H���g�� Logger ��ݒ肵�܂��
' -----------------------------------------------------------------------------
Public Sub SetDefaultLogger(defaultLogger__ As ILog)
    Set defaultLogger_ = defaultLogger__
End Sub

' -----------------------------------------------------------------------------
' �J�e�S���ʂ� Logger �̃C�j�V�����`�F�b�N���s���܂��B
' -----------------------------------------------------------------------------
Private Sub InitLoggers()
    If loggers Is Nothing Then
        Set loggers = CreateObject("Scripting.Dictionary")
    End If
End Sub

' -----------------------------------------------------------------------------
' �J�e�S���ʂ� Logger �����������܂��B
' -----------------------------------------------------------------------------
Public Sub ClearLoggers()
    Call InitLoggers
    Call loggers.RemoveAll
End Sub

' -----------------------------------------------------------------------------
' �J�e�S���ʂ� Logger ��o�^���܂��B
' -----------------------------------------------------------------------------
Public Sub RegistLogger(key As String, logger As ILog)
    Call InitLoggers
    Call loggers.Add(key, logger)
End Sub

' -----------------------------------------------------------------------------
' �J�e�S���ʂ� Logger �����o���܂��B
' -----------------------------------------------------------------------------
Public Function GetLogger(key As String) As ILog
    Call InitLoggers
    If Not loggers.Exists(key) Then
        Call GetDefaultLogger.LogFatal("[GetLogger] �L�[ '" & key & "' ��������܂���B")
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
