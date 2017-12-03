VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FallbackLogger"
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

' Dependency: ILog

Implements ILog

' -----------------------------------------------------------------------------
' �q�� Logger �̃R���N�V�������N���A���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_ClearChildren()
    ' NOP
End Sub

' -----------------------------------------------------------------------------
' �q�� Logger ��o�^���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_RegistChild(child As ILog)
    ' NOP
End Sub

' -----------------------------------------------------------------------------
' ���� Logger �̏o�͐�� Appender ���N���A���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_ClearAppenders()
    ' NOP
End Sub

' -----------------------------------------------------------------------------
' ���� Logger �̏o�͐�� Appender ��o�^���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_RegistAppender(appender As IAppender)
    ' NOP
End Sub

' -----------------------------------------------------------------------------
' �J���p�̃f�o�b�O���b�Z�[�W�̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsDebugEnabled() As Boolean
    ILog_IsDebugEnabled = True
End Property
' -----------------------------------------------------------------------------
' �J���p�̃f�o�b�O���b�Z�[�W�̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsDebugEnabled(isDebugEnabled__ As Boolean)
    ' NOP
End Property

' -----------------------------------------------------------------------------
' �V�X�e����~�͂��Ȃ����A���ƂȂ��Q�̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsErrorEnabled() As Boolean
    ILog_IsErrorEnabled = True
End Property
' -----------------------------------------------------------------------------
' �V�X�e����~�͂��Ȃ����A���ƂȂ��Q�̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsErrorEnabled(isErrorEnabled__ As Boolean)
    ' NOP
End Property

' -----------------------------------------------------------------------------
' �V�X�e����~����悤�Ȓv���I�ȏ�Q�̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsFatalEnabled() As Boolean
    ILog_IsFatalEnabled = True
End Property
' -----------------------------------------------------------------------------
' �V�X�e����~����悤�Ȓv���I�ȏ�Q�̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsFatalEnabled(isFatalEnabled__ As Boolean)
    ' NOP
End Property

' -----------------------------------------------------------------------------
' ���샍�O�Ȃǂ̏��̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsInfoEnabled() As Boolean
    ILog_IsInfoEnabled = True
End Property
' -----------------------------------------------------------------------------
' ���샍�O�Ȃǂ̏��̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsInfoEnabled(isInfoEnabled__ As Boolean)
    ' NOP
End Property

' -----------------------------------------------------------------------------
' ��Q�ł͂Ȃ����ӌx���̏o�͂��L�����ǂ������擾���܂��B
' -----------------------------------------------------------------------------
Private Property Get ILog_IsWarnEnabled() As Boolean
    ILog_IsWarnEnabled = True
End Property
' -----------------------------------------------------------------------------
' ��Q�ł͂Ȃ����ӌx���̏o�͂��L�����ǂ�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Private Property Let ILog_IsWarnEnabled(isWarnEnabled__ As Boolean)
    ' NOP
End Property

' -----------------------------------------------------------------------------
' �J���p�̃f�o�b�O���b�Z�[�W�̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogDebug(message_ As String)
    Debug.Print message_
End Sub

' -----------------------------------------------------------------------------
' �V�X�e����~�͂��Ȃ����A���ƂȂ��Q�̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogError(message_ As String)
    Debug.Print message_
End Sub

' -----------------------------------------------------------------------------
' �V�X�e����~����悤�Ȓv���I�ȏ�Q�̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogFatal(message_ As String)
    Debug.Print message_
End Sub

' -----------------------------------------------------------------------------
' ���샍�O�Ȃǂ̏��̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogInfo(message_ As String)
    Debug.Print message_
End Sub

' -----------------------------------------------------------------------------
' ��Q�ł͂Ȃ����ӌx���̏o�͂��s���܂��B
' -----------------------------------------------------------------------------
Private Sub ILog_LogWarn(message_ As String)
    Debug.Print message_
End Sub
