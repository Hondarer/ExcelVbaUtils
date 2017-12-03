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

' Dependency: IAppender, CLoggingEvent, LogLevelEnum, FileUtility

Implements IAppender

' ���̃I�u�W�F�N�g�̖��̂�ێ����܂��B
Private name_ As String

' ���� Appender �̏o�͐�p�X��ێ����܂��B
Private filePath_ As String

' -----------------------------------------------------------------------------
' DebugPrintAppender �N���X�̐V�����C���X�^���X�����������܂��B
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    name_ = "TextFileAppender"
    filePath_ = ThisWorkbook.Path & "\" & RemoveExtension(ThisWorkbook.Name) & ".log"
End Sub

' -----------------------------------------------------------------------------
' Appender ����܂��B
' -----------------------------------------------------------------------------
Private Sub IAppender_AppenderClose()
    ' NOP
End Sub

' -----------------------------------------------------------------------------
' CLoggingEvent ��ǉ����܂��B
' -----------------------------------------------------------------------------
Private Sub IAppender_DoAppend(loggingEvent As CLoggingEvent)
    
    Dim fileno As Long
    fileno = FreeFile
    
    On Error GoTo errHandler
        Open filePath_ For Append As fileno
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
' ���̃I�u�W�F�N�g�̖��̂��擾���܂��B
' -----------------------------------------------------------------------------
Public Property Get IAppender_Name() As String
    IAppender_Name = name_
End Property
' -----------------------------------------------------------------------------
' ���̃I�u�W�F�N�g�̖��̂�ݒ肵�܂��B
' -----------------------------------------------------------------------------
Public Property Let Name(name__ As String)
    name_ = name__
End Property

' -----------------------------------------------------------------------------
' �o�͐�t�@�C���p�X���擾���܂��B
' -----------------------------------------------------------------------------
Public Property Get filePath() As String
    filePath = filePath_
End Property
' -----------------------------------------------------------------------------
' �o�͐�t�@�C���p�X��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Public Property Let filePath(filePath__ As String)
    filePath_ = filePath__
End Property

