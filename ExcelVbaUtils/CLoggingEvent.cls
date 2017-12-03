VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoggingEvent"
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

' Dependency: TimeUtility

' �^�C���X�^���v��ێ����܂��B
Dim timeStamp_ As String

' ���O���x����ێ����܂��B
Dim loglevel_ As LogLevel

' �R���s���[�^�[����ێ����܂��B
Dim machineName_ As String

' �[������ێ����܂��B
Dim clientName_ As String

' �h���C�������܂ރ��[�U�[����ێ����܂��B
Dim userDomainAndName_ As String

' ���b�Z�[�W��ێ����܂��B
Dim message_ As String

' -----------------------------------------------------------------------------
' CLoggingEvent �N���X�̐V�����C���X�^���X�����������܂��B
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    timeStamp_ = LocalNowToString(LONGTIME_WITH_MSEC)
    machineName_ = Environ("COMPUTERNAME")
    clientName_ = Environ("CLIENTNAME")
    userDomainAndName_ = Environ("USERDOMAIN") & "\" & Environ("USERNAME")
End Sub

' -----------------------------------------------------------------------------
' �^�C���X�^���v���擾���܂��B
' -----------------------------------------------------------------------------
Public Property Get TimeStamp() As String
    TimeStamp = timeStamp_
End Property

' -----------------------------------------------------------------------------
' �R���s���[�^�[�����擾���܂��B
' -----------------------------------------------------------------------------
Public Property Get MachineName() As String
    MachineName = machineName_
End Property

' -----------------------------------------------------------------------------
' �[�������擾���܂��B
' -----------------------------------------------------------------------------
Public Property Get ClientName() As String
    ClientName = clientName_
End Property

' -----------------------------------------------------------------------------
' �h���C�������܂ރ��[�U�[�����擾���܂��B
' -----------------------------------------------------------------------------
Public Property Get UserDomainAndName() As String
    UserDomainAndName = userDomainAndName_
End Property

' -----------------------------------------------------------------------------
' ���b�Z�[�W���擾���܂��B
' -----------------------------------------------------------------------------
Public Property Get LogLevel() As LogLevel
    LogLevel = loglevel_
End Property
' -----------------------------------------------------------------------------
' ���b�Z�[�W��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let LogLevel(loglevel__ As LogLevel)
    loglevel_ = loglevel__
End Property

' -----------------------------------------------------------------------------
' ���b�Z�[�W���擾���܂��B
' -----------------------------------------------------------------------------
Public Property Get message() As String
    message = message_
End Property
' -----------------------------------------------------------------------------
' ���b�Z�[�W��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let message(message__ As String)
    message_ = message__
End Property
