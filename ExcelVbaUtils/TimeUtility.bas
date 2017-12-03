Attribute VB_Name = "TimeUtility"
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

' Dependency: PInvoke
 
#Const ENABLE_TEST_METHODS = 1

' -----------------------------------------------------------------------------
' ������ϊ��̃t�H�[�}�b�g��ʂ�\���܂��B
' -----------------------------------------------------------------------------
Public Enum TimeUtility_ToStringType
    LONGTIME_WITH_MSEC = 0
End Enum
 
' -----------------------------------------------------------------------------
' �����̎擾���s���܂��B
' <IN> st As SYSTEMTIME �擾�����������i�[����\���́B
' <IN> localTime As Boolean ���n���Ԃ��Q�Ƃ���ꍇ�� True�AUTC �̏ꍇ�� False�B
' -----------------------------------------------------------------------------
Private Sub GetTimeCore(ByRef st As SYSTEMTIME, localTime As Boolean)
    If localTime <> True Then
        Call GetSystemTime(st)
    Else
        Call GetLocalTime(st)
    End If
End Sub
 
' -----------------------------------------------------------------------------
' ���n���Ԃ𕶎���Ƃ��ĕԂ��܂��B
' <IN> toStringType As TimeUtility_ToStringType ������ϊ��̃t�H�[�}�b�g��ʁB
' <OUT> String ���n���Ԃ��t�H�[�}�b�g����������B
' -----------------------------------------------------------------------------
Public Function LocalNowToString(toStringType As TimeUtility_ToStringType) As String

    Dim st As SYSTEMTIME
    
    Call GetTimeCore(st, True)

    LocalNowToString = Format(st.wYear, "0000") & "/" & _
                       Format(st.wMonth, "00") & "/" & _
                       Format(st.wDay, "00") & " " & _
                       Format(st.wHour, "00") & ":" & _
                       Format(st.wMinute, "00") & ":" & _
                       Format(st.wSecond, "00") & "." & _
                       Format(st.wMilliseconds, "000")

End Function

