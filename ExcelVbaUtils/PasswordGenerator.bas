Attribute VB_Name = "PasswordGenerator"
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

' �p�X���[�h�ɗp���镶���̑g�ݍ��킹��\���܂��B
Public Enum gpPasswordType
    ' �A���t�@�x�b�g�̂݁B
    gpAlphabetOnly = 0
    ' �A���t�@�x�b�g�Ɛ����B
    gpAlphabetAndNumeric = 1
    ' �A���t�@�x�b�g�Ɛ����ƋL���B
    gpIncludeSymbol = 2
End Enum

' �ԈႦ�����ȕ����͈Ӑ}�I�ɊO���A�Ώۂɂ��Ă��Ȃ�
Private Const GPELEMENTS_ALPHABET = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ"
Private Const GPELEMENTS_NUMERIC = "0123456789"
Private Const GPELEMENTS_SYMBOLS = "!#$%&()+-./:<=>?[]^_|" ' @ �� AD �ł֑͋�����

Private Const GPMINLENGTH = 8

' -----------------------------------------------------------------------------
' �p�X���[�h������𐶐����܂��B
' <IN> passwordType As gpPasswordType �p�X���[�h�ɗp���镶���̑g�ݍ��킹�B
' <IN>passwordLength As Long �p�X���[�h�̒����B�Œᕶ�����������ꍇ�́A�Œᕶ�����ɕ␳����܂��B
' <OUT> String �������ꂽ�p�X���[�h������B
' -----------------------------------------------------------------------------
Public Function GenaretePassword(passwordType As gpPasswordType, passwordLength As Long) As String
    
    Dim result As String
    Dim count As Long
    
    Dim numerics As Long
    Dim symbols As Long
    
    Dim insertIndex As Long

    ' �Œᕶ�����������ꍇ�́A�Œᕶ�����ɕ␳
    If passwordLength < GPMINLENGTH Then
        passwordLength = GPMINLENGTH
    End If

    ' �����n��̏�����
    Randomize
    
    ' �܂܂�鐔���̕��������Z�o
    If passwordType = gpAlphabetAndNumeric Or passwordType = gpIncludeSymbol Then
        ' �S�̂� 25% ���x(�������A�Œᐔ 1)
        numerics = Int(Rnd * passwordLength / 4) + 1
    End If
    
    ' �܂܂��L���̕��������Z�o
    If passwordType = gpIncludeSymbol Then
        ' �S�̂� 12% ���x(�������A�Œᐔ 1)
        symbols = Int(Rnd * passwordLength / 8) + 1
    End If
    
    ' �A���t�@�x�b�g�̃p�X���[�h�𐶐�
    For count = 1 To (passwordLength - numerics - symbols)
        result = result & Mid(GPELEMENTS_ALPHABET, Int(Rnd * Len(GPELEMENTS_ALPHABET)) + 1, 1)
    Next
    
    ' ���������̃p�X���[�h�𐶐����đ}��
    For count = 1 To numerics
        insertIndex = Int(Rnd * (Len(result) + 1))
        result = Left(result, insertIndex) & Mid(GPELEMENTS_NUMERIC, Int(Rnd * Len(GPELEMENTS_NUMERIC)) + 1, 1) & Mid(result, insertIndex + 1)
    Next
    
    ' �L�������̃p�X���[�h�𐶐����đ}��
    For count = 1 To symbols
        insertIndex = Int(Rnd * (Len(result) + 1))
        result = Left(result, insertIndex) & Mid(GPELEMENTS_SYMBOLS, Int(Rnd * Len(GPELEMENTS_SYMBOLS)) + 1, 1) & Mid(result, insertIndex + 1)
    Next
    
    GenaretePassword = result

End Function

#If ENABLE_TEST_METHODS = 1 Then

' -----------------------------------------------------------------------------
' GenaretePassword ���\�b�h�̃e�X�g���s���܂��B
' -----------------------------------------------------------------------------
Public Sub GenaretePasswordTest()
    Debug.Print GenaretePassword(gpIncludeSymbol, 8)
End Sub

#End If

