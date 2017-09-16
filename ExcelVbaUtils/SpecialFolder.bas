Attribute VB_Name = "SpecialFolder"
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

' -----------------------------------------------------------------------------
' �X�y�V�����t�H���_�̎��ʎq��\���܂��B
' https://msdn.microsoft.com/ja-jp/library/cc364490.aspx
' -----------------------------------------------------------------------------
Public Const SPECIALFOLDERS_ALLUSERS_DESKTOP = "AllUsersDesktop"
Public Const SPECIALFOLDERS_ALLUSERS_STARTMENU = "AllUsersStartMenu"
Public Const SPECIALFOLDERS_ALLUSERS_PROGRAMS = "AllUsersPrograms"
Public Const SPECIALFOLDERS_ALLUSERS_STARTUP = "AllUsersStartup"
Public Const SPECIALFOLDERS_DESKTOP = "Desktop"
Public Const SPECIALFOLDERS_FAVORITES = "Favorites"
Public Const SPECIALFOLDERS_FONTS = "Fonts"
Public Const SPECIALFOLDERS_MYDOCUMENTS = "MyDocuments"
Public Const SPECIALFOLDERS_NETHOOD = "NetHood"
Public Const SPECIALFOLDERS_PRINTHOOD = "PrintHood"
Public Const SPECIALFOLDERS_PROGRAMS = "Programs"
Public Const SPECIALFOLDERS_RECENT = "Recent"
Public Const SPECIALFOLDERS_SENDTO = "SendTo"
Public Const SPECIALFOLDERS_STARTMENU = "StartMenu"
Public Const SPECIALFOLDERS_STARTUP = "Startup"
Public Const SPECIALFOLDERS_TEMPLATES = "Templates"

' -----------------------------------------------------------------------------
' �X�y�V�����t�H���_�̃p�X���擾���܂��B
' <IN> specialFolderName As Variant SPECIALFOLDERS_ �Ŏn�܂�X�y�V�����t�H���_�̎��ʎq�B
' <OUT> String �X�y�V�����t�H���_�̃p�X�B
' -----------------------------------------------------------------------------
Public Function GetSpecialFolder(ByVal specialFolderName As Variant) As String

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' MEMO: SpecialFolders �̃p�����[�^�� ByVal Variant �łȂ��Ɛ��퓮�삵�Ȃ�
    GetSpecialFolder = wsh.SpecialFolders(specialFolderName)
     
    Set wsh = Nothing

End Function

#If ENABLE_TEST_METHODS = 1 Then

' -----------------------------------------------------------------------------
' GetSpecialFolder ���\�b�h�̃e�X�g���s���܂��B
' -----------------------------------------------------------------------------
Public Sub GetSpecialFolderTest()

    Debug.Print GetSpecialFolder("MyDocuments")
    
End Sub

#End If

