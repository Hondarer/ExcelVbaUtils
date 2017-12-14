Attribute VB_Name = "FileUtility"
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

' FileSystemObject ��ێ����܂��B
Dim fso As Object

' -----------------------------------------------------------------------------
' �w�肳�ꂽ�t�H���_�����݂��邩�Ԃ��܂��B
' <IN> folderName As String �`�F�b�N����t�H���_���B
' <OUT> Boolean �t�H���_�����݂���ꍇ�� True�A���݂��Ȃ��ꍇ�� False�B
' -----------------------------------------------------------------------------
Public Function FolderExists(folderName As String) As Boolean
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    FolderExists = fso.FolderExists(folderName)
    
    Set fso = Nothing

End Function

' -----------------------------------------------------------------------------
' �w�肳�ꂽ�t�@�C�������݂��邩�Ԃ��܂��B
' <IN> folderName As String �`�F�b�N����t�H���_���B
' <OUT> Boolean �t�@�C�������݂���ꍇ�� True�A���݂��Ȃ��ꍇ�� False�B
' -----------------------------------------------------------------------------
Public Function FileExists(fileName As String) As Boolean
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    FileExists = fso.FileExists(fileName)
    
    Set fso = Nothing

End Function

' -----------------------------------------------------------------------------
' �t�@�C�����̊g���q������Ԃ��܂��B
' <IN> filePath As String �Ώۂ̃t�@�C���p�X�B
' <OUT> String �s���I�h���܂ފg���q�����B�s���I�h��������Ȃ��ꍇ�́A�󕶎��B
' -----------------------------------------------------------------------------
Public Function GetExtension(filePath As String) As String

    If InStr(filePath, ".") = 0 Then
        Exit Function
    End If
    
    If InStrRev(filePath, ".") < InStrRev(filePath, "\") Then
        ' �t���p�X�ŁA��ʂ̃t�H���_���Ƀs���I�h���܂܂�Ă���ꍇ
        Exit Function
    End If

    GetExtension = Mid(filePath, InStrRev(filePath, "."))

End Function

' -----------------------------------------------------------------------------
' �t�@�C�����̊g���q��������������Ԃ��܂��
' <IN> filePath As String �Ώۂ̃t�@�C���p�X�B
' <OUT> String �g���q����菜�����t�@�C���p�X�B�s���I�h��������Ȃ��ꍇ�́A���͂����̂܂ܕԂ��܂��B
' -----------------------------------------------------------------------------
Public Function RemoveExtension(filePath As String) As String

    RemoveExtension = Left(filePath, Len(filePath) - Len(GetExtension(filePath)))

End Function

' -----------------------------------------------------------------------------
' ���[�N�u�b�N����̃p�X���A��΃p�X�ɕϊ����܂��B
' <IN> workbookPath As String ���[�N�u�b�N����̑��΃p�X���A��΃p�X�B
' <OUT> String �������ꂽ��΃p�X�B
' -----------------------------------------------------------------------------
Public Function GetAbsolutePathNameFromThisWorkbookPath(workbookPath As String) As String

    ' �J�����g�f�B���N�g�����u�b�N�̃p�X�ɐݒ肷��
    Call SetCurrentDirectory(ThisWorkbook.Path)
    ' �p�X������������
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    GetAbsolutePathNameFromThisWorkbookPath = fso.GetAbsolutePathName(workbookPath)

End Function

' -----------------------------------------------------------------------------
' �w�肳�ꂽ�T�u�t�H���_�[�����݂��邩�`�F�b�N���A
' ���݂��Ȃ��ꍇ�͍쐬���܂��B
' <IN> dirPath As String �`�F�b�N����t�H���_�̃p�X�B�_���p�X�̏ꍇ�̓J�����g�f�B���N�g���̃J�����g�t�H���_����ɂ��܂��B
' <OUT> Boolean �f�B���N�g�������݂��邩�A�쐬�ɐ��������ꍇ�� True�B�쐬�Ɏ��s�����ꍇ�� False�BTrue �̏ꍇ�ł��A�o�͂Ɏ��s����\�������邽�߁A�o�͎��̃G���[�`�F�b�N�͕K�����{���Ă��������B
' -----------------------------------------------------------------------------
Public Function TryMakeDir(dirPath As String) As Boolean
    
    Dim rtc As Long
    
    ' ���łɖړI�̃t�H���_�����邩
    If PathIsDirectory(dirPath) = True Then
        TryMakeDir = True
        Exit Function
    End If
    
    rtc = SHCreateDirectoryEx(0&, dirPath, 0&)
    
    ' ����ɍ쐬�ł����ꍇ NO_ERROR(0)
    ' �r�����t�@�C���ōċA�쐬�Ɏ��s�����ꍇ ERROR_PATH_NOT_FOUND(3)
    ' ���Ƀf�B���N�g��������ꍇ ERROR_ALREADY_EXISTS(183) (�ŏI�K�w���t�@�C���̏ꍇ�� ERROR_ALREADY_EXISTS �̂��߁A�������ɐ����Ƃ͂����Ȃ�)
    ' �������A���Y�t�H���_�Ƀt�@�C���̐����������邩�ǂ����́A�����ł̓`�F�b�N���Ă��Ȃ��B
    
    If rtc <> NO_ERROR Then
        ' log
    End If
    
    ' �ŏI�K�w���t�@�C���̏ꍇ�Ȃǂ�z�肵�āAAPI �ōŏI�`�F�b�N
    TryMakeDir = PathIsDirectory(dirPath)

End Function

#If ENABLE_TEST_METHODS = 1 Then

Public Sub RemoveExtensionTest()
    Debug.Print GetExtension("aaa.txt")
    Debug.Print RemoveExtension("aaa.txt")
    Debug.Print GetExtension("bbb")
    Debug.Print RemoveExtension("bbb")
    Debug.Print GetExtension(".\ccc.txt")
    Debug.Print RemoveExtension(".\ccc.txt")
    Debug.Print GetExtension(".\ddd")
    Debug.Print RemoveExtension(".\ddd")
End Sub

Public Sub mkdirtest()
    
    Debug.Print TryMakeDir(GetAbsolutePathNameFromThisWorkbookPath("log\sub1\sub2"))
    
End Sub

#End If


