Attribute VB_Name = "FileDialog"
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

' Dependency: FileUtility

#Const ENABLE_TEST_METHODS = 1

' reference:
' Excel �t�@�C���p�X�A�t�H���_�p�X��I�����Z���Ɋi�[����Q�ƃ{�^��
' http://qiita.com/boss_ape/items/1733fe6317e4566fdebb

' -----------------------------------------------------------------------------
' �t�H���_�̑I���_�C�A���O��\�����܂��B
' �����t�H���_�����݂��Ȃ��ꍇ�́A�w�肳�ꂽ�X�y�V�����t�H���_�������t�H���_�Ƃ��ĕ\�����܂��B
' <IN> defaultPath As String �����t�H���_�B�w�肵�Ȃ��ꍇ�͋󕶎����w�肵�܂��B
' <IN> fallbackSpecialFolder As String SPECIALFOLDERS_ �Ŏn�܂�X�y�V�����t�H���_�̎��ʎq�BdefaultPath ���󔒂܂��͖����ȏꍇ�ɍ̗p����܂��B
' <OUT> String �I�����ꂽ�t�H���_�B�L�����Z�����ꂽ�ꍇ�͋󕶎���Ԃ��܂��B
' -----------------------------------------------------------------------------
Public Function SelectFolderWithDialog(defaultPath As String, fallbackSpecialFolder As String) As String
    
    Dim ofdFolderDlg As Office.FileDialog
    Dim openPath As String

    ' �����t�H���_�̐ݒ�
    If Len(defaultPath) > 0 Then
        ' ������ "\" �폜
        If Right(defaultPath, 1) = "\" Then
            openPath = Left(defaultPath, Len(defaultPath) - 1)
        Else
            openPath = defaultPath
        End If

        ' �t�H���_���݃`�F�b�N
        If Not FolderExists(openPath) Then
            openPath = GetSpecialFolder(fallbackSpecialFolder)
        End If
    Else
        openPath = GetSpecialFolder(fallbackSpecialFolder)
    End If

    ' �t�H���_�I���_�C�A���O�ݒ�
    Set ofdFolderDlg = Application.FileDialog(msoFileDialogFolderPicker)
    With ofdFolderDlg
        ' �\������A�C�R���̑傫�����w��
        .InitialView = msoFileDialogViewDetails
        ' �t�H���_�����ʒu
        .InitialFileName = openPath & "\"
        ' �����I��s��
        .AllowMultiSelect = False
    End With

    ' �t�H���_�I���_�C�A���O�\��
    If ofdFolderDlg.Show() = -1 Then
        ' �t�H���_�p�X�ݒ�
        SelectFolderWithDialog = ofdFolderDlg.SelectedItems(1)
    End If

    Set ofdFolderDlg = Nothing
    
End Function

#If ENABLE_TEST_METHODS = 1 Then

' -----------------------------------------------------------------------------
' SelectFolderWithDialog ���\�b�h�̃e�X�g���s���܂��B
' -----------------------------------------------------------------------------
Public Sub SelectFolderWithDialogTest()

    Debug.Print "SelectFolderWithDialog=" & SelectFolderWithDialog("", SPECIALFOLDERS_MYDOCUMENTS)
    
End Sub

#End If
