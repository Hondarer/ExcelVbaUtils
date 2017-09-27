Attribute VB_Name = "BookUtility"
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

Public Const WORKBOOK_EXT = ".xlsx"
Public Const WORKBOOKWITHMACRO_EXT = ".xlsm"

' -----------------------------------------------------------------------------
' �ʂ̃v���Z�X�Ńu�b�N���J���܂��B
' <IN> filePath As String �J���u�b�N�̃p�X�B
' <OUT> Boolean �����ɐ��������ꍇ�ATrue�B���s�����ꍇ�AFalse�B
' -----------------------------------------------------------------------------
Public Function OpenWorkbookAsNewProcess(bookPath As String) As Boolean

    Dim excl As Object
    Dim lastErr As Long
    
    If FileExists(bookPath) = False Then
        OpenWorkbookAsNewProcess = False
        Exit Function
    End If
    
    Set excl = CreateObject("Excel.Application")
    On Error Resume Next
    Err = 0
    excl.Workbooks.Open fileName:=bookPath, ReadOnly:=True
    lastErr = Err
    On Error GoTo 0
    
    If lastErr = 0 Then
        excl.Visible = True
        OpenWorkbookAsNewProcess = True
    Else
        excl.Quit
        OpenWorkbookAsNewProcess = False
    End If
    
    Set excl = Nothing
    
End Function

' -----------------------------------------------------------------------------
' �u�b�N�ɖ��O�����ĕۑ����܂��B
' �ۑ��ł��Ȃ��ꍇ�ɂ͖₢���킹���s���A�ʔԂ�t�^�����K�؂Ȗ��̂ŕۑ����܂��B
' <IN> book As Workbook �ۑ�����u�b�N�B
' <IN> filePath As String �ۑ�����u�b�N�̃p�X�B�󕶎��̏ꍇ�́A���̃}�N�������삵�Ă���u�b�N�̃p�X�B
' <IN> fileName As String �g���q���܂ށA�ۑ�����u�b�N�̃t�@�C�����B
' <OUT> String �ۑ��ɐ��������ꍇ�A���̃u�b�N�̃t���p�X�B���s�����ꍇ�A�󕶎��B
' -----------------------------------------------------------------------------
Public Function SaveAsWorkBook(book As Workbook, filePath As String, fileName As String) As String

    Dim msgboxResult As VbMsgBoxResult
    Dim seq As Long '�d�����p�̒ʔ�
    Dim seqedName As String
    Dim lastErr As Long
    
    Dim extension As String
    
    If InStr(fileName, ".") = 0 Then
        Exit Function
    End If

    extension = Mid(fileName, InStrRev(fileName, "."))
    fileName = Left(fileName, Len(fileName) - Len(extension))

    ' �p�X���ȗ������ꍇ�̏���
    If filePath = "" Then
        filePath = ThisWorkbook.Path
    End If

    If FolderExists(filePath & "\" & fileName & extension) = True Then
        ' �����̃t�H���_�����݂���
        msgboxResult = MsgBox("���̏ꏊ�� '" & filePath & "\" & fileName & extension & "' �Ƃ������O�̃t�H���_�����ɂ���܂��B���O��ύX���ĕۑ����܂���?", vbOKCancel Or vbInformation)
        If msgboxResult = vbOK Then
            ' ���j�[�N�Ȗ��̂�����
            Do
                seq = seq + 1
                seqedName = fileName & "(" & seq & ")"
                If (Not FolderExists(filePath & "\" & seqedName & extension)) And (Not FileExists(filePath & "\" & seqedName & extension)) Then
                    ' ���j�[�N�Ȗ��̂���������
                    Exit Do
                End If
            Loop
            Call book.SaveAs(filePath & "\" & seqedName & extension)
        Else
            ' �L�����Z�����ꂽ
            Call book.Close(False)
            Exit Function
        End If
    ElseIf FileExists(filePath & "\" & fileName & extension) = True Then
        ' �t�@�C�������݂���
        msgboxResult = MsgBox("���̏ꏊ�� '" & filePath & "\" & fileName & extension & "' �Ƃ������O�̃t�@�C�������ɂ���܂��B�u�������܂���?", vbYesNoCancel Or vbInformation)
        If msgboxResult = vbYes Then
            ' �u������
            Application.DisplayAlerts = False
            
            On Error Resume Next
            Err = 0
            Call book.SaveAs(filePath & "\" & fileName & extension)
            lastErr = Err
            On Error GoTo 0
            Application.DisplayAlerts = True
            
            ' �u�������悤�Ƃ������N�����J���Ă��铙
            If lastErr <> 0 Then
                msgboxResult = MsgBox("'" & filePath & "\" & fileName & extension & "' �̕ۑ��Ɏ��s���܂����B���O��ύX���ĕۑ����܂���?", vbOKCancel Or vbInformation)
                If msgboxResult = vbOK Then
                    ' ���j�[�N�Ȗ��̂�����
                    Do
                        seq = seq + 1
                        seqedName = fileName & "(" & seq & ")"
                        If (Not FolderExists(filePath & "\" & seqedName & extension)) And (Not FileExists(filePath & "\" & seqedName & extension)) Then
                            ' ���j�[�N�Ȗ��̂���������
                            Exit Do
                        End If
                    Loop
                    Call book.SaveAs(filePath & "\" & seqedName & extension)
                Else
                    ' �L�����Z�����ꂽ
                    Call book.Close(False)
                    Exit Function
                End If
            End If
            
        ElseIf msgboxResult = vbNo Then
            ' ���j�[�N�Ȗ��̂�����
            Do
                seq = seq + 1
                seqedName = fileName & "(" & seq & ")"
                If (Not FolderExists(filePath & "\" & seqedName & extension)) And (Not FileExists(filePath & "\" & seqedName & extension)) Then
                    ' ���j�[�N�Ȗ��̂���������
                    Exit Do
                End If
            Loop
            Call book.SaveAs(filePath & "\" & seqedName & extension)
        Else
            ' �L�����Z�����ꂽ
            Call book.Close(False)
            Exit Function
        End If
    Else
        ' ����P�[�X
        Call book.SaveAs(filePath & "\" & fileName & extension)
    End If

    SaveAsWorkBook = book.FullName
    
    Call book.Close

End Function

#If ENABLE_TEST_METHODS = 1 Then

' -----------------------------------------------------------------------------
' SaveAsWorkBook ���\�b�h�̃e�X�g���s���܂��B
' -----------------------------------------------------------------------------
Public Sub SaveAsWorkBookTest()

    Dim newBook As Workbook
    Set newBook = Workbooks.Add(xlWBATWorksheet)
    
    Call SaveAsWorkBook(newBook, "", "Test")
    
    Set newBook = Nothing
    
End Sub

#End If

