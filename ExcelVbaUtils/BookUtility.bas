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

Private Const LONG_MAXVALUE = 2147483647

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
    Dim seqedName As String
    
    If InStr(fileName, ".") = 0 Then
        Exit Function
    End If

    ' �p�X���ȗ������ꍇ�̏���
    If filePath = "" Then
        filePath = ThisWorkbook.Path
    End If

    If FolderExists(filePath & "\" & fileName) = True Then
        ' �����̃t�H���_�����݂���
        msgboxResult = MsgBox("���̏ꏊ�� '" & filePath & "\" & fileName & "' �Ƃ������O�̃t�H���_�����ɂ���܂��B���O��ύX���ĕۑ����܂���?", vbOKCancel Or vbInformation)
        If msgboxResult = vbOK Then
        
            ' ���j�[�N�Ȗ��̂�����
            seqedName = FindUniqueFileName(filePath, fileName)
            If seqedName = "" Then
                ' �����Ɏ��s
                Call book.Close(False)
                Exit Function
            End If
            
            ' �ۑ�����
            If SaveAsWorkBookCore(book, filePath, seqedName, True) <> True Then
                ' �ۑ��Ɏ��s
                Call book.Close(False)
                Exit Function
            End If
        Else
            ' �L�����Z�����ꂽ
            Call book.Close(False)
            Exit Function
        End If
    ElseIf FileExists(filePath & "\" & fileName) = True Then
        ' �t�@�C�������݂���
        msgboxResult = MsgBox("���̏ꏊ�� '" & filePath & "\" & fileName & "' �Ƃ������O�̃t�@�C�������ɂ���܂��B�u�������܂���?", vbYesNoCancel Or vbInformation)
        If msgboxResult = vbYes Then
            
            ' �u�������悤�Ƃ������N�����J���Ă��铙
            If SaveAsWorkBookCore(book, filePath, fileName, False) <> True Then
                msgboxResult = MsgBox("'" & filePath & "\" & fileName & "' �̕ۑ��Ɏ��s���܂����B���O��ύX���ĕۑ����܂���?", vbOKCancel Or vbInformation)
                If msgboxResult = vbOK Then
                    
                    ' ���j�[�N�Ȗ��̂�����
                    seqedName = FindUniqueFileName(filePath, fileName)
                    If seqedName = "" Then
                        ' �����Ɏ��s
                        Call book.Close(False)
                        Exit Function
                    End If
                
                    ' �ۑ�����
                    If SaveAsWorkBookCore(book, filePath, seqedName, True) <> True Then
                        ' �ۑ��Ɏ��s
                        Call book.Close(False)
                        Exit Function
                    End If
                Else
                    ' �L�����Z�����ꂽ
                    Call book.Close(False)
                    Exit Function
                End If
            End If
            
        ElseIf msgboxResult = vbNo Then
            
            ' ���j�[�N�Ȗ��̂�����
            seqedName = FindUniqueFileName(filePath, fileName)
            If seqedName = "" Then
                ' �����Ɏ��s
                Call book.Close(False)
                Exit Function
            End If
        
            ' �ۑ�����
            If SaveAsWorkBookCore(book, filePath, seqedName, True) <> True Then
                ' �ۑ��Ɏ��s
                Call book.Close(False)
                Exit Function
            End If
        Else
            ' �L�����Z�����ꂽ
            Call book.Close(False)
            Exit Function
        End If
    Else
        ' �t�@�C�����d���`�F�b�N OK �̏ꍇ
        ' �ۑ�����
        If SaveAsWorkBookCore(book, filePath, fileName, True) <> True Then
            ' �ۑ��Ɏ��s
            Call book.Close(False)
            Exit Function
        End If
    End If

    SaveAsWorkBook = book.FullName
    
    Call book.Close

End Function

' -----------------------------------------------------------------------------
' �t�^�\�ȃt�@�C�������������ԋp���܂��B
' <IN> filePath As String �ۑ�����u�b�N�̃p�X�B�󕶎��̏ꍇ�́A���̃}�N�������삵�Ă���u�b�N�̃p�X�B
' <IN> fileName As String �g���q���܂ށA�ۑ�����u�b�N�̃t�@�C�����B
' <OUT> String �t�@�C�����̌����ɐ��������ꍇ�A�g���q���܂ށA�ۑ�����u�b�N�̃t�@�C�����B���s�����ꍇ�A�󕶎��B
' -----------------------------------------------------------------------------
Private Function FindUniqueFileName(filePath As String, fileName As String) As String
    
    Dim seq As Long '�d�����p�̒ʔ�
    Dim baseName As String
    Dim seqedName As String
    Dim extension As String
    
    baseName = RemoveExtension(fileName)
    extension = GetExtension(fileName)

    Do
        If seq = 0 Then
            seqedName = baseName
        Else
            seqedName = baseName & "(" & seq & ")"
        End If
        
        If (Not FolderExists(filePath & "\" & seqedName & extension)) And _
           (Not FileExists(filePath & "\" & seqedName & extension)) And _
           (Not WorkbookExists(seqedName & extension)) Then
            ' ���j�[�N�Ȗ��̂���������
            Exit Do
        End If
        
        If seq = LONG_MAXVALUE Then
            Call MsgBox("'" & filePath & "\" & fileName & "' �̕ۑ��Ɏ��s���܂����B" & vbCrLf & _
                        "�ʔԂ��ő�l�ɒB���܂����B", vbOKOnly Or vbExclamation)
            Exit Function
        End If
        
        seq = seq + 1
    Loop
    
    FindUniqueFileName = seqedName & extension

End Function

' -----------------------------------------------------------------------------
' ���̃v���Z�X�Ŏw�肳�ꂽ�u�b�N�������łɊJ����Ă��邩�ǂ�����Ԃ��܂��B
' <IN> workbookName As String �`�F�b�N����u�b�N���B
' <OUT> Boolean ���Ƀu�b�N���J����Ă���ꍇ�ATrue�B�J����Ă��Ȃ��ꍇ�AFalse�B
' -----------------------------------------------------------------------------
Public Function WorkbookExists(workbookName As String) As Boolean
    
    Dim book As Workbook
    
    For Each book In Application.Workbooks
        If book.Name = workbookName Then
            WorkbookExists = True
            Exit Function
        End If
    Next
    
    WorkbookExists = False
    
End Function

' -----------------------------------------------------------------------------
' �u�b�N�ɖ��O�����ĕۑ����܂��B
' <IN> book As Workbook �ۑ�����u�b�N�B
' <IN> filePath As String �ۑ�����u�b�N�̃p�X�B�󕶎��̏ꍇ�́A���̃}�N�������삵�Ă���u�b�N�̃p�X�B
' <IN> fileName As String �g���q���܂ށA�ۑ�����u�b�N�̃t�@�C�����B
' <IN> showDialog As Boolean �ۑ��Ɏ��s�����ۂɃ_�C�A���O��\�����邩�ǂ����B
' <OUT> Boolean �ۑ��ɐ��������ꍇ�ATrue�B���s�����ꍇ�AFalse�B
' -----------------------------------------------------------------------------
Private Function SaveAsWorkBookCore(book As Workbook, filePath As String, fileName As String, showDialog As Boolean) As Boolean

    Dim lastErr As Long
    Dim lastErrDescription As String

    ' �u�b�N��ۑ�����
    Application.DisplayAlerts = False
    On Error Resume Next
    Err = 0
    Call book.SaveAs(filePath & "\" & fileName)
    lastErr = Err
    lastErrDescription = Err.description
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' �ۑ��Ɏ��s������
    If lastErr <> 0 Then
    
        ' ���s�����̂Ń_�C�A���O��\��
        If showDialog = True Then
            Call MsgBox("'" & filePath & "\" & fileName & "' �̕ۑ��Ɏ��s���܂����B" & vbCrLf & vbCrLf & _
                        lastErrDescription, vbOKOnly Or vbExclamation)
        End If
        
        SaveAsWorkBookCore = False
        
    Else
    
        ' �ۑ��ɐ���
        SaveAsWorkBookCore = True
        
    End If

End Function

#If ENABLE_TEST_METHODS = 1 Then

' -----------------------------------------------------------------------------
' SaveAsWorkBook ���\�b�h�̃e�X�g���s���܂��B
' -----------------------------------------------------------------------------
Public Sub SaveAsWorkBookTest()

    Dim newBook As Workbook
    Set newBook = Workbooks.Add(xlWBATWorksheet)
    
    Call SaveAsWorkBook(newBook, "", "Test.xlsx")
    
    Set newBook = Nothing
    
End Sub

#End If

