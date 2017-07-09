Attribute VB_Name = "Module1"
Option Explicit
' -----------------------------------------------------------------------------
' ExcelChildProcessController
' https://github.com/Hondarer/ExcelChildProcessController
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

' -----------------------------------------------------------------------------
' �R�}���h�v�����v�g�̃T���v�� 1 �����s���܂��B
' -----------------------------------------------------------------------------
Public Sub Sample1()

    Dim func As New SampleFunction1
    
    Call func.Execute
    
End Sub

' -----------------------------------------------------------------------------
' �R�}���h�v�����v�g�̃T���v�� 2 �����s���܂��B
' -----------------------------------------------------------------------------
Public Sub Sample2()

    Dim func As New SampleFunction2
    
    Call func.Execute
    
End Sub

' -----------------------------------------------------------------------------
' �f�[�^�x�[�X����̒l�擾�ȈՃT���v�������s���܂��B
' -----------------------------------------------------------------------------
Public Sub SimpleQueryFunctionSample()

    Dim func As New SimpleQueryFunction
    
    ' �ݒ�l
    ' �V�[�g������Ȃ�萔�ŗ^����Ȃ肨�D����
    func.Username = "hr"
    func.Password = "tiger"
    func.Tlsname = "XE"
    
    ' �擾�����
'    Call func.AddColumns("ROWID")
    Call func.AddColumns("REGION_ID")
    Call func.AddColumns("REGION_NAME")
    
    ' from ��
    func.From = "REGIONS"
    
'    ' where ��
'    func.Where = "REGION_ID = 4"
    
    ' order by ��
    ' �����I�Ɏw�肵�Ȃ��ꍇ�́A�t���[�����[�N�ɂ�� ROWID �� order �����B
    ' �e�[�u���֐��Ȃ� ROWID �����݂��Ȃ��ꍇ�͕K���w�肷�邱�ƁB
    func.Orderby = "REGION_ID"
    
'    ' �^�C���A�E�g�̐ݒ�
'    ' �����I�Ɏw�肵�Ȃ��ꍇ�́A�t���[�����[�N�ɂ��
'    ' �f�t�H���g�̃^�C���A�E�g(60 �b)���w�肳���B�P�ʂ�[ms]�B
'    func.QueryTimeoutMilliSeconds = 60& * 1000&
    
    ' �₢���킹���s
    Call func.Execute
    
    ' �Z���ɏ����o���ɂ͈ȉ��̂悤�ɂ���
    ' �w�b�_�s�����邽�߁A�s�̐��̓f�[�^ + 1 �ƂȂ�A���� !
    
    Dim record As Long
    Dim column As Long
    
    For record = 0 To func.GetRecordsCount
        For column = 0 To func.GetColumnsCount - 1
            ThisWorkbook.Worksheets("Sheet1").Cells(record + 1, column + 1).Value = func.GetResult(record, column)
        Next
    Next
    
End Sub
