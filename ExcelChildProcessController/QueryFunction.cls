VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'
' �f�[�^�x�[�X����̒l�擾�T���v��
'
' VBA �ɂ͌p���̊T�O���������߁A����ē����̏������s���B
'
' -----------------------------------------------------------------------------

Implements IProcessControlCallbacks

Dim base As QueryFunctionBase

Dim from_ As String
Dim where_ As String
Dim orderby_ As String

' -----------------------------------------------------------------------------
' ���̃I�u�W�F�N�g�̃X�^�[�^�X�ω����󂯎�� IStateHandler �C���^�[�t�F�[�X���擾���܂��B
' -----------------------------------------------------------------------------
Property Get StateHandler() As IStateHandler
    Set StateHandler = base.StateHandler
End Property
' -----------------------------------------------------------------------------
' ���̃I�u�W�F�N�g�̃X�^�[�^�X�ω����󂯎�� IStateHandler �C���^�[�t�F�[�X��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let StateHandler(stateHandler__ As IStateHandler)
    base.StateHandler = stateHandler__
End Property

' -----------------------------------------------------------------------------
' ���[�U�[�����擾���܂��B
' -----------------------------------------------------------------------------
Property Get Username() As String
    Username = base.Username
End Property
' -----------------------------------------------------------------------------
' ���[�U�[����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let Username(username__ As String)
    base.Username = username__
End Property

' -----------------------------------------------------------------------------
' �p�X���[�h��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let Password(password__ As String)
    base.Password = password__
End Property

' -----------------------------------------------------------------------------
' �l�b�g�T�[�r�X�����擾���܂��B
' -----------------------------------------------------------------------------
Property Get Tlsname() As String
    Tlsname = base.Tlsname
End Property
' -----------------------------------------------------------------------------
' �l�b�g�T�[�r�X����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let Tlsname(tlsname__ As String)
    base.Tlsname = tlsname__
End Property

' -----------------------------------------------------------------------------
' from ����擾���܂��B
' -----------------------------------------------------------------------------
Property Get From() As String
    From = from_
End Property
' -----------------------------------------------------------------------------
' from ���ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let From(from__ As String)
    from_ = from__
End Property

' -----------------------------------------------------------------------------
' where ����擾���܂��B
' -----------------------------------------------------------------------------
Property Get Where() As String
    Where = where_
End Property
' -----------------------------------------------------------------------------
' where ���ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let Where(where__ As String)
    where_ = where__
End Property

' -----------------------------------------------------------------------------
' order by ����擾���܂��B
' -----------------------------------------------------------------------------
Property Get Orderby() As String
    Orderby = orderby_
End Property
' -----------------------------------------------------------------------------
' order by ���ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let Orderby(orderby__ As String)
    orderby_ = orderby__
End Property

' -----------------------------------------------------------------------------
' �񐔂��擾���܂��B
' -----------------------------------------------------------------------------
Property Get ColumnCount() As Long
    ColumnCount = base.ColumnCount()
End Property

' -----------------------------------------------------------------------------
' �₢���킹�^�C���A�E�g[ms]���擾���܂��B
' -----------------------------------------------------------------------------
Property Get QueryTimeoutMilliSeconds() As Long
    QueryTimeoutMilliSeconds = base.QueryTimeoutMilliSeconds
End Property
' -----------------------------------------------------------------------------
' �₢���킹�^�C���A�E�g[ms]��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let QueryTimeoutMilliSeconds(queryTimeout__ As Long)
    base.QueryTimeoutMilliSeconds = queryTimeout__
End Property

' -----------------------------------------------------------------------------
' �V�����C���X�^���X�����������܂��B
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set base = New QueryFunctionBase
End Sub

' -----------------------------------------------------------------------------
' ���ǉ����܂��B
' -----------------------------------------------------------------------------
Public Sub AddColumns(columnName As String, isDummyColumn As Boolean)
    Call base.AddColumns(columnName, isDummyColumn)
End Sub

' -----------------------------------------------------------------------------
' ������������܂��B
' -----------------------------------------------------------------------------
Public Sub ClearColumns()
    Call base.ClearColumns
End Sub

' -----------------------------------------------------------------------------
' �񖼂��擾���܂��B
' -----------------------------------------------------------------------------
Public Function GetColumnName(column As Long) As String
    GetColumnName = base.GetColumnName(column)
End Function

' -----------------------------------------------------------------------------
' �@�\�����s���܂��B
' -----------------------------------------------------------------------------
Public Function Execute() As Long
    Execute = base.ExecuteWithCallbackObject(Me)
End Function

' -----------------------------------------------------------------------------
' �l��Ԃ��܂��B
' -----------------------------------------------------------------------------
Public Function GetResult() As Variant
    GetResult = base.GetResult
End Function

' -----------------------------------------------------------------------------
' �l��1�Ԃ��܂��B
' -----------------------------------------------------------------------------
Public Function GetResultOne(record As Long, column As Long) As Variant
    GetResultOne = base.GetResult(record, column)
End Function

' -----------------------------------------------------------------------------
' �s����Ԃ��܂��B�񖼂������s�͊܂܂�܂���B
' -----------------------------------------------------------------------------
Public Function GetRecordsCount() As Long
    GetRecordsCount = base.GetRecordsCount
End Function

' -----------------------------------------------------------------------------
' �񐔂�Ԃ��܂��B
' -----------------------------------------------------------------------------
Public Function GetColumnsCount() As Long
    GetColumnsCount = base.GetColumnsCount
End Function

' -----------------------------------------------------------------------------
' SQL ������𐶐����܂��B
' -----------------------------------------------------------------------------
Private Function GetSqlSelectString() As String
    
    Dim column As Long
    
    GetSqlSelectString = "select "
    
    For column = 0 To base.GetColumnsCount - 1
    
        GetSqlSelectString = GetSqlSelectString & base.GetColumnName(column)
        
        If column < (base.GetColumnsCount - 1) Then
            GetSqlSelectString = GetSqlSelectString & " || ',' || "
        End If
        
    Next

    GetSqlSelectString = GetSqlSelectString & " from " & from_
        
    If where_ <> "" Then
        GetSqlSelectString = GetSqlSelectString & " where " & where_
    End If

    If orderby_ <> "" Then
        GetSqlSelectString = GetSqlSelectString & " order by " & orderby_
    Else
        GetSqlSelectString = GetSqlSelectString & " order by ROWID"
    End If
    
    GetSqlSelectString = GetSqlSelectString & ";"
    
End Function

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X�Ɛڑ����ꂽ�ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnConnected(sender As ProcessController)
    
    base.sql = GetSqlSelectString()
    
    Call base.IProcessControlCallbacks_.OnConnected(sender)

End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X�Ƃ̒ʐM���������ԍs���Ȃ������ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnDeepIdle(sender As ProcessController)
    Call base.IProcessControlCallbacks_.OnDeepIdle(sender)
End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X���I�������ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnDisconnected(sender As ProcessController)
    Call base.IProcessControlCallbacks_.OnDisconnected(sender)
End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X���Z���ԒʐM�����Ȃ������ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnIdle(sender As ProcessController, lineString As String)
    Call base.IProcessControlCallbacks_.OnIdle(sender, lineString)
End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X�̕W���G���[�o�͂� 1 �s�o�͂��ꂽ�ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnRecieveLineFromStderr(sender As ProcessController, recvErrString As String)
    Call base.IProcessControlCallbacks_.OnRecieveLineFromStderr(sender, recvErrString)
End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X�̕W���o�͂� 1 �s�o�͂��ꂽ�ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnRecieveLineFromStdout(sender As ProcessController, recvOutString As String)
    Call base.IProcessControlCallbacks_.OnRecieveLineFromStdout(sender, recvOutString)
End Sub



