VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryFunctionBase"
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
' �f�[�^�x�[�X����̒l�擾�T���v���̊��N���X
'
' -----------------------------------------------------------------------------

Implements IProcessControlCallbacks

Private Const DEFAULT_COLUMN_DELIMITER = vbTab

Private Const QUERY_TIMEOUT = 60& * 1000& ' [ms]

Public controller As New ProcessController

Public isError As Boolean
Public errorDescription As String

Dim stateHandler_ As IStateHandler

Dim records As Long
Dim columns As Long
Dim realColumns As Long
Dim columnNames() As String
Dim dummyColumn() As Boolean

' ���ʂ�ێ����܂��B
Dim result() As Variant

Dim username_ As String
Dim password_ As String
Dim tlsname_ As String
Dim sql_ As String
Dim columnDelimiter_ As String

Dim queryTimeout_ As Long

Public Enum ExecutionPhase
    Ready = 0
    SqlSended = 1
    Recieveing = 2
End Enum

Private phase As ExecutionPhase

' -----------------------------------------------------------------------------
' ���̃I�u�W�F�N�g�� IProcessControlCallbacks �C���^�[�t�F�[�X���擾���܂��B
' -----------------------------------------------------------------------------
Property Get IProcessControlCallbacks_() As IProcessControlCallbacks
    Set IProcessControlCallbacks_ = Me
End Property

' -----------------------------------------------------------------------------
' ���̃I�u�W�F�N�g�̃X�^�[�^�X�ω����󂯎�� IStateHandler �C���^�[�t�F�[�X���擾���܂��B
' -----------------------------------------------------------------------------
Property Get StateHandler() As IStateHandler
    If Not stateHandler_ Is Nothing Then
        Set StateHandler = stateHandler_
    End If
End Property
' -----------------------------------------------------------------------------
' ���̃I�u�W�F�N�g�̃X�^�[�^�X�ω����󂯎�� IStateHandler �C���^�[�t�F�[�X��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let StateHandler(stateHandler__ As IStateHandler)
    Set stateHandler_ = stateHandler__
End Property

' -----------------------------------------------------------------------------
' ���[�U�[�����擾���܂��B
' -----------------------------------------------------------------------------
Property Get Username() As String
    Username = username_
End Property
' -----------------------------------------------------------------------------
' ���[�U�[����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let Username(username__ As String)
    username_ = username__
End Property

' -----------------------------------------------------------------------------
' �p�X���[�h��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let Password(password__ As String)
    password_ = password__
End Property

' -----------------------------------------------------------------------------
' �l�b�g�T�[�r�X�����擾���܂��B
' -----------------------------------------------------------------------------
Property Get Tlsname() As String
    Tlsname = tlsname_
End Property
' -----------------------------------------------------------------------------
' �l�b�g�T�[�r�X����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let Tlsname(tlsname__ As String)
    tlsname_ = tlsname__
End Property

' -----------------------------------------------------------------------------
' SQL ��������擾���܂��B
' -----------------------------------------------------------------------------
Property Get sql() As String
    sql = sql_
End Property
' -----------------------------------------------------------------------------
' SQL �������ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let sql(sql__ As String)
    sql_ = sql__
End Property

' -----------------------------------------------------------------------------
' ���؂蕶������擾���܂��B
' -----------------------------------------------------------------------------
Property Get ColumnDelimiter() As String
    ColumnDelimiter = columnDelimiter_
End Property
' -----------------------------------------------------------------------------
' ���؂蕶�����ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let ColumnDelimiter(columnDelimiter__ As String)
    columnDelimiter_ = columnDelimiter__
End Property

' -----------------------------------------------------------------------------
' �񐔂��擾���܂��B
' -----------------------------------------------------------------------------
Property Get ColumnCount() As Long
    ColumnCount = columns
End Property

' -----------------------------------------------------------------------------
' �₢���킹�^�C���A�E�g[ms]���擾���܂��B
' -----------------------------------------------------------------------------
Property Get QueryTimeoutMilliSeconds() As Long
    QueryTimeoutMilliSeconds = queryTimeout_
End Property
' -----------------------------------------------------------------------------
' �₢���킹�^�C���A�E�g[ms]��ݒ肵�܂��B
' -----------------------------------------------------------------------------
Property Let QueryTimeoutMilliSeconds(queryTimeout__ As Long)
    queryTimeout_ = queryTimeout__
End Property

' -----------------------------------------------------------------------------
' �V�����C���X�^���X�����������܂��B
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()

    ' ��̏�����
    Call ClearColumns
    
    ' �^�C���A�E�g�̃f�t�H���g�l��ݒ�
    queryTimeout_ = QUERY_TIMEOUT
    
    ' �f�t�H���g�̗��؂��ݒ�
    columnDelimiter_ = DEFAULT_COLUMN_DELIMITER
    
End Sub

' -----------------------------------------------------------------------------
' ���ǉ����܂��B
' -----------------------------------------------------------------------------
Public Sub AddColumns(columnName As String, isDummyColumn As Boolean)

    If columns = 0 Then
        columnNames(0) = columnName
        dummyColumn(0) = isDummyColumn
    Else
        ReDim Preserve columnNames(columns)
        columnNames(columns) = columnName
        ReDim Preserve dummyColumn(columns)
        dummyColumn(columns) = isDummyColumn
    End If
    
    columns = columns + 1
    
    If isDummyColumn = False Then
        realColumns = realColumns + 1
    End If

End Sub

' -----------------------------------------------------------------------------
' ������������܂��B
' -----------------------------------------------------------------------------
Public Sub ClearColumns()

    columns = 0
    realColumns = 0
    ReDim columnNames(0)
    ReDim dummyColumn(0)

End Sub

' -----------------------------------------------------------------------------
' �񖼂��擾���܂��B
' -----------------------------------------------------------------------------
Public Function GetColumnName(column As Long) As String
    GetColumnName = columnNames(column)
End Function

' -----------------------------------------------------------------------------
' �@�\�����s���܂��B
' -----------------------------------------------------------------------------
Public Function Execute() As Long
    Execute = ExecuteWithCallbackObject(Me)
End Function

' -----------------------------------------------------------------------------
' �R�[���o�b�N�Ώۂ̃I�u�W�F�N�g���w�肵�ċ@�\�����s���܂��B
' -----------------------------------------------------------------------------
Public Function ExecuteWithCallbackObject(callbackObject As IProcessControlCallbacks) As Long
    
    ' ��ʃf�[�^���Q�Ƃ���ۂɂ́A�f�o�b�O�o�͂ɂ�鐫�\�ቺ�������ł��Ȃ��̂�
    ' �f�o�b�O�o�͂𖳌��ɂ���
    controller.DisableIoTrace = True
    
    ' [username]/[passowrd]@[tlsname]
    ExecuteWithCallbackObject = controller.ExecuteProcess("sqlplus.exe " & username_ & "/" & password_ & "@" & tlsname_, callbackObject)
    
End Function

' -----------------------------------------------------------------------------
' �l��Ԃ��܂��B
' -----------------------------------------------------------------------------
Public Function GetResult() As Variant
    ' �V�[�g�ɒ���t���邱�Ƃ��l�������s��ɕϊ����Ă���Ԃ�
    GetResult = WorksheetFunction.Transpose(result)
End Function

' -----------------------------------------------------------------------------
' �l��1�Ԃ��܂��B
' -----------------------------------------------------------------------------
Public Function GetResultOne(record As Long, column As Long) As Variant
    GetResultOne = result(column, record)
End Function

' -----------------------------------------------------------------------------
' �s����Ԃ��܂��B�񖼂������s�͊܂܂�܂���B
' -----------------------------------------------------------------------------
Public Function GetRecordsCount() As Long
    GetRecordsCount = records
End Function

' -----------------------------------------------------------------------------
' �񐔂�Ԃ��܂��B
' -----------------------------------------------------------------------------
Public Function GetColumnsCount() As Long
    GetColumnsCount = columns
End Function

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X�Ɛڑ����ꂽ�ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnConnected(sender As ProcessController)

    ' �񖼂����ʂɒ�������
    Dim column As Long
    ReDim Preserve result(columns - 1, 0)
    For column = 0 To columns - 1
        result(column, 0) = columnNames(column)
    Next

    ' �^�C���A�E�g��ݒ�
    Call sender.SetDeepIdleTimeoutMilliseconds(queryTimeout_)

End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X�Ƃ̒ʐM���������ԍs���Ȃ������ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnDeepIdle(sender As ProcessController)
    Call sender.RequestTerminate(EXIT_FAILURE)
    Call MsgBox("Oracle ���삪�^�C���A�E�g���܂����B", Buttons:=vbCritical, Title:="�G���[")
End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X���I�������ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnDisconnected(sender As ProcessController)
End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X���Z���ԒʐM�����Ȃ������ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnIdle(sender As ProcessController, lineString As String)

    If Left(lineString, 5) = "SQL> " Then
        Call IProcessControlCallbacks_OnRecieveLineFromStdout(sender, lineString)
    End If

End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X�̕W���G���[�o�͂� 1 �s�o�͂��ꂽ�ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnRecieveLineFromStderr(sender As ProcessController, recvErrString As String)
End Sub

' -----------------------------------------------------------------------------
' �Ώۂ̃v���Z�X�̕W���o�͂� 1 �s�o�͂��ꂽ�ۂɔ������܂��B
' -----------------------------------------------------------------------------
Private Sub IProcessControlCallbacks_OnRecieveLineFromStdout(sender As ProcessController, recvOutString As String)

    ' �G���[�����������ꍇ�͂��̓��e��ێ����āA�v���O�������I��������
    If (recvOutString = "ERROR:") Or (Right(recvOutString, 12) = "�G���[���������܂����B:") Then
        isError = True
        sender.RequestTerminate EXIT_FAILURE
    ElseIf isError = True And errorDescription = "" Then
        errorDescription = recvOutString
        GetDefaultLogger.LogError "[QueryFunctionBase][IProcessControlCallbacks_OnRecieveLineFromStdout] Oracle �ɂăG���[���������܂����B" & errorDescription
        Call MsgBox("Oracle �ɂăG���[���������܂����B" & vbCrLf & errorDescription, Buttons:=vbCritical, Title:="�G���[")
        
        ' �f�[�^���폜����
        ReDim Preserve result(columns - 1, 0)
        records = 0
        
        Exit Sub
    ElseIf isError = True Then
        Exit Sub
    End If

    ' �v�����v�g���󂯂邱�Ƃɂ��t�F�[�Y��i�߂�
    If Left(recvOutString, 5) = "SQL> " Then
        Select Case phase
        
            Case Ready
            
                If Not StateHandler Is Nothing Then
                    Call StateHandler.OnStatusChanged("SQL �𑗐M���Ă��܂�...")
                End If
                
                ' ����̗\��
                sender.SendLineToStdin "set linesize 32767"
                sender.SendLineToStdin "set head off"
                sender.SendLineToStdin "set newpage 0"
                sender.SendLineToStdin "set feedback off"
                
                GetDefaultLogger.LogInfo "[QueryFunctionBase][IProcessControlCallbacks_OnRecieveLineFromStdout] sql=" & sql
                
                sender.SendLineToStdin sql
                sender.SendLineToStdin "exit"
                
                If Not StateHandler Is Nothing Then
                    Call StateHandler.OnStatusChanged("�₢���킹���ʂ�ҋ@���Ă��܂�...")
                End If
                
                phase = SqlSended
                
            Case SqlSended
                phase = Recieveing
                
        End Select
    End If
    
    ' �f�[�^��M���̏���
    If phase = Recieveing Then
    
        Dim column As Long
        Dim data() As String
            
        ' "set newpage 0" �ɂ��t�������A���y�[�W�̍폜
        ' &hC(Form Feed(���y�[�W))�����o���Ă����܂Ŏ�菜��
        ' ���y�[�W�̓f�[�^�̊J�n����ɗ��p���Ă��邽�߁A�ӎ��I�ɕt�����Ă���
        If InStr(recvOutString, Chr(&HC)) > 0 Then
            recvOutString = Mid(recvOutString, InStr(recvOutString, Chr(&HC)) + 1)
        End If
        
        data = Split(recvOutString, columnDelimiter_)
        
        ' �G���[������Ȃǂł͂Ȃ��A�f�[�^���ł��邱��
        ' (��؂�ꂽ�v�f�̌����\�z�����񐔂ƈ�v���Ă��邱��)
        If UBound(data) = (realColumns - 1) Then
        
'           GetDefaultLogger.LogDebug "> " & data(0) & " / " & data(1)

            records = records + 1
            ReDim Preserve result(columns - 1, records)
    
            If Not StateHandler Is Nothing Then
                If (records Mod 1000) = 1 Then
                    Call StateHandler.OnStatusChanged("�₢���킹���ʂ���M���Ă��܂�...(" & records & ")")
                End If
            End If
            
            Dim datacolumn As Long
            
            For column = 0 To columns - 1
                If dummyColumn(column) = False Then
                    result(column, records) = data(datacolumn)
                    datacolumn = datacolumn + 1
                End If
            Next
        
        End If
        
    End If
    
End Sub


