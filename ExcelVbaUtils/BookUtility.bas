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
' 以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル
' （以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを
' 無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、
' 変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、
' およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
'
' 上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な
' 部分に記載するものとします。
'
' ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、
' 何らの保証もなく提供されます。
' ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての
' 保証も含みますが、それに限定されるものではありません。
' 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、
' ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の
' 扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わない
' ものとします。
'
' -----------------------------------------------------------------------------

' Dependency: FileUtility

#Const ENABLE_TEST_METHODS = 1

Public Const WORKBOOK_EXT = ".xlsx"
Public Const WORKBOOKWITHMACRO_EXT = ".xlsm"

' -----------------------------------------------------------------------------
' 別のプロセスでブックを開きます。
' <IN> filePath As String 開くブックのパス。
' <OUT> Boolean 処理に成功した場合、True。失敗した場合、False。
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
' ブックに名前をつけて保存します。
' 保存できない場合には問い合わせを行い、通番を付与した適切な名称で保存します。
' <IN> book As Workbook 保存するブック。
' <IN> filePath As String 保存するブックのパス。空文字の場合は、このマクロが動作しているブックのパス。
' <IN> fileName As String 拡張子を含む、保存するブックのファイル名。
' <OUT> String 保存に成功した場合、そのブックのフルパス。失敗した場合、空文字。
' -----------------------------------------------------------------------------
Public Function SaveAsWorkBook(book As Workbook, filePath As String, fileName As String) As String

    Dim msgboxResult As VbMsgBoxResult
    Dim seq As Long '重複回避用の通番
    Dim seqedName As String
    Dim lastErr As Long
    
    Dim extension As String
    
    If InStr(fileName, ".") = 0 Then
        Exit Function
    End If

    extension = Mid(fileName, InStrRev(fileName, "."))
    fileName = Left(fileName, Len(fileName) - Len(extension))

    ' パスを省略した場合の処理
    If filePath = "" Then
        filePath = ThisWorkbook.Path
    End If

    If FolderExists(filePath & "\" & fileName & extension) = True Then
        ' 同名のフォルダが存在する
        msgboxResult = MsgBox("この場所に '" & filePath & "\" & fileName & extension & "' という名前のフォルダが既にあります。名前を変更して保存しますか?", vbOKCancel Or vbInformation)
        If msgboxResult = vbOK Then
            ' ユニークな名称を検索
            Do
                seq = seq + 1
                seqedName = fileName & "(" & seq & ")"
                If (Not FolderExists(filePath & "\" & seqedName & extension)) And (Not FileExists(filePath & "\" & seqedName & extension)) Then
                    ' ユニークな名称が見つかった
                    Exit Do
                End If
            Loop
            Call book.SaveAs(filePath & "\" & seqedName & extension)
        Else
            ' キャンセルされた
            Call book.Close(False)
            Exit Function
        End If
    ElseIf FileExists(filePath & "\" & fileName & extension) = True Then
        ' ファイルが存在する
        msgboxResult = MsgBox("この場所に '" & filePath & "\" & fileName & extension & "' という名前のファイルが既にあります。置き換えますか?", vbYesNoCancel Or vbInformation)
        If msgboxResult = vbYes Then
            ' 置き換え
            Application.DisplayAlerts = False
            
            On Error Resume Next
            Err = 0
            Call book.SaveAs(filePath & "\" & fileName & extension)
            lastErr = Err
            On Error GoTo 0
            Application.DisplayAlerts = True
            
            ' 置き換えようとしたが誰かが開いている等
            If lastErr <> 0 Then
                msgboxResult = MsgBox("'" & filePath & "\" & fileName & extension & "' の保存に失敗しました。名前を変更して保存しますか?", vbOKCancel Or vbInformation)
                If msgboxResult = vbOK Then
                    ' ユニークな名称を検索
                    Do
                        seq = seq + 1
                        seqedName = fileName & "(" & seq & ")"
                        If (Not FolderExists(filePath & "\" & seqedName & extension)) And (Not FileExists(filePath & "\" & seqedName & extension)) Then
                            ' ユニークな名称が見つかった
                            Exit Do
                        End If
                    Loop
                    Call book.SaveAs(filePath & "\" & seqedName & extension)
                Else
                    ' キャンセルされた
                    Call book.Close(False)
                    Exit Function
                End If
            End If
            
        ElseIf msgboxResult = vbNo Then
            ' ユニークな名称を検索
            Do
                seq = seq + 1
                seqedName = fileName & "(" & seq & ")"
                If (Not FolderExists(filePath & "\" & seqedName & extension)) And (Not FileExists(filePath & "\" & seqedName & extension)) Then
                    ' ユニークな名称が見つかった
                    Exit Do
                End If
            Loop
            Call book.SaveAs(filePath & "\" & seqedName & extension)
        Else
            ' キャンセルされた
            Call book.Close(False)
            Exit Function
        End If
    Else
        ' 正常ケース
        Call book.SaveAs(filePath & "\" & fileName & extension)
    End If

    SaveAsWorkBook = book.FullName
    
    Call book.Close

End Function

#If ENABLE_TEST_METHODS = 1 Then

' -----------------------------------------------------------------------------
' SaveAsWorkBook メソッドのテストを行います。
' -----------------------------------------------------------------------------
Public Sub SaveAsWorkBookTest()

    Dim newBook As Workbook
    Set newBook = Workbooks.Add(xlWBATWorksheet)
    
    Call SaveAsWorkBook(newBook, "", "Test")
    
    Set newBook = Nothing
    
End Sub

#End If

