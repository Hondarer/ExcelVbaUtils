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

' Dependency: None

#Const ENABLE_TEST_METHODS = 1

' パスワードに用いる文字の組み合わせを表します。
Public Enum gpPasswordType
    ' アルファベットのみ。
    gpAlphabetOnly = 0
    ' アルファベットと数字。
    gpAlphabetAndNumeric = 1
    ' アルファベットと数字と記号。
    gpIncludeSymbol = 2
End Enum

' 間違えそうな文字は意図的に外し、対象にしていない
Private Const GPELEMENTS_ALPHABET = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ"
Private Const GPELEMENTS_NUMERIC = "0123456789"
Private Const GPELEMENTS_SYMBOLS = "!#$%&()+-./:<=>?[]^_|" ' @ は AD では禁則文字

Private Const GPMINLENGTH = 8

' -----------------------------------------------------------------------------
' パスワード文字列を生成します。
' <IN> passwordType As gpPasswordType パスワードに用いる文字の組み合わせ。
' <IN>passwordLength As Long パスワードの長さ。最低文字数を下回る場合は、最低文字数に補正されます。
' <OUT> String 生成されたパスワード文字列。
' -----------------------------------------------------------------------------
Public Function GenaretePassword(passwordType As gpPasswordType, passwordLength As Long) As String
    
    Dim result As String
    Dim count As Long
    
    Dim numerics As Long
    Dim symbols As Long
    
    Dim insertIndex As Long

    ' 最低文字数を下回る場合は、最低文字数に補正
    If passwordLength < GPMINLENGTH Then
        passwordLength = GPMINLENGTH
    End If

    ' 乱数系列の初期化
    Randomize
    
    ' 含まれる数字の文字数を算出
    If passwordType = gpAlphabetAndNumeric Or passwordType = gpIncludeSymbol Then
        ' 全体の 25% 程度(ただし、最低数 1)
        numerics = Int(Rnd * passwordLength / 4) + 1
    End If
    
    ' 含まれる記号の文字数を算出
    If passwordType = gpIncludeSymbol Then
        ' 全体の 12% 程度(ただし、最低数 1)
        symbols = Int(Rnd * passwordLength / 8) + 1
    End If
    
    ' アルファベットのパスワードを生成
    For count = 1 To (passwordLength - numerics - symbols)
        result = result & Mid(GPELEMENTS_ALPHABET, Int(Rnd * Len(GPELEMENTS_ALPHABET)) + 1, 1)
    Next
    
    ' 数字部分のパスワードを生成して挿入
    For count = 1 To numerics
        insertIndex = Int(Rnd * (Len(result) + 1))
        result = Left(result, insertIndex) & Mid(GPELEMENTS_NUMERIC, Int(Rnd * Len(GPELEMENTS_NUMERIC)) + 1, 1) & Mid(result, insertIndex + 1)
    Next
    
    ' 記号部分のパスワードを生成して挿入
    For count = 1 To symbols
        insertIndex = Int(Rnd * (Len(result) + 1))
        result = Left(result, insertIndex) & Mid(GPELEMENTS_SYMBOLS, Int(Rnd * Len(GPELEMENTS_SYMBOLS)) + 1, 1) & Mid(result, insertIndex + 1)
    Next
    
    GenaretePassword = result

End Function

#If ENABLE_TEST_METHODS = 1 Then

' -----------------------------------------------------------------------------
' GenaretePassword メソッドのテストを行います。
' -----------------------------------------------------------------------------
Public Sub GenaretePasswordTest()
    Debug.Print GenaretePassword(gpIncludeSymbol, 8)
End Sub

#End If

