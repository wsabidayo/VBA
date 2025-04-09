Attribute VB_Name = "modTest_clsEnhancedString"
Option Explicit

'/**
' * @description clsEnhancedStringのテストを実行する
' */
Public Sub Test_clsEnhancedString()
    ' すべてのテストケースを実行
    Test_clsEnhancedString_Initialize
    Test_clsEnhancedString_Value
    Test_clsEnhancedString_Length
    Test_clsEnhancedString_Mutable
    Test_clsEnhancedString_Concat
    Test_clsEnhancedString_ToUpperCase
    Test_clsEnhancedString_ToLowerCase
    Test_clsEnhancedString_Trim
    Test_clsEnhancedString_LTrim
    Test_clsEnhancedString_RTrim
    Test_clsEnhancedString_Substring
    Test_clsEnhancedString_PadLeft
    Test_clsEnhancedString_PadRight
    Test_clsEnhancedString_Repeat
    Test_clsEnhancedString_Includes
    Test_clsEnhancedString_IndexOf
    Test_clsEnhancedString_StartsWith
    Test_clsEnhancedString_EndsWith
    Test_clsEnhancedString_Replace
    Test_clsEnhancedString_Split
    Test_clsEnhancedString_Template
    Test_clsEnhancedString_Reverse
    Test_clsEnhancedString_Test
    Test_clsEnhancedString_Match
    Test_clsEnhancedString_ChainMethods
    
    MsgBox "すべてのテストが完了しました", vbInformation, "テスト完了"
End Sub

'/**
' * @description テスト結果を検証する
' * @param {Boolean} pCondition - 検証条件
' * @param {String} pTestName - テスト名
' */
Private Sub AssertTrue(ByVal pCondition As Boolean, ByVal pTestName As String)
    If pCondition Then
        Debug.Print "PASS: " & pTestName
    Else
        Debug.Print "FAIL: " & pTestName
        ' オプション: テスト失敗時に実行を停止
        ' Debug.Assert False
    End If
End Sub

'/**
' * @description Initialize メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Initialize()
    Dim lvStr As New clsEnhancedString
    
    ' 初期化前のデフォルト値をテスト
    AssertTrue lvStr.Value = "", "Initialize - Default Value Should Be Empty"
    AssertTrue lvStr.Mutable = False, "Initialize - Default Mutable Should Be False"
    
    ' 初期化後の値をテスト
    lvStr.Initialize "Hello", True
    AssertTrue lvStr.Value = "Hello", "Initialize - Value Should Be Set"
    AssertTrue lvStr.Mutable = True, "Initialize - Mutable Should Be Set"
    
    ' オプションパラメータをテスト
    Dim lvStr2 As New clsEnhancedString
    lvStr2.Initialize "World"
    AssertTrue lvStr2.Value = "World", "Initialize - Optional Param - Value Should Be Set"
    AssertTrue lvStr2.Mutable = False, "Initialize - Optional Param - Default Mutable Should Be False"
End Sub

'/**
' * @description Value プロパティのテスト
' */
Public Sub Test_clsEnhancedString_Value()
    Dim lvStr As New clsEnhancedString
    
    ' 値の設定と取得をテスト
    lvStr.Value = "Test Value"
    AssertTrue lvStr.Value = "Test Value", "Value - Should Get And Set Correctly"
    
    ' 空の値をテスト
    lvStr.Value = ""
    AssertTrue lvStr.Value = "", "Value - Should Handle Empty String"
    
    ' 特殊文字をテスト
    lvStr.Value = "!@#$%^&*()"
    AssertTrue lvStr.Value = "!@#$%^&*()", "Value - Should Handle Special Characters"
End Sub

'/**
' * @description Length プロパティのテスト
' */
Public Sub Test_clsEnhancedString_Length()
    Dim lvStr As New clsEnhancedString
    
    ' 空文字列の長さをテスト
    lvStr.Initialize ""
    AssertTrue lvStr.length = 0, "Length - Empty String Should Have Length 0"
    
    ' 通常の文字列の長さをテスト
    lvStr.Initialize "Hello World"
    AssertTrue lvStr.length = 11, "Length - Should Calculate Correctly"
    
    ' 特殊文字を含む文字列の長さをテスト
    lvStr.Initialize "こんにちは世界"
    AssertTrue lvStr.length = 7, "Length - Should Handle Non-ASCII Characters"
End Sub

'/**
' * @description Mutable プロパティのテスト
' */
Public Sub Test_clsEnhancedString_Mutable()
    Dim lvStr As New clsEnhancedString
    
    ' デフォルト値をテスト
    AssertTrue lvStr.Mutable = False, "Mutable - Default Should Be False"
    
    ' 値の設定と取得をテスト
    lvStr.Mutable = True
    AssertTrue lvStr.Mutable = True, "Mutable - Should Get And Set True"
    
    lvStr.Mutable = False
    AssertTrue lvStr.Mutable = False, "Mutable - Should Get And Set False"
    
    ' Initialize経由での設定をテスト
    lvStr.Initialize "Test", True
    AssertTrue lvStr.Mutable = True, "Mutable - Should Set Through Initialize"
End Sub

'/**
' * @description Concat メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Concat()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な連結をテスト
    lvStr.Initialize "Hello"
    Set lvResult = lvStr.Concat(" World")
    AssertTrue lvResult.Value = "Hello World", "Concat - Should Concatenate Strings"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "Hello", "Concat - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.Concat("!")
    AssertTrue lvResult.Value = "Hello!", "Concat - Should Concatenate In Mutable Mode"
    AssertTrue lvStr.Value = "Hello!", "Concat - Original String Should Change In Mutable Mode"
    
    ' 空文字列との連結をテスト
    Set lvResult = lvStr.Concat("")
    AssertTrue lvResult.Value = "Hello!", "Concat - Should Handle Empty String"
End Sub

'/**
' * @description ToUpperCase メソッドのテスト
' */
Public Sub Test_clsEnhancedString_ToUpperCase()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な大文字変換をテスト
    lvStr.Initialize "hello world"
    Set lvResult = lvStr.ToUpperCase
    AssertTrue lvResult.Value = "HELLO WORLD", "ToUpperCase - Should Convert To Upper Case"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "hello world", "ToUpperCase - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.ToUpperCase
    AssertTrue lvResult.Value = "HELLO WORLD", "ToUpperCase - Should Convert In Mutable Mode"
    AssertTrue lvStr.Value = "HELLO WORLD", "ToUpperCase - Original String Should Change In Mutable Mode"
    
    ' 既に大文字の文字列をテスト
    lvStr.Initialize "ALREADY UPPERCASE"
    Set lvResult = lvStr.ToUpperCase
    AssertTrue lvResult.Value = "ALREADY UPPERCASE", "ToUpperCase - Should Handle Already Uppercase String"
    
    ' 数字と特殊文字をテスト
    lvStr.Initialize "Hello 123 !@#"
    Set lvResult = lvStr.ToUpperCase
    AssertTrue lvResult.Value = "HELLO 123 !@#", "ToUpperCase - Should Handle Numbers And Special Characters"
End Sub

'/**
' * @description ToLowerCase メソッドのテスト
' */
Public Sub Test_clsEnhancedString_ToLowerCase()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な小文字変換をテスト
    lvStr.Initialize "HELLO WORLD"
    Set lvResult = lvStr.ToLowerCase
    AssertTrue lvResult.Value = "hello world", "ToLowerCase - Should Convert To Lower Case"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "HELLO WORLD", "ToLowerCase - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.ToLowerCase
    AssertTrue lvResult.Value = "hello world", "ToLowerCase - Should Convert In Mutable Mode"
    AssertTrue lvStr.Value = "hello world", "ToLowerCase - Original String Should Change In Mutable Mode"
    
    ' 既に小文字の文字列をテスト
    lvStr.Initialize "already lowercase"
    Set lvResult = lvStr.ToLowerCase
    AssertTrue lvResult.Value = "already lowercase", "ToLowerCase - Should Handle Already Lowercase String"
    
    ' 数字と特殊文字をテスト
    lvStr.Initialize "HELLO 123 !@#"
    Set lvResult = lvStr.ToLowerCase
    AssertTrue lvResult.Value = "hello 123 !@#", "ToLowerCase - Should Handle Numbers And Special Characters"
End Sub

'/**
' * @description Trim メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Trim()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 両端の空白を削除するテスト
    lvStr.Initialize "  Hello World  "
    Set lvResult = lvStr.Trim
    AssertTrue lvResult.Value = "Hello World", "Trim - Should Remove Whitespace From Both Ends"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "  Hello World  ", "Trim - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.Trim
    AssertTrue lvResult.Value = "Hello World", "Trim - Should Trim In Mutable Mode"
    AssertTrue lvStr.Value = "Hello World", "Trim - Original String Should Change In Mutable Mode"
    
    ' 空白のみの文字列をテスト
    lvStr.Initialize "    "
    Set lvResult = lvStr.Trim
    AssertTrue lvResult.Value = "", "Trim - Should Handle String With Only Whitespace"
    
    ' 既にトリム済みの文字列をテスト
    lvStr.Initialize "NoWhitespace"
    Set lvResult = lvStr.Trim
    AssertTrue lvResult.Value = "NoWhitespace", "Trim - Should Handle String With No Whitespace"
End Sub

'/**
' * @description LTrim メソッドのテスト
' */
Public Sub Test_clsEnhancedString_LTrim()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 左側の空白を削除するテスト
    lvStr.Initialize "  Hello World  "
    Set lvResult = lvStr.LTrim
    AssertTrue lvResult.Value = "Hello World  ", "LTrim - Should Remove Whitespace From Left End Only"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "  Hello World  ", "LTrim - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.LTrim
    AssertTrue lvResult.Value = "Hello World  ", "LTrim - Should Trim In Mutable Mode"
    AssertTrue lvStr.Value = "Hello World  ", "LTrim - Original String Should Change In Mutable Mode"
    
    ' 空白のみの文字列をテスト
    lvStr.Initialize "    "
    Set lvResult = lvStr.LTrim
    AssertTrue lvResult.Value = "", "LTrim - Should Handle String With Only Whitespace"
    
    ' 左側に空白がない文字列をテスト
    lvStr.Initialize "NoLeftWhitespace  "
    Set lvResult = lvStr.LTrim
    AssertTrue lvResult.Value = "NoLeftWhitespace  ", "LTrim - Should Handle String With No Left Whitespace"
End Sub

'/**
' * @description RTrim メソッドのテスト
' */
Public Sub Test_clsEnhancedString_RTrim()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 右側の空白を削除するテスト
    lvStr.Initialize "  Hello World  "
    Set lvResult = lvStr.RTrim
    AssertTrue lvResult.Value = "  Hello World", "RTrim - Should Remove Whitespace From Right End Only"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "  Hello World  ", "RTrim - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.RTrim
    AssertTrue lvResult.Value = "  Hello World", "RTrim - Should Trim In Mutable Mode"
    AssertTrue lvStr.Value = "  Hello World", "RTrim - Original String Should Change In Mutable Mode"
    
    ' 空白のみの文字列をテスト
    lvStr.Initialize "    "
    Set lvResult = lvStr.RTrim
    AssertTrue lvResult.Value = "", "RTrim - Should Handle String With Only Whitespace"
    
    ' 右側に空白がない文字列をテスト
    lvStr.Initialize "  NoRightWhitespace"
    Set lvResult = lvStr.RTrim
    AssertTrue lvResult.Value = "  NoRightWhitespace", "RTrim - Should Handle String With No Right Whitespace"
End Sub

'/**
' * @description Substring メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Substring()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な部分文字列取得をテスト
    lvStr.Initialize "Hello World"
    Set lvResult = lvStr.Substring(0, 5)
    AssertTrue lvResult.Value = "Hello", "Substring - Should Extract Correct Substring With Length"
    
    ' 開始位置のみ指定するテスト
    Set lvResult = lvStr.Substring(6)
    AssertTrue lvResult.Value = "World", "Substring - Should Extract From Start Position To End"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "Hello World", "Substring - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.Substring(0, 5)
    AssertTrue lvResult.Value = "Hello", "Substring - Should Extract In Mutable Mode"
    AssertTrue lvStr.Value = "Hello", "Substring - Original String Should Change In Mutable Mode"
    
    ' 境界値のテスト
    lvStr.Initialize "Testing"
    Set lvResult = lvStr.Substring(0)
    AssertTrue lvResult.Value = "Testing", "Substring - Should Handle Start Index 0"
    
    Set lvResult = lvStr.Substring(0, 7)
    AssertTrue lvResult.Value = "Testing", "Substring - Should Handle Exact Length"
    
    ' これはVBAの仕様上、エラーにはならず、空文字列を返す（負のインデックスはゼロとして扱われるため）
    Set lvResult = lvStr.Substring(7)
    AssertTrue lvResult.Value = "", "Substring - Should Handle Start Index At End"
End Sub

'/**
' * @description PadLeft メソッドのテスト
' */
Public Sub Test_clsEnhancedString_PadLeft()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な左パディングをテスト
    lvStr.Initialize "123"
    Set lvResult = lvStr.PadLeft(5)
    AssertTrue lvResult.Value = "  123", "PadLeft - Should Pad With Spaces By Default"
    
    ' カスタム文字でのパディングをテスト
    Set lvResult = lvStr.PadLeft(5, "0")
    AssertTrue lvResult.Value = "00123", "PadLeft - Should Pad With Custom Character"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "123", "PadLeft - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.PadLeft(5, "*")
    AssertTrue lvResult.Value = "**123", "PadLeft - Should Pad In Mutable Mode"
    AssertTrue lvStr.Value = "**123", "PadLeft - Original String Should Change In Mutable Mode"
    
    ' 既に長い文字列をテスト
    lvStr.Initialize "12345"
    Set lvResult = lvStr.PadLeft(3)
    AssertTrue lvResult.Value = "12345", "PadLeft - Should Not Truncate If String Longer Than Width"
    
    ' 複数文字のパディング文字をテスト（最初の1文字のみが使用される）
    lvStr.Initialize "123"
    Set lvResult = lvStr.PadLeft(5, "AB")
    AssertTrue lvResult.Value = "AA123", "PadLeft - Should Use Only First Character Of Padding String"
End Sub

'/**
' * @description PadRight メソッドのテスト
' */
Public Sub Test_clsEnhancedString_PadRight()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な右パディングをテスト
    lvStr.Initialize "123"
    Set lvResult = lvStr.PadRight(5)
    AssertTrue lvResult.Value = "123  ", "PadRight - Should Pad With Spaces By Default"
    
    ' カスタム文字でのパディングをテスト
    Set lvResult = lvStr.PadRight(5, "0")
    AssertTrue lvResult.Value = "12300", "PadRight - Should Pad With Custom Character"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "123", "PadRight - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.PadRight(5, "*")
    AssertTrue lvResult.Value = "123**", "PadRight - Should Pad In Mutable Mode"
    AssertTrue lvStr.Value = "123**", "PadRight - Original String Should Change In Mutable Mode"
    
    ' 既に長い文字列をテスト
    lvStr.Initialize "12345"
    Set lvResult = lvStr.PadRight(3)
    AssertTrue lvResult.Value = "12345", "PadRight - Should Not Truncate If String Longer Than Width"
    
    ' 複数文字のパディング文字をテスト（最初の1文字のみが使用される）
    lvStr.Initialize "123"
    Set lvResult = lvStr.PadRight(5, "AB")
    AssertTrue lvResult.Value = "123AA", "PadRight - Should Use Only First Character Of Padding String"
End Sub

'/**
' * @description Repeat メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Repeat()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な繰り返しをテスト
    lvStr.Initialize "abc"
    Set lvResult = lvStr.Repeat(3)
    AssertTrue lvResult.Value = "abcabcabc", "Repeat - Should Repeat String Correctly"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "abc", "Repeat - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.Repeat(2)
    AssertTrue lvResult.Value = "abcabc", "Repeat - Should Repeat In Mutable Mode"
    AssertTrue lvStr.Value = "abcabc", "Repeat - Original String Should Change In Mutable Mode"
    
    ' ゼロ回の繰り返しをテスト
    lvStr.Initialize "test"
    Set lvResult = lvStr.Repeat(0)
    AssertTrue lvResult.Value = "", "Repeat - Zero Repetitions Should Result In Empty String"
    
    ' 1回の繰り返しをテスト
    Set lvResult = lvStr.Repeat(1)
    AssertTrue lvResult.Value = "test", "Repeat - One Repetition Should Equal Original String"
    
    ' 空文字列の繰り返しをテスト
    lvStr.Initialize ""
    Set lvResult = lvStr.Repeat(10)
    AssertTrue lvResult.Value = "", "Repeat - Repeating Empty String Should Result In Empty String"
End Sub

'/**
' * @description Includes メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Includes()
    Dim lvStr As New clsEnhancedString
    
    ' 含まれる部分文字列をテスト
    lvStr.Initialize "Hello World"
    AssertTrue lvStr.Includes("World") = True, "Includes - Should Return True For Included Substring"
    AssertTrue lvStr.Includes("Hello") = True, "Includes - Should Return True For Included Substring At Start"
    AssertTrue lvStr.Includes("llo Wo") = True, "Includes - Should Return True For Included Substring In Middle"
    
    ' 含まれない部分文字列をテスト
    AssertTrue lvStr.Includes("Goodbye") = False, "Includes - Should Return False For Non-included Substring"
    
    ' 大文字小文字の区別をテスト
    AssertTrue lvStr.Includes("world") = True, "Includes - Should Be Case Insensitive"
    
    ' 空文字列をテスト
    AssertTrue lvStr.Includes("") = True, "Includes - Empty String Should Always Be Included"
    
    ' 空の対象文字列をテスト
    lvStr.Initialize ""
    AssertTrue lvStr.Includes("test") = False, "Includes - Empty Target String Should Only Include Empty String"
    AssertTrue lvStr.Includes("") = True, "Includes - Empty Target String Should Include Empty String"
End Sub

'/**
' * @description IndexOf メソッドのテスト
' */
Public Sub Test_clsEnhancedString_IndexOf()
    Dim lvStr As New clsEnhancedString
    
    ' 基本的な位置検索をテスト
    lvStr.Initialize "Hello World"
    AssertTrue lvStr.IndexOf("World") = 6, "IndexOf - Should Return Correct Index"
    AssertTrue lvStr.IndexOf("Hello") = 0, "IndexOf - Should Return 0 For Start"
    AssertTrue lvStr.IndexOf("llo") = 2, "IndexOf - Should Return Correct Index For Substring"
    
    ' 含まれない部分文字列をテスト
    AssertTrue lvStr.IndexOf("Goodbye") = -1, "IndexOf - Should Return -1 For Non-included Substring"
    
    ' 大文字小文字の区別をテスト
    AssertTrue lvStr.IndexOf("world") = 6, "IndexOf - Should Be Case Insensitive"
    
    ' 空文字列をテスト
    AssertTrue lvStr.IndexOf("") = 0, "IndexOf - Empty String Should Be Found At Position 0"
    
    ' 空の対象文字列をテスト
    lvStr.Initialize ""
    AssertTrue lvStr.IndexOf("test") = -1, "IndexOf - Empty Target String Should Not Include Non-empty String"
    AssertTrue lvStr.IndexOf("") = 0, "IndexOf - Empty Target String Should Include Empty String At Position 0"
End Sub

'/**
' * @description StartsWith メソッドのテスト
' */
Public Sub Test_clsEnhancedString_StartsWith()
    Dim lvStr As New clsEnhancedString
    
    ' 基本的な先頭一致をテスト
    lvStr.Initialize "Hello World"
    AssertTrue lvStr.StartsWith("Hello") = True, "StartsWith - Should Return True For Starting Substring"
    AssertTrue lvStr.StartsWith("Hello World") = True, "StartsWith - Should Return True For Exact Match"
    
    ' 先頭以外の一致をテスト
    AssertTrue lvStr.StartsWith("World") = False, "StartsWith - Should Return False For Non-starting Substring"
    
    ' 大文字小文字の区別をテスト
    AssertTrue lvStr.StartsWith("HELLO") = False, "StartsWith - Should Be Case Sensitive"
    
    ' 空文字列をテスト
    AssertTrue lvStr.StartsWith("") = True, "StartsWith - Empty String Should Always Match At Start"
    
    ' 対象文字列より長い検索文字列をテスト
    AssertTrue lvStr.StartsWith("Hello World Plus") = False, "StartsWith - Should Return False For Longer Search String"
    
    ' 空の対象文字列をテスト
    lvStr.Initialize ""
    AssertTrue lvStr.StartsWith("test") = False, "StartsWith - Empty Target String Should Not Start With Non-empty String"
    AssertTrue lvStr.StartsWith("") = True, "StartsWith - Empty Target String Should Start With Empty String"
End Sub

'/**
' * @description EndsWith メソッドのテスト
' */
Public Sub Test_clsEnhancedString_EndsWith()
    Dim lvStr As New clsEnhancedString
    
    ' 基本的な末尾一致をテスト
    lvStr.Initialize "Hello World"
    AssertTrue lvStr.EndsWith("World") = True, "EndsWith - Should Return True For Ending Substring"
    AssertTrue lvStr.EndsWith("Hello World") = True, "EndsWith - Should Return True For Exact Match"
    
    ' 末尾以外の一致をテスト
    AssertTrue lvStr.EndsWith("Hello") = False, "EndsWith - Should Return False For Non-ending Substring"
    
    ' 大文字小文字の区別をテスト
    AssertTrue lvStr.EndsWith("WORLD") = False, "EndsWith - Should Be Case Sensitive"
    
    ' 空文字列をテスト
    AssertTrue lvStr.EndsWith("") = True, "EndsWith - Empty String Should Always Match At End"
    
    ' 対象文字列より長い検索文字列をテスト
    AssertTrue lvStr.EndsWith("More Hello World") = False, "EndsWith - Should Return False For Longer Search String"
    
    ' 空の対象文字列をテスト
    lvStr.Initialize ""
    AssertTrue lvStr.EndsWith("test") = False, "EndsWith - Empty Target String Should Not End With Non-empty String"
    AssertTrue lvStr.EndsWith("") = True, "EndsWith - Empty Target String Should End With Empty String"
End Sub

'/**
' * @description Replace メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Replace()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な置換をテスト
    lvStr.Initialize "Hello World"
    Set lvResult = lvStr.Replace("o", "0")
    AssertTrue lvResult.Value = "Hell0 W0rld", "Replace - Should Replace All Occurrences"
    
    ' グローバル置換のON/OFFをテスト
    Set lvResult = lvStr.Replace("o", "0", True, False)
    AssertTrue lvResult.Value = "Hell0 World", "Replace - Should Replace Only First Occurrence When Global=False"
    
    ' 大文字小文字の区別をテスト
    Set lvResult = lvStr.Replace("O", "0", False)
    AssertTrue lvResult.Value = "Hello World", "Replace - Should Consider Case When IgnoreCase=False"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "Hello World", "Replace - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.Replace("o", "0")
    AssertTrue lvResult.Value = "Hell0 W0rld", "Replace - Should Replace In Mutable Mode"
    AssertTrue lvStr.Value = "Hell0 W0rld", "Replace - Original String Should Change In Mutable Mode"
    
    ' 正規表現パターンでの置換をテスト
    lvStr.Initialize "Hello 123 World 456"
    Set lvResult = lvStr.Replace("\d+", "NUM")
    AssertTrue lvResult.Value = "Hello NUM World NUM", "Replace - Should Handle Regex Patterns"
    
    ' 存在しないパターンの置換をテスト
    Set lvResult = lvStr.Replace("xyz", "abc")
    AssertTrue lvResult.Value = "Hello 123 World 456", "Replace - Should Return Unchanged String When Pattern Not Found"
    
    ' 空文字列への置換をテスト
    Set lvResult = lvStr.Replace("\s", "")
    AssertTrue lvResult.Value = "Hello123World456", "Replace - Should Allow Replacement With Empty String"
End Sub

'/**
' * @description Split メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Split()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As Variant
    
    ' 基本的な分割をテスト
    lvStr.Initialize "a,b,c"
    lvResult = lvStr.Split(",")
    AssertTrue UBound(lvResult) = 2, "Split - Should Return Correct Number Of Elements"
    AssertTrue lvResult(0) = "a", "Split - First Element Should Be Correct"
    AssertTrue lvResult(1) = "b", "Split - Second Element Should Be Correct"
    AssertTrue lvResult(2) = "c", "Split - Third Element Should Be Correct"
    
    ' 複数の区切り文字を持つ文字列をテスト
    lvStr.Initialize "Hello World Test"
    lvResult = lvStr.Split(" ")
    AssertTrue UBound(lvResult) = 2, "Split - Should Handle Multiple Delimiters"
    AssertTrue lvResult(0) = "Hello", "Split - First Element Should Be Correct With Space Delimiter"
    AssertTrue lvResult(1) = "World", "Split - Second Element Should Be Correct With Space Delimiter"
    AssertTrue lvResult(2) = "Test", "Split - Third Element Should Be Correct With Space Delimiter"
    
    ' 区切り文字がない場合をテスト
    lvStr.Initialize "NoDelimiter"
    lvResult = lvStr.Split(",")
    AssertTrue UBound(lvResult) = 0, "Split - Should Return Single Element When No Delimiter Present"
    AssertTrue lvResult(0) = "NoDelimiter", "Split - First Element Should Be Entire String When No Delimiter"
    
    ' 空文字列をテスト
    lvStr.Initialize ""
    lvResult = lvStr.Split(",")
    AssertTrue UBound(lvResult) = -1, "Split - Should Handle Empty String"
    
    ' 連続する区切り文字をテスト
    lvStr.Initialize "a,,c"
    lvResult = lvStr.Split(",")
    AssertTrue UBound(lvResult) = 2, "Split - Should Handle Consecutive Delimiters"
    AssertTrue lvResult(0) = "a", "Split - First Element Should Be Correct With Consecutive Delimiters"
    AssertTrue lvResult(1) = "", "Split - Second Element Should Be Empty With Consecutive Delimiters"
    AssertTrue lvResult(2) = "c", "Split - Third Element Should Be Correct With Consecutive Delimiters"
End Sub

'/**
' * @description Template メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Template()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的なテンプレート置換をテスト
    lvStr.Initialize "Hello, {0}! Today is {1}."
    Set lvResult = lvStr.Template("World", "Monday")
    AssertTrue lvResult.Value = "Hello, World! Today is Monday.", "Template - Should Replace Placeholders"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "Hello, {0}! Today is {1}.", "Template - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.Template("User", "Tuesday")
    AssertTrue lvResult.Value = "Hello, User! Today is Tuesday.", "Template - Should Replace In Mutable Mode"
    AssertTrue lvStr.Value = "Hello, User! Today is Tuesday.", "Template - Original String Should Change In Mutable Mode"
    
    ' 多数のプレースホルダーをテスト
    lvStr.Initialize "{0} {1} {2} {3} {4}"
    Set lvResult = lvStr.Template("A", "B", "C", "D", "E")
    AssertTrue lvResult.Value = "A B C D E", "Template - Should Handle Multiple Placeholders"
    
    ' パラメータが足りない場合をテスト
    lvStr.Initialize "{0} {1} {2}"
    Set lvResult = lvStr.Template("A")
    AssertTrue lvResult.Value = "A {1} {2}", "Template - Should Leave Unfilled Placeholders"
    
    ' パラメータが多い場合をテスト
    lvStr.Initialize "{0}"
    Set lvResult = lvStr.Template("A", "B", "C")
    AssertTrue lvResult.Value = "A", "Template - Should Ignore Extra Parameters"
    
    ' プレースホルダーがない場合をテスト
    lvStr.Initialize "No placeholders"
    Set lvResult = lvStr.Template("A", "B")
    AssertTrue lvResult.Value = "No placeholders", "Template - Should Return Unchanged String When No Placeholders"
    
    ' 数値以外のプレースホルダーをテスト
    lvStr.Initialize "Hello, {name}!"
    Set lvResult = lvStr.Template("World")
    AssertTrue lvResult.Value = "Hello, {name}!", "Template - Should Only Replace Numeric Placeholders"
End Sub

'/**
' * @description Reverse メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Reverse()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的な文字列反転をテスト
    lvStr.Initialize "Hello"
    Set lvResult = lvStr.Reverse
    AssertTrue lvResult.Value = "olleH", "Reverse - Should Reverse String"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "Hello", "Reverse - Original String Should Not Change In Immutable Mode"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    Set lvResult = lvStr.Reverse
    AssertTrue lvResult.Value = "olleH", "Reverse - Should Reverse In Mutable Mode"
    AssertTrue lvStr.Value = "olleH", "Reverse - Original String Should Change In Mutable Mode"
    
    ' 空文字列をテスト
    lvStr.Initialize ""
    Set lvResult = lvStr.Reverse
    AssertTrue lvResult.Value = "", "Reverse - Should Handle Empty String"
    
    ' パリンドローム（回文）をテスト
    lvStr.Initialize "radar"
    Set lvResult = lvStr.Reverse
    AssertTrue lvResult.Value = "radar", "Reverse - Should Handle Palindromes"
    
    ' 複数単語をテスト
    lvStr.Initialize "Hello World"
    Set lvResult = lvStr.Reverse
    AssertTrue lvResult.Value = "dlroW olleH", "Reverse - Should Handle Spaces And Multiple Words"
    
    ' 特殊文字をテスト
    lvStr.Initialize "Hello! @#$"
    Set lvResult = lvStr.Reverse
    AssertTrue lvResult.Value = "$#@ !olleH", "Reverse - Should Handle Special Characters"
End Sub

'/**
' * @description Test メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Test()
    Dim lvStr As New clsEnhancedString
    
    ' 基本的な正規表現テストをテスト
    lvStr.Initialize "Hello World"
    AssertTrue lvStr.Test("World") = True, "Test - Should Return True For Simple Match"
    AssertTrue lvStr.Test("Goodbye") = False, "Test - Should Return False When Not Matched"
    
    ' 正規表現パターンをテスト
    AssertTrue lvStr.Test("^Hello") = True, "Test - Should Support Start Anchor"
    AssertTrue lvStr.Test("World$") = True, "Test - Should Support End Anchor"
    AssertTrue lvStr.Test("H.+W") = True, "Test - Should Support Regex Quantifiers"
    AssertTrue lvStr.Test("\w+\s\w+") = True, "Test - Should Support Word And Whitespace Patterns"
    
    ' 大文字小文字の区別をテスト
    AssertTrue lvStr.Test("hello", True) = True, "Test - Should Be Case Insensitive By Default"
    AssertTrue lvStr.Test("hello", False) = False, "Test - Should Be Case Sensitive When Specified"
    
    ' 複雑なパターンをテスト
    lvStr.Initialize "test123@example.com"
    AssertTrue lvStr.Test("^\w+@\w+\.\w{2,}$") = True, "Test - Should Support Complex Patterns"
    
    ' 空文字列をテスト
    lvStr.Initialize ""
    AssertTrue lvStr.Test(".+") = False, "Test - Empty String Should Not Match Non-empty Pattern"
    AssertTrue lvStr.Test("^$") = True, "Test - Empty String Should Match Empty Pattern"
End Sub

'/**
' * @description Match メソッドのテスト
' */
Public Sub Test_clsEnhancedString_Match()
    Dim lvStr As New clsEnhancedString
    Dim lvMatches As MatchCollection
    
    ' 基本的なマッチをテスト
    lvStr.Initialize "Hello World"
    Set lvMatches = lvStr.Match("World")
    AssertTrue lvMatches.Count = 1, "Match - Should Return Correct Number Of Matches"
    AssertTrue lvMatches(0).Value = "World", "Match - Should Return Correct Match Value"
    
    ' 複数マッチをテスト
    lvStr.Initialize "test test test"
    Set lvMatches = lvStr.Match("test")
    AssertTrue lvMatches.Count = 3, "Match - Should Return All Matches"
    AssertTrue lvMatches(0).Value = "test", "Match - First Match Should Be Correct"
    AssertTrue lvMatches(1).Value = "test", "Match - Second Match Should Be Correct"
    AssertTrue lvMatches(2).Value = "test", "Match - Third Match Should Be Correct"
    
    ' 正規表現パターンをテスト
    lvStr.Initialize "Hello 123 World 456"
    Set lvMatches = lvStr.Match("\d+")
    AssertTrue lvMatches.Count = 2, "Match - Should Handle Regex Patterns"
    AssertTrue lvMatches(0).Value = "123", "Match - Should Extract Numbers Correctly"
    AssertTrue lvMatches(1).Value = "456", "Match - Should Extract All Number Groups"
    
    ' グローバルマッチのON/OFFをテスト
    Set lvMatches = lvStr.Match("\d+", True, False)
    AssertTrue lvMatches.Count = 1, "Match - Should Return Only First Match When Global=False"
    AssertTrue lvMatches(0).Value = "123", "Match - First Match Should Be Correct When Global=False"
    
    ' 大文字小文字の区別をテスト
    lvStr.Initialize "Test TEST test"
    Set lvMatches = lvStr.Match("test", True)
    AssertTrue lvMatches.Count = 3, "Match - Should Be Case Insensitive By Default"
    
    Set lvMatches = lvStr.Match("test", False)
    AssertTrue lvMatches.Count = 1, "Match - Should Be Case Sensitive When Specified"
    AssertTrue lvMatches(0).Value = "test", "Match - Should Match Only Lowercase When Case Sensitive"
    
    ' マッチがない場合をテスト
    lvStr.Initialize "Hello World"
    Set lvMatches = lvStr.Match("xyz")
    AssertTrue lvMatches.Count = 0, "Match - Should Return Empty Collection When No Matches"
    
    ' 空文字列をテスト
    lvStr.Initialize ""
    Set lvMatches = lvStr.Match(".")
    AssertTrue lvMatches.Count = 0, "Match - Empty String Should Not Match Non-empty Pattern"
End Sub

'/**
' * @description メソッドチェーンのテスト
' */
Public Sub Test_clsEnhancedString_ChainMethods()
    Dim lvStr As New clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    ' 基本的なメソッドチェーンをテスト
    lvStr.Initialize "  hello world  "
    Set lvResult = lvStr.Trim().ToUpperCase()
    AssertTrue lvResult.Value = "HELLO WORLD", "ChainMethods - Should Support Basic Method Chaining"
    
    ' イミュータブルモードでは元の値は変更されないことをテスト
    AssertTrue lvStr.Value = "  hello world  ", "ChainMethods - Original String Should Not Change In Immutable Mode"
    
    ' 複雑なメソッドチェーンをテスト
    Set lvResult = lvStr.Trim().ToUpperCase().Substring(0, 5).PadRight(10, "-")
    AssertTrue lvResult.Value = "HELLO-----", "ChainMethods - Should Support Complex Method Chaining"
    
    ' ミュータブルモードをテスト
    lvStr.Mutable = True
    lvStr.Trim().ToUpperCase
    AssertTrue lvStr.Value = "HELLO WORLD", "ChainMethods - Should Apply All Changes In Mutable Mode"
    
    ' ミュータブルモードでの複雑なチェーンをテスト
    lvStr.Initialize "  hello world  ", True
    lvStr.Trim().Replace("world", "everyone").ToUpperCase().Concat ("!")
    AssertTrue lvStr.Value = "HELLO EVERYONE!", "ChainMethods - Should Apply All Changes In Complex Chain With Mutable Mode"
End Sub
