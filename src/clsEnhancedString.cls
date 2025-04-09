VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnhancedString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/**
' * @class clsEnhancedString
' * @description 拡張文字列クラス - モダンな文字列操作メソッドとチェーン呼び出しをサポート
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello")
' * Debug.Print str.ToUpperCase().Concat(", World!").Value ' 出力: "HELLO, World!"
' */

'/**
' * @typedef ThisType
' * @property {String} Value - 文字列の値
' * @property {Boolean} Mutable - ミュータブル操作を許可するかどうか
' */
Private Type ThisType
    Value As String
    Mutable As Boolean
End Type

Private mThis As ThisType

'/**
' * @description クラスの初期化処理
' */
Private Sub Class_Initialize()
    mThis.Value = ""
    mThis.Mutable = False
End Sub

'/**
' * @description クラスの終了処理
' */
Private Sub Class_Terminate()
    mThis.Value = ""
    mThis.Mutable = False
End Sub

'/**
' * @description 拡張文字列を初期化する
' * @param {String} pValue - 初期値
' * @param {Boolean} [pMutable=False] - ミュータブル操作を許可するかどうか
' */
Public Sub Initialize(ByVal pValue As String, Optional ByVal pMutable As Boolean = False)
    mThis.Value = pValue
    mThis.Mutable = pMutable
End Sub

'/**
' * @description 文字列の値を取得する
' * @returns {String} 現在の文字列値
' */
Public Property Get Value() As String
    Value = mThis.Value
End Property

'/**
' * @description 文字列の値を設定する
' * @param {String} pValue - 設定する文字列値
' */
Public Property Let Value(ByVal pValue As String)
    mThis.Value = pValue
End Property

'/**
' * @description 文字列の長さを取得する
' * @returns {Long} 文字列の長さ
' */
Public Property Get length() As Long
    length = VBA.Len(mThis.Value)
End Property

'/**
' * @description ミュータブルプロパティを取得する
' * @returns {Boolean} ミュータブル操作が許可されているかどうか
' */
Public Property Get Mutable() As Boolean
    Mutable = mThis.Mutable
End Property

'/**
' * @description ミュータブルプロパティを設定する
' * @param {Boolean} pMutable - ミュータブル操作を許可するかどうか
' */
Public Property Let Mutable(ByVal pMutable As Boolean)
    mThis.Mutable = pMutable
End Property

'/**
' * @description 文字列を連結する
' * @param {String} pText - 連結する文字列
' * @returns {clsEnhancedString} 連結された新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello")
' * Debug.Print str.Concat(" World").Value ' 出力: "Hello World"
' */
Public Function Concat(ByVal pText As String) As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = GetInstance
    
    lvResult.Value = mThis.Value & pText
    
    Set Concat = lvResult
End Function

'/**
' * @description 文字列を大文字に変換する
' * @returns {clsEnhancedString} 大文字に変換された新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("hello")
' * Debug.Print str.ToUpperCase().Value ' 出力: "HELLO"
' */
Public Function ToUpperCase() As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = GetInstance
    
    lvResult.Value = VBA.UCase(mThis.Value)
    
    Set ToUpperCase = lvResult
End Function

'/**
' * @description 文字列を小文字に変換する
' * @returns {clsEnhancedString} 小文字に変換された新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("HELLO")
' * Debug.Print str.ToLowerCase().Value ' 出力: "hello"
' */
Public Function ToLowerCase() As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = GetInstance
    
    lvResult.Value = VBA.LCase(mThis.Value)
    
    Set ToLowerCase = lvResult
End Function

'/**
' * @description 文字列の両端の空白を削除する
' * @returns {clsEnhancedString} トリムされた新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("  Hello  ")
' * Debug.Print str.Trim().Value ' 出力: "Hello"
' */
Public Function Trim() As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = GetInstance
    
    lvResult.Value = VBA.Trim(mThis.Value)
    
    Set Trim = lvResult
End Function

'/**
' * @description 文字列の左側の空白を削除する
' * @returns {clsEnhancedString} 左トリムされた新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("  Hello  ")
' * Debug.Print str.LTrim().Value ' 出力: "Hello  "
' */
Public Function LTrim() As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = GetInstance
    
    lvResult.Value = VBA.LTrim(mThis.Value)
    
    Set LTrim = lvResult
End Function

'/**
' * @description 文字列の右側の空白を削除する
' * @returns {clsEnhancedString} 右トリムされた新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("  Hello  ")
' * Debug.Print str.RTrim().Value ' 出力: "  Hello"
' */
Public Function RTrim() As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = GetInstance
    
    lvResult.Value = VBA.RTrim(mThis.Value)
    
    Set RTrim = lvResult
End Function

'/**
' * @description 文字列の部分文字列を取得する
' * @param {Long} pStartIndex - 開始インデックス（0ベース）
' * @param {Long} [pLength=-1] - 取得する文字数、-1の場合は最後まで
' * @returns {clsEnhancedString} 部分文字列を含む新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello World")
' * Debug.Print str.Substring(6).Value ' 出力: "World"
' * Debug.Print str.Substring(0, 5).Value ' 出力: "Hello"
' */
Public Function Substring(ByVal pStartIndex As Long, Optional ByVal pLength As Long = -1) As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = GetInstance
    
    If pLength = -1 Then
        lvResult.Value = VBA.Mid(mThis.Value, pStartIndex + 1)
    Else
        lvResult.Value = VBA.Mid(mThis.Value, pStartIndex + 1, pLength)
    End If
    
    Set Substring = lvResult
End Function

'/**
' * @description 指定した幅になるように文字列の左側を指定文字でパディングする
' * @param {Long} pTotalWidth - 目標の文字列幅
' * @param {String} [pPaddingChar=" "] - パディングに使用する文字
' * @returns {clsEnhancedString} パディングされた新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("123")
' * Debug.Print str.PadLeft(5).Value ' 出力: "  123"
' * Debug.Print str.PadLeft(5, "0").Value ' 出力: "00123"
' */
Public Function PadLeft(ByVal pTotalWidth As Long, Optional ByVal pPaddingChar As String = " ") As clsEnhancedString
    Dim lvResult As clsEnhancedString
    Dim lvPadding As String
    
    Set lvResult = GetInstance
    
    If Len(mThis.Value) >= pTotalWidth Then
        lvResult.Value = mThis.Value
    Else
        lvPadding = String(pTotalWidth - Len(mThis.Value), Left(pPaddingChar, 1))
        lvResult.Value = lvPadding & mThis.Value
    End If
    
    Set PadLeft = lvResult
End Function

'/**
' * @description 指定した幅になるように文字列の右側を指定文字でパディングする
' * @param {Long} pTotalWidth - 目標の文字列幅
' * @param {String} [pPaddingChar=" "] - パディングに使用する文字
' * @returns {clsEnhancedString} パディングされた新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("123")
' * Debug.Print str.PadRight(5).Value ' 出力: "123  "
' * Debug.Print str.PadRight(5, "0").Value ' 出力: "12300"
' */
Public Function PadRight(ByVal pTotalWidth As Long, Optional ByVal pPaddingChar As String = " ") As clsEnhancedString
    Dim lvResult As clsEnhancedString
    Dim lvPadding As String
    
    Set lvResult = GetInstance
    
    If Len(mThis.Value) >= pTotalWidth Then
        lvResult.Value = mThis.Value
    Else
        lvPadding = String(pTotalWidth - Len(mThis.Value), Left(pPaddingChar, 1))
        lvResult.Value = mThis.Value & lvPadding
    End If
    
    Set PadRight = lvResult
End Function

'/**
' * @description 文字列を指定回数繰り返す
' * @param {Long} pCount - 繰り返し回数
' * @returns {clsEnhancedString} 繰り返された新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("abc")
' * Debug.Print str.Repeat(3).Value ' 出力: "abcabcabc"
' */
Public Function Repeat(ByVal pCount As Long) As clsEnhancedString
    Dim lvResult As clsEnhancedString
    Dim i As Long
    Dim lvTemp As String
    
    Set lvResult = GetInstance
    lvTemp = ""
    
    For i = 1 To pCount
        lvTemp = lvTemp & mThis.Value
    Next i
    
    lvResult.Value = lvTemp
    
    Set Repeat = lvResult
End Function

'/**
' * @description 文字列に指定したテキストが含まれるかどうかを確認する
' * @param {String} pFindText - 検索するテキスト
' * @returns {Boolean} テキストが含まれる場合はTrue、そうでない場合はFalse
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello World")
' * Debug.Print str.Includes("World") ' 出力: True
' * Debug.Print str.Includes("Goodbye") ' 出力: False
' */
Public Function Includes(ByVal pFindText As String) As Boolean
    If pFindText = "" Then
        Includes = True
    Else
        Includes = (InStr(1, mThis.Value, pFindText, vbTextCompare) > 0)
    End If
End Function

'/**
' * @description 文字列内での指定したテキストの位置を検索する（0ベース）
' * @param {String} pFindText - 検索するテキスト
' * @returns {Long} 見つかった位置（0ベース）。見つからない場合は-1
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello World")
' * Debug.Print str.IndexOf("World") ' 出力: 6
' * Debug.Print str.IndexOf("Goodbye") ' 出力: -1
' */
Public Function IndexOf(ByVal pFindText As String) As Long
    If pFindText = "" Then
        IndexOf = 0
    Else
        IndexOf = InStr(1, mThis.Value, pFindText, vbTextCompare) - 1
    End If
End Function

'/**
' * @description 文字列が指定したテキストで始まるかどうかを確認する
' * @param {String} pFindText - 検索するテキスト
' * @returns {Boolean} 指定したテキストで始まる場合はTrue、そうでない場合はFalse
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello World")
' * Debug.Print str.StartsWith("Hello") ' 出力: True
' * Debug.Print str.StartsWith("World") ' 出力: False
' */
Public Function StartsWith(ByVal pFindText As String) As Boolean
    If Len(pFindText) > Len(mThis.Value) Then
        StartsWith = False
    Else
        StartsWith = (Left(mThis.Value, VBA.Len(pFindText)) = pFindText)
    End If
End Function

'/**
' * @description 文字列が指定したテキストで終わるかどうかを確認する
' * @param {String} pFindText - 検索するテキスト
' * @returns {Boolean} 指定したテキストで終わる場合はTrue、そうでない場合はFalse
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello World")
' * Debug.Print str.EndsWith("World") ' 出力: True
' * Debug.Print str.EndsWith("Hello") ' 出力: False
' */
Public Function EndsWith(ByVal pFindText As String) As Boolean
    If Len(pFindText) > Len(mThis.Value) Then
        EndsWith = False
    Else
        EndsWith = (Right(mThis.Value, VBA.Len(pFindText)) = pFindText)
    End If
End Function

'/**
' * @description 正規表現パターンに一致する部分を置換する
' * @param {String} pPattern - 正規表現パターン
' * @param {String} pReplaceText - 置換するテキスト
' * @param {Boolean} [pIgnoreCase=True] - 大文字と小文字を区別しないかどうか
' * @param {Boolean} [pGlobal=True] - すべてのマッチを置換するかどうか
' * @returns {clsEnhancedString} 置換された新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello World")
' * Debug.Print str.Replace("o", "0").Value ' 出力: "Hell0 W0rld"
' * Debug.Print str.Replace("l", "L", True, False).Value ' 出力: "HeLlo World"
' */
Public Function Replace(ByVal pPattern As String, ByVal pReplaceText As String, Optional ByVal pIgnoreCase As Boolean = True, Optional ByVal pGlobal As Boolean = True) As clsEnhancedString
    Dim lvReg As RegExp
    Dim lvResult As clsEnhancedString
    
    Set lvReg = New RegExp
    Set lvResult = GetInstance
    
    With lvReg
        .Pattern = pPattern
        .IgnoreCase = pIgnoreCase
        .Global = pGlobal
    End With
    
    lvResult.Value = lvReg.Replace(mThis.Value, pReplaceText)
    
    Set Replace = lvResult
End Function

'/**
' * @description 文字列を指定した区切り文字で分割する
' * @param {String} pDelimiter - 区切り文字
' * @returns {Variant} 分割された文字列の配列
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("a,b,c")
' * Dim arr As Variant
' * arr = str.Split(",")
' * ' arr(0) = "a", arr(1) = "b", arr(2) = "c"
' */
Public Function Split(ByVal pDelimiter As String) As Variant
    Split = VBA.Split(mThis.Value, pDelimiter)
End Function

'/**
' * @description テンプレート文字列内のプレースホルダーを指定した値で置換する
' * @param {ParamArray} pTexts - プレースホルダーに代入する値（{0}, {1}, ...）
' * @returns {clsEnhancedString} プレースホルダーが置換された新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello, {0}! Today is {1}.")
' * Debug.Print str.Template("World", "Monday").Value ' 出力: "Hello, World! Today is Monday."
' */
Public Function Template(ParamArray pTexts() As Variant) As clsEnhancedString
    Dim lvResult As clsEnhancedString
    Dim i As Long
    
    Set lvResult = GetInstance
    
    lvResult.Value = mThis.Value
    
    For i = LBound(pTexts) To UBound(pTexts)
        lvResult.Value = VBA.Replace(lvResult.Value, "{" & i & "}", CStr(pTexts(i)))
    Next i
    
    Set Template = lvResult
End Function

'/**
' * @description 文字列を反転する
' * @returns {clsEnhancedString} 反転された新しい文字列インスタンス
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello")
' * Debug.Print str.Reverse().Value ' 出力: "olleH"
' */
Public Function Reverse() As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = GetInstance
    
    lvResult.Value = VBA.StrReverse(mThis.Value)
    
    Set Reverse = lvResult
End Function

'/**
' * @description 文字列が正規表現パターンに一致するかどうかをテストする
' * @param {String} pPattern - 正規表現パターン
' * @param {Boolean} [pIgnoreCase=True] - 大文字と小文字を区別しないかどうか
' * @returns {Boolean} パターンに一致する場合はTrue、そうでない場合はFalse
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("test@example.com")
' * Debug.Print str.Test("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") ' 出力: True (メールアドレス形式チェック)
' */
Public Function Test(ByVal pPattern As String, Optional ByVal pIgnoreCase As Boolean = True) As Boolean
    Dim lvReg As RegExp
    
    Set lvReg = New RegExp
    
    With lvReg
        .Pattern = pPattern
        .IgnoreCase = pIgnoreCase
        .Global = True
    End With
    
    Test = lvReg.Test(mThis.Value)
    
    Set lvReg = Nothing
End Function

'/**
' * @description 正規表現パターンに一致するすべてのマッチを取得する
' * @param {String} pPattern - 正規表現パターン
' * @param {Boolean} [pIgnoreCase=True] - 大文字と小文字を区別しないかどうか
' * @param {Boolean} [pGlobal=True] - すべてのマッチを検索するかどうか
' * @returns {MatchCollection} 正規表現マッチのコレクション
' * @example
' * Dim str As New clsEnhancedString
' * str.Initialize("Hello World")
' * Dim matches As MatchCollection
' * Set matches = str.Match("\w+")
' * ' matches(0).Value = "Hello", matches(1).Value = "World"
' */
Public Function Match(ByVal pPattern As String, Optional ByVal pIgnoreCase As Boolean = True, Optional ByVal pGlobal As Boolean = True) As MatchCollection
    Dim lvReg As RegExp
    Dim i As Long
    
    Set lvReg = New RegExp
    
    With lvReg
        .Pattern = pPattern
        .IgnoreCase = pIgnoreCase
        .Global = pGlobal
    End With
    
    Set Match = lvReg.Execute(mThis.Value)
End Function

'/**
' * @description インスタンスの取得メソッド（内部使用）
' * @returns {clsEnhancedString} ミュータブルの場合は自身、非ミュータブルの場合は新しいインスタンス
' * @private
' */
Private Function GetInstance() As clsEnhancedString
    If mThis.Mutable Then
        Set GetInstance = Me
    Else
        Set GetInstance = New clsEnhancedString
    End If
End Function
