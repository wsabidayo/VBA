Attribute VB_Name = "modFactory"
Option Explicit

'/**
' * @description 拡張文字列クラスのインスタンスを作成するファクトリ関数
' * @param {String} pValue - 初期化する文字列値
' * @param {Boolean} [pMutable=False] - ミュータブル操作を許可するかどうか
' * @returns {clsEnhancedString} 初期化された拡張文字列クラスのインスタンス
' * @example
' * Dim str As clsEnhancedString
' * Set str = CreateEnhancedString("Hello", True)
' * Debug.Print str.ToUpperCase().Concat(", World!").Value ' 出力: "HELLO, World!"
' */
Public Function CreateEnhancedString(ByVal pValue As String, Optional ByVal pMutable As Boolean = False) As clsEnhancedString
    Dim lvResult As clsEnhancedString
    
    Set lvResult = New clsEnhancedString
    
    lvResult.Initialize pValue, pMutable
    
    Set CreateEnhancedString = lvResult
End Function
