# VBA EnhancedString

モダンなプログラミング言語のような文字列操作機能をVBAで実現するクラスライブラリです。

[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## 概要

VBAの標準文字列処理機能は限られており、多くの操作では複数のステップが必要になります。このライブラリは、JavaScriptやC#などの現代的なプログラミング言語で一般的な文字列操作メソッドを提供し、コードの可読性と生産性を向上させます。

## 特徴

- **モダンな文字列操作メソッド** - JavaScriptやC#のような現代的な文字列操作APIを提供
- **メソッドチェーン** - メソッドを連続して呼び出すことができるフルエントインターフェース
- **イミュータブル/ミュータブル** - 必要に応じて変更可能・不可能なインスタンスを選択可能
- **正規表現サポート** - パターンマッチングや置換に正規表現を使用可能
- **JSDoc形式のドキュメント** - すべてのメソッドとプロパティは詳細なドキュメントコメント付き

## 使用例

```vb
' 新しいインスタンスを作成
Dim str As New clsEnhancedString
str.Initialize "  Hello World  "

' メソッドチェーンを使用した連続操作
Dim result As String
result = str.Trim().ToUpperCase().Replace("WORLD", "VBA").Concat("!").Value

Debug.Print result  ' 出力: "HELLO VBA!"

' 検索操作
If str.Includes("World") Then
    Debug.Print "「World」が含まれています"
End If

' 正規表現による検証
Dim email As New clsEnhancedString
email.Initialize("test@example.com")

If email.Test("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") Then
    Debug.Print "有効なメールアドレスです"
End If
```

## インストール方法

1. このリポジトリから`clsEnhancedString.cls`をダウンロード
2. VBAプロジェクトに`clsEnhancedString.cls`をインポート
3. VBEのツール→参照設定から「Microsoft VBScript Regular Expressions 5.5」を選択

## 主要なメソッド

- `Concat` - 文字列の連結
- `ToUpperCase`/`ToLowerCase` - 大文字/小文字変換
- `Trim`/`LTrim`/`RTrim` - 空白の削除
- `Substring` - 部分文字列の取得
- `Includes`/`IndexOf` - 検索操作
- `StartsWith`/`EndsWith` - 先頭/末尾の確認
- `Replace` - パターン置換
- `Test`/`Match` - 正規表現による検証とマッチング
- その他多数のメソッド

## ライセンス

MITライセンスの下で公開されています。詳細は[LICENSE](LICENSE)ファイルをご参照ください。
