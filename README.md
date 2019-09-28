WSHLibrary
==========

A library for WSH(Windows Script Host).


## Description
+ FileUtility.js - WSH(JScript)用ファイルライブラリ
	+ ファイル／フォルダの作成
	+ ファイルの読み込み／書き込み
	+ ファイル／フォルダの検索
	+ ファイル／フォルダのコピー／移動／削除
+ ErrorUtility.js - WSH(JScript)用エラーライブラリ
	+ エラーの文字列化
	+ 実行履歴の文字列化
	+ stack, captureStackTraceの疑似実現
+ ErrorUtility.FileInfo.js - WSH(JScript)用エラーファイル情報付加
	+ エラーにファイル情報を付加
+ Console.js - WSH(JScript)用コンソール
	+ consoleの疑似実現
	+ ログ出力用の機能提供
+ Console.Animation.js - WSH(JScript)用コンソール
	+ アニメーション機能の実現
	+ ただし、標準出力の1行文字列に限る
+ EncodeUtility.js - WSH(JScript)用変換ライブラリ
	+ 文字列／16進数文字列／Base64文字列→バイト配列
	+ バイト配列→文字列／16進数文字列／Base64文字列
	+ md5／sha1のハッシュ計算
	+ Rijndael暗号(AES)の暗号化／復号
+ EasyCrypto.js - WSH(JScript)用簡易暗号化／復号
	+ 簡易暗号化／復号機能の提供
+ HTTPUtility.js - WSH(JScript)用HTTPライブラリ
	+ HTTP/HTTPSの取得
	+ HTML/XMLの解析
+ ShellUtility.js - WSH(JScript)用シェル関連
	+ 送る、ゴミ箱へ送る、zip圧縮、zip解凍
+ WebBrowser.js - WSH(JScript)用ブラウザ操作補助
	+ ブラウザ操作補助
	+ 操作ログの出力/キャッシュ
+ WMIUtility.js - WSH(JScript)用WMIライブラリ
	+ wmiObjectの取得
	+ wmiObjectをObject化して取得
+ DBLite.js - WSH(JScript)用簡易データベース
	+ SQLite(DB)の操作

元々ブログで公開していたコードです。使用例等は、ブログの方を参照願います。  
https://www.bugbugnow.net/search/label/WSHLibrary


## Copyright
Copyright (c) 2018 toshi (https://www.bugbugnow.net/p/profile.html)  
Released under the MIT license.  
see https://opensource.org/licenses/MIT
