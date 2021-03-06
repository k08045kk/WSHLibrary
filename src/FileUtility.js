/*!
 * FileUtility.js v17
 *
 * Copyright (c) 2018 toshi (https://github.com/k08045kk)
 *
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */

/**
 * WSH(JScript)用ファイル関連ライブラリ
 * @requires    module:ActiveXObject('Scripting.FileSystemObject')
 * @requires    module:ActiveXObject('WScript.Shell')
 * @requires    module:ActiveXObject('ADODB.Stream')
 * @requires    module:JSON
 * @requires    module:WScript
 * @auther      toshi (https://github.com/k08045kk)
 * @version     17
 * @see         1 - add - 初版
 * @see         2 - update - Move関数のMoveFile第3引数を削除 - 動作しなかったため
 * @see         3 - update - PrivateUnderscore.jsを導入 - ソフトに共通処理化するため
 * @see         4 - update - find関数関連を修正 - 単純化のため
 * @see         5 - fix - getUniqFilePath()修正 - 拡張子あり時ファイルパスでなくファイル名を返していた
 * @see         6 - update - fs, shをクラス内部確保に変更
 * @see         7 - fix - パス長が規定値を超えた場合、エラーする
 * @see         8 - fix - UTF-8(BOMあり)の3文字未満ファイルをBOMを付けたまま読込んでいた
 * @see         8 - fix - UTF-16BE(BOMなし)の文字なしテキストを読込むとハングする
 * @see         8 - update - UTF-16BEのBOM付き読み書きに対応
 * @see         9 - update - awaitDelete関数を追加
 * @see         9 - update - リファクタリング
 * @see         10 - fix - 移動先ファイルを削除直後の場合、Move関数が動作しない可能性がある
 * @see         10 - update - awaitDelete関数を削除 - 無意味だと判明したため
 * @see         11 - update - リファクタリング
 * @see         12 - update - Move関数のエラー時に待機時間を確保する
 * @see         12 - fix - パス長が規定値を超えた場合、エラーする
 * @see         13 - fix - ファイルあり時、上書きなしでエラーする
 * @see         14 - update - リファクタリング
 * @see         15 - update - fileExists, folderExists追加
 * @see         16 - update - getFileFolders追加
 * @see         16 - update - storeTextWithoutBOM, storeJSONWithoutBOM追加
 * @see         16 - update - storeTextWithBOM, storeJSONWithBOM追加
 * @see         17 - update - _FileUtility_createFolder()を再帰処理しないように修正
 */
(function(root, factory) {
  if (!root.FileUtility) {
    root.FileUtility = factory();
  }
})(this, function() {
  "use strict";
  
  var global = Function('return this')();
  var fs = new ActiveXObject('Scripting.FileSystemObject');
  var sh = void 0;
  var _this = void 0;
  
  /**
   * PrivateUnderscore.js
   * @version   7
   */
  {
    try {
      sh = new ActiveXObject('WScript.Shell');
    } catch (e) {}
    function _FileUtility_createFolder(folderpath) {
      var ret = false,
          buffer = [],
          path = fs.GetAbsolutePathName(folderpath);
      while (path != '' && !(fs.FolderExists(path) || fs.FileExists(path))) {
        buffer.push(path);
        path = fs.GetParentFolderName(path);
      }
      if (fs.FolderExists(path)) {
        while ((path = buffer.pop()) != null) {
          try {
            fs.CreateFolder(path);
            ret = true;
          } catch (e) {
            // ファイルが見つかりません(パス長問題) || パスが見つかりません(パス不正 || 存在しない)
            ret = false;
            break;
          }
        }
      } // else ファイルが存在する場合、子フォルダを作成できない
      return ret;
    };
    function _FileUtility_createFileFolder(filepath) {
      return _FileUtility_createFolder(fs.GetParentFolderName(filepath));
    };
  }
  
  /**
   * コンストラクタ
   * @constructor
   */
  _this = function FileUtility_constructor() {};
  
  // ADODB関連定義
  _this.AUTODETECT= '_autodetect_all';  // 文字コード自動検出
  _this.ASCII     = 'us-ascii';
  _this.EUC_JP    = 'euc-jp';
  _this.SHIFT_JIS = 'shift-jis';
  _this.UTF_8     = 'utf-8';
  _this.UTF_16    = 'utf-16';           // UTF-16LE
  _this.UTF_16LE  = 'utf-16le';         // UTF-16LE
  _this.UTF_16BE  = 'utf-16be';         // UTF-16BE
  _this.StreamTypeEnum = {
    adTypeBinary: 1,    // バイナリデータ
    adTypeText:   2     // テキストデータ(既定)
  };
  _this.ConnectModeEnum = {
    adModeUnknown:        0,    // 設定が不明であることを表します。(既定)
    adModeRead:           1,    // 読み取り専用の権限
    adModeWrite:          2,    // 書き込み専用の権
    adModeReadWrite:      3,    // 読み取り/書き込み両方の権限
    adModeShareDenyRead:  4,    // 読み取り権限による他ユーザー禁止
    adModeShareDenyWrite: 8,    // 書き込み権限による他ユーザー禁止
    adModeShareExclusive: 12,   // ほかのユーザーの接続を禁止します
    adModeShareDenyNone:  16    // 他のユーザーにも読み取り/書き込みの許可
  };
  _this.StreamReadEnum = {
    adReadAll:    -1,   // ストリームからすべてのバイトを読取(既定)
    adReadLine:   -2    // ストリームから次の行を読み取り
  };
  _this.StreamWriteEnum = {
    adWriteChar:  0,    // テキスト文字列を書き込みます。(既定)
    adWriteLine:  1     // テキスト文字列と行区切り文字書き込み
  };
  _this.SaveOptionsEnum = {
    adSaveCreateNotExist:   1,  // ファイルない時、新しいファイル作成(既定)
    adSaveCreateOverWrite:  2   // ファイルある時、ファイルが上書き
  }
  _this.LineSeparatorsEnum = {
    adCRLF: -1, // 改行文字: CRLF(既定)
    adLF:   10, // 改行文字: LF
    adCR:   13  // 改行文字: CR
  };
  
  // FileSystemObject関連定義
  _this.OpenTextFileIomode = {
    ForReading:   1,    // ファイルを読み取り専用として開きます
    ForWriting:   2,    // ファイルを書き込み専用として開きます
    ForAppending: 8     // ファイルを開き、ファイルの最後に追加
  };
  _this.FileAttributes = {
    Normal:     0,      // 標準ファイル
    ReadOnly:   1,      // 読み取り専用ファイル
    Hidden:     2,      // 隠しファイル
    System:     4,      // システム・ファイル
    Volume:     8,      // ディスクドライブ・ボリューム・ラベル
    Directory:  16,     // フォルダ／ディレクトリ
    Archive:    32,     // 前回のバックアップ以降に変更されていれば1
    Alias:      64,     // リンク／ショートカット
    Compressed: 128     // 圧縮ファイル
  };
  
  // FileUtility_find関数用の定義
  // order(検索順序)
  var folderorder = 0x01; // フォルダを調べる
  var fileorder   = 0x02; // ファイルを調べる
  var suborder    = 0x03; // サブフォルダを調べる
  // 前方検索
  _this.orderFront= (folderorder <<  0) | (fileorder <<  8) | (suborder << 16);
  // 後方検索
  _this.orderPost = (folderorder <<  8) | (fileorder << 16) | (suborder <<  0);
  
  // target(検索対象)
  var subfolder   = 0x01; // サブフォルダを調べる
  var filelist    = 0x02; // ファイルを調べる
  var folderlist  = 0x04; // フォルダを調べる
  var rootfolder  = 0x08; // ルートフォルダを調べる
  var branchfolder= 0x10; // 中間フォルダを調べる
  var leaffolder  = 0x20; // 末端フォルダを調べる
  // 直下ファイル
  _this.targetFiles           =           filelist;
  // サブフォルダ含む全ファイル
  _this.targetAllFiles        = subfolder|filelist;
  // 直下フォルダ
  _this.targetFolders         =                    folderlist|           branchfolder;
  // 末端フォルダ
  _this.targetLeafFolders     = subfolder|         folderlist|                        leaffolder;
  // ルート以外全フォルダ
  _this.targetSubFolders      = subfolder|         folderlist|           branchfolder;
  // ルート含む全フォルダ
  _this.targetAllFolders      = subfolder|         folderlist|rootfolder|branchfolder;
  
  /**
   * フォルダパスからフォルダを作成する
   * @param {string} folderpath - フォルダパス
   * @return {boolean} フォルダ作成有無
   */
  _this.createFolder = function FileUtility_createFolder(folderpath) {
    return _FileUtility_createFolder(folderpath);
  };
  
  /**
   * ファイルパスからフォルダを作成する
   * @param {string} filepath - ファイルパス
   * @return {boolean} フォルダ作成有無
   */
  _this.createFileFolder = function FileUtility_createFileFolder(filepath) {
    return _FileUtility_createFileFolder(filepath);
  };
  
  /**
   * 空ファイルを作成する
   * 既に存在する場合、何もしない。
   * @param {string} filepath - ファイルパス
   * @return {boolean} ファイル作成有無
   */
  _this.createFile = function FileUtility_createFile(filepath) {
    var ret = false;
    if (!_this.Exists(filepath)) {
      _this.storeText('', filepath, false, _this.ASCII);
      ret = true;
    }
    return ret;
  };
  
  /**
   * 空のフォルダを判定する
   * @param {string} folderpath - フォルダパス
   * @return {boolean} 成否
   */
  _this.isEmptyFolder = function FileUtility_isEmptyFolder(folderpath) {
    var ret = false,
        folder,
        fullpath = fs.GetAbsolutePathName(folderpath);
    if (fs.FolderExists(fullpath)) {
      folder = fs.GetFolder(fullpath);
      // ファイルなし && フォルダなし
      try {
        ret = folder.Files.Count == 0 && folder.SubFolders.Count == 0;
      } catch (e) {
        // パス長問題の場合、falseを返す
      }
    }
    return ret;
  };
  
  /**
   * 存在しないファイルパスを返す
   * @param {string} [opt_folderpath=sh.CurrentDirectory] - フォルダパス
   * @param {string} [opt_ext=''] - 拡張子
   * @return {string} 存在しないファイルパス
   */
  _this.getTempFilePath = function FileUtility_getTempFilePath(opt_folderpath, opt_ext) {
    var folderpath = (opt_folderpath != null)?
            fs.GetAbsolutePathName(opt_folderpath):
            sh.CurrentDirectory;
    var ext = (opt_ext && opt_ext.length!=0)? '.'+opt_ext: '';
    
    // 存在しないファイル名を返す
    var fullpath = '';
    do {
      fullpath = fs.BuildPath(folderpath, fs.GetTempName()+ext);
    } while (fs.FileExists(fullpath) || fs.FolderExists(fullpath));
    return fullpath;
    // 補足:fs.GetTempName()関数は、単純にランダムな文字列を返すだけで、
    //      存在有無を意識しないため、存在するかの判定をする。(未確認)
  };
  // 存在しないファイル名称を返す
  _this.getTempFileName = function FileUtility_getTempFileName(opt_folderpath, opt_ext) {
    return fs.GetFileName(_this.getTempFilePath(opt_folderpath, opt_ext));
  };
  
  /**
   * ユニークなファイルパスを返す
   * ファイル名を維持しつつ、末尾に数値を付加したユニークなファイルパスを返す。
   * @param {string} filepath - ファイルパス
   * @return {string} ユニークなファイルパス
   */
  _this.getUniqFilePath = function FileUtility_getUniqFilePath(filepath) {
    var fullpath = fs.GetAbsolutePathName(filepath);
    var folderpath = fs.GetParentFolderName(fullpath);
    var uniqpath = fullpath;
    
    if (fs.FileExists(fullpath) || fs.FolderExists(fullpath)) {
      var ext = fs.GetExtensionName(fullpath);
      if (ext.length !== 0) {
        ext = '.'+ext;
        fullpath = fs.BuildPath(folderpath, fs.GetBaseName(fullpath));
      }
      
      var idx = 1;
      do {
        idx++;
        uniqpath = fullpath + '_'+idx + ext;
      } while (fs.FileExists(uniqpath) || fs.FolderExists(uniqpath));
    }
    return uniqpath;
  };
  // ユニークなファイル名称を返す
  _this.getUniqFileName = function FileUtility_getUniqFileName(path) {
    return fs.GetFileName(_this.getUniqFilePath(path));
  };
  
  /**
   * 有効なファイル名を返す
   * 無効な文字を削除する。
   * @param {string} name - ファイル名
   * @return {string} 有効なファイル名
   */
  _this.getValidFileName = function FileUtility_getValidFileName(name) {
    var marks = ['\\','/',':','*','?','"','<','>','|'];
    for (var i=0; i<marks.length; i++) {
      name = name.replace(marks[i], '');
    }
    return name;
  };
  
  /**
   * ファイルを読み込み
   * @private
   * @param {number} type - 種別(1:Bynary/2:Text)
   * @param {string} path - ファイルパス
   * @param {string} [opt_charset='_autodetect_all'] - 文字セット
   * @return {(string|Variant)} ファイルデータ
   */
  function FileUtility_loadFile(type, path, opt_charset) {
    var ret, fullpath, 
        charset = opt_charset,
        skip = false;
    
    if (charset == null) { charset = _this.AUTODETECT; }
    charset = charset.toLowerCase();
    
    fullpath = fs.GetAbsolutePathName(path);
    if (!fs.FileExists(fullpath)) {
      // ファイルなし
      return null;
    } else if (fs.GetFile(fullpath).size === 0) {
      // 空ファイル
      return (type == _this.StreamTypeEnum.adTypeText)? '': null;
    }
    
    var sr = new ActiveXObject('ADODB.Stream');
    sr.Type = type;
    if (type == _this.StreamTypeEnum.adTypeText) {
      if (charset == _this.AUTODETECT || charset == _this.UTF_16BE) {
        // BOMを確認してUTF-8とUTF-16だけ、手動で判定する
        // UTF-16BEは、BOMあり時にBOM削除されないため、手動でスキップする
        var pre = new ActiveXObject('ADODB.Stream');
        pre.Type = _this.StreamTypeEnum.adTypeText;
        pre.Charset = _this.ASCII;
        pre.Open();
        pre.LoadFromFile(fullpath);
        var bom = [];
        bom.push(pre.EOS || escape(pre.ReadText(1)));
        bom.push(pre.EOS || escape(pre.ReadText(1)));
        bom.push(pre.EOS || escape(pre.ReadText(1)));
        if (charset == _this.UTF_16BE) {
          if (bom[0] == '%7E' && bom[1]== '%7F') {
            skip = true;
          }
        } else if (bom[0] == 'o'   && bom[1]== '%3B' && bom[2]== '%3F') {
          charset = _this.UTF_8;
        } else if (bom[0] == '%7F' && bom[1]== '%7E') {
          charset = _this.UTF_16LE;
        } else if (bom[0] == '%7E' && bom[1]== '%7F') {
          charset = _this.UTF_16BE;
          skip = true;
        }
        pre.Close();
        pre = null;
      }
      sr.Charset = charset;
    }
    
    // ファイルから読み出し
    sr.Open();
    sr.LoadFromFile(fullpath);
    if (skip) {
      // 先頭一文字(BOM)を空読み
      sr.ReadText(1);
      ret = sr.ReadText();
    } else {
      ret = (type == _this.StreamTypeEnum.adTypeText)? sr.ReadText(): sr.Read();
    }
    
    // 終了処理
    sr.Close();
    sr = null;
    return ret;
    // 備考：修正が発生した場合、PrivateUnderscore.jsの_loadText()も合わせて修正する
  };
  
  /**
   * バイナリファイルを読み込み
   * @param {string} path - ファイルパス
   * @return {Variant} byte配列
   */
  _this.loadBinary = function FileUtility_loadBinary(path) {
    return FileUtility_loadFile(_this.StreamTypeEnum.adTypeBinary, path);
  };
  
  /**
   * テキストファイルを読み込み
   * @param {string} path - ファイルパス
   * @param {string} [opt_charset='_autodetect_all'] - 文字コード
   * @return {string} テキスト
   */
  _this.loadText = function FileUtility_loadText(path, opt_charset) {
    return FileUtility_loadFile(_this.StreamTypeEnum.adTypeText, path, opt_charset);
  };
  
  /**
   * ファイルに書き込み
   * フォルダが存在しない場合、自動で作成します。
   * @private
   * @param {number} type - 種別(1:Bynary/2:Text)
   * @param {(string|Variant)} src - 書き込むデータ
   * @param {string} path - 書き込むパス
   * @param {(boolean|number)} [opt_option=1] - オプション(1:上書きなし/2:上書きあり)
   * @param {string} [opt_charset='utf-8'] - 文字コード
   * @param {boolean} [opt_bom=true] - BOM(true/false)
   * @return {boolean} 結果
   */
  function FileUtility_storeFile(type, src, path, opt_option, opt_charset, opt_bom) {
    var option  = (opt_option === true)?  _this.SaveOptionsEnum.adSaveCreateOverWrite:
                  (opt_option === false)? _this.SaveOptionsEnum.adSaveCreateNotExist:
                  (opt_option == null)?   _this.SaveOptionsEnum.adSaveCreateNotExist: opt_option;
    var charset = (opt_charset== null)?   _this.UTF_8: opt_charset;
    var bom     = (opt_bom    == null)?   true: opt_bom;
    var fullpath, skip, bin;
    var ret = true;
    
    // 前処理
    charset = charset.toLowerCase();
    skip = {};
    skip[_this.UTF_8] = 3;
    skip[_this.UTF_16] = 2;
    skip[_this.UTF_16LE] = 2;
    // UTF-16BEは、スキップ不要(ADODB.StreamがBOMを書き込まないため)
    fullpath = fs.GetAbsolutePathName(path);
    
    // (存在しない場合)フォルダを作成する
    _this.createFileFolder(fullpath);
    
    // ファイルに書き込む。
    var sr = new ActiveXObject('ADODB.Stream');
    sr.Type = type;
    if (type == _this.StreamTypeEnum.adTypeText) {
      sr.Charset = charset;
      sr.Open();
      sr.WriteText(src);
      if (bom === true && charset == _this.UTF_16BE) {
        // ADODB.Streamは、UTF-16BEのBOMを書き込まないため、自力でBOMを書き込む
        // LEのBOMを確保
        var le = new ActiveXObject('ADODB.Stream');
        le.Type = _this.StreamTypeEnum.adTypeText;
        le.Charset = _this.UTF_16LE;
        le.Open();
        le.WriteText('');
        le.Position = 0;
        le.Type = _this.StreamTypeEnum.adTypeBinary;
        
        // BEのバイナリを確保
        var be = sr;
        be.Position = 0;
        be.Type = _this.StreamTypeEnum.adTypeBinary;
        bin = be.Read();
        be.Close();
        be  = null;
        sr  = null;
        
        // 再度BOMありを書き込み
        sr = new ActiveXObject('ADODB.Stream');
        sr.Type = _this.StreamTypeEnum.adTypeBinary;
        sr.Open();
        
        // BOM(LEの1Byteと2Byteが逆)を書き込み
        // BEのバイナリを書き込み
        le.Position = 1;
        sr.Write(le.Read(1));
        le.Position = 0;
        sr.Write(le.Read(1));
        if (bin != null)  sr.Write(bin);
        
        le.Close();
        le  = null;
      }
      if (bom === false && skip[charset]) {
        // BOMなし書込処理
        var pre = sr;
        pre.Position = 0;
        pre.Type = _this.StreamTypeEnum.adTypeBinary;
        // skipバイト(BOM)を読み飛ばす
        pre.Position = skip[charset];
        bin = pre.Read();
        pre.Close();
        pre = null;
        sr  = null;
        
        // 再度BOMなしを書き込み
        sr = new ActiveXObject('ADODB.Stream');
        sr.Type = _this.StreamTypeEnum.adTypeBinary;
        sr.Open();
        if (bin != null)  sr.Write(bin);
      }
    } else {
      // 上記以外(バイナリ)の時
      sr.Open();
      sr.Write(src);
    }
    try {
      sr.SaveToFile(fullpath, option);
    } catch (e) {       // ADODB.Stream: ファイルへ書き込めませんでした。
      // ファイルあり時、上書きなし
      ret = false;
    }
    sr.Close();
    sr = null;
    return ret;
    // 補足:LineSeparatorプロパティは、全行読み出しのため、無意味
  };
  
  /**
   * バイナリファイルを書き込む
   * @param {Variant} bytes - 書き込むデータ
   * @param {string} path - ファイルへのパス
   * @param {number} [opt_option=1] - オプション(1:上書きなし/2:上書きあり)
   * @return {boolean} 結果
   */
  _this.storeBinary = function FileUtility_storeBinary(bytes, path, opt_option) {
    return FileUtility_storeFile(_this.StreamTypeEnum.adTypeBinary, bytes, path, opt_option);
  };
  
  /**
   * テキストファイルに書き込む
   * 注意：BOMの標準をtrueからfalseに変更する可能性があります。
   *       BOM有無を意識する場合、明示的に指定するか専用関数を使用してください。
   * @param {string} text - 書き込むデータ
   * @param {string} path - ファイルへのパス
   * @param {number} [opt_option=1] - オプション(1:上書きなし/2:上書きあり)
   * @param {string} [opt_charset='utf-8'] - 文字コード
   * @param {boolean} [opt_bom=true] - BOM(true/false)
   * @return {boolean} 結果
   */
  _this.storeText = function FileUtility_storeText(text, path, opt_option, opt_charset, opt_bom) {
    return FileUtility_storeFile(_this.StreamTypeEnum.adTypeText,text,path,opt_option,opt_charset, opt_bom);
  };
  _this.storeTextWithBOM = function FileUtility_storeTextWithBOM(text, path, opt_option, opt_charset) {
    return _this.storeText(text, path, opt_option, opt_charset, true);
  };
  _this.storeTextWithoutBOM = function FileUtility_storeTextWithoutBOM(text, path, opt_option, opt_charset) {
    return _this.storeText(text, path, opt_option, opt_charset, false);
  };
  
  /**
   * 書き込み禁止を判定します。
   * @param {string} path - ファイルへのパス
   * @return {boolean} 結果(true:書き込み禁止)
   */
  _this.isWriteProtected = function FileUtility_isWriteProtected(path) {
    var fullpath = fs.GetAbsolutePathName(path);
    return fs.FileExists(fullpath) 
        && !!(fs.GetFile(fullpath).Attributes & _this.FileAttributes.ReadOnly);
  };
  
  /**
   * 検索ファイルフォルダを判定する
   * @callback findCallback
   * @param {string} fullpath - ファイルフォルダのフルパス
   * @param {boolean} folder - フォルダ(true:フォルダ/false:ファイル)
   * @return {boolean} 結果(true:処理中断/false:なし)
   */
  
  /**
   * ファイル/フォルダ検索
   * callback関数がtrueと解釈できる値を返した場合、
   * 処理を中断してcallback関数の戻り値を戻す。
   * Windowsの最大パス長(260)の場合、該当ファイル/フォルダは呼び出されない。
   * @private
   * @param {findCallback} callback - コールバック関数
   * @param {string} path - 検索開始パス
   * @param {number} target - 検索対象(targetFiles/targetAllFiles/...)
   * @param {number} order - 検索順序(orderFront:前方検索/orderPost:後方検索)
   * @param {Object} opt_extensionSet - 拡張子セット
   * @return {*} callback関数の戻り値
   */
  function FileUtility_findMain(callback, fullpath, target, order, opt_extensionSet) {
    var ret = void 0,
        callee = FileUtility_findMain,
        folder, flag, e, path, ext, ntarget, isLeaf;
    
    if (fs.FolderExists(fullpath)) {
      folder = fs.GetFolder(fullpath);
      // 検索順序の残件がある時 && フォルダが存在する時(削除されることもある) && 中断でない時
      for (flag=order; flag!=0 && fs.FolderExists(fullpath) && !ret; flag=flag>>>8) {
        switch (flag & 0xff) {
        case folderorder:       // フォルダ
          if (target & folderlist) {
            if (target & rootfolder) {
              ret = callback(fullpath, true);
            } else if (target & leaffolder) {
              isLeaf = false
              try {
                isLeaf = (folder.SubFolders.Count == 0);
              } catch (error) {}// パス長問題
              if (isLeaf) {
                ret = callback(fullpath, true);
              }
            }
          }
          break;
        case fileorder:         // ファイル
          if (target & filelist) {
            if (!opt_extensionSet) {
              for (e=new Enumerator(folder.Files); !e.atEnd() && !ret; e.moveNext()) {
                ret = callback(e.item().Path, false);
              }
              e = null;
            } else {
              for (e=new Enumerator(folder.Files); !e.atEnd() && !ret; e.moveNext()) {
                path = e.item().Path;
                ext = fs.GetExtensionName(path).toLowerCase();
                if (opt_extensionSet[ext] === true) {
                  ret = callback(path, false);
                }
              }
              e = null;
            }
          }
          break;
        case suborder:          // サブフォルダ
          if (target & subfolder) {
            for (e=new Enumerator(folder.SubFolders); !e.atEnd() && !ret; e.moveNext()) {
              path = e.item().Path;
              ntarget = (target & branchfolder)? target|rootfolder: target&(~rootfolder);
                                // 枝を追加する場合、ルート(枝)追加
                                // 上記以外、ルートを削除
              ret = callee(callback, path, ntarget, order, opt_extensionSet);
                                // サブフォルダ(再帰呼出し)
            }
            e = null;
          } else if ((target & folderlist) && (target & branchfolder)) {
            for (e=new Enumerator(folder.SubFolders); !e.atEnd() && !ret; e.moveNext()) {
              // サブなしフォルダ一覧
              ret = callback(e.item().Path, true);
            }
            e = null;
          }
          break;
        }
      }
    } else if (fs.FileExists(fullpath)) {
      if (target & filelist) {  // ファイル一覧
        ret = callback(fullpath, false);
      }
    }
    return ret;
  };
  /**
   * ファイル/フォルダ検索
   * @param {findCallback} callback - コールバック関数
   * @param {string} path - 検索開始パス
   * @param {number} target - 検索対象(targetFiles/targetAllFiles/...)
   * @param {number} order - 検索順序(orderFront:前方検索/orderPost:後方検索)
   * @param {string[]} opt_extensions - 拡張子(例:['jpg', 'png'])
   * @return {boolean} callback関数の戻り値
   */
  _this.find = function FileUtility_find(callback, path, target, opt_order, opt_extensions) {
    var order = (opt_order)? opt_order: _this.orderFront;
    var fullpath = fs.GetAbsolutePathName(path);
    var extensionSet = void 0;
    if (opt_extensions) {
      extensionSet = {};
      for (var i=0; i<opt_extensions.length; i++) {
        extensionSet[opt_extensions[i].toLowerCase()] = true;
      }
    }
    return FileUtility_findMain(callback, fullpath, target, order, extensionSet);
  };
  // 直下ファイル検索
  _this.findFile = function FileUtility_findFile(callback, path, opt_order, opt_extensions) {
    return _this.find(callback, path, _this.targetFiles, opt_order, opt_extensions);
  };
  // サブフォルダ含む全ファイル検索
  _this.findAllFile = function FileUtility_findAllFile(callback, path, opt_order, opt_extensions) {
    return _this.find(callback, path, _this.targetAllFiles, opt_order, opt_extensions);
  };
  // 直下フォルダ検索
  _this.findFolder = function FileUtility_findFolder(callback, path, opt_order) {
    return _this.find(callback, path, _this.targetFolders, opt_order);
  };
  // 末端フォルダ検索
  _this.findLeafFolder = function FileUtility_findLeafFolder(callback, path, opt_order) {
    return _this.find(callback, path, _this.targetLeafFolders, opt_order);
  };
  // ルート以外全フォルダ検索
  _this.findSubFolder = function FileUtility_findSubFolder(callback, path, opt_order) {
    return _this.find(callback, path, _this.targetSubFolders, opt_order);
  };
  // ルート含む全フォルダ検索
  _this.findAllFolder = function FileUtility_findAllFolder(callback, path, opt_order) {
    return _this.find(callback, path, _this.targetAllFolders, opt_order);
  };
  
  /**
   * ファイルパスの一覧を返す
   * Windowsの最大パス長(260)の場合、該当ファイルは配列に含まれない。
   * @param {string} folderpath - フォルダパス
   * @return {string[]} ファイルパスの配列
   */
  _this.getFiles = function FileUtility_getFiles(folderpath) {
    var list = [];
    var fullpath = fs.GetAbsolutePathName(folderpath);
    if (fs.FolderExists(fullpath)) {
      var e;
      for (e=new Enumerator(fs.GetFolder(fullpath).Files); !e.atEnd(); e.moveNext()) {
        list.push(e.item().Path);
      }
      e = null;
    }
    return list;
  };
  
  /**
   * フォルダパスの一覧を返す
   * Windowsの最大パス長(260)の場合、該当フォルダは配列に含まれない。
   * @param {string} folderpath - フォルダパス
   * @return {string[]} フォルダパスの配列
   */
  _this.getFolders = function FileUtility_getFolders(folderpath) {
    var list = [];
    var fullpath = fs.GetAbsolutePathName(folderpath);
    if (fs.FolderExists(fullpath)) {
      var e;
      for (e=new Enumerator(fs.GetFolder(fullpath).SubFolders); !e.atEnd(); e.moveNext()){
        list.push(e.item().Path);
      }
      e = null;
    }
    return list;
  };
  
  /**
   * ファイルフォルダパスの一覧を返す
   * Windowsの最大パス長(260)の場合、該当フォルダは配列に含まれない。
   * @param {string} folderpath - フォルダパス
   * @return {string[]} ファイルフォルダパスの配列
   */
  _this.getFileFolders = function FileUtility_getFileFolders(folderpath) {
    return _this.getFiles(folderpath).concat(_this.getFolders(folderpath));
  };
  
  /**
   * 存在確認
   * @param {string} src - 移動前のパス(例:'C:\\A\\B.txt')
   * @param {string} dst - 移動後のパス(例:'C:\\B\\C.txt')
   * @return {boolean} 複製完了有無
   */
  _this.exists =
  _this.Exists = function FileUtility_Exists(path) {
    var fullpath = fs.GetAbsolutePathName(path);
    return (fs.FileExists(fullpath) || fs.FolderExists(fullpath));  // ファイル/フォルダあり
  };
  _this.fileExists =
  _this.FileExists = function FileUtility_FileExists(path) {
    return fs.FileExists(fs.GetAbsolutePathName(path));
  };
  _this.folderExists =
  _this.FolderExists = function FileUtility_FolderExists(path) {
    return fs.FolderExists(fs.GetAbsolutePathName(path));
  };
  
  /**
   * 複製
   * @param {string} src - 複製前のパス(例:'C:\\A\\B.txt')
   * @param {string} dst - 複製後のパス(例:'C:\\B\\C.txt')
   * @param {boolean} [opt_overwite=false] - 複製先にファイルがある場合、上書きする
   * @return {boolean} 複製完了有無
   */
  _this.copy =
  _this.Copy = function FileUtility_Copy(src, dst, opt_overwite) {
    var overwite = (opt_overwite === true);
    var ret = false;
    
    // 誤動作防止用(''は、カレントディレクトリ扱いになる)
    if (!src || src == '' || !dst || dst == '') {  return ret; }
    try {
      src = fs.GetAbsolutePathName(src);
      dst = fs.GetAbsolutePathName(dst);
      if (_this.Exists(dst) && !overwite) {     // 移動先が存在する
      } else if (fs.FileExists(src)) {          // 移動元ファイルがある
        _this.createFileFolder(dst);
        fs.CopyFile(src, dst, overwite);
        ret = true;
      } else if (fs.FolderExists(src)) {        // 移動元フォルダがある
        _this.createFileFolder(dst);
        fs.CopyFolder(src, dst, overwite);
        ret = true;
      }
    } catch (e) {}
    return ret;
  };
  
  /**
   * 移動
   * 移動先のファイル/フォルダがある場合、移動は失敗します。
   * ボリューム(ドライブ)間のフォルダ移動は、動作しないことがあります。
   * ボリューム(ドライブ)間のフォルダ移動は、Copy()とDelete()で代用可能です。
   * @param {string} src - 移動前のパス(例:'C:\\A\\B.txt')
   * @param {string} dst - 移動後のパス(例:'C:\\B\\C.txt')
   * @return {boolean} 移動完了有無
   */
  _this.move =
  _this.Move = function FileUtility_Move(src, dst) {
    var ret = false;
    
    // 誤動作防止用(''は、カレントディレクトリ扱いになる)
    if (!src || src == '' || !dst || dst == '') {  return ret; }
    src = fs.GetAbsolutePathName(src);
    dst = fs.GetAbsolutePathName(dst);
    while (!_this.Exists(dst)) {                // 移動先が存在しない
      try {
        if (fs.FileExists(src)) {               // 移動元ファイルがある
          _this.createFileFolder(dst);
          fs.MoveFile(src, dst);
          ret = true;
        } else if (fs.FolderExists(src)) {      // 移動元フォルダがある
          _this.createFileFolder(dst);
          fs.MoveFolder(src, dst);
          ret = true;
        }
      } catch (e) {
        if (e.number == -2146828230) {
          // Error: 既に同名のファイルが存在しています。
          // 移動先が存在しないことは、確認済みのため、再度繰り返す。
          // 置き換え(削除直後に移動)で本現象が多発する。
          // Existsの挙動が、IF上のファイルの存在有無であり、
          // 内部処理上の存在有無ではない可能性がある。
          // そのため、内部処理が完了するまで繰り返す。
          if ('WScript' in global) {
            WScript.Sleep(10);
          }
          continue;
        }
        // 書き込みできません。(-2146828218)
        // Windowsの最大パス長(260)以上のファイルが含まれる
      }
      break;
    }
    return ret;
  };
  
  /**
   * 名称変更
   * 移動先のファイル/フォルダがある場合、移動は失敗します。
   * @param {string} src - 移動前のパス(例:'C:\\A\\B.txt')
   * @param {string} name - 変更後の名称
   * @return {boolean} 変更完了有無
   */
  _this.rename =
  _this.Rename = function FileUtility_Rename(src, name) {
    var parent= fs.GetParentFolderName(src);
    var dist  = fs.BuildPath(parent, name);
    return _this.Move(src, dist);               // リネーム(移動)
  };
  
  /**
   * ファイル/フォルダ削除
   * ファイル/フォルダを削除します。ゴミ箱に移動するわけではありません。
   * 最初のエラーが発生した時点で処理を中止します。
   * エラーが発生するまでに行った処理を、取り消したり元に戻したりする処理は一切行われません。
   * 例: opt_force=falseで読み取り専用ファイルが含まれるフォルダを削除する。
   *     この場合、戻り値は、falseとなります。
   *     ただし、削除の順序によっては、読み取り専用でないファイルは削除されます。
   * @param {string} path - パス
   * @param {boolean} [opt_force=false] - 読み取り専用を削除する
   * @return {boolean} 削除完了有無
   */
  _this.del =
  _this.Delete = function FileUtility_Delete(path, opt_force) {
    var ret = false,
        force = (opt_force === true);
    
    // 誤動作防止用(''は、カレントディレクトリ扱いになる)
    if (!path || path == '') {  return ret; }
    try {
      var fullpath = fs.GetAbsolutePathName(path);
      if (fs.FileExists(fullpath)) {
        fs.DeleteFile(fullpath, force);
        ret  = true;
      } else if (fs.FolderExists(fullpath)) {
        fs.DeleteFolder(fullpath, force);
        ret = true;
      }
    } catch (e) {}
    return ret;
  };
  // FileUtility.delete(...);は、IE8-で使用不可
  // FileUtility['delete'](...);は、使用可能
  // ただし、WSHはIE8環境のため、本ライブラリではdelete()を準備しない
  
  /**
   * JSONファイルを読み込んで返します。
   * @param {string} path - ファイルパス
   * @param {string} [opt_charset='_autodetect_all'] - 文字コード
   * @return {Object} 読み込んだJSONデータ(null:読み込み失敗)
   */
  _this.loadJSON = function FileUtility_loadJSON(path, opt_charset) {
    if (global.JSON) {
      var ret  = null;
      var text = _this.loadText(path, opt_charset);
      if (text && text.length !== 0) {
        try {
          ret = JSON.parse(text);
        } catch (e) {}
      }
      return ret;
    } else {
      throw new ReferenceError('The variable "JSON" is not declared.');
    }
  };
  
  /**
   * JSONファイルを書き込みます。
   * 注意：BOMの標準をtrueからfalseに変更する可能性があります。
   *       BOM有無を意識する場合、明示的に指定するか専用関数を使用してください。
   * @param {Object} json - 書き込むデータ
   * @param {string} path - ファイルへのパス
   * @param {number} [opt_option=1] - オプション(1:上書きなし/2:上書きあり)
   * @param {string} [opt_charset='utf-8'] - 文字コード
   * @param {boolean} [opt_bom=true] - BOM(true/false)
   * @return {boolean} 結果
   */
  _this.storeJSON = function FileUtility_storeJSON(json, path, opt_option, opt_charset, opt_bom) {
    if (global.JSON) {
      return _this.storeText(JSON.stringify(json), path, opt_option, opt_charset, opt_bom);
    } else {
      throw new ReferenceError('The variable "JSON" is not declared.');
    }
  };
  _this.storeJSONWithBOM = function FileUtility_storeJSONWithBOM(json, path, opt_option, opt_charset) {
    return _this.storeJSON(json, path, opt_option, opt_charset, true);
  };
  _this.storeJSONWithoutBOM = function FileUtility_storeJSONWithoutBOM(json, path, opt_option, opt_charset) {
    return _this.storeJSON(json, path, opt_option, opt_charset, false);
  };
  
  return _this;
  
  // 覚書1: Windowsの最大パス長(260)対策
  // 最大パス長のファイル/フォルダを処理する場合、エラーが発生する場合がある。
  // fs.FileExists(), fs.FolderExists(), fs.GetAbsolutePathName()は、エラーとならない。
  // fs.FileExists(), fs.FolderExists()は、ファイル/フォルダがある場合でも、falseを返す。
  // fs.GetFolder(fullpath).Files.Countは、エラーとなる可能背がある。
  // (以前エラーしたものがエラーしなくなている。100%発生ではない？)
  // new Enumerator(fs.GetFolder(fullpath).Files).atEnd()は、エラーとならない。
  // ただし、ファイルがある場合でも、trueを返す。
  // (パス長問題に該当しないファイルは、取得できる(atEnd()でtrueとならない))
  // 本ライブラリでは、Windowsの最大パス長のファイル/フォルダを指定した場合でも、
  // ハングしないように設計する。(そのため、最大パス長のファイル/フォルダを処理しないことがある)
  //
  // 問題環境作成方法(下記の操作をエクスプローラ上で行う)
  // ./a/b            # フォルダを作成する
  // ./a123456789/b   # 全体のパス長が260桁を超えるように、親フォルダの名称変更する
  //                  # 全体のパス長が260桁を超えるように、作成はできない
  //                  # 全体のパス長が260桁を超えるように、名称変更はできない
  
  // 覚書2: Existsの結果について
  // Exists(fs.FileExists/fs.FolderExists)の結果は、
  // まだファイルが存在する時点でも、ファイルが存在しない結果(false)となることがある。
  // 上記は、Delete直後にExists(false)を確認し、Moveを実行した時、
  // 「既に同名のファイルが存在しています。」エラーが発生したため、発覚した。
  // （エラー後、Moveの再実行を繰り返すとMoveが動作する）
  // このため、Existsの挙動が、IF上のファイルの存在有無であり、
  // 内部処理上の存在有無ではない可能性がある。
  
  // 覚書3: ""の扱いについて
  // ""(空文字)は、パスとしては、カレントディレクトリを表す。
  // ""を間違って入力して削除処理等が動作すると困るため、本ライブラリでは""を処理しない。
  
  // 覚書4: ドライブ直下の親フォルダ
  // fs.GetParentFolderName()をドライブ直下（例: C:）で使用すると""（空文字）を出力する。
  // ""は、カレントディレクトリを表すため、いろいろと処理が狂う可能性がある。
  
  // 覚書5: ファイル・フォルダ名に対象文字を含む場合、処理できないことがある。
  // Right Double Quotation Mark (”: 8221, U+201D)
  // ？ Fullwidth Question Mark (？: 65311, U+FF1F)
  // move関数で確認済み（なぜか、exists, copy, delete関数は動作した）
  
});
