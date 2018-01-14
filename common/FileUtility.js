/*!
 * WSH(JScript)用ファイル関連ライブラリ
 * @auther      toshi(http://www.bugbugnow.net/)
 * @license     MIT License
 * @version     1
 */
(function(factory) {
  var global = Function('return this')();
  var fs = new ActiveXObject('Scripting.FileSystemObject');
  var sh = new ActiveXObject('WScript.Shell');
  global.FileUtility= global.FileUtility || factory(global, fs, sh);
  global.fu = global.fu || global.FileUtility;
})(function FileUtility_factory(global, fs, sh) {
  "use strict";
  
  var _this = void 0;
  _this = function FileUtility_constructor() {};
  
  var separator   = '\\';
  
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
    adModeShareDenyNone:  16   // 他のユーザーにも読み取り/書き込みの許可
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
    Normal:     0,      // 読み取り専用ファイル
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
  // target(検索対象)
  var subfolder   = 0x01; // サブフォルダを調べる
  var filelist    = 0x02; // ファイルを調べる
  var folderlist  = 0x04; // フォルダを調べる
  var rootfolder  = 0x08; // ルートフォルダを調べる
  var branchfolder= 0x10; // 中間フォルダを調べる
  var leaffolder  = 0x20; // 末端フォルダを調べる
  // order(検索順序)
  var folderorder = 0x01; // フォルダを調べる
  var fileorder   = 0x02; // ファイルを調べる
  var suborder    = 0x03; // サブフォルダを調べる
  // 前方検索
  _this.orderFront= (folderorder <<  0) | (fileorder <<  8) | (suborder << 16);
  // 後方検索
  _this.orderPost = (folderorder <<  8) | (fileorder << 16) | (suborder <<  0);
  // 直下のファイル
  _this.targetFiles           =           filelist;
  // サブフォルダを含む全ファイル
  _this.targetAllFiles        = subfolder|filelist;
  // 直下のフォルダ
  _this.targetFolders         =                    folderlist|           branchfolder;
  // 末端フォルダ
  _this.targetLeafFolders     = subfolder|         folderlist|                        leaffolder;
  // ルート以外の全フォルダ
  _this.targetSubFolders      = subfolder|         folderlist|           branchfolder;
  // ルートを含む全フォルダ
  _this.targetAllFolders      = subfolder|         folderlist|rootfolder|branchfolder;
  
  /**
   * フォルダパスからフォルダを作成する
   * @param {string} folderpath         フォルダパス
   * @return {boolean}                  フォルダ作成有無
   */
  _this.createFolder = function FileUtility_createFolder(folderpath) {
    var callee = FileUtility_createFolder,
        ret = false,
        fullpath = fs.GetAbsolutePathName(folderpath);
    if (!_this.Exists(fullpath)) {
      var parentpath = fs.GetParentFolderName(fullpath);
      // ドライブ直下ではない時
      if (parentpath != '') {
        // 再帰：親フォルダ作成
        callee(fs.GetParentFolderName(fullpath));
      }
      fs.CreateFolder(fullpath);
      ret = true;
    }
    return ret;
  };
  
  /**
   * ファイルパスからフォルダを作成する
   * @param {string} filepath           ファイルパス
   * @return {boolean}                  フォルダ作成有無
   */
  _this.createFileFolder = function FileUtility_createFileFolder(filepath) {
    var ret = false;
    var fullpath  = fs.GetAbsolutePathName(filepath);
    var parentpath= fs.GetParentFolderName(fullpath);
    if (parentpath != '') {
      ret = _this.createFolder(parentpath);
    }
    return ret;
  };
  
  
  /**
   * 空ファイルを作成する
   * 既に存在する場合、何もしない。
   * @param {string} filepath           ファイルパス
   * @return {boolean}                  ファイル作成有無
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
   * 空のフォルダパス判定する
   * @param {string} folderpath         フォルダパス
   * @return {boolean}                  成否
   */
  _this.isEmptyFolder = function FileUtility_isEmptyFolder(folderpath) {
    var ret = false,
        folder,
        fullpath = fs.GetAbsolutePathName(folderpath);
    if (fs.FolderExists(fullpath)) {
      folder = fs.GetFolder(fullpath);
      // ファイルなし && フォルダなし
      ret = ((folder.Files.Count == 0) && (folder.SubFolders.Count == 0));
    }
    return ret;
  };
  
  /**
   * 存在しないファイルパスを返す。
   * @param {string} [opt_folderpath=sh.CurrentDirectory]       フォルダパス
   * @param {string} [opt_ext='']       拡張子
   * @return {string}                   存在しないファイルパス
   */
  _this.getTempFilePath = function FileUtility_getTempFilePath(opt_folderpath, opt_ext) {
    var folderpath = (opt_folderpath)?
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
    //      存在有無を意識しないため、存在するかの判定をする。
  };
  _this.getTempFileName = function FileUtility_getTempFileName(opt_folderpath, opt_ext) {
    return fs.GetFileName(_this.getTempFilePath(opt_folderpath, opt_ext));
  };
  
  /**
   * ユニークなファイルパスを返す。
   * ファイル名を維持しつつ、末尾に数値を付加したユニークなファイルパスを返す。
   * @param {string} path               フォルダパス
   * @return {string}                   ユニークなファイルパス
   */
  _this.getUniqFilePath = function FileUtility_getUniqFilePath(path) {
    var uniqpath = fs.GetAbsolutePathName(path);        // フルフォルダパス
    
    if (fs.FileExists(uniqpath) || fs.FolderExists(uniqpath)) {  // ファイル存在チェック
      var fullpath = uniqpath;
      var ext = fs.GetExtensionName(uniqpath);          // 拡張子
      if (ext.length !== 0) {
        ext = '.'+ext;
        fullpath = fs.GetBaseName(fullpath);
      }
      
      var idx = 1;
      do {
        idx++;
        uniqpath = fullpath + '_'+idx + ext;
      } while (fs.FileExists(uniqpath) || fs.FolderExists(uniqpath));
    }
    return uniqpath;
  };
  _this.getUniqFileName = function FileUtility_getUniqFileName(path) {
    return fs.GetFileName(_this.getUniqFilePath(path));
  };
  
  /**
   * 有効なファイル名を返す。
   * 無効な文字を削除します。
   * @param {string} name               ファイル名
   * @return {string}                   有効なファイル名
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
   * @param {integer} type              種別(1:Bynary/2:Text)
   * @param {string} path               ファイルパス
   * @param {string} [opt_charset='_autodetect_all']    文字セット
   * @return {string|Variant}           ファイルデータ
   */
  function FileUtility_loadFile(type, path, opt_charset) {
    var __this = _this;  // jsc.exe(errorJS1187対策)
    var charset = (opt_charset != null)? opt_charset: __this.AUTODETECT;
    var ret,
        sr, pre, bom, fullpath;
    
    fullpath = fs.GetAbsolutePathName(path);
    if (fs.FileExists(fullpath) === false) {
      // ファイルなしの時
      return null;
    } else if (fs.GetFile(fullpath).size === 0) {
      // ファイルの中身なしの時
      return (type == __this.StreamTypeEnum.adTypeText)? '': null;
    }
    
    sr = new ActiveXObject('ADODB.Stream');
    sr.Type = type;
    
    if (type == __this.StreamTypeEnum.adTypeText) {
      if (charset == __this.AUTODETECT) {
        // BOMを確認してUTF-8とUTF-16だけ、手動で判定する
        pre = new ActiveXObject('ADODB.Stream');
        pre.Type = __this.StreamTypeEnum.adTypeText;
        pre.Charset = __this.ASCII;
        pre.Open();
        pre.LoadFromFile(fullpath);
        bom = pre.ReadText(3);
        if (bom.length < 2) {
        } else if (escape(bom.substr(0, 2)) == '%7F%7E') {
          charset = __this.UTF_16LE;
        } else if (escape(bom.substr(0, 2)) == '%7E%7F') {
          charset = __this.UTF_16BE;
        } else if (bom.length < 3) {
        } else if (escape(bom) == 'o%3B%3F') {
          charset = __this.UTF_8;
        }
        pre.Close();
        pre = null;
      }
      sr.Charset = charset;
    }
    
    // ファイルから読み出し
    sr.Open();
    sr.LoadFromFile(fullpath);
    ret = (type == __this.StreamTypeEnum.adTypeText)? sr.ReadText(): sr.Read();
    
    // 終了処理
    sr.Close();
    sr = null;
    return ret;
  }
  
  /**
   * バイナリファイルを読み込み
   * @param {string} path               ファイルパス
   * @return {Variant}                  byte配列
   */
  _this.loadBinary = function FileUtility_loadBinary(path) {
    return FileUtility_loadFile(_this.StreamTypeEnum.adTypeBinary, path);
  };
  
  /**
   * テキストファイルを読み込み
   * @param {string} path               ファイルパス
   * @param {string} [opt_charset='_autodetect_all']    文字コード
   * @return {string}                   ファイルデータ(文字列)
   */
  _this.loadText = function FileUtility_loadText(path, opt_charset) {
    return FileUtility_loadFile(_this.StreamTypeEnum.adTypeText, path, opt_charset);
  };
  
  /**
   * ファイルに書き込み
   * フォルダが存在しない場合、自動で作成します。
   * 文字コードがUTF-16BEの場合、BOM=trueとしてもBOMが書き込まれない。
   * ADODB.Stream側が書き込んでくれない。JScript側で書き込むのは難しいため、仕様扱いとする。
   * @private
   * @param {integer} type              種別(1:Bynary/2:Text)
   * @param {string|Variant} src        書き込むデータ
   * @param {string} path               書き込むパス
   * @param {integer} [opt_option=1]    オプション(1:上書きなし/2:上書きあり)
   * @param {string} [opt_charset='utf-8']      文字コード
   * @param {boolean} [opt_bom=true]    BOM(true/false)
   */
  function FileUtility_storeFile(type, src, path, opt_option, opt_charset, opt_bom) {
    var option  = (opt_option === true)?  _this.SaveOptionsEnum.adSaveCreateOverWrite:
                  (opt_option === false)? _this.SaveOptionsEnum.adSaveCreateNotExist:
                  (opt_option != null)?   opt_option:  _this.SaveOptionsEnum.adSaveCreateNotExist;
    var charset = (opt_charset!= null)?   opt_charset: _this.UTF_8;
    var bom     = (opt_bom    != null)?   opt_bom: true;
    var skip, fullpath,
        sr, pre, bin;
    
    // 前処理
    skip = {};
    skip[_this.UTF_8] = 3;
    skip[_this.UTF_16]= 2;
    fullpath = fs.GetAbsolutePathName(path);
    
    // (存在しない場合)フォルダを作成する
    _this.createFileFolder(fullpath);
    
    // ファイルに書き込む。
    sr = new ActiveXObject('ADODB.Stream');
    sr.Type = type;
    if (type == _this.StreamTypeEnum.adTypeText) {
      sr.Charset = charset;
      sr.Open();
      sr.WriteText(src);
      if ((bom === false) && skip[charset]) {
        // BOMなし書込処理
        pre = sr;
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
        sr.Write(bin);
      }
    } else {
      // 上記以外(バイナリ)の時
      sr.Open();
      sr.Write(src);
    }
    sr.SaveToFile(fullpath, option);
    sr.Close();
    sr = null;
    // 補足:LineSeparatorプロパティは、全行読み出しのため、無意味
  }
  
  /**
   * バイナリファイルを書き込む
   * @param {Variant} bytes             書き込むデータ
   * @param {string} path               ファイルへのパス
   * @param {integer} [opt_option=1]    オプション(1:上書きなし/2:上書きあり)
   */
  _this.storeBinary = function FileUtility_storeBinary(bytes, path, opt_option) {
    FileUtility_storeFile(_this.StreamTypeEnum.adTypeBinary, bytes, path, opt_option);
  };
  
  /**
   * テキストファイルに書き込む
   * @param {string} text               書き込むデータ
   * @param {string} path               ファイルへのパス
   * @param {integer} [opt_option=1]    オプション(1:上書きなし/2:上書きあり)
   * @param {string} [opt_charset='utf-8']      文字コード
   * @param {boolean} [opt_bom=true]    BOM(true/false)
   */
  _this.storeText = function FileUtility_storeText(text, path, opt_option, opt_charset, opt_bom) {
    FileUtility_storeFile(_this.StreamTypeEnum.adTypeText, text, path, opt_option, opt_charset, opt_bom);
  };
  
  /**
   * 書き込み禁止を判定します。
   * @param {string} path               ファイルへのパス
   * @return {boolean}                  結果(true:書き込み禁止)
   */
  _this.isWriteProtected = function FileUtility_isWriteProtected(path) {
    var ret = false;
    var fullpath = fs.GetAbsolutePathName(path);
    if (fs.FileExists(fullpath)) {
      // 読み取り専用ファイルを判定
      ret = !!(fs.GetFile(fullpath).Attributes & _this.FileAttributes.ReadOnly);
    }
    return ret;
  };
  
  /**
   * ファイル/フォルダ検索
   * @private
   * @param {Function} callback         コールバック関数(該当ファイル/フォルダ時に呼び出す)
   *                                    想定関数(function (fullpath, folder) {})
   *                                    fullpath:ファイルフォルダのフルパス
   *                                    folder:(true:フォルダ/false:ファイル)
   *                                    return:true判定するものを返した場合、処理を中断する
   * @param {string} path               検索開始パス
   * @param {string} target             検索対象(targetFiles/targetAllFiles/...)
   * @param {integer} order             検索順序(orderFront:前方検索/orderPost:後方検索)
   * @param {string[]} [opt_extensions=]拡張子(例:['jpg', 'png'])
   * @return {boolean}                  callback関数の戻り値
   */
  function FileUtility_findMain(callback, fullpath, target, order, opt_extensions) {
    var callee = FileUtility_findMain,
        ret = void 0,
        folder, flag, e, path, ext, ntarget;
    
    if (fs.FolderExists(fullpath)) {
      folder = fs.GetFolder(fullpath);
      // 検索順序の残件がある時 && フォルダが存在する時(削除されることもある) && 中断でない時
      for (flag=order; flag!=0 && fs.FolderExists(fullpath) && !ret; flag=flag>>>8) {
        switch (flag & 0xff) {
        case folderorder: // フォルダ
          if (target & folderlist) {
            if (target & rootfolder) {
              ret = callback(fullpath, true);
            } else if ((target & leaffolder) && (folder.SubFolders.Count == 0)) {
              ret = callback(fullpath, true);
            }
          }
          break;
        case fileorder: // ファイル
          // ファイル一覧の時
          if (target & filelist) {
            if (!opt_extensions) {
              for (e=new Enumerator(folder.Files); !e.atEnd() && !ret; e.moveNext()) {
                ret = callback(e.item().Path, false);
              }
            } else {
              for (e=new Enumerator(folder.Files); !e.atEnd() && !ret; e.moveNext()) {
                path = e.item().Path;
                ext = fs.GetExtensionName(path).toLowerCase();
                if (opt_extensions.indexOf(ext) !== -1) {
                  ret = callback(path, false);
                }
              }
            }
          }
          break;
        case suborder:
          if (target & subfolder) {
            for (e=new Enumerator(folder.SubFolders); !e.atEnd() && !ret; e.moveNext()) {
              path = e.item().Path;
              ntarget = (target & branchfolder)? target|rootfolder: target&(~rootfolder);
                                // 枝を追加する場合、ルート(枝)追加
                                // 上記以外、ルートを削除
              ret = callee(callback, path, ntarget, order, opt_extensions);
                                // サブフォルダ(再帰呼出し)
            }
          } else if ((target & folderlist) && (target & branchfolder)) {
            for (e=new Enumerator(folder.SubFolders); !e.atEnd() && !ret; e.moveNext()) {
              // サブなしフォルダ一覧
              ret = callback(e.item().Path, true);
            }
          }
          break;
        }
      }
    } else if (fs.FileExists(fullpath)) {
      // ファイル一覧の時
      if (target & filelist) {
        ret = callback(fullpath, false);
      }
    }
    
    return ret;
  };
  _this.find = function FileUtility_find(callback, path, target, opt_order, opt_extensions) {
    var order = (opt_order)? opt_order: _this.orderFront;
    var fullpath = fs.GetAbsolutePathName(path);
    if (opt_extensions) {
      for (var i=0; i<opt_extensions.length; i++) {
        opt_extensions[i] = opt_extensions[i].toLowerCase();
      }
    }
    return FileUtility_findMain(callback, fullpath, target, order, opt_extensions);
  };
  _this.findFile = 
  function FileUtility_findFile(callback, path, opt_target, opt_order, opt_extensions) {
    var target = (opt_target != null)? opt_target: _this.targetAllFiles;
    return _this.find(callback, path, target, opt_order, opt_extensions);
  };
  _this.findFolder = 
  function FileUtility_findFolder(callback, path, opt_target, opt_order) {
    var target = (opt_target != null)? opt_target: _this.targetSubFolders;
    return _this.find(callback, path, target, opt_order);
  };
  
  
  /**
   * ファイルパスの一覧を返す
   * @param {string} folderpath         フォルダパス
   * @return {string[]}                 ファイルパスの配列
   */
  _this.getFiles = function FileUtility_getFiles(folderpath) {
    var list = [];
    var fullpath = fs.GetAbsolutePathName(folderpath);
    if (fs.FolderExists(fullpath)) {
      for (var e=new Enumerator(fs.GetFolder(fullpath).Files); !e.atEnd(); e.moveNext()) {
        list.push(e.item().Path);
      }
    }
    return list;
  };
  
  /**
   * フォルダパスの一覧を返す
   * @param {string} folderpath         フォルダパス
   * @return {string[]}                 フォルダパスの配列
   */
  _this.getFolders = function FileUtility_getFolders(folderpath) {
    var list = [];
    var fullpath = fs.GetAbsolutePathName(folderpath);
    if (fs.FolderExists(fullpath)) {
      for (var e=new Enumerator(fs.GetFolder(fullpath).SubFolders); !e.atEnd(); e.moveNext()){
        list.push(e.item().Path);
      }
    }
    return list;
  };
  
  /**
   * 存在確認
   * パスが260桁以上の場合、存在したとしてもfalseを返すことがある
   * @param {string} src                移動前のパス(例:'C:\\A\\B.txt')
   * @param {string} dst                移動後のパス(例:'C:\\B\\C.txt')
   * @return {boolean}                  複製完了有無
   */
  _this.exists =
  _this.Exists = function FileUtility_Exists(path) {
    var fullpath = fs.GetAbsolutePathName(path);
    return (fs.FileExists(fullpath) || fs.FolderExists(fullpath));  // ファイル/フォルダあり
  }
  /**
   * 複製
   * 同期？
   * @param {string} src                複製前のパス(例:'C:\\A\\B.txt')
   * @param {string} dst                複製後のパス(例:'C:\\B\\C.txt')
   * @param {boolean} [opt_overwite=false]      複製先にファイルがある場合、上書きする
   * @return {boolean}                  複製完了有無
   */
  _this.copy =
  _this.Copy = function FileUtility_Copy(src, dst, opt_overwite) {
    var overwite = (opt_overwite === true);
    var ret = false;
    src = fs.GetAbsolutePathName(src);
    dst = fs.GetAbsolutePathName(dst);
    
    try {
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
    } catch (e) {
      if (e && e.number === -2146828218) {
        // Error: 書き込みできません。
        // 原因:コピー元フォルダ内のファイル/フォルダパスが260桁以上
        //      コピー先フォルダ内のファイル/フォルダパスが260桁以上
      }
    }
    return ret;
  };
  /**
   * 移動
   * 同期？
   * 移動元パスが260桁以上の場合、存在したとしても移動しないことがある
   * @param {string} src                移動前のパス(例:'C:\\A\\B.txt')
   * @param {string} dst                移動後のパス(例:'C:\\B\\C.txt')
   * @return {boolean}                  移動完了有無
   */
  _this.move =
  _this.Move = function FileUtility_Move(src, dst) {
    var ret = false;
    src = fs.GetAbsolutePathName(src);
    dst = fs.GetAbsolutePathName(dst);
    
    try {
      if (_this.Exists(dst)) {                  // 移動先が存在する
      } else if (fs.FileExists(src)) {          // 移動元ファイルがある
        _this.createFileFolder(dst);
        fs.MoveFile(src, dst, true);
        ret = true;
      } else if (fs.FolderExists(src)) {        // 移動元フォルダがある
        _this.createFileFolder(dst);
        fs.MoveFolder(src, dst);
        ret = true;
      }
    } catch (e) {
      if (e && e.number === -2146828218) {
        // Error: 書き込みできません。
        // 原因:移動元フォルダ内のファイル/フォルダパスが260桁以上
        //      移動先フォルダ内のファイル/フォルダパスが260桁以上
      }
    }
    return ret;
    // 上書き機能はいる？
  };
  /**
   * 名称変更
   * 同期？
   * 移動元パスが260桁以上の場合、存在したとしても移動しないことがある
   * @param {string} src                移動前のパス(例:'C:\\A\\B.txt')
   * @param {string} name               変更後の名称
   * @return {boolean}                  変更完了有無
   */
  _this.rename =
  _this.Rename = function FileUtility_Rename(src, name) {
    var parent= fs.GetParentFolderName(src);
    var dist  = fs.BuildPath(parent, name);
    return _this.Move(src, dist);               // リネーム(移動)
  };
  /**
   * ファイル/フォルダ削除
   * 非同期
   * @param {string} path               パス
   * @param {boolean} [opt_force=false] 読み取り専用を削除する
   * @return {boolean}                  削除完了有無
   */
  _this.del =
  _this.Delete = function FileUtility_Delete(path, opt_force) {
    var ret = false,
        force = (opt_force === true),
        fullpath, file, folder;
    
    if (!path || path == '') {
      // 誤動作防止用
      // ''は、カレントディレクトリ扱いになる
      return ret;
    }
    try {
      fullpath = fs.GetAbsolutePathName(path);
      if (fs.FileExists(fullpath)) {
        file = fs.GetFile(fullpath);
        file.Delete(force);
        ret  = true;
      } else if (fs.FolderExists(fullpath)) {
        folder = fs.GetFolder(fullpath);
        folder.Delete(force);
        ret = true;
      }
    } catch (e) {
      if (e && e.number === -2146828218) {
        // Error: 書き込みできません。
        // 原因:forceなし書き込み禁止ファイルである
        //      forceなし書き込み禁止ファイルを含むフォルダである
      }
    }
    return ret;
    // 補足:forceについて、英語のドキュメントのみ書いてある(・_・;)
  };
  
  /**
   * JSONファイルを読み込んで返します。
   * @param {string} path               ファイルパス
   * @param {string} [opt_charset='_autodetect_all']    文字コード
   * @return {Object}                   読み込んだJSONデータ(null:読み込み失敗)
   */
  _this.loadJSON = function FileUtility_loadJSON(path, opt_charset) {
    if ('JSON' in global) {
      var ret  = null;
      var text = _this.loadText(path, opt_charset);
      if (text && text.length !== 0) {
        try {
          ret = JSON.parse(text);
        } catch (e) {}
      }
      return ret;
    } else {
      throw new Exception('The variable "JSON" is not declared.');
    }
  };
  /**
   * JSONファイルを書き込みます。
   * @param {Object} json               書き込むデータ
   * @param {string} path               ファイルへのパス
   * @param {integer} [opt_option=1]    オプション(1:上書きなし/2:上書きあり)
   * @param {string} [opt_charset='utf-8']      文字コード
   * @param {boolean} [opt_bom=true]    BOM(true/false)
   */
  _this.storeJSON = function FileUtility_storeJSON(json, path, opt_option, opt_charset, opt_bom) {
    if ('JSON' in global) {
      _this.storeText(JSON.stringify(json), path, opt_option, opt_charset, opt_bom);
    } else {
      throw new Exception('The variable "JSON" is not declared.');
    }
  };
  
  return _this;
});
