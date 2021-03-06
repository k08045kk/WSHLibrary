/*!
 * ShellUtility.js v1
 *
 * Copyright (c) 2019 toshi (https://github.com/k08045kk)
 *
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */

/**
 * WSH(JScript)用シェル関連
 * 送る、ゴミ箱へ送る、zip圧縮、zip解凍の機能を提供する。
 * @requires    module:ActiveXObject('Scripting.FileSystemObject')
 * @requires    module:ActiveXObject('Shell.Application')
 * @requires    module:WScript
 * @auther      toshi (https://github.com/k08045kk)
 * @version     1
 * @see         1 - add - 初版
 */
(function(root, factory) {
  if (!root.ShellUtility) {
    root.ShellUtility = factory();
  }
})(this, function() {
  "use strict";
  
  // -------------------- private --------------------
  
  var fs = new ActiveXObject('Scripting.FileSystemObject');
  var shell = new ActiveXObject("Shell.Application");
  var _this = void 0;
  
  /**
   * PrivateUnderscore.js
   * @version   5
   */
  {
    function _isString(obj) {
      return Object.prototype.toString.call(obj) === '[object String]';
    };
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
  }
  
  // -------------------- static --------------------
  
  /**
   * コンストラクタ
   * @constructor
   */
  _this = function ShellUtility_constructor() {};
  
  // Shell.Application関連定義
  _this.ShellSpecialFolderConstants = {
    ssfDESKTOP:   0,    // デスクトップ(仮想)
    ssfIE:        1,    // Internet Explorer
    ssfPROGRAMS:  2,    // プログラム
    ssfCONTROLS:  3,    // コントロールパネル
    ssfPRINTERS:  4,    // プリンタ
    ssfPERSONAL:  5,    // マイドキュメント
    ssfFAVORITES: 6,    // お気に入り
    ssfSTARTUP:   7,    // スタートアップ
    ssfRECENT:    8,    // 最近使ったファイル
    ssfSENDTO:    9,    // 送る
    ssfBITBUCKET: 10,   // ごみ箱
    ssfSTARTMENU: 11,   // スタートメニュー
    ssfDESKTOPDIRECTORY: 16,  // デスクトップ(フォルダ)
    ssfDRIVES:    17,   // マイコンピュータ
    ssfNETWORK:   18,   // ネットワークコンピュータ
    ssfNETHOOD:   19,   // NetHood
    ssfFONTS:     20,   // フォント
    ssfTEMPLATES: 21    // テンプレート
  };
  
  /**
   * 送る
   * @param {number} sendto - 送り先(ShellSpecialFolderConstants)
   * @param {string} path - ファイル/フォルダパス
   * @return {boolean} 成否
   */
  _this.sendTo = function ShellUtility_sendTo(sendto, path) {
    var ret = false;
    var fullpath = fs.GetAbsolutePathName(path);
    var isFile   = fs.FileExists(fullpath);
    if (isFile ||  fs.FolderExists(fullpath)) {
      shell.NameSpace(sendto).MoveHere(fullpath, 4);
      
      // 完了待機(ファイル存在確認)
      if (isFile) { while (fs.FileExists(fullpath)) {   WScript.Sleep(100); } }
      else {        while (fs.FolderExists(fullpath)) { WScript.Sleep(100); } }
      
      shell = null;
      ret = true;
    }
    return ret;
  }
  
  /**
   * ゴミ箱へ送る
   * ファイル・フォルダをゴミ箱へ送る（削除する）
   * @param {string} path - ファイル/フォルダパス
   * @return {boolean} 成否
   */
  _this.sendToTrash = function ShellUtility_sendToTrash(path) {
    return _this.sendTo(_this.ShellSpecialFolderConstants.ssfBITBUCKET, path);
  };
  
  /**
   * Explorerを取得する
   * @return {Explorer[]} Explorerの配列
   */
  _this.explorers = function ShellUtility_explorers() {
    var explorers = [];
    var windows = shell.Windows();
    for (var i=0; i<windows.Count; i++) {
      var item = windows.Item(i);
      if (item) {               // なぜか、nullを取得することがある
        explorers.push(item);
      }
    }
    return explorers;
  }
  
  /**
   * 指定URLのExplorerを返す
   * 「file:///C:/test」などでExplorerを取得できる
   * 「https://www.bugbugnow.net/」などでInternetExplorerを取得できる
   * @return {Explorer} Explorer
   */
  _this.explorer = function ShellUtility_explorer(url) {
    var explorer = null;
    var windows = shell.Windows();
    for (var i=0; i<windows.Count; i++) {
      var item = windows.Item(i);
      if (item && item.LocationURL == url) {
        explorer = item;
        break;
      }
    }
    return explorer;
  }
  
  /**
   * zip圧縮
   * Windows標準機能のみでzip圧縮する。(Microsoftサポート対象外)
   * 圧縮が終わるまで処理を抜けない。
   * @param {string} zipfile - 作成するzipファイルパス
   * @param {string[]} files - 格納ファイルパス配列(フォルダも可)
   * @return {boolean}          成否
   */
  _this.zip = function ShellUtility_zip(zipfile, files) {
    var names = [];
    for (var i=0; i<files.length; i++) {
      // ファイルの存在確認
      files[i] = fs.GetAbsolutePathName(files[i]);
      if (!fs.FileExists(files[i]) && !fs.FolderExists(files[i])) {
        return false;                           // ファイルがない
      }
      // ファイル名重複確認
      var name = fs.GetFileName(files[i]);
      for (var n=0; n<i; n++) {
        if (names[n] == name) {
          return false;                         // 同名ファイルがある
        }
      }
      names.push(name);
    }
    
    // zipファイル作成
    var zipfullpath = fs.GetAbsolutePathName(zipfile);
    if (!fs.FileExists(zipfullpath) && !fs.FolderExists(zipfullpath)) {
      var ts = null;
      try {
        ts = fs.OpenTextFile(zipfullpath, 2, true);
        ts.Write(String.fromCharCode(80,75,5,6,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0));
      } catch (e) {
        throw e;
      } finally {
        try {
          if (ts != null) {  ts.Close();  }     // ファイル閉じる
        } catch (e) {
          throw e;
        } finally {
          ts = null;
        }
      }
      // zipファイル完成待ち
      while (true) {
        try {
          WScript.Sleep(100);
          fs.OpenTextFile(zipfullpath, 8, false).Close();
          WScript.Sleep(100);
          break;
        } catch (e) {}
      }
    } else {
      return false;
    }
    
    // ファイル格納
    var zipfolder = shell.NameSpace(zipfullpath);
    for (var i=0; i<files.length; i++) {
      zipfolder.CopyHere(files[i], 4);          // zipへコピー(進歩バー非表示)
      
      // 書き込み待機(Sleepがないと動作しない、正常なzipが作成されない)
      while (true) {
        try {
          WScript.Sleep(100);
          fs.OpenTextFile(zipfullpath, 8, false).Close();
          WScript.Sleep(100);
          break;
        } catch (e) {}
      }
    }
    
    return true;
  };
  
  /**
   * zip解凍
   * Windows標準機能のみでzip解凍する。(Microsoftサポート対象外)
   * 解凍が終わるまで処理を抜けない。
   * @param {string} zipfile - 作成するzipファイルパス
   * @param {string|boolean} [opt_output=true] - 出力パス(true:zipフォルダ作成/false:直下に出力)
   * @return {boolean} 成否
   */
  _this.unzip = function ShellUtility_unzip(zipfile, opt_output) {
    // zip確認
    var zipfullpath = fs.GetAbsolutePathName(zipfile);
    if (!fs.FileExists(zipfullpath) || fs.GetExtensionName(zipfullpath).toLowerCase() != 'zip') {
      return false;
    }
    
    var output = opt_output;
    if (!(output === true || output === false || _isString(output))) {
      output = true;
    }
    var zipfolder = shell.NameSpace(zipfullpath);
    
    // 出力確認
    var outfullpath = null;
    if (output === true) {                      // フォルダ作成
      var name = fs.GetBaseName(zipfullpath);
      var path = fs.BuildPath(fs.GetParentFolderName(zipfullpath), name);
      if (fs.FileExists(path) || fs.FolderExists(path)) {
        return false;
      }
      
      var newfolder = true;
      var items = zipfolder.Items();
      if (items.Count === 1) {                  // 中身が1つのみの時
        var item = items.item();
        if (item.IsFolder) {                    // フォルダの時
          if (name == item.Name) {              // 同名の時
            newfolder = false;                  // フォルダ未作成
          }
        }
      }
      
      outfullpath = path;
      if (newfolder) {
        _FileUtility_createFolder(path);
      }
    } else {                                    // フォルダ未作成
      if (output === false) {
        outfullpath = fs.GetParentFolderName(zipfullpath);
      } else if (_isString(output)) {
        outfullpath = output;
      }
      for(var e=new Enumerator(zipfolder.Items());!e.atEnd();e.moveNext()) {
        var item = e.item();
        var path = fs.BuildPath(outfullpath, item.Name);
        if (fs.FileExists(path) || fs.FolderExists(path)) {     // 重複ファイル確認
          return false;
        }
      }
    }
    
    // ファイル解凍
    var folder = shell.NameSpace(outfullpath);
    for(var e=new Enumerator(zipfolder.Items());!e.atEnd();e.moveNext()) {
      folder.CopyHere(e.item(), 4);             // zipからコピー(進歩バー非表示)
    }
    
    return true;
  };
  
  return _this;
});
