/**
 * 簡易暗号
 * ボリューム(HDD等)のGUIDを使用した簡易暗号、
 * 保存ボリュームを変更しなければ、正常に復号可能となる。
 * 他人のPC上では失敗する。(別ドライブへ移動した場合も失敗する)
 * 
 * ソースor設定ファイルにパスワード等を
 * 平文のまま記入しない目的で使用することを想定して作成した。
 * @requires    module:ActiveXObject('Scripting.FileSystemObject')
 * @requires    module:WScript
 * @requires    module:EncodeUtility.js
 * @requires    module:WMIUtility.js
 * @auther      toshi(https://www.bugbugnow.net/)
 * @license     MIT License
 * @version     1
 */
(function(global, factory) {
  if (!global.EasyCrypto) {
    global.fs = global.fs || new ActiveXObject('Scripting.FileSystemObject');
    global.sh = global.sh || new ActiveXObject('WScript.Shell');
    global.EasyCrypto = factory(global, 
                                global.fs, 
                                global.sh, 
                                global.EncodeUtility, 
                                global.WMIUtility);
  }
})(this, function(global, fs, sh, EncodeUtility, WMIUtility) {
  "use strict";
  
  var _ModuleDirectory;
  /**
   * PrivateUnderscore.js
   * @version   1
   */
  {
    _ModuleDirectory = (global.ModuleDirectory)?
          global.ModuleDirectory:
          (WScript && WScript.ScriptFullName)?
            fs.getParentFolderName((function() {
              // MakeExe/ModulePath.js
              var ret,
                  T, fHandle;
              ret = WScript.ScriptFullName;
              T   = ret.toLowerCase();
              T   = T.substring(0,T.length-3);
              if (T.charAt(T.length-1) == '.') {
                T = T.substring(0,T.length-1);
              }
              if (T.substring(T.length-4,T.length) == '.tmp') {
                try {
                  fHandle = fs.OpenTextFile(T);
                  T = fHandle.ReadLine();
                  fHandle.Close();
                  ret = T;
                } catch (e) {}
              }
              return ret;
            })()):
            sh.CurrentDirectory;
  }
  
  var _this = void 0;
  var _guid = null;
  var _prefix = 'EasyCrypto:';
  
  _this = function EasyCrypto_constrcutor() {};
  
  
  /**
   * 暗号化有無
   * @param {string} text - 暗号化文字列
   * @return {boolean} 暗号化有無
   */
  _this.isEncrypted = function EasyCrypto_isEncrypted(text) {
    return text.startsWith(_prefix);
  };
  
  
  /**
   * 鍵取得
   * 実行ファイル格納ボリュームのGUID文字列を返す。
   * @return {string} 暗号化鍵
   */
  _this.getKey = function EasyCrypto_getKey() {
    if (_guid == null) {
      var drive = fs.GetDriveName(_ModuleDirectory);
      var drives= WMIUtility.getQueryProperties(
                                  "SELECT DeviceID FROM Win32_Volume"
                                + " WHERE DriveLetter = '"+drive+"'");
      for (var i=0; i<drives.length; i++) {
        _guid = drives[i].DeviceID.match(/{(.+)}/)[1].split('-').join('');
        break;
      }
      drives = null;
    }
    return _guid;
  };
  
  /**
   * 暗号化
   * @param {string} text - 平文
   * @return {string} 暗号化文字列
   */
  _this.encrypt = function EasyCrypto_encrypt(text) {
    var uid = EncodeUtility.hex2bin(_this.getKey());
    var src = EncodeUtility.str2bin(text);
    var bin = EncodeUtility.encrypt(uid, src);
    return _prefix + EncodeUtility.bin2base64(bin);
  };
  
  /**
   * 復号
   * encrypt()で暗号化した文字列を復号する。
   * 平文を引数に渡した場合、平文のまま返す。
   * 復号途中でエラーが発生した場合、平文のまま返す。
   * 平文以外が戻った場合でも、復号成功を保証するものではない。
   * @param {string} text - 暗号化文字列
   * @return {string} 復号文字列
   */
  _this.decrypt = function EasyCrypto_decrypt(text) {
    var ret = text;
    if (_this.isEncrypted(text)) {
      try {
        var uid = EncodeUtility.hex2bin(_this.getKey());
        var src = EncodeUtility.base642bin(text.substr(_prefix.length));
        var bin = EncodeUtility.decrypt(uid, src);
        ret = EncodeUtility.bin2str(bin);
      } catch (e) {  // base64復号エラー or ブロック不正 or ...
        console.printStackTrace(e, console.FINER);
      }
    }
    return ret;
  };
  
  return _this;
});
