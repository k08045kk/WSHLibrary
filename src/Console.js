/*!
 * Console.js v18
 *
 * Copyright (c) 2018 toshi (https://github.com/k08045kk)
 *
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */

/**
 * WSH(JScript)用コンソール
 * @requires    module:WScript
 * @requires    ActiveXObject('Scripting.FileSystemObject')
 * @requires    ActiveXObject('WScript.Shell')
 * @requires    ErrorUtility.js
 * @auther      toshi (https://github.com/k08045kk)
 * @license     MIT License
 * @version     18
 * @since       13 - update - countReset/timeLogに対応
 * @since       13 - update - print/printlnのグローバル変数を廃止
 * @since       13 - update - 読込み時のコマンドライン引数解析を廃止
 * @since       14 - update - ErrorUtility.js v5対応（trace()の引数を削減）
 * @since       15 - update - setFormat()を追加
 * @since       15 - update - trace()でErrorUtilityあり時にErrorUtilityを使用するように変更
 * @since       16 - update - _FileUtility_createFolder()を再帰処理しないように修正
 * @since       17 - 20210109 - fix isNewLine()がsubstrのバグによって動作しないことがある
 * @since       18 - 20210829 - Console.Popup()を追加
 */
(function(root, factory) {
  if (!root.Console) {
    root.Console = factory(root.ErrorUtility);
    root.console = Console.getConsole();
  }
})(this, function(ErrorUtility) {
  "use strict";
  
  // -------------------- private --------------------
  
  var fs = void 0;
  var sh = void 0;
  var _this = void 0;
  var _newline = true;  // 改行直後
  var _consoleSet = {}; // コンソールセット
  
  /**
   * PrivateUnderscore.js
   * @version   7
   */
  {
    try {
      fs = new ActiveXObject('Scripting.FileSystemObject');
    } catch (e) {}
    try {
      sh = new ActiveXObject('WScript.Shell');
    } catch (e) {}
    // see https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Array/indexOf
    function _Array_indexOf(array, searchElement, fromIndex) {
      if (array == null) throw new TypeError('"array" is null or not defined');
      var o = Object(array);
      var len = o.length >>> 0;
      if (len === 0) return -1;
      var n = fromIndex | 0;
      if (n >= len) return -1;
      var k = n >= 0 ? n : Math.max(len + n, 0);
      for (; k < len; k++)  if (k in o && o[k] === searchElement) return k;
      return -1;
    };
    // see https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Number/isInteger
    function _Number_isInteger(value) {
      return typeof value === 'number' && isFinite(value) && Math.floor(value) === value;
    };
    // see https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Date/toISOString
    function _Date_toISOString() {
      function pad(number) {
        return ((number < 10)? '0': '') + number;
      }
      return this.getUTCFullYear() +
        '-' + pad(this.getUTCMonth() + 1) +
        '-' + pad(this.getUTCDate()) +
        'T' + pad(this.getUTCHours()) +
        ':' + pad(this.getUTCMinutes()) +
        ':' + pad(this.getUTCSeconds()) +
        '.' + (this.getUTCMilliseconds() / 1000).toFixed(3).slice(2, 5) +
        'Z';
    };
    function _isString(obj) {
      return Object.prototype.toString.call(obj) === '[object String]';
    };
    function _isError(obj) {
      return Object.prototype.toString.call(obj) === '[object Error]';
    };
    function _getCallee() {
      var args = arguments;
      var func = args.callee;
      return func.caller;
    };
    function _getFunctionName(func) {
      var name = 'anonymous';
      if (func === Function || func === Function.prototype.constructor) {
        name = 'Function';
      } else if (func !== Function.prototype) {
        var match = ('' + func).match(/^(?:\s|[^\w])*function\s+([^\(\s]*)\s*/);
        if (match != null && match[1] !== '') {
          name = match[1];
        }
      }
      return name;
    };
    function _mlength(str) {
      var len=0,
          i;
      for (i=0; i<str.length; i++) {
        len += (str.charCodeAt(i) > 255) ? 2: 1;
      }
      return len;
    };
    function _dateFormat(format, opt_date, opt_prefix, opt_suffix) {
      var pre = (opt_prefix != null)? opt_prefix: '';
      var suf = (opt_suffix != null)? opt_suffix: '';
      var fmt = {};
      fmt[pre+'yyyy'+suf] = function(date) { return ''  + date.getFullYear(); };
      fmt[pre+'MM'+suf]   = function(date) { return('0' +(date.getMonth() + 1)).slice(-2); };
      fmt[pre+'dd'+suf]   = function(date) { return('0' + date.getDate()).slice(-2); };
      fmt[pre+'hh'+suf]   = function(date) { return('0' +(date.getHours() % 12)).slice(-2); };
      fmt[pre+'HH'+suf]   = function(date) { return('0' + date.getHours()).slice(-2); };
      fmt[pre+'mm'+suf]   = function(date) { return('0' + date.getMinutes()).slice(-2); };
      fmt[pre+'ss'+suf]   = function(date) { return('0' + date.getSeconds()).slice(-2); };
      fmt[pre+'SSS'+suf]  = function(date) { return('00'+ date.getMilliseconds()).slice(-3); };
      fmt[pre+'yy'+suf]   = function(date) { return(''  + date.getFullYear()).slice(-2); };
      fmt[pre+'M'+suf]    = function(date) { return ''  +(date.getMonth() + 1); };
      fmt[pre+'d'+suf]    = function(date) { return ''  + date.getDate(); };
      fmt[pre+'h'+suf]    = function(date) { return ''  +(date.getHours() % 12); };
      fmt[pre+'H'+suf]    = function(date) { return ''  + date.getHours(); };
      fmt[pre+'m'+suf]    = function(date) { return ''  + date.getMinutes(); };
      fmt[pre+'s'+suf]    = function(date) { return ''  + date.getSeconds(); };
      fmt[pre+'S'+suf]    = function(date) { return ''  + date.getMilliseconds(); };
      
      var date = opt_date;
      if (date == null) {
        date = new Date();
      } else if (typeof date === 'number' && isFinite(date) && Math.floor(date) === date) {
        date = new Date(date);
      } else if (Object.prototype.toString.call(date) === '[object String]') {
        date = new Date(date);
      }
      
      var result = format;
      for (var key in fmt) {
        if (fmt.hasOwnProperty(key)) {
          result = result.replace(key, fmt[key](date));
        }
      }
      return result;
    };
    function _error(opt_message, opt_root, opt_error) {
      if (ErrorUtility != null) {
        return ErrorUtility.create(opt_message, opt_root, opt_error);
      } else if (Object.prototype.toString.call(opt_error) === '[object Error]') {
        return new opt_error(opt_message);
      } else {
        var e = new Error(opt_message);
        if (Object.prototype.toString.call(opt_error) === '[object String]') {
          e.name = opt_error;
        }
        return e;
      }
    };
    function _errormessage(error) {
      return ''
          + (error.name? error.name: 'UnknownError')
          + '('
            // 機能識別符号(上位16bit).エラーコード(下位16bit)
          + (error.number? ((error.number>>16)&0xFFFF)+'.'+(error.number&0xFFFF): '')
          + ')'
          + (error.message? ': '+error.message: '');
    };
    function _stack(callee, message, prefix) {
      callee = callee || _getCallee();
      var stack = [message];
      var funcs = []; // 再帰呼び出し検出用
      for (var func=callee.caller; func; func=func.caller) {
        // 再帰呼び出し検出
        if (_Array_indexOf(funcs, func) !== -1) {
          stack.push(prefix+_getFunctionName(func)+'()...');
          break;
        }
        funcs.push(func);
        stack.push(prefix+_getFunctionName(func)+'()');
      }
      return stack.join('\n');
    };
    function _toString(message, isString) {
      var callee = _toString;
      function propertyList(obj) {
        var a = [];
        for (var key in obj) {
          try {
            if (obj.hasOwnProperty(key)) {  a.push(key+':'+callee(obj[key], true));  }
          } catch (e) {}  // native method
        }
        return a;
      }
      
      var msg = ''+message;
      switch (Object.prototype.toString.call(message)) {
      case '[object Null]':
      case '[object Undefined]':
      case '[object Boolean]':
      case '[object Number]':
      case '[object RegExp]':
        break;
      case '[object String]':
        msg = (isString === true)? '"' + msg + '"': msg;
        break;
      case '[object Date]':
        msg = _Date_toISOString.call(message);
        msg = (isString === true)? '"' + msg + '"': msg;
        break;
      case '[object Math]':
        msg = 'Math';
        break;
      case '[object Error]':
        msg = (message.name == null)? 'Error': ''+message.name;
        msg += '{'+propertyList(message).join(',')+'}';
        break;
      case '[object Function]':
        msg = _getFunctionName(message)+'()';
        var temp = propertyList(message);
        if (temp.length != 0) {
          msg += '{'+temp.join(',')+'}';
        }
        break;
      case '[object Array]':
        msg = '[';
        for (var i=0; i<message.length; i++) {
          msg += (i? ',': '') + callee(message[i], true);
        }
        msg += ']';
        break;
      default:
        if (message === Function('return this')()) {
          msg = 'Global';
        } else if (typeof message === 'function' || typeof message === 'object' && !!message) {
          msg = '{'+propertyList(message).join(',')+'}';
        }
        break;
      }
      return msg;
    };
    function _Process_getNamedArgument(name, def, min, max) {
      if (WScript.Arguments.Named.Exists(name)) {
        var arg = WScript.Arguments.Named.Item(name);
        
        // 型が一致する場合、代入する
        if (def === void 0) {                   // 未定義の時
          def = (arg === void 0)? true: arg;
        } else if (typeof def == typeof arg) {  // string or boolean の時
          def = arg;
        } else if (typeof def == 'number') {
          try {
            arg = new Number(arg);
            if (isNaN(arg)) {
            } else if (min !== void 0 && arg < min) {
            } else if (max !== void 0 && arg > max) {
            } else {
              def = arg;
            }
          } catch (e) {}  // 変換失敗
        }
      }
      return def;
    };
    function _Process_getScriptPath(opt_ext) {
      var parent= fs.GetParentFolderName(WScript.ScriptFullName);
      var base  = fs.GetBaseName(WScript.ScriptFullName);
      var ext   = (!opt_ext)? fs.GetExtensionName(WScript.ScriptFullName): opt_ext;
      if (ext.length != 0 && ext.substr(0, 1) !== '.') {
        ext = '.'+ext;
      }
      return fs.BuildPath(parent, base + ext);
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
    function _FileUtility_createFileFolder(filepath) {
      return _FileUtility_createFolder(fs.GetParentFolderName(filepath));
    };
  }
  
  
  
  // -------------------- static --------------------
  
  _this = function Console_constructor() {
    this.initialize.apply(this, arguments);
  };
  
  // 出力レベル
  _this.FATAL   = 8;  // [FATAL] 致命的   重大な障害を示す(異常終了を伴うようなもの)
  _this.ERROR   = 7;  // [ERROR] エラー   予期しない他の実行時エラー
  _this.WARN    = 6;  // [WARN]  警告     潜在的な問題を示す
  _this.INFO    = 5;  // [INFO]  情報     既定
  _this.CONF    = 4;  // [CONF]  構成     静的な構成メッセージ
  _this.CONFIG  = _this.CONF;
  _this.FINE    = 3;  // [FINE]  普通     トレースメッセージ
  _this.FINER   = 2;  // [FINER] 詳細     詳細なトレースメッセージ
  _this.FINEST  = 1;  // [FINEST]最も詳細 非常に詳細なトレースメッセージ
  //_this.TRACE = 1;  // [TRACE] トレース ログ上のみ出力(廃止)
  // TRACEを廃止する。以下の処理で代用が可能
  // var trace = Console.getConsole('trace');
  // trace.propertySet.stdout = false;
  // trace.propertySet.stderr = false;
  // trace.addOutFile();
  // trace.log('Hello World.');
  
  // 出力レベル(setLevel関数用)
  _this.OFF     = 9;  // 非出力
  _this.ALL     = 0;  // すべて
  
  // ポップアップの標準タイトル
  _this.title   = WScript.ScriptName || 'Windows Script Host';
  
  
  /**
   * コンソールインスタンスの取得
   * @param {string} [opt_name='default'] - 機能名
   * @return {Console} コンソール
   */
  _this.getConsole = function Console_getConsole(opt_name) {
    var name = (opt_name)? opt_name: 'default';
    var console = _consoleSet[name];
    if (!console) {                      // 未作成ならば作成
      console = new _this(name);
      _consoleSet[name] = console;
    }
    return console;
  };
  
  /**
   * ビープ音出力
   * ビープ音の出力を待機しない(非同期実行)
   * 連続でビープ音出力しても、指定回数のビープ音を知覚できないことがありえる。
   * WScript.Sleep();等で一定時間、時間を置いてから使用すること。
   * @deprecated Windows環境に依存する
   */
  _this.beep = 
  _this.Beep = 
  _this.prototype.beep = 
  _this.prototype.Beep = function Console_beep() {
    sh.Run('rundll32 user32.dll,MessageBeep', 0, false);
  };
  
  /**
   * グラフィカルなポップアップを表示
   * 詳細な仕様は、WScript.ShellオブジェクトのPopup関数を参照
   * @param {string} message - メッセージ
   * @param {string} [opt_wait=0] - 待機時間（0:無制限/n:秒）
   * @param {string} [opt_title=null] - タイトル
   * @param {string} [opt_type=0] - ボタンとアイコンの種類
   * @return {number} 結果
   */
  _this.popup = 
  _this.Popup = 
  _this.prototype.popup = 
  _this.prototype.Popup = function Console_popup(message, opt_wait, opt_title, opt_type) {
    var wait = opt_wait || 0;
    var title = opt_title || _this.title;
    var type = opt_type || 0;
    return sh.Popup(message, wait, title, type);
  };
  
  
  
  // -------------------- public --------------------
  
  /**
   * 初期化
   * @param {string} name - 機能名
   */
  _this.prototype.initialize = function Console_initialize(name) {
    this.name = name;           // コンソール名
    this.streams = [];          // ストリーム配列
    this.groups = [];           // グループ配列
    this.counterSet = {};       // カウンターセット
    this.timerSet = {};         // タイマーセット
    this.propertySet = {};      // ログ出力有無
    this.propertySet.stdout = true;     // 標準出力
    this.propertySet.stderr = true;     // 標準エラー
    this.propertySet.group = true;      // グループ出力
    this.propertySet.format = '${indent}${message}';    // フォーマット
    // 例: '[${yyyy}/${MM}/${dd} ${HH}:${mm}:${ss}]${prefix}${group} ${indent}${message}';
    this.propertySet.newline = false;   // ログ強制改行
    this.propertySet.preindent = true;  // 改行後のインデント出力
    this.prefixSet = {};        // ログ出力用の接頭詞セット
    this.prefixSet[_this.FATAL] = '[FATAL]';
    this.prefixSet[_this.ERROR] = '[ERROR]';
    this.prefixSet[_this.WARN]  = '[WARN]';
    this.prefixSet[_this.INFO]  = '[INFO]';
    this.prefixSet[_this.CONF]  = '[CONF]';
    this.prefixSet[_this.FINE]  = '[FINE]';
    this.prefixSet[_this.FINER] = '[FINER]';
    this.prefixSet[_this.FINEST]= '[FINEST]';
    //this.prefixSet[_this.TRACE] = '[TRACE]';
    
    this.setLevel(_this.INFO);  // 出力レベル
  };
  
  /**
   * ストリーム追加
   * @param {TextStreamObject} stream - ストリーム
   */
  _this.prototype.addStream = function Console_addStream(stream) {
    this.streams.push(stream);
  };
  
  /**
   * ファイル出力を追加
   * @param {string} [opt_path='./{FileBaseName}(.{this.name}).log' - 出力ファイルパス
   * @param {boolean} [opt_append=false] - 追記モード
   * @param {boolean} [opt_format=null] - 文字コード(null:SystemDefault/true:UTF-16/false:ASCII)
   */
  _this.prototype.addOutFile = function Console_addOutFile(opt_path, opt_append, opt_format) {
    var path = opt_path;
    if (path == null) {
      // 拡張子(+ファイル無効文字削除)
      var ext = (this.name == 'default')? 'log': this.name+'.log';
      var marks = ['\\','/',':','*','?','"','<','>','|'];
      for (var i=0; i<marks.length; i++) {
        ext = ext.replace(marks[i], '');
      }
      path = _Process_getScriptPath(ext);
    }
    var append = opt_append === true;
    
    var fullpath = fs.GetAbsolutePathName(path);
    var iomode   = (append === true)? 8: 2;     // FileUtility.OpenTextFileIomode参照
//                      FileUtility.OpenTextFileIomode.ForAppending:  // 追加
//                      FileUtility.OpenTextFileIomode.ForWriting;    // 新規
    var create   = true;    // ない場合、新規作成
    var format   = -2;      // 文字コード(-2:SystemDefault/-1:UTF-16/0:ASCII))
    if (opt_format === true) {  format = -1; }
    if (opt_format === false) { format =  0; }
    
    _FileUtility_createFileFolder(fullpath);
    this.addStream(fs.OpenTextFile(fullpath, iomode, create, format));
  };
  
  /**
   * 出力レベル設定
   * @param {(number|string)} level - 出力レベル
   */
  _this.prototype.setLevel = function Console_setLevel(level) {
    if (_Number_isInteger(level)) {
      this.level = level;
    } else if (_isString(level)) { 
      level = _this[level.toUpperCase()];
      if (_Number_isInteger(level) && _this.ALL <= level && level <= _this.OFF) {
        this.level = level;
      }
    } else {
      //throw new TypeError('"level" is not the intended type. (level='+level+')');
    }
  };
  
  /**
   * 出力フォーマット設定
   * @param {string} format - フォーマット
   */
  _this.prototype.setFormat = function Console_setFormat(format) {
    this.propertySet.format = format;
  };
  
  /**
   * 出力判定
   * @param {number} [opt_level=Console.INFO] - 出力レベル
   */
  _this.prototype.isOutput = function Console_isOutput(opt_level) {
    var level = (opt_level == null)? _this.INFO: opt_level;
    return (this.level <= level);
  };
  
  /**
   * 行頭判定
   * @deprecated Console.Animation.js用
   * @return {boolean} 行頭？
   */
  _this.prototype.isNewLine = function Console_isNewLine() {
    return _newline;
  };
  
  /**
   * ストリーム書き出し
   * 制約: print(), println()以外から、_printCore()を呼び出さないこと。
   * @protected
   * @param {!string} message - 出力文字列
   * @param {number} [opt_level=Console.INFO] - 出力レベル
   */
  _this.prototype._printCore = function Console__printCore(message, opt_level) {
    var level = (opt_level == null)? _this.INFO: opt_level;
    
    if (this.isOutput(level)) {
      if (this.propertySet.stderr && _this.ERROR <= level) {
        // 標準エラー
        try {
          WScript.StdErr.Write(message);
        } catch(e) {}
      } else if (this.propertySet.stdout) {
        // 標準出力
        try {
          WScript.StdOut.Write(message);
        } catch(e) {}
      }
      // ファイル出力
      for (var i=0; i<this.streams.length; i++) {
        try {
          this.streams[i].Write(message);
        } catch(e) {}
      }
    }
  };
  
  /**
   * 文字列化する
   * オブジェクト等を文字列化する。
   * @protected
   * @param {*} [message=''] - メッセージ
   * @return {string} 出力文字列
   */
  _this.prototype._toString = function Console__toString(message) {
    return (message === void 0)? '': _toString(message);
  };
  
  /**
   * 出力
   * @param {string} [message=''] - 出力文字列
   * @param {number} [opt_level=Console.INFO] - 出力レベル
   */
  _this.prototype.print = function Console_print(message, opt_level) {
    var level = (opt_level == null)? _this.INFO: opt_level;
    var msg = this._toString(message);
    
    this._printCore(msg, opt_level);
    //_newline = msg.endsWith('\n');
    _newline = (msg.substring(msg.length-1, msg.length) === '\n');
  };
  _this.prototype.println = function Console_println(message, opt_level) {
    var msg = this._toString(message);
    this.print(msg+'\n', opt_level);
  };
  
  /**
   * ログ出力
   * 残余引数 or arguments対応の可能性があります。
   * そのため、出力レベルの使用は、非推奨とします。
   * INFO以外の出力レベルは、個別の関数を使用してください。
   * @param {string} message - 出力文字列
   * @param {number} [opt_level=Console.INFO] - 出力レベル（非推奨）
   */
  _this.prototype.log = function Console_log(message, opt_level) {
    var level = (opt_level == null)? _this.INFO: opt_level;
    var msg = this._toString(message);
    var newline = (this.propertySet.newline && !this.isNewLine())? '\n': '';
    var prefix  = (this.prefixSet[level])? this.prefixSet[level]: '';
    var indent  = Array(this.groups.length+1).join('  ');
    var group   = (this.groups.length!=0)? this.groups[this.groups.length-1]: '';
    
    // メッセージ以外を置換
    var format = _dateFormat(this.propertySet.format, new Date(), '${', '}')
                  .replace('${prefix}', prefix)
                  .replace('${group}', group);
    
    // 行頭以外の改行に続く新規行をインデント
    if (this.propertySet.preindent) {
      var idx = format.indexOf('${indent}');
      var preindent = indent;
      if (idx == -1) {
        idx = format.indexOf('${message}');
        preindent = '';
      }
      if (idx > 0 || preindent != '') {
        var index = (idx > 0)? _mlength(format.substr(0, idx)): 0;
        preindent += (index > 0)? Array(index+1).join(' '): '';
        msg = msg.replace(/(?!\n$)\n/g, '\n'+preindent);
      }
    }
    
    // 残りを置換
    format = format.replace('${indent}', indent)
                   .replace('${message}', msg);
    
    // 出力
    this.println(newline+format, level);
    // this.propertySet.format
    // [${yyyy}/${MM}/${dd} ${HH}:${mm}:${ss}]${prefix}${group} ${indent}${message}
    // ${indent                                                         }${message}
  };
  _this.prototype.fatal = function Console_fatal(message) {
    this.log(message, _this.FATAL);
  };
  _this.prototype.error = function Console_error(message) {
    this.log(message, _this.ERROR);
  };
  _this.prototype.warn = function Console_warn(message) {
    this.log(message, _this.WARN);
  };
  _this.prototype.info = function Console_info(message) {
    this.log(message, _this.INFO);
  };
  _this.prototype.conf = 
  _this.prototype.config = function Console_config(message) {
    this.log(message, _this.CONF);
  };
  _this.prototype.fine = function Console_fine(message) {
    this.log(message, _this.FINE);
  };
  _this.prototype.finer = function Console_finer(message) {
    this.log(message, _this.FINER);
  };
  _this.prototype.finest = function Console_finest(message) {
    this.log(message, _this.FINEST);
  };
  
  /**
   * 画面クリア
   * 実現できないため、特に何もしない。
   */
  _this.prototype.clear = function Console_clear() {
    ;
  };
  
  /**
   * グループ
   * 以降の出力をインデントする。
   * @param {string} [opt_label=''] - ラベル
   */
  _this.prototype.group = function Console_group(opt_label) {
    var label = (opt_label == void 0)? '': opt_label;
    if (this.propertySet.group && label.length != 0) {  this.config(label);  }
    this.groups.push(label);
    // label
    //   
  };
  _this.prototype.groupEnd = function Console_groupEnd() {
    this.groups.pop();
  };
  
  /**
   * カウント
   * 特定のcount()を呼び出した回数を記録する。
   * @param {string} opt_label - ラベル
   * @param {number} [opt_level=Console.INFO] - 出力レベル
   */
  _this.prototype.count = function Console_count(opt_label, opt_level) {
    var label = opt_label? opt_label: 'default';
    var msg = label+': ';
    if (!this.counterSet[label]) {  this.counterSet[label] = 0; }
    this.counterSet[label]++;
    this.log(msg+this.counterSet[label], opt_level);
    // label: ${count}
    // ${count}
  };
  _this.prototype.countReset = function Console_countReset(opt_label, opt_level) {
    var label = opt_label? opt_label: 'default';
    var msg = label+': ';
    this.counterSet[label] = 0;
    this.log(msg+this.counterSet[label], opt_level);
  };
  
  /**
   * 時間
   * 操作の所要時間を追跡するために使用できるタイマーを開始する。
   * @param {string} opt_label - ラベル
   * @param {number} [opt_level=Console.INFO] - 出力レベル
   */
  _this.prototype.time = function Console_time(opt_label) {
    var label = opt_label? opt_label: 'default';
    this.timerSet[label] = new Date().getTime();
  };
  _this.prototype.timeLog = function Console_timeLog(opt_label, opt_level) {
    var label = opt_label? opt_label: 'default';
    var msg = label+': ';
    if (this.timerSet[label]) {
      var elapsed = new Date().getTime() - this.timerSet[label];
      this.log(msg+elapsed+'ms', opt_level);
    } else {
      this.log('Timer "'+label+'" does not exist.', opt_level);
    }
    // label: ${time}ms
    // ${time}ms
  };
  _this.prototype.timeEnd = function Console_timeEnd(opt_label, opt_level) {
    var label = opt_label? opt_label: 'default';
    this.timeLog(label, opt_level);
    delete this.timerSet[label];
  };
  
  /**
   * アサート
   * アサーションがfalseの場合、エラーメッセージを出力する。
   * @param {boolean} assertion - アサーション
   * @param {string} message - エラーメッセージ
   */
  _this.prototype.assert = function Console_assert(assertion, message) {
    if (assertion === false) {
      var msg = '';
      var separator = '';
      if (ErrorUtility != null) {
        msg += '[';
        msg += ErrorUtility.trace(2, true);
        separator = '] '
      }
      msg += separator;
      msg += 'Assertion failed';
      msg += (message)? ': '+message: '.';
      this.warn(msg);
      // [関数名(引数, ...) - ファイル名:行数:列数] Assertion failed.
      // [関数名(引数, ...) - ファイル名:行数:列数] Assertion failed: ${msg}
    }
  };
  
  /**
   * スタックトレースを出力
   * @param {number} [opt_level=Console.INFO] - 出力レベル
   */
  _this.prototype.trace = function Console_trace(opt_level) {
    if (ErrorUtility != null) {
      var e = ErrorUtility.captureStackTrace(new Error('console.trace()'), _getCallee());
      this.log(ErrorUtility.stack(e, false), opt_level);
    } else {
      this.log(_stack(_getCallee(), 'console.trace()', '    at '), opt_level);
    }
  };
  
  /**
   * エラー出力
   * @param {Error} e - エラー
   */
  _this.prototype.printStackTrace = 
  _this.prototype.errorStackTrace = function Console_errorStackTrace(e) {
    if (!_isError(e)) {
      e = _error('"e" is not a type of Error.', null, 'TypeError');
    }
    if (ErrorUtility != null) {
      if (!e.stackframes) {
        ErrorUtility.captureStackTrace(e, _getCallee());
      }
      this.error(ErrorUtility.stack(e));
    } else {
      this.error(_stack(_getCallee(), _errormessage(e), '    at '));
    }
  };
  
  /**
   * コマンドライン引数の解析
   * 標準のコマンドライン引数解析を提供します。
   * 標準出力の有無やファイル保存機能を提供します。
   * 補足：v12以前は、読込み時に実行していた処理を提供します。
   *       ライブラリの自由度向上のため、読込み時の自動実行機能を廃止しました。
   */
  _this.commandLineArgumentAnalysis = function Console_commandLineArgumentAnalysis() {
    var console = _this.getConsole();
    
    // 標準出力のロガーを設定
    console.propertySet.stdout = _Process_getNamedArgument('stdout', true);
    console.propertySet.stderr = _Process_getNamedArgument('stderr', true);
    console.setLevel(_Process_getNamedArgument('stdout', 'INFO'));
    
    // ファイル出力設定
    console.propertySet.format = _Process_getNamedArgument('logformat', console.propertySet.format);
    var logfile = _Process_getNamedArgument('logfile', false);
    if (fs && logfile !== false) {
      var logpath = _isString(logfile)? logfile: _Process_getNamedArgument('logfilepath');
      var append  = _Process_getNamedArgument('logappend');
      var format  = _Process_getNamedArgument('logencode');
      if (_isString(logpath)) {
        // logpath    新規書き込み
        console.addOutFile(logpath, (append === true), format);
      } else if (_Process_getNamedArgument('logrotate', false)) {
        // ./log/{FileBaseName}.{yyyMMdd}.log    追加書き込み
        var fullpath= _Process_getScriptPath();
        var base    = fs.GetBaseName(fullpath);
        var date    = _dateFormat('yyyyMMdd', new Date());
        var folder  = fs.BuildPath(fs.GetParentFolderName(fullpath), 'log');
        var name    = base+'.'+date+'.log';
        console.addOutFile(fs.BuildPath(folder, name), (append !== false), format);
      } else {
        // ./{FileBaseName}.log    新規書き込み
        console.addOutFile(_Process_getScriptPath('log'), (append === true), format);
      }
    }
    // コマンド指定例
    // cscript appname.wsf /stdout- /logfile+
    // 標準出力なし、ログファイル出力あり（./{FileBaseName}.log）、新規書き込み
    // cscript appname.wsf /stdout- /logfile:"C:\test\test.log"
    // 標準出力なし、ログファイル出力あり（C:\test\test.log）、新規書き込み
    // cscript appname.wsf /stdout- /logfile+ /logrotate+ /logappend+ /logencode+
    // 標準出力なし、ログファイル出力あり（./log/{FileBaseName}.{yyyMMdd}.log）、追加書き込み、UTF-16
    // cscript appname.wsf /logformat:"[${yyyy}/${MM}/${dd} ${HH}:${mm}:${ss}]${prefix}${group} ${indent}${message}"
    // 標準出力あり、ログ出力なし、フォーマット変更あり
  };
  
  return _this;
});
