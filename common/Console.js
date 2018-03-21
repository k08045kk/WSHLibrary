// https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/String/endsWith
String.prototype.endsWith = String.prototype.endsWith || function(searchString, position) {
  var subjectString = this.toString();
  if (typeof position !== 'number' || !isFinite(position) || Math.floor(position) !== position || position > subjectString.length) {
    position = subjectString.length;
  }
  position -= searchString.length;
  var lastIndex = subjectString.lastIndexOf(searchString, position);
  return lastIndex !== -1 && lastIndex === position;
};
// https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Function/bind
Function.prototype.bind = Function.prototype.bind || function (oThis) {
  if (typeof this !== 'function') {
    throw new TypeError('Function.prototype.bind - what is trying to be bound is not callable');
  }
  var aArgs = Array.prototype.slice.call(arguments, 1), 
    fToBind = this, 
    fNOP = function () {},
    fBound = function () {
      return fToBind.apply(this instanceof fNOP && oThis
        ? this
        : oThis,
        aArgs.concat(Array.prototype.slice.call(arguments)));
    };
  fNOP.prototype = this.prototype;
  fBound.prototype = new fNOP();
  return fBound;
};

/**
 * WSH(JScript)用コンソール
 * @requires    module:ActiveXObject('Scripting.FileSystemObject')
 * @requires    module:ActiveXObject('WScript.Shell')
 * @requires    module:ActiveXObject('htmlfile')
 * @requires    module:ErrorUtility.js
 * @requires    module:WScript
 * @auther      toshi(http://www.bugbugnow.net/)
 * @license     MIT License
 * @version     1
 */
(function(global, factory) {
  if (!global.Console) {
    global.fs = global.fs || new ActiveXObject('Scripting.FileSystemObject');
    global.sh = global.sh || new ActiveXObject('WScript.Shell');
    global.Console = factory(global, global.fs, global.sh, global.ErrorUtility);
    
    global.console = Console.getConsole();
    print   = console.print.bind(console);
    println = console.println.bind(console);
  }
})(this, function(global, fs, sh, ErrorUtility) {
  "use strict";
  
  /**
   * PrivateUnderscore.js
   * @version   1
   */
  {
    if (typeof(document ) === 'undefined') {
      document = new ActiveXObject('htmlfile');
      document.write('<html><head></head><body></body></html>');
    }
    if (typeof(window   ) === 'undefined') {
      window = document.parentWindow;
      setTimeout  = function(callback, millisec){
        return window.setTimeout((function(params){
          return function(){callback.apply(null, params);};
        })([].slice.call(arguments,2)), millisec);
      };
      setInterval = function(callback, millisec){
        return window.setInterval((function(params){
          return function(){callback.apply(null, params);};
        })([].slice.call(arguments,2)), millisec);
      };
      clearTimeout  = window.clearTimeout;
      clearInterval = window.clearInterval;
    }
    // https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Number/isInteger
    function _isInteger(value) {
      return typeof value === 'number' && isFinite(value) && Math.floor(value) === value;
    };
    function _isString(obj) {
      return Object.prototype.toString.call(obj) === '[object String]';
    };
    function _isError(obj) {
      return Object.prototype.toString.call(obj) === '[object Error]';
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
      } else if (date === 'number' && isFinite(date) && Math.floor(date) === date) {
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
    }
    function _errormessage(e) {
      var msg = '';
      msg += (e.name)? e.name: 'UnknownError';
      msg += '(';
      // 機能識別符号.エラーコード
      msg += (e.number)? ((e.number>>16)&0xFFFF)+'.'+(e.number&0xFFFF): '';
      msg += ')';
      msg += (e.message)? ': '+e.message: '';
      return msg;
    }
    function _WScript_getNamedArgument(name, def, min, max) {
      if (WScript.Arguments.Named.Exists(name)) {
        var arg = WScript.Arguments.Named.Item(name);
        
        // 型が一致する場合、代入する
        if (def === void 0) {                   // 未定義の時
          def = arg;
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
    function _getScriptPath(opt_ext) {
      var ext = (opt_ext && opt_ext.length !== 0)? '.'+opt_ext: '';
      var parent = fs.GetParentFolderName(WScript.ScriptFullName);
      var base = fs.GetBaseName(WScript.ScriptFullName);
      return fs.BuildPath(parent, base + ext);
    };
    function _FileUtility_createFolder(folderpath) {
      var callee = _FileUtility_createFolder,
          ret = false,
          fullpath = fs.GetAbsolutePathName(folderpath);
      if (!(fs.FolderExists(fullpath) || fs.FileExists(fullpath))) {
        var parentpath = fs.GetParentFolderName(fullpath);
        if (parentpath != '') {
          callee(fs.GetParentFolderName(fullpath));
        }
        fs.CreateFolder(fullpath);
        ret = true;
      }
      return ret;
    };
    function _FileUtility_createFileFolder(filepath) {
      var ret = false;
      var fullpath  = fs.GetAbsolutePathName(filepath);
      var parentpath= fs.GetParentFolderName(fullpath);
      if (parentpath != '') {
        ret = _FileUtility_createFolder(parentpath);
      }
      return ret;
    };
  }
  
  var _this = void 0;
  var _newline = true;          // 改行直後
  var _tempMessage = '';        // 一時表示中の文字列
  var _consoleSet = {};         // コンソールのインスタンスセット
  var _animeSet = {};           // アニメーションのインスタンスセット
  
  _this = function Console_constructor() {
    this.initialize.apply(this, arguments);
  };
  
  // 出力レベル
  _this.OFF   = 10; // 非出力
  _this.FATAL = 9;  // [FATAL]致命的    重大な障害を示す(異常終了を伴うようなもの)
  _this.ERROR = 8;  // [ERROR]エラー    予期しない他の実行時エラー
  _this.WARN  = 7;  // [WARN] 警告      潜在的な問題を示す
  _this.INFO  = 6;  // [INFO] 情報(既)
  _this.CONFIG= 5;  // [CONF] 構成      静的な構成メッセージ
  _this.FINE  = 4;  // [FINE] 普通      トレースメッセージ
  _this.FINER = 3;  // [FINER]詳細      詳細なトレースメッセージ
  _this.FINEST= 2;  // [FINE]最も詳細   非常に詳細なトレースメッセージ
  _this.TRACE = 1;  // [TRACE]トレース  ログ上のみ出力
  _this.ALL   = 0;  // すべて
  _this.TEMP  = -1; // 一時             コンソール上のみ出力し、次の出力時に削除する
  
  /**
   * コンソールインスタンスの取得
   * @param {string} [opt_name='default'] - 機能名
   * @returns {Console} コンソール
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
   * 連続でビープ音出力しても、指定回数のビープ音を知覚できないことがありえる
   * @deprecated Windows環境に依存する
   */
  _this.beep = 
  _this.Beep = 
  _this.prototype.beep = 
  _this.prototype.Beep = function Console_beep() {
    sh.Run('rundll32 user32.dll,MessageBeep', 0, false);
  }
  
  /**
   * 初期化
   * @param {string} name - 機能名
   */
  _this.prototype.initialize = function Console_initialize(name) {
    this.name = name;           // コンソール名
    this.stdout = true;         // 標準出力有無
    this.streams = [WScript.StdOut];    // ストリーム配列
    this.groups = [];           // グループ配列
    this.counterSet = {};       // カウンターセット
    this.timerSet = {};         // タイマーセット
    //this.logformat = '[${yyyy}/${MM}/${dd} ${HH}:${mm}:${ss}]${prefix} ${message}';
    this.logprefix = false;     // ログ出力用の接頭詞有無
    this.prefixSet = {};        // ログ出力用の接頭詞セット
    this.prefixSet[_this.FATAL] = '[致命的]';
    this.prefixSet[_this.ERROR] = '[エラー]';
    this.prefixSet[_this.WARN]  = '[警告]';
    this.prefixSet[_this.INFO]  = '[情報]';
    this.prefixSet[_this.CONFIG]= '[構成]';
    this.prefixSet[_this.FINE]  = '[普通]';
    this.prefixSet[_this.FINER] = '[詳細]';
    this.prefixSet[_this.FINEST]= '[最も詳細]';
    this.prefixSet[_this.TRACE] = '[トレース]';
    
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
   * @param {string} [opt_path='./{FileBaseName}.log' - 出力ファイルパス
   * @param {boolean} [opt_append=false] - 追記モード
   * @param {boolean} [opt_format=null] - 文字コード(null:SystemDefault/true:UTF-16/false:ASCII)
   */
  _this.prototype.addOutFile = function Console_addOutFile(opt_path, opt_append, opt_format) {
    var path = (opt_path != null)? opt_path: _getScriptPath('log');
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
   * @param {number} level - 出力レベル
   */
  _this.prototype.setLevel = function Console_setLevel(level) {
    if (_isInteger(level)) {
      this.level = level;
    } else if (_isInteger(_this[level]) &&_this.ALL <= _this[level] && _this[level] <= _this.OFF) {
      this.level = _this[level];
    } else {
      //throw new TypeError('"level" is not the intended type. (level='+level+')');
    }
  };
  
  /**
   * 出力判定
   * @param {number} [opt_level=INFO(6)] - 出力レベル
   */
  _this.prototype.isOutput = function Console_isOutput(opt_level) {
    var level = (opt_level != null)? opt_level: _this.INFO;
    return (this.level <= level);
  };
  
  /**
   * ストリーム書き出し(一時表示対応)
   * ただし、一時表示中にWScript.Echo(), WScript.Stdout.Write()等で別途出力すると、表示が崩れる。
   * 制約: print(), println()以外から、printCore()を呼び出さないこと。
   * @param {!string} message - 出力文字列
   * @param {number} [opt_level=INFO(6)] - 出力レベル
   */
  _this.prototype.printCore = function Console_printCore(message, opt_level) {
    var level = opt_level;
    var msg   = message;
    var tlen0 = _mlength(_tempMessage);
    if (level == null) { level = _this.INFO; }
    
    // グループ(インデント)
    var indent = Array(this.groups.length+1).join('  ');// インデント作成
    msg = (_newline)? indent+msg: msg;                  // 行頭のインデント
    msg = msg.replace(/(?!\n$)\n/g, '\n'+indent);       // 末尾以外の改行直後にインデント
    
    if (level === _this.TEMP) {
      // 一時表示(標準出力の削除+出力)
      var tlen1 = _mlength(_tempMessage.replace(/\s+$/,''));
      var mlen  = _mlength(msg);
      var lack  = Math.max(0, tlen1 - mlen);
      if (this.stdout) {
        // 標準出力の出力
        this.streams[0].Write(
                Array(tlen0+1).join('\b')       // カーソルを戻す(前回の一時表示分)
              + msg                             // 表示
              + Array(lack+1).join(' ')         // 不足分をスペース埋め
              + Array(lack+1).join('\b'));      // カーソルを戻す(不足分)
        _tempMessage = msg;
      }
    } else if (this.isOutput(level)) {
      // 通常出力
      if (tlen0 > 0) {
        if (this.stdout && level !== _this.TRACE) {
          // 標準出力の削除
          this.streams[0].Write(
                Array(tlen0+1).join('\b')       // カーソルを戻す(前回の一時表示分)
              + Array(tlen0+1).join(' ')        // 不足分をスペース埋め
              + Array(tlen0+1).join('\b'));     // カーソルを戻す(不足分)
          _tempMessage = '';
        }
      }
      
      var i = (!this.stdout || level === _this.TRACE)? 1: 0;
      for (; i<this.streams.length; i++) {
        try {
          // 出力
          this.streams[i].Write(msg);
        } catch(e) {}
      }
    }
  };
  _this.prototype.print = function Console_print(message, opt_level) {
    message = (message !== void 0)? message: '';
    this.printCore(message, opt_level);
    _newline = message.endsWith('\n');
  };
  _this.prototype.println = function Console_println(message, opt_level) {
    this.printCore(((message !== void 0)? message: '')+'\n', opt_level);
    _newline = true;
  };
  
  /**
   * グループ
   * 以降の出力をインデントする。
   * @param {string} opt_label - ラベル
   */
  _this.prototype.group = function Console_group(opt_label) {
    if (opt_label != void 0 && (''+opt_label).length != 0) {  this.println('['+opt_label+']');  }
    this.groups.push(opt_label);
    // [label]
    //   
  }
   _this.prototype.groupEnd = function Console_groupEnd() {
    this.groups.pop();
  }
  
  /**
   * カウント
   * 特定のcount()を呼び出した回数を記録する。
   * @param {string} opt_label - ラベル
   * @param {number} [opt_level=INFO(6)] - 出力レベル
   */
  _this.prototype.count = function Console_count(opt_label, opt_level) {
    var msg = (opt_label != null && (''+opt_label).length != 0)? opt_label+': ': '';
    if (!_this.counterSet[opt_label]) {  this.counterSet[opt_label] = 0; }
    this.counterSet[opt_label]++;
    this.println(msg+this.counterSet[opt_label], opt_level);
    // label: ${count}
    // ${count}
  };
  
  /**
   * 時間
   * 操作の所要時間を追跡するために使用できるタイマーを開始する。
   * @param {string} opt_label - ラベル
   * @param {number} [opt_level=INFO(6)] - 出力レベル
   */
  _this.prototype.time = function Console_time(opt_label) {
    _this.timerSet[opt_label] = Date.now();
  };
  _this.prototype.timeEnd = function Console_timeEnd(opt_label, opt_level) {
    var msg = (opt_label != null && (''+opt_label).length != 0)? opt_label+': ': '';
    if (!this.timerSet[opt_label]) {  this.timerSet[opt_label] = 0; }
    var elapsed = Date.now() - this.timerSet[opt_label];
    this.println(msg+elapsed+'ms', opt_level);
    // label: ${time}ms
    // ${time}ms
  };
  
  /**
   * 一時表示
   * @param {string} message - 出力文字列
   * @param {boolean} [opt_newline=false] - 新規行
   */
  _this.prototype.temp = function Console_temp(message, opt_newline) {
    // 新規行 && 改行直後でない -> 改行する
    if (opt_newline === true && !_newline) { this.println(''); }
    this.print(message, _this.TEMP);
  };
  
  /**
   * 出力
   * @param {string} message - 出力文字列
   * @param {number} [opt_level=INFO(6)] - 出力レベル
   */
  _this.prototype.log = function Console_log(message, opt_level) {
    var level = (opt_level != null)? opt_level: _this.INFO;
    var msg = (message !== void 0)? message: '';
    if (_isString(this.logformat)) {
      var format = _dateFormat(this.logformat, new Date(), '${', '}')
          .replace('${prefix}', (this.prefixSet[level])? this.prefixSet[level]: '')
          .replace('${message}', msg);
      this.println(format, level);
    } else {
      var pre = (this.logprefix && this.prefixSet[level])? this.prefixSet[level]+' ': '';
      this.println(pre+msg, level);
    }
    // ${prefix} ${message}
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
  _this.prototype.config = function Console_config(message) {
    this.log(message, _this.CONFIG);
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
  _this.prototype.trace = function Console_trace(message) {
    this.log(message, _this.TRACE);
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
        msg += ErrorUtility.trace(2, true, true);
        separator = '] '
      }
      msg += separator;
      msg += 'Assertion failed';
      msg += (message)? ': '+message: '.';
      this.warn(msg);
      // [関数名(引数, ...) (ファイル名:行数:列数)] Assertion failed.
      // [関数名(引数, ...) (ファイル名:行数:列数)] Assertion failed: ${msg}
    }
  };
  
  /**
   * エラー出力
   * @param {Error} e - エラー
   */
  _this.prototype.printStackTrace = function Console_printStackTrace(e) {
    if (!_isError(e)) {
      e = _error('"e" is not a type of Error.', null, 'TypeError');
    }
    if (ErrorUtility != null) {
      if (!e.stackframes) {
        ErrorUtility.captureStackTrace(e);
      }
      this.error(ErrorUtility.stack(e));
    } else {
      this.error(_errormessage(e));
    }
  };
  
  /**
   * アニメーション
   * 一時表示を利用した、CUIアニメーションを行う。
   * 同期処理中は、アニメーションを更新できない可能性がある。
   * 一時表示中に、別途出力すると、表示が崩れる可能性がある。
   */
  var Animation = (function Console$Animation_factory(_parent) {
    "use strict";
    
    var _this = void 0;
    
    _this = function Console$Animation_constructor() {
      this.initialize.apply(this, arguments);
    };
    
    // 初期化
    _this.prototype.initialize = function Console$Animation_initialize() {
      // 変更可の変数(ただし、start()前に変更しておくこと)
      this.delay    = 0;          // 初回待機時間(ms単位)
      this.interval = 1000;       // 継続待機時間(ms単位)
      this.count    = 0;          // カウント値(0<=value)
      this.addition = 1;          // 加算値(0<value)
      this.max      = -1;         // 最大回数(-1:無効)
      //this.createAnimeText;       // 表示文字列作成関数
      
      // 変更不可の変数
      this.delayId    = null;     // private: timeoutのid
      this.intervalId = null;     // private: intervalのid
      this.console    = _parent.getConsole();
    }
    // クローン作成
    _this.prototype.clone = function Console$Animation_clone() {
      var obj  = Object(this);
      var copy = new _this();
      for (var attr in obj) {
        if (obj.hasOwnProperty(attr)) copy[attr] = obj[attr];
      }
      return copy;
    }
    // アニメーションの文字列作成
    _this.prototype.createAnimeText = function Console$Animation_createAnimeText() {
      return ''+this.count;
    };
    // カウンタをリセットする
    _this.prototype.reset = function Console$Animation_reset(count) {
      if (this.delayId !== null || this.intervalId !== null) {
        this.stop();                            // 停止する
      }
      this.count = (count)? count: 0;           // 数値の設定
    };
    // カウンタを開始する(非同期)
    // @returns {boolean} 動作中
    _this.prototype.start = function Console$Animation_start() {
      if (this.delayId !== null || this.intervalId !== null) {
        return true;                            // 稼働中
      }
      if ((0 <= this.max) && (this.max <= this.count)) {
        return false;                           // 未稼働(既に回数超過)
      }
      
      // 周期処理を登録
      var _this = this;
      var _args = arguments;
      function intervalfunc() {
        _this.count += _this.addition;  // カウントアップ
        if (_this.intervalId === null) {
        } else if ((0 <= _this.max) && (_this.max <= _this.count)) {
          _this.stop();                 // 停止
        } else {
          // 表示更新
          _this.console.temp(_this.createAnimeText.apply(_this, _args));
        }
      }
      if (_this.delay > 0) {
        _this.delayId = window.setTimeout(
          function () {
            if (_this.delayId === null) {
            } else {
              // 初回表示
              this.console.temp(_this.createAnimeText.apply(_this, _args));
              _this.intervalId = window.setInterval(intervalfunc, _this.interval);
            }
          }, _this.delay);
      } else {
        _this.intervalId = window.setInterval(intervalfunc, _this.interval);
      }
      return true;                              // 稼働開始
    };
    // カウンタを開始する(疑似同期)
    _this.prototype.exec = function Console$Animation_exec() {
      this.start.apply(this, arguments);
      
      if (this.delayId !== null) {
        WScript.Sleep(this.delay);
      }
      while (this.intervalId !== null) {        // 完了を待機
        WScript.Sleep(this.interval);
      }
    };
    // カウンタを停止する
    _this.prototype.stop = function Console$Animation_stop() {
      var ret = -1;
      if (this.delayId !== null || this.intervalId !== null) {
        if (this.delayId !== null) {
          window.clearTimeout(this.delayId);
        }
        if (this.intervalId !== null) {
          window.clearInterval(this.intervalId);
          this.console.temp('');                // 一時表示を削除
        }
        
        this.delayId = null;
        this.intervalId = null;
        ret = this.count;
      }
      return ret;
    };
    
    return _this;
  })(_this);
    
  /**
   * アニメーションインスタンスの取得
   * @param {string} opt_name - 機能名
   * @returns {Console$Animation} アニメーション
   */
  _this.getAnimation = function Console_getAnimation(opt_name) {
    var anime = null;
    if (opt_name) {
      anime = _animeSet[opt_name];
    }
    if (!anime) {                       // 未作成ならば作成
      anime = new Animation();
      if (opt_name) {
        // サンプル(clone()使用することを想定)
        switch (opt_name) {
        case 'countup':
          anime.digits  = 3;            // 最低桁数
          anime.padding = ' ';          // パディング文字
          anime.bformat = '[';
          anime.aformat = 's]';
          anime.createAnimeText = function Console$Countup_createAnimeText() {
            var digits = Math.max(this.digits, (this.max+'').length);
            return ''
              + this.bformat
              + (Array(digits+1).join(this.padding)+(this.count+1)).slice(-digits)
              + this.aformat;
          };
          break;
        case 'countdown':
          anime.digits  = 3;            // 最低桁数
          anime.padding = ' ';          // パディング文字
          anime.bformat = '[';
          anime.aformat = 's]';
          anime.createAnimeText = function Console$Countdown_createAnimeText() {
            var digits = Math.max(this.digits, (this.max+'').length);
            return ''
              + this.bformat
              + (Array(digits+1).join(this.padding)+(this.max-this.count)).slice(-digits)
              + this.aformat;
          };
          break;
        case 'animation':
          anime.animations = ['   ','.  ','.. ','...'];
          //anime.animations = ['-','\\','|','/'];  // フォント次第で￥になるため、注意
          anime.createAnimeText = function Console$AnimeText_createAnimeText() {
            return this.animations[this.count%this.animations.length];
          };
          break;
        case 'sleep':
          anime.createAnimeText = function Console$Sleep_createAnimeText() {
            var digits = Math.max(3, (this.max+'').length);
            return ''
              + '[Console.Sleep('+(this.max*this.interval+this.delay)+'ms):'
              + (Array(digits+1).join(' ')+(this.max-this.count)).slice(-digits)
              + 's]';
          };
          break;
        }
        _animeSet[opt_name] = anime;
      }
    }
    return anime;
  };
    
  /**
   * アニメーション付き待機
   * Console.getAnimation('sleep')のアニメーションを使用して指定時間待機する。
   * @param {number} milliseconds - 待機時間(ms)
   */
  _this.sleep = 
  _this.Sleep = 
  _this.prototype.sleep = 
  _this.prototype.Sleep = function Console_sleep(milliseconds) {
    var anime = _this.getAnimation('sleep');
    anime.count = 0;
    anime.max   = Math.floor(milliseconds / anime.interval);
    anime.delay = 0;
    anime.exec();
  };
  
  (function Console_main() {
    var console = _this.getConsole();
    
    // 標準出力のロガーを設定
    console.stdout = _WScript_getNamedArgument('stdout', true);
    console.setLevel(_WScript_getNamedArgument('stdout', 'INFO').toUpperCase());
    
    // ファイル出力設定
    console.logformat = _WScript_getNamedArgument('logformat', null);
    if (_WScript_getNamedArgument('logfile', false)) {
      var logpath = _WScript_getNamedArgument('logfilepath');
      var append  = _WScript_getNamedArgument('logappend');
      var format  = _WScript_getNamedArgument('logencode');
      if (_isString(logpath)) {
        // logpath    新規書き込み
        console.addOutFile(logpath, (append === true), format);
      } else if (_WScript_getNamedArgument('logrotate', false)) {
        // ./log/{FileBaseName}.{yyyMMd}.log    追加書き込み
        var now = new Date();
        var fullpath  = _getScriptPath();
        var folderpath= fs.GetParentFolderName(fullpath);
        var base      = fs.GetBaseName(fullpath);
        var date      = _dateFormat('yyyyMMdd', now);
        var folder    = fs.BuildPath(folderpath, 'log');
        var name      = base+'.'+date+'.log';
        console.addOutFile(fs.BuildPath(folder, name), (append !== false), format);
      } else {
        // ./{FileBaseName}.log    新規書き込み
        console.addOutFile(_getScriptPath('log'), (append === true), format);
      }
    }
  })();
  
  return _this;
});
