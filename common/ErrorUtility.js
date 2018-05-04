/*!
 * WSH(JScript)用エラー出力ライブラリ
 * @requires    Process.js
 * @requires      module:ActiveXObject('Scripting.FileSystemObject')
 * @requires      module:ActiveXObject('WScript.Shell')
 * @requires      module:ActiveXObject('MSXML.DOMDocument')
 * @requires      module:WScript
 * @requires      module:FileUtility.js
 * @requires        module:ActiveXObject('ADODB.Stream')
 * @auther      toshi(https://www.bugbugnow.net/)
 * @license     MIT License
 * @version     4
 * @see         2 - update - PrivateUnderscore.jsを導入 - 共通処理化するため
 * @see         3 - update - JSONを排除 - 依存排除のため
 * @see         4 - update - 文字列作成処理の外部制御を用意にする
 */
(function(root, factory) {
  if (!root.ErrorUtility) {
    root.ErrorUtility = factory();
  }
})(this, function(Process) {
  "use strict";
  
  var _this = void 0;
  
  /**
   * PrivateUnderscore.js
   * @version   1
   */
  {
    // https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/indexOf
    function _Array_indexOf(array, searchElement, fromIndex) {
      if (array == null) throw new TypeError('"array" is null or not defined');
      var o = Object(array);
      var len = o.length >>> 0;
      if (len === 0)	return -1;
      var n = fromIndex | 0;
      if (n >= len)	return -1;
      var k = n >= 0 ? n : Math.max(len + n, 0);
      for (; k < len; k++)  if (k in o && o[k] === searchElement) return k;
      return -1;
    };
    // https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Number/isInteger
    function _Number_isInteger(value) {
      return typeof value === 'number' && isFinite(value) && Math.floor(value) === value;
    };
    // https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date/toISOString
    function _Date_toISOString() {
      function pad(number) {
        return ((number < 10)? '0': '')+number;
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
    function _getGlobalObject(obj, name) {
      return (obj != null)? obj: Function('return this')()[name];
    };
    function _isString(obj) {
      return Object.prototype.toString.call(obj) === '[object String]';
    };
    function _isFunction(obj) {
      return Object.prototype.toString.call(obj) === '[object Function]';
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
    function _shortMessage(type, msg, max) {
      function _msubstr(str, start, length) {
        var s = 0;
        if (start != 0) {
          var mstart = Math.abs(start),
              direct = start/mstart,
              len = 0,
              i1 = (start < 0)? str.length-1: 0;
          for (; 0<=i1 && i1<str.length; i1+=direct) {
            len += (str.charCodeAt(i1) > 255)? 2: 1;
            if (mstart < len) {
              i1 -= direct;
              break;
            }
          }
          s = Math.max(0, i1);
        }
        var i2 = str.length;
        if (length !== void 0) {
          var len = 0;
          for (i2=s; i2<str.length; i2++) {
            len += (str.charCodeAt(i2) > 255) ? 2: 1;
            if (length < len) {
              break;
            }
          }
        }
        return str.substr(s, i2-s);
      };
      
      msg = msg+'';
      var m = _msubstr(msg, 0, max);
      if (m.length < msg.length) {
        if (max <= 2) {
          m = '..'.substr(0, max);
        } else if (type === 'middle') {
          var m1 = _msubstr(msg, 0, Math.ceil((max/2)-1)),
              m2 = (Math.floor((max/2)-1) == 0) ? 
                    '': 
                    _msubstr(msg,-Math.floor((max/2)-1)),
              dn = max - _mlength(m1) - _mlength(m2);
          m = m1 + Array(dn + 1).join('.') + m2;
        } else if (type === 'head') {
          m = _msubstr(msg, -(max - 2));
          m = Array(max - _mlength(m) + 1).join('.') + m;
        } else {
          m = _msubstr(msg, 0, max - 2);
          m = m + Array(max - _mlength(m) + 1).join('.');
        }
      }
      return m;
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
  }
  
  /**
   * コンストラクタ
   * @constructor
   */
  _this = function ErrorUtility_constrcutor() {};
  
  // 変数
  _this.stackTraceLimit = -1;                   // 最大トレース数(-1:無制限)
  _this.stackTraceArgumentMessageLimit = 128;   // 引数文字列最大桁数(全角考慮)
  _this.stackTracePrefix = '    at ';           // スタックトレースの前置詞
  
  /**
   * スタックトレース抽出
   * 自身の直前の関数からトレースする。
   * 第2引数を指定することで、指定した関数の次関数からトレースする。
   * 再帰呼び出しがある場合、トーレスを中断する。
   * @param {Error} error - エラー
   * @param {Function} opt_root - トレースを開始する直前の関数
   * @return {Error} エラー
   */
  _this.captureStackTrace = function ErrorUtility_captureStackTrace(error, opt_root) {
    var callee = _getCallee();
    error = (_isError(error))? error: _this.create('"error" is not Error.', callee, 'TypeError');
    
    var stack = [];
    var funcs = []; // 再帰呼び出し検出用
    
    // トレース開始関数決定
    var func  = callee;
    if (opt_root) {
      // ルート関数までスキップ
      for (;func && func!==opt_root; func=func.caller) {
        // 再帰呼び出し検出
        if (_Array_indexOf(funcs, func) !== -1) {
          error.stackframes = stack;
          return error;
        }
        funcs.push(func);
      }
    }
    func = func.caller;
    
    for (; func; func=func.caller) {
      // 再帰呼び出し検出
      if (_Array_indexOf(funcs, func) !== -1) {
        stack.push([null, func]);
        break;
      }
      funcs.push(func);
      
      // スタックに追加
      var frame = [func];
      for (var i=0; i<func.arguments.length; i++) {
        frame.push(func.arguments[i]);
      }
      stack.push(frame);
    }
    error.stackframes = stack;
    return error;
    // 補足:再帰があると、それ以前までたどれないため、中断する
    // 補足:関数内で引数の値を変更すると、変更後の値を表示する
  };
  
  /**
   * 文字列化する
   * オブジェクト等を文字列化する。
   * @protected
   * @param {*} message - メッセージ
   * @param {*} isString - 文字列表現するか
   * @return {string} 出力文字列
   */
  _this._toString = function ErrorUtility__toString(message, isString) {
    return _toString(message, isString);
  };
  
  /**
   * トレース
   * frameから文字列を作成する
   * 動作例
   *  呼出し: func(true, 123, "abc", new Date(0), null, void 0, 
   *            {a:0,b:"x",c:[]}, [0,"a",{}], function f0(){}, Math, /test/g, new Error("x"))
   *  表示:   func(true, 123, "abc", "1970-01-01T00:00:00.000Z", null, undefined, 
   *            {a:0,b:"x",c:[]}, [0,"a",{}], f0(), Math, /test/g, Error{message:"x"})
   * @param {(Array|number)} frame - エラー情報
   * @param {Function} frame[0] - 関数
   * @param {*} frame[n] - 引数
   * @param {boolean} [opt_argsinfo=false] - 引数表示(true:表示/false:非表示)
   * @param {(boolean|string)} [opt_exinfo=false] - 拡張情報(true:拡張情報/false:なし/string:文字列)
   * @return {string} エラー文字列
   */
  _this.trace = function ErrorUtility_trace(frame, opt_argsinfo, opt_exinfo) {
    var msg = '';
    
    // frameを間接指定(本関数を0とする)
    if (_Number_isInteger(frame)) {
      var frames = _this.create().stackframes;
      if (0 <= frame && frame < frames.length) {
        frame = frames[frame];
      } else {
        throw new RangeError('"frame" is out of the range of "stackframes".');
      }
    }
    
    // 関数名
    msg += _getFunctionName(frame[0]) + '(';
    
    // 引数
    if (opt_argsinfo === true && _this.stackTraceArgumentMessageLimit >= 2) {
      var args = frame.slice(1);
      var m;
      for (var i=0; i<args.length; i++) {
        m = _this._toString(args[i], true);
        m = _shortMessage('middle', m, _this.stackTraceArgumentMessageLimit);
        msg += ((i !== 0)? ', ': '') + m;
      }
    }
    
    msg += ')';
    
    // 拡張情報(ファイル名:行数:列数/文字列)
    if (opt_exinfo === true) {
      // 検索関数がある時
      var _Process = _getGlobalObject(Process, 'Process');
      if (_Process && _Process.searchScripts) {
        try {
          var scripts = _Process.searchScripts(frame[0]);
          if (scripts.length == 1) {
            msg += ' - '+scripts[0].name+':'+scripts[0].row+':'+scripts[0].column+'';
          }
          // 補足:0個の可能性あり、2個以上の可能性あり、全部表示するのもあり？
        } catch (e) {}  // エラー出力処理中であるため、なにもしない
      }
    } else if (_isString(opt_exinfo)) {
      msg += ' - ' + opt_exinfo;
    }
    return msg;
    // 関数名(引数, ...) - ファイル名:行数:列数
    // 関数名(引数, ...) - 文字列
  };
  
  /**
   * スタックトレース
   * stackframesから文字列を作成する。
   * stackframesが未登録の場合、空文字列を返す。
   * @param {Error} error - エラー
   * @param {boolean} [opt_error=true] - エラーであるか(false:エラー名称を非表示)
   * @return {string} エラー文字列
   */
  _this.stack = function ErrorUtility_stack(error, opt_error) {
    var callee = _getCallee();
    if (!_isError(error)) { error = _this.create('"error" is not Error.', callee, 'TypeError'); }
    
    var msgs = [];
    var frames = error.stackframes;
    if (frames == null) {
      return '';
    }
    
    // スタックトレース前の文字列
    if (opt_error !== false) {
      msgs.push(_errormessage(error));
    } else if (error.message && error.message.length != 0) {
      msgs.push(error.message);
    }
    
    // スタックトレース
    for (var i=0; i<frames.length; i++) {
      var msg = _this.stackTracePrefix;
      if (_this.stackTraceLimit >= 0 && i >= _this.stackTraceLimit) {
        break;                      // 最大数に到達
      }
      if (frames[i][0] === null) {                // 再帰呼び出し
        msgs.push(msg + _getFunctionName(frames[i][1]) + '()...');
        break;
        // 補足:引数は、再帰呼び出しの値であるため、表示しない
      }
      msg += _this.trace(frames[i], true, true);
      msgs.push(msg);
      // 補足:例「    at 関数名(引数, ...) - ファイル名:行数:列数」
      // 注意:行数と列数は、対象関数の先頭であり、エラー発生箇所ではない
    }
    return msgs.join('\n');
  };
  
  /**
   * エラー作成
   * エラー作成と同時にキャプチャする。
   * エラー作成箇所は、本関数を直接コールすること。
   * @param {string} opt_message - エラーメッセージ
   * @param {Function} opt_root - トレースを開始する直前の関数(null:本関数)
   * @param {(string|Function)} opt_error - エラー
   * @return {Error} エラー
   */
  _this.create = function ErrorUtility_create(opt_message, opt_root, opt_error) {
    var error = (_isFunction(opt_error))? new opt_error(opt_message): new Error(opt_message);
    
    // 不要な変数の削除(必要な場合、別途追加すること)
    // Error, TypeErrorで挙動が異なる。
    // 共通化と文字列表現の見栄えのために削除する。
    delete error.number;
    delete error.description;
    
    // 名称変更
    if (_isString(opt_error)) {
      error.name = opt_error;
    }
    
    // stackframe作成
    var root = (_isFunction(opt_root))? opt_root: _getCallee();
    _this.captureStackTrace(error, root);
    
    // 補助情報を追加
    var _Process = _getGlobalObject(Process, 'Process');
    if (_Process && _Process.searchScripts) {
      try {
        var scripts = _Process.searchScripts(root.caller);
        if (scripts.length === 1) {
          error.fileName = scripts[0].name;
          error.lineNumber = scripts[0].row;
          error.columnNumber = scripts[0].column;
        }
      } catch (e) {}  // エラー出力処理中であるため、なにもしない
    }
    
    return error;
  };
  
  return _this;
});
