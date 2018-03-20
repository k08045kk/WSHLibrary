/*!
 * WSH(JScript)用エラー出力ライブラリ
 * @requires    module:JSON
 * @requires    module:Process.js
 * @auther      toshi(http://www.bugbugnow.net/)
 * @license     MIT License
 * @version     2
 * @see         2 - PrivateUnderscore.jsを導入 - ソフトに共通処理化するため
 */
(function(global, factory) {
  if (!global.ErrorUtility) {
    global.ErrorUtility = factory(global, global.JSON);
  }
})(this, function(global, JSON) {
  "use strict";
  
  /**
   * PrivateUnderscore.js
   * @version   1
   */
  {
    // json3.js用
    global.dontEnums = global.dontEnums || [
      'toString',
      'toLocaleString',
      'valueOf',
      'hasOwnProperty',
      'isPrototypeOf',
      'propertyIsEnumerable',
      'constructor'
    ];
    // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/indexOf
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
    // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date/toISOString
    function _Date_toISOString(date) {
      function pad(number) {
        return ((number < 10)? '0': '')+number;
      }
      return date.getUTCFullYear() +
        '-' + pad(date.getUTCMonth() + 1) +
        '-' + pad(date.getUTCDate()) +
        'T' + pad(date.getUTCHours()) +
        ':' + pad(date.getUTCMinutes()) +
        ':' + pad(date.getUTCSeconds()) +
        '.' + (date.getUTCMilliseconds() / 1000).toFixed(3).slice(2, 5) +
        'Z';
    };
    // https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/Number/isInteger
    function _isInteger(value) {
      return typeof value === 'number' && isFinite(value) && Math.floor(value) === value;
    };
    function _isObject(obj) {
      var type = typeof obj;
      return type === 'function' || type === 'object' && !!obj;
    };
    function _isString(obj) {
      return Object.prototype.toString.call(obj) === '[object String]';
    };
    function _isDate(obj) {
      return Object.prototype.toString.call(obj) === '[object Date]';
    };
    function _isArray(obj) {
      return Object.prototype.toString.call(obj) === '[object Array]';
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
      var msg = '';
      msg += (error.name)? error.name: 'UnknownError';
      msg += '(';
      // 機能識別符号.エラーコード
      msg += (error.number)? ((error.number>>16)&0xFFFF)+'.'+(error.number&0xFFFF): '';
      msg += ')';
      msg += (error.message)? ': '+error.message: '';
      return msg;
    }
  }
  
  var _this = void 0;
  _this = function ErrorUtility_constrcutor() {};
  
  // 変数
  _this.stackTraceLimit = -1;                   // 最大トレース数(-1:無制限)
  _this.stackTraceArgumentMessageLimit = 128;   // 引数文字列最大桁数(全角考慮)
  
  /**
   * スタックトレース抽出
   * 自身の直前の関数からトレースする。
   * 第2引数を指定することで、指定した関数の次関数からトレースする。
   * 再帰呼び出しがある場合、トーレスを中断する。
   * @param {Error} error - エラー
   * @param {Function} opt_root - トレースを開始する直前の関数
   * @returns {Error} エラー
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
   * トレース
   * frameから文字列を作成する
   * 動作例
   *  エラー発生関数呼び出し: 
   *    func(true, 123, "abc", new Date(0), null, void 0, 
   *          {a:0,b:"x",c:[]}, [0,"a",{}], function f0(){}, Math, /test/, new Error("x"))
   *  JSONあり: 
   *    func(true, 123, "abc", "1970-01-01T00:00:00.000Z", null, undefined, 
   *          {"a":0,"b":"x","c":[]}, [0,"a",{}], f0(), {}, {}, {"description":"x","message":"x"})
   *  JSONなし: 
   *    func(true, 123, "abc", "1970-01-01T00:00:00.000Z", null, undefined, 
   *          [object Object], [object Array], f0(), [object Math], [object RegExp], [object Error])
   * @param {(Array|number)} frame - エラー情報
   * @param {Function} frame[0] - 関数
   * @param {*} frame[n] - 引数
   * @param {boolean} [opt_argsinfo=false] - 引数表示(true:表示/false:非表示)
   * @param {(boolean|string)} [opt_exinfo=false] - 拡張情報(true:拡張情報/false:なし/string:文字列)
   * @returns {string} エラー文字列
   */
  _this.trace = function ErrorUtility_trace(frame, opt_argsinfo, opt_exinfo) {
    var msg = '';
    
    // frameを間接指定(本関数を0とする)
    if (_isInteger(frame)) {
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
      for (var i=0; i<args.length; i++) {
        var m = args[i]+'';
        if (_isFunction(args[i])) {
          m = _getFunctionName(args[i])+'()';
        } else if (JSON != null) {
          try {
            m = JSON.stringify(args[i]);
          } catch (e) {}  // 文字列化できない要素対策(WScript固有変数など)
        } else if (_isString(args[i])) {
          m = '"'+args[i]+'"';
        } else if (_isDate(args[i])) {
          m = '"'+_Date_toISOString(args[i])+'"';
        } else if (_isObject(args[i]) || _isArray(args[i])) {
          m = Object.prototype.toString.call(args[i]);
        }
        m = _shortMessage('middle', m, _this.stackTraceArgumentMessageLimit);
        msg += ((i !== 0)? ', ': '') + m;
      }
    }
    
    msg += ')';
    
    // 拡張情報(ファイル名:行数:列数/文字列)
    if (opt_exinfo === true) {
      // 検索関数がある時
      if (global.Process && global.Process.searchScripts) {
        try {
          var scripts = global.Process.searchScripts(frame[0]);
          if (scripts.length == 1) {
            var script = scripts[0];
            var filename = script.name;
            msg += ' - '+filename+':'+script.row+':'+script.column+'';
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
   * stackframesから文字列を作成する
   * @param {Error} error - エラー
   * @param {boolean} [opt_error=true] - エラーであるか(false:エラー名称を非表示)
   * @returns {string} エラー文字列
   */
  _this.stack = function ErrorUtility_stack(error, opt_error) {
    var callee = _getCallee();
    if (!_isError(error)) { error = _this.create('"error" is not Error.', callee, 'TypeError'); }
    
    var msgs = [];
    var frames = error.stackframes;
    if (frames == null) { frames = _this.captureStackTrace(error, callee).stackframes;  }
    
    // スタックトレース前の文字列
    if (opt_error !== false) {
      msgs.push(_errormessage(error));
    } else if (error.message) {
      msgs.push(error.message);
    }
    
    // スタックトレース
    for (var i=0; i<frames.length; i++) {
      var msg = '    at '
      if (_this.stackTraceLimit >= 0 && i >= _this.stackTraceLimit) {
        break;                      // 最大数に到達
      }
      if (frames[i][0] === null) {                // 再帰呼び出し
        msgs.push(msg + _getFunctionName(frames[i][1]) + '()...');
        break;
        // 補足:引数は、最後の呼び出しの値であるため、表示しない
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
   * @param {?Function} opt_root - トレースを開始する直前の関数(null:本関数)
   * @param {(string|Error)} opt_error - エラー名称
   * @return {Error} エラー
   */
  _this.create = function ErrorUtility_create(opt_message, opt_root, opt_error) {
    var error = (_isError(opt_error))? new opt_error(): new Error();
    
    if (_isString(opt_error)) { error.name = opt_error; }
    if (opt_message != null) {  error.message = opt_message;  }
    
    var root = (_isFunction(opt_root))? opt_root: _getCallee();
    _this.captureStackTrace(error, root);
    
      if (global.Process && global.Process.searchScripts) {
      try {
        var scripts = global.Process.searchScripts(root.caller);
        if (scripts.length === 1) {
          var script = scripts[0];
          error.fileName = script.name;
          error.lineNumber = script.row;
          error.columnNumber = script.column;
        }
      } catch (e) {}  // エラー出力処理中であるため、なにもしない
    }
    
    return error;
  };
  
  return _this;
});
