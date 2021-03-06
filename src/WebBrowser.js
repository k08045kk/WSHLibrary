/*!
 * WebBrowser.js v5
 *
 * Copyright (c) 2018 toshi (https://github.com/k08045kk)
 *
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */

/**
 * WSH(JScript)用ブラウザ操作補助
 * Console.jsと組み合わせることで実行中ログ出力機能を追加する。
 * Console.Animation.jsと組み合わせることで待機待ちをアニメーション表示する。
 * ErrorUtility.jsと組み合わせることで呼び出し元関数の引数までトレースする。
 * @requires    module:ActiveXObject("InternetExplorer.Application")
 * @requires    module:JSON
 * @requires    module:WScript
 * @requires    ErrorUtility.js
 * @requires    Console.js
 * @requires    Console.Animation.js - v5+
 * @auther      toshi (https://github.com/k08045kk)
 * @version     5
 * @see         1.20180319 - add - 初版
 * @see         2.20190823 - update - Navigateを閲覧中の場合、再閲覧しないように仕様変更
 * @see         2.20190823 - update - doNavigateを追加
 * @see         2.20190823 - update - アニメーション関連をリファクタリング
 * @see         2.20190823 - update - コマンドライン引数を自動で確認しないように変更
 * @see         2.20190918 - fix - Navigateの戻り値がなくなっていた
 * @see         3.20191031 - update - doSubmitNotParamsLog()を追加、パスワード漏洩対策
 * @see         4.20191214 - fix - WebBrowser.css()が動作していなかった
 * @see         4.20191214 - update - js(), css()のラップ関数を準備
 * @see         5.20200116 - fix - _random() が定義済みだったが、間違って random() が使用されている
 * @see         5.20200202 - update - _getParentElement() の処理改善
 */
(function(root, factory) {
  if (!root.WebBrowser) {
    root.WebBrowser = factory(root.ErrorUtility, root.Console);
  }
})(this, function(ErrorUtility, Console) {
  "use strict";
  
  // -------------------- private --------------------
  
  var _this = void 0;
  
  /**
   * PrivateUnderscore.js
   * @version   4
   */
  {
    function _isString(obj) {
      return Object.prototype.toString.call(obj) === '[object String]';
    };
    function _isArray(obj) {
      return Object.prototype.toString.call(obj) === '[object Array]';
    };
    function _isElement(obj) {
      return !!(obj && obj.nodeType === 1);
    };
    function _getCallee() {
      var args = arguments;
      var func = args.callee;
      return func.caller;
    };
    function _extend(dst, src, undefinedOnly) {
      if (dst != null && src != null) {
        for (var key in src) {
          if (src.hasOwnProperty(key)) {
            if (!undefinedOnly || dst[key] === void 0) {
              dst[key] = src[key];
            }
          }
        }
      }
      return dst;
    };
    function _random(min, max) {
      return min + Math.floor(Math.random() * (max - min + 1));
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
    function _getParentElement(element, tagName) {
      for (; element; element=element.parentElement) {
        if (element.tagName === tagName) {      // 大文字
          return element;
        }
      }
      return null;
    }
  }
  
  
  // -------------------- static --------------------
  
  _this = function WebBrowser_constructor() {
    this.initialize.apply(this, arguments);
  };
  
  _this.READYSTATE_UNINITIALIZED= 0;  // ReadyState: 未完了状態
  _this.READYSTATE_LOADING      = 1;  // ReadyState: ロード中状態
  _this.READYSTATE_LOADED       = 2;  // ReadyState: ロード完了状態、ただし操作不可能状態
  _this.READYSTATE_INTERACTIVE  = 3;  // ReadyState: 操作可能状態
  _this.READYSTATE_COMPLETE     = 4;  // ReadyState: 全データ読み込み完了状態
  
  /**
   * UserAgent
   * Desktop: IE11(Windows)
   * Mobile:  Firefox(Android)
   * TODO:随時更新が必要
   *      https://developer.mozilla.org/ja/docs/Web/HTTP/Gecko_user_agent_string_reference
   */
  _this.Desktop = 'Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; Touch; rv:11.0) like Gecko';
  _this.Mobile  = 'Mozilla/5.0 (Android 4.4; Mobile; rv:41.0) Gecko/41.0 Firefox/41.0';
  _this.UserAgent = _this.Desktop;
  
  _this.animation = false;      // 待機アニメーションを使用する
  
  /**
   * 親要素のタグを探す
   * @param {Element} element - 要素
   * @param {string} tagName - 探すタグ
   * @return {Element} 親要素
   */
  _this.getParentElement = 
  _this.prototype.getParentElement = function WebBrowser_getParentElement(element, tagName) {
    return _getParentElement(element, tagName.toUpperCase());
  };
  
  /**
   * IEの読み込みを待機
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {number} [opt_state=3] - 状態(READYSTATE参照)
   * @param {number} [opt_timeout=60000] - 待機時間(ms)
   */
  _this.wait = function WebBrowser_static_wait(ie, opt_state, opt_timeout) {
    var state = (opt_state == null)? _this.READYSTATE_INTERACTIVE: opt_state;
    var timeout = (opt_timeout == null)? 60*1000: opt_timeout;
    
    if (ie.Busy || (ie.readystate < state)) {   // 重複呼び出し対策
      var begin = new Date().getTime();
      var waiting = null;
      if (Console && Console.getAnimation && _this.animation) {
        // アニメーションスリープ設定
        waiting = Console.getAnimation('WebBrowser.wait');
        if (waiting.created !== true) {
          waiting.delay = 1000;
          waiting.prefix = '';
          waiting.suffix = '';
          waiting.animations = ['   ','.  ','.. ','...'];
          waiting.createAnimeText = function WebBrowser_createAnimeText() {
            return this.prefix+this.animations[this.count%this.animations.length]+this.suffix;
          };
          waiting.created = true;
        }
      }
      
      if (waiting) {
        waiting.reset();
        waiting.start();
      }
      for (var i=0; i<2; i++) {                 // 安全のため、2回実行
        // busyの場合、あるいは、読み込み中の場合は、100ミリ秒スリープする
        while (ie.Busy || (ie.readystate < state)) {
          WScript.Sleep(100);
          if (begin+timeout < new Date().getTime()) {
            // タイムアウト時間経過
            var message = 'time out. (Busy='+ie.Busy+', readystate='+ie.readystate+')';
            throw (ErrorUtility)? ErrorUtility.create(message): new Error(message);
          }
        }
      }
      if (waiting) {
        waiting.stop();
      }
      
      // 安全のため、最短300msを確保
      WScript.Sleep(300);
    }
    // 補足:デフォをREADYSTATE_INTERACTIVE(3)とする。
    //      完全な完了は、4であるが、4にならないページ等もあり、(例:楽天)
    //      3でも問題が起こらないことも多いため、デフォは3とする。
    //      4が必須の場合、明示的に再度wait(4)すること。
  };
  
  /**
   * 関数実行
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {Function} func - 実行する関数
   */
  _this.js = function WebBrowser_static_js(ie, func) {
    _this.wait(ie);     // 安全のため、待機
    
    // 動的実行
    var d = ie.document;
    var e = d.createElement('script');
    e.type = 'text/javascript';
    e.text = '('+func.toString()+')();';
    if (d.head) {       d.head.appendChild(e);  }
    else if (d.body) {  d.body.appendChild(e);  }
    else {
      var message = 'append failed. (head='+d.head+', body='+d.body+')'
      throw (ErrorUtility)? ErrorUtility.create(message): new Error(message);
    }
    // 補足:setTimeoutと異なり、HTML上に書き込むため、
    //      他のスクリプトに干渉する可能性がある
    
    _this.wait(ie);     // 安全のため、待機
    // 補足:ブックマークレット(Navigateのurlとして指定)とすると、
     //     wait()で読み込みが完了しないため、本方法を使用する。
  };
  
  /**
   * CSS追加
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {string} text - 追加するコード
   */
  _this.css = function WebBrowser_static_css(ie, text) {
    _this.wait(ie);
    
    var d = ie.document;
    var e = d.createElement('style');
    e.type = 'text/css';
    var rule = d.createTextNode(text);
    e.appendChild(rule);
    if (d.head) {       d.head.appendChild(e);  }
    else if (d.body) {  d.body.appendChild(e);  }
    else {
      var message = 'append failed. (head='+d.head+', body='+d.body+')'
      throw (ErrorUtility)? ErrorUtility.create(message): new Error(message);
    }
    
    _this.wait(ie);
  };
  
  /**
   * ページ遷移後の共通処理
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {string} confirm - confirmの戻り値
   * @param {string} confirm - promptの戻り値
   */
  _this.postNavigate = function WebBrowser_static_postNavigate(ie, confirm, prompt) {
    var ret = false;
    try {
      // 自動化の妨げとなる関数を上書き
      var func = ''
            + 'window.alert = function (message) {console.log("alert(): "+message);};'
            + 'window.onunload = function () {console.log("onunload()");};'
            + 'window.onbeforeunload = function () {console.log("onbeforeunload()");};';
      if (confirm === void 0) { confirm = true;  }      // デフォでOKとする
      if (confirm !== null) {                           // 標準指定(null)でない時
        func += ''
            + 'window.confirm = function () {'
            + '  console.log(\'confirm('+confirm+')\');'
            + '  return('+confirm+');'
            + '};';
      }
      if (prompt === void 0) { prompt = '""';  }        // デフォで空文字とする
      if (prompt !== null) {                            // 標準指定(null)でない時
        func += ''
            + 'window.prompt = function () {'
            + '  console.log(\'prompt('+prompt+')\');'  // ダブルクオートの2重対応
            + '  return('+prompt+');'
            + '};';
      }
      _this.js(ie, new Function(func));
      // 補足:onload以前は、対象外(Seleniumでも無理っぽいので諦めたほうが良さそう)
      
      _this.wait(ie);   // 安全のため、待機
      ret = true;
    } catch (e) {}      // エラーは、無視
    return ret;
  };
  
  
  // -------------------- local --------------------
  
  /**
   * コンストラクタ
   * @constructor
   */
  _this.prototype.initialize = function WebBrowser_initialize(param) {
    if (!param) { param = {}; }
    
    this.ie = null;     // IE
    this.console = Console && Console.getConsole();     // 出力コンソール
    this.setup = {};    // IEの設定
    
    // 初期化後の設定 > コマンドライン引数 > コンストラクタ引数 > 初期値
    _extend(this.setup, param);
    _extend(this.setup, {
        debug:false,            // デバッグモード
        visible:true,           // 表示有無
        top:-1, left:-1, width:-1, height:-1,   // 表示位置
        headers:void 0,         // ヘッダー
        confirm:true,           // ダイアログ応答の戻り地
        prompt:'""',            // ダイアログ応答の入力文字列
        sleep: {
          min: 5*1000,
          max: 10*1000
        },
        cache: {                // キャッシュ保持
          querySelectorAll:false,// querySelectorAllを収集する
          max: 25               // 最大キャッシュ保持数
        }
    }, true);
    
    this.clearCache();
    this.addCacheSeparator();
  };
  
  /**
   * デストラクタ
   * IEを終了する(必須ではない)。
   * 本関数を呼び出さないまま、プログラムを終了すると、IEが起動したままとなる。
   * @destructor
   */
  _this.prototype.quit = 
  _this.prototype.Quit = function WebBrowser_Quit() {
    this.addCacheSeparator();
    if (this.ie !== null) {
      // 終了しない方法も必要？(必要になってから考える)
      this.ie.Quit();
      this.ie = null;
      this.url = '';
    }
  };
  
  /**
   * コマンドライン引数の解析
   */
  _this.prototype.commandLineArgumentAnalysis = function WebBrowser_commandLineArgumentAnalysis() {
    this.addCacheSeparator();
    this.setup.debug   = _Process_getNamedArgument('debug', this.setup.debug);
    this.setup.visible = _Process_getNamedArgument('visible', this.setup.visible);
    this.setup.top     = _Process_getNamedArgument('top', this.setup.top, 0);
    this.setup.left    = _Process_getNamedArgument('left', this.setup.left, 0);
    this.setup.width   = _Process_getNamedArgument('width', this.setup.width, 1);
    this.setup.height  = _Process_getNamedArgument('height', this.setup.height, 1);
  };
  
  /**
   * キャッシュ関連
   * TODO: 外部に切り出したい
   * @deprecated 仕様変更の可能性大
   */
  function getCallerFunctionName() {
    return (ErrorUtility)?
        ErrorUtility.trace(3, true):                    // 関数名 + 引数
        _getFunctionName(_getCallee().caller.caller);   // 関数名のみ
  };
  _this.prototype.clearCache = function WebBrowser_clearCache() {
    this.cache = null;
    this.cache = [];
    this.addCacheSeparator(JSON.stringify(this.setup));
  };
  _this.prototype.addCache = function WebBrowser_addCache(cacheData, var_args) {
    var catchArray = [];
    for (var i=0; i<arguments.length; i++) {
      catchArray.push(arguments[i]);
    }
    if (this.setup.cache) {
      this.cache.push(catchArray);
      if (this.setup.cache.max != -1 && this.setup.cache.max < this.cache.length) {
        this.cache.shift();
      }
    }
    if (this.console && this.setup.debug === true) {
      for (var i=0; i<catchArray.length; i++) {
        if (i != 0) {  this.console.print(', ');  }
        this.console.print(catchArray[i]);
      }
    }
  };
  _this.prototype.addCacheEnd = function WebBrowser_addCacheEnd(cacheData, var_args) {
    if (this.setup.cache) {
      Array.prototype.push.apply(this.cache[this.cache.length-1], arguments);
    }
    if (this.console && this.setup.debug === true) {
      for (var i=0; i<arguments.length; i++) {
        this.console.print(', ');
        this.console.print(arguments[i]);
      }
    }
  };
  _this.prototype.addCacheSeparator = function WebBrowser_addCacheSeparator(message) {
    var carcheArray = [getCallerFunctionName()];
    if (message) {  carcheArray.push(message);  }
    carcheArray.push(this.getURL());
    this.addCache.apply(this, carcheArray);
    if (this.console && this.setup.debug === true) {
      this.console.println();
    }
  };
  _this.prototype.getCacheData = function WebBrowser_getCacheData() {
    var msgs = [];
    for (var i=0; i<this.cache.length; i++) {
      msgs.push(this.cache[i].join(', '));
    }
    return msgs.join('\n');
  };
  
  /**
   * 前処理
   */
  _this.prototype.init = function WebBrowser_init() {
    if (this.ie != null) {
      // IEが動作するか確認
      // 動作しない場合、終了する
      try {
        this.ie.document.querySelector('body');
      } catch (e) {
        try {
          this.ie.Quit();
        } catch (e2) {}
        this.ie = null;
      }
    }
    if (this.ie === null) {
      // IEの作成
      // 2回作成を試みる
      for (var i=1; i>=0; i--) {
        try {
          this.ie = new ActiveXObject('InternetExplorer.Application');
          break;
        } catch (e) {
          if (i === 0) {  throw (ErrorUtility)? ErrorUtility.captureStackTrace(e): e;  }
        }
        if (this.console && this.console.sleep && _this.animation) {
          var anime = Console.getAnimation('WebBrowser.init', 'Console.countdown');
          if (anime.created !== true) {
            anime.created = true;
          }
          anime.reset(30*1000);
          anime.exec();
        } else {
          WScript.Sleep(30*1000);
        }
      }
      
      // IE初期化
      if (this.setup.top >= 0) {  this.ie.Top   = this.setup.top;   }
      if (this.setup.left >= 0) { this.ie.Left  = this.setup.left;  }
      if (this.setup.width > 0) { this.ie.Width = this.setup.width; }
      if (this.setup.height > 0) {this.ie.Height= this.setup.height;}
      this.ie.Visible = this.setup.visible;
      this.ie.Navigate('about:blank');  // 空ページ表示
    }
    _this.wait(this.ie);
    this.url = this.getURL();
    
    // キャッシュ(コマンド保存)
    this.addCache(getCallerFunctionName());
    // 補足:読み込み完了状態で処理に進む
    //      読み込み中に、以降の処理を行うとエラーとなるため
  };
  
  /**
   * 後処理
   * @param {(boolean|number)} ret - 結果(成功(true)の場合、待機する。キャッシュに保存する)
   * @param {boolean} sleep - 待機有無(falseでないければ、待機する)
   */
  _this.prototype.finish = function WebBrowser_finish(ret, sleep) {
    _this.wait(this.ie);
    
    // キャッシュ(結果保存)
    this.addCacheEnd(ret, this.getURL());
    if (this.console && this.setup.debug === true) {
      this.console.println();
    }
    
    // ランダムスリープ(人間ぽさを演出するため)
    if (ret !== false) {
      var min = this.setup.sleep.min;
      var max = this.setup.sleep.max;
      if (sleep !== false && min && max) {
        var random = _random(min, max);
        if (this.console && this.console.sleep && _this.animation) {
          var anime = Console.getAnimation('WebBrowser.finish', 'Console.countdown');
          if (anime.created !== true) {
            anime.created = true;
          }
          anime.reset(random);
          anime.exec();
        } else {
          WScript.Sleep(random);
        }
      }
    }
    this.url = this.getURL();
  };
  
  /**
   * IEを取得
   * 取得した、IEを操作してもキャッシュに保存しない。
   * @return {ActiveXObject} IE
   */
  _this.prototype.getIE = function WebBrowser_getIE() {
    return this.ie;
  };
  
  /**
   * URLを取得
   * 現在表示中のURLを取得する。
   * @return {string} URL
   */
  _this.prototype.getURL = function WebBrowser_getURL() {
    return (this.ie !== null)? this.ie.LocationURL: '';
  };
  
  /**
   * Documentを取得
   * 取得した、ドキュメントを操作してもキャッシュに保存しない。
   * @return {Element} document
   */
  _this.prototype.getDocument = function WebBrowser_getDocument() {
    return (this.ie !== null)? this.ie.Document: null;
  };
  
  /**
   * 設定を保存
   * 本クラスが使用する設定を保存する。
   * 設定は、キャッシュに書き出す。
   * @param {string} name - 名称
   * @param {*} value - 値
   */
  _this.prototype.setSetupValue = function WebBrowser_setSetupValue(name, value) {
    this.addCacheSeparator();
    this.setup[name] = value;
  };
  _this.prototype.setHeaders = function WebBrowser_setHeaders(header, var_args) {
    var headers = '';
    for (var i=0; i<arguments.length; i++) {
      headers += arguments[i]+'\n';
    }
    this.setSetupValue('headers', headers);
  };
  
  _this.prototype.setConfirmAnswer = function WebBrowser_setConfirmAnswer(answer) {
    this.setSetupValue('confirm', answer);
  };
  
  _this.prototype.setPromptAnswer = function WebBrowser_setPromptAnswer(answer) {
    this.setSetupValue('prompt', JSON.stringify(answer));
  };
  
  _this.prototype.setRandomSleep = function WebBrowser_setRandomSleep(min, max) {
    this.setSetupValue('sleep', {min:min, max:max});
  };
  
  /**
   * IEの読み込みを待機
   * @param {number} state - 状態(READYSTATE参照)
   * @param {number} timeout - 待機時間(ms)
   */
  _this.prototype.wait = function WebBrowser_wait(state, timeout) {
    _this.wait(this.ie, state, timeout);
  };
  _this.prototype.waitUninitialized = function WebBrowser_waitUninitialized(timeout) {
    this.wait(_this.READYSTATE_UNINITIALIZED, timeout);
  };
  _this.prototype.waitLoading = function WebBrowser_waitLoading(timeout) {
    this.wait(_this.READYSTATE_LOADING, timeout);
  };
  _this.prototype.waitLoader = function WebBrowser_waitLoader(timeout) {
    this.wait(_this.READYSTATE_LOADED, timeout);
  };
  _this.prototype.waitInteractive = function WebBrowser_waitInteractive(timeout) {
    this.wait(_this.READYSTATE_INTERACTIVE, timeout);
  };
  _this.prototype.waitComplete = function WebBrowser_waitComplete(timeout) {
    this.wait(_this.READYSTATE_COMPLETE, timeout);
  };
  
  _this.prototype.js = function WebBrowser_js(func) {
    return _this.js(this.ie, func);
  };
  
  _this.prototype.css = function WebBrowser_css(text) {
    return _this.css(this.ie, text);
  };
  
  _this.prototype.postNavigate = function WebBrowser_postNavigate(opt_confirm, opt_prompt) {
    var confirm = (opt_confirm != null)? opt_confirm: this.setup.confirm;
    var prompt = (opt_prompt != null)? opt_prompt: this.setup.prompt;
    
    return _this.postNavigate(this.ie, confirm, prompt);
  };
  
  /**
   * 閲覧する
   * 既に閲覧中の場合、再閲覧しない。
   * @param {string} url - URL
   * @param {number} flags - ウィンドウ詳細設定
   * @param {string} targetFrameName - ウィンドウ表示設定
   * @param {string} postData - POSTデータ設定
   * @param {string} headers - HTTPヘッダー
   * @return {boolean} 成否
   */
  _this.prototype.Navigate = 
  function WebBrowser_Navigate(url, flags, targetFrameName, postData, headers) {
    var ret = false;
    if (this.getURL() != url) {
      ret = this.doNavigate.apply(this, arguments);
    }
    return ret;
  };
  
  /**
   * 強制的に閲覧する
   * 既に閲覧中の場合、再閲覧する。
   * @param {string} url - URL
   * @param {number} flags - ウィンドウ詳細設定
   * @param {string} targetFrameName - ウィンドウ表示設定
   * @param {string} postData - POSTデータ設定
   * @param {string} headers - HTTPヘッダー
   * @return {boolean} 成否
   */
  _this.prototype.doNavigate = 
  function WebBrowser_doNavigate(url, flags, targetFrameName, postData, headers) {
    headers = (headers !== void 0)? headers: this.setup.headers;
    
    this.init();
    
    this.ie.Navigate(url, flags, targetFrameName, postData, headers);
    var ret = this.postNavigate();
    
    this.finish(ret);
    return ret;
  };
  
  function query2element(element, query) {
    if (_isString(query)) {             // 文字列(クエリー)
      return element.querySelector(query);
    } else if (_isElement(query)) {     // 要素
      return query;
    }
    return null;
  };
  
  function query2elements(element, query) {
    if (_isString(query)) {             // 文字列(クエリー)
      return element.querySelectorAll(query);
    } else if (_isArray(query)) {
      return query;
    } else if (_isElement(query)) {     // 要素
      return [query];
    }
    return [];
  };
  
  /**
   * アンカーを閲覧
   * targetをクリアする。(_blankがあると別ウィンドウを表示するため)
   * @param {(string|Element)} element - 要素
   * @param {number} [number=0] - 要素番号
   * @return {boolean} 成否
   */
  _this.prototype.doAnchor = function WebBrowser_doAnchor(elements, number) {
    var ret = false;
    
    this.init();
    
    elements = query2elements(this.ie.document, elements);
    if (elements.length != 0) {
      if (number == null) {
        number = 0;
      } else if (number < 0) {
        number = _random(0, elements.length-1);
      }
      var element = elements[number%elements.length];
      this.addCacheEnd(number%elements.length);
      
      var a = _getParentElement(element, 'a');
      if (a !== null) {
        // 親アンカーのターゲットを削除(_blank対策)
        a.target = '';
      }
      
      if (a !== null && a.href) {
        this.addCacheEnd('navi');
        this.ie.Navigate(a.href, null, null, null, this.setup.headers);
      } else {
        this.addCacheEnd('click');
        element.click();
        // 補足:onclickは、アンカーにかぎらずあるため、elementをクリックする
      }
      ret = this.postNavigate();
    }
    
    this.finish(ret);
    return ret;
  };
  
  /**
   * 要素をクリックする
   * 補足:ヘッダ偽装未対応
   * @param {(string|Element)} elements - 要素
   * @param {number} [number=0] - 要素番号
   * @return {boolean} 成否
   */
  _this.prototype.doClick = function WebBrowser_doClick(elements, number) {
    number = number != null ? number: 0;
    var ret = false;
    
    this.init();
    
    elements = query2elements(this.ie.document, elements);
    if (elements.length != 0) {
      if (number < 0) {
        number = _random(0, elements.length-1);
      }
      var element = elements[number%elements.length];
      this.addCacheEnd(number%elements.length);
      
      // 親アンカーのターゲットを削除(_blank対策)
      var a = _getParentElement(element, 'a');
      if (a !== null) {
        a.target = '';
      }
      
      element.click();
      ret = this.postNavigate();
    }
    
    this.finish(ret);
    return ret;
  };
  
  
  /**
   * formにパラメータを入力する
   * @param {(string|Element)} form - 要素
   * @param {Object} params - パラメータ
   */
  function WebBrowser_setFormParams(form, params) {
    if (form !== null) {
      // パラメータ設定
      if (params != null) {
        for (var key in params) {
          if (params.hasOwnProperty(key)) {
            form[key].value = params[key];
          }
        }
        WScript.Sleep(500);  // 安全のため、待機(テスト時の目視もしやすい)
      }
    }
  };
  
  /**
   * formに入力してからsubmitする
   * @param {(string|Element)} form - 要素
   * @param {Object} params - パラメータ
   * @param {(string|Element)} [submit=null] - 確定要素(formからのクエリーでも可)
   * @return {boolean} 成否
   */
  _this.prototype.submit = 
  _this.prototype.doSubmit = function WebBrowser_doSubmit(form, params, submit) {
    var ret = false;
    
    this.init();
    
    form = query2element(this.ie.document, form);
    if (form !== null) {
      // パラメータ設定
      WebBrowser_setFormParams(form, params);
      
      // サブミット
      submit = query2element(form, submit);
      if (submit != null) {                     // submit指定あり時
        this.addCacheEnd('arg');
        submit.submit();
      } else if (_isElement(form.submit)) {     // submitが要素の時
        this.addCacheEnd('submit');
        form.submit.click();
        // name='submit'の場合、form.submit()できないため、代替処置
      } else {
        this.addCacheEnd('form');
        form.submit();
      }
      ret = this.postNavigate();
    }
    
    this.finish(ret);
    return ret;
    // 補足:nameがsubmitの要素がある場合、form.submit()が上書きされているため、注意
  };
  _this.prototype.doSubmitNotParamsLog = function WebBrowser_doSubmitNotParamsLog(form, params, submit) {
    var form2 = query2element(this.ie.document, form);
    if (form2 !== null) {
      // パラメータ設定
      WebBrowser_setFormParams(form2, params);
    }
    return this.doSubmit(form, null, submit);
    // パラメータをログ出力しないようにする
    // パスワード対策
  };
  
  /**
   * querySelectorAll
   * @param {string} query - クエリ
   * @param {Element} [element=document] - 親要素
   * @return {Element[]} 要素配列
   */
  _this.prototype.querySelectorAll = function WebBrowser_querySelectorAll(query, element) {
    if (this.setup.cache && this.setup.cache.querySelectorAll === true) {
      this.init();
    }
    
    var elements = (element)? 
                      element.querySelectorAll(query):
                      this.ie.document.querySelectorAll(query);
    
    if (this.setup.cache && this.setup.cache.querySelectorAll === true) {
      this.finish(elements.length, false);
    }
    return elements;
  };
  
  _this.prototype.querySelector = function WebBrowser_querySelector(query, element) {
    var elements = this.querySelectorAll(query, element);
    return (elements.length) ? elements[0]: null;
  };
  
  return _this;
});
