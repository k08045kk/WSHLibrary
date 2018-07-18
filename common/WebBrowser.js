/*!
 * WebBrowser.js v1
 *
 * Copyright (c) 2018 toshi
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */

/**
 * WSH(JScript)用ブラウザ操作補助
 * @requires    module:ActiveXObject("InternetExplorer.Application")
 * @requires    module:JSON
 * @requires    module:WScript
 * @requires    ErrorUtility.js
 * @requires    Console.js
 * @requires    Console.Animation.js
 * @auther      toshi(https://www.bugbugnow.net/)
 * @license     MIT License
 * @version     1
 * @see         1 - add - 初版
 */
(function(root, factory) {
  if (!root.WebBrowser) {
    root.WebBrowser = factory(root.ErrorUtility, root.Console);
  }
})(this, function(ErrorUtility, Console) {
  "use strict";
  
  var _this = void 0;
  
  /**
   * PrivateUnderscore.js
   * @version   4
   */
  {
    function _isString(obj) {
      return Object.prototype.toString.call(obj) === '[object String]';
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
    function _getParentElement(element, tag) {
      var callee = _getParentElement;
      if (element === null) {                     // 要素がない時
      } else if (element.tagName === null) {      // 要素がない時
      } else if (element.tagName.toLowerCase() === tag.toLowerCase()) {  // タグ一致
        return element;                           // 要素
      } else {
        return callee(element.parentElement, tag);// 再帰呼び出し(親要素)
      }
      return null;
    }
  }
  
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
   * @param {string} tag - 探すタグ
   * @return {Element} 親要素
   */
  _this.getParentElement = 
  _this.prototype.getParentElement = function WebBrowser_getParentElement(element, tag) {
    return _getParentElement(element, tag);
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
          waiting.prefix = 'WebBrowser.wait()';
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
    //      4が必須の場合、明示的に再度wait(4)しなおすこと。
  };
  
  /**
   * 関数実行
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {Function} func - 実行する関数
   */
  _this.js = function WebBrowser_js(ie, func) {
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
  _this.css = function WebBrowser_css(ie, text) {
    _this.wait(ie);
    
    var d = ie.document;
    var e = d.createElement('style');
    e.type = 'text/css';
    e.styleSheet.cssText = text;
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
        debug:false, visible:true, top:-1, left:-1, width:-1, height:-1,
        headers:void 0, confirm:true, prompt:'""', 
        'random.sleep.min':5*1000, 'random.sleep.max':10*1000
    }, true);
    this.setup.debug   = _Process_getNamedArgument('wb.debug', this.setup.debug);
    this.setup.visible = _Process_getNamedArgument('wb.visible', this.setup.visible);
    this.setup.top     = _Process_getNamedArgument('wb.top', this.setup.top, 0);
    this.setup.left    = _Process_getNamedArgument('wb.left', this.setup.left, 0);
    this.setup.width   = _Process_getNamedArgument('wb.width', this.setup.width, 1);
    this.setup.height  = _Process_getNamedArgument('wb.height', this.setup.height, 1);
    
    this.clearCache();
  };
  
  /**
   * デストラクタ
   * IEを終了する(必須ではない)。
   * 本関数を呼び出さないまま、プログラムを終了すると、IEが起動したままとなる。
   * @destructor
   */
  _this.prototype.quit = 
  _this.prototype.Quit = function WebBrowser_Quit() {
    if (this.ie !== null) {
      // 終了しない方法も必要？(必要になってから考える)
      this.ie.Quit();
      this.ie = null;
      this.url = '';
    }
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
    if (this.setup.cache !== false) {
      this.cache.push(catchArray);
    }
    if (this.console && this.setup.debug === true) {
      for (var i=0; i<catchArray.length; i++) {
        if (i != 0) {  this.console.print(', ');  }
        this.console.print(catchArray[i]);
      }
    }
  };
  _this.prototype.addCacheEnd = function WebBrowser_addCacheEnd(cacheData, var_args) {
    if (this.setup.cache !== false) {
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
          this.console.sleep(30*1000);
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
    this.url = this.ie.LocationURL;
    
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
    
    // ランダムスリープ(人間ぽさを演出するため)
    if (ret || this.url != this.ie.LocationURL) {
      var min = this.setup['random.sleep.min'];
      var max = this.setup['random.sleep.max'];
      if (sleep !== false && min && max) {
        if (this.console && this.console.sleep && _this.animation) {
          this.console.sleep(_random(min, max));
        } else {
          WScript.Sleep(_random(min, max));
        }
      }
    }
    this.url = this.ie.LocationURL;
    
    // キャッシュ(結果保存)
    this.addCacheEnd(ret, this.getURL());
    if (this.console && this.setup.debug === true) {
      this.console.println();
    }
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
    this.setSetupValue('random.sleep.min', min);
    this.setSetupValue('random.sleep.max', max);
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
  
  _this.prototype.postNavigate = function WebBrowser_postNavigate() {
    return _this.postNavigate(this.ie, this.setup.confirm, this.setup.prompt);
  };
  
  /**
   * 閲覧
   * @param {string} url - URL
   * @param {number} flags - ウィンドウ詳細設定
   * @param {string} targetFrameName - ウィンドウ表示設定
   * @param {string} postData - POSTデータ設定
   * @param {string} headers - HTTPヘッダー
   * @return {boolean} 成否
   */
  _this.prototype.Navigate = 
  function WebBrowser_Navigate(url, flags, targetFrameName, postData, headers) {
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
  
  /**
   * アンカー(href)を閲覧
   * targetをクリアする。(_blankがあると別ウィンドウを表示するため)
   * @param {(string|Element)} element - 要素
   * @return {boolean} 成否
   */
  _this.prototype.doAnchor = function WebBrowser_doAnchor(element) {
    var ret = false;
    
    this.init();
    
    element = query2element(this.ie.document, element);
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
    this.finish(ret);
    return ret;
  };
  
  /**
   * 要素をクリック
   * 補足:ヘッダ偽装未対応
   * @param {(string|Element)} element - 要素
   * @return {boolean} 成否
   */
  _this.prototype.doClick = function WebBrowser_doClick(element) {
    var ret = false;
    
    this.init();
    
    element = query2element(this.ie.document, element);
    if (element !== null) {
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
      if (params != null) {
        for (var key in params) {
          if (params.hasOwnProperty(key)) {
            form[key].value = params[key];
          }
        }
        WScript.Sleep(500);  // 安全のため、待機(テスト時の目視もしやすい)
      }
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
  
  /**
   * querySelectorAll
   * @param {string} query - クエリ
   * @param {Element} [element=document] - 親要素
   * @return {Element[]} 要素配列
   */
  _this.prototype.querySelectorAll = function WebBrowser_querySelectorAll(query, element) {
    this.init();
    
    var elements = (element)? 
                      element.querySelectorAll(query):
                      this.ie.document.querySelectorAll(query);
    
    this.finish(elements.length, false);
    return elements;
  };
  
  _this.prototype.querySelector = function WebBrowser_querySelector(query, element) {
    var elements = this.querySelectorAll(query, element);
    return (elements.length) ? elements[0]: null;
  };
  
  return _this;
});
