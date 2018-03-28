/**
 * Web関連
 * @requires    module:WScript
 * @requires    module:ActiveXObject("InternetExplorer.Application")
 * @requires    module:ErrorUtility.js
 * @auther      toshi(https://www.bugbugnow.net/)
 * @license     MIT License
 * @version     1
 */
 // TODO: キャッシュの切り離し
(function(global, factory) {
  if (!global.WebBrowser) {
    global.WebBrowser = factory(global.ErrorUtility);
  }
})(this, function(ErrorUtility) {
  "use strict";
  
  var _this = void 0;
  
  /**
   * PrivateUnderscore.js
   * @version   1
   */
  {
    function _isString(obj) {
      return Object.prototype.toString.call(obj) === '[object String]';
    };
    function _isElement(obj) {
      return !!(obj && obj.nodeType === 1);
    };
    function _random(min, max) {
      return min + Math.floor(Math.random() * (max - min + 1));
    };
    function _Process_getNamedArgument(name, def, min, max) {
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
    /**
     * 親ノードのtagを探す。
     * @param {Element} element - 要素
     * @return {string} tag - 探すタグ
     */
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
    function _getParentAnchor(element) {
      return _getParentElement(element, 'a');
    };
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
   * @deprecated 変更の可能性大
   */
  _this.IE11    = 'Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; Touch; rv:11.0) like Gecko';
  _this.iPhone  = 'User-Agent: Mozilla/5.0 (iPhone; CPU iPhone OS 9_0_1 like Mac OS X) AppleWebKit/601.1.46 (KHTML, like Gecko) Version/9.0 Mobile/13A404 Safari/601.1';	// iPhone(iOS 9.0.1)
  _this.UserAgent = _this.IE11;
  
  /**
   * IEの読み込みを待機
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {number} [opt_state=3] - 状態(READYSTATE参照)
   * @param {number} [opt_timeout=60000] - 待機時間(ms)
   */
  _this.wait = function WebBrowser_static_wait(ie, opt_state, opt_timeout) {
    var state = (opt_state != null)? opt_state: _this.READYSTATE_INTERACTIVE;
    var timeout = (opt_timeout != null)? opt_timeout: 60*1000;
    
    if (ie.Busy || (ie.readystate < state)) {   // 重複呼び出し対策
      var begin = Date.now();
      var waiting = Console.getAnimation('WebBrowser.wait');
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
      
      waiting.start();
      for (var i=0; i<2; i++) {                 // 安全のため、2回実行
        // busyの場合、あるいは、読み込み中の場合は、100ミリ秒スリープする
        while (ie.Busy || (ie.readystate < state)) {
          WScript.Sleep(100);
          if (begin+timeout < Date.now()) {
            // タイムアウト時間経過
            var message = 'time out. (Busy='+ie.Busy+', readystate='+ie.readystate+')';
            throw ErrorUtility.create(message, null, 'WebBrowserError');
          }
        }
      }
      waiting.stop();
      
      // 安全のため、最短300msを確保
      WScript.Sleep(300);
    }
    // 補足:デフォをREADYSTATE_INTERACTIVE(3)とする
    //      完全な完了は、4であるが、4にならないページ等もあり、(楽天など)
    //      3でも問題が起こらないことも多いため、デフォは3とする。
    //      4が必須の場合、明示的に再度wait(4)しなおすこと
  };
  
  /**
   * 関数実行
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {Function} func - 実行する関数
   * @return {boolean} 成否
   */
  _this.js = function WebBrowser_js(ie, func) {
    _this.wait(ie);              // 安全のため、待機
    
    // setTimeout実行
//    var script = 'javascript:(' 
//      + func.toString() 
//      + ')();'
//    ;
//    ie.document.Script.setTimeout(script, 0);  // 実行(0ms後)
//    WScript.Sleep(10);              // 実行待ち
    // 補足:Sleepなしの場合、即時実行ではないため、
    //      タイミングにより、直後のクリックイベント等が先に実行されることがある
    //      そのため、confirmの置き換えが間に合わない可能性がある
    //      Sleepしたとしても、確実にsetTimeoutを実行するわけではない
    
    // 動的実行
    var d = ie.document;
    var e = d.createElement('script');
    e.type = 'text/javascript';
    e.text = '('+func.toString()+')();';
    if (d.head) {        d.head.appendChild(e);  }
    else if (d.body) {  d.body.appendChild(e);  }
    else {  throw ErrorUtility.create('append failed. (head='+d.head+', body='+d.body+')');  }
    // 補足:setTimeoutと異なり、HTML上に書き込むため、
    //      他のスクリプトに干渉する可能性がある
    
    _this.wait(ie);    // 安全のため、待機
    // 補足:ブックマークレット(Navigateのurlとして指定)とすると、
     //      wait()で読み込みが完了しないため、辞めた方がいい。
  };
  
  /**
   * CSS追加
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {Function} func - 実行する関数
   * @return {boolean} 成否
   */
  _this.css = function WebBrowser_css(ie, text) {
    _this.wait(ie);
    
    var d = ie.document;
    var e = d.createElement('style');
    e.type = 'text/css';
    e.styleSheet.cssText = text;
    if (d.head) {        d.head.appendChild(e);  }
    else if (d.body) {  d.body.appendChild(e);  }
    else {  throw ErrorUtility.create('append failed. (head='+d.head+', body='+d.body+')');  }
    
    _this.wait(ie);
  };
  
  /**
   * ページ移動後の共通処理
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
            + '  console.log("confirm('+confirm+')");'
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
      
      _this.wait(ie);  // 安全のため、待機
      ret = true;
    } catch (e) {
      if (e.fileName !== 'WebBrowser.js') {  // ファイル外のエラーの時
        throw e;
      }
    }
    return ret;
  };
  
  /**
   * コンストラクタ
   * @constructor
   */
  _this.prototype.initialize = function WebBrowser_initialize() {
    this.ie = null;     // IE
    this.list = [];     // 閲覧リスト
    this.console = Console.getConsole();  // 出力コンソール
    this.setup = {};    // IEの設定
    this.setup.debug   = _Process_getNamedArgument('debug', false);
    this.setup.visible = _Process_getNamedArgument('visible', true);
    this.setup.top     = _Process_getNamedArgument('top', -1, 0);
    this.setup.left    = _Process_getNamedArgument('left', -1, 0);
    this.setup.width   = _Process_getNamedArgument('width', -1, 1);
    this.setup.height  = _Process_getNamedArgument('height', -1, 1);
    
    this.setup.headers = void 0;
    this.setup.confirm = true;
    this.setup.prompt  = '""';
    
    this.clearCache();
    this.setRandomSleep(5*1000, 10*1000);       // finish()の追加待機時間(安全重視)
  };
  
  /**
   * デストラクタ
   * IEを終了する(必須ではない)。
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
   * @deprecated 仕様変更の可能性大
   */
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
    if (this.setup.debug === true) {
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
    if (this.setup.debug === true) {
      for (var i=0; i<arguments.length; i++) {
        this.console.println(', ');
        this.console.print(arguments[i]);
      }
    }
  };
  _this.prototype.addCacheSeparator = function WebBrowser_addCacheSeparator(message) {
    var carcheArray = [ErrorUtility.trace(2, true)];
    if (message) {  carcheArray.push(message);  }
    carcheArray.push(this.getURL());
    this.addCache.apply(this, carcheArray);
    if (this.setup.debug === true) {
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
          if (i === 0) {  throw ErrorUtility.captureStackTrace(e);  }
        }
        this.console.sleep(30*1000);
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
    this.addCache(ErrorUtility.trace(2, true));
    // 補足:読み込み完了状態で処理に進む
    //    読み込み中に、以降の処理を行うとエラーとなるため
  };
  
  /**
   * 後処理
   */
  _this.prototype.finish = function WebBrowser_finish(ret, sleep) {
    _this.wait(this.ie);
    
    // ランダムスリープ(人間ぽさを演出するため)
    if (ret || this.url != this.ie.LocationURL) {               // 成功 or ページ移動
      if (sleep !== false && this.setup.rsmin && this.setup.rsmax) {
        this.console.sleep(_random(this.setup.rsmin, this.setup.rsmax));
      }
    }
    this.url = this.ie.LocationURL;
    
    // キャッシュ(結果保存)
    this.addCacheEnd(ret, this.getURL());
    if (this.setup.debug === true) {
      this.console.println();
    }
  };
  
  _this.prototype.getIE = function WebBrowser_getIE() {
    return this.ie;
  };
  
  _this.prototype.getURL = function WebBrowser_getURL() {
    return (this.ie !== null)? this.ie.LocationURL: '';
  };
  
  _this.prototype.getDocument = function WebBrowser_getDocument() {
    return (this.ie !== null)? this.ie.Document: null;
  };
  
  _this.prototype.setSetupValue = function WebBrowser_setSetupValue(name, value) {
    this.addCacheSeparator();
    this.setup[name] = value;
  }
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
    this.setSetupValue('rsmin', min);
    this.setSetupValue('rsmax', max);
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
   * @param {ActiveXObject} ie - InternetExplorer.Application
   * @param {Element} element - 要素
   * @return {boolean} 成否
   */
  _this.prototype.doAnchor = function WebBrowser_doAnchor(element) {
    var ret  = false;
    
    this.init();
    
    element = query2element(this.ie.document, element);
    var a  = _getParentAnchor(element);
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
   * @param {Element} element - 要素
   * @return {boolean} 成否
   */
  _this.prototype.doClick = function WebBrowser_doClick(element) {
    var ret = false;
    
    this.init();
    
    element = query2element(this.ie.document, element);
    if (element !== null) {
      // 親アンカーのターゲットを削除(_blank対策)
      var a = _getParentAnchor(element);
      if (a !== null) {
        a.target = '';
      }
      
      element.click();
//      var ev = this.ie.Document.createEvent('MouseEvents');  // イベントオブジェクトを作成
//      ev.initEvent('click', false, true);       // イベントの内容を設定
//      element.dispatchEvent(ev);                // イベントを発火させる
      
      ret = this.postNavigate();
    }
    
    this.finish(ret);
    return ret;
  };
  
  /**
   * formに入力してからsubmitする
   * @param {Element} form - 要素
   * @param {Object} params - パラメータ
   * @param {Element} submit - 確定要素(formからのクエリーでも可)
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
        var keys = Object.keys(params);
        for (var i=0; i<keys.length; i++) {
          form[keys[i]].value = params[keys[i]];
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
  
  _this.prototype.querySelectorAll = function WebBrowser_querySelectorAll(query) {
    this.init();
    
    var elements = this.ie.document.querySelectorAll(query);
    
    this.finish(elements.length, false);
    return elements;
  };
  
  _this.prototype.querySelector = function WebBrowser_querySelector(query) {
    var elements = this.querySelectorAll(query);
    return (elements.length) ? elements[0] : null;
  };
  
  /**
   * 閲覧リストの要素を閲覧
   * @param {string[]} opt_list - 閲覧リスト(閲覧後削除します)
   * @return {boolean} 成否
   */
  _this.prototype.listNavigate = 
  function WebBrowser_listNavigate(opt_list, flags, targetFrameName, postData, headers) {
    var ret  = true;
    var list = opt_list;
    var i = 0;
    
    this.addCache(ErrorUtility.trace(1, true));
    if (opt_list == null) {
      list = this.list;
    }
    
    for (i=0; i<list.length && ret !== true; i++) {
      ret = this.Navigate(list[i], flags, targetFrameName, postData, headers);
    }
    
    // リストを削除(完了分+失敗分)
    if (0 < i && i <= list.length) {
      list.splice(0, i);
    }
    return ret;
  };
  
  /**
   * 閲覧リストに追加
   * @param {(string|Element)} element - URL or 要素
   */
  _this.prototype.addNavigateList = 
  function WebBrowser_addNavigateList(element) {
    if (_isElement(element)) {
      element = query2element(this.ie.document, element);
      var a  = _getParentAnchor(element);
      if (a && a.href) {
        this.list.push(a.href);
      }
    } else {
      this.list.push(''+element);
    }
  };
  
  return _this;
});
