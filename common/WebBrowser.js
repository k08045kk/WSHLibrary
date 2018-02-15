// require("../common/Utility.js");
// require("../common/WebUtility.js");
// require("../common/WebApp.js");

/**
 * Web関連
 * @constructor
 * @auther toshi2limu@gmail.com (toshi)
 */
(function(factory) {
	global.WebBrowser = factory();
})(function WebBrowser_factory() {
	"use strict";
	
	var _this = void 0;
	var _super= Object.prototype;
	
	_this = function WebBrowser_constructor() {
		this.initialize.apply(this, arguments);
	};
	_this.prototype = Object.create(_super);
	
	_this.READYSTATE_UNINITIALIZED	= 0; 	// ReadyState: 未完了状態
	_this.READYSTATE_LOADING		= 1; 	// ReadyState: ロード中状態
	_this.READYSTATE_LOADED			= 2;	// ReadyState: ロード完了状態、ただし操作不可能状態
	_this.READYSTATE_INTERACTIVE	= 3;	// ReadyState: 操作可能状態
	_this.READYSTATE_COMPLETE		= 4;	// ReadyState: 全データ読み込み完了状態
	
	/**
	 * 親ノードのtagを探す。
	 * @public
	 * @param {Element}		要素
	 * @return {Element}	アンカータグ
	 */
	_this.getParentElement = function WebBrowser_getParentElement(element, tag) {
		if (element === null) {												// 要素がない時
		} else if (element.tagName === null) {								// 要素がない時
		} else if (element.tagName.toLowerCase() === tag.toLowerCase()) {	// タグ一致
			return element;													// 要素
		} else {
			return arguments.callee(element.parentElement, tag);			// 再帰呼び出し(親要素)
		}
		return null;
	}
	_this.getParentAnchor = function WebBrowser_getParentAnchor(element) {
		return _this.getParentElement(element, 'a');
	};
	
	/**
	 * IEの読み込みを待機する。
	 * @public
	 * @param {ActiveXObject} IEApp	InternetExplorer.Application
	 * @param {number} state 		状態(enum.ReadyState)
	 */
	_this.wait = function WebBrowser_wait(ie, state, timeout) {
		state	= (state   !== void 0)? state: _this.READYSTATE_INTERACTIVE;	// 操作可能状態
		timeout = (timeout !== void 0)? timeout: 60*1000;
		
		if (ie.Busy || (ie.readystate < state)) {	// 重複呼び出し対策
			var begin = Date.now();
			for (var i=0; i<2; i++) {										// 安全のため、2回実行
				// busyの場合、あるいは、読み込み中の場合は、100ミリ秒スリープする
				WScript.Sleep(100);
				while (ie.Busy || (ie.readystate < state)) {
					WScript.Sleep(100);
					if (begin+timeout < Date.now()) {		// タイムアウト時間経過
						throw new ExError("time out. (Busy="+ie.Busy
													+", readystate="+ie.readystate+")");
					}
//					println("wait: "+(begin + 60*1000)+", "+Date.now());
				}
			}
			WScript.Sleep(300);	// 安全のため、最短500msを確保
		}
		// 補足:デフォをREADYSTATE_INTERACTIVE(3)とする
		//		完全な完了は、4であるが、4にならないページ等もあり、(楽天など)
		//		3でも問題が起こらないことも多いため、デフォは3とする
		//		4が必須の場合、明示的に再度wait(4)しなおすこと
	};
	
	/**
	 * 関数実行
	 * @public
	 * @param {ActiveXObject} IEApp	InternetExplorer.Application
	 * @param {Function} func 		実行する関数
	 * @return {boolean}			成否
	 */
	_this.js = function WebBrowser_js(ie, func) {
		_this.wait(ie);							// 安全のため、待機
		
		// setTimeout実行
//		var script = "javascript:(" 
//			+ func.toString() 
//			+ ")();"
//		;
//		ie.document.Script.setTimeout(script, 0);	// 実行(0ms後)
//		WScript.Sleep(10);							// 実行待ち
		// 補足:Sleepなしの場合、即時実行ではないため、
		//		タイミングにより、直後のクリックイベント等が先に実行されることがある
		//		そのため、confirmの置き換えが間に合わない可能性がある
		// 補足:Sleepしたとしても、確実にsetTimeoutを実行するわけではない
		
		// 動的実行
		var d = ie.document;
		var e = d.createElement("script");
		e.type = "text/javascript";
		e.text = "("+func.toString()+")();";
		if (d.head) {		    d.head.appendChild(e);	}
		else if (d.body) {	d.body.appendChild(e);	}
		else {				throw new ExError("append failed. (head="+d.head+", body="+d.body+")");	}
//		d.documentElement.firstChild.appendChild(e);
		// 補足:setTimeoutと異なり、HTML上に書き込むため、
		//		他のスクリプトに干渉する可能性がある
		
		_this.wait(ie);		// 安全のため、待機
		// 補足:ブックマークレット(Navigateのurlとして指定)とすると、
	 	//		wait()で読み込みが完了しないため、辞めた方がいい。
		// 補足:
	};
	
	/**
	 * CSS追加
	 * @public
	 * @param {ActiveXObject} IEApp	InternetExplorer.Application
	 * @param {Function} func 		実行する関数
	 * @return {boolean}			成否
	 */
	_this.css = function WebBrowser_css(ie, text) {
		_this.wait(ie);
		
		var d = ie.document;
		var e = d.createElement('style');
		e.type = "text/css";
		e.styleSheet.cssText = text;
		if (d.head) {		    d.head.appendChild(e);	}
		else if (d.body) {	d.body.appendChild(e);	}
		else {				throw new ExError("append failed. (head="+d.head+", body="+d.body+")");	}
//		d.documentElement.firstChild.appendChild(e);
		
		_this.wait(ie);
	};
	
	/**
	 * ページ移動後の共通処理
	 * @public
	 * @param {ActiveXObject} IEApp	InternetExplorer.Application
	 */
	_this.postNavigate = function WebBrowser_postNavigate(ie, confirm, prompt) {
		var ret = false;
		try {
			// 自動化の妨げとなる関数を上書き
			var func =	  'window.alert = function (message) {console.log("alert(): "+message);};'
						+ 'window.onunload = function () {console.log("onunload()");};'
						+ 'window.onbeforeunload = function () {console.log("onbeforeunload()");};';
			if (confirm === void 0) { confirm = true;	}	// デフォでOKとする
			if (confirm !== null) {							// 標準指定(null)でない時
				func +=   'window.confirm = function () {'
						+ '  console.log("confirm('+confirm+')");'
						+ '  return('+confirm+');'
						+ '};';
			}
			if (prompt === void 0) { prompt = '""';	}		// デフォで空文字とする
			if (prompt !== null) {							// 標準指定(null)でない時
				func +=   'window.prompt = function () {'
						+ '  console.log(\'prompt('+prompt+')\');'	// ダブルクオートの2重対応
						+ '  return('+prompt+');'
						+ '};';
			}
			_this.js(ie, new Function(func));
			// 補足:onload以前は、対象外(Seleniumでも無理っぽいので諦めたほうが良さそう)
			
			_this.wait(ie);	// 安全のため、待機
			ret = true;
		} catch (e) {
			if (e.fileName !== 'WebBrowser.js') {	// ファイル外のエラーの時
				throw e;
			}
		}
		return ret;
	};
	
	/**
	 * コンストラクタ
	 * @public
	 * @constructor
	 */
	_this.prototype.initialize = function WebBrowser_initialize() {
		this.ie = null;
		this.setup = {};
		this.setup.visible	= WSUtility.getNamedArgument("visible", true);
//		this.setup.top		= WSUtility.getNamedArgument("top",		0, 0);
//		this.setup.left		= WSUtility.getNamedArgument("left",	0, 0);
//		this.setup.width	= WSUtility.getNamedArgument("width",	0, 0);
//		this.setup.height	= WSUtility.getNamedArgument("height",	0, 0);
		
		this.setup.headers	= void 0;
		this.setup.confirm	= true;
		this.setup.prompt	= '""';
		this.clearCache();
	};
	
	/**
	 * デストラクタ
	 * @public
	 * @destructor
	 */
	_this.prototype.quit = 
	_this.prototype.Quit = function WebBrowser_Quit() {
		if (this.ie !== null) {
			// 終了しない方法も必要？(必要になってから考える)
			this.ie.Quit();
			this.ie = null;
			this.url = '';
			this.clearCache();
		}
	};
	_this.prototype.debugMode = function WebBrowser_debugMode(debug) {
		this.debug = (debug !== false);
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
		this.cache.push(catchArray);
		if (this.debug === true) {
			for (var i=0; i<catchArray.length; i++) {
				if (i != 0) {	console.print(', ');	}
				console.print(catchArray[i]);
			}
		}
	};
	_this.prototype.addCacheEnd = function WebBrowser_addCacheEnd(cacheData, var_args) {
		Array.prototype.push.apply(this.cache[this.cache.length-1], arguments);
		if (this.debug === true) {
			for (var i=0; i<arguments.length; i++) {
				console.println(', ');
				console.print(arguments[i]);
			}
		}
	};
	_this.prototype.addCacheSeparator = function WebBrowser_addCacheSeparator(message) {
		var carcheArray = [Atom.trace(new ExError().stackframes[1], true)];
		if (message) {	carcheArray.push(message);	}
		carcheArray.push(this.getURL());
		this.addCache.apply(this, carcheArray);
		if (this.debug === true) {
			console.println();
		}
	};
	
	_this.prototype.getCacheData = function WebBrowser_getCacheData() {
		var msgs = [];
		for (var i=0; i<this.cache.length; i++) {
			msgs.push(this.cache[i].join(", "));
		}
		return msgs.join("\n");
	};
	
	/**
	 * 前処理
	 * @public
	 * @return {ActiveXObject}		InternetExplorer.Application
	 */
	_this.prototype.init = function WebBrowser_init() {
		if (this.ie != null) {
			try {
				this.ie.document.querySelector('body');
			} catch (e) {
println('not body.');
				try {
					this.ie.Quit();
				} catch (e2) {}
				this.ie = null;
			}
		}
		if (this.ie === null) {
			for (var i=1; i>=0; i--) {	// 2回作成を試みる
				try {
					this.ie = new ActiveXObject("InternetExplorer.Application");
					break;
				} catch (e) {
					if (i === 0) {	throw Atom.captureStackTrace(e);	}
				}
				WSUtility.Sleep(30*1000);
			}
			
			// IE初期化
//			if (this.setup.top !== 0) {		this.ie.Top		= this.setup.top;		}
//			if (this.setup.left !== 0) {	this.ie.Left	= this.setup.left;		}
//			if (this.setup.width !== 0) {	this.ie.Width	= this.setup.width;		}
//			if (this.setup.height !== 0) {	this.ie.Height	= this.setup.height;	}
			this.ie.Navigate("about:blank");	// 空ページ表示
			this.ie.Visible = this.setup.visible;
		}
		_this.wait(this.ie);
		this.url = this.ie.LocationURL;
		
		// キャッシュ(コマンド保存)
		this.addCache(Atom.trace(new ExError().stackframes[1], true));
		// 補足:読み込み完了状態で処理に進む
		//		読み込み中に、以降の処理を行うとエラーとなるため
	};
	
	/**
	 * 後処理
	 * @public
	 * @return {ActiveXObject}		InternetExplorer.Application
	 */
	_this.prototype.finish = function WebBrowser_finish(ret, sleep) {
		_this.wait(this.ie);
		
		// ランダムスリープ(人間ぽさを演出するため)
		if (ret || this.url != this.ie.LocationURL) {								// 成功 or ページ移動
			if (sleep !== false && this.setup.rsmin && this.setup.rsmax) {
				WSUtility.Sleep(Atom.random(this.setup.rsmin, this.setup.rsmax));	// CUIありスリープ
			}
		}
		this.url = this.ie.LocationURL;
		
		// キャッシュ(結果保存)
		this.addCacheEnd(ret, this.getURL());
		if (this.debug === true) {
			console.println();
		}
	};
	
	/**
	 * @public
	 * @destructor
	 */
	_this.prototype.getIE = function WebBrowser_getIE() {
		return this.ie;
	};
	
	_this.prototype.getURL = function WebBrowser_getURL() {
		return (this.ie !== null)? this.ie.LocationURL: "";
	};
	
	_this.prototype.getDocument = function WebBrowser_getDocument() {
		return (this.ie !== null)? this.ie.Document: null;
	};
	
	_this.prototype.setHeaders = function WebBrowser_setHeaders(header, var_args) {
		this.addCacheSeparator();
		this.setup.headers = '';
		for (var i=0; i<arguments.length; i++) {
			this.setup.headers += arguments[i]+"\n";
		}
	};
	
	_this.prototype.setConfirmAnswer = function WebBrowser_setConfirmAnswer(answer) {
		this.addCacheSeparator();
		this.setup.confirm = answer;
	};
	
	_this.prototype.setPromptAnswer = function WebBrowser_setPromptAnswer(answer) {
		this.addCacheSeparator();
		this.setup.prompt = JSON.stringify(answer);
	};
	
	_this.prototype.setRandomSleep = function WebBrowser_setRandomSleep(min, max) {
		this.addCacheSeparator();
		this.setup.rsmin = min;
		this.setup.rsmax = max;
	};
	
	_this.prototype.wait = function WebBrowser_prototype_wait(state, timeout) {
		_this.wait(this.ie, state, timeout);
	};
	
	_this.prototype.postNavigate = function WebBrowser_prototype_postNavigate() {
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
		if (Atom.isString(query)) {							// 文字列(クエリー)
			return element.querySelector(query);
		} else if (Atom.isElement(query)) {					// 要素
			return query;
		}
		return null;
	};
	
	/**
	 * アンカー(href)を閲覧
	 * targetをクリアする。(_blankがあると別ウィンドウを表示するため)
	 * @public
	 * @param {ActiveXObject} IEApp		InternetExplorer.Application
	 * @param {Element} element 		要素
	 * @return {boolean}				成否
	 */
	_this.prototype.doAnchor = function WebBrowser_doAnchor(element) {
		var ret	= false;
		
		this.init();
		
		element = query2element(this.ie.document, element);
		var a	= _this.getParentAnchor(element);
		if (a !== null) {
			// 親アンカーのターゲットを削除(_blank対策)
			a.target = "";
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
	 * @public
	 * @param {Element} element 		要素
	 * @return {boolean}				成否
	 */
	_this.prototype.doClick = function WebBrowser_doClick(element) {
		var ret = false;
		
		this.init();
		
		element = query2element(this.ie.document, element);
		if (element !== null) {
			// 親アンカーのターゲットを削除(_blank対策)
			var a = _this.getParentAnchor(element);
			if (a !== null) {
				a.target = "";
			}
			
			element.click();
//			var ev = this.ie.Document.createEvent("MouseEvents");	// イベントオブジェクトを作成
//			ev.initEvent("click", false, true);						// イベントの内容を設定
//			element.dispatchEvent(ev);								// イベントを発火させる
			
			ret = this.postNavigate();
		}
		
		this.finish(ret);
		return ret;
	};
	
	/**
	 * formに入力してからsubmitする
	 * @public
	 * @param {Element} form	 		要素
	 * @param {Object} params 			パラメータ
	 * @param {Element} submit 			確定要素(formからのクエリーでも可)
	 * @return {boolean}				成否
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
				WScript.Sleep(500);	// 安全のため、待機(テスト時の目視もしやすい)
			}
			// サブミット
			submit = query2element(form, submit);
			if (submit != null) {											// submit指定あり時
				this.addCacheEnd("arg");
				submit.submit();
			} else if (Atom.isElement(form.submit)) {	// submitが要素の時
				this.addCacheEnd("submit");
				form.submit.click();
				// name="submit"の場合、form.submit()できないため、代替処置
			} else {
				this.addCacheEnd("form");
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
	 * 指定したURLを巡回する。
	 * @deprecated 廃止の可能性あり
	 * @param {ActiveXObject} IEApp	InternetExplorer.Application
	 * @param {Array} urls 			URLの配列
	 * @param {string} au	 		ユーザエージェント
	 * @param {Function} func 		コールバック関数
	 */
	_this.prototype.autoPatroler = function (urls, ua) {
		for (var i=0; i<urls.length; i++) {
			try {
				this.Navigate(urls[i]);
			} catch (e) {
				console.printStackTrace(e);
				this.Quit();
				WScript.Sleep(10*1000);
			}
		}
	};
	
	/**
	 * 指定したURLを巡回して、指定の要素をクリックする。
	 * @deprecated 廃止の可能性あり
	 * @param {ActiveXObject} IEApp	InternetExplorer.Application
	 * @param {Array} urls 			URLの配列
	 * @param {string} quary 		クエリー
	 */
	_this.prototype.autoClicker = function (urls, quary) {
		for (var i=0; i<urls.length; i++) {
			try {
				ret = this.Navigate(urls[i])
					 && this.doClick(quary)
					 && true;
			} catch (e) {
				console.printStackTrace(e);
				this.Quit();
				WScript.Sleep(10*1000);
			}
		}
	};
	
	return _this;
});
