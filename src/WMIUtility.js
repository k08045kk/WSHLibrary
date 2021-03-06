/*!
 * WMIUtility.js v1
 *
 * Copyright (c) 2018 toshi (https://github.com/k08045kk)
 *
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */

/**
 * WSH(JScript)用WMIライブラリ
 * WMI(Windows Management Instrumentation)
 * @requires    module:ActiveXObject("WbemScripting.SWbemLocator")
 * @auther      toshi (https://github.com/k08045kk)
 * @version     1
 * @see         1.20180702 - add - 初版
 */
(function(global, factory) {
  if (!global.WMIUtility) {
    global.WMIUtility = factory();
  }
})(this, function() {
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
    function _isArray(obj) {
      return Object.prototype.toString.call(obj) === '[object Array]';
    };
  }
  
  /**
   * コンストラクタ
   */
  _this = function WMIUtility_constructor() {};
  
  _this.locator = new ActiveXObject("WbemScripting.SWbemLocator");
  _this.service = _this.locator.ConnectServer();
  
  
  /**
   * 終了化
   * 終了化後、クラス内の変数/関数にアクセスしないこと。
   */
  _this.quit = function WMIUtility_quit() {
    _this.service = null;
    _this.locator = null;
  };
  
  /**
   * 選択
   * 配列の最初の要素を返す。
   * @param {string} query - クエリ
   * @return {wmiObject} アイテム
   */
  _this.select = function WMIUtility_select(query) {
    var items = _this.selects(query);
    return (items.length != 0)? items[0]: null;
  };
  
  /**
   * 選択
   * @param {string} query - クエリ
   * @return {wmiObject[]} アイテム配列
   */
  _this.selects = function WMIUtility_selects(query) {
    var ret = [];
    var set = _this.service.ExecQuery(query);
    for (var e=new Enumerator(set); !e.atEnd(); e.moveNext()) {
      ret.push(e.item());
    }
    set = null;
    return ret;
  };
  
  /**
   * プロパティ一覧
   * @param {(string|wmiObject)} query - クエリ or アイテム
   * @return {Object} アイテムプロパティのオブジェクト({Name: Value, ...})
   */
  _this.getProperty = function WMIUtility_getProperty(query) {
    var ret = null;
    
    if (_isString(query)) {      // クエリを処理する
      var set = _this.service.ExecQuery(query);
      for (var e=new Enumerator(set); !e.atEnd(); e.moveNext()) {
        query = e.item();
        break;
      }
      set = null;
    }
    if (query != null) {        // アイテムを処理する
      ret = {};
      for (var e=new Enumerator(query.Properties_); !e.atEnd(); e.moveNext()) {
        var item = e.item();
        ret[item.Name] = item.Value;
      }
    }
    return ret;
  }
  
  /**
   * プロパティ一覧
   * @param {(string|wmiObject[])} query - クエリ or アイテム配列
   * @return {Object[]} アイテムプロパティのオブジェクト([{Name: Value, ...}, ...])
   */
  _this.getProperties = function WMIUtility_getProperties(query) {
    var ret = null;
    
    if (_isString(query)) {      // クエリを処理する
      query = _this.selects(query);
    }
    if (_isArray(query)) {       // アイテム配列を処理する
      ret = [];
      for (var i=0; i<query.length; i++) {
        ret.push(_this.getProperty(query[i]));
      }
    }
    return ret;
  };
  
  return _this;
});
