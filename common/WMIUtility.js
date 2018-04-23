/**
 * WSH(JScript)用WMIライブラリ
 * WMI(Windows Management Instrumentation)
 * @requires    module:ActiveXObject("WbemScripting.SWbemLocator")
 * @auther      toshi(https://www.bugbugnow.net/)
 * @license     MIT License
 * @version     1
 */
(function(global, factory) {
  if (!global.WMIUtility) {
    global.WMIUtility = factory();
  }
})(this, function() {
  "use strict";
  
  var _this = void 0;
  var _locator = new ActiveXObject("WbemScripting.SWbemLocator");
  var _service = _locator.ConnectServer();
  
  _this = function WMIUtility_constructor() {};
  
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
    var set = _service.ExecQuery(query);
    
    for (var e=new Enumerator(set); !e.atEnd(); e.moveNext()) {
      ret.push(e.item());
    }
    set = null;
    return ret;
  };
  
  /**
   * プロパティ一覧
   * @param {wmiObject} obj - アイテム
   * @return {Object} アイテムプロパティのオブジェクト({Name: Value, ...})
   */
  _this.getProperties = function WMIUtility_getProperties(obj) {
    var ret = {};
    if (obj != null) {
      for (var e=new Enumerator(obj.Properties_); !e.atEnd(); e.moveNext()) {
        var item = e.item();
        ret[item.Name] = item.Value;
      }
    }
    item = null;
    return ret;
  };
  
  /**
   * プロパティ一覧
   * @param {string} query - クエリ
   * @return {Object} アイテムプロパティのオブジェクト({Name: Value, ...})
   */
  _this.getQueryProperty = function WMIUtility_getQueryProperty(query) {
    var ret = null;
    var set = _service.ExecQuery(query);
    
    for (var e=new Enumerator(set); !e.atEnd(); e.moveNext()) {
      ret = _this.getProperties(e.item());
      break;
    }
    set = null;
    return ret;
  }
  
  /**
   * プロパティ一覧
   * @param {string} query - クエリ
   * @return {Array} アイテムプロパティのオブジェクト([{Name: Value, ...}, ...])
   */
  _this.getQueryProperties = function WMIUtility_getQueryProperties(query) {
    var ret = [];
    var set = _service.ExecQuery(query);
    
    for (var e=new Enumerator(set); !e.atEnd(); e.moveNext()) {
      ret.push(_this.getProperties(e.item()));
    }
    set = null;
    return ret;
  };
  
  return _this;
});
