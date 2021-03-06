/*!
 * DBLite.js v4
 *
 * Copyright (c) 2018 toshi (https://github.com/k08045kk)
 *
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */

/**
 * 簡易データベース
 * SQLiteを使用した簡易データベース、
 * キーバリュー方式で使用できることを目標とする。
 * DBやSQL文を意識しなくても使用できることを目標とする。
 * [SQLite ODBC Driver]のインストールが別途必要です。
 * @requires    module:ActiveXObject('Scripting.FileSystemObject')
 * @requires    module:ActiveXObject('ADODB.Command')
 * @requires    module:ActiveXObject("ADODB.Connection")
 * @requires    module:JSON
 * @requires    Console.js
 * @requires      ErrorUtility.js
 * @auther      toshi (https://github.com/k08045kk)
 * @version     4
 * @see         1.20180611 - add - 初版
 * @see         2.20180723 - fix - パス長が規定値を超えた場合、エラーする
 * @see         3.20190927 - update - keysIterator, valuesIteratorを追加
 * @see         4.20200203 - update - _FileUtility_createFolder()を再帰処理しないように修正
 */
 
(function(root, factory) {
  if (!root.DBLite) {
    root.DBLite = factory(root.JSON, 
                          root.Console);
  }
})(this, function DBLite_factory(JSON, Console) {
  "use strict";
  
  var fs = new ActiveXObject('Scripting.FileSystemObject');
  var _this = void 0;
  
  /**
   * PrivateUnderscore.js
   * @version   7
   */
  {
    function _isObject(obj) {
      var type = typeof obj;
      return type === 'function' || type === 'object' && !!obj;
    };
    function _isArray(obj) {
      return Object.prototype.toString.call(obj) === '[object Array]';
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
  
  _this = function DBLite_constructor() {
    this.initialize.apply(this, arguments);
  };
  
  function createTablesCommand(_this) {         // テーブル一覧取得
    return ["SELECT name FROM sqlite_master WHERE type = 'table';"];
  };
  function createSchemaCommand(_this, table) {  // スキーマ取得
    return ["SELECT sql FROM sqlite_master WHERE type = 'table' AND name = ?;", table];
  };
  function createCreateTableCommand(_this, table) {
    return ["CREATE TABLE IF NOT EXISTS ["+table+"] ("  // (存在しない場合に、)テーブル作成
        + _this.field.key+" TEXT NOT NULL UNIQUE, "     // 文字列(NULLなし, 重複なし)
        + _this.field.val+" TEXT);"];                   // 文字列
  };
  function createDropTableCommand(_this, table) {       // (存在する場合に、)テーブル削除
    return ["DROP TABLE IF EXISTS ["+table+"];"];
  };
  function createReplaceCommand(_this, table, key, val) {       // レコードの置き換え
    return ["REPLACE INTO ["+table+"]"
        + " ("+_this.field.key+", "+_this.field.val+") VALUES (?, ?);", key, val];
  };
  function createInsertCommand(_this, table, key, val) {        // レコード挿入(ある場合、無処理)
    return ["INSERT OR IGNORE INTO ["+table+"]"
        + " ("+_this.field.key+", "+_this.field.val+") VALUES (?, ?);", key, val];
  };
  function createSelectValueCommand(_this, table, key) {        // 値の取得
    return ["SELECT "+_this.field.val+" FROM ["+table+"]"
        + " WHERE "+_this.field.key+" = ?;", key];
  };
  function createSelectExistsCommand(_this, table, field, fieldValue) { // 条件合致の最初1件取得
    return ["SELECT "+_this.field.key+" FROM ["+table+"]"
        + " WHERE "+field+" = ? LIMIT 1;", fieldValue];
  };
  function createSelectCountCommand(_this, table) {                     // レコード数の取得
    return ["SELECT COUNT("+_this.field.key+") FROM ["+table+"];"];
  };
  function createSelectKeysCommand(_this, table, limit, offset) {       // 鍵一覧の取得
    return ["SELECT ALL "+_this.field.key+" FROM ["+table+"]"
        + ((limit  !== void 0) ? " LIMIT " +limit : "")
        + ((offset !== void 0) ? " OFFSET "+offset: "")
        + ";"];
  };
  function createSelectUniqValuesCommand(_this, table, limit, offset) { // ユニークな値一覧取得
    return ["SELECT DISTINCT "+_this.field.val+" FROM ["+table+"]"
        + ((limit  !== void 0) ? " LIMIT " +limit : "")
        + ((offset !== void 0) ? " OFFSET "+offset: "")
        + ";"];
  };
  function createDeleteCommand(_this, table, key) {                     // レコードの削除
    return ["DELETE FROM ["+table+"] WHERE "+_this.field.key+" = ?;", key];
  };
  
  /**
   * コマンド作成
   * @private
   * @param {Array} command - コマンド([コマンド文字列, 引数...])
   * @return {ADODB.Command} 結果
   */
  function createCmd(_this, command) {
    var cmd = null;                             // ADODB.Command
    try {
      cmd = new ActiveXObject('ADODB.Command'); // 「Command」オブジェクト生成
      cmd.ActiveConnection = _this.con;         // 「Connection」オブジェクト設定
      cmd.CommandType = 1;                      // コマンドタイプをテキスト(1)に設定
      cmd.Prepared = true;                      // 「Prepared」を真に設定
      
      // 発行するコマンド文字列(SQL文)を設定 
      cmd.CommandText = command[0];             // コマンドを設定
      for (var i=1; i<command.length; i++) {
        cmd.Parameters(i-1).Value = command[i]; // 引数を設定
      }
    } catch (e) {
      if (_this.console) _this.console.errorStackTrace(e);
      cmd = null;
    }
    return cmd;
  }
  
  /**
   * データベースレコードセットの反復子
   */
  var Iterator = (function DBLite$Iterator_factory() {
    "use strict";
    
    var _this = function DBLite$Iterator_constructor() {
      this.initialize.apply(this, arguments);
    };
    
    /**
     * コンストラクタ
     * @param {ADODB.Recordset} rs - レコードセット
     */
    _this.prototype.initialize = function DBLite$Iterator_initialize(master, rs) {
      this.master = master;
      this.rs     = rs;                 // ADODB.Recordset
      this.count  = 0;
      this.fields = [];                 // フィールド名一覧
      try {
        for (var f=0; f<this.rs.Fields.Count; f++) {    // フィールド数分ループ
          this.fields.push(this.rs.Fields(f).Name);
        }
      } catch (e) {}
    };
    
    /**
     * クローズ
     * クローズ後、hasNext()は、falseを返します。
     */
    _this.prototype.close = function DBLite$Iterator_close() {
      if (this.rs !== null) {
        try {
          if (this.rs.State > 0) {      // 開いている場合
            this.rs.Close();
          }
        } catch (e) {
          if (this.master && this.master.console) this.master.console.errorStackTrace(e);
        } finally {
          this.rs = null;               // レコードセットを開放
          this.master = null;
        }
      }
    };
    /**
     * 次があるか
     * 次のレコードがある場合、trueを返します。
     * 次のレコードがない場合、Recordsetをクローズします。
     * @return {boolean} 結果
     */
    _this.prototype.hasNext = function DBLite$Iterator_hasNext() {
      var ret = false;
      if (this.rs !== null) {
        ret = !this.rs.EOF;
        if (ret === false) {            // 次がない時
          this.close();                 // 閉じる
        }
      }
      return ret;
    };
    /**
     * 次のレコードへ
     * フィールドが1つの場合、値を返します。
     * フィールドが2つ以上の場合、コンテナに値を格納します。
     * 次のレコードがない場合、Recordsetをクローズします。
     * @param {Object} container - コンテナ
     * @return {boolean|string} 結果
     *                          コンテナありの場合、false:失敗/true:成功
     *                          コンテナなしの場合、null:失敗/string:成功
     */
    _this.prototype.next = function DBLite$Iterator_next(containar) {
      var ret = (_isObject(containar))? false: null;    // コンテナあり:false/コンテナなし:null
      try {
        if (this.hasNext()) {
          if (containar !== void 0) {          // コンテナありの時
            for (var f=0; f<this.rs.Fields.Count; f++) {// フィールド数分ループ
              containar[this.rs.Fields(f).Name].push(this.rs.Fields(f).Value);
            }
            ret = true;
          } else if (this.rs.Fields.Count === 1) {      // フィールドが1つの時
            ret = this.rs.Fields(0).Value;
          }
          this.count++;                         // カウンタを進める
          this.rs.MoveNext();                   // 次へ進める
        }
      } catch (e) {
        if (this.master && this.master.console) this.master.console.errorStackTrace(e);
        this.close();
      }
      return ret;
    };
    /**
     * コンテナ作成
     * フィールド名を鍵とする、空の配列を格納したオブジェクトを返します。
     * next関数の引数として使用します。
     * @return {Object} container - コンテナ
     */
    _this.prototype.container = function DBLite$Iterator_container(obj) {
      if (obj == null)  obj = {}                // 既定値(={})
      for (var f=0; f<this.fields.length; f++) {// フィールド数分ループ
        if (obj[this.fields[f]] === void 0) {  // フィールドがない時
          obj[this.fields[f]] = [];
        }
      }
      return obj;
    };
    /**
     * 要素の取得数
     * next関数によって、いままでに取得した個数を返します。
     * @return {number} カウンタ
     */
    _this.prototype.counter = function DBLite$Iterator_counter() {
      return this.count;
    };
    
    return _this;
  })();
  
  /**
   * コンストラクタ
   * @constructor
   * @param {string} path - データベースのパス(相対パスで指定可能)
   */
  _this.prototype.initialize = function DBLite_initialize(path) {
    this.field    = {"key":"key","val":"val"};          // フィールドの名称
    this.con      = null;                               // ADODB.Connection
    this.fullpath = fs.GetAbsolutePathName(path);       // ファイルの絶対パス
    this.stack    = {};                                 // コマンドのスタック
    this.console  = void 0;                             // エラー出力用コンソール
    if (Console) {
      this.console = Console.getConsole();
    }
  };
  
  /**
   * オープン
   * @return {boolean} 結果(true:接続成功or接続済み)
   */
  _this.prototype.open = function DBLite_open() {
    var ret = false;
    try {
      if (this.con === null) {
        _FileUtility_createFileFolder(this.fullpath);    // フォルダを作成
        
        // SQLiteにドライバ経由で接続
        // <http://www.ch-werner.de/sqliteodbc/>のドライバが必要
        this.con = new ActiveXObject("ADODB.Connection");
        this.con.ConnectionString =             // 接続文字列設定
              "DRIVER=SQLite3 ODBC Driver;"     // ODBCドライバ
            + "DATABASE="+this.fullpath+";";    // データベース
        this.con.Open();                        // 接続開始
        ret = true;
      } else {
        ret = (this.con.State > 0);             // 接続中の時
      }
    } catch (e) {
      if (this.console) this.console.errorStackTrace(e);
      ret = false;
    }
    return ret;
  };
  
  /**
   * クローズ
   * pushの残りを書き込む。
   * @return {boolean} 結果(true切断成功or未接続)
   */
  _this.prototype.close = function DBLite_close() {
    var ret = false;
    try {
      if (this.con !== null) {
        if (this.con.State > 0) {       // 接続中の時
          this.flushAll();              // push済みのデータを書き込み
          this.con.Close();             // 切断
        }
        this.con = null;                // 開放
      }
      ret = true;
    } catch (e) {
      if (this.console) this.console.errorStackTrace(e);
    }
    return ret;
  };
  
  /**
   * コマンド実行
   * Iteratorは、0件ヒットの可能性がある。
   * @param {boolean} isResult - 結果の出力有無
   * @param {Array} command - コマンド([コマンド文字列, 引数, ...])
   * @return {(boolean|Object)} 結果(null|true|Iterator)
   */
  _this.prototype.execute = function DBLite_execute(command) {
    var ret = null;
    if (this.con === null || this.con.State === 0) {    // 未接続
      return ret;
    }
    var cmd = null;                     // ADODB.Command
    var rs  = null;                     // ADODB.Recordset
    try {
      cmd = createCmd(this, command);   // ADODB.Command作成
      rs  = cmd.Execute();              // 実行
      if (rs.State > 0) {               // レコードがある時(戻り値ありの時)
        ret = new Iterator(this, rs);
      } else {
        ret = true;                     // 戻り値なしの実行結果
      }
    } catch (e) {
      if (this.console) this.console.errorStackTrace(e);
      ret = null;
    } finally {
      rs  = null;
      cmd = null;
    }
    return ret;
  };
  
  /**
   * トランザクションあり複数コマンド実行
   * 引数に設定した、コマンドをすべて実行する。
   * 処理は、トランザクションで行うため、失敗するとロールバックする。
   * @param {Array} command - コマンド([コマンド文字列, 引数, ...])
   * @return {boolean} 結果
   */
  _this.prototype.exec = function DBLite_exec(command_args) {
    var ret = false,
        i = 0, 
        command = null;
    try {
      if (this.open()) {                        // 接続
        this.con.BeginTrans();                  // トランザクション開始
        for (i=0; i<arguments.length; i++) {
          command = arguments[i];
          if (this.execute(command) !== true) { // 失敗の時
            throw null;
          }
        }
        this.con.CommitTrans();                 // トランザクション、コミット
        ret = true;
      }
    } catch (e1) {
      e1.message += "\n" + i+"/"+arguments.length+":"+JSON.stringify(command);
      if (this.console) this.console.errorStackTrace(e1);
      try {
        this.con.RollbackTrans();               // トランザクション、ロールバック
      } catch (e2) {
        e2.message += "\n" + "RollbackTrans";
        if (this.console) this.console.errorStackTrace(e2);
      }
    }
    return ret;
  };
  
  /**
   * 反復子作成
   * コマンドの結果を反復子として返します。
   * 反復子は、hasNext()がfalseを返す。
   * または、明示的にclose()を実施する必要があります。終了処理が必要です。
   * iteratorは、取得時のDBの状態を返します。
   * そのため、取得後に挿入/削除があっても挿入/削除前の値を返します。
   * @param {Array} command - コマンド([コマンド文字列, 引数, ...])
   * @return {Iterator} Iterator
   */
  _this.prototype.iterator = function DBLite_iterator(command) {
    var ret = null;
    if (this.open()) {                  // 接続
      var ite = this.execute(command);  // コマンド実行
      if (ite && ite !== true && ite.constructor === Iterator) {
        ret = ite;                      // レコードがある時
      }
    }
    return (ret !== null)? ret: new Iterator(this, null);       // 取得失敗時は、空の反復子を返す
  };
  
  /**
   * データ参照(単数)
   * 取得レコードが1つである場合、値を返します。
   * コンテナ(空のオブジェクト)を指定した場合、コンテナに結果を格納します。
   * ただし、成功/失敗に関わらず、コンテナに値を格納するため、戻り値で成功/失敗を判定します。
   * @param {Array} command - コマンド([コマンド文字列, 引数, ...])
   * @param {Object} container - コンテナ
   * @return {(boolean|number|string)} 結果
   *                                   コンテナありの場合、成功:true/失敗:false
   *                                   コンテナなしの場合、成功:値/失敗:null
   */
  _this.prototype.select = function DBLite_select(command, container) {
    var ret = (container !== void 0)? false: null;      // コンテナあり:false, なし:null
    var ite = this.iterator(command);   // 反復子取得
    if (container !== void 0) {         // コンテナがある時
      ite.container(container);         // コンテナを初期化
    }
    var obj = ite.next(container);                      // 1番目を取得
    if ((container !== void 0 && obj === true)          // コンテナあり成功の時
     || (container === void 0 && obj !== null)) {       // コンテンなし成功の時
      if (!ite.hasNext()) {             // 2番目がない時
        ret = obj;
      }
    }
    ite.close();                        // 明示的に閉じる
    ite = null;
    return ret;
  };
  
  /**
   * データ参照(複数)
   * 取得レコードを配列として返します。
   * コンテナ(空のオブジェクト)を指定した場合、コンテナに結果を格納します。
   * ただし、成功/失敗に関わらず、コンテナに値を格納するため、戻り値で成功/失敗を判定します。
   * @param {Array} command - コマンド([コマンド文字列, 引数, ...])
   * @param {Object} container - コンテナ
   * @return {(boolean|Array)} 結果
   *                           コンテナありの場合、成功:true/失敗:false
   *                           コンテナなしの場合、成功:値/失敗:null
   */
  _this.prototype.selects = function DBLite_selects(command, container) {
    var ret = (container !== void 0)? false: [];// コンテナあり:false, なし:null
    var ite = this.iterator(command);           // 反復子取得
    if (container !== void 0) {                 // コンテナがある時
      ite.container(container);                 // コンテナを初期化
    }
    var obj = ite.next(container);              // 1番目を取得
    if (container !== void 0 && obj === true) { // コンテナあり成功の時
      while (ite.next(container));              // コンテナに格納
      ret = true;
    }
    if (container === void 0 && obj !== null) { // コンテナなし成功の時
      ret = [obj];
      while ((obj=ite.next()) !== null) {       // 末尾までループ
        ret.push(obj);                          // 配列に格納
      }
    }
    ite.close();                                // 明示的に閉じる
    ite = null;
    return ret;
  };
  
  /**
   * 作成済みのテーブル名を取得
   * @return {Array} テーブル名一覧
   */
  _this.prototype.getTableNames = function DBLite_getTableNames() {
    return this.selects(createTablesCommand(this));
  };
  
  /**
   * テーブル有無
   * @param {string} name - テーブル名
   * @return {boolean} 結果
   */
  _this.prototype.existsTable = function DBLite_existsTable(name) {
    return (this.select(createSchemaCommand(this, name)) !== null);
  };
  
  /**
   * テーブル追加
   * @param {string} name - テーブル名
   * @return {TableLite} TableLite
   */
  _this.prototype.getTable = 
  _this.prototype.createTable = function DBLite_createTable(name) {
    return (this.exec(createCreateTableCommand(this, name)) === true)?
        new TableLite(this, name):
        null;
  };
  
  /**
   * テーブル削除
   * @param {string} name - テーブル名
   * @return {boolean} 結果
   */
  _this.prototype.deleteTable = 
  _this.prototype.dropTable = function DBLite_dropTable(name) {
    return this.exec(createDropTableCommand(this, name));
  };
  
  /**
   * 鍵がテーブルに含まれるか
   * @param {string} table - テーブル名
   * @param {string} key - 鍵
   * @return {boolean} 結果
   */
  _this.prototype.containsKey = function DBLite_containsKey(table, key) {
    var command = createSelectExistsCommand(this, table, this.field.key, key);
    return (this.select(command) !== null);
  };
  
  /**
   * 値がテーブルに含まれるか
   * @param {string} table - テーブル名
   * @param {Object} val - 値
   * @return {boolean} 結果
   */
  _this.prototype.containsValue = function DBLite_containsValue(table, val) {
    var command = createSelectExistsCommand(this, table, this.field.val, JSON.stringify(val));
    return (this.select(command) !== null);
  };
  
  /**
   * 鍵がテーブルに含まれるか
   * containsKey関数のラップ関数
   * @param {string} table - テーブル名
   * @param {string} key - 鍵
   * @return {boolean} 結果
   */
  _this.prototype.contains = function DBLite_contains(table, key) {
    return this.containsKey(table, key);
  };
  
  /**
   * 値の設定
   * 鍵と値を配列とした場合、配列内の要素を鍵と値として複数設定する。
   * 処理はトランザクションとして実行するため、1つでも失敗すると未処理となります。
   * 本関数を短時間に大量に呼び出した場合、
   * 呼び出し毎にファイルアクセスが発生するため、遅延が発生します。
   * 短時間に大量に呼び出す場合、push, flush関数の使用をおすすめします。
   * @param {string} table - テーブル名
   * @param {(string|Array)} key - 鍵
   * @param {(Object|Array)} val - 値
   * @param {boolean} [overwrite=true] - 上書き有無
   * @return {boolean} 結果
   */
  _this.prototype.set = function DBLite_set(table, key, val, overwrite) {
    if (!_isArray(key)) {       // 配列でない時
      key = [key];              // 配列化
      val = [val];
    }
    var createCommand = (overwrite !== false)? createReplaceCommand: createInsertCommand;
    var commands = [];
    for (var i=0; i<key.length; i++) {
      commands.push(createCommand(this, table, key[i], JSON.stringify(val[i])));
    }
    return this.exec.apply(this, commands);
  };
  
  /**
   * 値の取得
   * @param {string} table - テーブル名
   * @param {string} key - 鍵
   * @param {Object} def - 取得失敗時の値
   * @return {Object} 値 or 取得失敗時の値
   */
  _this.prototype.get = function DBLite_get(table, key, def) {
    var text = this.select(createSelectValueCommand(this, table, key));
    return (text !== null) ? JSON.parse(text): def;
  };
  
  /**
   * レコード件数
   * @param {string} table - テーブル名
   * @return {number} レコードの件数
   */
  _this.prototype.size = function DBLite_size(table) {
    return this.select(createSelectCountCommand(this, table));
  };
  
  /**
   * 鍵の配列
   * 取得する鍵の配列が大きい場合、メモリ領域を圧迫することに注意してください。
   * 取得する鍵の配列が大きい場合、Iteratorの使用をおすすめします。
   * 注意：Iterator使用の場合、専用関数を使用してください。limit=trueは、廃止の可能性があります。
   * @param {string} table - テーブル名
   * @param {string} key - 鍵
   * @param {(boolean|number)} limit - 取得最大レコード数(trueならば、戻り値(Iterator))
   * @param {number} offset - レコード取得開始位置
   * @return {Array} 鍵の配列 or null
   */
  _this.prototype.keys = function DBLite_keys(table, limit, offset) {
    return (limit === true) ?
        this.iterator(createSelectKeysCommand(this, table)):                // true: Iterator
        this.selects(createSelectKeysCommand(this, table, limit, offset));  // true以外: 配列
  };
  _this.prototype.keysIterator = function DBLite_keysIterator(table, limit, offset) {
    return this.iterator(createSelectKeysCommand(this, table, limit, offset));
  };
  
  /**
   * ユニークな値の配列
   * 同一の値をまとめて、ユニークな配列を返します。
   * 注意：Iterator使用の場合、専用関数を使用してください。limit=trueは、廃止の可能性があります。
   * @param {string} table - テーブル名
   * @param {string} key - 鍵
   * @param {(boolean|number)} limit - 取得最大レコード数(trueならば、戻り値(Iterator))
   * @param {number} offset - レコード取得開始位置
   * @return {Array} 値の配列 or null
   */
  _this.prototype.values = function DBLite_values(table, limit, offset) {
    var ret = null;
    if (limit === true) {
      ret = this.iterator(createSelectUniqValuesCommand(this, table));
    } else {
      var vals = this.selects(createSelectUniqValuesCommand(this, table, limit, offset));
      ret = [];
      for (var i=0; i<vals.length; i++) {
        ret.push(JSON.parse(vals[i]));  // オブジェクトに復元
      }
    }
    return ret;
  };
  _this.prototype.valuesIterator = function DBLite_valuesIterator(table, limit, offset) {
    return this.iterator(createSelectUniqValuesCommand(this, table, limit, offset));
  };
  
  /**
   * レコード削除
   * @param {string} table テーブル名
   * @param {string} key - 鍵
   * @return {boolean} 結果
   */
  _this.prototype.del = function DBLite_del(table, key) {
    return this.exec(createDeleteCommand(this, table, key));
  };
  
  /**
   * スタック追加
   * 追加したコマンドは、flush()で実行します。
   * close()時に、すべてのコマンドをflush()します。(成功を保証するものではありません)
   * 補足:set関数では、set毎にファイルアクセスが発生するため、処理が遅い。
   * @param {string} table - テーブル名
   * @param {Array} command - コマンド([コマンド文字列, 引数, ...])
   */
  _this.prototype.pushCommand = function DBLite_pushCommand(table, command) {
    if (this.stack[table] === void 0)  this.stack[table] = [];  // 未使用時に配列を準備
    this.stack[table].push(command);                            // コマンドを登録
  };
  _this.prototype.pushSet = function DBLite_pushSet(table, key, val, overwrite) {
    var createCommand = (overwrite !== false)? createReplaceCommand: createInsertCommand;
    this.pushCommand(table, createCommand(this, table, key, JSON.stringify(val)));
  };
  _this.prototype.pushDel = function DBLite_pushDel(table, key) {
    this.pushCommand(table, createDeleteCommand(this, table, key));
  };
  _this.prototype.push = function DBLite_push(table, key, val, overwrite) {
    this.pushSet(table, key, val, overwrite);
  };
  
  /**
   * スタック実行
   * スタック内のコマンドを実行する。
   * 限界値(cv)を設定した場合、限界値に達していなければ処理を行わない。
   * @param {string} table - テーブル名
   * @param {number} cv - 限界値
   * @return {boolean} 結果(true/false:実行結果、null:未処理)
   */
  _this.prototype.flush = function DBLite_flush(table, cv) {
    var ret = true;                     // 未処理(=null)
    if (this.stack[table] !== void 0) { // スタックがある時
      if (cv === void 0 || cv <= this.stack[table].length) {    // 臨界値以上の時
        ret = this.exec.apply(this, this.stack[table]);
        this.stack[table] = [];
        // 補足:成功/失敗に関わらず、スタックをクリアする。
        //      残しておいてもどっちみち次も失敗するため。
      }
    }
    return ret;
  };
  _this.prototype.flushAll = function DBLite_flushAll() {
    var ret = true;                     // 初期値
    for (var table in this.stack) {     // stackの鍵(テーブル名)でループ
      ret = (this.flush(table) && ret); // flush実行(一度でも失敗(=false))
    }
    return ret;
  };
  
  var TableLite = (function DBLite$TableLite_factory() {
    "use strict";
    
    var _this = function DBLite$TableLite_constructor() {
      this.initialize.apply(this, arguments);
    };
    
    /**
     * テーブルアクセス用のIF
     * @param {DBLite} master - 親クラス
     * @param {string} name - テーブル名
     */
    _this.prototype.initialize = function DBLite$TableLite_initialize(master, name) {
      this.master = master;
      this.name = name;
    }
    
    /**
     * テーブル名取得
     * @return {string} テーブル名
     */
    _this.prototype.getName = function DBLite$TableLite_getName() {
      return this.name;
    };
    
    /**
     * テーブルを閉じる
     */
    _this.prototype.close = function DBLite$TableLite_close() {
      this.master = null;
      this.name = null;
    };
    
    /**
     * テーブル削除
     * @return {boolean} 結果
     */
    _this.prototype.drop = function DBLite$TableLite_drop() {
      if (this.master.dropTable(this.name)) {
        this.close();
      }
      return ret;
    };
    
    // DBLiteのラップ関数郡
    _this.prototype.containsKey = function DBLite$TableLite_containsKey(key) {
      return this.master.containsKey(this.name, key);
    };
    _this.prototype.containsValue = function DBLite$TableLite_containsValue(val) {
      return this.master.containsValue(this.name, val);
    };
    _this.prototype.contains = function DBLite$TableLite_contains(key) {
      return this.master.contains(this.name, key);
    };
    _this.prototype.set = function DBLite$TableLite_set(key, val, overwrite) {
      return this.master.set(this.name, key, val, overwrite);
    };
    _this.prototype.get = function DBLite$TableLite_get(key, def) {
      return this.master.get(this.name, key, def);
    };
    _this.prototype.del = function DBLite$TableLite_del(key) {
      return this.master.del(this.name, key);
    };
    _this.prototype.size = function DBLite$TableLite_size() {
      return this.master.size(this.name);
    };
    _this.prototype.keys = function DBLite$TableLite_keys(limit, offset) {
      return this.master.keys(this.name, limit, offset);
    };
    _this.prototype.keysIterator = function DBLite$TableLite_keysIterator(limit, offset) {
      return this.master.keysIterator(this.name, limit, offset);
    };
    _this.prototype.values = function DBLite$TableLite_values(limit, offset) {
      return this.master.values(this.name, limit, offset);
    };
    _this.prototype.valuesIterator = function DBLite$TableLite_valuesIterator(limit, offset) {
      return this.master.valuesIterator(this.name, limit, offset);
    };
    _this.prototype.pushCommand = function DBLite$TableLite_pushCommand(command) {
      this.master.pushCommand(this.name, command);
    };
    _this.prototype.pushSet = function DBLite$TableLite_pushSet(key, val, overwrite) {
      this.master.pushSet(this.name, key, val, overwrite);
    };
    _this.prototype.pushDel = function DBLite$TableLite_pushDel(key) {
      this.master.pushDel(this.name, key);
    };
    _this.prototype.push = function DBLite$TableLite_push(key, val, overwrite) {
      this.master.push(this.name, key, val, overwrite);
    };
    _this.prototype.flush = function DBLite$TableLite_flush(cv) {
      return this.master.flush(this.name, cv);
    };
    
    return _this;
  })();
  
  return _this;
});
