"use strict";

/**
 * 簡易データベース
 * JSONと同じような感覚で使用できることを目標とする。
 * @auther toshi2limu@gmail.com (toshi)
 */
var DBLite = function (path) {
	this.initialize.apply(this, arguments);
};
(function() {
	/**
	 * コンストラクタ
	 * @override
	 * @public
	 * @constructor
	 * @param {string} path		データベースのパス(相対パスで指定可能)
	 */
	this.prototype.initialize = function (path) {
		this.con = null;									// コネクション
		this.fullpath = Utility.getAbsolutePathName(path);	// 絶対パス
	};
	/**
	 * オープン
	 * @return {boolean}		成否(接続成功 or 接続済み時にtrue)
	 */
	this.prototype.open = function () {
		var ret = false;
		try {
			if (this.con == null) {
				Utility.createFolder(this.fullpath);			// フォルダを作成
				
				// SQLiteにドライバ経由で接続
				// <http://www.ch-werner.de/sqliteodbc/>のドライバが必要
				this.con = new ActiveXObject("ADODB.Connection");
				this.con.ConnectionString = 					// 接続文字列設定
						  "DRIVER=SQLite3 ODBC Driver;"			// ODBCドライバ
						+ "DATABASE="+this.fullpath+";";		// データベース
				this.con.Open();								// 接続開始
				ret = true;
			} else {
				ret = (this.con.State == 1);					// 接続中の時
			}
		} catch (e) {
			println("Error(" + (e.number & 0xFFFF) + "):" + e.message);
			println("    DBLite.open()");
		}
		return ret;
	};
	/**
	 * クローズ
	 * @return {boolean}		成否(切断成功 or 未接続時にtrue)
	 */
	this.prototype.close = function () {
		var ret = false;
		try {
			if (this.con != null) {
				if (this.con.State == 1) {					// 接続中の時
					this.con.Close();						// 切断
					this.con = null;
					ret = true;
				}
			} else {										// 未初期化
				ret = true;
			}
		} catch (e) {
			println("Error(" + (e.number & 0xFFFF) + "):" + e.message);
			println("    DBLite.close()");
		}
		return ret;
	};
	/**
	 * コマンド実行
	 * @param {string} command	コマンド
	 * @return {boolean}		成否
	 */
	this.prototype.execute = function (command) {
		var ret = false;
		if (this.open() == false) {					// 接続
			throw new Error("DBLite.execute.open()");
		}
		try {
			var cmd = new ActiveXObject("ADODB.Command");	// 「Command」オブジェクト生成
			cmd.ActiveConnection = this.con;				// 「Connection」オブジェクト設定
			cmd.CommandType = 1;							// コマンドタイプをテキスト(1)に設定
			cmd.Prepared = true;							// 「Prepared」を真に設定
			
			// 発行するコマンド文字列(SQL文)を設定 
			cmd.CommandText = command;						// コマンドを設定
			for (var i=1; i<arguments.length; i++) {
				cmd.Parameters(i-1).Value = arguments[i];	// 引数を設定
			}
			
			cmd.Execute();									// 実行
			ret = true;
		} catch (e) {
			println("Error(" + (e.number & 0xFFFF) + "):" + e.message);
			print("    DBLite.execute(command=\""+command+"\"");
			for (var i=1; i<arguments.length; i++) {
				print(", "+arguments[i]);
			}
			println(")");
		}
		return ret;
	};
	/**
	 * コマンド実行
	 * @param {Array} commands	[コマンド, 引数, ...]の形式で格納, 第2引数以降も同様の形式
	 * @return {boolean}		成否
	 */
	this.prototype.executeMultiCommit = function (commands) {
		var ret = false;
		var command = null;
		
		if (this.open() == false) {						// 接続
			throw new Error("DBLite.execute.open()");
		}
		this.con.BeginTrans();							// トランザクション開始
		try {
			for (var i=0; i<arguments.length; i++) {
				command = arguments[i];
				var cmd = new ActiveXObject("ADODB.Command");	// 「Command」オブジェクト生成
				cmd.ActiveConnection = this.con;				// 「Connection」オブジェクト設定
				cmd.CommandType = 1;							// コマンドタイプをテキスト(1)に設定
				cmd.Prepared = true;							// 「Prepared」を真に設定
				
				// 発行するコマンド文字列(SQL文)を設定 
				cmd.CommandText = command[0];						// コマンドを設定
				for (var k=1; k<command.length; k++) {
					cmd.Parameters(k-1).Value = command[k];	// 引数を設定
				}
				cmd.Execute();									// 実行
			}
			this.con.CommitTrans();							// トランザクション、コミット
			ret = true;
		} catch (e) {
			println("Error(" + (e.number & 0xFFFF) + "):" + e.message);
			if (command != null) {
				print("    DBLite.executeMultiCommit(command=\""+command[0]+"\"");
				for (var i=1; i<command.length; i++) {
					print(", "+command[i]);
				}
				println(")");
			}
			try {
				this.con.RollbackTrans();					// トランザクション、ロールバック
			} catch (e) {
				println("Error(" + (e.number & 0xFFFF) + "):" + e.message);
				println("    DBLite.executeMultiCommit.rollback()");
			}
		}
		return ret;
	};
	/**
	 * コマンド実行
	 * @param {string} command	コマンド
	 * @return {boolean}		成否
	 */
	this.prototype.executeCommit = function () {
		return this.executeMultiCommit(_.toArray(arguments));
	};
	/**
	 * コマンド実行
	 * @param {string} command	コマンド
	 * @return {Object}			値({キー: [データ, ...], ...}) or null
	 */
	this.prototype.executeSelect = function (command) {
		var ret = null;
		var rs  = null;
		if (this.open() == false) {						// 接続
			throw new Error("DBLite.execute.open()");
		}
		try {
			var cmd = new ActiveXObject("ADODB.Command");	// 「Command」オブジェクト生成
			cmd.ActiveConnection = this.con;				// 「Connection」オブジェクト設定
			cmd.CommandType = 1;							// コマンドタイプをテキスト(1)に設定
			cmd.Prepared = true;							// 「Prepared」を真に設定
			
			// 発行するコマンド文字列(SQL文)を設定 
			cmd.CommandText = command;						// コマンドを設定
			for (var i=1; i<arguments.length; i++) {
				cmd.Parameters(i-1).Value = arguments[i];	// 引数を設定
			}
			
			rs = cmd.Execute();								// 実行
			
			ret = {};
			for (var i=0; i<rs.Fields.Count; i++) {			// フィールド数分ループ
				ret[rs.Fields(i).Name] = [];
			}
			while (!rs.EOF) {								// 終端まで繰り返す。
				for (var i=0; i<rs.Fields.Count; i++) {		// フィールド数分ループ
					ret[rs.Fields(i).Name].push(rs.Fields(i).Value);
				}
				rs.MoveNext();								// 次の行へ移動
			}
		} catch (e) {
			println("Error(" + (e.number & 0xFFFF) + "):" + e.message);
			print("    DBLite.executeSelect(command=\""+command+"\"");
			for (var i=1; i<arguments.length; i++) {
				print(", "+arguments[i]);
			}
			println(")");
			ret = null;
		} finally {
			try {
				if ((rs != null) && (rs.State == 1)) {
					rs.Close();								// 「Recordset」を閉じる
				}
			} catch (e) {
				println("Error(" + (e.number & 0xFFFF) + "):" + e.message);
				println("    DBLite.executeSelect.close()");
				ret = null;
			}
		}
		return ret;
	};
	/**
	 * コマンド実行
	 * @param {string} command	コマンド
	 * @return {integer|string}	値 or null
	 */
	this.prototype.executeSelectSimple = function (command) {
		var obj = this.executeSelect.apply(this, arguments);
		if (obj != null) {
			var keys = Object.keys(obj);
			if (keys.length == 1) {							// 名前が1つのみ
				if (obj[keys[0]].length == 1) {				// 要素が1つのみ
					return obj[keys[0]][0];					// 要素1つのみを返す。
				}
			}
		}
		return null;										// 想定外の戻り値
	};
	/**
	 * テーブル追加
	 * @param {string} table	テーブル名
	 * @param {string} type		鍵の型("TEXT","INT"など)
	 * @return {boolean}		成否
	 */
	this.prototype.createTable = function (table) {
		return this.execute("CREATE TABLE IF NOT EXISTS ["+table+"] ("// テーブル作成
							+ "idx INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE, "	// 重複なし(数値)
							+ "key TEXT UNIQUE, "					// 重複なし(文字列)
							+ "val TEXT);"							// 文字列(JSON)
							);
	};
	/**
	 * テーブル削除
	 * @param {string} table	テーブル名
	 * @return {boolean}		成否
	 */
	this.prototype.deleteTable = function (table) {
		return this.executeCommit("DROP TABLE IF EXISTS ["+table+"];");
	};
	/**
	 * 挿入(置換)
	 * @param {string} table	テーブル名
	 * @param {string} key		鍵
	 * @param {Object} val		値
	 * @return {boolean}		成否
	 */
	this.prototype.set = function (table, key, val) {
		return this.executeCommit("REPLACE INTO ["+table+"] (key, val) VALUES (?, ?);", 
									key, JSON.stringify(val));
	};
	/**
	 * 挿入(置換)
	 * @param {string} table	テーブル名
	 * @param {Array} key		鍵
	 * @param {Array} val		値
	 * @return {boolean}		成否
	 */
	this.prototype.setArray = function (table, key, val) {
		var min = Math.min(key.length, val.length);
		var commandText = "REPLACE INTO ["+table+"] (key, val) VALUES (?, ?);";
		var commands = [];
		for (var i=0; i<min; i++) {
			commands.push([commandText, key[i], JSON.stringify(val[i])]);
		}
		return this.executeMultiCommit.apply(this, commands);
	};
	/**
	 * 値(抽出)
	 * @param {string} table	テーブル名
	 * @param {string} key		鍵
	 * @return {Object}			値 or null
	 */
	this.prototype.get = function (table, key) {
		return JSON.parse(
				this.executeSelectSimple("SELECT val FROM ["+table+"] WHERE key = ?;", key));
	};
	/**
	 * 値(抽出)
	 * @param {string} table	テーブル名
	 * @param {Array} key		鍵
	 * @return {Array}			値 or null
	 */
	this.prototype.getArray = function (table, key) {
		var arg = [command, key[0]];
		var command = "SELECT val FROM ["+table+"] WHERE key IN (?";
		for (var i=1; i<key.length; i++) {
			command += ", ?";
			arg.push(key[i]);
		}
		command += ");";
		arg[0] = command;
		var obj = this.executeSelect.apply(this, arg);
		if (obj != null) {
			var keys = Object.keys(obj);
			if (keys.length == 1) {
				return obj[keys[0]];
			}
		}
		return null;
	};
	/**
	 * 鍵の配列
	 * TODO: 要テスト(min, max)
	 * @param {string} table	テーブル名
	 * @param {string} key		鍵
	 * @return {Array}			鍵の配列 or null
	 */
	this.prototype.getKeys = function (table, min, max) {
		var ret  = null;
		var arg = [];
		if (min && max) {
			arg.push("SELECT key FROM ["+table+"] WHERE idx >= min AND idx <= max;");
			arg.push(min);
			arg.push(max);
		} else if (min) {
			arg.push("SELECT key FROM ["+table+"] WHERE idx >= min;");
			arg.push(min);
		} else if (max) {
			arg.push("SELECT key FROM ["+table+"] WHERE idx <= max;");
			arg.push(max);
		} else {
			arg.push("SELECT key FROM ["+table+"];");
		}
		var keys =  this.executeSelect.apply(this, arg);
		if (keys == null) {
		} else if (keys.key == null) {
		} else {
			ret = keys.key;
		}
		return ret;
	};
	/**
	 * レコード件数
	 * @param {string} table	テーブル名
	 * @return {integer}		レコードの件数
	 */
	this.prototype.getLength = function (table) {
		return this.executeSelectSimple("SELECT COUNT(val) FROM ["+table+"];");
	};
	/**
	 * レコード削除
	 * @param {string} table	テーブル名
	 * @param {string} key		鍵
	 * @return {boolean}		成否
	 */
	this.prototype.del = function (table, key) {
		return this.executeCommit("DELETE FROM ["+table+"] WHERE key = ?;", key);
	};
	/**
	 * インデックスの振り直し
	 * TODO: 要テスト
	 * @param {string} table	テーブル名
	 * @return {boolean}		成否
	 */
	this.prototype.update = function (table) {
		return this.executeMultiCommit(
			["UPDATE test SET idx = (SELECT COUNT(*) FROM test s2 WHERE s2.idx <= test.idx);"], 
			["DELETE FROM sqlite_sequence WHERE name = '"+table+"';"]);
				// command1: idxを1から連番振り直し
				// command2: 連番をリセット(次は最大値から)
		// 補足: REPLACEのたびに連番が加算される。
	};
}).call(DBLite, null);
