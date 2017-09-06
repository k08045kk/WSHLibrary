/**
 * WMI(Windows Management Instrumentation)
 * @auther toshi2limu@gmail.com (toshi)
 */
(function(global, factory) {
	global.WMIUtility = factory();
})(this, function WMIUtility_factory() {
	"use strict";
	
	var _this = void 0;
	_this = function WMIUtility_constructor() {};
	
	_this.mypid = null;	// 自身のプロセスID
	
	var locator = new ActiveXObject("WbemScripting.SWbemLocator");
	var service = locator.ConnectServer();
	
	_this.getDate = function WMIUtility_getDate(str) {
		// 想定値:"20170902135602.687944+540"
		var m = str.match(/^(\d+)(\d\d)(\d\d)(\d\d)(\d\d)(\d\d)\./);
		return (m!=null)? new Date(m[1]|0, m[2]-1|0, m[3]|0, m[4]|0, m[5]|0, m[6]|0): new Date(str);
	};
	_this.getSelfProcessId = function WMIUtility_getSelfProcessId() {
		if (_this.mypid == null) {												// 1回のみ実行
			var text = '(function(){while(true)WScript.Sleep(60*1000);})();';	// 停止プログラム
			var temp = FileUtility.getTempFilePath(null, "jse");
			FileUtility.storeText(text, temp, true, FileUtility.UTF_16);		// プログラム保存
			
			var obj = sh.Exec('cscript "'+temp+'"');
			var process = _this.get(obj.ProcessId);
			_this.mypid = _this.getProperty(process, "ParentProcessId");		// 親プロセス取得
			_this.Terminate(process);	// 強制終了
			process	= null;
			obj		= null;
			
			fs.DeleteFile(temp);		// プログラム削除
		}
		return _this.mypid;
	};
	_this.getProcessId = function WMIUtility_getProcessId(name) {
		if (!name) {
			return _this.getSelfProcessId();
		}
		var ret = null;
		var set = service.ExecQuery("SELECT * FROM Win32_Process"
									+ " WHERE Caption = 'cscript.exe' OR Caption = 'wscript.exe'");
		
		process: for (var e=new Enumerator(set); !e.atEnd(); e.moveNext()) {
			var item = e.item();
			if (item.CommandLine) {
				var commands = item.CommandLine.split(" ");
				for (var i=0; i<commands.length; i++) {
					if (commands[i].startsWith("/wminame:")) {
						if (commands[i].substr(9) == name) {
							ret = item.ProcessId;
							break process;
						}
					}
				}
			}
		}
		
		set = null;
		return ret;
	};
	_this.get = function WMIUtility_get(pid) {
		var ret = null;
		var set = service.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = '"+pid+"'");
		
		for (var e=new Enumerator(set); !e.atEnd(); e.moveNext()) {
			ret = e.item();
			break;	// pidの重複はないはず
		}
		
		set = null;
		return ret;
	};
	_this.getProperties = function WMIUtility_getProperties(pid) {
		var obj = {};
		var process = (Atom.isObject(pid))? pid: _this.get(pid);
		if (process != null) {
			for (var e=new Enumerator(process.Properties_); !e.atEnd(); e.moveNext()) {
				var item = e.item();
				obj[item.Name] = item.Value;
			}
		}
		process = null;
		return obj;
	};
	_this.getProperty = function WMIUtility_getProperty(pid, property) {
		var ret = null;
		var process = (Atom.isObject(pid))? pid: _this.get(pid);
		if (process != null) {
			ret = process[property];
		}
		process = null;
		return ret;
	};
	_this.Terminate = function WMIUtility_Terminate(pid) {
		var ret = -1;
		var process = (Atom.isObject(pid))? pid: _this.get(pid);
		if (process != null) {
			ret = process.Terminate();
		}
		process = null;
		return ret;
	};
	
	_this.Create = function WMIUtility_Create(command, current, startup) {
		var process = service.Get("Win32_Process");
		var ret = process.Create(command, current, startup);
//		ret = (ret == 0)? process.ProcessId: -ret;
		process = null;
		return ret;
		// 補足:Create関数の第4引数のProcessIdを取得したかったが、out引数のため、無理？
	};
	
	_this.Run = function WMIUtility_Create(commandline, current) {
		current = (current)? current: sh.CurrentDirectory;
		
		// 実行プログラム
		function exec(commandline, current) {
			var sh = new ActiveXObject("WScript.Shell");
			sh.CurrentDirectory = current;
			var obj = sh.Exec(commandline);
			if (obj && obj.ProcessId) {
				WScript.Quit(obj.ProcessId);
			}
			obj = null;
			WScript.Quit(-1);
		}
		current		= current.split('\\').join('\\\\');
		commandline	= commandline.split('\\').join('\\\\');
		var text='('+exec.toString()+'})("'+commandline+'", "'+current+'");';
		
		// 一時実行ファイル作成 & 実行 & 削除
		var temp = FileUtility.getTempFilePath(null, "jse");
		FileUtility.storeText(text, temp, true, FileUtility.UTF_16);
		
		var bkup = sh.CurrentDirectory;
		sh.CurrentDirectory = fs.GetParentFolderName(temp);
		var ret = sh.Run('cscript "'+temp+'"', 0, true);
		sh.CurrentDirectory = bkup;
		
		fs.DeleteFile(temp);
		return ret;
	};
	
	return _this;
});
