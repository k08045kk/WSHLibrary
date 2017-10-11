// require("../common/Utility.js");
// require("../common/EncodeUtility.js");

/**
 * 簡易暗号化
 * @auther toshi2limu@gmail.com (toshi)
 */
(function(global, factory) {
	global.EasyCrypto = global.EasyCrypto || factory();
})(this, function EasyCrypto_factory() {
	"use strict";
	
	var _this = void 0;
	_this = function EasyCrypto_constrcutor() {};
	
	var prefix = 'EasyCrypto:';
	
	var locator = new ActiveXObject("WbemScripting.SWbemLocator");
	var service = locator.ConnectServer();
	
	function getGUID(fullpath) {
		var drive	= fs.GetDriveName(fullpath);
		var set	= service.ExecQuery("SELECT DeviceID FROM Win32_Volume WHERE DriveLetter = '"+drive+"'");
		
		var ret = [];
		for (var e1=new Enumerator(set); !e1.atEnd(); e1.moveNext()) {
			var obj = {};
			for (var e2=new Enumerator(e1.item().Properties_); !e2.atEnd(); e2.moveNext()) {
				var item2 = e2.item();
				obj[item2.Name] = item2.Value;
			}
			ret.push(obj);
		}
		return (ret.length == 1 && ret[0].DeviceID)? 
						ret[0].DeviceID.match(/{(.+)}/)[1].split("-").join(""): null;
	}
	
	// 「byte配列」暗号化(Rijndael暗号)
	_this.encrypt = function EasyCrypto_encrypt(text) {
		var guid= getGUID(ModulePath());
		var bin	= EncodeUtility.encrypt(EncodeUtility.hex2bin(guid), EncodeUtility.str2bin(text));
		return prefix + EncodeUtility.bin2base64(bin);
	};
	
	// 「byte配列」暗号化(Rijndael暗号)
	_this.decrypt = function EasyCrypto_decrypt(text) {
		ret = text;
		if (text.startsWith(prefix)) {
			var guid = EncodeUtility.hex2bin(getGUID(ModulePath()));
			var code = EncodeUtility.base642bin(text.substr(prefix.length));
			var bin  = EncodeUtility.decrypt(guid, code);
			ret = EncodeUtility.bin2str(bin);
		}
		return ret;
	};
	
	return _this;
});
