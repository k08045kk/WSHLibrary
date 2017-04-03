// import: ActiveXObject("Msxml2.DOMDocument");
// import: ActiveXObject("System.Text.UTF8Encoding");
// import: ActiveXObject("System.Security.Cryptography.MD5CryptoServiceProvider");
// import: ActiveXObject("System.Security.Cryptography.SHA1CryptoServiceProvider");
// import: ActiveXObject("System.Security.Cryptography.SHA256CryptoServiceProvider");

/**
 * 変換系
 * @auther toshi2limu@gmail.com (toshi)
 */
var EncodeUtility = function () {
};
(function() {
	"use strict";
	
	var doc		= null;
	var encorde	= null;
	var md5		= null;
	var sha1	= null;
	var sha256	= null;
	
	// 文字列→値
	function text2value(type, text) {
		if (doc == null) {
			doc = new ActiveXObject("Msxml2.DOMDocument");
		}
		var element = doc.createElement("temp");
		element.dataType		= type;						// 種別を設定
		element.text			= text;						// 変換元を設定
		var ret = element.nodeTypedValue;					// 値を取り出す
		element = null;
		return ret;
	};
	
	// 値→文字列
	function value2text(type, value) {
		if (doc == null) {
			doc = new ActiveXObject("Msxml2.DOMDocument");
		}
		var element = doc.createElement("temp");
		element.dataType		= type;						// 種別を設定
		element.nodeTypedValue	= value;					// 変換元を設定
		var ret = element.text;								// 値を取り出す
		element = null;
		return ret;
	};
	
	/**
	 * 「文字列」から「byte配列」へ変換
	 * @public
	 * @param {string} string	文字列(UTF-8)
	 * @return {Array}			byte配列
	 */
	this.string2bytes = function string2bytes(string) {
		if (encorde == null) {
			encorde = new ActiveXObject("System.Text.UTF8Encoding");
		}
		return encorde.GetBytes_4(string);
	};
	
	/**
	 * 「byte配列」から「文字列」へ変換
	 * @public
	 * @param {Array} bytes		byte配列
	 * @return {string}			文字列(UTF-8)
	 */
	this.bytes2string = function bytes2string(bytes) {
		if (encorde == null) {
			encorde = new ActiveXObject("System.Text.UTF8Encoding");
		}
		return encorde.GetString(bytes);
	};
	
	// 「byte配列」から「16進数文字列」
	this.bytes2hex = function bytes2hex(bytes) {
		return value2text("bin.hex", bytes);
	}
	
	// 「16進数文字列」から「byte配列」
	this.hex2bytes = function hex2bytes(hex) {
		return text2value("bin.hex", hex);
	}
	
	// 「byte配列」から「Base64文字列」
	this.bytes2hex = function bytes2base64(bytes) {
		return value2text("bin.base64", bytes);
	}
	
	// 「Base64文字列」から「byte配列」
	this.hex2bytes = function base642bytes(base64) {
		return text2value("bin.base64", base64);
	}
	
	function crypto(provider, bytes) {
		provider.ComputeHash_2(bytes);
		var hash = provider.Hash;
		provider.Clear();
		return hash;
	};
	// md5
	this.md5 = function md5(bytes) {
		if (md5 == null) {
			md5 = new ActiveXObject("System.Security.Cryptography.MD5CryptoServiceProvider");
		}
		return crypto(md5, bytes);
	};
	
	// sha1
	this.sha1 = function sha1(bytes) {
		if (sha1 == null) {
			sha1 = new ActiveXObject("System.Security.Cryptography.SHA1CryptoServiceProvider");
		}
		return crypto(sha1, bytes);
	};
	
	// sha256
	this.sha256 = function sha256(bytes) {
		if (sha256 == null) {
			sha256 = new ActiveXObject("System.Security.Cryptography.SHA256CryptoServiceProvider");
		}
		return crypto(sha256, bytes);
	};
}).call(EncodeUtility);
