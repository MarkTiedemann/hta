var localStorage = (function () {
	var sh = new ActiveXObject("WScript.Shell");
	var root = "HKCU\\Software\\localStorage\\";
	return {
		getItem: function (key) {
			try {
				return sh.regRead(root + key);
			} catch (e) {
				return null;
			}
		},
		setItem: function (key, value) {
			sh.regWrite(root + key, value);
		},
		removeItem: function (key) {
			try {
				sh.regDelete(root + key);
			} catch (e) {}
		},
		clear: function () {
			try {
				sh.regDelete(root);
			} catch (e) {}
		}
	}
})();
