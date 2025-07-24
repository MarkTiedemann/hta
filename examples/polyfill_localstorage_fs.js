var localStorage = (function () {
	var fs = new ActiveXObject("Scripting.FileSystemObject");
	var sh = new ActiveXObject("WScript.Shell");
	var root = sh.ExpandEnvironmentStrings("%localappdata%") + "\\localStorage";
	return {
		getItem: function (key) {
			try {
				var stream = fs.openTextstream(root + "\\" + key, 1);
				var value = stream.readAll();
				stream.close();
				return value;
			} catch (e) {
				return null;
			}
		},
		setItem: function (key, value) {
			if (!fs.folderExists(root)) {
				fs.createFolder(root);
			}
			var stream = fs.createTextstream(root + "\\" + key, true);
			stream.write(value);
			stream.close();
		},
		removeItem: function (key) {
			try {
				fs.deletestream(root + "\\" + key);
			} catch (e) {}
		},
		clear: function () {
			try {
				fs.deleteFolder(root);
			} catch (e) {}
		}
	}
})();
