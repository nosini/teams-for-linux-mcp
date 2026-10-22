'use strict';

const os = require('node:os');

/**
 * Show the home directory as "~" in text meant for logs. The path stays usable
 * in a shell or file picker without writing the account name to the log.
 * @param {*} text - Text that may contain absolute paths
 * @param {string} [home] - Home directory to replace
 * @returns {*} The text with every home-directory prefix shown as "~"
 */
function tildify(text, home = os.homedir()) {
	if (typeof text !== 'string' || !home || home === '/') return text;
	let result = '';
	let index = 0;
	for (let found = text.indexOf(home); found !== -1; found = text.indexOf(home, index)) {
		const next = text[found + home.length];
		// Only whole path segments: "/home/al" must not match "/home/alice".
		const boundary = next === undefined || next === '/' || next === '\\' || /[\s'"`,)]/.test(next);
		result += text.slice(index, found) + (boundary ? '~' : home);
		index = found + home.length;
	}
	return result + text.slice(index);
}

module.exports = { tildify };
