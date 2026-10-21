'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { tildify } = require('../../app/utils/homePath');

describe('tildify', () => {
	it('shows the home directory as ~ in paths and error messages', () => {
		assert.equal(tildify('/home/alice/.config/teams-for-linux/mcp/bridge.sock', '/home/alice'), '~/.config/teams-for-linux/mcp/bridge.sock');
		assert.equal(tildify('/home/alice', '/home/alice'), '~');
		assert.equal(tildify("EACCES: permission denied, mkdir '/home/alice/.config/x' (copy of /home/alice/y)", '/home/alice'), "EACCES: permission denied, mkdir '~/.config/x' (copy of ~/y)");
	});

	it('replaces only whole path segments and leaves other text unchanged', () => {
		assert.equal(tildify('/home/alice2/file /home/alice', '/home/alice'), '/home/alice2/file ~');
		assert.equal(tildify('/srv/mcp/bridge.sock', '/home/alice'), '/srv/mcp/bridge.sock');
		assert.equal(tildify('/srv/mcp', '/'), '/srv/mcp');
		assert.equal(tildify(undefined, '/home/alice'), undefined);
	});
});
