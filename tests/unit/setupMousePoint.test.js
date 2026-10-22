const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { setupMousePoint } = require("../../app/office/setupMousePoint");

describe("named setup control frame positioning", () => {
  const local = { point: { x: 5, y: 7 }, viewport: { width: 100, height: 80 } };
  it("projects only through the actual parent chain and keeps probes passive", async () => {
    let calls = 0;
    const root = { url: "https://tenant.sharepoint.com/file", executeJavaScript: async (source, gesture) => {
      assert.equal(gesture, false);
      assert.ok(source.includes('"childUrl":"https://euc-excel.officeapps.live.com/x/edit"'));
      calls++;
      return { point: { x: 25, y: 37 }, viewport: { width: 800, height: 600 } };
    } };
    const editor = { url: "https://euc-excel.officeapps.live.com/x/edit", parent: root, name: "editor" };
    assert.deepEqual(await setupMousePoint({ mainFrame: root }, editor, local, () => true), { x: 25, y: 37 });
    assert.equal(calls, 1);
  });

  it("rejects detached ancestry, navigation, invalid points and frames outside the contents", async () => {
    const root = { url: "https://tenant.sharepoint.com/file" };
    const editor = { url: "https://euc-excel.officeapps.live.com/x/edit", parent: root };
    const contents = { mainFrame: root };
    for (const point of [null, { point: { x: -1, y: 0 }, viewport: { width: 1, height: 1 } }, { point: { x: 0, y: NaN }, viewport: { width: 1, height: 1 } }]) {
      assert.equal(await setupMousePoint(contents, editor, point, () => true), null);
    }
    assert.equal(await setupMousePoint(contents, editor, local, () => false), null);
    root.detached = true;
    assert.equal(await setupMousePoint(contents, editor, local, () => true), null);
    root.detached = false;
    root.executeJavaScript = async () => { editor.url += "?navigated"; return local; };
    assert.equal(await setupMousePoint(contents, editor, local, () => true), null);
    editor.parent = null;
    assert.equal(await setupMousePoint(contents, editor, local, () => true), null);
  });
});
