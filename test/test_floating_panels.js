const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const script = fs.readFileSync(require.resolve('../static/appearance-menu.js'), 'utf8');
function events() {
  const listeners = new Map();
  return {
    listeners,
    addEventListener(name, fn) {
      if (!listeners.has(name)) listeners.set(name, new Set());
      listeners.get(name).add(fn);
    },
    removeEventListener(name, fn) { listeners.get(name)?.delete(fn); },
    dispatch(name) { listeners.get(name)?.forEach(fn => fn()); }
  };
}
function node() {
  return {
    parentNode: null, nextSibling: null, children: [],
    appendChild(child) {
      if (child.parentNode) child.parentNode.children = child.parentNode.children.filter(x => x !== child);
      this.children.push(child); child.parentNode = this;
    },
    insertBefore(child, next) {
      this.appendChild(child);
      if (next) {
        this.children.pop();
        this.children.splice(this.children.indexOf(next), 0, child);
      }
    },
    remove() {
      if (this.parentNode) this.parentNode.children = this.parentNode.children.filter(x => x !== this);
      this.parentNode = null;
    }
  };
}
function fixture({width = 320, height = 568, native = true, viewport = null} = {}) {
  const window = {...events(), innerHeight: height, visualViewport: viewport && {...events(), ...viewport}};
  const document = {
    documentElement: {dataset: {}, clientWidth: width}, body: node(),
    readyState: 'loading', addEventListener() {}, createElement: node
  };
  const parent = node(), sibling = node(), properties = {}, attributes = {};
  const panel = {...node(), hidden: true, scrollHeight: 450, opened: false,
    style: {setProperty(key, value) { properties[key] = parseFloat(value); }},
    setAttribute(key, value) { attributes[key] = value; },
    matches() { return this.opened; }
  };
  if (native) {
    panel.showPopover = function () { this.opened = true; };
    panel.hidePopover = function () { this.opened = false; };
  }
  parent.appendChild(panel); parent.appendChild(sibling); panel.nextSibling = sibling;
  let rect = {top: 50, bottom: 96, right: 308};
  const anchor = {getBoundingClientRect: () => rect};
  vm.runInNewContext(script, {window, document, localStorage: {getItem() {return null;}}, setInterval() {}});
  return {
    controller: window.NeuronaAppearance.floatingPanel(panel, anchor),
    panel, parent, sibling, document, properties, attributes, window,
    move(next) { rect = next; }
  };
}
function assertBounds(f, {width = 320, height = 568, left = 0, top = 0} = {}) {
  const p = f.properties;
  assert.ok(p['--menu-x'] >= left + 12);
  assert.ok(p['--menu-y'] >= top + 12);
  assert.ok(p['--menu-x'] + p['--menu-width'] <= left + width - 12);
  assert.ok(p['--menu-y'] + p['--menu-height'] <= top + height - 12);
}
test('narrow phone menu fits the viewport and listeners are cleaned after repeated opens', () => {
  const f = fixture();
  f.controller.open(); f.controller.open();
  assertBounds(f); assert.equal(f.panel.hidden, false); assert.equal(f.panel.opened, true);
  assert.equal(f.window.listeners.get('resize').size, 1);
  f.controller.close();
  assert.equal(f.panel.hidden, true); assert.equal(f.panel.opened, false);
  assert.equal(f.window.listeners.get('resize').size, 0);
  assert.equal(f.window.listeners.get('scroll').size, 0);
});
test('short landscape menu scrolls within the screen, even when the anchor is offscreen', () => {
  const f = fixture({width: 740, height: 375});
  f.move({top: 230, bottom: 276, right: 728});
  f.controller.open(); assertBounds(f, {width: 740, height: 375});
  f.move({top: -150, bottom: -104, right: 800});
  f.window.dispatch('scroll');
  assertBounds(f, {width: 740, height: 375});
});
test('visual viewport offsets and resize are respected', () => {
  const f = fixture({viewport: {width: 300, height: 300, offsetLeft: 20, offsetTop: 100}});
  f.controller.open(); assertBounds(f, {width: 300, height: 300, left: 20, top: 100});
  Object.assign(f.window.visualViewport, {width: 280, height: 220, offsetTop: 150});
  f.window.visualViewport.dispatch('resize');
  assertBounds(f, {width: 280, height: 220, left: 20, top: 150});
  f.controller.close();
  assert.equal(f.window.visualViewport.listeners.get('resize').size, 0);
});
test('older browsers use a body portal and restore the panel in its original order', () => {
  const f = fixture({native: false});
  f.controller.open(); assertBounds(f);
  assert.notEqual(f.panel.parentNode, f.parent);
  assert.equal(f.panel.parentNode.parentNode, f.document.body);
  assert.equal(f.attributes.popover, undefined);
  f.controller.close();
  assert.equal(f.panel.parentNode, f.parent);
  assert.deepEqual(f.parent.children, [f.panel, f.sibling]);
  assert.equal(f.document.body.children.length, 0);
  f.controller.open(); f.controller.close();
  assert.equal(f.document.body.children.length, 0);
});
