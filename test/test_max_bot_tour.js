'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { Tour, STEPS, placement } = require('../static/max-bot-tour.js');

function fixture(search = '') {
  const nodes = new Map(), listeners = {}, timers = [], events = [];
  let document;
  class Node {
    constructor(id) { this.id = id; this.hidden = false; this.disabled = false; this.isConnected = true; this.textContent = ''; this.style = {}; this.listeners = {}; this.attributes = {}; this.classes = new Set(); this.classList = { add: v => this.classes.add(v), remove: v => this.classes.delete(v), toggle: (v, enabled) => enabled ? this.classes.add(v) : this.classes.delete(v) }; }
    addEventListener(name, fn) { this.listeners[name] = fn; }
    getClientRects() { return this.hidden ? [] : [this.getBoundingClientRect()]; }
    getBoundingClientRect() { return this.id === 'max-tour-card' ? { left: 0, top: 0, width: 370, height: 360 } : { left: 570, top: 150, right: 910, bottom: 245, width: 340, height: 95 }; }
    querySelectorAll() { return [new Node('dot1'), new Node('dot2'), new Node('dot3'), new Node('dot4')]; }
    setAttribute(key, value) { this.attributes[key] = value; }
    contains(node) { return node?.id.startsWith('max-tour-'); }
    focus() { document.activeElement = this; events.push('focus:' + this.id); }
    scrollIntoView() { events.push('scroll:' + this.id); }
  }
  const node = id => { if (!nodes.has(id)) nodes.set(id, new Node(id)); return nodes.get(id); };
  const launch = node('launch'); const dialogs = [];
  document = {
    hidden: false, activeElement: launch, body: node('body'),
    getElementById: node,
    querySelector: selector => selector === '[data-max-tour-start]' ? launch : null,
    querySelectorAll: selector => selector === '[data-max-tour-start]' ? [launch] : dialogs,
    addEventListener: (key, fn) => { listeners[key] = fn; },
  };
  const win = { document, location: { search }, scrollY: 0, innerWidth: 1100, innerHeight: 800,
    addEventListener: (key, fn) => { listeners[key] = fn; },
    scrollTo: () => events.push('scroll-page'), requestAnimationFrame: () => 1, cancelAnimationFrame() {},
    setTimeout: fn => { timers.push(fn); return timers.length; }, clearTimeout() {},
    fetch: () => { throw new Error('Tutorial must not call API'); },
  };
  node('max-bot-code-panel').hidden = true; node('max-bot-pending').hidden = true;
  node('max-tour-layer').hidden = true;
  return { win, node, launch, dialogs, timers, listeners, events, tour: new Tour(win) };
}

test('tutorial never opens automatically without an explicit tour link', () => {
  const f = fixture(); f.tour.init(); assert.equal(f.tour.open, false); assert.equal(f.node('max-tour-layer').hidden, true);
  f.launch.listeners.click(); assert.equal(f.tour.open, true); assert.equal(f.node('max-tour-count').textContent, 'Шаг 1 из 4');
});
test('four steps, back, repeat, and finish do not execute any binding action', () => {
  const f = fixture('?tour=1'); f.tour.init();
  for (let i = 0; i < 4; i++) { assert.equal(f.tour.index, i); assert.equal(f.node('max-tour-title').textContent, STEPS[i].title); f.node('max-tour-next').listeners.click(); }
  assert.equal(f.tour.open, false); assert.equal(f.win.document.activeElement.id, 'max-bot-create');
  f.launch.listeners.click(); f.node('max-tour-next').listeners.click(); f.node('max-tour-prev').listeners.click(); assert.equal(f.tour.index, 0);
});
test('Escape closes walkthrough and returns visible focus to launcher', () => {
  const f = fixture('?tour=1'); f.tour.init(); let prevented = false;
  f.listeners.keydown({ key: 'Escape', preventDefault() { prevented = true; } });
  assert.equal(prevented, true); assert.equal(f.tour.open, false); assert.equal(f.win.document.activeElement, f.launch);
  assert.ok(f.events.includes('scroll:launch')); assert.equal(f.win.document.body.classes.has('max-tour-open'), false);
});
test('visible controls are highlighted, unavailable controls have stable guide fallback', () => {
  const f = fixture(); f.tour.init(); f.launch.listeners.click(); assert.equal(f.tour.target().id, 'max-bot-create');
  f.node('max-bot-create').disabled = true; assert.equal(f.tour.target().id, 'max-guide-code');
  f.tour.show(1); assert.equal(f.tour.target().id, 'max-guide-send');
  f.node('max-bot-code-panel').hidden = false; assert.equal(f.tour.target().id, 'max-bot-code-panel');
  f.tour.show(2); assert.equal(f.tour.target().id, 'max-guide-confirm');
  f.node('max-bot-pending').hidden = false; assert.equal(f.tour.target().id, 'max-bot-pending');
});
test('another modal delays a requested tour and is not closed by the tutorial', () => {
  const f = fixture('?tour=1'); const modal = f.node('release-modal'); f.dialogs.push(modal); f.tour.init();
  assert.equal(f.tour.open, false); assert.equal(f.timers.length, 1); assert.equal(modal.hidden, false);
  modal.hidden = true; f.timers.shift()(); assert.equal(f.tour.open, true);
});
test('tutorial uses configured bot name as text and never builds an external URL', () => {
  const f = fixture(); f.node('max-bot-name').textContent = 'Ваш бот в MAX: <img src=x>'; f.tour.init(); f.launch.listeners.click(); f.tour.show(1);
  assert.match(f.node('max-tour-hint').textContent, /<img src=x>/); assert.equal(f.node('max-tour-hint').children, undefined);
});
test('card placement stays within desktop and mobile viewports', () => {
  for (const [width, height, cardWidth, cardHeight] of [[1366, 900, 370, 360], [360, 780, 328, 410], [320, 568, 288, 470]]) {
    const result = placement({ left: 24, right: width - 24, top: 130, bottom: 210 }, { width: cardWidth, height: cardHeight }, { width, height });
    assert.ok(result.x >= 16); assert.ok(result.x + cardWidth <= width - 16); assert.ok(result.y >= 16); assert.ok(result.y + cardHeight <= height - 16);
  }
});
test('tour keyboard arrows apply only while focus is in the guide', () => {
  const f = fixture('?tour=1'); f.tour.init(); f.win.document.activeElement = f.launch;
  f.listeners.keydown({ key: 'ArrowRight', preventDefault() { throw new Error('Outside guide'); } }); assert.equal(f.tour.index, 0);
  f.win.document.activeElement = f.node('max-tour-next'); f.listeners.keydown({ key: 'ArrowRight', preventDefault() {} }); assert.equal(f.tour.index, 1);
});
