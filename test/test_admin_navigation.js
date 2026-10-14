// Run: node --test test/test_admin_navigation.js
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const {createController, sections} = require(process.env.ADMIN_NAVIGATION_SCRIPT || path.join(__dirname, '../static/admin-navigation.js'));

class Element {
  constructor(dataset = {}) {
    this.dataset = dataset;
    this.hidden = false;
    this.attributes = new Map();
    this.listeners = new Map();
    this.focusCalls = [];
    this.scrollCalls = [];
  }
  setAttribute(name, value) { this.attributes.set(name, String(value)); }
  removeAttribute(name) { this.attributes.delete(name); }
  addEventListener(name, callback) {
    const listeners = this.listeners.get(name) || [];
    listeners.push(callback);
    this.listeners.set(name, listeners);
  }
  focus(options) { this.focusCalls.push(options); }
  scrollIntoView(options) { this.scrollCalls.push(options); }
  reset() { throw new Error('Navigation must not reset a form'); }
  replaceChildren() { throw new Error('Navigation must not replace mounted content'); }
}

function fixture({hash = '', initial = 'users'} = {}) {
  const keys = Object.keys(sections);
  const panels = keys.map(adminPanel => new Element({adminPanel}));
  const links = keys.map(adminLink => new Element({adminLink}));
  const shell = new Element({initialSection: initial});
  shell.querySelectorAll = selector => {
    if (selector === '[data-admin-panel]') return panels;
    if (selector === '[data-admin-link]') return links;
    throw new Error('Unexpected selector: ' + selector);
  };
  const user = new Element({dirty: 'false'});
  user.elements = {username: {value: 'selected-user'}, email: {value: 'unsaved@example.test'},
    password: {value: 'synthetic-user-password'}, is_active: {checked: false},
    modules: [{value: 'mingkh', checked: true}, {value: 'zips', checked: false}]};
  const integration = new Element({dirty: 'false'});
  integration.elements = {service: {value: 'zip_cloud'}, username: {value: 'synthetic-cloud-user'},
    password: {value: 'synthetic-cloud-password'}, source_url: {value: 'https://cloud.mosreg.ru/apps/files/?dir=/&fileid=10968596'},
    folder_path: {value: '/Shared/Unsaved'}};
  const nodes = new Map([['adminConsole', shell], ['adminTitle', new Element()], ['adminDescription', new Element()],
    ['adminMain', new Element()], ['userForm', user], ['integrationForm', integration]]);
  const emitted = [], events = new Map(), historyEntries = [], location = {hash};
  const document = {body: new Element(), title: '', getElementById: id => nodes.get(id) || null,
    dispatchEvent: event => emitted.push({section: event.detail.section, visible: panels.filter(panel => !panel.hidden).map(panel => panel.dataset.adminPanel), type: event.type})};
  const env = {document, location,
    CustomEvent: class {constructor(type, options) {this.type = type; this.detail = options.detail;}},
    history: {pushState(state, title, url) {historyEntries.push({state, title, url}); location.hash = url;}},
    addEventListener(name, callback) {const list = events.get(name) || []; list.push(callback); events.set(name, list);}};
  const controller = createController(env);
  function emit(name, event = {}) {for (const callback of events.get(name) || []) callback(event);}
  function click(section, options = {}) {
    const link = links.find(item => item.dataset.adminLink === section);
    const event = {button: 0, metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, currentTarget: link,
      prevented: false, preventDefault() {this.prevented = true;}, ...options};
    for (const callback of link.listeners.get('click') || []) callback(event);
    return event;
  }
  return {controller, document, panels, links, events, emitted, historyEntries, location, user, integration,
    node: id => nodes.get(id), emit, click};
}

function assertSelected(f, key) {
  assert.deepEqual(f.panels.filter(panel => !panel.hidden).map(panel => panel.dataset.adminPanel), [key]);
  assert.deepEqual(f.links.filter(link => link.attributes.get('aria-current') === 'page').map(link => link.dataset.adminLink), [key]);
  assert.equal(f.document.body.dataset.adminSection, key);
  assert.equal(f.node('adminTitle').textContent, sections[key].title);
  assert.equal(f.node('adminDescription').textContent, sections[key].description);
  assert.match(f.document.title, /Админ-панель — Нейрона$/);
}

test('one panel is selected on startup and section changes dispatch after visibility is updated', () => {
  const f = fixture();
  f.controller.start();
  assertSelected(f, 'users');
  for (const key of ['statistics', 'feedback', 'integrations', 'users']) {
    assert.equal(f.click(key).prevented, true);
    assertSelected(f, key);
    assert.equal(f.location.hash, '#' + key);
    assert.deepEqual(f.emitted.at(-1), {type: 'admin:sectionchange', section: key, visible: [key]});
  }
  assert.equal(f.node('adminTitle').focusCalls.length, 4);
  assert.equal(f.node('adminMain').scrollCalls.length, 4);
});

test('deep links take precedence and legacy feedback starts in feedback without a hash', () => {
  for (const [options, expected] of [
    [{hash: '#feedback'}, 'feedback'], [{initial: 'feedback'}, 'feedback'],
    [{initial: 'feedback', hash: '#integrations'}, 'integrations'],
    [{initial: 'unknown', hash: '#unknown'}, 'users']
  ]) {
    const f = fixture(options); f.controller.start(); assertSelected(f, expected);
    assert.equal(f.historyEntries.length, 0);
    assert.equal(f.node('adminTitle').focusCalls.length, 0);
  }
});

test('browser Back, Forward, direct hashes and an empty hash restore the requested section', () => {
  const f = fixture(); f.controller.start(); f.click('statistics'); f.click('feedback');
  f.location.hash = '#statistics'; f.emit('popstate'); f.emit('hashchange'); assertSelected(f, 'statistics');
  f.location.hash = '#feedback'; f.emit('popstate'); f.emit('hashchange'); assertSelected(f, 'feedback');
  f.location.hash = '#integrations'; f.emit('hashchange'); assertSelected(f, 'integrations');
  f.location.hash = ''; f.emit('popstate'); f.emit('hashchange'); assertSelected(f, 'users');
  assert.equal(f.historyEntries.length, 2, 'history restoration must not push another history entry');
});

test('the skip-to-content anchor keeps the current section open', () => {
  const f = fixture(); f.controller.start(); f.click('feedback');
  f.location.hash = '#adminMain'; f.emit('hashchange');
  assertSelected(f, 'feedback');
});

test('both dirty forms and their mounted controls survive navigation through all four sections', () => {
  const f = fixture(); f.controller.start();
  f.user.dataset.dirty = 'true'; f.integration.dataset.dirty = 'true';
  const userControls = f.user.elements, integrationControls = f.integration.elements;
  const before = JSON.stringify([userControls, integrationControls]);
  for (const key of ['statistics', 'feedback', 'integrations', 'users']) f.click(key);
  assert.equal(f.node('userForm'), f.user); assert.equal(f.node('integrationForm'), f.integration);
  assert.equal(f.user.elements, userControls); assert.equal(f.integration.elements, integrationControls);
  assert.equal(JSON.stringify([userControls, integrationControls]), before);
  assert.equal(f.user.dataset.dirty, 'true'); assert.equal(f.integration.dataset.dirty, 'true');
});

test('modifier clicks and non-primary buttons retain native browser behaviour', () => {
  for (const options of [{ctrlKey: true}, {metaKey: true}, {shiftKey: true}, {altKey: true}, {button: 1}, {button: 2}]) {
    const f = fixture(); f.controller.start();
    assert.equal(f.click('feedback', options).prevented, false);
    assertSelected(f, 'users'); assert.equal(f.historyEntries.length, 0);
  }
});

test('beforeunload warns only while either form has unsaved changes, even when its panel is hidden', () => {
  const f = fixture(); f.controller.start(); f.click('statistics');
  function unload() {const event = {prevented: false, preventDefault() {this.prevented = true;}}; f.emit('beforeunload', event); return event;}
  assert.deepEqual(unload().prevented, false);
  for (const form of [f.user, f.integration]) {
    form.dataset.dirty = 'true'; const event = unload();
    assert.equal(event.prevented, true); assert.equal(event.returnValue, '');
    form.dataset.dirty = 'false'; assert.equal(unload().prevented, false);
  }
});

test('repeated start and selecting the current section do not add handlers or history entries', () => {
  const f = fixture(); f.controller.start(); f.controller.start();
  for (const listeners of f.events.values()) assert.equal(listeners.length, 1);
  for (const link of f.links) assert.equal(link.listeners.get('click').length, 1);
  f.click('users'); assert.equal(f.historyEntries.length, 0);
  f.click('feedback'); assert.equal(f.historyEntries.length, 1);
});
