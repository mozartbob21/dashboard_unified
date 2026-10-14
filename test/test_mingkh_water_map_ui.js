/* Run with Node; synthetic status/data only, no browser or portal requests. */
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'static', 'mingkh-water-map.js'), 'utf8');
const statusSource = source.slice(source.indexOf("const refreshButton = document.getElementById('live-refresh');"), source.indexOf('\nif (refreshButton) {'));

async function statusTests() {
  const label = {textContent:'', dataset:{}};
  let state = {busy:false, automatic_enabled:true, interval_minutes:120, dataset_version:'v1'};
  let fetches = 0, loads = 0, timerDelay = null;
  const context = vm.createContext({
    document:{hidden:false, getElementById:id => id === 'dataset-update-status' ? label : null},
    fetch: async (url, options) => {
      assert.equal(url, '/mingkh/api/water-map/status');
      assert.equal(options.method, undefined, 'Readers never submit update requests');
      fetches++;
      return {ok:true, json:async () => state};
    },
    setTimeout:(_, delay) => {timerDelay = delay; return 1;}, clearTimeout:() => {},
    archiveRefreshing:false, syncArchiveControls:() => {}, datasetReady:true,
    loadedDatasetVersion:'v1', start:async () => {loads++; context.loadedDatasetVersion=state.dataset_version; return true;},
  });
  vm.runInContext(statusSource, context);
  await vm.runInContext('refreshStatus()', context);
  assert.equal(fetches, 1, 'Ordinary readers get live status without a refresh button');
  assert.match(label.textContent, /каждые 2 ч/);
  assert.equal(loads, 0);
  assert.equal(timerDelay, 60000);
  state = {...state, busy:true};
  await vm.runInContext('refreshStatus()', context);
  assert.equal(label.dataset.state, 'busy');
  assert.equal(timerDelay, 3000);
  state = {...state, busy:false, dataset_version:'v2'};
  await vm.runInContext('refreshStatus()', context);
  assert.equal(loads, 1, 'A changed dataset version refreshes the open map');
  state = {...state, error:'Admin: save credentials in users settings'};
  await vm.runInContext('refreshStatus()', context);
  assert.match(label.textContent, /Показаны сохранённые данные/);
  assert.doesNotMatch(label.textContent, /credentials|settings|Admin/);
  state = {...state, error:null, automatic_enabled:false};
  await vm.runInContext('refreshStatus()', context);
  assert.match(label.textContent, /приостановлено/);
  assert.doesNotMatch(label.textContent, /каждые/);
  state = {...state, automatic_enabled:true, interval_minutes:45};
  await vm.runInContext('refreshStatus()', context);
  assert.match(label.textContent, /каждые 45 мин/);
  context.document.hidden = true;
  const before = fetches;
  await vm.runInContext('refreshStatus()', context);
  assert.equal(fetches, before, 'Hidden pages pause polling until visible');
}

async function preserveFilterTests() {
  const startSource = source.slice(source.indexOf('async function start() {'), source.indexOf('const importForm ='));
  const selected = new Set(['selected district']);
  const form = {from:'2026-09-01', to:'2026-09-15', cats:new Set(['category'])};
  const area = {kind:'circle', lat:55.7, lon:37.6, r:500};
  const checkbox = {checked:true};
  let builds=0, applies=0;
  const context = vm.createContext({
    datasetReady:true, datasetLoading:false, loadedDatasetVersion:'v1', importForm:null,
    document:{querySelector:() => null, querySelectorAll:() => [], getElementById:() => ({hidden:false})},
    fetch:async () => ({ok:true, json:async () => ({rows:[]}), headers:{get:() => 'v2'}}),
    loadData:() => true, msBuild:() => {builds++;}, msBox:() => ({querySelectorAll:() => [{dataset:{o:'other district'}, querySelector:() => checkbox}]}),
    MS:{omsu:{sel:selected}, fact:{sel:null}, org:{sel:null}}, F:form, geoBox:area,
    buildGroups:() => {}, applyFilters:() => {applies++;},
    quick:() => {throw new Error('Manual dates were reset');},
    wire:() => {throw new Error('Existing listeners were reattached');},
    initMap:() => {throw new Error('Map viewport was reset');}, setUpdateStatus:() => {},
  });
  vm.runInContext(startSource, context);
  assert.equal(await vm.runInContext('start()', context), true);
  assert.equal(builds, 3);
  assert.equal(applies, 1);
  assert.equal(context.MS.omsu.sel, selected);
  assert.equal(context.F, form);
  assert.equal(context.geoBox, area);
  assert.equal(context.F.from, '2026-09-01');
  assert.equal(context.loadedDatasetVersion, 'v2');
}

(async () => {await statusTests(); await preserveFilterTests(); console.log('Map status and filter preservation tests passed');})().catch(error => {console.error(error); process.exitCode=1;});
