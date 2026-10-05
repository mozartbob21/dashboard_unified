const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname,'../static/map-basemap.js'),'utf8');
function fixture(config={}){
  const timers=new Map(), notes=[], removed=[], controls=[];let next=1;
  const map={removeLayer:layer=>removed.push(layer)};
  const layer={events:{},redraws:0,on(events){Object.assign(this.events,events);return this;},
    off(){this.events={};return this;},fire(event){this.events[event]?.();},
    addTo(target){assert.equal(target,map);this.fire('loading');return this;},
    redraw(){this.redraws++;this.fire('loading');return this;}};
  let url,options;
  const context=vm.createContext({window:{},
    document:{getElementById:()=>({textContent:JSON.stringify(config)})},
    L:{tileLayer:(u,o)=>{url=u;options=o;return layer;},
      control:opts=>{const c={options:opts,addTo(){this.node=this.onAdd();controls.push(this);},remove(){this.removed=true;}};return c;},
      DomUtil:{create:(tag,className,parent)=>{const el={tag,className,children:[],setAttribute(k,v){this[k]=v;}};parent?.children.push(el);return el;}},
      DomEvent:{disableClickPropagation(){},disableScrollPropagation(){}}},
    setTimeout:fn=>{const id=next++;timers.set(id,fn);return id;},clearTimeout:id=>timers.delete(id)});
  vm.runInContext(source,context);
  const control=context.window.NeuronaBasemap.create(map,{onStatus:note=>notes.push(note)});
  return {layer,control,map,notes,timers,removed,controls,url,options,timeout(){const pending=[...timers.values()];timers.clear();pending.forEach(fn=>fn());}};
}
test('Yandex uses official endpoint, spherical projection and safely encoded browser key',()=>{
  const key='synthetic&projection=wrong#<x>{z}';
  const f=fixture({provider:'yandex',api_key:key});
  const u=new URL(f.url);
  assert.equal(u.origin,'https://tiles.api-maps.yandex.ru');
  assert.equal(u.pathname,'/v1/tiles/');
  assert.equal(u.searchParams.get('projection'),'web_mercator');
  assert.equal(u.searchParams.get('apikey'),key);
  for(const axis of ['x','y','z'])assert.equal(u.searchParams.get(axis),'{'+axis+'}');
  assert.equal(u.searchParams.get('lang'),'ru_RU');assert.equal(u.searchParams.get('l'),'map');
  assert.equal(f.options.referrerPolicy,'strict-origin-when-cross-origin');
  assert.equal(f.options.crossOrigin,'anonymous');assert.equal(f.options.keepBuffer,0);
  assert.match(f.options.attribution,/Яндекс Карты/);
  const logo=f.controls[0];assert.equal(logo.options.position,'bottomleft');
  const link=logo.node.children[0];assert.equal(link.href,'https://yandex.ru/maps/');
  assert.equal(link.children[0].src,'/static/vendor/yandex-map-logo.svg');
  f.control.destroy();assert.equal(logo.removed,true);
});
test('Yandex failure is actionable and does not expose the key or replace the provider',()=>{
  const f=fixture({provider:'yandex',api_key:'synthetic-private-test-key'});
  f.layer.fire('tileerror');f.layer.fire('load');
  assert.match(f.notes.at(-1),/ключ Tiles API/);
  assert.doesNotMatch(f.notes.at(-1),/synthetic-private-test-key/);
  f.control.retry();assert.equal(f.layer.redraws,1);assert.equal(f.removed.length,0);
  f.layer.fire('tileload');assert.equal(f.notes.at(-1),'');
});
test('server configuration warning survives successful fallback tile loads',()=>{
  const f=fixture({provider:'osm',notice:'Ключ Яндекс Карт не настроен.'});
  assert.equal(f.notes.at(-1),'Ключ Яндекс Карт не настроен.');
  f.layer.fire('tileerror');f.layer.fire('load');assert.match(f.notes.at(-1),/недоступна/);
  f.control.retry();f.layer.fire('tileload');assert.equal(f.notes.at(-1),'Ключ Яндекс Карт не настроен.');
  assert.equal(f.controls.length,0);
});
test('EDDS retry retains persistent configuration notice',()=>{
  const html=fs.readFileSync(path.join(__dirname,'../services/edds/dashboard.html'),'utf8');
  const start=html.indexOf("  const tr = document.getElementById('tile-retry');");
  assert.ok(start>=0);
  const handler=html.slice(start,html.indexOf('\n  });',start)+6);
  let click,rendered;
  const ctx=vm.createContext({tileNote:'Ключ Яндекс Карт не настроен.',MAP:{},
    document:{getElementById:()=>({addEventListener:(event,fn)=>{click=fn;}})},
    useTiles(){},renderMapStatus(){rendered=ctx.tileNote;}});
  vm.runInContext(handler,ctx);click({preventDefault(){}});
  assert.equal(rendered,'Ключ Яндекс Карт не настроен.');
});
test('unsupported provider and missing browser config never request arbitrary hosts',()=>{
  for(const config of [null,{provider:'https://invalid.example',api_key:'x'},{provider:'yandex',api_key:''}]){
    const f=fixture(config);assert.match(f.url,/^https:\/\/tile.openstreetmap.org\//);
    assert.equal(f.controls.length,0);
  }
});
test('documented OSM endpoint with attribution, origin-only Referer and no cache busting',()=>{
  const f=fixture();assert.equal(f.url,'https://tile.openstreetmap.org/{z}/{x}/{y}.png');
  assert.equal(f.options.referrerPolicy,'strict-origin-when-cross-origin');
  assert.equal(f.options.crossOrigin,'anonymous');assert.match(f.options.attribution,/openstreetmap.org\/copyright/);
  assert.equal(f.options.maxZoom,18);
});
test('all failed tiles do not become successful at Leaflet load event',()=>{
  const f=fixture();f.layer.fire('tileerror');f.layer.fire('tileerror');f.layer.fire('load');
  assert.match(f.notes.at(-1),/недоступна/);assert.equal(f.timers.size,0);
});
test('silent timeout reports failure and a late real tile clears the warning',()=>{
  const f=fixture();f.timeout();assert.match(f.notes.at(-1),/недоступна/);
  f.layer.fire('tileload');assert.equal(f.notes.at(-1),'');assert.equal(f.timers.size,0);
});
test('new viewport is monitored after an earlier successful load',()=>{
  const f=fixture();f.layer.fire('tileload');f.layer.fire('load');
  f.layer.fire('loading');f.layer.fire('tileerror');f.layer.fire('load');assert.match(f.notes.at(-1),/недоступна/);
});
test('retry redraws only basemap and recovers after failures',()=>{
  const f=fixture();f.timeout();f.control.retry();assert.equal(f.notes.at(-1),'');
  assert.equal(f.layer.redraws,1);assert.equal(f.removed.length,0);
  f.layer.fire('tileload');f.layer.fire('load');assert.equal(f.notes.at(-1),'');assert.equal(f.timers.size,0);
});
test('destroy cancels pending status and removes only the basemap layer',()=>{
  const f=fixture();f.control.destroy();f.timeout();f.layer.fire('tileerror');f.control.retry();
  assert.equal(f.notes.length,0);assert.equal(f.layer.redraws,0);assert.deepEqual(f.removed,[f.layer]);
});
test('both production maps use the shared basemap and retire old 2GIS tiles',()=>{
  for(const name of ['services/edds/dashboard.html','static/mingkh-water-map.js']){
    const code=fs.readFileSync(path.join(__dirname,'..',name),'utf8');
    assert.match(code,/NeuronaBasemap\.create\(MAP/);assert.doesNotMatch(code,/maps\.2gis\.com/);
  }
});
