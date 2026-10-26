const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require.resolve('../services/water_dashboard/scraper.py'), 'utf8');
const script = source.match(/WIDGETS_JS = """([\s\S]*?)"""/)[1].replaceAll('\\\\', '\\');
function widget(header, values, busy=false) {
 return {
  getAttribute:()=>busy?'true':'false',
  querySelector:s=>s==='.widget-header'?{innerText:header}:null,
  querySelectorAll:()=>values.map(([title,value])=>({innerText:title+'\n'+value,querySelector:s=>({innerText:s.includes('value')?value:title})}))
 };
}
function collect(widgets) {
 return JSON.parse(JSON.stringify(vm.runInNewContext('('+script+')()', {document:{querySelectorAll:()=>widgets}})));
}
test('meeting indicator reads its own title when header is empty',()=>{
 assert.deepEqual(collect([widget('',[['Проведено совещаний','37']])]).map(x=>[x.label,x.value]),[['Проведено совещаний','37']]);
});
test('labeled header remains authoritative, adjacent content and loading widgets are excluded',()=>{
 const values=collect([widget('Всего задач\nЕще 0',[['','100']]),widget('',[['','50']]),widget('',[['Проведено совещаний','999']],true)]);
 assert.equal(values.length,1);assert.equal(values[0].label,'Всего задач');assert.equal(values[0].value,'100');
});
