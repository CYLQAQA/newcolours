const fs=require('fs'),vm=require('vm');
const html=fs.readFileSync('index.html','utf8');
const allScripts=[...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)];
const pageScript=allScripts.filter(m=>m[1].length>1000).map(m=>m[1]).pop();
const elements=new Map();
let idCounter = 0;
function makeEl(tag) {
  const el = {
    id: 'el'+(idCounter++),
    tagName: (tag||'DIV').toUpperCase(),
    style:{}, classList:{add:()=>{},remove:()=>{},contains:()=>false,toggle:()=>{}},
    textContent:'', value:'', dataset:{}, innerHTML:'',
    appendChild(c){ if (c) { this.children.push(c); c.parentElement = this; } return c; },
    remove(){}, click(){}, addEventListener(){},
    querySelector(){return makeEl();},
    querySelectorAll(){return [];},
    closest(){return {style:{}};},
    checked:false, files:[], type:'text', hidden:false,
    children:[], parentElement:null,
  };
  return el;
}
const document={ getElementById:(id)=>{
  if(!elements.has(id)) { const el=makeEl(); el.id=id; elements.set(id, el); }
  return elements.get(id);
}, createElement:(tag)=>makeEl(tag), createTextNode:(t)=>({nodeValue:t,textContent:t,data:t,parentElement:null}), createDocumentFragment:()=>({appendChild(){}}), querySelectorAll:()=>[], addEventListener:()=>{}, body:{classList:{add:()=>{},remove:()=>{},toggle:()=>{},contains:()=>false}, appendChild(){}, remove(){}} };

const harness = pageScript + `
globalThis.__sim = function() {
  RAW.length = 0;
  RAW.push({CompanyName:'A', SKU:'X-A', Quantity:1, Total:10, Currency:'USD', Country:'US', OrderStatus:'Shipped', SaleType:'Sale', CreatedAt:new Date(2026,4,10)});
  RAW.push({CompanyName:'A', SKU:'Y-A', Quantity:1, Total:10, Currency:'USD', Country:'US', OrderStatus:'Shipped', SaleType:'Sale', CreatedAt:new Date(2026,4,11)});
  RAW.push({CompanyName:'A', SKU:'Z-A', Quantity:1, Total:10, Currency:'USD', Country:'US', OrderStatus:'Cancelled', SaleType:'Sale', CreatedAt:new Date(2026,4,12)});
  enrichRows(RAW);
  populateFilters();
  const cl = document.getElementById('cl-condition');
  const cbs = cl.querySelectorAll('input[type=checkbox]');
  return { count: cbs.length, init: [...cbs].map(c=>({v:c.dataset.value,on:c.checked})) };
};
globalThis.__test = function() {
  const cl = document.getElementById('cl-condition');
  const cbs = [...cl.querySelectorAll('input[type=checkbox]')];
  const damaged = cbs.find(c => c.dataset.value === 'Damaged');
  const states = [];
  states.push({step:'start', FILTERS_condition:[...FILTERS.condition], dom:cbs.map(c=>({v:c.dataset.value,on:c.checked}))});
  damaged.checked = false;
  onFilterChange({target: damaged});
  states.push({step:'unchecked Damaged', FILTERS_condition:[...FILTERS.condition], dom:cbs.map(c=>({v:c.dataset.value,on:c.checked})), active:activeRows().length});
  damaged.checked = true;
  onFilterChange({target: damaged});
  states.push({step:'re-checked Damaged', FILTERS_condition:[...FILTERS.condition], dom:cbs.map(c=>({v:c.dataset.value,on:c.checked})), active:activeRows().length});
  return states;
};
`;
const ctx=vm.createContext({document,window:{addEventListener:()=>{}},XLSX:{read:()=>({SheetNames:[]}),sheet_to_json:()=>[],utils:{}},Chart:function(){this.destroy=()=>{};this.resize=()=>{};},JSZip:function(){this.folder=()=>this;this.file=()=>this;this.generateAsync=async()=>Buffer.alloc(0);},URL:{createObjectURL:()=>'',revokeObjectURL:()=>{}},localStorage:{getItem:()=>null,setItem:()=>{}},alert:()=>{},setTimeout:fn=>fn(),setInterval:()=>{},clearInterval:()=>{},console});
ctx.globalThis=ctx;vm.runInContext(harness,ctx);
const init = ctx.__sim();
console.log('initial:', JSON.stringify(init.init));
const steps = ctx.__test();
for (const s of steps) {
  console.log(s.step);
  console.log('  FILTERS.condition:', s.FILTERS_condition);
  console.log('  DOM:', JSON.stringify(s.dom));
  if (s.active !== undefined) console.log('  activeRows():', s.active);
}