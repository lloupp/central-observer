// Regression test: an incorrect selector must not prevent the menu from initializing.
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
const elements=new Map();
function element(key){
 if(!elements.has(key))elements.set(key,{
  value:'',textContent:'',innerHTML:'',disabled:false,
  listeners:{},addEventListener(type,fn){this.listeners[type]=fn},
  classList:{toggle(){},add(){},remove(){},contains(){return false}}
 });
 return elements.get(key);
}
const nav=element('nav');nav.dataset={view:'estagios'};nav.textContent='Estagiários';
const document={
 querySelector:element,
 querySelectorAll:selector=>selector==='[data-view]'?[nav]:[],
 addEventListener(){},
 body:element('body')
};
const context={document,Date,console,setTimeout,clearTimeout,
 fetch:async(url)=>({ok:true,status:200,text:async()=>JSON.stringify(url.includes('/api/modelos')?{modelos:[]}:{registros:[]})})
};
vm.createContext(context);
const src=fs.readFileSync(path.join(__dirname,'../static/app.js'),'utf8');
vm.runInContext(src+'\ninit();this.getState=()=>state',context);
assert.equal(typeof nav.listeners.click,'function','Navigation handler must be installed');
nav.listeners.click();
assert.equal(context.getState().active,'estagios');
assert.equal(element('#breadcrumb').textContent,'Estagiários');
console.log('OK: menu initializes and navigates');
