const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
process.env.TZ='America/Sao_Paulo';
const elements=new Map();
function element(k){if(!elements.has(k))elements.set(k,{value:'',innerHTML:'',textContent:'',classList:{toggle(){},add(){},remove(){}}});return elements.get(k)}
const document={querySelector:element,querySelectorAll:()=>[],addEventListener:()=>{}};
const context={document,Date,console,setTimeout,clearTimeout};
vm.createContext(context);
const source=fs.readFileSync(path.join(__dirname,'../static/app.js'),'utf8');
vm.runInContext(source+'\nthis.subject={localTodayIso,upcomingStarts,showAllUpcoming,renderDirectory,renderDashboard,state}',context);
const {localTodayIso,upcomingStarts,showAllUpcoming,renderDirectory,renderDashboard,state}=context.subject;
assert.equal(localTodayIso(new Date('2026-10-16T01:00:00Z')),'2026-10-15');
const records=[
{id:1,nome:'Passado',inicio:'2026-10-14',status:'ativo'},
{id:2,nome:'Novembro',inicio:'2026-11-01',status:'ativo'},
{id:3,nome:'Hoje B',inicio:'2026-10-15',status:'ativo'},
{id:4,nome:'Hoje A',inicio:'2026-10-15',status:'ativo'},
{id:5,nome:'Cancelado',inicio:'2026-10-16',status:'cancelado'},
{id:6,nome:'Certificado enviado',inicio:'2026-10-17',status:'ativo',status_certificado:'enviado'},
{id:7,nome:'Recusado',inicio:'2026-10-16',status:'ativo',situacao_aprovacao:'recusado'},
{id:8,nome:'Sem data',inicio:'',status:'ativo'},
{id:9,nome:'Proximo',inicio:'2026-10-16',status:'ativo'}
];
assert.deepEqual([...upcomingStarts(records,'2026-10-15')].map(r=>r.id),[4,3,9,2]);
assert.deepEqual([...upcomingStarts(records,'2026-11-02')].map(r=>r.id),[]);
state.registros=records.map(r=>({email:'',especialidade:'',situacao_documentos:'pendente',etapa:'Aguardando aprovação',fim:'2026-11-02',...r}));
element('#search').value='Busca antiga';
element('#filter').value='Aguardando documentos';
showAllUpcoming();
assert.equal(state.active,'estagios');
assert.equal(element('#search').value,'');
assert.equal(element('#filter').value,'');
assert.equal(element('#period-filter').value,'upcoming');
state.registros=[
{id:101,nome:'Anterior',inicio:'2020-01-01',status:'ativo',etapa:'Aguardando aprovação'},
{id:102,nome:'Futuro',inicio:'2099-01-01',status:'ativo',etapa:'Aguardando aprovação'}
].map(r=>({email:'',especialidade:'',situacao_documentos:'pendente',fim:'2099-02-01',...r}));
renderDirectory();
assert.equal(element('#student-body').innerHTML.includes('Futuro'),true);
assert.equal(element('#student-body').innerHTML.includes('Anterior'),false);
assert.equal(element('#count-result').textContent,'1 próximo(s) início(s)');
element('#period-filter').value='';
renderDirectory();
assert.equal(element('#student-body').innerHTML.includes('Anterior'),true);
state.registros=Array.from({length:9},(_,i)=>({
id:i+1,nome:'Estudante '+(i+1),inicio:'2099-01-'+String(9-i).padStart(2,'0'),
status:'ativo',status_certificado:'pendente',situacao_documentos:'pendente',
etapa:'Aguardando aprovação',alertas_importacao:'',especialidade:'',email:'',fim:'2099-02-01'
}));
renderDashboard();
const html=element('#upcoming').innerHTML;
assert.equal((html.match(/class="up-row"/g)||[]).length,6);
assert.ok(html.indexOf('Estudante 9')<html.indexOf('Estudante 4'));
assert.equal(html.includes('Estudante 1'),false);
console.log('OK: proximos inicios, ordenacao, data local, filtro e historico');
