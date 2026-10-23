// 기능 확인용 로컬 서버입니다. 학교 데이터·Google 계정에 연결하지 않습니다.
import http from 'node:http';
import fs from 'node:fs';
import crypto from 'node:crypto';
import {createApp}from './tests/mock-app.mjs';
const app=createApp(),T='teacher@cnsa.hs.kr',A='student-a@cnsa.hs.kr',B='student-b@cnsa.hs.kr';
for(const [email,name,no]of [[T,'미리보기 교사',''],[A,'가상 학생 가','20001'],[B,'가상 학생 나','20002']]){
 app.request(email,{action:'bootstrap'});app.request(email,{action:'profile',name,studentNo:no});if(email!==T)app.request(T,{action:'member',email,status:'approved'});
}
const id=crypto.randomUUID();
app.request(B,{action:'save',id,revision:0,title:'화면 확인용 가상 활동 기록',activityDate:'2026-10-23',submit:true,data:{scene:'선택한 장면과 그 이유가 이곳에 표시됩니다.',questions:[{question:'친구들에게 던질 질문이 표시됩니다.',answer:'이 질문에 대해 먼저 작성한 자신의 답이 표시됩니다.'}],connection:'학생이 장면과 아렌트의 사상을 연결해 설명한 내용이 표시됩니다.',future:'다음 독서에서 탐구하고 싶은 점이 표시됩니다.'}});
app.request(T,{action:'publish',entryId:id,revision:1,fields:['scene','questions','connection','future']});
let html=fs.readFileSync(new URL('Index.html',import.meta.url),'utf8');
html=html.replace(/<\?!= include_\('([^']+)'\); \?>/g,(_,n)=>fs.readFileSync(new URL(n+'.html',import.meta.url),'utf8')).replace('<?= initialView ?>','');
const bridge=`<style>.preview-bar{background:#143b59;color:white;padding:9px 15px;font:12px/1.7 sans-serif;text-align:center;position:relative;z-index:90}.preview-bar button{background:white;border:0;border-radius:5px;padding:5px 10px;color:#143b59;margin-left:5px;cursor:pointer}.preview-bar span{margin-right:8px}</style><div class="preview-bar"><span>화면 검토용 · 가상 자료 · 학교 계정과 저장소에 연결되지 않습니다.</span><button id="preview-student">학생 화면</button><button id="preview-teacher">교사 화면</button></div><script>const previewUser=sessionStorage.getItem('movie-preview-role')||'student-a@cnsa.hs.kr';document.getElementById('preview-student').onclick=()=>{sessionStorage.setItem('movie-preview-role','student-a@cnsa.hs.kr');location.reload()};document.getElementById('preview-teacher').onclick=()=>{sessionStorage.setItem('movie-preview-role','teacher@cnsa.hs.kr');location.reload()};window.google={script:{run:{withSuccessHandler(fn){const handler={success:fn,withFailureHandler(f){this.failure=f;return this;},api(body){fetch('/__preview/api',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({user:previewUser,body})}).then(r=>r.json()).then(this.success).catch(this.failure);}};return handler;}}}};</script>`;
html=html.replace('<body data-initial-view="">','<body data-initial-view="">'+bridge);
const server=http.createServer(async(req,res)=>{
 if(req.method==='GET'&&req.url==='/'){res.writeHead(200,{'Content-Type':'text/html; charset=utf-8','Cache-Control':'no-store'});return res.end(html);}
 if(req.method==='POST'&&req.url==='/__preview/api'){
  let raw='';for await(const chunk of req){raw+=chunk;if(raw.length>70000){res.writeHead(413);return res.end();}}
  try{const x=JSON.parse(raw);if(![T,A,B].includes(x.user))throw new Error('Unknown preview user');const result=app.call(x.user,x.body);res.writeHead(200,{'Content-Type':'application/json','Cache-Control':'no-store'});res.end(JSON.stringify(result));}catch(e){res.writeHead(400);res.end('{}');}return;
 }
 res.writeHead(404);res.end('Not found');
});
server.listen(4173,'0.0.0.0',()=>console.log('Movie app preview available at http://localhost:4173'));
