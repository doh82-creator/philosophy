import vm from 'node:vm';
import fs from 'node:fs';
import crypto from 'node:crypto';
const source=fs.readFileSync(new URL('../Code.gs',import.meta.url),'utf8');
export function createApp(){
 let active='',uuid=0;
 const props={DB_ID:'test-book',TEACHER_EMAILS:JSON.stringify(['teacher@cnsa.hs.kr']),SCHOOL_DOMAIN:'cnsa.hs.kr'};
 class Range{
  constructor(sheet,r,c,n,m){Object.assign(this,{sheet,r,c,n,m});}
  getValues(){return Array.from({length:this.n},(_,i)=>Array.from({length:this.m},(_,j)=>this.sheet.values[this.r+i-1]?.[this.c+j-1]??''));}
  setValues(values){for(let i=0;i<values.length;i++){const target=this.r+i-1;this.sheet.values[target]||=[];for(let j=0;j<values[i].length;j++){const v=values[i][j];this.sheet.values[target][this.c+j-1]=typeof v==='string'&&/^'[=+\-@\t\r]/.test(v)?v.slice(1):v;}}return this;}
  setNumberFormat(){return this;}setFontWeight(){return this;}setBackground(){return this;}
 }
 class Sheet{
  constructor(name,headers){this.name=name;this.values=[headers];}
  getDataRange(){return new Range(this,1,1,this.values.length,this.values[0].length);}
  getRange(...args){return new Range(this,...args);}getLastRow(){return this.values.length;}
  setFrozenRows(){return this;}
 }
 const book={sheets:{},getSheetByName(n){return this.sheets[n]},getUrl(){return 'https://example.invalid/test-book'}};
 const context=vm.createContext({console,Date,JSON,Number,Object,Array,String,Error,RegExp,isNaN,
  Session:{getActiveUser:()=>({getEmail:()=>active})},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]??null,setProperties:v=>Object.assign(props,v)})},
  SpreadsheetApp:{openById:id=>{if(id!=='test-book')throw new Error('Unexpected DB');return book;},flush(){}},
  LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})},
  Utilities:{getUuid:()=>crypto.randomUUID()}
 });
 vm.runInContext(source,context,{filename:'Code.gs'});
 for(const [name,headers]of Object.entries(context.SCHEMA_))book.sheets[name]=new Sheet(name,headers);
 function call(email,body){active=email;return JSON.parse(JSON.stringify(context.api(body)));}
 function request(email,body){const r=call(email,body);if(!r.ok)throw Object.assign(new Error(r.error),{code:r.code});return r.result;}
 return {call,request,book,context,props};
}
