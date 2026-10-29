/** 영화 한나 아렌트 활동 기록 웹앱. 모든 권한은 서버에서 검사합니다. */
var SCHOOL_DOMAIN_ = 'cnsa.hs.kr';
var SCHEMA_ = {
  Members: ['email','name','studentNo','status','createdAt','updatedAt'],
  Entries: ['id','email','title','activityDate','revision','status','data','createdAt','updatedAt','submittedAt','feedback','feedbackAt'],
  Versions: ['id','entryId','email','revision','data','title','activityDate','submittedAt'],
  Publications: ['entryId','revision','author','title','activityDate','fields','data','publishedAt','publishedBy','active'],
  Seminars: ['id','email','revision','data','published','publishedRevision','publishedAt','updatedAt']
};
var FIELD_LABELS_ = {
  scene: '인상적인 장면',
  questions: '영화를 보면서 들었던 질문과 나의 답',
  connection: '이 영화 속 장면과 한나 아렌트의 사상이 연결되는 지점은?',
  future: '향후 더 알아보고 싶은 점'
};

/** 편집기의 INSTALLER_EMAIL에 지정한 교사만 초기 설정을 실행합니다. */
function initializeApp() {
  var expected=String(PropertiesService.getScriptProperties().getProperty('INSTALLER_EMAIL')||'').toLowerCase();
  var current=String(Session.getActiveUser().getEmail()||'').toLowerCase();
  if(!expected || current!==expected) throw new Error('프로젝트 설정의 INSTALLER_EMAIL에 지정된 교사만 초기 설정을 실행할 수 있습니다.');
  return install_();
}

/** 편집기에서 학교 계정으로 한 번 실행합니다. 브라우저에서 호출할 수 없습니다. */
function install_() {
  var email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email || email.split('@')[1] !== SCHOOL_DOMAIN_) throw new Error('cnsa.hs.kr 학교 계정으로 편집기에서 실행해 주세요.');
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('DB_ID');
  if (existing) {
    var teachers = JSON.parse(props.getProperty('TEACHER_EMAILS') || '[]');
    if (teachers.indexOf(email) < 0) throw new Error('설치한 교사만 초기 설정을 확인할 수 있습니다.');
    ensureReadingSchema_({book:SpreadsheetApp.openById(existing)});
    console.log('기존 저장소를 유지합니다. 스프레드시트: ' + SpreadsheetApp.openById(existing).getUrl());
    return;
  }
  var sheet = SpreadsheetApp.create('한나 아렌트 영화 활동 기록');
  Object.keys(SCHEMA_).forEach(function(name, i) {
    var tab = i === 0 ? sheet.getSheets()[0].setName(name) : sheet.insertSheet(name);
    tab.getRange(1,1,1,SCHEMA_[name].length).setValues([SCHEMA_[name]]).setFontWeight('bold').setBackground('#e7f0f8');
    tab.setFrozenRows(1);
  });
  props.setProperties({DB_ID:sheet.getId(),TEACHER_EMAILS:JSON.stringify([email]),SCHOOL_DOMAIN:SCHOOL_DOMAIN_});
  console.log('설치 완료. 저장용 스프레드시트는 학생에게 공유하지 않습니다.');
  console.log(sheet.getUrl());
}

function doGet(e) {
  var template=HtmlService.createTemplateFromFile('Index');
  var requested=e && e.parameter && e.parameter.view;
  template.initialView=['write','archive','community','teacher','reading'].indexOf(requested)>=0?requested:'';
  template.initialRound=Number(e&&e.parameter&&e.parameter.round)||0;
  if(!Number.isInteger(template.initialRound)||template.initialRound<1||template.initialRound>8)template.initialRound=0;
  return template.evaluate()
    .setTitle('현대철학 동아리 활동 기록')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function include_(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }
function fail_(message, code) { var e = new Error(message); e.appCode = code || 'INVALID'; throw e; }
function text_(v,max) {
  if (v == null) return '';
  if (typeof v !== 'string') fail_('입력 형식을 확인해 주세요.');
  var t = v.trim();
  if (t.length > max) fail_('입력 가능한 분량을 넘었습니다. 내용을 줄여 주세요.');
  return t;
}
function now_() { return new Date().toISOString(); }
function json_(s, fallback) { try { return JSON.parse(s || ''); } catch (e) { return fallback; } }
function id_() { return Utilities.getUuid(); }
function publicMember_(m, teacher) {
  return {email:m.email,name:m.name,studentNo:m.studentNo,status:teacher?'approved':m.status,role:teacher?'teacher':'student'};
}
function identity_() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('DB_ID')) fail_('초기 설정이 필요합니다. 담당 교사에게 알려 주세요.','SETUP');
  // getEffectiveUser는 배포자의 계정이므로 학생 식별에 사용하지 않습니다.
  var email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) fail_('학교 계정을 확인하지 못했습니다. 학교 계정으로 접속했는지, 배포 설정이 학교 조직으로 제한되어 있는지 확인해 주세요.','AUTH');
  var domain = props.getProperty('SCHOOL_DOMAIN') || SCHOOL_DOMAIN_;
  if (email.split('@')[1] !== domain) fail_('학교 구글 계정으로 접속해 주세요.','AUTH');
  return {email:email,isTeacher:json_(props.getProperty('TEACHER_EMAILS'),[]).indexOf(email)>=0,dbId:props.getProperty('DB_ID')};
}
function db_(identity) { return {book:SpreadsheetApp.openById(identity.dbId),cache:{}}; }
function rows_(db,table) {
  if (db.cache[table]) return db.cache[table];
  var tab=db.book.getSheetByName(table);
  if (!tab) fail_('저장소의 시트 구성을 확인해 주세요.','STORAGE');
  var values=tab.getDataRange().getValues();
  if (values[0].join('|')!==SCHEMA_[table].join('|')) fail_('저장소의 열 구성을 확인해 주세요.','STORAGE');
  var result=[];
  for(var i=1;i<values.length;i++) {
    if (!values[i][0]) continue;
    var row={_row:i+1}; SCHEMA_[table].forEach(function(k,j){row[k]=values[i][j];}); result.push(row);
  }
  db.cache[table]=result; return result;
}
function plainCell_(v) { return typeof v==='string' && /^[=+\-@\t\r]/.test(v) ? "'"+v : v; }
function put_(db,table,obj) {
  var items=rows_(db,table), tab=db.book.getSheetByName(table);
  var row=Object.assign({},obj), pos=row._row || tab.getLastRow()+1;
  tab.getRange(pos,1,1,SCHEMA_[table].length).setNumberFormat('@')
    .setValues([SCHEMA_[table].map(function(k){return plainCell_(row[k]==null?'':row[k]);})]);
  row._row=pos;
  var i=items.findIndex(function(x){return x._row===pos;});
  if(i>=0) items[i]=row; else items.push(row);
  return row;
}
function member_(ctx) {
  var m=rows_(ctx.db,'Members').find(function(x){return x.email===ctx.email;});
  if(!m) m=put_(ctx.db,'Members',{email:ctx.email,name:'',studentNo:'',status:ctx.isTeacher?'approved':'pending',createdAt:now_(),updatedAt:now_()});
  if(ctx.isTeacher && m.status!=='approved') m=put_(ctx.db,'Members',Object.assign({},m,{status:'approved'}));
  ctx.member=m; return m;
}
function approved_(ctx) { if(!ctx.isTeacher && ctx.member.status!=='approved') fail_('담당 교사의 구성원 승인을 기다려 주세요.','MEMBERSHIP'); }
function teacher_(ctx) { if(!ctx.isTeacher) fail_('교사만 사용할 수 있습니다.','FORBIDDEN'); }
function entry_(ctx,id) {
  var e=rows_(ctx.db,'Entries').find(function(x){return x.id===id;});
  if(!e || (!ctx.isTeacher && e.email!==ctx.email)) fail_('이 기록을 열람할 수 없습니다.','FORBIDDEN');
  return e;
}
// 영화 활동지는 계정별 하나입니다. 예전 중복 기록은 원본을 남겨 읽기 전용으로 제공합니다.
function primaryMovie_(entries,email) {
  return entries.filter(function(e){return e.email===email && recordKind_(json_(e.data,{}))==='movie';})
    .sort(function(a,b){
      var aSubmitted=!!a.submittedAt || a.status==='submitted', bSubmitted=!!b.submittedAt || b.status==='submitted';
      return Number(bSubmitted)-Number(aSubmitted) || String(b.submittedAt||'').localeCompare(String(a.submittedAt||''))
        || String(b.updatedAt||'').localeCompare(String(a.updatedAt||'')) || String(a.id).localeCompare(String(b.id));
    })[0] || null;
}
function previousCopyOf_(ctx,e) {
  if(recordKind_(json_(e.data,{}))!=='movie')return '';
  var primary=primaryMovie_(rows_(ctx.db,'Entries'),e.email);
  return primary && primary.id!==e.id ? primary.id : '';
}
function entryView_(e,ctx) {
  var m=rows_(ctx.db,'Members').find(function(x){return x.email===e.email;}) || {};
  return {id:e.id,email:e.email,name:m.name||'',studentNo:m.studentNo||'',title:e.title,activityDate:e.activityDate,
    revision:Number(e.revision),status:e.status,data:json_(e.data,{}),createdAt:e.createdAt,updatedAt:e.updatedAt,
    submittedAt:e.submittedAt,feedback:e.feedback||'',feedbackAt:e.feedbackAt||'',previousCopyOf:previousCopyOf_(ctx,e)};
}
function publicationView_(p) {
  return {entryId:p.entryId,revision:Number(p.revision),author:p.author,title:p.title,activityDate:p.activityDate,
    fields:json_(p.fields,[]),data:json_(p.data,{}),publishedAt:p.publishedAt};
}
function state_(ctx) {
  var me=publicMember_(ctx.member,ctx.isTeacher);
  var out={user:me,entries:[],publications:[],members:[],seminars:[],labels:FIELD_LABELS_,recordLabels:READING_LABELS_};
  if(!ctx.isTeacher && ctx.member.status!=='approved') return out;
  var members=rows_(ctx.db,'Members'), es=rows_(ctx.db,'Entries');
  out.entries=es.filter(function(x){return ctx.isTeacher || x.email===ctx.email;})
    .sort(function(a,b){return String(b.updatedAt).localeCompare(String(a.updatedAt));}).map(function(e){return entryView_(e,ctx);});
  out.publications=rows_(ctx.db,'Publications').filter(function(p){
    if(String(p.active)!=='true') return false;
    var e=es.find(function(x){return x.id===p.entryId;});
    return !!e && members.some(function(m){return m.email===e.email && m.status==='approved';});
  }).sort(function(a,b){return String(b.publishedAt).localeCompare(String(a.publishedAt));}).map(publicationView_);
  if(ctx.isTeacher) out.members=members.map(function(m){return publicMember_(m,ctx.teacherEmails.indexOf(m.email)>=0);});
  out.seminars=seminarsState_(ctx);
  return out;
}
function cleanData_(raw,submit,ctx,entry) {
  raw=raw||{};
  var kind=recordKind_(raw);
  if(entry && recordKind_(json_(entry.data,{}))!==kind)fail_('기록의 활동 유형은 변경할 수 없습니다.');
  if(kind!=='movie')return cleanReadingRecord_(ctx,raw,submit,entry);
  if(!Array.isArray(raw.questions) || raw.questions.length<1 || raw.questions.length>10) fail_('질문과 답은 1개부터 10개까지 작성할 수 있습니다.');
  var d={scene:text_(raw.scene,6000),questions:raw.questions.map(function(q){
    if(!q || typeof q!=='object') fail_('질문 형식을 확인해 주세요.');
    return {question:text_(q.question,1500),answer:text_(q.answer,5000)};
  }),connection:text_(raw.connection,7000),future:text_(raw.future,3000)};
  if(JSON.stringify(d).length>40000) fail_('전체 작성 분량이 너무 깁니다. 내용을 줄여 주세요.');
  if(submit && (!d.scene || !d.connection || d.questions.some(function(q){return !q.question || !q.answer;})))
    fail_('인상적인 장면, 질문과 나의 답, 장면과 사상의 연결을 작성해 주세요. 추가한 질문에도 답을 적어 주세요.');
  return d;
}
function save_(ctx,b) {
  var id=text_(b.id,80);
  if(!/^[a-zA-Z0-9-]{16,80}$/.test(id)) fail_('기록 번호를 확인해 주세요.');
  var e=rows_(ctx.db,'Entries').find(function(x){return x.id===id;});
  if(e && e.email!==ctx.email) fail_('다른 학생의 기록은 수정할 수 없습니다.','FORBIDDEN');
  var expected=Number(b.revision);
  if(!Number.isInteger(expected) || expected<0 || (e?Number(e.revision)!==expected:expected!==0))
    fail_('다른 창에서 기록이 변경되었습니다. 현재 내용을 내려받은 뒤 새로고침해 주세요.','CONFLICT');
  if(!ctx.member.name || (!ctx.isTeacher && !ctx.member.studentNo)) fail_('먼저 이름과 학번을 입력해 주세요.','PROFILE');
  var submitted=b.submit===true, data=cleanData_(b.data,submitted,ctx,e);
  if(recordKind_(data)==='movie') {
    var primary=primaryMovie_(rows_(ctx.db,'Entries'),ctx.email);
    if(primary && (!e || e.id!==primary.id))fail_('영화 활동지는 한 개만 작성할 수 있습니다. 현재 입력 내용을 내려받아 보관한 뒤 새로고침하여 기존 활동지를 이어서 수정해 주세요.','CONFLICT');
  }
  if(data.kind && rows_(ctx.db,'Entries').some(function(x){var d=json_(x.data,{});return x.id!==id&&x.email===ctx.email&&d.kind===data.kind&&Number(d.round)===data.round;}))fail_('이 발제의 기록이 이미 있습니다. 나의 활동 기록에서 이어서 수정해 주세요.','CONFLICT');
  var day=text_(b.activityDate,10);
  if(!/^\d{4}-\d{2}-\d{2}$/.test(day) || isNaN(new Date(day+'T00:00:00Z').getTime())) fail_('활동일을 확인해 주세요.');
  var timestamp=now_();
  var next=Object.assign({},e||{id:id,email:ctx.email,createdAt:timestamp,feedback:'',feedbackAt:'',submittedAt:''},
    {title:text_(b.title,120)||'영화 한나 아렌트 활동 기록',activityDate:day,revision:expected+1,status:submitted?'submitted':'draft',data:JSON.stringify(data),updatedAt:timestamp});
  if(submitted) {
    next.submittedAt=timestamp;
    // 제출본과 공유본을 분리해 저장합니다. 공유 후 수정이 자동으로 공개되지 않습니다.
    put_(ctx.db,'Versions',{id:id_(),entryId:id,email:ctx.email,revision:next.revision,data:next.data,title:next.title,activityDate:day,submittedAt:timestamp});
  }
  next=put_(ctx.db,'Entries',next);
  return {entry:entryView_(next,ctx)};
}
function publish_(ctx,b) {
  teacher_(ctx);
  var e=entry_(ctx,text_(b.entryId,80));
  if(previousCopyOf_(ctx,e))fail_('보관된 이전 기록입니다. 현재 활동지의 제출본을 선택해 주세요.');
  if(e.status!=='submitted') fail_('학생이 제출한 기록만 공유할 수 있습니다. 작성 중인 기록은 공유하지 않습니다.');
  if(Number(b.revision)!==Number(e.revision)) fail_('기록이 변경되었습니다. 최신 제출본을 다시 확인해 주세요.','CONFLICT');
  var owner=rows_(ctx.db,'Members').find(function(m){return m.email===e.email;});
  if(!owner || owner.status!=='approved') fail_('승인된 구성원의 제출 기록만 공유할 수 있습니다.');
  var fields=Object.keys(labelsFor_(json_(e.data,{}))).filter(function(k){return Array.isArray(b.fields)&&b.fields.indexOf(k)>=0;});
  if(!fields.length) fail_('공유할 항목을 하나 이상 선택해 주세요.');
  var saved=rows_(ctx.db,'Versions').find(function(v){return v.entryId===e.id && Number(v.revision)===Number(e.revision);});
  if(!saved) fail_('제출본을 찾지 못했습니다. 학생이 다시 제출한 뒤 시도해 주세요.');
  var original=json_(saved.data,{}), snapshot=readingContext_(original);
  fields.forEach(function(k){snapshot[k]=original[k];});
  var previous=rows_(ctx.db,'Publications').find(function(p){return p.entryId===e.id;});
  put_(ctx.db,'Publications',Object.assign({},previous||{}, {entryId:e.id,revision:Number(e.revision),author:owner.name,
    title:saved.title,activityDate:saved.activityDate,fields:JSON.stringify(fields),data:JSON.stringify(snapshot),
    publishedAt:now_(),publishedBy:ctx.email,active:'true'}));
  return {ok:true};
}
function csvCell_(v) { var s=typeof v==='string'?v:String(v==null?'':v); if(/^[=+\-@\t\r]/.test(s))s="'"+s; return '"'+s.replace(/"/g,'""')+'"'; }
function export_(ctx) {
  teacher_(ctx);
  var rows=[['학번','이름','학교 이메일','활동일','제목','상태','인상적인 장면','질문과 나의 답','장면과 사상의 연결','향후 더 알아보고 싶은 점','교사 피드백','활동 유형','발제 번호','독서 활동 기록']];
  rows_(ctx.db,'Entries').forEach(function(e){
    var m=rows_(ctx.db,'Members').find(function(x){return x.email===e.email;})||{}, d=json_(e.data,{});
    rows.push([m.studentNo,m.name,e.email,e.activityDate,e.title,e.status==='submitted'?'제출':'작성 중',d.scene,
      (d.questions||[]).map(function(q,i){return (i+1)+'. '+q.question+'\n나의 답: '+q.answer;}).join('\n\n'),d.connection,d.future,e.feedback,recordKind_(d),d.round||'',recordKind_(d)==='movie'?'':JSON.stringify(d)]);
  });
  return {csv:'\uFEFF'+rows.map(function(r){return r.map(csvCell_).join(',');}).join('\r\n')};
}
/** 유일한 클라이언트 RPC 진입점. 클라이언트가 보낸 이메일·역할은 신뢰하지 않습니다. */
function api(b) {
  var lock;
  try {
    if(!b || typeof b!=='object' || Array.isArray(b) || JSON.stringify(b).length>55000) fail_('요청 형식을 확인해 주세요.');
    var who=identity_();
    lock=LockService.getScriptLock();
    if(!lock.tryLock(10000)) fail_('다른 기록을 저장 중입니다. 잠시 후 다시 시도해 주세요.','BUSY');
    var ctx={email:who.email,isTeacher:who.isTeacher,db:db_(who),teacherEmails:json_(PropertiesService.getScriptProperties().getProperty('TEACHER_EMAILS'),[])};
    ensureReadingSchema_(ctx.db);
    member_(ctx);
    var result, action=text_(b.action,30);
    if(action==='bootstrap') result=state_(ctx);
    else if(action==='profile') {
      var name=text_(b.name,60), no=text_(b.studentNo,20);
      if(!name || (!ctx.isTeacher && !no)) fail_('이름과 학번을 입력해 주세요.');
      ctx.member=put_(ctx.db,'Members',Object.assign({},ctx.member,{name:name,studentNo:no,updatedAt:now_()}));
      result={user:publicMember_(ctx.member,ctx.isTeacher)};
    } else {
      approved_(ctx);
      if(action==='save') result=save_(ctx,b);
      else if(action==='seminarClaim') result=claimSeminar_(ctx,b);
      else if(action==='seminarAssign') result=assignSeminar_(ctx,b);
      else if(action==='seminarSave') result=saveSeminar_(ctx,b);
      else if(action==='seminarUnpublish') result=unpublishSeminar_(ctx,b);
      else if(action==='publish') result=publish_(ctx,b);
      else if(action==='unpublish') {
        teacher_(ctx); entry_(ctx,text_(b.entryId,80));
        var post=rows_(ctx.db,'Publications').find(function(p){return p.entryId===b.entryId;});
        if(post) put_(ctx.db,'Publications',Object.assign({},post,{active:'false'}));
        result={ok:true};
      } else if(action==='member') {
        teacher_(ctx);
        var target=rows_(ctx.db,'Members').find(function(m){return m.email===text_(b.email,200).toLowerCase();});
        if(!target || ctx.teacherEmails.indexOf(target.email)>=0) fail_('학생 구성원을 선택해 주세요.');
        if(['approved','pending','blocked'].indexOf(b.status)<0) fail_('승인 상태를 확인해 주세요.');
        put_(ctx.db,'Members',Object.assign({},target,{status:b.status,updatedAt:now_()})); result={ok:true};
      } else if(action==='feedback') {
        teacher_(ctx); var targetEntry=entry_(ctx,text_(b.entryId,80));
        put_(ctx.db,'Entries',Object.assign({},targetEntry,{feedback:text_(b.feedback,5000),feedbackAt:now_()})); result={ok:true};
      } else if(action==='versions') {
        var own=entry_(ctx,text_(b.entryId,80));
        result={versions:rows_(ctx.db,'Versions').filter(function(v){return v.entryId===own.id && Number(v.revision)<=Number(own.revision);})
          .map(function(v){return {revision:Number(v.revision),title:v.title,activityDate:v.activityDate,submittedAt:v.submittedAt,data:json_(v.data,{})};}).reverse()};
      } else if(action==='export') result=export_(ctx);
      else fail_('지원하지 않는 요청입니다.');
    }
    SpreadsheetApp.flush();
    return {ok:true,result:result};
  } catch(e) {
    if(!e.appCode) console.error('Movie journal server error: '+e.name);
    return {ok:false,code:e.appCode||'SERVER',error:e.appCode?e.message:'처리하지 못했습니다. 입력 내용은 그대로 유지됩니다. 잠시 후 다시 시도해 주세요.'};
  } finally { if(lock) lock.releaseLock(); }
}

