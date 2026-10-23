/** 『인간의 조건』 발제 1~8, 토의 답변, 개인 성찰. 기존 영화 기록은 유지합니다. */
var READING_LABELS_={
  discussion:{questions:'토의 질문에 대한 나의 답'},
  reflection:{learned:'이번 활동에서 이해한 점',thinking:'토의를 거치며 달라지거나 유지한 생각',future:'더 탐구하고 싶은 질문'}
};
function ensureReadingSchema_(db){
  if(!db.book.getSheetByName('Seminars')){
    var tab=db.book.insertSheet('Seminars');
    tab.getRange(1,1,1,SCHEMA_.Seminars.length).setValues([SCHEMA_.Seminars]).setFontWeight('bold').setBackground('#e7f0f8');
    tab.setFrozenRows(1);
  }
}
function recordKind_(d){
  var k=d&&d.kind||'movie';
  if(['movie','discussion','reflection'].indexOf(k)<0)fail_('활동 유형을 확인해 주세요.');
  return k;
}
function labelsFor_(d){var k=recordKind_(d);return k==='movie'?FIELD_LABELS_:READING_LABELS_[k];}
function readingContext_(d){
  if(recordKind_(d)==='movie')return {};
  return {kind:d.kind,round:d.round,seminarRevision:d.seminarRevision,topic:d.topic};
}
function roundNumber_(n){var v=Number(n);if(!Number.isInteger(v)||v<1||v>8)fail_('발제 번호는 1~8입니다.');return v;}
function seminar_(ctx,n){return rows_(ctx.db,'Seminars').find(function(s){return Number(s.id)===n;});}
function defaultSeminar_(){return {chapter:'',startPage:'',endPage:'',activityDate:'',summary:'',questions:[]};}
function publishedSeminar_(ctx,s){
  if(!s||!s.publishedAt)return null;
  var owner=rows_(ctx.db,'Members').find(function(m){return m.email===s.email;});
  if(!owner||owner.status!=='approved')return null;
  return json_(s.published,null);
}
function seminarsState_(ctx){
  var members=rows_(ctx.db,'Members'), all=rows_(ctx.db,'Seminars');
  var mine=all.some(function(s){return s.email===ctx.email;});
  return Array.from({length:8},function(_,i){
    var n=i+1,s=all.find(function(x){return Number(x.id)===n;}),owner=members.find(function(m){return s&&m.email===s.email;});
    var editable=ctx.isTeacher||!!s&&s.email===ctx.email;
    var out={number:n,assigned:!!(s&&s.email),ownerName:owner&&owner.name||'',assignedToMe:!!s&&s.email===ctx.email,
      canClaim:!ctx.isTeacher&&!(s&&s.email)&&!mine,canEdit:editable,
      revision:editable&&s?Number(s.revision):0,draft:editable?json_(s&&s.data,defaultSeminar_()):null,
      publication:publishedSeminar_(ctx,s),publishedRevision:s?Number(s.publishedRevision)||0:0};
    if(ctx.isTeacher)out.ownerEmail=s&&s.email||'';
    return out;
  });
}
function assignSeminar_(ctx,b){
  teacher_(ctx);var n=roundNumber_(b.number),email=text_(b.email,200).toLowerCase();
  var owner=rows_(ctx.db,'Members').find(function(m){return m.email===email&&m.status==='approved'&&ctx.teacherEmails.indexOf(m.email)<0;});
  if(!owner)fail_('승인된 학생을 발제자로 선택해 주세요.');
  return setSeminarOwner_(ctx,n,email);
}
function claimSeminar_(ctx,b){
  if(ctx.isTeacher)fail_('담당 학생을 선택해 지정해 주세요.');
  if(!ctx.member.name||!ctx.member.studentNo)fail_('먼저 이름과 학번을 입력해 주세요.','PROFILE');
  return setSeminarOwner_(ctx,roundNumber_(b.number),ctx.email);
}
function setSeminarOwner_(ctx,n,email){
  var s=seminar_(ctx,n),all=rows_(ctx.db,'Seminars');
  if(s&&s.email&&s.email!==email)fail_('이미 담당자가 있는 발제입니다.','CONFLICT');
  if(all.some(function(x){return x.email===email&&Number(x.id)!==n;}))fail_('한 학생은 한 발제를 담당합니다. 이미 담당한 발제를 확인해 주세요.');
  if(!s||!s.email)put_(ctx.db,'Seminars',Object.assign({},s||{id:String(n),data:JSON.stringify(defaultSeminar_()),published:'',publishedRevision:0,publishedAt:''},
    {email:email,revision:s?Number(s.revision)+1:1,updatedAt:now_()}));
  return {ok:true};
}
function cleanSeminar_(raw,publish){
  raw=raw||{};
  var qs=raw.questions;
  if(!Array.isArray(qs)||qs.length>10)fail_('토의 질문은 최대 10개입니다.');
  var seen={};
  var d={chapter:text_(raw.chapter,200),startPage:text_(raw.startPage,8),endPage:text_(raw.endPage,8),activityDate:text_(raw.activityDate,10),summary:text_(raw.summary,12000),
    questions:qs.map(function(q){
      if(!q||typeof q!=='object')fail_('질문 형식을 확인해 주세요.');
      var id=text_(q.id,80);if(!/^[a-zA-Z0-9-]{8,80}$/.test(id)||seen[id])fail_('질문 번호를 확인해 주세요.');seen[id]=true;
      return {id:id,question:text_(q.question,1500),source:text_(q.source,1500)};
    })};
  if(d.activityDate&&(!/^\d{4}-\d{2}-\d{2}$/.test(d.activityDate)||isNaN(new Date(d.activityDate+'T00:00:00Z').getTime())))fail_('활동일을 확인해 주세요.');
  if(publish){
    if(!d.chapter||!d.summary||!d.activityDate||!d.questions.length||d.questions.some(function(q){return !q.question;}))fail_('담당 장·절, 활동일, 내용 요약과 한 개 이상의 토의 질문을 작성해 주세요.');
    if(!/^\d{1,6}$/.test(d.startPage)||!/^\d{1,6}$/.test(d.endPage)||Number(d.startPage)<1||Number(d.endPage)<Number(d.startPage))fail_('담당 시작·끝 쪽수를 확인해 주세요.');
  }
  if(JSON.stringify(d).length>40000)fail_('발제 자료의 전체 분량이 너무 깁니다.');
  return d;
}
function saveSeminar_(ctx,b){
  var n=roundNumber_(b.number),s=seminar_(ctx,n);
  if(!ctx.isTeacher&&(!s||s.email!==ctx.email))fail_('담당 발제자와 교사만 발제 자료를 수정할 수 있습니다.','FORBIDDEN');
  var expected=Number(b.revision);
  if(!Number.isInteger(expected)||expected<0||(s?Number(s.revision)!==expected:expected!==0))fail_('발제 자료가 다른 창에서 변경되었습니다. 내용을 내려받고 새로고침해 주세요.','CONFLICT');
  var publish=b.publish===true,d=cleanSeminar_(b.data,publish),stamp=now_();
  var next=Object.assign({},s||{id:String(n),email:'',published:'',publishedRevision:0,publishedAt:''},{revision:expected+1,data:JSON.stringify(d),updatedAt:stamp});
  if(publish){
    var owner=rows_(ctx.db,'Members').find(function(m){return m.email===next.email&&m.status==='approved';});
    if(!owner)fail_('승인된 담당 학생을 먼저 지정해 주세요.');
    next.published=JSON.stringify(Object.assign({},d,{author:owner.name}));next.publishedRevision=next.revision;next.publishedAt=stamp;
  }
  put_(ctx.db,'Seminars',next);
  return {seminars:seminarsState_(ctx)};
}
function unpublishSeminar_(ctx,b){
  var s=seminar_(ctx,roundNumber_(b.number));
  if(!s||(!ctx.isTeacher&&s.email!==ctx.email))fail_('담당 발제자와 교사만 공유를 해제할 수 있습니다.','FORBIDDEN');
  put_(ctx.db,'Seminars',Object.assign({},s,{publishedAt:'',updatedAt:now_()}));return {ok:true};
}
function cleanReadingRecord_(ctx,raw,submit,entry){
  var kind=recordKind_(raw),n=roundNumber_(raw.round),old=entry&&json_(entry.data,{}),context;
  if(old){
    if(old.kind!==kind||Number(old.round)!==n)fail_('기록의 발제 번호는 변경할 수 없습니다.');
    context=old;
  }else{
    var seminar=seminar_(ctx,n),pub=publishedSeminar_(ctx,seminar);
    if(!pub)fail_('공유된 발제 자료를 확인한 뒤 기록을 작성해 주세요.');
    if(Number(raw.seminarRevision)!==Number(seminar.publishedRevision))fail_('발제 자료가 새로 공유되었습니다. 최신 자료를 확인한 뒤 기록을 작성해 주세요.','CONFLICT');
    context={kind:kind,round:n,seminarRevision:Number(seminar.publishedRevision),topic:pub.chapter+' · '+pub.startPage+'~'+pub.endPage+'쪽',questions:pub.questions};
  }
  var d={kind:kind,round:n,seminarRevision:context.seminarRevision,topic:context.topic};
  if(kind==='discussion'){
    if(!Array.isArray(raw.questions)||raw.questions.length!==context.questions.length)fail_('작성 중인 토의 질문을 확인해 주세요.');
    var seen={};raw.questions.forEach(function(q){if(!q||seen[q.questionId])fail_('질문 번호가 중복되었습니다.');seen[q.questionId]=true;});
    d.questions=context.questions.map(function(q){
      var id=q.questionId||q.id,item=raw.questions.find(function(x){return x.questionId===id;});
      if(!item)fail_('토의 질문이 바뀌었습니다. 기록을 다시 열어 주세요.','CONFLICT');
      return {questionId:id,question:q.question,source:q.source||'',answer:text_(item.answer,5000)};
    });
    if(submit&&d.questions.some(function(q){return !q.answer;}))fail_('각 토의 질문에 대한 나의 답을 작성해 주세요.');
  }else{
    d.learned=text_(raw.learned,6000);d.thinking=text_(raw.thinking,6000);d.future=text_(raw.future,3000);
    if(submit&&(!d.learned||!d.thinking))fail_('이번 활동에서 이해한 점과 토의 후 나의 생각을 작성해 주세요.');
  }
  if(JSON.stringify(d).length>40000)fail_('전체 작성 분량이 너무 깁니다.');
  return d;
}
