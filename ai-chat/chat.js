/* Khantaishir AI — client-only widget. No API key is placed in the browser. */
(function(){
  'use strict';
  const meta = document.querySelector('meta[name="khantaishir-ai-endpoint"]');
  const endpoint = (meta && meta.getAttribute('content') || '').trim();
  // Do not display an AI button until a real Worker URL is configured.
  if (!/^https:\/\/[^\s/]+\/api\/chat\/?$/.test(endpoint)) return;
  const build = (tag, cls, parent, text) => {
    const el = document.createElement(tag);
    if (cls) el.className = cls;
    if (text !== undefined) el.textContent = text;
    if (parent) parent.appendChild(el);
    return el;
  };
  const launcher = build('button','kh-ai-launcher',document.body);
  launcher.type='button';
  launcher.setAttribute('aria-label','AI туслахтай чатлах');
  launcher.setAttribute('aria-expanded','false');
  launcher.setAttribute('aria-controls','khAI');
  build('span','kh-ai-dot',launcher,'AI');
  build('span','kh-ai-launch-text',launcher,'AI туслах');
  const panel = build('section','kh-ai-panel',document.body);
  panel.id='khAI';panel.hidden=true;
  panel.setAttribute('aria-label','Хантайшир Эко Блок AI туслах');
  const header = build('div','kh-ai-head',panel);
  const hgroup = build('div','',header);
  build('strong','',hgroup,'AI туслах');
  build('small','',hgroup,'Автомат AI туслах · худалдааны мэдээлэл');
  const close = build('button','kh-ai-close',header,'×');
  close.type='button';close.setAttribute('aria-label','Чат хаах');
  const feed = build('div','kh-ai-messages',panel);
  feed.setAttribute('role','log');feed.setAttribute('aria-live','polite');feed.setAttribute('aria-relevant','additions');
  const chips = build('div','kh-ai-chips',panel);
  const presets=['1 м³ хэдэн төгрөг вэ?','200 ба 300 мм ялгаа?','Байшингийн блок тооцох'];
  const form=build('form','kh-ai-form',panel);
  const field=build('textarea','',form);
  field.placeholder='Блокийн талаар асуугаарай...';
  field.rows=1;field.maxLength=500;field.required=true;
  field.setAttribute('aria-label','AI туслахаас асуух');
  const send=build('button','kh-ai-send',form,'Илгээх');
  send.type='submit';
  const foot=build('div','kh-ai-foot',panel);
  build('span','',foot,'AI алдаж болно. Хувийн мэдээлэл бүү оруул.');
  const link=build('a','',foot,'Үнийн санал авах');
  link.href='#quote';
  const messages=[];
  let busy=false;
  function bubble(text,who){
    const node=build('div','kh-ai-bubble '+who,feed,text);
    feed.scrollTop=feed.scrollHeight;
    return node;
  }
  bubble('Сайн байна уу! AAC блокийн хэмжээ, жишиг үнэ болон хэрэглээний талаар асуугаарай. Албан үнийн саналыг борлуулалтын ажилтан баталгаажуулна.','bot');
  presets.forEach((question)=>{
    const b=build('button','',chips,question);b.type='button';
    b.addEventListener('click',()=>ask(question));
  });
  function setBusy(on){
    busy=on;send.disabled=on;field.disabled=on;
    chips.querySelectorAll('button').forEach(b=>b.disabled=on);
    if(!on)field.focus({preventScroll:true});
  }
  function setOpen(open){
    panel.hidden=!open;
    launcher.setAttribute('aria-expanded',String(open));
    if(open)field.focus({preventScroll:true});else launcher.focus({preventScroll:true});
  }
  launcher.addEventListener('click',()=>setOpen(panel.hidden));
  close.addEventListener('click',()=>setOpen(false));
  document.addEventListener('keydown',e=>{
    if(e.key==='Escape'&&!panel.hidden)setOpen(false);
  });
  link.addEventListener('click',()=>setOpen(false));
  field.addEventListener('keydown',e=>{
    if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();if(!busy)form.requestSubmit();}
  });
  form.addEventListener('submit',e=>{e.preventDefault();ask(field.value);});
  async function ask(value){
    const question=String(value||'').trim().slice(0,500);
    if(!question||busy)return;
    if(/(?:\+?976[\s-]?)?\b\d{4}[\s-]?\d{4}\b/.test(question)||/[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/.test(question)){
      bubble('Утас, имэйл зэрэг мэдээллээ чатад оруулахгүйгээр "Үнийн санал авах" маягтаар илгээнэ үү.','bot');
      field.value='';return;
    }
    bubble(question,'user');field.value='';
    const pending=bubble('AI хариулж байна…','bot pending');
    const ctl=new AbortController();
    const timeout=setTimeout(()=>ctl.abort(),24000);
    setBusy(true);
    try{
      const history=messages.slice(-7);
      const payload=[...history,{role:'user',content:question}];
      const response=await fetch(endpoint,{
        method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({messages:payload}),signal:ctl.signal,
        credentials:'omit',cache:'no-store'
      });
      if(!response.ok)throw Error('AI түр хугацаанд хариулах боломжгүй байна.');
      const data=await response.json();
      if(typeof data.answer!=='string'||!data.answer.trim())throw Error('Хариулт хоосон байна.');
      pending.classList.remove('pending');
      pending.textContent=data.answer.slice(0,1500);
      messages.push({role:'user',content:question},{role:'assistant',content:data.answer.slice(0,1500)});
      if(messages.length>8)messages.splice(0,messages.length-8);
    }catch(_){
      pending.classList.remove('pending');
      pending.textContent='AI түр хугацаанд ажиллахгүй байна. Тооцоолуур ашиглах эсвэл 8600 0254 дугаарт холбогдоно уу.';
    }finally{
      clearTimeout(timeout);setBusy(false);feed.scrollTop=feed.scrollHeight;
    }
  }
})();
