/* Progressive motion enhancements; standalone so the original business logic stays intact. */
(function(){
'use strict';
function init(){
 const section=document.getElementById('kh3d');const stage=document.getElementById('khViewer');
 if(!section||!stage)return;
 const cube=stage.querySelector('.kh-cube'),anchor=stage.querySelector('.kh-cube-anchor');
 const sizeBtns=Array.from(section.querySelectorAll('.kh-width-btn'));
 const sizeLabel=document.getElementById('khSizeLabel'),spinBtn=document.getElementById('khRotate');
 const motion=window.matchMedia('(prefers-reduced-motion: reduce)');
 let size=300,rx=-19,ry=-30,dragging=false,lastX=0,lastY=0,raf=0;
 const desktop=window.matchMedia('(min-width:521px)');
 const changeAuto=(on)=>{
  const active=Boolean(on&&!motion.matches&&!document.hidden);
  stage.dataset.auto=String(active);
  spinBtn.setAttribute('aria-pressed',String(active));
  spinBtn.textContent=active?'⏸ Эргэлтийг зогсоох':'↻ Автоматаар эргүүлэх';
 };
 function setSize(value,syncProduct){
  size=(Number(value)===200)?200:300;
  const depth=desktop.matches?(size===200?88:132):(size===200?71:107);
  anchor.style.setProperty('--kh-depth',depth+'px');
  sizeLabel.textContent='600 × 240 × '+size+' мм';
  stage.setAttribute('aria-label','AAC блокийн 3D дүрслэл. 600 × 240 × '+size+' мм. Хулганаар чирэх, сумны товчлуураар эргүүлэх боломжтой.');
  sizeBtns.forEach(btn=>btn.setAttribute('aria-pressed',String(Number(btn.dataset.khSize)===size)));
  if(syncProduct){
   const productButton=document.querySelector('.size-link[data-size="'+size+'"]');
   if(productButton)productButton.click();
  }
 }
 sizeBtns.forEach(btn=>btn.addEventListener('click',()=>setSize(btn.dataset.khSize,true)));
 document.querySelectorAll('.size-link[data-size]').forEach(btn=>btn.addEventListener('click',()=>setSize(btn.dataset.size,false)));
 const responsive=()=>setSize(size,false);
 if(desktop.addEventListener)desktop.addEventListener('change',responsive);else desktop.addListener(responsive);
 function render(){raf=0;cube.style.setProperty('--kh-rx',rx+'deg');cube.style.setProperty('--kh-ry',ry+'deg');}
 function schedule(){if(!raf)raf=requestAnimationFrame(render);}
 function stopDrag(e){
  if(!dragging)return;dragging=false;
  if(e&&stage.hasPointerCapture&&stage.hasPointerCapture(e.pointerId))stage.releasePointerCapture(e.pointerId);
  stage.classList.remove('kh-dragging');
 }
 stage.addEventListener('pointerdown',e=>{
  if(e.target.closest('button')||e.button!==0)return;
  dragging=true;lastX=e.clientX;lastY=e.clientY;changeAuto(false);
  if(stage.setPointerCapture)stage.setPointerCapture(e.pointerId);
  stage.classList.add('kh-dragging');
 });
 stage.addEventListener('pointermove',e=>{
  if(!dragging)return;
  const dx=e.clientX-lastX,dy=e.clientY-lastY;lastX=e.clientX;lastY=e.clientY;
  ry+=dx*.55;rx=Math.max(-75,Math.min(65,rx-dy*.42));schedule();
 });
 stage.addEventListener('pointerup',stopDrag);
 stage.addEventListener('pointercancel',stopDrag);
 stage.addEventListener('lostpointercapture',()=>{dragging=false;stage.classList.remove('kh-dragging');});
 stage.addEventListener('keydown',e=>{
  if(e.target!==stage)return;
  if(!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(e.key))return;
  e.preventDefault();changeAuto(false);
  if(e.key==='ArrowLeft')ry-=15;if(e.key==='ArrowRight')ry+=15;
  if(e.key==='ArrowUp')rx=Math.max(-75,rx-10);if(e.key==='ArrowDown')rx=Math.min(65,rx+10);
  schedule();
 });
 spinBtn.addEventListener('click',()=>changeAuto(stage.dataset.auto!=='true'));
 document.addEventListener('visibilitychange',()=>{if(document.hidden){if(stage.dataset.auto==='true')stage.dataset.resume='true';changeAuto(false);}else if(stage.dataset.resume==='true'){delete stage.dataset.resume;changeAuto(true);}});
 const changed=()=>{if(motion.matches)changeAuto(false);};
 if(motion.addEventListener)motion.addEventListener('change',changed);else motion.addListener(changed);
 setSize(300,false);changeAuto(!motion.matches);
 // A subtle parallax for existing photography. Passive scrolling + one animation frame.
 const hero=document.querySelector('.hero'),about=document.querySelector('.about-picture.big');
 function updateParallax(){
  rafScroll=0;if(motion.matches||document.hidden)return;
  if(hero){const y=hero.getBoundingClientRect();if(y.bottom>0&&y.top<innerHeight)hero.style.setProperty('--kh-hero-y',Math.max(-70,Math.min(70,-y.top*.12)).toFixed(1)+'px');}
  if(about){const y=about.getBoundingClientRect();if(y.bottom>0&&y.top<innerHeight)about.style.setProperty('--kh-about-y',Math.max(-30,Math.min(30,(innerHeight/2-y.top-y.height/2)*.08)).toFixed(1)+'px');}
 }
 let rafScroll=0;
 function requestParallax(){if(!rafScroll&&!motion.matches)rafScroll=requestAnimationFrame(updateParallax);}
 addEventListener('scroll',requestParallax,{passive:true});addEventListener('resize',requestParallax,{passive:true});requestParallax();
 // Stagger smaller benefit cards; existing page already reveals its large sections.
 if(!motion.matches&&'IntersectionObserver' in window){
  document.body.classList.add('kh-motion-on');
  const items=Array.from(document.querySelectorAll('.benefit,.contact-tile:not([hidden]),.tech-strip-item'));
  items.forEach((el,i)=>{el.classList.add('kh-enter');el.style.transitionDelay=(i%4)*85+'ms';});
  const ob=new IntersectionObserver(entries=>entries.forEach(e=>{
   if(e.isIntersecting){e.target.classList.add('kh-visible');ob.unobserve(e.target);}
  }),{threshold:.06,rootMargin:'0px 0px 35px 0px'});
  items.forEach(x=>ob.observe(x));
 }
 // Animate only the static factory capacity; never alter calculator prices/results.
 const stat=document.querySelector('.hero-stats strong');
 if(stat&&/^50[,\s]?000/.test(stat.textContent.trim())&&!motion.matches){
  const original=stat.textContent;
  const count=()=>{const duration=1500,start=performance.now();
   function frame(t){const x=Math.min(1,(t-start)/duration),p=1-Math.pow(1-x,3);
    stat.textContent=Math.round(50000*p).toLocaleString('en-US')+' м³';
    if(x<1)requestAnimationFrame(frame);else stat.textContent=original;
   }
   requestAnimationFrame(frame);
  };
  if('IntersectionObserver' in window){
   const ob=new IntersectionObserver(entries=>{if(entries.some(x=>x.isIntersecting)){ob.disconnect();count();}},{threshold:.25});ob.observe(stat);
  }else count();
 }
 // Mouse-only card tilt — modest angles; original cards remain clickable.
 if(!motion.matches&&matchMedia('(hover:hover) and (pointer:fine)').matches){
  document.querySelectorAll('.product-shot,.contact-tile').forEach(card=>{
   card.classList.add('kh-tilt');
   card.addEventListener('pointermove',e=>{
    const r=card.getBoundingClientRect(),x=(e.clientX-r.left)/r.width-.5,y=(e.clientY-r.top)/r.height-.5;
    card.style.setProperty('--kh-tilt-x',(-y*4).toFixed(2)+'deg');
    card.style.setProperty('--kh-tilt-y',(x*5).toFixed(2)+'deg');
   },{passive:true});
   card.addEventListener('pointerleave',()=>{card.style.removeProperty('--kh-tilt-x');card.style.removeProperty('--kh-tilt-y');});
  });
 }
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();
