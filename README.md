[Programa Autoevaluacion2.html](https://github.com/user-attachments/files/28982062/Programa.Autoevaluacion2.html)
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ruta de Competencias</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --bg:#0F1B2D;--bg-card:#182944;--bg-card-hover:#21385C;--line:#2C4368;
    --text:#EFEAE0;--text-dim:#8FA3C4;--accent:#FFB454;--accent-soft:#3A2C18;
    --score3:#6FB7E0;--score4:#FFB454;--score5:#6FD9A0;--radius:18px;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  html,body{height:100%;}
  body{font-family:'Inter',sans-serif;background:radial-gradient(circle at 15% 0%,#1c2f4d 0%,transparent 45%),radial-gradient(circle at 85% 100%,#16273f 0%,transparent 50%),var(--bg);color:var(--text);min-height:100vh;display:flex;justify-content:center;-webkit-font-smoothing:antialiased;}
  #app{width:100%;max-width:780px;padding:28px 20px 60px;display:flex;flex-direction:column;min-height:100vh;}
  h1,h2,h3{font-family:'Space Grotesk',sans-serif;letter-spacing:-0.01em;}
  .eyebrow{font-family:'JetBrains Mono',monospace;font-size:12px;letter-spacing:0.18em;text-transform:uppercase;color:var(--accent);font-weight:700;}
  .top{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:28px;flex-wrap:wrap;gap:8px;}
  .brand{font-size:18px;font-weight:600;color:var(--text-dim);}.brand b{color:var(--text);}
  .hero{text-align:left;padding:36px 0 18px;}
  .hero h1{font-size:clamp(32px,6vw,46px);line-height:1.08;margin:10px 0 14px;}
  .hero p{color:var(--text-dim);font-size:16px;line-height:1.6;max-width:480px;}
  .grid-collab{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:14px;margin-top:30px;}
  .pick-card{background:var(--bg-card);border:1px solid var(--line);border-radius:var(--radius);padding:20px 16px;cursor:pointer;text-align:left;transition:transform .15s,border-color .15s,background .15s;color:var(--text);font-family:'Space Grotesk',sans-serif;}
  .pick-card:hover{background:var(--bg-card-hover);border-color:var(--accent);transform:translateY(-3px);}
  .pick-card .num{font-family:'JetBrains Mono',monospace;font-size:12px;color:var(--text-dim);display:block;margin-bottom:8px;}
  .pick-card .name{font-size:17px;font-weight:600;}
  .pick-card .status{margin-top:10px;font-size:12px;font-family:'JetBrains Mono',monospace;color:var(--score5);}
  .pick-card .status.pending{color:var(--text-dim);}
  .leader-row{margin-top:26px;border-top:1px solid var(--line);padding-top:22px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;}
  .leader-row p{color:var(--text-dim);font-size:14px;max-width:380px;}
  .btn{font-family:'Space Grotesk',sans-serif;font-weight:600;font-size:15px;border:none;border-radius:999px;padding:13px 26px;cursor:pointer;transition:transform .12s,opacity .12s;white-space:nowrap;}
  .btn:active{transform:scale(0.97);}
  .btn-primary{background:var(--accent);color:#1A1300;}.btn-primary:hover{opacity:0.92;}
  .btn-ghost{background:transparent;color:var(--text);border:1px solid var(--line);}.btn-ghost:hover{border-color:var(--accent);color:var(--accent);}
  .btn-danger{background:transparent;color:#FF8A8A;border:1px solid #4A2A2A;}.btn-danger:hover{background:#2A1414;}
  .btn[disabled]{opacity:0.35;cursor:not-allowed;}
  /* Trail */
  .trail{display:flex;align-items:center;margin-bottom:34px;overflow-x:auto;padding-bottom:6px;}
  .trail-node{display:flex;flex-direction:column;align-items:center;flex-shrink:0;width:64px;}
  .trail-line{flex:1;height:2px;background:var(--line);min-width:14px;margin-top:-22px;}.trail-line.done{background:var(--accent);}
  .dot{width:38px;height:38px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:700;border:2px solid var(--line);background:var(--bg-card);color:var(--text-dim);transition:all .2s;}
  .dot.current{border-color:var(--accent);color:var(--accent);box-shadow:0 0 0 5px var(--accent-soft);}
  .dot.done{border-color:transparent;color:#1A1300;}
  .dot.done.s3{background:var(--score3);}.dot.done.s4{background:var(--score4);}.dot.done.s5{background:var(--score5);}
  .trail-label{font-size:10px;color:var(--text-dim);margin-top:6px;text-align:center;line-height:1.2;max-width:64px;}
  /* Game */
  .station-card{background:var(--bg-card);border:1px solid var(--line);border-radius:24px;padding:30px 26px;flex:1;}
  .station-title{font-size:clamp(22px,4vw,30px);margin:10px 0 22px;line-height:1.2;}
  .option{display:block;width:100%;text-align:left;background:var(--bg);border:1.5px solid var(--line);border-radius:14px;padding:18px 20px;margin-bottom:12px;color:var(--text);font-size:15px;line-height:1.55;cursor:pointer;transition:border-color .15s,background .15s,transform .1s;font-family:'Inter',sans-serif;}
  .option:hover{border-color:var(--accent);background:#1C2E4C;}.option:active{transform:scale(0.99);}
  .option.selected{border-color:var(--accent);background:#1C2E4C;}
  .option.locked{cursor:default;}.option.locked:hover{border-color:var(--line);background:var(--bg);}
  .option.locked.selected{border-color:var(--accent);background:#1C2E4C;}
  .reveal{margin-top:22px;border-top:1px dashed var(--line);padding-top:22px;display:flex;align-items:center;gap:18px;flex-wrap:wrap;animation:fadeUp .35s ease;}
  @keyframes fadeUp{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:translateY(0);}}
  .score-badge{font-family:'JetBrains Mono',monospace;font-weight:700;font-size:34px;width:64px;height:64px;border-radius:16px;display:flex;align-items:center;justify-content:center;color:#1A1300;animation:pop .4s cubic-bezier(.34,1.56,.64,1);flex-shrink:0;}
  @keyframes pop{0%{transform:scale(0.4);opacity:0;}70%{transform:scale(1.08);}100%{transform:scale(1);opacity:1;}}
  .score-badge.s3{background:var(--score3);}.score-badge.s4{background:var(--score4);}.score-badge.s5{background:var(--score5);}
  .reveal-text{flex:1;min-width:160px;}.reveal-text .label{font-size:13px;color:var(--text-dim);margin-bottom:4px;}.reveal-text .value{font-size:15px;line-height:1.5;}
  .station-footer{display:flex;justify-content:flex-end;margin-top:24px;}
  /* Results */
  .results-head{padding:20px 0 8px;}.results-head h1{font-size:clamp(26px,5vw,38px);margin:8px 0 6px;}.results-head p{color:var(--text-dim);font-size:14px;}
  .total-box{background:var(--bg-card);border:1px solid var(--line);border-radius:20px;padding:26px;margin:22px 0;display:flex;align-items:center;gap:22px;flex-wrap:wrap;}
  .total-number{font-family:'JetBrains Mono',monospace;font-size:48px;font-weight:700;color:var(--accent);}
  .total-sub{color:var(--text-dim);font-size:13px;max-width:340px;line-height:1.5;}
  .bar-row{display:flex;align-items:center;gap:14px;margin-bottom:16px;}
  .bar-label{width:200px;flex-shrink:0;font-size:13px;color:var(--text);line-height:1.3;}
  .bar-track{flex:1;height:10px;background:var(--line);border-radius:6px;overflow:hidden;}
  .bar-fill{height:100%;border-radius:6px;transition:width .6s;}
  .bar-fill.s3{background:var(--score3);}.bar-fill.s4{background:var(--score4);}.bar-fill.s5{background:var(--score5);}
  .bar-val{width:24px;text-align:right;font-family:'JetBrains Mono',monospace;font-weight:700;font-size:14px;flex-shrink:0;}
  /* Chart canvas wrapper */
  .chart-wrap{background:var(--bg-card);border:1px solid var(--line);border-radius:20px;padding:22px 22px 16px;margin:24px 0 8px;}
  .chart-wrap h3{font-family:'Space Grotesk',sans-serif;font-size:14px;font-weight:600;color:var(--text-dim);margin-bottom:16px;letter-spacing:0.02em;}
  .chart-wrap canvas{display:block;width:100%;height:auto;}
  /* KPI strip */
  .kpi-strip{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;margin:16px 0;}
  .kpi{background:var(--bg-card);border:1px solid var(--line);border-radius:14px;padding:16px 14px;text-align:center;}
  .kpi .kv{font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;color:var(--accent);}
  .kpi .kl{font-size:11px;color:var(--text-dim);margin-top:4px;line-height:1.3;}
  /* Footer */
  .footer-actions{margin-top:32px;display:flex;gap:12px;flex-wrap:wrap;}
  /* PIN */
  .pin-box{background:var(--bg-card);border:1px solid var(--line);border-radius:20px;padding:34px 28px;max-width:360px;margin-top:40px;}
  .pin-box input{width:100%;background:var(--bg);border:1.5px solid var(--line);border-radius:12px;padding:14px 16px;color:var(--text);font-size:20px;font-family:'JetBrains Mono',monospace;letter-spacing:0.3em;text-align:center;margin:18px 0;}
  .pin-box input:focus{outline:none;border-color:var(--accent);}
  .pin-error{color:#FF8A8A;font-size:13px;margin-bottom:8px;min-height:18px;}
  /* Dashboard */
  .team-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:14px;margin-top:26px;}
  .team-card{background:var(--bg-card);border:1px solid var(--line);border-radius:16px;padding:20px;cursor:pointer;transition:border-color .15s,transform .15s;}
  .team-card:hover{border-color:var(--accent);transform:translateY(-2px);}
  .team-card .num{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--text-dim);}
  .team-card .name{font-family:'Space Grotesk',sans-serif;font-size:17px;font-weight:600;margin:8px 0 10px;}
  .team-card .score-pill{display:inline-flex;align-items:center;gap:8px;font-family:'JetBrains Mono',monospace;font-size:13px;color:var(--accent);}
  .team-card .score-pill.pending{color:var(--text-dim);}
  /* Export panel */
  .export-panel{background:var(--bg-card);border:1px solid var(--line);border-radius:18px;padding:22px;margin:24px 0 8px;}
  .export-panel h3{font-size:15px;margin-bottom:4px;}
  .export-hint{color:var(--text-dim);font-size:13px;margin-bottom:16px;}
  .check-list{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:18px;}
  .check-item{display:flex;align-items:center;gap:8px;background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:9px 14px;font-size:13px;cursor:pointer;font-family:'Inter',sans-serif;}
  .check-item input{accent-color:var(--accent);width:15px;height:15px;}
  .check-item.disabled{opacity:0.4;cursor:default;}
  .export-buttons{display:flex;gap:12px;flex-wrap:wrap;}
  .export-buttons .btn{font-size:13px;padding:11px 20px;}
  .back-link{background:none;border:none;color:var(--text-dim);font-family:'Space Grotesk',sans-serif;font-size:14px;font-weight:600;cursor:pointer;padding:0;margin-bottom:18px;display:inline-flex;align-items:center;gap:6px;}
  .back-link:hover{color:var(--accent);}
  .intro-card{background:var(--bg-card);border:1px solid var(--line);border-radius:22px;padding:32px 28px;margin-top:18px;}
  .intro-card ul{margin:18px 0 24px 0;padding-left:0;list-style:none;}
  .intro-card li{display:flex;gap:12px;margin-bottom:14px;color:var(--text-dim);font-size:14px;line-height:1.5;}
  .intro-card li .ico{font-family:'JetBrains Mono',monospace;color:var(--accent);font-weight:700;flex-shrink:0;}
  @media(max-width:520px){.bar-label{width:120px;font-size:12px;}.trail-label{display:none;}.total-number{font-size:40px;}}
</style>
</head>
<body>
<div id="app"></div>
<script>
/* ============ DATA ============ */
const CONCEPTS=[
  {title:"Orientación a Resultados",label:"Resultados",options:[
    {score:3,text:"Se compromete con su logro y toma riesgos calculados. Piensa y optimiza variables."},
    {score:4,text:"Demuestra estar un paso adelante, coordinando sus acciones con los objetivos de otros, para acelerar el logro de resultados."},
    {score:5,text:"Promueve y comparte buenas prácticas para facilitar el alcance de resultados o potenciar la mejora continua."}]},
  {title:"Trabajo en Equipo y Colaborativo",label:"Equipo",options:[
    {score:3,text:"Valora las opiniones y experiencia de los demás."},
    {score:4,text:"Desarrolla el espíritu de equipo, promueve la colaboración y la construcción de manera interdisciplinaria."},
    {score:5,text:"Fomenta la modalidad de trabajo en equipo como herramienta de alto impacto para el abordaje de los distintos temas apostando a la sinergia, cohesión y mirada compartida."}]},
  {title:"Transformación e Innovación",label:"Innovación",options:[
    {score:3,text:"Contribuye, acepta la implementación de los cambios creando una visión positiva."},
    {score:4,text:"Evalúa el entorno para anticiparse a los cambios y generar nuevas soluciones."},
    {score:5,text:"Es un referente y promueve la adaptación al cambio entre colaboradores."}]},
  {title:"Aprendizaje Continuo",label:"Aprendizaje",options:[
    {score:3,text:"Capitaliza las experiencias exitosas y los errores propios y de los demás, respondiendo de manera novedosa y ágil frente a situaciones desafiantes."},
    {score:4,text:"Demuestra una marcada tendencia a aprender constantemente, contagiando su innovación y agilidad en nuevas formas de hacer las cosas."},
    {score:5,text:"Es un referente en formas innovadoras de resolver situaciones nuevas y capitalizar los errores, generando y compartiendo conocimiento de manera constante."}]},
  {title:"Orientación al Cliente",label:"Cliente",options:[
    {score:3,text:"Piensa cómo adecuar respuestas, soluciones y alternativas en función de la necesidad del cliente."},
    {score:4,text:"Crea experiencias memorables de servicio en cada interacción, en cada momento de verdad."},
    {score:5,text:"Es un referente tanto interno como externo en experiencias memorables y fidelización de clientes."}]},
  {title:"Planificación y Organización",label:"Planificación",options:[
    {score:3,text:"Se involucra en la planificación y organiza su trabajo de una manera fluida."},
    {score:4,text:"Hace de la planificación su rutina de efectividad, utilizando información correcta y experiencias del pasado."},
    {score:5,text:"Promueve y entrena a otros en su metodología de planificación efectiva."}]}
];
const COLLABORATORS=["Colaborador 1","Colaborador 2","Colaborador 3","Colaborador 4","Colaborador 5"];
const MANAGER_PIN="2026";
const MAX=CONCEPTS.length*5; // 30

/* ============ STORAGE ============ */
let _mem={};
async function sGet(k){try{if(window.storage){const r=await window.storage.get(k,true);return r?JSON.parse(r.value):null;}}catch(e){}return _mem[k]||null;}
async function sSet(k,v){try{if(window.storage){await window.storage.set(k,JSON.stringify(v),true);return;}}catch(e){}_mem[k]=v;}
async function sDel(k){try{if(window.storage)await window.storage.delete(k,true);}catch(e){}delete _mem[k];}

/* ============ STATE ============ */
let S={screen:'home',currentUser:null,conceptIndex:0,order:[],answers:[],selectedOption:null,revealed:false,pinError:'',detailUser:null,detailResult:null,dashboardResults:{}};
function shuffle(a){a=[...a];for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[a[i],a[j]]=[a[j],a[i]];}return a;}

/* ============ CHART ENGINE (Canvas, no libs) ============ */
function scoreColor(s){return s<=3?'#6FB7E0':s===4?'#FFB454':'#6FD9A0';}
function lvl(s){return s===3?'Básico':s===4?'Destacado':'Referente';}
function pct(v,m){return Math.round((v/m)*100)+'%';}

/**
 * Draw a bar chart on a canvas element.
 * labels: string[], values: number[], max: number, barColors: string[]
 */
function drawBarChart(canvas, labels, values, max, barColors){
  const DPR=window.devicePixelRatio||1;
  const W=canvas.offsetWidth||700, H=260;
  canvas.width=W*DPR; canvas.height=H*DPR;
  canvas.style.height=H+'px';
  const ctx=canvas.getContext('2d');
  ctx.scale(DPR,DPR);

  const PAD={top:42,right:16,bottom:60,left:40};
  const cW=W-PAD.left-PAD.right;
  const cH=H-PAD.top-PAD.bottom;
  const n=values.length;
  const slotW=cW/n;
  const barW=Math.min(slotW*0.55,58);

  // background
  ctx.fillStyle='#182944'; ctx.fillRect(0,0,W,H);

  // grid lines + Y labels (0,1,2,3,4,5)
  for(let v=0;v<=max;v++){
    const y=PAD.top+cH-(v/max)*cH;
    ctx.strokeStyle= v===0?'#2C4368':'#1E3456';
    ctx.lineWidth=v===0?1.5:1;
    ctx.beginPath();ctx.moveTo(PAD.left,y);ctx.lineTo(PAD.left+cW,y);ctx.stroke();
    if(v%1===0){
      ctx.fillStyle='#8FA3C4';ctx.font=`11px 'JetBrains Mono', monospace`;
      ctx.textAlign='right';ctx.fillText(v,PAD.left-6,y+4);
    }
  }

  // bars
  values.forEach((val,i)=>{
    const barH=(val/max)*cH;
    const x=PAD.left+slotW*i+slotW/2-barW/2;
    const y=PAD.top+cH-barH;
    const r=Math.min(6,barW/4);
    ctx.fillStyle=barColors[i]||'#FFB454';
    // rounded top
    ctx.beginPath();
    ctx.moveTo(x+r,y);ctx.lineTo(x+barW-r,y);
    ctx.quadraticCurveTo(x+barW,y,x+barW,y+r);
    ctx.lineTo(x+barW,y+barH);ctx.lineTo(x,y+barH);ctx.lineTo(x,y+r);
    ctx.quadraticCurveTo(x,y,x+r,y);
    ctx.closePath();ctx.fill();

    // value label on top
    ctx.fillStyle='#EFEAE0';ctx.font=`bold 12px 'JetBrains Mono', monospace`;
    ctx.textAlign='center';
    const dispVal=typeof val==='number'&&!Number.isInteger(val)?val.toFixed(2):val;
    ctx.fillText(dispVal,x+barW/2,y-7);

    // X label (wrap)
    ctx.fillStyle='#8FA3C4';ctx.font=`11px Inter, sans-serif`;
    const words=labels[i].split(' ');
    let line='',ly=PAD.top+cH+14;
    words.forEach(w=>{
      const t=line?line+' '+w:w;
      if(t.length>10&&line){ctx.fillText(line,x+barW/2,ly);line=w;ly+=13;}
      else line=t;
    });
    ctx.fillText(line,x+barW/2,ly);
  });
}

/* Mount a chart into a .chart-wrap div created inside a parent selector */
function mountChart(parentEl, title, labels, values, max, barColors){
  const wrap=document.createElement('div');
  wrap.className='chart-wrap';
  wrap.innerHTML=`<h3>${title}</h3><canvas></canvas>`;
  parentEl.appendChild(wrap);
  const canvas=wrap.querySelector('canvas');
  // wait for layout
  requestAnimationFrame(()=>drawBarChart(canvas,labels,values,max,barColors));
  // redraw on resize
  const obs=new ResizeObserver(()=>drawBarChart(canvas,labels,values,max,barColors));
  obs.observe(canvas);
}

/* ============ RENDER ============ */
const app=document.getElementById('app');
function render(){
  if(S.screen==='home')return renderHome();
  if(S.screen==='intro')return renderIntro();
  if(S.screen==='game')return renderGame();
  if(S.screen==='results')return renderResults();
  if(S.screen==='pin')return renderPin();
  if(S.screen==='dashboard')return renderDashboard();
  if(S.screen==='detail')return renderDetail();
}
const topBar=()=>`<div class="top"><div class="brand">Ruta de <b>Competencias</b></div></div>`;

/* HOME */
async function renderHome(){
  let cards='';
  for(let i=0;i<COLLABORATORS.length;i++){
    const n=COLLABORATORS[i],res=await sGet('res:'+n);
    cards+=`<div class="pick-card" onclick="goUser('${n.replace(/'/g,"\\'")}')"><span class="num">0${i+1}</span><div class="name">${n}</div>${res?`<span class="status">Completado ✓</span>`:`<span class="status pending">Pendiente</span>`}</div>`;
  }
  app.innerHTML=`${topBar()}<div class="hero"><span class="eyebrow">Autoevaluación de competencias</span><h1>Una estación por cada competencia clave.</h1><p>Elegí tu nombre para recorrer las 6 estaciones. En cada una vas a leer tres formas de actuar y elegir la que más te representa hoy.</p></div><div class="grid-collab">${cards}</div><div class="leader-row"><p>¿Sos el líder del equipo? Ingresá con tu PIN para ver el avance y los resultados de cada persona.</p><button class="btn btn-ghost" onclick="goPin()">Acceso del líder</button></div>`;
}
window.goUser=async function(n){S.currentUser=n;const res=await sGet('res:'+n);if(res){S.answers=res.answers;S.screen='results';}else S.screen='intro';render();};
window.goPin=function(){S.screen='pin';S.pinError='';render();};
window.goHome=function(){S.currentUser=null;S.screen='home';render();};

/* INTRO */
function renderIntro(){
  app.innerHTML=`${topBar()}<button class="back-link" onclick="goHome()">← Volver</button><div class="hero" style="padding-top:6px"><span class="eyebrow">${S.currentUser}</span><h1>Antes de arrancar</h1></div><div class="intro-card"><ul><li><span class="ico">01</span>Vas a recorrer 6 estaciones, una por cada competencia.</li><li><span class="ico">02</span>En cada estación hay 3 formas de actuar. Elegí la que mejor te represente hoy, sin pensar en un número.</li><li><span class="ico">03</span>Al elegir, vas a ver el puntaje que corresponde a esa forma de actuar (3, 4 o 5).</li><li><span class="ico">04</span>Al final vas a ver tu propio resultado. Tu líder podrá verlo también, pero tus compañeros no.</li></ul><button class="btn btn-primary" onclick="startGame()">Comenzar recorrido</button></div>`;
}
window.startGame=function(){S.conceptIndex=0;S.answers=[];S.order=CONCEPTS.map(()=>shuffle([0,1,2]));S.selectedOption=null;S.revealed=false;S.screen='game';render();};

/* GAME */
function renderGame(){
  const idx=S.conceptIndex,C=CONCEPTS[idx],ord=S.order[idx];
  let trail='';
  CONCEPTS.forEach((c,i)=>{
    let cls='dot',cnt='0'+(i+1);
    if(i<idx){const sc=S.answers[i].score;cls+=` done s${sc}`;cnt=sc;}
    else if(i===idx)cls+=' current';
    trail+=`<div class="trail-node"><div class="${cls}">${cnt}</div><div class="trail-label">${c.label}</div></div>`;
    if(i<CONCEPTS.length-1)trail+=`<div class="trail-line ${i<idx?'done':''}"></div>`;
  });
  let opts='';
  ord.forEach(oi=>{
    const opt=C.options[oi],sel=S.selectedOption===oi;
    opts+=`<div class="option${S.revealed?' locked':''}${sel?' selected':''}" onclick="${S.revealed?'':`selectOpt(${oi})`}">${opt.text}</div>`;
  });
  let rev='';
  if(S.revealed){const sc=C.options[S.selectedOption].score;rev=`<div class="reveal"><div class="score-badge s${sc}">${sc}</div><div class="reveal-text"><div class="label">Tu puntaje en esta competencia</div><div class="value">Elegiste la forma de actuar correspondiente al nivel <b>${sc}</b>.</div></div></div>`;}
  app.innerHTML=`${topBar()}<div class="trail">${trail}</div><div class="station-card"><span class="eyebrow">Estación ${idx+1} de ${CONCEPTS.length}</span><h2 class="station-title">${C.title}</h2>${opts}${rev}<div class="station-footer"><button class="btn btn-primary" ${S.revealed?'':'disabled'} onclick="nextStation()">${idx===CONCEPTS.length-1?'Finalizar y guardar':'Siguiente estación'}</button></div></div>`;
}
window.selectOpt=function(oi){if(S.revealed)return;S.selectedOption=oi;S.revealed=true;render();};
window.nextStation=async function(){
  if(!S.revealed)return;
  const idx=S.conceptIndex,c=CONCEPTS[idx],sc=c.options[S.selectedOption].score;
  S.answers[idx]={title:c.title,score:sc};
  if(idx===CONCEPTS.length-1){
    const res={answers:S.answers,total:S.answers.reduce((a,b)=>a+b.score,0),date:new Date().toLocaleString('es-AR')};
    await sSet('res:'+S.currentUser,res);S.screen='results';
  }else{S.conceptIndex++;S.selectedOption=null;S.revealed=false;S.screen='game';}
  render();
};

/* RESULTS (personal) */
function scc(s){return s<=3?'s3':s===4?'s4':'s5';}
function renderResults(){
  const total=S.answers.reduce((a,b)=>a+b.score,0);
  let bars='';
  S.answers.forEach(a=>{const p=Math.round((a.score/5)*100);bars+=`<div class="bar-row"><div class="bar-label">${a.title}</div><div class="bar-track"><div class="bar-fill ${scc(a.score)}" style="width:${p}%"></div></div><div class="bar-val">${a.score}</div></div>`;});
  app.innerHTML=`${topBar()}<div class="results-head"><span class="eyebrow">${S.currentUser}</span><h1>Recorrido completo</h1><p>Este es tu resultado personal. Tu líder puede verlo; tus compañeros no.</p></div><div class="total-box"><div class="total-number">${total}<span style="font-size:22px;color:var(--text-dim)"> / ${MAX}</span></div><div class="total-sub">Suma de los puntajes obtenidos en las 6 competencias evaluadas.</div></div>${bars}<div id="chart-personal"></div><div class="footer-actions"><button class="btn btn-ghost" onclick="goHome()">Volver al inicio</button></div>`;
  mountChart(document.getElementById('chart-personal'),'Puntaje por competencia',S.answers.map(a=>a.title.split(' ')[0]),S.answers.map(a=>a.score),5,S.answers.map(a=>scoreColor(a.score)));
}

/* PIN */
function renderPin(){
  app.innerHTML=`${topBar()}<button class="back-link" onclick="goHome()">← Volver</button><div class="hero" style="padding-top:6px"><span class="eyebrow">Acceso del líder</span><h1>Ingresá tu PIN</h1></div><div class="pin-box"><div class="pin-error">${S.pinError}</div><input id="pinInput" type="password" inputmode="numeric" maxlength="6" placeholder="••••" autofocus/><button class="btn btn-primary" style="width:100%" onclick="checkPin()">Entrar</button></div>`;
  const inp=document.getElementById('pinInput');
  inp.addEventListener('keydown',e=>{if(e.key==='Enter')checkPin();});inp.focus();
}
window.checkPin=function(){const v=document.getElementById('pinInput').value;if(v===MANAGER_PIN){S.pinError='';S.screen='dashboard';render();}else{S.pinError='PIN incorrecto. Probá de nuevo.';render();}};

/* DASHBOARD */
async function renderDashboard(){
  const rm={};let cards='';
  for(let i=0;i<COLLABORATORS.length;i++){
    const n=COLLABORATORS[i],res=await sGet('res:'+n);
    rm[n]=res;
    const pill=res?`<span class="score-pill">${res.total} / ${MAX} · completado</span>`:`<span class="score-pill pending">Pendiente</span>`;
    cards+=`<div class="team-card" onclick="goDetail('${n.replace(/'/g,"\\'")}')"><span class="num">0${i+1}</span><div class="name">${n}</div>${pill}</div>`;
  }
  S.dashboardResults=rm;
  const checks=COLLABORATORS.map(n=>{const ok=!!rm[n];return`<label class="check-item ${ok?'':'disabled'}"><input type="checkbox" class="export-check" value="${n.replace(/"/g,'&quot;')}" ${ok?'checked':'disabled'}>${n}${ok?'':' (pendiente)'}</label>`;}).join('');
  app.innerHTML=`${topBar()}<button class="back-link" onclick="goHome()">← Volver al inicio</button><div class="hero" style="padding-top:6px"><span class="eyebrow">Vista del líder</span><h1>Estado del equipo</h1><p>Seleccioná a una persona para ver el detalle, o marcá colaboradores para exportar.</p></div><div class="export-panel"><h3>Exportar resultados</h3><div class="export-hint">Elegí uno o varios colaboradores (completados). Un colaborador = informe individual. Varios = comparativo de equipo con promedios.</div><div class="check-list">${checks}</div><div class="export-buttons"><button class="btn btn-ghost" onclick="doExportXLSX()">⬇ Datos en Excel</button><button class="btn btn-primary" onclick="doExportDashboard()">⬇ Dashboard con Gráficos</button></div></div><div id="comp-chart"></div><div class="team-grid">${cards}</div>`;

  // Comparative chart for completed collaborators
  const done=COLLABORATORS.filter(n=>rm[n]);
  if(done.length>1){
    mountChart(document.getElementById('comp-chart'),'Puntaje total por colaborador',done,done.map(n=>rm[n].total),MAX,done.map(()=>'#FFB454'));
  }
}
window.goDetail=async function(n){S.detailUser=n;S.detailResult=await sGet('res:'+n);S.screen='detail';render();};
window.goDashboard=function(){S.screen='dashboard';render();};
window.resetUser=async function(n){if(!confirm(`¿Seguro que querés reiniciar la autoevaluación de ${n}?`))return;await sDel('res:'+n);S.detailResult=null;render();};

/* DETAIL (manager) */
function renderDetail(){
  const n=S.detailUser,res=S.detailResult;
  if(!res){app.innerHTML=`${topBar()}<button class="back-link" onclick="goDashboard()">← Volver al equipo</button><div class="hero" style="padding-top:6px"><span class="eyebrow">${n}</span><h1>Todavía no completó su autoevaluación</h1></div>`;return;}
  let bars='';
  res.answers.forEach(a=>{const p=Math.round((a.score/5)*100);bars+=`<div class="bar-row"><div class="bar-label">${a.title}</div><div class="bar-track"><div class="bar-fill ${scc(a.score)}" style="width:${p}%"></div></div><div class="bar-val">${a.score}</div></div>`;});
  const topS=Math.max(...res.answers.map(a=>a.score));
  const sn=n.replace(/'/g,"\\'");
  app.innerHTML=`${topBar()}<button class="back-link" onclick="goDashboard()">← Volver al equipo</button>
    <div class="results-head"><span class="eyebrow">${n}</span><h1>Resultado de la autoevaluación</h1><p>Completado el ${res.date}.</p></div>
    <div class="kpi-strip">
      <div class="kpi"><div class="kv">${res.total}</div><div class="kl">Puntaje total<br>de ${MAX} posibles</div></div>
      <div class="kpi"><div class="kv">${pct(res.total,MAX)}</div><div class="kl">Del puntaje<br>máximo</div></div>
      <div class="kpi"><div class="kv" style="font-size:18px">${lvl(topS)}</div><div class="kl">Nivel<br>predominante</div></div>
    </div>
    ${bars}
    <div id="chart-detail"></div>
    <div class="footer-actions">
      <button class="btn btn-ghost" onclick="goDashboard()">Volver al equipo</button>
      <button class="btn btn-ghost" onclick="exportIndividualExcel('${sn}',S.detailResult)">⬇ Datos en Excel</button>
      <button class="btn btn-primary" onclick="exportIndividualDashboard('${sn}',S.detailResult)">⬇ Dashboard con Gráficos</button>
      <button class="btn btn-danger" onclick="resetUser('${sn}')">Reiniciar</button>
    </div>`;
  mountChart(document.getElementById('chart-detail'),'Puntaje por competencia',res.answers.map(a=>a.title.split(' ').slice(0,2).join(' ')),res.answers.map(a=>a.score),5,res.answers.map(a=>scoreColor(a.score)));
}

/* ============================================================
   CHART-TO-IMAGE for HTML Dashboard export
   ============================================================ */
function chartToBase64(labels, values, maxVal, barColors, W=700, H=300){
  const canvas=document.createElement('canvas');
  canvas.width=W;canvas.height=H;canvas.style.width=W+'px';canvas.style.height=H+'px';
  // Use the same drawing logic but synchronously
  drawBarChart(canvas,labels,values,maxVal,barColors);
  return canvas.toDataURL('image/png');
}

/* ============================================================
   HTML DASHBOARD EXPORT  (self-contained, opens in browser)
   ============================================================ */
function buildDashboardHTML(name, res, extraCharts){
  const total=res.total, topS=Math.max(...res.answers.map(a=>a.score));
  const chartImg=chartToBase64(res.answers.map(a=>a.title.split(' ').slice(0,2).join(' ')),res.answers.map(a=>a.score),5,res.answers.map(a=>scoreColor(a.score)));
  let rows=res.answers.map(a=>`<tr><td>${a.title}</td><td style="text-align:center;font-weight:700;color:${scoreColor(a.score)}">${a.score}</td><td style="text-align:center">${lvl(a.score)}</td><td>${pct(a.score,5)}</td></tr>`).join('');
  let extra=extraCharts||'';
  return`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Dashboard — ${name}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',Arial,sans-serif;background:#0F1B2D;color:#EFEAE0;padding:32px 24px;}
h1{font-size:28px;font-weight:700;color:#FFB454;margin-bottom:4px;}
.sub{color:#8FA3C4;font-size:14px;margin-bottom:24px;}
.kpi-row{display:flex;gap:16px;margin-bottom:28px;flex-wrap:wrap;}
.kpi{background:#182944;border-radius:14px;padding:18px 22px;flex:1;min-width:120px;border:1px solid #2C4368;}
.kpi .v{font-size:32px;font-weight:700;color:#FFB454;font-family:'Courier New',monospace;}
.kpi .l{font-size:12px;color:#8FA3C4;margin-top:4px;}
.section{background:#182944;border-radius:16px;padding:22px;margin-bottom:20px;border:1px solid #2C4368;}
.section h2{font-size:15px;font-weight:600;color:#8FA3C4;margin-bottom:16px;letter-spacing:.04em;text-transform:uppercase;}
.section img{width:100%;border-radius:10px;display:block;}
table{width:100%;border-collapse:collapse;}
th{background:#112238;color:#8FA3C4;font-size:11px;letter-spacing:.08em;text-transform:uppercase;padding:10px 12px;text-align:left;}
td{padding:12px 12px;font-size:13px;border-bottom:1px solid #1E3456;}
.footer{color:#2C4368;font-size:11px;text-align:center;margin-top:20px;}
@media print{body{background:#fff;color:#000;} .section{border:1px solid #ccc;} h1,.kpi .v{color:#b07a00;} th{background:#eee;color:#555;} .sub,.kpi .l{color:#666;} td{border-bottom:1px solid #ddd;}}
</style></head><body>
<h1>Autoevaluación de Competencias</h1>
<div class="sub">Colaborador: <b>${name}</b> &nbsp;·&nbsp; Fecha de realización: ${res.date}</div>
<div class="kpi-row">
  <div class="kpi"><div class="v">${total}</div><div class="l">Puntaje total (de ${MAX})</div></div>
  <div class="kpi"><div class="v">${pct(total,MAX)}</div><div class="l">Del máximo posible</div></div>
  <div class="kpi"><div class="v" style="font-size:20px">${lvl(topS)}</div><div class="l">Nivel predominante</div></div>
</div>
<div class="section"><h2>Gráfico de Competencias</h2><img src="${chartImg}" alt="Gráfico de competencias"></div>
<div class="section"><h2>Detalle por Competencia</h2>
<table><thead><tr><th>Competencia</th><th>Puntaje</th><th>Nivel</th><th>% del máximo</th></tr></thead>
<tbody>${rows}</tbody></table></div>
${extra}
<div class="footer">Generado el ${new Date().toLocaleString('es-AR')} · Sistema de Autoevaluación de Equipo</div>
</body></html>`;
}

function downloadHTML(html, filename){
  const blob=new Blob([html],{type:'text/html;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download=filename;a.click();
  setTimeout(()=>URL.revokeObjectURL(url),3000);
}

window.exportIndividualDashboard=function(name,res){
  if(!res){alert(`${name} todavía no completó su autoevaluación.`);return;}
  const html=buildDashboardHTML(name,res,'');
  downloadHTML(html,`Dashboard_${name.replace(/\s+/g,'_')}.html`);
};

window.doExportDashboard=function(){
  const names=Array.from(document.querySelectorAll('.export-check:checked')).map(c=>c.value);
  if(!names.length){alert('Seleccioná al menos un colaborador completado.');return;}
  const rm=S.dashboardResults;
  if(names.length===1){exportIndividualDashboard(names[0],rm[names[0]]);return;}

  // Comparative dashboard
  const avgs=CONCEPTS.map((_,i)=>names.reduce((s,n)=>s+rm[n].answers[i].score,0)/names.length);
  const avgTotal=avgs.reduce((a,b)=>a+b,0);

  // chart 1: totals per person
  const totalsImg=chartToBase64(names,names.map(n=>rm[n].total),MAX,names.map(()=>'#FFB454'));
  // chart 2: team averages per concept
  const avgsImg=chartToBase64(CONCEPTS.map(c=>c.label),avgs.map(v=>Number(v.toFixed(2))),5,avgs.map(v=>scoreColor(Math.round(v))));

  let compRows=CONCEPTS.map((c,i)=>`<tr><td>${c.title}</td>${names.map(n=>`<td style="text-align:center;font-weight:700;color:${scoreColor(rm[n].answers[i].score)}">${rm[n].answers[i].score}</td>`).join('')}<td style="text-align:center;font-weight:700;color:${scoreColor(Math.round(avgs[i]))}">${Number(avgs[i].toFixed(2))}</td></tr>`).join('');
  let totalRow=`<tr style="border-top:2px solid #2C4368"><td><b>TOTAL</b></td>${names.map(n=>`<td style="text-align:center;font-weight:700;color:#FFB454">${rm[n].total}</td>`).join('')}<td style="text-align:center;font-weight:700;color:#FFB454">${Number(avgTotal.toFixed(2))}</td></tr>`;
  let pctRow=`<tr><td style="color:#8FA3C4">% del máximo</td>${names.map(n=>`<td style="text-align:center;color:#8FA3C4">${pct(rm[n].total,MAX)}</td>`).join('')}<td style="text-align:center;color:#8FA3C4">${pct(avgTotal,MAX)}</td></tr>`;

  const html=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Dashboard Comparativo de Equipo</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',Arial,sans-serif;background:#0F1B2D;color:#EFEAE0;padding:32px 24px;}
h1{font-size:28px;font-weight:700;color:#FFB454;margin-bottom:4px;}
.sub{color:#8FA3C4;font-size:14px;margin-bottom:24px;}
.kpi-row{display:flex;gap:16px;margin-bottom:28px;flex-wrap:wrap;}
.kpi{background:#182944;border-radius:14px;padding:18px 22px;flex:1;min-width:120px;border:1px solid #2C4368;}
.kpi .v{font-size:28px;font-weight:700;color:#FFB454;font-family:'Courier New',monospace;}
.kpi .l{font-size:12px;color:#8FA3C4;margin-top:4px;}
.section{background:#182944;border-radius:16px;padding:22px;margin-bottom:20px;border:1px solid #2C4368;}
.section h2{font-size:15px;font-weight:600;color:#8FA3C4;margin-bottom:16px;letter-spacing:.04em;text-transform:uppercase;}
.section img{width:100%;border-radius:10px;display:block;}
table{width:100%;border-collapse:collapse;overflow-x:auto;display:block;}
th{background:#112238;color:#8FA3C4;font-size:11px;letter-spacing:.06em;text-transform:uppercase;padding:10px 12px;text-align:center;white-space:nowrap;}
th:first-child{text-align:left;}
td{padding:12px 10px;font-size:13px;border-bottom:1px solid #1E3456;white-space:nowrap;}
.footer{color:#2C4368;font-size:11px;text-align:center;margin-top:20px;}
@media print{body{background:#fff;color:#000;}.section{border:1px solid #ccc;}h1,.kpi .v{color:#b07a00;}th{background:#eee;color:#555;}.sub,.kpi .l{color:#666;}td{border-bottom:1px solid #ddd;}}
</style></head><body>
<h1>Dashboard Comparativo de Equipo</h1>
<div class="sub">Autoevaluación de Competencias · ${names.length} colaboradores · Generado el ${new Date().toLocaleString('es-AR')}</div>
<div class="kpi-row">
  <div class="kpi"><div class="v">${names.length}</div><div class="l">Colaboradores evaluados</div></div>
  <div class="kpi"><div class="v">${Number(avgTotal.toFixed(1))}</div><div class="l">Promedio total del equipo</div></div>
  <div class="kpi"><div class="v">${pct(avgTotal,MAX)}</div><div class="l">% del máximo (${MAX} pts)</div></div>
</div>
<div class="section"><h2>Puntaje total por colaborador</h2><img src="${totalsImg}"></div>
<div class="section"><h2>Promedio del equipo por competencia</h2><img src="${avgsImg}"></div>
<div class="section"><h2>Tabla comparativa</h2>
<table><thead><tr><th>Competencia</th>${names.map(n=>`<th>${n}</th>`).join('')}<th>PROMEDIO</th></tr></thead>
<tbody>${compRows}${totalRow}${pctRow}</tbody></table></div>
<div class="footer">Sistema de Autoevaluación de Equipo</div>
</body></html>`;
  downloadHTML(html,'Dashboard_Comparativo_Equipo.html');
};

/* ============================================================
   EXCEL EXPORT (datos — SheetJS desde cdnjs)
   ============================================================ */
function makeSheet(aoa,colW,rowH,merges,freeze){
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  if(colW)ws['!cols']=colW.map(w=>({wch:w}));
  if(rowH)ws['!rows']=rowH.map(h=>h?{hpt:h}:{});
  if(merges)ws['!merges']=merges;
  if(freeze)ws['!freeze']=freeze;
  return ws;
}
function bar(v,max){const f=Math.round((v/max)*10);return'█'.repeat(f)+'░'.repeat(10-f);}

window.exportIndividualExcel=function(name,res){
  if(!res){alert(`${name} todavía no completó su autoevaluación.`);return;}
  const total=res.total,p=pct(total,MAX),topS=Math.max(...res.answers.map(a=>a.score));
  const topC=res.answers.filter(a=>a.score===topS).map(a=>a.title).join(' / ');
  const aoa=[
    ['AUTOEVALUACIÓN DE COMPETENCIAS','','',''],
    [`Colaborador: ${name}   |   Fecha de realización: ${res.date}`,'','',''],
    ['','','',''],
    ['RESUMEN','','',''],
    ['Puntaje total',total,'Porcentaje',p],
    ['Nivel predominante',lvl(topS),'Mayor fortaleza',topC],
    ['','','',''],
    ['COMPETENCIA','PUNTAJE','NIVEL','PROGRESO (escala 3-5)'],
    ...res.answers.map(a=>[a.title,a.score,lvl(a.score),bar(a.score-2,3)]),
    ['','','',''],
    ['TOTAL',`${total} / ${MAX}`,p,''],
  ];
  const ws=makeSheet(aoa,[44,10,18,24],[22,16,8,16,18,18,8,16,...res.answers.map(()=>20),8,18],[{s:{r:0,c:0},e:{r:0,c:3}},{s:{r:1,c:0},e:{r:1,c:3}},{s:{r:3,c:0},e:{r:3,c:3}},{s:{r:aoa.length-1,c:3},e:{r:aoa.length-1,c:3}}],{xSplit:0,ySplit:8});
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Datos');
  XLSX.writeFile(wb,`Datos_${name.replace(/\s+/g,'_')}.xlsx`);
};

window.doExportXLSX=function(){
  const names=Array.from(document.querySelectorAll('.export-check:checked')).map(c=>c.value);
  if(!names.length){alert('Seleccioná al menos un colaborador completado.');return;}
  const rm=S.dashboardResults;
  if(names.length===1){exportIndividualExcel(names[0],rm[names[0]]);return;}

  const wb=XLSX.utils.book_new();
  const now=new Date().toLocaleString('es-AR');
  const avgs=CONCEPTS.map((_,i)=>names.reduce((s,n)=>s+rm[n].answers[i].score,0)/names.length);
  const avgTotal=avgs.reduce((a,b)=>a+b,0);

  // Sheet 1: Comparativo
  const compAoa=[
    ['COMPARATIVO DE EQUIPO — AUTOEVALUACIÓN',...names.map(()=>'')],
    [`Generado el ${now}`,...names.map(()=>'')],
    ['',...names.map(()=>'')],
    ['TOTALES','','',''],
    ['Colaborador',...names],
    ['Puntaje total',...names.map(n=>rm[n].total)],
    ['% del máximo',...names.map(n=>pct(rm[n].total,MAX))],
    ['',...names.map(()=>'')],
    ['DETALLE POR COMPETENCIA',...names.map(()=>'')],
    ['Competencia',...names],
    ...CONCEPTS.map((c,i)=>[c.title,...names.map(n=>rm[n].answers[i].score)]),
    ['',...names.map(()=>'')],
    ['TOTAL',...names.map(n=>rm[n].total)],
  ];
  const wsComp=makeSheet(compAoa,[40,...names.map(()=>14)],[22,14,8,14,16,22,18,8,16,16,...CONCEPTS.map(()=>20),8,20],[{s:{r:0,c:0},e:{r:0,c:names.length}},{s:{r:1,c:0},e:{r:1,c:names.length}},{s:{r:3,c:0},e:{r:3,c:names.length}},{s:{r:8,c:0},e:{r:8,c:names.length}}],{xSplit:1,ySplit:10});
  XLSX.utils.book_append_sheet(wb,wsComp,'Comparativo');

  // Sheet 2: Promedios
  const avgAoa=[
    ['PROMEDIO DEL EQUIPO POR COMPETENCIA','',''],
    [`${names.length} colaboradores · ${now}`,'',''],
    ['','',''],
    ['Promedio total',Number(avgTotal.toFixed(2)),pct(avgTotal,MAX)],
    ['','',''],
    ['Competencia','Promedio','Nivel'],
    ...CONCEPTS.map((c,i)=>[c.title,Number(avgs[i].toFixed(2)),lvl(Math.max(3,Math.min(5,Math.round(avgs[i]))))]),
    ['','',''],
    ['PROMEDIO GENERAL',Number(avgTotal.toFixed(2)),pct(avgTotal,MAX)],
  ];
  const wsAvg=makeSheet(avgAoa,[40,14,20],[22,14,8,20,8,16,...CONCEPTS.map(()=>20),8,20],[{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}],null);
  XLSX.utils.book_append_sheet(wb,wsAvg,'Promedios');

  // Individual sheets
  names.forEach(n=>{
    const r=rm[n],p2=pct(r.total,MAX),topS=Math.max(...r.answers.map(a=>a.score));
    const ia=[
      [`${n.toUpperCase()}`,'',''],
      [`Fecha: ${r.date}`,'',''],
      ['','',''],
      ['Puntaje',r.total,p2],
      ['Nivel',lvl(topS),''],
      ['','',''],
      ['Competencia','Puntaje','Nivel'],
      ...r.answers.map(a=>[a.title,a.score,lvl(a.score)]),
      ['','',''],
      ['TOTAL',`${r.total}/${MAX}`,p2],
    ];
    const iws=makeSheet(ia,[40,12,18],[20,14,8,16,16,8,16,...r.answers.map(()=>20),8,18],[{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}],null);
    XLSX.utils.book_append_sheet(wb,iws,n.replace(/\s+/g,'_').substring(0,31));
  });

  XLSX.writeFile(wb,'Datos_Comparativo_Equipo.xlsx');
};

/* ============ INIT ============ */
render();
</script>
</body>
</html>
