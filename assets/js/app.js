// ════════════════════════════════════════════════════
// CONFIG
// ════════════════════════════════════════════════════
const MSAL_CONFIG = {
  auth: {
    clientId: "9c04551b-3267-4a30-a212-ec9143838eb6",
    authority: "https://login.microsoftonline.com/deb13623-4618-4c6b-a268-3d5e4dec1c30",
    // Se toma de la URL donde esté publicada la encuesta. Debe estar registrada
    // EXACTAMENTE igual como Redirect URI tipo SPA en Entra ID.
    redirectUri: window.location.origin + window.location.pathname.replace(/index\.html$/, ""),
  },
  cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: true },
  system: { allowNativeBroker: false }
};
// Todos los permisos desde el login — ya tienen admin consent
const SCOPES = ["User.Read"];

const VALID_DOMAINS = [
  "celaque.net",
  "administracion.hn"
];

// Webhook de n8n que guarda ESTA encuesta (flujo nuevo, distinto al de mayITa Proyectos).
const N8N_WEBHOOK_URL = "PENDIENTE_URL_WEBHOOK_N8N";

const N8N_RANKING_SAVE_URL =
  "https://celaque.app.n8n.cloud/webhook/mayita/ranking";

const N8N_RANKING_TOP_URL =
  "https://celaque.app.n8n.cloud/webhook/mayita/ranking/top";

const APP_VERSION = "2.0";

// ════════════════════════════════════════════════════
// DATA — RETROALIMENTACIÓN DE PROCESOS
// ════════════════════════════════════════════════════

const SURVEY_NAME = "Retroalimentación de procesos";

// ── Preguntas ───────────────────────────────────────────────────────────
// Tipos soportados:
//   yesno  → botones Sí / No
//   single → una sola opción (avanza automáticamente al elegir)
//   multi  → varias opciones (requiere al menos una)
//   text   → respuesta abierta (obligatoria)
// Campos opcionales:
//   showIf        → { id, equals } muestra la pregunta solo si otra respuesta coincide
//   otroConTexto  → opción que pide especificar (ej. "Otros")
//   placeholder   → texto de ayuda del campo abierto
//   largo         → campo de texto más alto
// Los "id" se usan como nombres de columna/campo al guardar, evita cambiarlos
// después de que ya existan respuestas guardadas.
const SURVEY_QUESTIONS = [
  {
    id: "induccion",
    type: "yesno",
    text: "¿Recibiste inducción de tus procesos por parte de tu siguiente nivel?"
  },
  {
    id: "siguiente_nivel",
    type: "text",
    text: "¿Quién es tu siguiente nivel? (Nombre y puesto)",
    placeholder: "Ej: Nombre Apellido — Puesto"
  },
  {
    id: "conocimiento_procesos",
    type: "single",
    text: "¿Qué tan bien conoces tus procesos?",
    options: [
      "Poco conocido",
      "Conozco lo básico",
      "Familiarizado",
      "Bien familiarizado",
      "Muy familiarizado"
    ]
  },
  {
    id: "carga_laboral",
    type: "single",
    text: "¿Cómo sientes tu carga laboral actualmente?",
    options: ["Baja", "Media", "Alta"]
  },
  {
    id: "ajusta_tiempo",
    type: "yesno",
    text: "¿Te ajusta el tiempo para finalizar tus tareas en los días programados?"
  },
  {
    id: "por_que_no_ajusta_tiempo",
    type: "text",
    text: "¿Por qué no te ajusta el tiempo?",
    largo: true,
    showIf: { id: "ajusta_tiempo", equals: "No" }
  },
  {
    id: "dias_promedio",
    type: "single",
    text: "En promedio, ¿Cuántos días te tardas en completar tus procesos?",
    options: ["0-1 días", "1-3 días", "5-7 días", "8 o más días"]
  },
  {
    id: "procesos_mas_tiempo",
    type: "text",
    text: "¿Cuáles procesos te toman más tiempo?",
    largo: true
  },
  {
    id: "dia_mayor_carga",
    type: "single",
    text: "¿Qué día de la semana sientes mayor carga laboral?",
    options: ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
  },
  {
    id: "por_que_dia_carga",
    type: "text",
    text: "¿Por qué esos días se sienten con mayor carga?",
    largo: true
  },
  {
    id: "procesos_corridos",
    type: "yesno",
    text: "¿Tienes procesos corridos? (Procesos con más de 5 días)"
  },
  {
    id: "motivo_procesos_corridos",
    type: "single",
    text: "¿Cuál ha sido el principal motivo?",
    options: [
      "Dependencia del proceso",
      "Dependencia Interna",
      "Dependencia departamental",
      "Dependencia externa",
      "Falta de tiempo",
      "Entrenamiento",
      "Proceso montado",
      "Otros"
    ],
    otroConTexto: "Otros",
    showIf: { id: "procesos_corridos", equals: "Sí" }
  },
  {
    id: "desea_modificar_proceso",
    type: "yesno",
    text: "¿Tienes algún proceso que deseas modificar?"
  },
  {
    id: "nombre_proceso_modificar",
    type: "text",
    text: "Nombre del proceso que desea modificar",
    showIf: { id: "desea_modificar_proceso", equals: "Sí" }
  },
  {
    id: "tipo_modificacion",
    type: "multi",
    text: "¿Cuáles modificaciones serían?",
    options: ["Cambio en descripción", "Frecuencia", "Fecha"],
    showIf: { id: "desea_modificar_proceso", equals: "Sí" }
  },
  {
    id: "descripcion_cambio",
    type: "text",
    text: "Describe el cambio deseado",
    largo: true,
    showIf: { id: "desea_modificar_proceso", equals: "Sí" }
  }
];
// Flappy game uses canvas — no FOODS array needed
const MOTIV = [
  "¡Eres increíble! Cada reto es una oportunidad disfrazada. 🚀",
  "¡El éxito es la suma de pequeños esfuerzos repetidos! 💪",
  "¡Tú tienes el poder de hacer la diferencia hoy! ⭐",
  "¡Los campeones no se rinden, se reinventan! 🏆",
  "¡Cada día es una nueva oportunidad de brillar! ✨"
];

// ════════════════════════════════════════════════════
// STATE
// ════════════════════════════════════════════════════
let msalApp = null;
let msUser  = {name:"",email:"",id:"",token:""};
let answers = {};      // id de pregunta -> respuesta (string | string[])
let historyIds = [];   // preguntas ya contestadas, para el botón de regresar
let currentId = null;
let answering = false;
let enviando = false;   // evita envíos duplicados (doble clic)
let gameScoreFinal = 0;
let dbCount = 0;

// ════════════════════════════════════════════════════
// UTILS
// ════════════════════════════════════════════════════
function shuffle(a){return [...a].sort(()=>Math.random()-.5);}
function nameFromEmail(e){return e.split("@")[0].replace(/[._]/g," ").replace(/\b\w/g,c=>c.toUpperCase());}

function showScreen(id){
  document.querySelectorAll(".screen").forEach(s=>{s.classList.remove("active");});
  const el=document.getElementById(id);
  if(el){el.classList.add("active");}
}

// ════════════════════════════════════════════════════
// PARROT SVG
// ════════════════════════════════════════════════════
function parrotSVG(mood="idle", size=180){
  // Mood-based offsets for expression variation via overlays
  const cfg={
    idle:     {sparkle:false, tears:false, overlay:""},
    malo1:    {sparkle:false, tears:true,  overlay:"😢"},
    malo2:    {sparkle:false, tears:false, overlay:"😟"},
    regular:  {sparkle:false, tears:false, overlay:"😐"},
    bueno:    {sparkle:false, tears:false, overlay:"😊"},
    excelente:{sparkle:true,  tears:false, overlay:"🌟"},
    done:     {sparkle:true,  tears:false, overlay:""},
    eat:      {sparkle:false, tears:false, overlay:"😋"},
    disgust:  {sparkle:false, tears:true,  overlay:"😖"},
  };
  const m=cfg[mood]||cfg.idle;
  const h=Math.round(size*1.08);
  return `<svg width="${size}" height="${h}" viewBox="0 0 ${size} ${h}" class="parrot-svg parrot-idle" style="overflow:visible;">
    ${m.sparkle?`<text x="0" y="20" font-size="16" style="user-select:none">⭐</text><text x="${size-20}" y="14" font-size="12" style="user-select:none">✨</text>`:""}
    <image href="assets/img/mayita-parrot.png"
           x="0" y="0" width="${size}" height="${h}"
           style="image-rendering:auto;"/>
    ${m.tears?`<text x="${size*0.3}" y="${h*0.7}" font-size="${size*0.18}" style="user-select:none">😢</text>`:""}
    ${m.overlay&&m.overlay!=""?`<text x="${size*0.68}" y="${h*0.28}" font-size="${size*0.2}" style="user-select:none">${m.overlay}</text>`:""}
  </svg>`;
}

function setParrotBounce(el){
  el.querySelector("svg").classList.remove("parrot-idle","parrot-bounce");
  void el.querySelector("svg").offsetWidth;
  el.querySelector("svg").classList.add("parrot-bounce");
  setTimeout(()=>{
    if(el.querySelector("svg")){
      el.querySelector("svg").classList.remove("parrot-bounce");
      el.querySelector("svg").classList.add("parrot-idle");
    }
  },600);
}

// ════════════════════════════════════════════════════
// PARTICLES
// ════════════════════════════════════════════════════
function boom(){
  const emojis=["🎉","⭐","✨","💫","🎊","🦜","🏆"];
  for(let i=0;i<16;i++){
    const p=document.createElement("div");
    p.className="particle";
    p.textContent=emojis[i%emojis.length];
    const dx=(Math.random()-.5)*340,dy=-(Math.random()*220+60);
    p.style.setProperty("--dx",dx+"px");
    p.style.setProperty("--dy",dy+"px");
    p.style.setProperty("--rot",(Math.random()-.5)*720+"deg");
    document.body.appendChild(p);
    setTimeout(()=>p.remove(),1700);
  }
}

// ════════════════════════════════════════════════════
// MSAL LOGIN
// ════════════════════════════════════════════════════
async function initMSAL(){
  if(msalApp) return;
  msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
  await msalApp.initialize();
}

async function handleRedirectResult(){
  // Called on every page load — catches the token after Microsoft redirects back
  try{
    const result = await msalApp.handleRedirectPromise();
    if(result && result.accessToken){
      await processLoginResult(result);
    }
  }catch(e){
    console.error("Redirect handle error:", e);
    showScreen("s-choose");
    const errEl=document.getElementById("choose-err");
    const errTxt=document.getElementById("choose-err-text");
    errTxt.textContent = e.message||"Error al procesar el inicio de sesión.";
    errEl.classList.add("show");
    resetChooseButtons();
  }
}

async function processLoginResult(result){
  const meRes = await fetch("https://graph.microsoft.com/v1.0/me",{
    headers:{Authorization:`Bearer ${result.accessToken}`}
  });
  const me = await meRes.json();
  const email = (me.mail||me.userPrincipalName||"").toLowerCase();
  const domain = email.split("@")[1]||"";
  if(!VALID_DOMAINS.includes(domain)){
    showScreen("s-choose");
    const errEl=document.getElementById("choose-err");
    const errTxt=document.getElementById("choose-err-text");
    errTxt.textContent = "Cuenta no pertenece a celaque.net o administracion.hn";
    errEl.classList.add("show");
    resetChooseButtons();
    return;
  }
  msUser = {
    name: me.displayName||nameFromEmail(email),
    email: email,
    id: me.id,
    token: result.accessToken
  };
  if(sessionStorage.getItem("mayita-intent")==="skip"){
    sessionStorage.removeItem("mayita-intent");
    goToQuiz();
  }else{
    goToGame();
  }
}

const CHOOSE_BUTTON_DEFAULTS={
  "btn-choose-flappy": `<span style="font-size:1.6rem">🦜</span><span>Flappy mayITa</span>`,
  "btn-choose-tower": `<span style="font-size:1.6rem">🏗️</span><span>Torre de Bloques</span>`,
  "btn-choose-skip": `<span style="font-size:1.6rem">📋</span><span>Saltar juego → ir a encuesta</span>`
};

function resetChooseButtons(){
  Object.keys(CHOOSE_BUTTON_DEFAULTS).forEach(id=>{
    const btn=document.getElementById(id);
    btn.disabled=false;
    btn.innerHTML=CHOOSE_BUTTON_DEFAULTS[id];
  });
}

const IS_LOCAL_DEV = ["localhost","127.0.0.1"].includes(window.location.hostname);

async function msLogin(btn, intent){
  const errEl=document.getElementById("choose-err");
  document.getElementById("btn-choose-flappy").disabled=true;
  document.getElementById("btn-choose-tower").disabled=true;
  document.getElementById("btn-choose-skip").disabled=true;
  errEl.classList.remove("show");

  if(IS_LOCAL_DEV){
    // En local no hay forma de que Entra ID redirija de vuelta a localhost,
    // así que simulamos un login exitoso y entramos directo al destino.
    msUser = { name:"Dev Local", email:"dev.local@celaque.net", id:"dev-local", token:"fake-local-token" };
    if(intent==="skip") goToQuiz(); else goToGame();
    return;
  }

  sessionStorage.setItem("mayita-intent", intent||"game");
  btn.innerHTML="⏳ Redirigiendo a Microsoft…";
  try{
    await initMSAL();
    // loginRedirect — redirige la misma página, evita bloqueo de popups anidados
    await msalApp.loginRedirect({scopes: SCOPES});
    // Execution stops here — Microsoft redirects back to this page
  }catch(e){
    const errTxt=document.getElementById("choose-err-text");
    errTxt.textContent = e.message||"No se pudo iniciar sesión. Intenta de nuevo.";
    errEl.classList.add("show");
    resetChooseButtons();
  }
}
document.getElementById("btn-start").addEventListener("click",()=>showScreen("s-choose"));
document.getElementById("btn-choose-flappy").addEventListener("click",function(){msLogin(this,"game");});
document.getElementById("btn-choose-tower").addEventListener("click",goToTower);
document.getElementById("btn-choose-skip").addEventListener("click",function(){msLogin(this,"skip");});

// ════════════════════════════════════════════════════
// FLAPPY MAYITA — Canvas Game
// ════════════════════════════════════════════════════
let flappyRAF = null;
let flappyKeyHandler = null;

const FW=436, FH=320, GRAVITY=0.2, JUMP_VEL=-5.2;
const PIPE_W=62, PIPE_GAP=132, PIPE_SPEED_BASE=1.1, PIPE_FREQ=95, SCORE_WIN=50;
// Buildings data for pipe skins (cycling through all 11)
const BLDGS=[
  {name:"Zorzales"},
  {name:"Aura"},
  {name:"Céfiro Azul"},
  {name:"D. Artemisa"},
  {name:"Agalta"},
  {name:"Atlas"},
];

let fb={};
function fbReset(){
  fb={phase:"ready",bird:{x:90,y:FH/2,vy:0,angle:0,flap:0},pipes:[],score:0,frame:0,best:fb.best||0,particles:[],bldgIdx:0,speed:PIPE_SPEED_BASE};
}
fbReset();

function drawMacaw(ctx,x,y,angle,size,flap){
  if(!window._mf1){ window._mf1=new Image(); window._mf1.src="assets/img/embedded-0c09f732db.png"; }
  if(!window._mf2){ window._mf2=new Image(); window._mf2.src="assets/img/embedded-ef3aa5ab1b.png"; }
  if(!window._mf3){ window._mf3=new Image(); window._mf3.src="assets/img/embedded-815a2263d8.png"; }
  // 3-frame cycle: up → mid → down → mid → up...  (6 frames each)
  const cycle = Math.floor(flap/6) % 4;
  let img, iw, ih;
  if(cycle===0)      { img=window._mf1; iw=260; ih=305; }
  else if(cycle===1) { img=window._mf2; iw=260; ih=236; }
  else if(cycle===2) { img=window._mf3; iw=260; ih=221; }
  else               { img=window._mf2; iw=260; ih=236; }
  const w=size, h=Math.round(size*(ih/iw));
  ctx.save();
  ctx.translate(x,y);
  ctx.rotate(angle);
  ctx.drawImage(img,-w/2,-h/2,w,h);
  ctx.restore();
}

function drawBuilding(ctx, x, gapTop, bldgData){
  if(!window._b0){ window._b0=new Image(); window._b0.src="assets/img/embedded-eee2d4f6ec.png"; window._b0.bw=130; window._b0.bh=112; }
  if(!window._b1){ window._b1=new Image(); window._b1.src="assets/img/embedded-77546f2069.png"; window._b1.bw=130; window._b1.bh=172; }
  if(!window._b2){ window._b2=new Image(); window._b2.src="assets/img/embedded-be7afe9fb4.png"; window._b2.bw=130; window._b2.bh=151; }
  if(!window._b3){ window._b3=new Image(); window._b3.src="assets/img/embedded-9a0589189e.png"; window._b3.bw=130; window._b3.bh=100; }
  if(!window._b4){ window._b4=new Image(); window._b4.src="assets/img/embedded-e0c035d13f.png"; window._b4.bw=130; window._b4.bh=161; }
  if(!window._b5){ window._b5=new Image(); window._b5.src="assets/img/embedded-bed22e7fec.png"; window._b5.bw=130; window._b5.bh=193; }

  const imgs  = [window._b0,window._b1,window._b2,window._b3,window._b4,window._b5];
  const names = ["Zorzales","Aura","Céfiro Azul","D. Artemisa","Agalta","Atlas"];
  const idx   = names.indexOf(bldgData.name);
  const img   = imgs[idx >= 0 ? idx : 0];
  const BW    = PIPE_W + 10;
  const gapBot = gapTop + PIPE_GAP;

  if(img && img.complete && img.naturalWidth > 0){

    // ── TOP building: flip vertically, stretch to fill exactly gapTop height ──
    const topH = gapTop;
    if(topH > 10){
      ctx.save();
      ctx.translate(x, 0);
      ctx.scale(1, -1);               // flip so roof faces up
      ctx.drawImage(img, 0, -topH, BW, topH); // stretch to fill — no tiling
      ctx.restore();
      // Name label
      ctx.save();
      ctx.globalAlpha=0.82;
      ctx.fillStyle="rgba(0,0,0,0.5)";
      ctx.beginPath(); ctx.roundRect(x+2,gapTop-20,BW-4,14,4); ctx.fill();
      ctx.fillStyle="#fff"; ctx.font="bold 7px Arial"; ctx.textAlign="center";
      ctx.fillText(bldgData.name, x+BW/2, gapTop-10);
      ctx.restore();
    }

    // ── BOTTOM building: right-side up, stretch to fill exactly botH height ──
    const botH = FH - 40 - gapBot;
    if(botH > 10){
      ctx.drawImage(img, x, gapBot, BW, botH); // stretch — no tiling
    }

  } else {
    // Fallback while loading
    ctx.fillStyle="#1a6b1a";
    ctx.fillRect(x, 0, BW, gapTop);
    ctx.fillRect(x, gapBot, BW, FH-40-gapBot);
  }

  // Gap hint
  ctx.save();
  ctx.globalAlpha=0.5;
  ctx.fillStyle="rgba(255,255,255,0.08)";
  ctx.beginPath(); ctx.roundRect(x+5,gapTop+PIPE_GAP/2-9,BW-10,18,5); ctx.fill();
  ctx.fillStyle="rgba(255,255,255,0.6)"; ctx.font="bold 7px Arial"; ctx.textAlign="center";
  ctx.fillText("✈ VUELA", x+BW/2, gapTop+PIPE_GAP/2+3);
  ctx.restore();
}

function drawBg(ctx,frame){
  // Sky — dawn/dusk gradient
  const t = (Math.sin(frame*0.002)+1)/2; // 0-1 cycle
  const sky=ctx.createLinearGradient(0,0,0,FH);
  sky.addColorStop(0,`hsl(${220+t*30},${60+t*20}%,${8+t*12}%)`);
  sky.addColorStop(0.5,`hsl(${200+t*40},${50+t*20}%,${12+t*15}%)`);
  sky.addColorStop(1,`hsl(${180+t*30},${40}%,${6+t*8}%)`);
  ctx.fillStyle=sky; ctx.fillRect(0,0,FW,FH);

  // Moon
  ctx.save();
  ctx.globalAlpha=0.6;
  ctx.fillStyle="#fffde0";
  ctx.beginPath(); ctx.arc(380,30,14,0,Math.PI*2); ctx.fill();
  ctx.globalAlpha=0.15; ctx.fillStyle="#fff";
  ctx.beginPath(); ctx.arc(380,30,20,0,Math.PI*2); ctx.fill();
  ctx.restore();

  // Stars
  const stars=[[20,18],[55,10],[110,25],[170,8],[240,20],[300,12],[355,22],[405,14],[430,28],[70,40],[200,35],[340,38]];
  stars.forEach(([sx,sy])=>{
    ctx.globalAlpha=0.3+0.5*Math.abs(Math.sin(frame*0.025+sx*0.12));
    ctx.fillStyle="#fff"; ctx.fillRect(sx,sy,1.5,1.5);
  });
  ctx.globalAlpha=1;

  // Background city silhouette (far buildings)
  ctx.fillStyle="rgba(15,10,30,0.7)";
  const bgBuildings=[[0,80,30,FH-40],[35,90,25,FH-40],[65,70,40,FH-40],[110,85,20,FH-40],
    [135,75,30,FH-40],[170,65,25,FH-40],[200,80,35,FH-40],[240,55,20,FH-40],
    [265,80,30,FH-40],[300,70,40,FH-40],[345,85,25,FH-40],[375,60,30,FH-40],
    [410,75,26,FH-40]];
  bgBuildings.forEach(([bx,by,bw,bh])=>{
    ctx.fillRect(bx,by,bw,bh);
    // Tiny lit windows in bg
    ctx.fillStyle="rgba(255,220,100,0.4)";
    for(let r=0;r<3;r++) for(let c=0;c<2;c++){
      if(Math.sin(bx*r+frame*0.01)>0.1) ctx.fillRect(bx+3+c*8,by+5+r*10,4,5);
    }
    ctx.fillStyle="rgba(15,10,30,0.7)";
  });

  // Ground
  const gnd=ctx.createLinearGradient(0,FH-40,0,FH);
  gnd.addColorStop(0,"#1a1025"); gnd.addColorStop(1,"#0d0818");
  ctx.fillStyle=gnd; ctx.fillRect(0,FH-40,FW,40);
  // Road
  ctx.fillStyle="#1e1530"; ctx.fillRect(0,FH-30,FW,20);
  // Road markings
  ctx.fillStyle="rgba(255,255,150,0.25)";
  const roadOff=(frame*PIPE_SPEED_BASE*1.2)%50;
  for(let rx=FW-roadOff;rx>-50;rx-=50) ctx.fillRect(rx,FH-22,30,3);
  // Sidewalk
  ctx.fillStyle="#151020"; ctx.fillRect(0,FH-40,FW,10);
  // Street lights
  for(let lx=40-(frame*PIPE_SPEED_BASE)%80;lx<FW+40;lx+=80){
    ctx.strokeStyle="rgba(255,220,100,0.5)"; ctx.lineWidth=1.5;
    ctx.beginPath(); ctx.moveTo(lx,FH-40); ctx.lineTo(lx,FH-60); ctx.lineTo(lx+8,FH-62); ctx.stroke();
    ctx.fillStyle="rgba(255,230,100,0.6)";
    ctx.beginPath(); ctx.arc(lx+8,FH-62,2.5,0,Math.PI*2); ctx.fill();
    // Light cone
    ctx.globalAlpha=0.04; ctx.fillStyle="#ffe064";
    ctx.beginPath(); ctx.moveTo(lx+8,FH-62); ctx.lineTo(lx-10,FH-40); ctx.lineTo(lx+26,FH-40); ctx.closePath(); ctx.fill();
    ctx.globalAlpha=1;
  }
}

function flappyLoop(canvas){
  const ctx=canvas.getContext("2d");
  const scoreEl=document.getElementById("flappy-score");
  function loop(){
    if(!document.getElementById("flappy-canvas")){cancelAnimationFrame(flappyRAF);return;}
    flappyRAF=requestAnimationFrame(loop);
    fb.frame++;
    const b=fb.bird;
    if(fb.phase==="playing"){
      // Speed increases every 5 buildings passed
      fb.speed = PIPE_SPEED_BASE + Math.floor(fb.score/5)*0.15;
      b.vy=Math.min(b.vy+GRAVITY,10); b.y+=b.vy;
      b.angle=Math.max(-0.5,Math.min(1.2,b.vy*0.08)); b.flap++;
      if(fb.frame%PIPE_FREQ===0){
        const bldg = BLDGS[fb.bldgIdx % BLDGS.length];
        fb.pipes.push({x:FW+10,gapTop:50+Math.random()*(FH-PIPE_GAP-110),passed:false,bldg});
        fb.bldgIdx++;
      }
      fb.pipes.forEach(p=>{
        p.x-=fb.speed;
        if(!p.passed&&p.x+PIPE_W<b.x){
          p.passed=true; fb.score++;
          if(scoreEl) scoreEl.textContent=fb.score;
          for(let i=0;i<8;i++) fb.particles.push({x:b.x,y:b.y,vx:(Math.random()-.5)*5,vy:-2.5-Math.random()*2.5,life:50,e:["⭐","✨","💫","🏙️"][i%4]});
        }
      });
      fb.pipes=fb.pipes.filter(p=>p.x>-PIPE_W-20);
      // Collisions
      if(b.y+14>FH-40||b.y-14<0){killBird();return;}
      for(const p of fb.pipes){
        if(b.x+11>p.x+3&&b.x-11<p.x+PIPE_W-3){
          if(b.y-11<p.gapTop||b.y+11>p.gapTop+PIPE_GAP){killBird();return;}
        }
      }
      // Win!
      if(fb.score>=SCORE_WIN){
        fb.phase="done"; fb.best=Math.max(fb.best,fb.score);
        gameScoreFinal=fb.score;
        saveRanking(fb.score);
        setTimeout(()=>showFlappyResult(true),400);
      }
    }
    fb.particles=fb.particles.filter(p=>p.life>0);
    fb.particles.forEach(p=>{p.x+=p.vx;p.y+=p.vy;p.vy+=0.2;p.life--;});
    // Draw
    drawBg(ctx,fb.frame);
    fb.pipes.forEach(p=>drawBuilding(ctx,p.x,p.gapTop,p.bldg));
    fb.particles.forEach(p=>{ctx.globalAlpha=p.life/50;ctx.font="13px serif";ctx.fillText(p.e,p.x-7,p.y+7);});
    ctx.globalAlpha=1;
    if(fb.phase!=="done") drawMacaw(ctx,b.x,fb.phase==="ready"?b.y+Math.sin(fb.frame*.08)*5:b.y,fb.phase==="ready"?0:b.angle,34,b.flap);
    // HUD score
    if(fb.phase==="playing"){
      ctx.fillStyle="rgba(0,0,0,.55)"; ctx.beginPath(); ctx.roundRect(FW/2-38,8,76,34,10); ctx.fill();
      ctx.fillStyle="#ffd700"; ctx.font="bold 22px Nunito,Arial"; ctx.textAlign="center"; ctx.fillText(fb.score,FW/2,30);
      const spd=Math.round((fb.speed-PIPE_SPEED_BASE)/0.15)+1;
      if(spd>1){ ctx.fillStyle="rgba(255,140,0,0.9)"; ctx.font="bold 9px Arial"; ctx.fillText("VEL "+spd+"×",FW/2,48); }
      ctx.textAlign="left";
      const nextP=fb.pipes.find(p=>!p.passed&&p.x>b.x);
      if(nextP){
        ctx.fillStyle="rgba(0,0,0,0.45)"; ctx.beginPath(); ctx.roundRect(4,6,72,16,4); ctx.fill();
        ctx.fillStyle=nextP.bldg.accent; ctx.font="bold 8px Arial";
        ctx.fillText("▶ "+nextP.bldg.name,8,17);
      }
    }
    if(fb.phase==="ready"){
      ctx.fillStyle="rgba(0,0,0,.55)"; ctx.beginPath(); ctx.roundRect(FW/2-115,FH/2-46,230,92,14); ctx.fill();
      ctx.fillStyle="#ffd700"; ctx.font="bold 16px Nunito,Arial"; ctx.textAlign="center"; ctx.fillText("¡Flappy mayITa! 🦜",FW/2,FH/2-24);
      ctx.fillStyle="rgba(255,255,255,.75)"; ctx.font="12px Arial"; ctx.fillText("Vuela entre los edificios de Celaque",FW/2,FH/2-4);
      ctx.fillStyle="#ff8c42"; ctx.font="bold 12px Arial"; ctx.fillText("Tap  ·  Click  ·  Espacio para volar",FW/2,FH/2+18);
      ctx.fillStyle="rgba(255,255,255,.4)"; ctx.font="10px Arial"; ctx.fillText("¡Sin límite de edificios! Llega lo más lejos posible",FW/2,FH/2+36);
      ctx.textAlign="left";
    }
    if(fb.phase==="dead"){
      // Show result overlay in canvas while waiting for user click
      ctx.fillStyle="rgba(0,0,0,.7)"; ctx.beginPath(); ctx.roundRect(FW/2-115,FH/2-60,230,120,16); ctx.fill();
      ctx.fillStyle="#f87171"; ctx.font="bold 22px Nunito,Arial"; ctx.textAlign="center"; ctx.fillText("¡Chocaste! 💥",FW/2,FH/2-32);
      ctx.fillStyle="#ffd700"; ctx.font="bold 28px Nunito,Arial"; ctx.fillText(fb.score+" edificios",FW/2,FH/2+2);
      if(fb.score>fb.best-1&&fb.score>0){ ctx.fillStyle="#22c55e"; ctx.font="bold 10px Arial"; ctx.fillText("🏆 ¡NUEVO RÉCORD!",FW/2,FH/2+20); }
      else{ ctx.fillStyle="rgba(255,255,255,.5)"; ctx.font="10px Arial"; ctx.fillText("Récord: "+fb.best,FW/2,FH/2+20); }
      ctx.fillStyle="rgba(255,255,255,.6)"; ctx.font="11px Arial"; ctx.fillText("Toca para intentar de nuevo",FW/2,FH/2+42);
      ctx.textAlign="left";
    }
  }
  loop();
}

// ─── Ranking (localStorage) ──────────────────────────
// ─── Ranking general SharePoint / n8n ──────────────────────────
async function saveRanking(score) {
  try {
    const response = await fetch(N8N_RANKING_SAVE_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        usuario: {
          nombre: msUser.name,
          correo: msUser.email,
          entraId: msUser.id
        },
        puntaje: Number(score),
        fechaRegistro: new Date().toISOString(),
        origen: "GitHub Pages"
      })
    });

    const data = await response.json().catch(() => null);

    if (!response.ok || !data?.ok) {
      throw new Error(
        data?.mensaje ||
        `No se pudo guardar el puntaje: ${response.status}`
      );
    }

    return true;

  } catch (error) {
    console.error("Error guardando ranking:", error);
    return false;
  }
}

async function getRanking() {
  try {
    const response = await fetch(N8N_RANKING_TOP_URL, {
      method: "GET",
      headers: {
        "Accept": "application/json"
      },
      cache: "no-store"
    });

    const data = await response.json().catch(() => null);

    if (!response.ok || !data?.ok) {
      throw new Error(
        data?.mensaje ||
        `No se pudo consultar el ranking: ${response.status}`
      );
    }

    return Array.isArray(data.ranking)
      ? data.ranking
      : [];

  } catch (error) {
    console.error("Error consultando ranking:", error);
    return [];
  }
}

function formatearFechaRanking(fecha) {
  if (!fecha) return "";

  const date = new Date(fecha);

  if (Number.isNaN(date.getTime())) {
    return fecha;
  }

  return date.toLocaleDateString("es-HN");
}

async function rankingHTML() {
  const list = await getRanking();

  if (!list.length) {
    return `
      <div style="
        color:rgba(255,255,255,.4);
        font-size:.8rem;
        text-align:center;
        padding:8px;
      ">
        Sé el primero en el ranking 🦜
      </div>
    `;
  }

  const medals = ["🥇", "🥈", "🥉"];

  return list.map((r, i) => `
    <div style="
      display:flex;
      align-items:center;
      gap:8px;
      padding:5px 8px;
      border-radius:9px;
      background:${
        i === 0
          ? "rgba(255,215,0,.1)"
          : i === 1
          ? "rgba(192,192,192,.07)"
          : i === 2
          ? "rgba(205,127,50,.07)"
          : "rgba(255,255,255,.03)"
      };
      margin-bottom:3px;
    ">
      <span style="
        font-size:.9rem;
        width:20px;
        text-align:center;
      ">
        ${medals[i] || "#" + (i + 1)}
      </span>

      <span style="
        color:#fff;
        font-weight:800;
        font-size:.82rem;
        flex:1;
        overflow:hidden;
        text-overflow:ellipsis;
        white-space:nowrap;
      ">
        ${r.name || "Usuario"}
      </span>

      <span style="
        color:#ffd700;
        font-weight:900;
        font-size:.88rem;
      ">
        ${Number(r.score || 0)}
      </span>

      <span style="
        color:rgba(255,255,255,.28);
        font-size:.7rem;
      ">
        ${formatearFechaRanking(r.date)}
      </span>
    </div>
  `).join("");
}

async function killBird() {
  fb.bird.vy = JUMP_VEL * 0.4;
  fb.phase = "dead";
  fb.best = Math.max(fb.best, fb.score);

  await saveRanking(fb.score);

  setTimeout(() => {
    if (fb.phase === "dead") {
      showFlappyResult(false);
    }
  }, 1800);
}

function flappyJump(){
  if(fb.phase==="ready"){fb.phase="playing";fb.bird.vy=JUMP_VEL;return;}
  if(fb.phase==="playing"){fb.bird.vy=JUMP_VEL;return;}
  if(fb.phase==="dead"){fbReset();fb.phase="playing";fb.bird.vy=JUMP_VEL;}
}

async function showFlappyResult(won){
detenerControlesFlappy();
  gameScoreFinal=fb.score;
  const score=fb.score, best=fb.best;
  const isRecord=score>0&&score>=best;
  const mood=won?"excelente":score>=8?"excelente":score>=4?"bueno":"regular";
  const title=won?"¡Lo lograste! 🏆🎉":score>=8?"¡Increíble vuelo! 🏆":score>=4?"¡Buen vuelo! 🎉":"¡Buen intento! 😄";
  const rankingGeneral = await rankingHTML();
  document.getElementById("game-inner").innerHTML=`
    <div style="text-align:center">
      <div style="display:flex;justify-content:center;margin-bottom:6px">${parrotSVG(mood,96)}</div>
      <div style="font-family:'Syne',sans-serif;font-size:1.1rem;font-weight:800;color:#fff;margin-bottom:8px">${title}</div>
      <div style="background:rgba(255,255,255,.07);border-radius:14px;padding:10px 12px;margin-bottom:8px;display:flex;justify-content:space-around;align-items:center;">
        <div>
          <div style="color:rgba(255,255,255,.42);font-size:.65rem;font-weight:800;letter-spacing:1px;text-transform:uppercase;margin-bottom:3px">🏙️ Edificios</div>
          <div style="font-size:1.9rem;font-weight:900;color:#ffd700;line-height:1">${score}</div>
        </div>
        <div style="width:1px;height:34px;background:rgba(255,255,255,.15)"></div>
        <div>
          <div style="color:rgba(255,255,255,.42);font-size:.65rem;font-weight:800;letter-spacing:1px;text-transform:uppercase;margin-bottom:3px">🏅 Mi récord</div>
          <div style="font-size:1.9rem;font-weight:900;color:${isRecord?"#22c55e":"rgba(255,255,255,.5)"};line-height:1">${best}</div>
        </div>
        <div style="width:1px;height:34px;background:rgba(255,255,255,.15)"></div>
        <div>
          <div style="color:rgba(255,255,255,.42);font-size:.65rem;font-weight:800;letter-spacing:1px;text-transform:uppercase;margin-bottom:3px">🎯 Meta</div>
          <div style="font-size:1.9rem;font-weight:900;color:rgba(255,255,255,.45);line-height:1">${SCORE_WIN}</div>
        </div>
      </div>
      ${isRecord?`<div style="background:rgba(34,197,94,.12);border:1px solid rgba(34,197,94,.3);border-radius:10px;padding:6px;margin-bottom:6px;color:#86efac;font-size:.8rem;font-weight:800;">🏆 ¡Nuevo récord personal!</div>`:""}
      ${won||score>=4?`<div class="motivational" style="margin-bottom:8px;padding:10px 12px;">💫 ${MOTIV[Math.floor(Math.random()*MOTIV.length)]}</div>`:""}
      <div style="background:rgba(255,255,255,.05);border-radius:14px;padding:10px 10px;margin-bottom:10px;text-align:left;">
        <div style="color:rgba(255,255,255,.45);font-size:.65rem;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:7px;text-align:center;">🏆 Ranking mayITa</div>
       ${rankingGeneral}
      </div>
      <button class="btn-primary" style="margin-top:0" onclick="goToQuiz()">Ir a la encuesta 📋</button>
      <button onclick="fbReset();renderFlappyArena();" style="width:100%;padding:12px;border-radius:14px;border:1.5px solid rgba(255,255,255,.2);background:rgba(255,255,255,.06);color:rgba(255,255,255,.72);font-family:'Nunito',sans-serif;font-weight:800;font-size:.88rem;cursor:pointer;margin-top:9px;" onmouseover="this.style.background='rgba(255,255,255,.12)'" onmouseout="this.style.background='rgba(255,255,255,.06)'">
        Intentar de nuevo 🔄
      </button>
    </div>`;
}

function renderFlappyArena(){
detenerControlesFlappy();
  fbReset();
  document.getElementById("game-inner").innerHTML=`
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
      <div style="color:rgba(255,255,255,.5);font-size:.74rem;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;">Flappy mayITa 🦜</div>
      <div style="background:rgba(255,255,255,.08);border-radius:20px;padding:4px 14px;color:#ffd700;font-weight:900;font-size:.88rem;">🏙️ <span id="flappy-score">0</span></div>
    </div>
    <div style="display:flex;justify-content:center;">
      <canvas id="flappy-canvas" width="${FW}" height="${FH}" style="border-radius:18px;cursor:pointer;touch-action:none;max-width:100%;"></canvas>
    </div>
    <div style="text-align:center;margin-top:8px;color:rgba(255,255,255,.38);font-size:.76rem;font-weight:800;">Tap · Click · Espacio para volar · ¡Sin límite!</div>`;
  const canvas=document.getElementById("flappy-canvas");
  canvas.addEventListener("click",flappyJump);
  canvas.addEventListener("touchstart",e=>{e.preventDefault();flappyJump();},{passive:false});
 if (flappyKeyHandler) {
  document.removeEventListener("keydown", flappyKeyHandler);
}

flappyKeyHandler = (e) => {
  const tag = document.activeElement?.tagName?.toLowerCase();

  const escribiendo =
    tag === "input" ||
    tag === "textarea" ||
    document.activeElement?.isContentEditable;

  if (escribiendo) {
    return;
  }

  if (e.code === "Space" || e.code === "ArrowUp") {
    e.preventDefault();
    flappyJump();
  }
};

document.addEventListener("keydown", flappyKeyHandler);
  flappyLoop(canvas);
}
function detenerControlesFlappy() {
  cancelAnimationFrame(flappyRAF);

  if (flappyKeyHandler) {
    document.removeEventListener(
      "keydown",
      flappyKeyHandler
    );

    flappyKeyHandler = null;
  }
}

function goToGame(){
  document.getElementById("game-avatar-initial").textContent=msUser.name[0].toUpperCase();
  document.getElementById("game-user-name").textContent=msUser.name;
  document.getElementById("game-user-email").textContent=msUser.email;
  showScreen("s-game");
  renderGameReady();
}

function goToTower(){
  showScreen("s-tower");
  if(typeof initTowerBlocks==="function") initTowerBlocks();
}

document.getElementById("tower-back-button").addEventListener("click",function(){
  if(typeof towerGame!=="undefined" && towerGame) towerGame.resetToReady();
  if(typeof leaveTowerBlocks==="function") leaveTowerBlocks();
  showScreen("s-choose");
});

function renderGameReady(){
  detenerControlesFlappy();
  fbReset();
  document.getElementById("game-inner").innerHTML=`
    <div style="text-align:center;max-width:380px;margin:0 auto">
      <div id="game-parrot-ready" style="display:flex;justify-content:center;margin-bottom:18px"></div>
      <div style="color:#fff;font-family:'Syne',sans-serif;font-weight:800;font-size:1.15rem;margin-bottom:14px">¡Flappy mayITa! 🦜</div>
      <p style="color:rgba(255,255,255,.62);font-size:.87rem;line-height:1.8;margin-bottom:26px">
        Vuela como mayita entre los proyectos.<br>
        <strong style="color:#fbbf24">Tap  · Click · Espacio</strong> para volar.<br>
        ¡Pasa <strong style="color:#22c55e">todos los edificios</strong> que puedas!
      </p>
      <button class="btn-primary" style="margin-top:6px" onclick="renderFlappyArena()">¡Jugar! 🎮</button>
    </div>`;
  document.getElementById("game-parrot-ready").innerHTML=parrotSVG("idle",130);
}

// ════════════════════════════════════════════════════
// ENCUESTA — RETROALIMENTACIÓN DE PROCESOS
// ════════════════════════════════════════════════════

const $ = (id) => document.getElementById(id);

function esVisible(q) {
  if (!q.showIf) return true;
  return answers[q.showIf.id] === q.showIf.equals;
}

function preguntasVisibles() {
  return SURVEY_QUESTIONS.filter(esVisible);
}

function getPregunta(id) {
  return SURVEY_QUESTIONS.find((q) => q.id === id);
}

// Valor final de una respuesta tal como se guarda
function valorFinal(q) {
  const v = answers[q.id];
  if (q.type === "multi") return Array.isArray(v) ? v : [];
  if (q.otroConTexto && v === q.otroConTexto) {
    const extra = String(answers[q.id + "__otro"] || "").trim();
    return extra ? `${v}: ${extra}` : v;
  }
  return typeof v === "string" ? v.trim() : "";
}

function goToQuiz() {
  answers = {};
  historyIds = [];
  currentId = null;
  answering = false;
  enviando = false;
  showScreen("s-quiz");
  $("quiz-parrot").innerHTML = parrotSVG("idle", 180);
  mostrarPregunta(preguntasVisibles()[0].id);
}

function goBackQuiz() {
  if (answering) return;
  if (historyIds.length === 0) {
    showScreen("s-choose");
    return;
  }
  mostrarPregunta(historyIds.pop());
}

function renderProgreso() {
  const visibles = preguntasVisibles();
  const idx = Math.max(0, visibles.findIndex((q) => q.id === currentId));
  $("quiz-progress-label").textContent = `${SURVEY_NAME} · Pregunta ${idx + 1}`;
  const pct = Math.round((idx / visibles.length) * 100);
  $("quiz-dots").innerHTML = `
    <div style="width:96px;height:8px;border-radius:5px;background:rgba(255,255,255,.18);overflow:hidden;">
      <div style="height:100%;width:${pct}%;background:#ff8c42;border-radius:5px;transition:width .3s ease;"></div>
    </div>`;
}

function mostrarPregunta(id) {
  const q = getPregunta(id);
  if (!q) {
    finalizarEncuesta();
    return;
  }

  currentId = id;
  answering = false;
  $("quiz-back-btn").disabled = false;
  $("quiz-question").textContent = q.text;
  $("quiz-parrot").innerHTML = parrotSVG("idle", 180);
  renderProgreso();

  const cont = $("quiz-opts");
  cont.innerHTML = "";

  if (q.type === "yesno") {
    renderUnica(q, ["Sí", "No"], true);
  } else if (q.type === "single") {
    renderUnica(q, q.options, false);
  } else if (q.type === "multi") {
    renderMulti(q);
  } else {
    renderTexto(q);
  }
}

function avanzar() {
  if (enviando) return;
  const visibles = preguntasVisibles();
  const idx = visibles.findIndex((q) => q.id === currentId);
  const siguiente = visibles[idx + 1];
  historyIds.push(currentId);
  if (siguiente) {
    mostrarPregunta(siguiente.id);
  } else {
    finalizarEncuesta();
  }
}

// ── Opción única / Sí-No ────────────────────────────
function renderUnica(q, opciones, dosColumnas) {
  const cont = $("quiz-opts");
  cont.style.display = "grid";
  cont.style.gridTemplateColumns = dosColumnas ? "repeat(2, minmax(0, 1fr))" : "1fr";
  cont.style.gap = "10px";

  opciones.forEach((op) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "imp-btn";
    b.style.minHeight = "56px";
    b.style.justifyContent = dosColumnas ? "center" : "flex-start";
    b.style.textAlign = dosColumnas ? "center" : "left";
    b.style.fontWeight = "800";
    b.textContent = op;
    if (answers[q.id] === op) b.classList.add("sel");
    b.addEventListener("click", () => elegirUnica(q, op, b));
    cont.appendChild(b);
  });

  const extra = document.createElement("div");
  extra.id = "otro-panel";
  extra.style.gridColumn = "1 / -1";
  cont.appendChild(extra);

  if (q.otroConTexto && answers[q.id] === q.otroConTexto) {
    mostrarOtro(q);
  }
}

function elegirUnica(q, op, boton) {
  if (answering) return;
  answers[q.id] = op;

  $("quiz-opts")
    .querySelectorAll(".imp-btn")
    .forEach((b) => b.classList.remove("sel"));
  boton.classList.add("sel");

  if (q.otroConTexto && op === q.otroConTexto) {
    mostrarOtro(q);
    return;
  }

  delete answers[q.id + "__otro"];
  answering = true;
  $("quiz-back-btn").disabled = true;
  setTimeout(avanzar, 320);
}

function mostrarOtro(q) {
  const panel = $("otro-panel");
  panel.innerHTML = `
    <input id="otro-input" class="input" maxlength="300"
      placeholder="Especifica aquí..." style="margin-top:4px;" />
    <div id="otro-error" class="err-msg">⚠ Especifica el motivo para continuar.</div>
    <button id="btn-otro" type="button" class="btn-primary" style="margin-top:12px;">
      Continuar →
    </button>`;

  const input = $("otro-input");
  input.value = answers[q.id + "__otro"] || "";
  input.addEventListener("input", () => {
    if (input.value.trim()) $("otro-error").classList.remove("show");
  });
  $("btn-otro").addEventListener("click", () => {
    const texto = input.value.trim();
    if (!texto) {
      $("otro-error").classList.add("show");
      input.focus();
      return;
    }
    answers[q.id + "__otro"] = texto;
    avanzar();
  });
  setTimeout(() => input.focus(), 100);
}

// ── Selección múltiple ──────────────────────────────
function renderMulti(q) {
  const cont = $("quiz-opts");
  cont.style.display = "grid";
  cont.style.gridTemplateColumns = "1fr";
  cont.style.gap = "10px";

  const seleccion = new Set(Array.isArray(answers[q.id]) ? answers[q.id] : []);

  q.options.forEach((op) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "imp-btn";
    b.style.minHeight = "56px";
    b.innerHTML = `<div class="imp-check"></div><span></span>`;
    b.querySelector("span").textContent = op;

    const marcar = () => {
      const on = seleccion.has(op);
      b.classList.toggle("sel", on);
      b.querySelector(".imp-check").textContent = on ? "✓" : "";
    };
    marcar();

    b.addEventListener("click", () => {
      if (seleccion.has(op)) seleccion.delete(op);
      else seleccion.add(op);
      marcar();
      if (seleccion.size) $("multi-error").classList.remove("show");
    });
    cont.appendChild(b);
  });

  const pie = document.createElement("div");
  pie.style.gridColumn = "1 / -1";
  pie.innerHTML = `
    <div id="multi-error" class="err-msg">⚠ Selecciona al menos una opción.</div>
    <button id="btn-multi" type="button" class="btn-primary" style="margin-top:12px;">
      Continuar →
    </button>`;
  cont.appendChild(pie);

  $("btn-multi").addEventListener("click", () => {
    if (!seleccion.size) {
      $("multi-error").classList.add("show");
      return;
    }
    // Conserva el orden original de las opciones
    answers[q.id] = q.options.filter((o) => seleccion.has(o));
    avanzar();
  });
}

// ── Respuesta abierta ───────────────────────────────
function renderTexto(q) {
  const cont = $("quiz-opts");
  cont.style.display = "block";
  cont.style.gridTemplateColumns = "";
  cont.style.gap = "";

  cont.innerHTML = `
    <textarea id="texto-input" class="input" maxlength="1500"
      style="width:100%;min-height:${q.largo ? 130 : 84}px;margin-top:4px;resize:vertical;"></textarea>
    <div id="texto-error" class="err-msg" style="margin-top:7px;">
      ⚠ Escribe una respuesta para continuar.
    </div>
    <button id="btn-texto" type="button" class="btn-primary" style="margin-top:12px;">
      Continuar →
    </button>`;

  const input = $("texto-input");
  input.placeholder = q.placeholder || "Escribe tu respuesta aquí...";
  input.value = typeof answers[q.id] === "string" ? answers[q.id] : "";

  input.addEventListener("input", () => {
    if (input.value.trim()) $("texto-error").classList.remove("show");
  });
  $("btn-texto").addEventListener("click", () => {
    const valor = input.value.trim();
    if (!valor) {
      $("texto-error").classList.add("show");
      input.focus();
      return;
    }
    answers[q.id] = valor;
    avanzar();
  });
  setTimeout(() => input.focus(), 100);
}

// ════════════════════════════════════════════════════
// ENVÍO
// ════════════════════════════════════════════════════
function finalizarEncuesta() {
  if (enviando) return;
  enviando = true;
  answering = true;
  $("quiz-back-btn").disabled = true;
  $("quiz-opts").querySelectorAll("button").forEach((b) => (b.disabled = true));
  $("quiz-parrot").innerHTML = parrotSVG("done", 180);
  setTimeout(enviarEncuesta, 450);
}

function construirPayload() {
  const visibles = preguntasVisibles();

  const respuestas = visibles.map((q, i) => ({
    numeroPregunta: i + 1,
    id: q.id,
    pregunta: q.text,
    tipoRespuesta: q.type,
    respuesta: valorFinal(q)
  }));

  // Un campo por pregunta (vacío si no aplicó): útil para mapear a columnas
  // de SharePoint o a una sola fila por encuesta.
  const respuestasPorId = {};
  SURVEY_QUESTIONS.forEach((q) => {
    if (!esVisible(q)) {
      respuestasPorId[q.id] = "";
      return;
    }
    const v = valorFinal(q);
    respuestasPorId[q.id] = Array.isArray(v) ? v.join("; ") : v;
  });

  return {
    usuario: {
      nombre: msUser.name,
      correo: msUser.email,
      entraId: msUser.id
    },
    encuesta: SURVEY_NAME,
    respuestas,
    respuestasPorId,
    puntajeJuego: Number(gameScoreFinal || 0),
    origen: "mayITa Retroalimentación",
    versionApp: APP_VERSION
  };
}

async function guardarEncuestaEnN8n(payload) {
  // Mientras no exista el webhook real, en local se simula el guardado
  // para poder probar toda la interfaz.
  if (N8N_WEBHOOK_URL.startsWith("PENDIENTE")) {
    if (IS_LOCAL_DEV) {
      console.info("[DEV] Payload que se enviaría a n8n:", payload);
      await new Promise((r) => setTimeout(r, 700));
      return {
        ok: true,
        codigoEncuesta: "DEV-0001",
        respuestasGuardadas: payload.respuestas.length
      };
    }
    throw new Error("Falta configurar N8N_WEBHOOK_URL en assets/js/app.js.");
  }

  const response = await fetch(N8N_WEBHOOK_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  const texto = await response.text();
  let data = null;
  try {
    data = texto ? JSON.parse(texto) : null;
  } catch (error) {
    console.error("La respuesta de n8n no es JSON válido:", texto);
  }

  if (!response.ok) {
    throw new Error(data?.mensaje || `n8n respondió con estado ${response.status}`);
  }
  if (!data?.ok) {
    throw new Error(data?.mensaje || "n8n no confirmó el guardado.");
  }
  return data;
}

function estadoEnvio(estado, mensaje) {
  const cargando = estado === "enviando";
  $("sending-icon").textContent = cargando ? "⏳" : "⚠️";
  $("sending-title").textContent = cargando ? "Guardando encuesta…" : "No pudimos guardar la encuesta";
  $("sending-sub").textContent = cargando
    ? "mayITa está registrando tus respuestas de forma segura."
    : mensaje || "Intenta nuevamente.";
  $("send-actions").style.display = cargando ? "none" : "flex";
}

async function enviarEncuesta() {
  enviando = true;
  showScreen("s-sending");
  estadoEnvio("enviando");

  const payload = construirPayload();
  console.log("Payload mayITa Retroalimentación:", payload);

  try {
    const resultado = await guardarEncuestaEnN8n(payload);
    dbCount = Number(resultado.respuestasGuardadas || payload.respuestas.length);

    $("done-title").textContent = `¡Gracias, ${msUser.name}! 🎉`;
    $("done-parrot").innerHTML = parrotSVG("done", 160);

    const badge = $("teams-status-badge");
    badge.className = "teams-status ok";
    badge.textContent = "✓ Guardada";

    $("teams-msg-text").textContent =
      "Tu retroalimentación fue registrada correctamente.";

    const codigo = resultado.codigoEncuesta ? `Encuesta ${resultado.codigoEncuesta} · ` : "";
    $("db-sub-text").textContent =
      `${codigo}${dbCount} respuestas guardadas · ${new Date().toLocaleString("es-HN")}`;

    showScreen("s-done");
  } catch (error) {
    console.error("Error al guardar la encuesta:", error);
    enviando = false;
    estadoEnvio("error", `${error.message || "Intenta nuevamente."}`);
  }
}

// Botones de la pantalla de envío (reintentar / volver a revisar)
$("btn-send-retry").addEventListener("click", enviarEncuesta);
$("btn-send-back").addEventListener("click", () => {
  enviando = false;
  answering = false;
  showScreen("s-quiz");
  mostrarPregunta(historyIds.pop() ?? preguntasVisibles()[0].id);
});

// ════════════════════════════════════════════════════
// RESET
// ════════════════════════════════════════════════════
function resetAll(){

  msUser = {
    name:"",
    email:"",
    id:"",
    token:""
  };

  answers = {};
  historyIds = [];
  currentId = null;
  answering = false;
  enviando = false;
  gameScoreFinal = 0;


  detenerControlesFlappy();

  fbReset();


  if (
    typeof leaveTowerBlocks ===
    "function"
  ) {
    leaveTowerBlocks();
  }


  document
    .getElementById("intro-err")
    .classList
    .remove("show");


  document
    .getElementById("choose-err")
    .classList
    .remove("show");


  resetChooseButtons();

  showScreen("s-intro");
}

// ════════════════════════════════════════════════════
// INIT
// ════════════════════════════════════════════════════
document.getElementById("intro-parrot").innerHTML=parrotSVG("idle",180);

// Init MSAL and handle redirect result on every page load
(async()=>{
  try{
    await initMSAL();
    await handleRedirectResult();
  }catch(e){ console.error("MSAL init error:", e); }
})();
