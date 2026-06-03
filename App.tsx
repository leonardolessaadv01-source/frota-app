import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Storage keys ─────────────────────────────────────────────────────────────
const KEYS = {
  vehicles:  "frota_v6_vehicles",
  manut:     "frota_v6_manut",
  clientes:  "frota_v6_clientes",
  contratos: "frota_v6_contratos",
  docs:      "frota_v6_docs",
  fotos:     "frota_v6_fotos",
  reservas:  "frota_v6_reservas",
};

const FLEET_INIT = [
  { id:1,  placa:"EYA8E54", modelo:"VW Voyage",    ano:"2023", cor:"", status:"alugado",    km:"123000", kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"750", mensal:"", caucao:"" },
  { id:2,  placa:"GKD7B27", modelo:"VW Voyage",    ano:"2023", cor:"", status:"alugado",    km:"120000", kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"750", mensal:"", caucao:"" },
  { id:3,  placa:"QXC2D83", modelo:"Fiat Mobi",    ano:"2021", cor:"", status:"alugado",    km:"280000", kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"650", mensal:"", caucao:"" },
  { id:4,  placa:"SIV1D24", modelo:"Fiat Mobi",    ano:"2024", cor:"", status:"alugado",    km:"48000",  kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"650", mensal:"", caucao:"" },
  { id:5,  placa:"SWL1E93", modelo:"VW Polo Track",ano:"",     cor:"", status:"disponivel", km:"20000",  kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"850", mensal:"", caucao:"" },
  { id:6,  placa:"TIY9F98", modelo:"VW Polo Track",ano:"",     cor:"", status:"disponivel", km:"20000",  kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"850", mensal:"", caucao:"" },
  { id:7,  placa:"SUQ3J76", modelo:"VW Polo Track",ano:"",     cor:"", status:"disponivel", km:"20000",  kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"850", mensal:"", caucao:"" },
  { id:8,  placa:"SJB3B02", modelo:"Fiat Argo",    ano:"",     cor:"", status:"disponivel", km:"130000", kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"830", mensal:"", caucao:"" },
  { id:9,  placa:"E000B06", modelo:"VW Voyage",    ano:"",     cor:"", status:"disponivel", km:"120000", kmRevisao:"", proxRevisao:"", ipva:"", seguro:"", locatario:"", obs:"", diaria:"", semanal:"750", mensal:"", caucao:"" },
];

const STATUS_MAP = {
  disponivel:{ label:"Disponível",  color:"#22c55e", bg:"rgba(34,197,94,.12)",  dot:"#16a34a" },
  alugado:   { label:"Alugado",     color:"#60a5fa", bg:"rgba(96,165,250,.12)", dot:"#3b82f6" },
  em_reparo: { label:"Em Reparo",   color:"#fb923c", bg:"rgba(251,146,60,.12)", dot:"#f97316" },
  inativo:   { label:"Inativo",     color:"#6b7280", bg:"rgba(107,114,128,.12)",dot:"#4b5563" },
};
const MANUT_TIPOS = ["Troca de óleo","Revisão geral","Troca de pneus","Freios","Suspensão","Elétrica","Funilaria/Pintura","Outro"];
const PGTO_TIPOS  = ["Dinheiro","PIX","Cartão de crédito","Cartão de débito","Transferência","Outro"];
const DOC_TIPOS   = ["IPVA","Licenciamento","Seguro","DPVAT","Multa","Revisão programada","Outro"];
const MESES       = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];

// ─── Helpers ──────────────────────────────────────────────────────────────────
const fmt    = v => v ? "R$ "+Number(v).toLocaleString("pt-BR",{minimumFractionDigits:2}) : "—";
const fmtN   = v => v ? Number(v) : 0;
const dias   = d => { if(!d) return null; return Math.ceil((new Date(d)-new Date())/86400000); };
const ptDate = d => d ? new Date(d+"T12:00:00").toLocaleDateString("pt-BR") : "—";
const ls     = {
  get:(k,def=[])=>{ try{const r=localStorage.getItem(k);return r?JSON.parse(r):def;}catch(e){return def;} },
  set:(k,v)=>{ try{localStorage.setItem(k,JSON.stringify(v));}catch(e){} }
};
const isoDate = d => d ? d.slice(0,10) : "";
const today   = () => new Date().toISOString().slice(0,10);

// ─── Shared UI ────────────────────────────────────────────────────────────────
const iSty = { background:"#0f1117", border:"1px solid #2d3148", borderRadius:10, padding:"10px 14px", color:"#e2e8f0", fontSize:14, width:"100%", outline:"none", fontFamily:"inherit" };
const lSty = { fontSize:11, color:"#64748b", fontWeight:700, display:"block", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.5px" };
const cSty = { background:"#13161f", border:"1px solid #1e2130", borderRadius:14 };

function Toast({msg,color}){ return msg?<div style={{position:"fixed",bottom:28,left:"50%",transform:"translateX(-50%)",background:color||"#22c55e",color:"#fff",padding:"10px 24px",borderRadius:30,fontWeight:700,fontSize:14,zIndex:9999,boxShadow:"0 4px 24px rgba(0,0,0,.5)",animation:"fadeUp .2s ease"}}>{msg}</div>:null; }
function StatusTag({status}){ const s=STATUS_MAP[status]||STATUS_MAP.inativo; return <span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 11px",borderRadius:20,fontSize:11,fontWeight:700,background:s.bg,color:s.color,border:`1px solid ${s.color}33`}}><span style={{width:6,height:6,borderRadius:"50%",background:s.dot,flexShrink:0}}/>{s.label}</span>; }
function FI({label,id,f,set,placeholder,type="text",...r}){ return <div><label style={lSty}>{label}</label><input value={f[id]||""} onChange={e=>set(id,e.target.value)} placeholder={placeholder} type={type} style={iSty} onFocus={e=>e.target.style.borderColor="#6366f1"} onBlur={e=>e.target.style.borderColor="#2d3148"} {...r}/></div>; }
function FS({label,id,f,set,children}){ return <div><label style={lSty}>{label}</label><select value={f[id]||""} onChange={e=>set(id,e.target.value)} style={{...iSty,cursor:"pointer"}}>{children}</select></div>; }
function FM({label,id,f,set}){ return <div><label style={lSty}>{label}</label><div style={{position:"relative"}}><span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"#64748b",fontSize:13,fontFamily:"'DM Mono',monospace",pointerEvents:"none"}}>R$</span><input value={f[id]||""} onChange={e=>set(id,e.target.value)} type="number" min="0" step="0.01" placeholder="0,00" style={{...iSty,paddingLeft:38,fontFamily:"'DM Mono',monospace"}} onFocus={e=>e.target.style.borderColor="#6366f1"} onBlur={e=>e.target.style.borderColor="#2d3148"}/></div></div>; }
function G2({children,style}){ return <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,...style}}>{children}</div>; }
function Card({children,style,onClick}){ return <div style={{...cSty,padding:18,...style}} onClick={onClick}>{children}</div>; }
function ST({children,color="#6366f1"}){ return <div style={{fontSize:11,fontWeight:800,color,textTransform:"uppercase",letterSpacing:"1.5px",marginBottom:14}}>{children}</div>; }
function BP({onClick,children,c="linear-gradient(135deg,#6366f1,#8b5cf6)"}){ return <button onClick={onClick} style={{background:c,color:"#fff",border:"none",padding:"10px 18px",borderRadius:12,fontWeight:700,fontSize:14,cursor:"pointer"}}>{children}</button>; }
function BG({onClick,children}){ return <button onClick={onClick} style={{background:"transparent",color:"#94a3b8",border:"1px solid #2d3148",borderRadius:12,padding:"10px 16px",fontSize:13,cursor:"pointer"}}>{children}</button>; }
function BD({onClick,children}){ return <button onClick={onClick} style={{background:"#1a0808",color:"#ef4444",border:"1px solid #3d1515",padding:"10px 16px",borderRadius:12,fontWeight:700,fontSize:14,cursor:"pointer"}}>{children}</button>; }
function Back({onClick}){ return <button onClick={onClick} style={{background:"transparent",color:"#94a3b8",border:"1px solid #2d3148",borderRadius:8,padding:"7px 14px",fontSize:12,cursor:"pointer",marginBottom:16}}>← Voltar</button>; }

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({vehicles,manutencoes,contratos,reservas,onNav}){
  const counts={total:vehicles.length,disponivel:0,alugado:0,em_reparo:0};
  vehicles.forEach(v=>{if(counts[v.status]!==undefined)counts[v.status]++;});
  const pct=counts.total>0?Math.round((counts.alugado/counts.total)*100):0;
  const semPot=vehicles.filter(v=>v.status!=="inativo").reduce((s,v)=>s+fmtN(v.semanal),0);
  const semAtual=vehicles.filter(v=>v.status==="alugado").reduce((s,v)=>s+fmtN(v.semanal),0);
  const totalManut=manutencoes.reduce((s,m)=>s+fmtN(m.custo),0);
  const reservasPend=reservas.filter(r=>r.status==="pendente");
  const alertas=[];
  vehicles.forEach(v=>{
    const dr=dias(v.proxRevisao); if(dr!==null&&dr<=30) alertas.push({label:`Revisão: ${v.modelo} (${v.placa})`,dias:dr,cor:"#f97316"});
    const di=dias(v.ipva);        if(di!==null&&di<=60) alertas.push({label:`IPVA: ${v.modelo} (${v.placa})`,dias:di,cor:"#fbbf24"});
    const ds=dias(v.seguro);      if(ds!==null&&ds<=60) alertas.push({label:`Seguro: ${v.modelo} (${v.placa})`,dias:ds,cor:"#a78bfa"});
  });
  contratos.forEach(c=>{const df=dias(c.dataFim);if(df!==null&&df>=0&&df<=7&&c.status==="ativo"){const v=vehicles.find(x=>x.id===c.vehicleId);alertas.push({label:`Contrato vence: ${v?.placa||""}`,dias:df,cor:"#60a5fa"});}});

  return <div>
    <div style={{marginBottom:20}}>
      <h1 style={{fontSize:22,fontWeight:800,letterSpacing:"-.5px"}}>Painel da Frota</h1>
      <p style={{color:"#64748b",fontSize:13,marginTop:4}}>Visão geral em tempo real</p>
    </div>
    {reservasPend.length>0&&<div style={{background:"#0a1f2e",border:"1px solid #1e3a4d",borderRadius:12,padding:"12px 18px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
      <span style={{fontWeight:700,color:"#60a5fa",fontSize:13}}>📲 {reservasPend.length} reserva{reservasPend.length!==1?"s":""} pendente{reservasPend.length!==1?"s":""} aguardando confirmação</span>
      <button onClick={()=>onNav("reservas")} style={{background:"transparent",color:"#60a5fa",border:"1px solid #1e3a4d",padding:"6px 14px",borderRadius:8,fontWeight:700,fontSize:12,cursor:"pointer"}}>Ver reservas →</button>
    </div>}
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(110px,1fr))",gap:10,marginBottom:14}}>
      {[["🚘","Total",counts.total,"#a5b4fc"],["✅","Disponíveis",counts.disponivel,"#22c55e"],["🔑","Alugados",counts.alugado,"#60a5fa"],["🔧","Em Reparo",counts.em_reparo,"#fb923c"]].map(([e,l,v,c])=>(
        <Card key={l} style={{padding:"16px 18px"}}><div style={{fontSize:20,marginBottom:6}}>{e}</div><div style={{fontSize:26,fontWeight:800,color:c,fontFamily:"'DM Mono',monospace"}}>{v}</div><div style={{fontSize:11,color:"#64748b",marginTop:2}}>{l}</div></Card>
      ))}
    </div>
    <Card style={{borderColor:"#1e3a2e",marginBottom:12}}>
      <div style={{fontWeight:700,fontSize:14,marginBottom:12}}>💵 Receita Semanal</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
        {[["Potencial",semPot,"#a5b4fc"],["Alugados",semAtual,"#22c55e"],["Manutenção",totalManut,"#fb923c"]].map(([l,v,c])=>(
          <div key={l} style={{background:"#0f1117",borderRadius:10,padding:"12px 10px",textAlign:"center"}}><div style={{fontSize:10,color:"#64748b",marginBottom:4,fontWeight:600}}>{l}</div><div style={{fontFamily:"'DM Mono',monospace",fontWeight:800,fontSize:13,color:c}}>{fmt(v)}</div></div>
        ))}
      </div>
    </Card>
    <Card style={{marginBottom:12}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><span style={{fontWeight:700,fontSize:14}}>Taxa de Ocupação</span><span style={{fontFamily:"'DM Mono',monospace",fontSize:22,fontWeight:800,color:"#a5b4fc"}}>{pct}%</span></div>
      <div style={{background:"#0f1117",borderRadius:8,height:10,overflow:"hidden"}}><div style={{background:"linear-gradient(90deg,#6366f1,#a78bfa)",height:"100%",borderRadius:8,width:`${pct}%`,transition:"width .6s"}}/></div>
      <div style={{display:"flex",justifyContent:"space-between",marginTop:8,fontSize:12,color:"#64748b"}}><span>{counts.alugado} alugados</span><span>{counts.disponivel} disponíveis</span></div>
    </Card>
    {alertas.length>0&&<div style={{background:"#1a1108",border:"1px solid #3d2e00",borderRadius:14,padding:16,marginBottom:12}}>
      <div style={{fontWeight:700,color:"#fbbf24",marginBottom:10,fontSize:14}}>⚠️ Alertas ({alertas.length})</div>
      {alertas.sort((a,b)=>a.dias-b.dias).map((a,i)=>(
        <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid #2d2000",fontSize:13}}><span>{a.label}</span><span style={{color:a.dias<=0?"#ef4444":a.cor,fontWeight:800,fontFamily:"'DM Mono',monospace"}}>{a.dias<=0?"VENCIDA":a.dias+"d"}</span></div>
      ))}
    </div>}
    <Card>
      <div style={{fontWeight:700,fontSize:14,marginBottom:14}}>Acesso Rápido</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
        {[["🚘","Frota","lista"],["📋","Contratos","contratos"],["🔧","Manutenções","manutencoes"],["👥","Clientes","clientes"],["📄","Docs","documentos"],["💰","Financeiro","financeiro"],["📅","Calendário","calendario"],["📲","Reservas","reservas"]].map(([e,l,s])=>(
          <button key={s} onClick={()=>onNav(s)} style={{background:"#0f1117",border:"1px solid #2d3148",borderRadius:10,padding:"10px 6px",cursor:"pointer",color:"#e2e8f0",textAlign:"center"}}><div style={{fontSize:18,marginBottom:3}}>{e}</div><div style={{fontSize:10,fontWeight:600,color:"#94a3b8"}}>{l}</div></button>
        ))}
      </div>
    </Card>
  </div>;
}

// ─── FOTOS ────────────────────────────────────────────────────────────────────
function FotosVeiculo({vehicleId,fotos,onSave,label="Fotos"}){
  const inputRef=useRef();
  const minhas=fotos.filter(f=>f.vehicleId===vehicleId);

  const handleFile=async(e)=>{
    const files=[...e.target.files];
    for(const file of files){
      const reader=new FileReader();
      reader.onload=ev=>{
        const nova={id:Date.now()+Math.random(),vehicleId,dataUrl:ev.target.result,nome:file.name,data:today(),label};
        onSave(nova);
      };
      reader.readAsDataURL(file);
    }
    e.target.value="";
  };

  return <div>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
      <div style={{fontSize:11,fontWeight:800,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.5px"}}>{label} ({minhas.length})</div>
      <button onClick={()=>inputRef.current.click()} style={{background:"rgba(99,102,241,.12)",color:"#a5b4fc",border:"1px solid rgba(99,102,241,.3)",padding:"5px 12px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer"}}>📸 Adicionar</button>
    </div>
    <input ref={inputRef} type="file" accept="image/*" multiple onChange={handleFile} style={{display:"none"}}/>
    {minhas.length===0
      ?<div style={{background:"#0f1117",borderRadius:10,padding:"20px",textAlign:"center",color:"#374151",fontSize:13}}>Nenhuma foto adicionada</div>
      :<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(100px,1fr))",gap:8}}>
        {minhas.map(f=>(
          <div key={f.id} style={{position:"relative",borderRadius:10,overflow:"hidden",aspectRatio:"1",background:"#0f1117"}}>
            <img src={f.dataUrl} alt={f.nome} style={{width:"100%",height:"100%",objectFit:"cover"}}/>
            <button onClick={()=>onSave({...f,_delete:true})} style={{position:"absolute",top:4,right:4,background:"rgba(0,0,0,.7)",color:"#ef4444",border:"none",borderRadius:6,width:22,height:22,fontSize:12,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
            <div style={{position:"absolute",bottom:0,left:0,right:0,background:"rgba(0,0,0,.6)",fontSize:9,color:"#fff",padding:"3px 6px",textAlign:"center"}}>{f.data}</div>
          </div>
        ))}
      </div>
    }
  </div>;
}

// ─── ASSINATURA DIGITAL ───────────────────────────────────────────────────────
function AssinaturaDigital({onSave,onClose,titulo="Assinatura do Contrato"}){
  const canvasRef=useRef();
  const drawing=useRef(false);
  const [hasDrawn,setHasDrawn]=useState(false);

  const getPos=(e,canvas)=>{
    const rect=canvas.getBoundingClientRect();
    const src=e.touches?e.touches[0]:e;
    return {x:(src.clientX-rect.left)*(canvas.width/rect.width),y:(src.clientY-rect.top)*(canvas.height/rect.height)};
  };

  useEffect(()=>{
    const canvas=canvasRef.current;
    const ctx=canvas.getContext("2d");
    ctx.fillStyle="#0f1117";
    ctx.fillRect(0,0,canvas.width,canvas.height);
    ctx.strokeStyle="#e2e8f0";
    ctx.lineWidth=2.5;
    ctx.lineCap="round";
    ctx.lineJoin="round";
  },[]);

  const start=e=>{e.preventDefault();drawing.current=true;const canvas=canvasRef.current;const ctx=canvas.getContext("2d");const p=getPos(e,canvas);ctx.beginPath();ctx.moveTo(p.x,p.y);};
  const move=e=>{e.preventDefault();if(!drawing.current)return;setHasDrawn(true);const canvas=canvasRef.current;const ctx=canvas.getContext("2d");const p=getPos(e,canvas);ctx.lineTo(p.x,p.y);ctx.stroke();};
  const end=e=>{e.preventDefault();drawing.current=false;};
  const limpar=()=>{const canvas=canvasRef.current;const ctx=canvas.getContext("2d");ctx.fillStyle="#0f1117";ctx.fillRect(0,0,canvas.width,canvas.height);setHasDrawn(false);};
  const salvar=()=>{if(!hasDrawn)return;const dataUrl=canvasRef.current.toDataURL("image/png");onSave(dataUrl);};

  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.85)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <div style={{background:"#13161f",border:"1px solid #1e2130",borderRadius:16,padding:20,width:"100%",maxWidth:480}}>
      <div style={{fontWeight:800,fontSize:15,marginBottom:4}}>{titulo}</div>
      <div style={{color:"#64748b",fontSize:12,marginBottom:16}}>Assine com o dedo ou mouse no campo abaixo</div>
      <canvas ref={canvasRef} width={440} height={180} style={{width:"100%",height:160,borderRadius:10,border:"1px solid #2d3148",touchAction:"none",cursor:"crosshair",display:"block"}}
        onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end}
        onTouchStart={start} onTouchMove={move} onTouchEnd={end}/>
      <div style={{display:"flex",gap:10,marginTop:14,flexWrap:"wrap"}}>
        <BP onClick={salvar} c={hasDrawn?"linear-gradient(135deg,#22c55e,#16a34a)":"#2d3148"}>✅ Confirmar Assinatura</BP>
        <button onClick={limpar} style={{background:"transparent",color:"#94a3b8",border:"1px solid #2d3148",borderRadius:12,padding:"10px 16px",fontSize:13,cursor:"pointer"}}>🗑 Limpar</button>
        <BG onClick={onClose}>Cancelar</BG>
      </div>
    </div>
  </div>;
}

// ─── CALENDÁRIO DE DISPONIBILIDADE ───────────────────────────────────────────
function Calendario({vehicles,contratos,reservas,onNav}){
  const now=new Date();
  const [mes,setMes]=useState(now.getMonth());
  const [ano,setAno]=useState(now.getFullYear());

  const diasNoMes=new Date(ano,mes+1,0).getDate();
  const primeiroDia=new Date(ano,mes,1).getDay();

  const ocupacao={};
  const addOcup=(id,placa,modelo,de,ate,tipo)=>{
    if(!de||!ate)return;
    const d=new Date(de+"T12:00:00"); const a=new Date(ate+"T12:00:00");
    for(let dt=new Date(d);dt<=a;dt.setDate(dt.getDate()+1)){
      const key=`${ano}-${String(mes+1).padStart(2,"0")}-${String(dt.getDate()).padStart(2,"0")}`;
      if(!ocupacao[key])ocupacao[key]=[];
      ocupacao[key].push({id,placa,modelo,tipo});
    }
  };

  contratos.filter(c=>c.status==="ativo"||c.status==="encerrado").forEach(c=>{
    const v=vehicles.find(x=>x.id===c.vehicleId);
    if(v) addOcup(v.id,v.placa,v.modelo,c.dataInicio,c.dataFim,"contrato");
  });
  reservas.filter(r=>r.status==="confirmada"||r.status==="pendente").forEach(r=>{
    const v=vehicles.find(x=>x.id===r.vehicleId);
    if(v) addOcup(v.id,v.placa,v.modelo,r.dataInicio,r.dataFim,"reserva");
  });

  const cells=[];
  for(let i=0;i<primeiroDia;i++) cells.push(null);
  for(let d=1;d<=diasNoMes;d++) cells.push(d);

  const [selected,setSelected]=useState(null);
  const selKey=selected?`${ano}-${String(mes+1).padStart(2,"0")}-${String(selected).padStart(2,"0")}`:null;
  const selOcup=selKey?ocupacao[selKey]||[]:[];

  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:10}}>
      <h1 style={{fontSize:20,fontWeight:800}}>📅 Calendário de Disponibilidade</h1>
      <div style={{display:"flex",gap:8,alignItems:"center"}}>
        <button onClick={()=>{if(mes===0){setMes(11);setAno(a=>a-1);}else setMes(m=>m-1);}} style={{background:"#1e2130",color:"#e2e8f0",border:"none",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:14}}>‹</button>
        <span style={{fontWeight:700,fontSize:14,minWidth:120,textAlign:"center"}}>{MESES[mes]} {ano}</span>
        <button onClick={()=>{if(mes===11){setMes(0);setAno(a=>a+1);}else setMes(m=>m+1);}} style={{background:"#1e2130",color:"#e2e8f0",border:"none",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:14}}>›</button>
      </div>
    </div>

    {/* Legenda */}
    <div style={{display:"flex",gap:14,marginBottom:14,flexWrap:"wrap",fontSize:12}}>
      {[["#22c55e","Disponível"],["#3b82f6","Contrato"],["#f97316","Reserva"],["#8b5cf6","Hoje"]].map(([c,l])=>(
        <div key={l} style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:10,height:10,borderRadius:3,background:c}}/><span style={{color:"#94a3b8"}}>{l}</span></div>
      ))}
    </div>

    {/* Grade */}
    <Card style={{padding:14,marginBottom:14}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3,marginBottom:6}}>
        {["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"].map(d=><div key={d} style={{textAlign:"center",fontSize:10,color:"#64748b",fontWeight:700,padding:"4px 0"}}>{d}</div>)}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3}}>
        {cells.map((d,i)=>{
          if(!d) return <div key={i}/>;
          const key=`${ano}-${String(mes+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
          const ocup=ocupacao[key]||[];
          const isToday=d===now.getDate()&&mes===now.getMonth()&&ano===now.getFullYear();
          const isSel=d===selected;
          const hasContrato=ocup.some(o=>o.tipo==="contrato");
          const hasReserva=ocup.some(o=>o.tipo==="reserva");
          const bg=isToday?"#4c1d95":isSel?"#1e2130":hasContrato?"rgba(59,130,246,.15)":hasReserva?"rgba(249,115,22,.15)":"transparent";
          const borderColor=isToday?"#8b5cf6":isSel?"#6366f1":hasContrato?"#3b82f6":hasReserva?"#f97316":ocup.length===0?"#1e2130":"#22c55e22";
          return <div key={i} onClick={()=>setSelected(d===selected?null:d)} style={{borderRadius:8,border:`1px solid ${borderColor}`,background:bg,padding:"6px 4px",textAlign:"center",cursor:"pointer",transition:"all .1s"}}>
            <div style={{fontSize:13,fontWeight:isToday||isSel?800:500,color:isToday?"#c4b5fd":ocup.length>0?"#e2e8f0":"#64748b"}}>{d}</div>
            {ocup.length>0&&<div style={{fontSize:9,color:hasContrato?"#60a5fa":hasReserva?"#fb923c":"#22c55e",marginTop:2,fontWeight:700}}>{ocup.length}</div>}
          </div>;
        })}
      </div>
    </Card>

    {/* Detalhe do dia selecionado */}
    {selected&&<Card>
      <div style={{fontWeight:700,fontSize:14,marginBottom:12}}>
        {String(selected).padStart(2,"0")}/{String(mes+1).padStart(2,"0")}/{ano}
        {selOcup.length===0&&<span style={{color:"#22c55e",marginLeft:8,fontSize:13}}>— Todos disponíveis</span>}
      </div>
      {selOcup.length===0
        ?<div style={{color:"#64748b",fontSize:13}}>{vehicles.length} veículos disponíveis neste dia.</div>
        :<div style={{display:"flex",flexDirection:"column",gap:8}}>
          {selOcup.map((o,i)=>(
            <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:"#0f1117",borderRadius:10,padding:"10px 14px"}}>
              <div><div style={{fontWeight:700,fontSize:13}}>{o.modelo}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:"#94a3b8"}}>{o.placa}</div></div>
              <span style={{fontSize:11,fontWeight:700,color:o.tipo==="contrato"?"#60a5fa":"#fb923c",background:o.tipo==="contrato"?"rgba(59,130,246,.12)":"rgba(249,115,22,.12)",border:`1px solid ${o.tipo==="contrato"?"rgba(59,130,246,.3)":"rgba(249,115,22,.3)"}`,padding:"3px 10px",borderRadius:20}}>
                {o.tipo==="contrato"?"Contrato":"Reserva"}
              </span>
            </div>
          ))}
          {vehicles.filter(v=>!selOcup.find(o=>o.id===v.id)&&v.status!=="inativo"&&v.status!=="em_reparo").map(v=>(
            <div key={v.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:"#0f1117",borderRadius:10,padding:"10px 14px"}}>
              <div><div style={{fontWeight:700,fontSize:13}}>{v.modelo}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:"#94a3b8"}}>{v.placa}</div></div>
              <span style={{fontSize:11,fontWeight:700,color:"#22c55e",background:"rgba(34,197,94,.12)",border:"1px solid rgba(34,197,94,.3)",padding:"3px 10px",borderRadius:20}}>Disponível</span>
            </div>
          ))}
        </div>
      }
    </Card>}
  </div>;
}

// ─── RESERVAS ONLINE ──────────────────────────────────────────────────────────
const emptyReserva={vehicleId:"",clienteNome:"",clienteTel:"",clienteEmail:"",dataInicio:"",dataFim:"",obs:"",status:"pendente"};

function Reservas({vehicles,reservas,clientes,onSave,onDelete,showToast}){
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [f,setF]=useState({...emptyReserva});
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const p={f,set};
  const [showLink,setShowLink]=useState(false);
  const [filterStatus,setFilterStatus]=useState("todos");

  const link=`${window.location.origin}${window.location.pathname}?reserva=1`;

  const openNew=()=>{setF({...emptyReserva,dataInicio:today()});setEditing(null);setShowForm(true);};
  const openEdit=r=>{setF({...r});setEditing(r.id);setShowForm(true);};
  const save=()=>{
    if(!f.vehicleId||!f.clienteNome||!f.dataInicio){showToast("Veículo, cliente e data são obrigatórios","#ef4444");return;}
    onSave({...f,id:editing||Date.now()});setShowForm(false);showToast("📲 Reserva salva!");
  };

  const confirmar=r=>{onSave({...r,status:"confirmada"});showToast("✅ Reserva confirmada!");};
  const cancelar=r=>{onSave({...r,status:"cancelada"});showToast("❌ Reserva cancelada");};

  const filtered=filterStatus==="todos"?reservas:reservas.filter(r=>r.status===filterStatus);
  const pendentes=reservas.filter(r=>r.status==="pendente");
  const getV=id=>{const v=vehicles.find(v=>v.id===id||v.id===Number(id));return v?`${v.placa} – ${v.modelo}`:"—";};

  const statusColor={pendente:"#fbbf24",confirmada:"#22c55e",cancelada:"#6b7280"};
  const statusBg={pendente:"rgba(251,191,36,.12)",confirmada:"rgba(34,197,94,.12)",cancelada:"rgba(107,114,128,.12)"};

  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
      <div><h1 style={{fontSize:20,fontWeight:800}}>📲 Reservas Online</h1><p style={{color:"#64748b",fontSize:13,marginTop:3}}>{reservas.length} total · {pendentes.length} pendentes</p></div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>setShowLink(v=>!v)} style={{background:"rgba(96,165,250,.12)",color:"#60a5fa",border:"1px solid rgba(96,165,250,.3)",padding:"9px 14px",borderRadius:10,fontWeight:700,fontSize:12,cursor:"pointer"}}>🔗 Link</button>
        <BP onClick={openNew} c="linear-gradient(135deg,#3b82f6,#60a5fa)">+ Nova</BP>
      </div>
    </div>

    {/* Link para o cliente */}
    {showLink&&<Card style={{borderColor:"#1e3a4d",marginBottom:16}}>
      <div style={{fontWeight:700,fontSize:14,marginBottom:8}}>🔗 Link de Reserva para Clientes</div>
      <p style={{color:"#64748b",fontSize:13,marginBottom:12}}>Envie este link para o cliente preencher os dados. A reserva chega aqui como "Pendente" para você confirmar.</p>
      <div style={{background:"#0f1117",border:"1px solid #2d3148",borderRadius:10,padding:"10px 14px",fontFamily:"'DM Mono',monospace",fontSize:12,color:"#a5b4fc",wordBreak:"break-all",marginBottom:10}}>{link}</div>
      <button onClick={()=>{navigator.clipboard?.writeText(link);showToast("✅ Link copiado!");}} style={{background:"rgba(96,165,250,.12)",color:"#60a5fa",border:"1px solid rgba(96,165,250,.3)",padding:"8px 16px",borderRadius:10,fontWeight:700,fontSize:13,cursor:"pointer"}}>📋 Copiar Link</button>
      <div style={{marginTop:10,fontSize:12,color:"#64748b"}}>💡 Você pode enviar via WhatsApp, e-mail ou Instagram.</div>
    </Card>}

    {/* Pendentes em destaque */}
    {pendentes.length>0&&<div style={{background:"#1a1a08",border:"1px solid #3d3800",borderRadius:14,padding:16,marginBottom:14}}>
      <div style={{fontWeight:700,color:"#fbbf24",marginBottom:10,fontSize:14}}>⏳ Aguardando Confirmação ({pendentes.length})</div>
      {pendentes.map(r=>(
        <div key={r.id} style={{background:"#0f1117",borderRadius:10,padding:"12px 14px",marginBottom:8}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8,marginBottom:8}}>
            <div><div style={{fontWeight:700,fontSize:14}}>{r.clienteNome}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#94a3b8"}}>{getV(r.vehicleId)}</div></div>
            <div style={{fontSize:12,color:"#64748b"}}>{ptDate(r.dataInicio)} → {ptDate(r.dataFim)}</div>
          </div>
          {r.clienteTel&&<div style={{fontSize:12,color:"#64748b",marginBottom:8}}>📱 {r.clienteTel}</div>}
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>confirmar(r)} style={{background:"rgba(34,197,94,.12)",color:"#22c55e",border:"1px solid rgba(34,197,94,.3)",padding:"7px 14px",borderRadius:8,fontWeight:700,fontSize:12,cursor:"pointer"}}>✅ Confirmar</button>
            <button onClick={()=>cancelar(r)} style={{background:"rgba(239,68,68,.08)",color:"#ef4444",border:"1px solid rgba(239,68,68,.2)",padding:"7px 14px",borderRadius:8,fontWeight:700,fontSize:12,cursor:"pointer"}}>❌ Cancelar</button>
            <button onClick={()=>openEdit(r)} style={{background:"transparent",color:"#64748b",border:"1px solid #2d3148",padding:"7px 12px",borderRadius:8,fontSize:12,cursor:"pointer"}}>✏️</button>
          </div>
        </div>
      ))}
    </div>}

    {/* Formulário */}
    {showForm&&<Card style={{borderColor:"#1e2a3d",marginBottom:16}}>
      <ST color="#60a5fa">{editing?"✏️ Editar Reserva":"➕ Nova Reserva"}</ST>
      <G2 style={{marginBottom:16}}>
        <FS label="Veículo *" id="vehicleId" {...p}><option value="">Selecione...</option>{vehicles.filter(v=>v.status==="disponivel"||String(v.id)===String(f.vehicleId)).map(v=><option key={v.id} value={v.id}>{v.placa} – {v.modelo}</option>)}</FS>
        <FI label="Nome do Cliente *" id="clienteNome" placeholder="Nome completo" {...p}/>
        <FI label="Telefone / WhatsApp" id="clienteTel" placeholder="(11) 99999-9999" {...p}/>
        <FI label="E-mail" id="clienteEmail" type="email" {...p}/>
        <FI label="Data Início *" id="dataInicio" type="date" {...p}/>
        <FI label="Data Fim" id="dataFim" type="date" {...p}/>
        <FS label="Status" id="status" {...p}><option value="pendente">Pendente</option><option value="confirmada">Confirmada</option><option value="cancelada">Cancelada</option></FS>
        <FI label="Observações" id="obs" {...p}/>
      </G2>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><BP onClick={save} c="linear-gradient(135deg,#3b82f6,#60a5fa)">💾 Salvar</BP>{editing&&<BD onClick={()=>{if(confirm("Excluir?"))onDelete(editing);setShowForm(false);}}>🗑</BD>}<BG onClick={()=>setShowForm(false)}>Cancelar</BG></div>
    </Card>}

    {/* Filtro + lista */}
    <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{...iSty,width:"auto",minWidth:160,marginBottom:12,cursor:"pointer"}}>
      <option value="todos">Todas</option><option value="pendente">Pendentes</option><option value="confirmada">Confirmadas</option><option value="cancelada">Canceladas</option>
    </select>
    <div style={{display:"flex",flexDirection:"column",gap:8}}>
      {filtered.filter(r=>r.status!=="pendente").length===0?<div style={{textAlign:"center",color:"#64748b",padding:"30px 0"}}>Nenhuma reserva</div>
      :filtered.filter(r=>r.status!=="pendente").sort((a,b)=>b.dataInicio?.localeCompare(a.dataInicio)).map(r=>(
        <div key={r.id} onClick={()=>openEdit(r)} style={{...cSty,padding:"13px 16px",cursor:"pointer",transition:"transform .12s"}} onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"} onMouseLeave={e=>e.currentTarget.style.transform=""}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div><div style={{fontWeight:700,fontSize:14}}>{r.clienteNome}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#94a3b8",marginTop:2}}>{getV(r.vehicleId)}</div></div>
            <span style={{fontSize:11,fontWeight:700,color:statusColor[r.status],background:statusBg[r.status],border:`1px solid ${statusColor[r.status]}33`,padding:"3px 10px",borderRadius:20}}>{r.status.toUpperCase()}</span>
          </div>
          <div style={{display:"flex",gap:12,marginTop:8,fontSize:12,color:"#64748b",flexWrap:"wrap"}}>
            {r.dataInicio&&<span>📅 {ptDate(r.dataInicio)}</span>}{r.dataFim&&<span>🏁 {ptDate(r.dataFim)}</span>}{r.clienteTel&&<span>📱 {r.clienteTel}</span>}
          </div>
        </div>
      ))}
    </div>
  </div>;
}

// ─── FROTA ────────────────────────────────────────────────────────────────────
const emptyV={placa:"",modelo:"",ano:"",cor:"",status:"disponivel",km:"",kmRevisao:"",proxRevisao:"",ipva:"",seguro:"",locatario:"",obs:"",diaria:"",semanal:"",mensal:"",caucao:""};

function Frota({vehicles,fotos,onSave,onDelete,onFoto}){
  const [search,setSearch]=useState("");
  const [fst,setFst]=useState("todos");
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [showDetail,setShowDetail]=useState(null);
  const [f,setF]=useState({...emptyV});
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const p={f,set};

  const openNew=()=>{setF({...emptyV});setEditing(null);setShowDetail(null);setShowForm(true);};
  const openEdit=v=>{setF({...v});setEditing(v.id);setShowDetail(null);setShowForm(true);};
  const openDetail=v=>{setShowDetail(v);setShowForm(false);};
  const save=()=>{
    if(!f.placa||!f.modelo){alert("Placa e modelo são obrigatórios");return;}
    onSave({...f,id:editing||Date.now()});setShowForm(false);setEditing(null);
  };
  const del=()=>{if(!confirm("Excluir veículo?"))return;onDelete(editing);setShowForm(false);};

  const filtered=vehicles.filter(v=>(fst==="todos"||v.status===fst)&&(v.placa?.toLowerCase().includes(search.toLowerCase())||v.modelo?.toLowerCase().includes(search.toLowerCase())));

  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18}}>
      <div><h1 style={{fontSize:20,fontWeight:800}}>🚘 Frota Completa</h1><p style={{color:"#64748b",fontSize:13,marginTop:3}}>{vehicles.length} veículos</p></div>
      <BP onClick={openNew}>+ Novo</BP>
    </div>
    <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Placa ou modelo..." style={{...iSty,flex:1,minWidth:160}}/>
      <select value={fst} onChange={e=>setFst(e.target.value)} style={{...iSty,width:"auto",minWidth:130,cursor:"pointer"}}><option value="todos">Todos</option><option value="disponivel">Disponível</option><option value="alugado">Alugado</option><option value="em_reparo">Em Reparo</option><option value="inativo">Inativo</option></select>
    </div>

    {showDetail&&!showForm&&<Card style={{marginBottom:18}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div><h2 style={{fontSize:18,fontWeight:800}}>{showDetail.modelo}</h2><span style={{fontFamily:"'DM Mono',monospace",color:"#94a3b8",fontSize:14}}>{showDetail.placa}</span></div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}><button onClick={()=>openEdit(showDetail)} style={{background:"linear-gradient(135deg,#6366f1,#8b5cf6)",color:"#fff",border:"none",padding:"8px 14px",borderRadius:10,fontWeight:700,fontSize:12,cursor:"pointer"}}>✏️ Editar</button><StatusTag status={showDetail.status}/><button onClick={()=>setShowDetail(null)} style={{background:"transparent",color:"#64748b",border:"none",fontSize:18,cursor:"pointer"}}>×</button></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
        {[["Ano",showDetail.ano],["Cor",showDetail.cor],["KM Atual",showDetail.km?Number(showDetail.km).toLocaleString()+" km":""],["Próx. Revisão",ptDate(showDetail.proxRevisao)],["IPVA",ptDate(showDetail.ipva)],["Seguro",ptDate(showDetail.seguro)],["Semanal",showDetail.semanal?`R$ ${showDetail.semanal}`:""],["Locatário",showDetail.locatario]].filter(([,v])=>v).map(([l,v])=>(
          <div key={l} style={{background:"#0f1117",borderRadius:10,padding:"10px 14px"}}><div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>{l}</div><div style={{fontWeight:600,fontSize:13}}>{v}</div></div>
        ))}
      </div>
      {/* Fotos do veículo */}
      <ST color="#a5b4fc">📸 Fotos</ST>
      <G2 style={{marginBottom:16}}>
        <FotosVeiculo vehicleId={showDetail.id} fotos={fotos.filter(f=>f.label==="saida"||f.label==="Fotos")} onSave={onFoto} label="saida"/>
        <FotosVeiculo vehicleId={showDetail.id} fotos={fotos.filter(f=>f.label==="retorno")} onSave={onFoto} label="retorno"/>
      </G2>
    </Card>}

    {showForm&&<Card style={{marginBottom:18}}>
      <ST>{editing?"✏️ Editar Veículo":"➕ Novo Veículo"}</ST>
      <G2 style={{marginBottom:16}}><FI label="Placa *" id="placa" placeholder="EYA8E54" {...p}/><FI label="Modelo *" id="modelo" placeholder="VW Voyage" {...p}/><FI label="Ano" id="ano" {...p}/><FI label="Cor" id="cor" {...p}/><FI label="KM Atual" id="km" type="number" {...p}/><FI label="KM Próx. Revisão" id="kmRevisao" type="number" {...p}/><FI label="Data Revisão" id="proxRevisao" type="date" {...p}/><FS label="Status" id="status" {...p}><option value="disponivel">Disponível</option><option value="alugado">Alugado</option><option value="em_reparo">Em Reparo</option><option value="inativo">Inativo</option></FS><FI label="Venc. IPVA" id="ipva" type="date" {...p}/><FI label="Venc. Seguro" id="seguro" type="date" {...p}/><FI label="Locatário Atual" id="locatario" {...p}/><FI label="Observações" id="obs" {...p}/></G2>
      <ST color="#22c55e">💰 Valores</ST>
      <G2 style={{marginBottom:20}}><FM label="Diária" id="diaria" {...p}/><FM label="Semanal" id="semanal" {...p}/><FM label="Mensal" id="mensal" {...p}/><FM label="Caução" id="caucao" {...p}/></G2>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><BP onClick={save}>💾 Salvar</BP>{editing&&<BD onClick={del}>🗑 Excluir</BD>}<BG onClick={()=>setShowForm(false)}>Cancelar</BG></div>
    </Card>}

    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      {filtered.length===0?<div style={{textAlign:"center",color:"#64748b",padding:"50px 0"}}>Nenhum veículo encontrado</div>
      :filtered.map(v=>{
        const nFotos=fotos.filter(f=>f.vehicleId===v.id).length;
        return <div key={v.id} onClick={()=>openDetail(showDetail?.id===v.id?null:v)} style={{...cSty,padding:"14px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,cursor:"pointer",transition:"transform .12s",borderColor:showDetail?.id===v.id?"#6366f1":"#1e2130"}} onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"} onMouseLeave={e=>e.currentTarget.style.transform=""}>
          <div style={{display:"flex",alignItems:"center",gap:12,minWidth:0}}>
            <div style={{width:10,height:10,borderRadius:"50%",background:STATUS_MAP[v.status]?.dot,flexShrink:0}}/>
            <div>
              <div style={{fontWeight:700,fontSize:14}}>{v.modelo} <span style={{fontWeight:400,color:"#64748b"}}>{v.ano}</span></div>
              <div style={{display:"flex",gap:8,marginTop:3,flexWrap:"wrap"}}>
                <span style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#94a3b8"}}>{v.placa}</span>
                {v.semanal&&<span style={{fontSize:12,color:"#22c55e",fontFamily:"'DM Mono',monospace"}}>R${v.semanal}/sem</span>}
                {v.locatario&&<span style={{fontSize:12,color:"#60a5fa"}}>👤 {v.locatario}</span>}
                {nFotos>0&&<span style={{fontSize:11,color:"#64748b"}}>📸 {nFotos}</span>}
              </div>
            </div>
          </div>
          <StatusTag status={v.status}/>
        </div>;
      })}
    </div>
  </div>;
}

// ─── CONTRATOS ────────────────────────────────────────────────────────────────
const emptyC={vehicleId:"",clienteId:"",dataInicio:"",dataFim:"",valorTotal:"",formaPgto:"PIX",kmSaida:"",kmRetorno:"",combustivelSaida:"",combustivelRetorno:"",avariasSaida:"",avariasRetorno:"",status:"ativo",obs:"",assinatura:null};

function Contratos({vehicles,clientes,contratos,fotos,onSave,onDelete,onFoto,showToast}){
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [f,setF]=useState({...emptyC});
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const p={f,set};
  const [showAssin,setShowAssin]=useState(false);

  const openNew=()=>{setF({...emptyC,dataInicio:today()});setEditing(null);setShowForm(true);};
  const openEdit=c=>{setF({...c});setEditing(c.id);setShowForm(true);};
  const save=()=>{
    if(!f.vehicleId||!f.dataInicio){showToast("Veículo e data de início são obrigatórios","#ef4444");return;}
    onSave({...f,id:editing||Date.now()});setShowForm(false);showToast("📋 Contrato salvo!");
  };
  const del=()=>{if(!confirm("Excluir?"))return;onDelete(editing);setShowForm(false);};
  const salvarAssin=dataUrl=>{set("assinatura",dataUrl);setShowAssin(false);showToast("✅ Assinatura capturada!");};

  const getV=id=>vehicles.find(v=>v.id===id||v.id===Number(id));
  const getC=id=>clientes.find(c=>c.id===id||c.id===Number(id));
  const receitaTotal=contratos.reduce((s,c)=>s+fmtN(c.valorTotal),0);

  return <div>
    {showAssin&&<AssinaturaDigital titulo="Assinatura do Contrato" onSave={salvarAssin} onClose={()=>setShowAssin(false)}/>}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
      <div><h1 style={{fontSize:20,fontWeight:800}}>📋 Contratos</h1><p style={{color:"#64748b",fontSize:13,marginTop:3}}>{contratos.length} total</p></div>
      <BP onClick={openNew} c="linear-gradient(135deg,#3b82f6,#60a5fa)">+ Novo</BP>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>
      {[["Ativos",contratos.filter(c=>c.status==="ativo").length,"#22c55e"],["Encerrados",contratos.filter(c=>c.status==="encerrado").length,"#60a5fa"],["Receita Total",fmt(receitaTotal),"#fb923c"]].map(([l,v,c])=>(
        <Card key={l}><div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:6}}>{l}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:typeof v==="number"?22:15,fontWeight:800,color:c}}>{v}</div></Card>
      ))}
    </div>
    {showForm&&<Card style={{borderColor:"#1e2a3d",marginBottom:16}}>
      <ST color="#60a5fa">{editing?"✏️ Editar":"➕ Novo Contrato"}</ST>
      <G2 style={{marginBottom:14}}>
        <FS label="Veículo *" id="vehicleId" {...p}><option value="">Selecione...</option>{vehicles.map(v=><option key={v.id} value={v.id}>{v.placa} – {v.modelo}</option>)}</FS>
        <FS label="Cliente" id="clienteId" {...p}><option value="">Avulso</option>{clientes.map(c=><option key={c.id} value={c.id}>{c.nome}</option>)}</FS>
        <FI label="Data Início *" id="dataInicio" type="date" {...p}/><FI label="Data Fim" id="dataFim" type="date" {...p}/>
        <FM label="Valor Total" id="valorTotal" {...p}/><FS label="Pagamento" id="formaPgto" {...p}>{PGTO_TIPOS.map(t=><option key={t} value={t}>{t}</option>)}</FS>
        <FS label="Status" id="status" {...p}><option value="ativo">Ativo</option><option value="encerrado">Encerrado</option><option value="cancelado">Cancelado</option></FS>
        <FI label="Observações" id="obs" {...p}/>
      </G2>
      <ST color="#94a3b8">🚗 Checklist Saída</ST>
      <G2 style={{marginBottom:14}}>
        <FI label="KM Saída" id="kmSaida" type="number" {...p}/>
        <FS label="Combustível" id="combustivelSaida" {...p}><option value="">—</option>{["Cheio","3/4","Meio","1/4","Reserva"].map(o=><option key={o} value={o}>{o}</option>)}</FS>
        <div style={{gridColumn:"1/-1"}}><FI label="Avarias na Saída" id="avariasSaida" placeholder="Riscos, amassados..." {...p}/></div>
      </G2>
      {/* Fotos saída */}
      {f.vehicleId&&<div style={{marginBottom:14}}>
        <FotosVeiculo vehicleId={Number(f.vehicleId)||f.vehicleId} fotos={fotos} onSave={onFoto} label="saida"/>
      </div>}
      <ST color="#94a3b8">🔄 Checklist Retorno</ST>
      <G2 style={{marginBottom:14}}>
        <FI label="KM Retorno" id="kmRetorno" type="number" {...p}/>
        <FS label="Combustível" id="combustivelRetorno" {...p}><option value="">—</option>{["Cheio","3/4","Meio","1/4","Reserva"].map(o=><option key={o} value={o}>{o}</option>)}</FS>
        <div style={{gridColumn:"1/-1"}}><FI label="Avarias no Retorno" id="avariasRetorno" {...p}/></div>
      </G2>
      {/* Fotos retorno */}
      {f.vehicleId&&<div style={{marginBottom:16}}>
        <FotosVeiculo vehicleId={Number(f.vehicleId)||f.vehicleId} fotos={fotos} onSave={onFoto} label="retorno"/>
      </div>}
      {/* Assinatura */}
      <ST color="#a78bfa">✍️ Assinatura Digital</ST>
      <div style={{marginBottom:20}}>
        {f.assinatura
          ?<div style={{display:"flex",alignItems:"center",gap:14,background:"#0f1117",borderRadius:10,padding:"12px 16px"}}>
            <img src={f.assinatura} alt="assinatura" style={{height:50,borderRadius:6,border:"1px solid #2d3148"}}/>
            <div style={{flex:1}}><div style={{fontSize:13,fontWeight:600,color:"#22c55e"}}>✅ Assinatura capturada</div></div>
            <button onClick={()=>setShowAssin(true)} style={{background:"transparent",color:"#64748b",border:"1px solid #2d3148",padding:"6px 12px",borderRadius:8,fontSize:12,cursor:"pointer"}}>Refazer</button>
          </div>
          :<button onClick={()=>setShowAssin(true)} style={{background:"rgba(167,139,250,.1)",color:"#a78bfa",border:"1px solid rgba(167,139,250,.3)",padding:"11px 20px",borderRadius:12,fontWeight:700,fontSize:13,cursor:"pointer"}}>✍️ Capturar Assinatura do Cliente</button>
        }
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><BP onClick={save} c="linear-gradient(135deg,#3b82f6,#60a5fa)">💾 Salvar</BP>{editing&&<BD onClick={del}>🗑</BD>}<BG onClick={()=>setShowForm(false)}>Cancelar</BG></div>
    </Card>}

    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      {contratos.length===0?<div style={{textAlign:"center",color:"#64748b",padding:"50px 0"}}>Nenhum contrato</div>
      :[...contratos].sort((a,b)=>(b.dataInicio||"").localeCompare(a.dataInicio||"")).map(c=>{
        const v=getV(c.vehicleId); const cl=getC(c.clienteId);
        const df=dias(c.dataFim);
        return <div key={c.id} onClick={()=>openEdit(c)} style={{...cSty,padding:"14px 18px",cursor:"pointer",transition:"transform .12s"}} onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"} onMouseLeave={e=>e.currentTarget.style.transform=""}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:8,flexWrap:"wrap",gap:8}}>
            <div>
              <div style={{fontWeight:700,fontSize:14}}>{v?`${v.placa} – ${v.modelo}`:"—"}</div>
              {cl&&<div style={{fontSize:12,color:"#60a5fa",marginTop:2}}>👤 {cl.nome}</div>}
            </div>
            <div style={{textAlign:"right"}}>
              {c.valorTotal&&<div style={{fontFamily:"'DM Mono',monospace",fontWeight:800,fontSize:15,color:"#22c55e"}}>{fmt(c.valorTotal)}</div>}
              <div style={{display:"flex",gap:6,alignItems:"center",justifyContent:"flex-end",marginTop:2}}>
                {c.assinatura&&<span style={{fontSize:10,color:"#a78bfa"}}>✍️</span>}
                <span style={{fontSize:11,fontWeight:700,color:c.status==="ativo"?"#22c55e":c.status==="encerrado"?"#60a5fa":"#6b7280"}}>{c.status?.toUpperCase()}</span>
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:12,fontSize:12,color:"#64748b",flexWrap:"wrap"}}>
            {c.dataInicio&&<span>📅 {ptDate(c.dataInicio)}</span>}{c.dataFim&&<span>🏁 {ptDate(c.dataFim)}</span>}{c.formaPgto&&<span>💳 {c.formaPgto}</span>}
            {df!==null&&df>=0&&df<=7&&c.status==="ativo"&&<span style={{color:"#f97316",fontWeight:700}}>⚠️ Vence em {df}d</span>}
          </div>
        </div>;
      })}
    </div>
  </div>;
}

// ─── MANUTENÇÕES ──────────────────────────────────────────────────────────────
const emptyM={vehicleId:"",data:"",tipo:"",descricao:"",km:"",custo:"",oficina:""};
function Manutencoes({vehicles,manutencoes,onSave,onDelete,showToast}){
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [filterV,setFilterV]=useState("todos");
  const [f,setF]=useState({...emptyM});
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const p={f,set};
  const openNew=()=>{setF({...emptyM,data:today()});setEditing(null);setShowForm(true);};
  const openEdit=m=>{setF({...m});setEditing(m.id);setShowForm(true);};
  const save=()=>{if(!f.vehicleId||!f.tipo){showToast("Veículo e tipo são obrigatórios","#ef4444");return;}onSave({...f,id:editing||Date.now()});setShowForm(false);showToast("🔧 Manutenção salva!");};
  const filtered=filterV==="todos"?manutencoes:manutencoes.filter(m=>String(m.vehicleId)===filterV);
  const totalGasto=filtered.reduce((s,m)=>s+fmtN(m.custo),0);
  const getV=id=>{const v=vehicles.find(v=>String(v.id)===String(id));return v?`${v.placa} – ${v.modelo}`:"—";};
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
      <div><h1 style={{fontSize:20,fontWeight:800}}>🔧 Manutenções</h1><p style={{color:"#64748b",fontSize:13,marginTop:3}}>{manutencoes.length} registros</p></div>
      <BP onClick={openNew} c="linear-gradient(135deg,#f97316,#fb923c)">+ Registrar</BP>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
      <Card style={{borderColor:"#2d1e0f"}}><div style={lSty}>Registros</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:24,fontWeight:800,color:"#fb923c"}}>{manutencoes.length}</div></Card>
      <Card style={{borderColor:"#2d1e0f"}}><div style={lSty}>Total gasto</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:18,fontWeight:800,color:"#fb923c"}}>{fmt(totalGasto)}</div></Card>
    </div>
    <select value={filterV} onChange={e=>setFilterV(e.target.value)} style={{...iSty,width:"auto",minWidth:220,marginBottom:14,cursor:"pointer"}}><option value="todos">Todos os veículos</option>{vehicles.map(v=><option key={v.id} value={String(v.id)}>{v.placa} – {v.modelo}</option>)}</select>
    {showForm&&<Card style={{borderColor:"#2d1e0f",marginBottom:16}}>
      <ST color="#fb923c">{editing?"✏️ Editar":"➕ Novo"}</ST>
      <G2 style={{marginBottom:16}}>
        <FS label="Veículo *" id="vehicleId" {...p}><option value="">Selecione...</option>{vehicles.map(v=><option key={v.id} value={String(v.id)}>{v.placa} – {v.modelo}</option>)}</FS>
        <FI label="Data *" id="data" type="date" {...p}/>
        <FS label="Tipo *" id="tipo" {...p}><option value="">Selecione...</option>{MANUT_TIPOS.map(t=><option key={t} value={t}>{t}</option>)}</FS>
        <FI label="Oficina" id="oficina" {...p}/><FI label="KM" id="km" type="number" {...p}/><FM label="Custo" id="custo" {...p}/>
        <div style={{gridColumn:"1/-1"}}><FI label="Descrição" id="descricao" placeholder="Detalhes..." {...p}/></div>
      </G2>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><BP onClick={save} c="linear-gradient(135deg,#f97316,#fb923c)">💾 Salvar</BP>{editing&&<BD onClick={()=>{if(confirm("Excluir?"))onDelete(editing);setShowForm(false);}}>🗑</BD>}<BG onClick={()=>setShowForm(false)}>Cancelar</BG></div>
    </Card>}
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      {filtered.length===0?<div style={{textAlign:"center",color:"#64748b",padding:"50px 0"}}>Nenhuma manutenção</div>
      :[...filtered].sort((a,b)=>(b.data||"").localeCompare(a.data||"")).map(m=>(
        <div key={m.id} onClick={()=>openEdit(m)} style={{...cSty,padding:"14px 18px",cursor:"pointer",transition:"transform .12s"}} onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"} onMouseLeave={e=>e.currentTarget.style.transform=""}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
            <div><div style={{fontWeight:700,fontSize:14}}>{m.tipo}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#94a3b8",marginTop:2}}>{getV(m.vehicleId)}</div></div>
            <div style={{textAlign:"right"}}>{m.custo&&<div style={{fontFamily:"'DM Mono',monospace",fontWeight:800,fontSize:15,color:"#fb923c"}}>{fmt(m.custo)}</div>}<div style={{fontSize:12,color:"#64748b",marginTop:2}}>{ptDate(m.data)}</div></div>
          </div>
          <div style={{display:"flex",gap:10,fontSize:12,color:"#64748b",flexWrap:"wrap"}}>{m.oficina&&<span>🏪 {m.oficina}</span>}{m.km&&<span style={{fontFamily:"'DM Mono',monospace"}}>📍 {Number(m.km).toLocaleString()} km</span>}{m.descricao&&<span>📝 {m.descricao}</span>}</div>
        </div>
      ))}
    </div>
  </div>;
}

// ─── CLIENTES ─────────────────────────────────────────────────────────────────
const emptyCl={nome:"",cpf:"",telefone:"",email:"",cnh:"",cnhValidade:"",endereco:"",obs:""};
function Clientes({clientes,contratos,onSave,onDelete,showToast}){
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [search,setSearch]=useState("");
  const [f,setF]=useState({...emptyCl});
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const p={f,set};
  const openNew=()=>{setF({...emptyCl});setEditing(null);setShowForm(true);};
  const openEdit=c=>{setF({...c});setEditing(c.id);setShowForm(true);};
  const save=()=>{if(!f.nome){showToast("Nome é obrigatório","#ef4444");return;}onSave({...f,id:editing||Date.now()});setShowForm(false);showToast("👤 Cliente salvo!");};
  const filtered=clientes.filter(c=>c.nome?.toLowerCase().includes(search.toLowerCase())||c.cpf?.includes(search)||c.telefone?.includes(search));
  const getLoc=id=>contratos.filter(c=>String(c.clienteId)===String(id)).length;
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
      <div><h1 style={{fontSize:20,fontWeight:800}}>👥 Clientes</h1><p style={{color:"#64748b",fontSize:13,marginTop:3}}>{clientes.length} cadastrados</p></div>
      <BP onClick={openNew} c="linear-gradient(135deg,#8b5cf6,#a78bfa)">+ Novo</BP>
    </div>
    <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Nome, CPF ou telefone..." style={{...iSty,marginBottom:14}}/>
    {showForm&&<Card style={{borderColor:"#2d1f3d",marginBottom:16}}>
      <ST color="#a78bfa">{editing?"✏️ Editar":"➕ Novo Cliente"}</ST>
      <G2 style={{marginBottom:20}}>
        <FI label="Nome *" id="nome" placeholder="Nome completo" {...p}/><FI label="Telefone" id="telefone" placeholder="(11) 99999-9999" {...p}/>
        <FI label="CPF" id="cpf" placeholder="000.000.000-00" {...p}/><FI label="E-mail" id="email" type="email" {...p}/>
        <FI label="Nº CNH" id="cnh" {...p}/><FI label="Validade CNH" id="cnhValidade" type="date" {...p}/>
        <div style={{gridColumn:"1/-1"}}><FI label="Endereço" id="endereco" {...p}/></div>
        <div style={{gridColumn:"1/-1"}}><FI label="Observações" id="obs" placeholder="Histórico, inadimplência..." {...p}/></div>
      </G2>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><BP onClick={save} c="linear-gradient(135deg,#8b5cf6,#a78bfa)">💾 Salvar</BP>{editing&&<BD onClick={()=>{if(confirm("Excluir?"))onDelete(editing);setShowForm(false);}}>🗑</BD>}<BG onClick={()=>setShowForm(false)}>Cancelar</BG></div>
    </Card>}
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      {filtered.length===0?<div style={{textAlign:"center",color:"#64748b",padding:"50px 0"}}>Nenhum cliente</div>
      :filtered.map(c=>{const dc=dias(c.cnhValidade);const loc=getLoc(c.id);return <div key={c.id} onClick={()=>openEdit(c)} style={{...cSty,padding:"14px 18px",cursor:"pointer",transition:"transform .12s"}} onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"} onMouseLeave={e=>e.currentTarget.style.transform=""}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div><div style={{fontWeight:700,fontSize:15}}>{c.nome}</div><div style={{display:"flex",gap:10,marginTop:4,flexWrap:"wrap",fontSize:12,color:"#64748b"}}>{c.telefone&&<span>📱 {c.telefone}</span>}{c.cpf&&<span>🪪 {c.cpf}</span>}</div></div>
          <div style={{textAlign:"right",flexShrink:0}}><div style={{fontFamily:"'DM Mono',monospace",fontSize:13,color:"#a78bfa",fontWeight:700}}>{loc} locação{loc!==1?"ões":""}</div>{dc!==null&&dc<=60&&<div style={{fontSize:11,color:dc<=0?"#ef4444":"#fbbf24",fontWeight:700,marginTop:4}}>CNH {dc<=0?"VENCIDA":"vence "+dc+"d"}</div>}</div>
        </div>
        {c.obs&&<div style={{fontSize:12,color:"#94a3b8",marginTop:6,borderTop:"1px solid #1e2130",paddingTop:6}}>📝 {c.obs}</div>}
      </div>;})}
    </div>
  </div>;
}

// ─── DOCUMENTOS ───────────────────────────────────────────────────────────────
const emptyD={vehicleId:"",tipo:"",vencimento:"",numero:"",valor:"",obs:""};
function Documentos({vehicles,docs,onSave,onDelete,showToast}){
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [filterV,setFilterV]=useState("todos");
  const [f,setF]=useState({...emptyD});
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const p={f,set};
  const save=()=>{if(!f.vehicleId||!f.tipo){showToast("Veículo e tipo obrigatórios","#ef4444");return;}onSave({...f,id:editing||Date.now()});setShowForm(false);showToast("📄 Documento salvo!");};
  const filtered=filterV==="todos"?docs:docs.filter(d=>String(d.vehicleId)===filterV);
  const vencendo=docs.filter(d=>{const dv=dias(d.vencimento);return dv!==null&&dv<=30;});
  const getV=id=>{const v=vehicles.find(v=>String(v.id)===String(id));return v?`${v.placa} – ${v.modelo}`:"—";};
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
      <div><h1 style={{fontSize:20,fontWeight:800}}>📄 Documentos</h1><p style={{color:"#64748b",fontSize:13,marginTop:3}}>{docs.length} registros</p></div>
      <BP onClick={()=>{setF({...emptyD});setEditing(null);setShowForm(true);}} c="linear-gradient(135deg,#fbbf24,#f59e0b)">+ Adicionar</BP>
    </div>
    {vencendo.length>0&&<div style={{background:"#1a1108",border:"1px solid #3d2e00",borderRadius:14,padding:16,marginBottom:14}}>
      <div style={{fontWeight:700,color:"#fbbf24",marginBottom:10,fontSize:14}}>⚠️ Vencendo em breve ({vencendo.length})</div>
      {vencendo.map(d=>{const dv=dias(d.vencimento);return <div key={d.id} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:"1px solid #2d2000",fontSize:13}}><span>{d.tipo} — {getV(d.vehicleId)}</span><span style={{color:dv<=0?"#ef4444":"#f97316",fontWeight:700}}>{dv<=0?"VENCIDO":dv+"d"}</span></div>;})}
    </div>}
    <select value={filterV} onChange={e=>setFilterV(e.target.value)} style={{...iSty,width:"auto",minWidth:220,marginBottom:14,cursor:"pointer"}}><option value="todos">Todos</option>{vehicles.map(v=><option key={v.id} value={String(v.id)}>{v.placa} – {v.modelo}</option>)}</select>
    {showForm&&<Card style={{borderColor:"#3d2e00",marginBottom:16}}>
      <ST color="#fbbf24">{editing?"✏️ Editar":"➕ Novo Documento"}</ST>
      <G2 style={{marginBottom:20}}>
        <FS label="Veículo *" id="vehicleId" {...p}><option value="">Selecione...</option>{vehicles.map(v=><option key={v.id} value={String(v.id)}>{v.placa} – {v.modelo}</option>)}</FS>
        <FS label="Tipo *" id="tipo" {...p}><option value="">Selecione...</option>{DOC_TIPOS.map(t=><option key={t} value={t}>{t}</option>)}</FS>
        <FI label="Vencimento" id="vencimento" type="date" {...p}/><FI label="Número" id="numero" {...p}/>
        <FM label="Valor" id="valor" {...p}/><FI label="Observações" id="obs" {...p}/>
      </G2>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><BP onClick={save} c="linear-gradient(135deg,#fbbf24,#f59e0b)">💾 Salvar</BP>{editing&&<BD onClick={()=>{if(confirm("Excluir?"))onDelete(editing);setShowForm(false);}}>🗑</BD>}<BG onClick={()=>setShowForm(false)}>Cancelar</BG></div>
    </Card>}
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      {filtered.length===0?<div style={{textAlign:"center",color:"#64748b",padding:"50px 0"}}>Nenhum documento</div>
      :[...filtered].sort((a,b)=>(a.vencimento||"").localeCompare(b.vencimento||"")).map(d=>{
        const dv=dias(d.vencimento);const ac=dv===null?null:dv<=0?"#ef4444":dv<=30?"#f97316":dv<=60?"#fbbf24":null;
        return <div key={d.id} onClick={()=>{setF({...d});setEditing(d.id);setShowForm(true);}} style={{...cSty,padding:"14px 18px",cursor:"pointer",transition:"transform .12s",borderColor:ac?ac+"33":"#1e2130"}} onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"} onMouseLeave={e=>e.currentTarget.style.transform=""}>
          <div style={{display:"flex",justifyContent:"space-between"}}>
            <div><div style={{fontWeight:700,fontSize:14}}>{d.tipo}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#94a3b8",marginTop:2}}>{getV(d.vehicleId)}</div></div>
            <div style={{textAlign:"right"}}>{d.valor&&<div style={{fontFamily:"'DM Mono',monospace",fontWeight:800,fontSize:14,color:"#fbbf24"}}>{fmt(d.valor)}</div>}{d.vencimento&&<div style={{fontSize:12,color:ac||"#64748b",fontWeight:ac?700:400,marginTop:2}}>{ac&&(dv<=0?"⛔ ":"⚠️ ")}{ptDate(d.vencimento)}</div>}</div>
          </div>
        </div>;
      })}
    </div>
  </div>;
}

// ─── FINANCEIRO ───────────────────────────────────────────────────────────────
function Financeiro({vehicles,manutencoes,contratos}){
  const semPot=vehicles.filter(v=>v.status!=="inativo").reduce((s,v)=>s+fmtN(v.semanal),0);
  const semAtual=vehicles.filter(v=>v.status==="alugado").reduce((s,v)=>s+fmtN(v.semanal),0);
  const totalManut=manutencoes.reduce((s,m)=>s+fmtN(m.custo),0);
  const totalContr=contratos.reduce((s,c)=>s+fmtN(c.valorTotal),0);
  const lucro=totalContr-totalManut;
  return <div>
    <div style={{marginBottom:22}}><h1 style={{fontSize:20,fontWeight:800}}>💰 Financeiro</h1></div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
      {[["Receita Semanal Potencial",fmt(semPot),"#22c55e","#1e3a2e"],["Receita Semanal Atual",fmt(semAtual),"#60a5fa","#1e2a3d"],["Receita Total (Contratos)",fmt(totalContr),"#a5b4fc","#1e1e3d"],["Total em Manutenção",fmt(totalManut),"#fb923c","#2d1e0f"]].map(([l,v,c,b])=>(
        <Card key={l} style={{borderColor:b}}><div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:6}}>{l}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:18,fontWeight:800,color:c}}>{v}</div></Card>
      ))}
    </div>
    <Card style={{borderColor:lucro>=0?"#1e3a2e":"#3d1515",marginBottom:14}}>
      <div style={{fontSize:10,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:6}}>Lucro Estimado (Contratos − Manutenção)</div>
      <div style={{fontFamily:"'DM Mono',monospace",fontSize:24,fontWeight:800,color:lucro>=0?"#22c55e":"#ef4444"}}>{fmt(Math.abs(lucro))}</div>
      <div style={{fontSize:12,color:"#64748b",marginTop:4}}>{lucro>=0?"✅ Positivo":"❌ Negativo"}</div>
    </Card>
    {vehicles.map(v=>{
      const vM=manutencoes.filter(m=>String(m.vehicleId)===String(v.id)).reduce((s,m)=>s+fmtN(m.custo),0);
      const vC=contratos.filter(c=>String(c.vehicleId)===String(v.id)).reduce((s,c)=>s+fmtN(c.valorTotal),0);
      return <Card key={v.id} style={{marginBottom:10}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:10,height:10,borderRadius:"50%",background:STATUS_MAP[v.status]?.dot}}/><div><div style={{fontWeight:700,fontSize:14}}>{v.modelo}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#94a3b8"}}>{v.placa}</div></div></div>
          <StatusTag status={v.status}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
          {[["SEMANAL",v.semanal,"#22c55e"],["MENSAL",v.mensal,"#60a5fa"],["CONTRATOS",vC||null,"#a5b4fc"],["MANUTENÇÃO",vM||null,"#fb923c"]].map(([l,val,c])=>(
            <div key={l} style={{background:"#0f1117",borderRadius:8,padding:"10px 8px",textAlign:"center"}}><div style={{fontSize:9,color:"#64748b",marginBottom:4,fontWeight:700}}>{l}</div><div style={{fontFamily:"'DM Mono',monospace",fontSize:12,fontWeight:800,color:c}}>{val?fmt(val):"—"}</div></div>
          ))}
        </div>
      </Card>;
    })}
  </div>;
}

// ─── EXPORT EXCEL ─────────────────────────────────────────────────────────────
function exportAll(vehicles,manutencoes,clientes,contratos,docs,reservas){
  const wb=XLSX.utils.book_new();
  const ws=(data,cols)=>{const s=XLSX.utils.aoa_to_sheet(data);s["!cols"]=cols.map(w=>({wch:w}));return s;};
  const getV=id=>{const v=vehicles.find(v=>String(v.id)===String(id));return v?`${v.placa} – ${v.modelo}`:""};
  const getC=id=>{const c=clientes.find(c=>String(c.id)===String(id));return c?c.nome:""};
  XLSX.utils.book_append_sheet(wb,ws([["Placa","Modelo","Ano","Cor","KM","Revisão","IPVA","Seguro","Status","Locatário","Diária","Semanal","Mensal","Caução"],...vehicles.map(v=>[v.placa,v.modelo,v.ano,v.cor,v.km?Number(v.km):"",ptDate(v.proxRevisao),ptDate(v.ipva),ptDate(v.seguro),STATUS_MAP[v.status]?.label,v.locatario,v.diaria?Number(v.diaria):"",v.semanal?Number(v.semanal):"",v.mensal?Number(v.mensal):"",v.caucao?Number(v.caucao):""]) ],[12,16,6,10,12,14,12,12,14,18,10,10,10,10]),"Frota");
  XLSX.utils.book_append_sheet(wb,ws([["Veículo","Cliente","Início","Fim","Valor","Pagamento","KM Saída","KM Retorno","Status","Assinatura"],...contratos.map(c=>[getV(c.vehicleId),getC(c.clienteId),ptDate(c.dataInicio),ptDate(c.dataFim),c.valorTotal?Number(c.valorTotal):"",c.formaPgto,c.kmSaida?Number(c.kmSaida):"",c.kmRetorno?Number(c.kmRetorno):"",c.status,c.assinatura?"Sim":"Não"])],[22,18,12,12,14,14,12,12,12,10]),"Contratos");
  XLSX.utils.book_append_sheet(wb,ws([["Data","Veículo","Tipo","Descrição","KM","Oficina","Custo"],...manutencoes.map(m=>[ptDate(m.data),getV(m.vehicleId),m.tipo,m.descricao,m.km?Number(m.km):"",m.oficina,m.custo?Number(m.custo):""]) ],[12,22,18,28,12,18,12]),"Manutenções");
  XLSX.utils.book_append_sheet(wb,ws([["Nome","CPF","Telefone","E-mail","CNH","Val.CNH","Endereço"],...clientes.map(c=>[c.nome,c.cpf,c.telefone,c.email,c.cnh,ptDate(c.cnhValidade),c.endereco])],[22,16,14,24,14,14,28]),"Clientes");
  XLSX.utils.book_append_sheet(wb,ws([["Veículo","Tipo","Vencimento","Número","Valor"],...docs.map(d=>[getV(d.vehicleId),d.tipo,ptDate(d.vencimento),d.numero,d.valor?Number(d.valor):""]) ],[22,16,14,14,12]),"Documentos");
  XLSX.utils.book_append_sheet(wb,ws([["Cliente","Telefone","Veículo","Início","Fim","Status","Obs"],...reservas.map(r=>[r.clienteNome,r.clienteTel,getV(r.vehicleId),ptDate(r.dataInicio),ptDate(r.dataFim),r.status,r.obs])],[20,14,22,12,12,12,20]),"Reservas");
  XLSX.writeFile(wb,`FrotaManager_${new Date().toLocaleDateString("pt-BR").replace(/\//g,"-")}.xlsx`);
}

// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App(){
  const [vehicles,setVehicles]   = useState(()=>ls.get(KEYS.vehicles,FLEET_INIT));
  const [manutencoes,setManut]   = useState(()=>ls.get(KEYS.manut));
  const [clientes,setClientes]   = useState(()=>ls.get(KEYS.clientes));
  const [contratos,setContratos] = useState(()=>ls.get(KEYS.contratos));
  const [docs,setDocs]           = useState(()=>ls.get(KEYS.docs));
  const [fotos,setFotos]         = useState(()=>ls.get(KEYS.fotos));
  const [reservas,setReservas]   = useState(()=>ls.get(KEYS.reservas));
  const [section,setSection]     = useState("dashboard");
  const [toast,setToast]         = useState({msg:"",color:""});

  useEffect(()=>ls.set(KEYS.vehicles,vehicles),[vehicles]);
  useEffect(()=>ls.set(KEYS.manut,manutencoes),[manutencoes]);
  useEffect(()=>ls.set(KEYS.clientes,clientes),[clientes]);
  useEffect(()=>ls.set(KEYS.contratos,contratos),[contratos]);
  useEffect(()=>ls.set(KEYS.docs,docs),[docs]);
  useEffect(()=>ls.set(KEYS.fotos,fotos),[fotos]);
  useEffect(()=>ls.set(KEYS.reservas,reservas),[reservas]);

  const showToast=(msg,color="#22c55e")=>{setToast({msg,color});setTimeout(()=>setToast({msg:"",color:""}),2400);};
  const goTo=s=>setSection(s);

  const crudSave=(setter)=>item=>setter(p=>p.find(x=>x.id===item.id)?p.map(x=>x.id===item.id?item:x):[...p,item]);
  const crudDel=(setter)=>id=>setter(p=>p.filter(x=>x.id!==id));

  const handleFoto=foto=>{
    if(foto._delete) setFotos(p=>p.filter(f=>f.id!==foto.id));
    else setFotos(p=>p.find(f=>f.id===foto.id)?p.map(f=>f.id===foto.id?foto:f):[...p,foto]);
  };

  const NAV=[
    {id:"dashboard",label:"📊 Painel"},{id:"lista",label:"🚘 Frota"},{id:"contratos",label:"📋 Contratos"},
    {id:"manutencoes",label:"🔧 Manutenções"},{id:"clientes",label:"👥 Clientes"},{id:"documentos",label:"📄 Docs"},
    {id:"financeiro",label:"💰 Financeiro"},{id:"calendario",label:"📅 Calendário"},{id:"reservas",label:"📲 Reservas"},
  ];

  const pendentes=reservas.filter(r=>r.status==="pendente").length;

  return <div style={{fontFamily:"'DM Sans','Segoe UI',sans-serif",background:"#0f1117",color:"#e2e8f0",minHeight:"100vh"}}>
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500;700&display=swap');
      *{box-sizing:border-box;margin:0;padding:0}
      ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:#1a1d27}::-webkit-scrollbar-thumb{background:#2d3148;border-radius:3px}
      @keyframes fadeUp{from{opacity:0;transform:translate(-50%,10px)}to{opacity:1;transform:translate(-50%,0)}}
      select option{background:#1a1d27;color:#e2e8f0}
    `}</style>

    {/* Header */}
    <div style={{background:"#13161f",borderBottom:"1px solid #1e2130",padding:"0 12px",display:"flex",alignItems:"center",justifyContent:"space-between",height:54,position:"sticky",top:0,zIndex:100}}>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:32,height:32,background:"linear-gradient(135deg,#6366f1,#8b5cf6)",borderRadius:9,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>🚗</div>
        <div style={{fontWeight:800,fontSize:14}}>FrotaManager</div>
      </div>
      <div style={{display:"flex",gap:1,alignItems:"center",flexWrap:"wrap"}}>
        {NAV.map(n=>(
          <button key={n.id} onClick={()=>goTo(n.id)} style={{background:section===n.id?"#1e2130":"transparent",color:section===n.id?"#a5b4fc":"#64748b",border:"none",padding:"5px 7px",borderRadius:7,fontWeight:700,fontSize:11,cursor:"pointer",position:"relative"}}>
            {n.label}
            {n.id==="reservas"&&pendentes>0&&<span style={{position:"absolute",top:1,right:1,width:7,height:7,borderRadius:"50%",background:"#f97316"}}/>}
          </button>
        ))}
        <button onClick={()=>exportAll(vehicles,manutencoes,clientes,contratos,docs,reservas)} style={{background:"rgba(34,197,94,.12)",color:"#22c55e",border:"1px solid rgba(34,197,94,.25)",padding:"5px 9px",borderRadius:7,fontWeight:700,fontSize:11,cursor:"pointer",marginLeft:4}}>📥 Excel</button>
      </div>
    </div>

    <div style={{maxWidth:920,margin:"0 auto",padding:"20px 14px"}}>
      {section==="dashboard"   && <Dashboard vehicles={vehicles} manutencoes={manutencoes} contratos={contratos} reservas={reservas} onNav={goTo}/>}
      {section==="lista"       && <Frota vehicles={vehicles} fotos={fotos} onSave={crudSave(setVehicles)} onDelete={crudDel(setVehicles)} onFoto={handleFoto}/>}
      {section==="contratos"   && <Contratos vehicles={vehicles} clientes={clientes} contratos={contratos} fotos={fotos} onSave={crudSave(setContratos)} onDelete={crudDel(setContratos)} onFoto={handleFoto} showToast={showToast}/>}
      {section==="manutencoes" && <Manutencoes vehicles={vehicles} manutencoes={manutencoes} onSave={crudSave(setManut)} onDelete={crudDel(setManut)} showToast={showToast}/>}
      {section==="clientes"    && <Clientes clientes={clientes} contratos={contratos} onSave={crudSave(setClientes)} onDelete={crudDel(setClientes)} showToast={showToast}/>}
      {section==="documentos"  && <Documentos vehicles={vehicles} docs={docs} onSave={crudSave(setDocs)} onDelete={crudDel(setDocs)} showToast={showToast}/>}
      {section==="financeiro"  && <Financeiro vehicles={vehicles} manutencoes={manutencoes} contratos={contratos}/>}
      {section==="calendario"  && <Calendario vehicles={vehicles} contratos={contratos} reservas={reservas} onNav={goTo}/>}
      {section==="reservas"    && <Reservas vehicles={vehicles} reservas={reservas} clientes={clientes} onSave={crudSave(setReservas)} onDelete={crudDel(setReservas)} showToast={showToast}/>}
    </div>

    <Toast msg={toast.msg} color={toast.color}/>
  </div>;
}
