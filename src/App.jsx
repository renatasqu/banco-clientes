import { useState, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

const D = {
  bg:        "#0D1117",
  sidebar:   "#161B22",
  card:      "#1C2333",
  border:    "#30363D",
  borderSub: "#21262D",
  primary:   "#00C896",
  blue:      "#3B82F6",
  purple:    "#8B5CF6",
  orange:    "#F59E0B",
  red:       "#EF4444",
  text:      "#C9D1D9",
  textSub:   "#8B949E",
  textDim:   "#484F58",
  white:     "#FFFFFF",
};

const CLUSTER_COLORS = [
  { color:"#00C896", dim:"#00C89618" },
  { color:"#3B82F6", dim:"#3B82F618" },
  { color:"#8B5CF6", dim:"#8B5CF618" },
  { color:"#F59E0B", dim:"#F59E0B18" },
];

const GOOGLE = `@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;800;900&family=Poppins:wght@400;500;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
::-webkit-scrollbar{width:4px;}
::-webkit-scrollbar-thumb{background:#30363D;border-radius:4px;}
input{color:#fff !important;}
@keyframes navBlink{0%,100%{background:#25D36620;color:#25D366;border-color:#25D36640;box-shadow:0 0 14px #25D36640;}50%{background:#25D36640;color:#fff;border-color:#25D366;box-shadow:0 0 22px #25D366;}}
@keyframes pulse{0%,100%{transform:scale(1);}50%{transform:scale(1.12);}}
`;
const BODY  = "'Poppins', sans-serif";
const TITLE = "'Montserrat', sans-serif";

// ── UTILS ────────────────────────────────────────────────────
function kMeans(data, k, maxIter=100) {
  let centroids = data.slice(0,k).map(p=>[...p]);
  let labels    = new Array(data.length).fill(0);
  for (let iter=0; iter<maxIter; iter++) {
    const next = data.map(pt => {
      let best=0, min=Infinity;
      centroids.forEach((c,ci)=>{ const d=pt.reduce((s,v,i)=>s+(v-c[i])**2,0); if(d<min){min=d;best=ci;} });
      return best;
    });
    if (next.every((l,i)=>l===labels[i])) break;
    labels = next;
    centroids = Array.from({length:k},(_,ci)=>{
      const pts=data.filter((_,i)=>labels[i]===ci);
      if(!pts.length) return centroids[ci];
      return Array.from({length:data[0].length},(_,d)=>pts.reduce((s,p)=>s+p[d],0)/pts.length);
    });
  }
  return {labels,centroids};
}

function normalize(rows,cols) {
  const mins=cols.map(c=>Math.min(...rows.map(r=>+r[c]||0)));
  const maxs=cols.map(c=>Math.max(...rows.map(r=>+r[c]||0)));
  return rows.map(r=>cols.map((c,i)=>maxs[i]===mins[i]?0:((+r[c]||0)-mins[i])/(maxs[i]-mins[i])));
}

function parseCsv(text) {
  const lines=text.trim().split("\n");
  const headers=lines[0].split(",").map(h=>h.replace(/^\uFEFF/,"").trim());
  return lines.slice(1).map(line=>{
    const vals=line.split(","); const obj={};
    headers.forEach((h,i)=>(obj[h]=(vals[i]||"").trim()));
    return obj;
  });
}

function exportCSV(data,filename) {
  const keys=Object.keys(data[0]);
  const csv=[keys.join(","),...data.map(r=>keys.map(k=>`"${(r[k]||"").toString().replace(/"/g,'""')}"`).join(","))].join("\n");
  Object.assign(document.createElement("a"),{href:URL.createObjectURL(new Blob([csv],{type:"text/csv"})),download:filename}).click();
}
function exportJSON(data,filename) {
  Object.assign(document.createElement("a"),{href:URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:"application/json"})),download:filename}).click();
}
function exportXLSX(data,filename) {
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Dados");
  XLSX.writeFile(wb,filename);
}

function useIsMobile() {
  const [m,setM]=useState(window.innerWidth<768);
  useEffect(()=>{ const h=()=>setM(window.innerWidth<768); window.addEventListener("resize",h); return ()=>window.removeEventListener("resize",h); },[]);
  return m;
}

// ── ATOMS ────────────────────────────────────────────────────
function Card({children,style,glow}) {
  return <div style={{background:D.card,border:`1px solid ${glow||D.border}`,borderRadius:14,padding:20,boxShadow:glow?`0 0 20px ${glow}25`:"0 2px 8px rgba(0,0,0,0.3)",...style}}>{children}</div>;
}

function Btn({children,onClick,disabled,color,style}) {
  const bg=color||D.primary;
  return <button onClick={onClick} disabled={disabled} style={{background:disabled?D.border:bg,color:disabled?D.textDim:"#000",border:"none",borderRadius:10,padding:"12px 20px",cursor:disabled?"not-allowed":"pointer",fontFamily:TITLE,fontWeight:700,fontSize:12,textTransform:"uppercase",letterSpacing:".06em",boxShadow:disabled?"none":`0 4px 14px ${bg}40`,transition:"all .2s",...style}}>{children}</button>;
}

function ExportBar({data,basename,label}) {
  if(!data?.length) return null;
  return (
    <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:20,alignItems:"center"}}>
      <span style={{fontFamily:BODY,fontSize:11,color:D.textSub}}>⬇ Exportar {label}:</span>
      {[{l:"CSV",f:()=>exportCSV(data,`${basename}.csv`),c:D.primary},{l:"JSON",f:()=>exportJSON(data,`${basename}.json`),c:D.blue},{l:"XLSX",f:()=>exportXLSX(data,`${basename}.xlsx`),c:D.purple}].map(b=>(
        <button key={b.l} onClick={b.f} style={{fontFamily:TITLE,fontWeight:700,fontSize:11,padding:"5px 14px",borderRadius:8,border:`1px solid ${b.c}50`,background:`${b.c}15`,color:b.c,cursor:"pointer"}}>{b.l}</button>
      ))}
    </div>
  );
}

function Sparkline({values,color}) {
  const w=100,h=36;
  if(!values||values.length<2) return null;
  const min=Math.min(...values),max=Math.max(...values);
  const norm=values.map(v=>max===min?.5:(v-min)/(max-min));
  const pts=norm.map((v,i)=>`${(i/(values.length-1))*w},${h-(v*(h-4)+2)}`).join(" ");
  return <svg width={w} height={h}><polyline points={pts} fill="none" stroke={color} strokeWidth="2" strokeLinecap="round"/><circle cx={(norm.length-1)/(norm.length-1)*w} cy={h-(norm[norm.length-1]*(h-4)+2)} r="3" fill={color}/></svg>;
}

// ── SIDEBAR ──────────────────────────────────────────────────
const NAV=[
  {id:"importar",icon:"📤",label:"Importar"},
  {id:"dashboard",icon:"📊",label:"Dashboard"},
  {id:"clientes",icon:"👥",label:"Clientes"},
  {id:"clusters",icon:"🫧",label:"Clusters"},
  {id:"mensagens",icon:"💬",label:"Mensagens",aiLabel:"AI WHATSAPP",aiIcon:"📱"},
];

function Sidebar({tab,onChange,isMobile,open,onClose,hasMessages}) {
  if(isMobile&&!open) return null;
  return (
    <>
      {isMobile&&<div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:40}}/>}
      <div style={{position:isMobile?"fixed":"relative",top:0,left:0,bottom:0,zIndex:50,width:220,minWidth:220,background:D.sidebar,borderRight:`1px solid ${D.border}`,display:"flex",flexDirection:"column",height:isMobile?"100vh":"auto",overflowY:"auto"}}>
        <div style={{padding:"24px 20px 20px",borderBottom:`1px solid ${D.borderSub}`}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:36,height:36,borderRadius:10,background:D.primary,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,fontWeight:900,color:"#000"}}>C</div>
            <div>
              <div style={{fontFamily:TITLE,fontWeight:900,fontSize:16,color:D.white}}>Cluster<span style={{color:D.primary}}>CRM</span></div>
              <div style={{fontFamily:BODY,fontSize:10,color:D.textSub}}>AI · K-Means · WhatsApp</div>
            </div>
          </div>
        </div>
        <nav style={{padding:"12px 10px",flex:1}}>
          {NAV.map(item=>{
            const active=tab===item.id;
            const isMsg=item.id==="mensagens";
            const showAI=isMsg&&hasMessages;
            const label=showAI?item.aiLabel:item.label;
            const icon=showAI?item.aiIcon:item.icon;
            return <button key={item.id} onClick={()=>{onChange(item.id);if(isMobile)onClose();}} style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:10,marginBottom:2,border:"none",background:active?`${D.primary}18`:"transparent",color:active?D.primary:D.textSub,cursor:"pointer",fontFamily:BODY,fontWeight:active?600:400,fontSize:13,textAlign:"left",borderLeft:active?`2px solid ${D.primary}`:"2px solid transparent",transition:"all .15s",...(showAI&&!active?{animation:"navBlink 1.8s ease-in-out infinite",borderRadius:10,border:`1px solid #25D36640`,background:"#25D36620",color:"#25D366"}:{})}}>
              <span style={{fontSize:16,...(showAI&&!active?{animation:"pulse 1.8s ease-in-out infinite"}:{})}}>{icon}</span>
              <span style={{fontWeight:showAI&&!active?700:undefined}}>{label}</span>
            </button>;
          })}
        </nav>
        <div style={{padding:"16px 20px",borderTop:`1px solid ${D.borderSub}`}}>
          <div style={{fontFamily:BODY,fontSize:10,color:D.textSub}}>🔒 Dados apenas no browser</div>
        </div>
      </div>
    </>
  );
}

function MobileTopbar({onMenu}) {
  return (
    <div style={{height:56,background:D.sidebar,borderBottom:`1px solid ${D.border}`,display:"flex",alignItems:"center",padding:"0 16px",gap:12,position:"sticky",top:0,zIndex:30}}>
      <button onClick={onMenu} style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:8,padding:"6px 10px",color:D.textSub,cursor:"pointer",fontSize:16}}>☰</button>
      <div style={{fontFamily:TITLE,fontWeight:900,fontSize:16,color:D.white}}>Cluster<span style={{color:D.primary}}>CRM</span></div>
    </div>
  );
}

// ── LABEL helpers ─────────────────────────────────────────────
const sectionTitle = (extra={}) => ({fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",letterSpacing:".1em",marginBottom:12,...extra});
const bodyText     = (extra={}) => ({fontFamily:BODY,fontSize:12,color:D.white,...extra});
const mutedText    = (extra={}) => ({fontFamily:BODY,fontSize:12,color:D.textSub,...extra});
const h1style      = (extra={}) => ({fontFamily:TITLE,fontWeight:900,fontSize:28,color:D.white,...extra});
const h2style      = (extra={}) => ({fontFamily:TITLE,fontWeight:800,fontSize:18,color:D.white,...extra});
const h3style      = (extra={}) => ({fontFamily:TITLE,fontWeight:700,fontSize:14,color:D.white,...extra});

// ── SUGGESTED CLUSTER PROFILES ───────────────────────────────
const SUGGESTED_PROFILES=[
  {nome:"Conservador",cor:CLUSTER_COLORS[0].color,emoji:"🛡️",desc:"Baixo risco, renda estável, pouco engajamento digital. Prefere produtos seguros como CDB e poupança."},
  {nome:"Investidor",cor:CLUSTER_COLORS[1].color,emoji:"📈",desc:"Alta renda, tolerância média a risco, investe em fundos e ações. Busca rentabilidade."},
  {nome:"Premium",cor:CLUSTER_COLORS[2].color,emoji:"💎",desc:"Alto saldo, muito engajado, usa múltiplos produtos. Cliente de alto valor estratégico."},
  {nome:"Digital",cor:CLUSTER_COLORS[3].color,emoji:"📱",desc:"Jovem, alta frequência no app, renda menor mas crescente. Engajamento digital elevado."},
];

// ── APP ──────────────────────────────────────────────────────
export default function App() {
  const isMobile=useIsMobile();
  const [onboarding,setOnboarding]=useState(true);
  const [sidebarOpen,setSidebarOpen]=useState(false);
  const [tab,setTab]=useState("importar");
  const [rows,setRows]=useState([]);
  const [fileName,setFileName]=useState("");
  const [k,setK]=useState(4);
  const [clusters,setClusters]=useState(null);
  const [profiles,setProfiles]=useState([]);
  const [messages,setMessages]=useState([]);
  const [apiKey,setApiKey]=useState("");
  const [showKey,setShowKey]=useState(false);
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState("");
  const [activeCluster,setActiveCluster]=useState(0);
  const [copied,setCopied]=useState(null);
  const [dragOver,setDragOver]=useState(false);
  const [clusterNames,setClusterNames]=useState(["Conservador","Investidor","Premium","Digital"]);

  const NUM_COLS=["idade","renda_mensal","saldo_medio","valor_investido","tempo_relacionamento","produtos_contratados","transacoes_mes","uso_app","score_engajamento"];

  const processFile=file=>{
    setFileName(file.name);
    const ext=file.name.split(".").pop().toLowerCase();
    if(ext==="xlsx"||ext==="xls"){
      const r=new FileReader();
      r.onload=e=>{const wb=XLSX.read(e.target.result,{type:"array"});setRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));setClusters(null);setProfiles([]);setMessages([]);};
      r.readAsArrayBuffer(file);
    } else {
      const r=new FileReader();
      r.onload=e=>{setRows(parseCsv(e.target.result));setClusters(null);setProfiles([]);setMessages([]);};
      r.readAsText(file);
    }
  };

  const runClustering=()=>{
    if(!rows.length) return;
    const avail=NUM_COLS.filter(c=>rows[0][c]!==undefined);
    const norm=normalize(rows,avail);
    const {labels,centroids}=kMeans(norm,k);
    setClusters({labels,grouped:Array.from({length:k},(_,ci)=>({id:ci,clients:rows.filter((_,i)=>labels[i]===ci),centroid:centroids[ci]})),cols:avail});
    setProfiles([]);setMessages([]);setTab("clusters");
  };

  const analyzeWithAI=useCallback(async()=>{
    if(!clusters||!apiKey.trim()) return;
    setLoading(true);setError("");
    try {
      const avg=(g,f)=>{const v=g.clients.map(c=>+c[f]||0).filter(x=>x>0);return v.length?(v.reduce((a,b)=>a+b,0)/v.length).toFixed(0):"N/A";};
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
        body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:2000,
          system:`CRM bancário. APENAS JSON válido, sem markdown.\nFormato: [{"cluster":1,"nome":"Perfil","descricao":"2 linhas","tom":"formal|informal|técnico","mensagem":"WhatsApp 2-3 linhas com emoji"}]`,
          messages:[{role:"user",content:`Clusters (use os nomes sugeridos como base):\n${clusters.grouped.map((g,i)=>`Cluster ${g.id+1} (sugestão de nome: "${clusterNames[i]||`Cluster ${i+1}`}"): ${g.clients.length} clientes | Idade:${avg(g,"idade")} | Renda:R$${avg(g,"renda_mensal")} | Saldo:R$${avg(g,"saldo_medio")} | Engaj:${avg(g,"score_engajamento")}`).join("\n")}`}]}),
      });
      const data=await res.json();
      if(data.error) throw new Error(data.error.message);
      const parsed=JSON.parse(data.content.map(c=>c.text||"").join("").replace(/```json|```/g,"").trim());
      setProfiles(parsed);
      const msgs=[];
      parsed.forEach(p=>{clusters.grouped[p.cluster-1].clients.slice(0,3).forEach(client=>{msgs.push({cluster:p.cluster,nome:client.nome||"Cliente",cidade:client.cidade||"",mensagem:p.mensagem.replace(/Ana|cliente/gi,(client.nome||"").split(" ")[0]||"Cliente"),tom:p.tom,perfil:p.nome});});});
      setMessages(msgs);
      setTab("mensagens");
    } catch(e){setError(e.message);}
    finally{setLoading(false);}
  },[clusters,apiKey,clusterNames]);

  const copyMsg=(text,idx)=>{navigator.clipboard.writeText(text);setCopied(idx);setTimeout(()=>setCopied(null),2000);};
  const avf=(g,f)=>{const v=g.clients.map(r=>+r[f]||0).filter(x=>x>0);return v.length?(v.reduce((a,b)=>a+b,0)/v.length).toFixed(0):"N/A";};
  const clientesExport=clusters?rows.map((r,i)=>({...r,cluster:clusters.labels[i]+1})):rows;
  const clustersExport=clusters?clusters.grouped.flatMap(g=>{const p=profiles[g.id];return g.clients.map(c=>({...c,cluster:g.id+1,perfil:p?.nome||""}));}):[];
  const mensagensExport=messages.map(m=>({cluster:m.cluster,perfil:m.perfil,nome:m.nome,cidade:m.cidade,tom:m.tom,mensagem:m.mensagem}));
  const P=isMobile?16:28;

  // ── ONBOARDING ──────────────────────────────────────────────
  if(onboarding) return (
    <div style={{minHeight:"100vh",background:D.bg,display:"flex",alignItems:"center",justifyContent:"center",padding:24,fontFamily:BODY}}>
      <style>{GOOGLE}</style>
      <div style={{width:"100%",maxWidth:600}}>
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{width:72,height:72,borderRadius:18,background:`${D.primary}20`,border:`2px solid ${D.primary}60`,margin:"0 auto 16px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:36}}>🏦</div>
          <div style={{fontFamily:TITLE,fontWeight:900,fontSize:isMobile?34:48,color:D.white,letterSpacing:"-2px",lineHeight:1}}>
            Cluster<span style={{color:D.primary}}>CRM</span>
          </div>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,marginTop:10,letterSpacing:".04em"}}>
            SEGMENTAÇÃO K-MEANS · PERFIS COM IA · MENSAGENS WHATSAPP
          </div>
        </div>

        <Card style={{marginBottom:16,padding:24}} glow={D.primary}>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.primary,textTransform:"uppercase",letterSpacing:".1em",marginBottom:12}}>
            O QUE O CLUSTERCRM FAZ
          </div>
          <p style={{fontFamily:BODY,fontSize:14,color:D.white,lineHeight:1.8,marginBottom:20}}>
            Carregue uma base de clientes em <span style={{color:D.primary,fontWeight:600}}>CSV ou Excel</span>. O app usa <span style={{color:D.blue,fontWeight:600}}>K-Means</span> para agrupar clientes por comportamento financeiro. A <span style={{color:D.purple,fontWeight:600}}>Claude API</span> analisa cada grupo e gera mensagens prontas para o <span style={{color:D.primary,fontWeight:600}}>WhatsApp</span>.
          </p>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)",gap:10}}>
            {[["📂","IMPORTE","CSV ou Excel"],["🔮","SEGMENTE","K-Means automático"],["🤖","ANALISE","IA gera perfis"],["📱","DISPARE","WhatsApp pronto"]].map(([icon,title,desc])=>(
              <div key={title} style={{background:D.bg,borderRadius:10,padding:"14px 10px",textAlign:"center",border:`1px solid ${D.border}`}}>
                <div style={{fontSize:28,marginBottom:8}}>{icon}</div>
                <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",marginBottom:4}}>{title}</div>
                <div style={{fontFamily:BODY,fontSize:12,color:D.white}}>{desc}</div>
              </div>
            ))}
          </div>
        </Card>

        <Card style={{marginBottom:24,padding:20}}>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.primary,textTransform:"uppercase",letterSpacing:".1em",marginBottom:12}}>
            O QUE VOCÊ PRECISA
          </div>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:10}}>
            {[["✅",D.primary,"CSV ou Excel com dados dos clientes"],["✅",D.primary,"Colunas numéricas: renda, saldo, engajamento…"],["🔑",D.blue,"API Key da Anthropic (grátis para testar)"],["💡",D.textSub,"Coluna 'nome' dos clientes (opcional)"]].map(([icon,color,text])=>(
              <div key={text} style={{display:"flex",gap:10,alignItems:"flex-start"}}>
                <span style={{color,flexShrink:0,fontSize:18}}>{icon}</span>
                <span style={{fontFamily:BODY,fontSize:13,color:D.white,lineHeight:1.6}}>{text}</span>
              </div>
            ))}
          </div>
        </Card>

        <Btn onClick={()=>setOnboarding(false)} style={{width:"100%",padding:"16px 0",fontSize:14,borderRadius:14,textTransform:"uppercase",letterSpacing:".08em"}}>
          COMEÇAR AGORA →
        </Btn>
        <div style={{fontFamily:BODY,fontSize:11,color:D.textSub,textAlign:"center",marginTop:12}}>
          🔒 Seus dados ficam apenas no browser
        </div>
      </div>
    </div>
  );

  // ── IMPORTAR ────────────────────────────────────────────────
  const PageImportar=()=>(
    <div style={{padding:P}}>
      <div style={{marginBottom:24}}>
        <div style={h1style({fontSize:isMobile?22:28})}>Importar Base</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",letterSpacing:".08em",marginTop:6}}>
          SIGA OS 4 PASSOS PARA SEGMENTAR E GERAR MENSAGENS
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:isMobile?"repeat(2,1fr)":"repeat(4,1fr)",gap:10,marginBottom:28}}>
        {[[D.primary,"📂","1. IMPORTE","CSV ou Excel"],[D.blue,"🔮","2. SEGMENTE","Rode o K-Means"],[D.purple,"🔑","3. API KEY","Anthropic"],[D.primary,"🤖","4. ANALISE","IA + WhatsApp"]].map(([color,icon,title,desc])=>(
          <Card key={title} style={{padding:"14px 12px",borderColor:`${color}40`,textAlign:"center"}}>
            <div style={{fontSize:24,marginBottom:8}}>{icon}</div>
            <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",marginBottom:4}}>{title}</div>
            <div style={{fontFamily:BODY,fontSize:12,color:D.white}}>{desc}</div>
          </Card>
        ))}
      </div>

      <Card style={{marginBottom:16,padding:0,overflow:"hidden"}}>
        <div style={{padding:"14px 20px",borderBottom:`1px solid ${D.border}`}}>
          <div style={sectionTitle()}>PASSO 1 — ARQUIVO DE CLIENTES</div>
        </div>
        <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
          onDrop={e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)processFile(f);}}
          onClick={()=>document.getElementById("fi").click()}
          style={{margin:16,border:`2px dashed ${dragOver?D.primary:D.border}`,borderRadius:12,padding:isMobile?"24px 16px":"36px 24px",textAlign:"center",cursor:"pointer",background:dragOver?`${D.primary}10`:D.bg,transition:"all .2s"}}>
          <input id="fi" type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&processFile(e.target.files[0])}/>
          <div style={{fontSize:36,marginBottom:10}}>{fileName?"✅":"📂"}</div>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",marginBottom:8}}>
            {fileName||"ARRASTE OU CLIQUE PARA SELECIONAR"}
          </div>
          <div style={{display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
            {[[".CSV",D.primary],[".XLSX",D.blue],[".XLS",D.purple]].map(([l,c])=>(
              <span key={l} style={{background:`${c}18`,color:c,border:`1px solid ${c}40`,padding:"3px 10px",borderRadius:8,fontFamily:TITLE,fontWeight:700,fontSize:11}}>{l}</span>
            ))}
          </div>
        </div>
        {rows.length>0&&(
          <div style={{margin:"0 16px 16px",background:`${D.primary}15`,border:`1px solid ${D.primary}40`,borderRadius:10,padding:12,fontFamily:BODY,fontSize:12,color:D.primary,display:"flex",alignItems:"center",gap:8}}>
            ✅ <strong>{rows.length} clientes</strong> carregados — {Object.keys(rows[0]).length} colunas
          </div>
        )}
      </Card>

      <Card style={{marginBottom:16}}>
        <div style={sectionTitle()}>PASSO 2 — NÚMERO DE SEGMENTOS (K)</div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
          {[2,3,4,5].map(n=>(
            <button key={n} onClick={()=>setK(n)} style={{width:52,height:52,borderRadius:12,border:`2px solid ${k===n?D.primary:D.border}`,background:k===n?`${D.primary}20`:D.bg,color:k===n?D.primary:D.textSub,cursor:"pointer",fontFamily:TITLE,fontWeight:900,fontSize:20,boxShadow:k===n?`0 0 14px ${D.primary}40`:"none",transition:"all .2s"}}>{n}</button>
          ))}
        </div>
        <div style={bodyText()}>Recomendado: 3–4 para bases de 100–500 clientes</div>
      </Card>

      <Card style={{marginBottom:16}}>
        <div style={sectionTitle()}>PERFIS SUGERIDOS — NOMEIE SEUS CLUSTERS</div>
        <div style={{fontFamily:BODY,fontSize:12,color:D.textSub,marginBottom:14}}>Esses são os 4 perfis típicos de clientes bancários. Edite os nomes para personalizar seus clusters antes de rodar a IA.</div>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)",gap:10}}>
          {SUGGESTED_PROFILES.map((p,i)=>{
            const cl=CLUSTER_COLORS[i%4];
            const isActive=k>i;
            return (
              <div key={i} style={{background:isActive?cl.dim:D.bg,border:`1px solid ${isActive?cl.color:D.borderSub}`,borderRadius:12,padding:"12px 12px 14px",opacity:isActive?1:0.4,transition:"all .2s"}}>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
                  <span style={{fontSize:18}}>{p.emoji}</span>
                  <span style={{fontFamily:TITLE,fontWeight:700,fontSize:10,color:cl.color,textTransform:"uppercase"}}>CLUSTER {i+1}</span>
                </div>
                <input
                  value={clusterNames[i]||""}
                  onChange={e=>{const n=[...clusterNames];n[i]=e.target.value;setClusterNames(n);}}
                  disabled={!isActive}
                  placeholder={p.nome}
                  style={{width:"100%",background:"transparent",border:"none",borderBottom:`1px solid ${isActive?cl.color+"50":D.borderSub}`,padding:"4px 0",fontFamily:TITLE,fontWeight:700,fontSize:13,color:isActive?cl.color:D.textDim,outline:"none",marginBottom:8,cursor:isActive?"text":"default"}}
                />
                <div style={{fontFamily:BODY,fontSize:11,color:D.textSub,lineHeight:1.6}}>{p.desc}</div>
              </div>
            );
          })}
        </div>
      </Card>

      <Card style={{marginBottom:20}}>
        <div style={sectionTitle()}>PASSO 3 — ANTHROPIC API KEY</div>
        <div style={{position:"relative"}}>
          <input type={showKey?"text":"password"} value={apiKey} onChange={e=>setApiKey(e.target.value)}
            placeholder="sk-ant-..." style={{width:"100%",padding:"12px 48px 12px 16px",background:D.bg,border:`1px solid ${D.border}`,borderRadius:10,fontFamily:BODY,fontSize:12,outline:"none",boxSizing:"border-box"}}/>
          <button onClick={()=>setShowKey(!showKey)} style={{position:"absolute",right:14,top:"50%",transform:"translateY(-50%)",background:"transparent",border:"none",cursor:"pointer",fontSize:18,color:D.textSub}}>{showKey?"🙈":"👁"}</button>
        </div>
        <div style={{...bodyText(),marginTop:6}}>
          Fica apenas no browser.{" "}
          <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noreferrer" style={{color:D.blue}}>Obter chave →</a>
        </div>
      </Card>

      <div style={{display:"flex",gap:12,flexDirection:isMobile?"column":"row"}}>
        <Btn onClick={runClustering} disabled={!rows.length} style={{flex:1,padding:"14px 0"}}>🔮 RODAR K-MEANS</Btn>
        <Btn onClick={analyzeWithAI} disabled={!clusters||!apiKey.trim()||loading} color={D.blue} style={{flex:1,padding:"14px 0",color:"#fff"}}>
          {loading?"⚙️ ANALISANDO…":"🤖 ANALISAR COM IA"}
        </Btn>
      </div>
      {error&&<div style={{marginTop:12,background:`${D.red}18`,border:`1px solid ${D.red}40`,borderRadius:10,padding:12,fontFamily:BODY,fontSize:12,color:D.red}}>⚠ {error}</div>}
    </div>
  );

  // ── DASHBOARD ───────────────────────────────────────────────
  const PageDashboard=()=>(
    <div style={{padding:P}}>
      <div style={{marginBottom:24}}>
        <div style={h1style({fontSize:isMobile?22:28})}>Dashboard</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",letterSpacing:".08em",marginTop:6}}>VISÃO CONSOLIDADA DA BASE SEGMENTADA</div>
      </div>
      <ExportBar data={clientesExport.length?clientesExport:null} basename="dashboard" label="clientes"/>
      <div style={{display:"grid",gridTemplateColumns:`repeat(${isMobile?2:4},1fr)`,gap:12,marginBottom:24}}>
        {[{label:"CLIENTES",value:rows.length||0,color:D.primary,icon:"👥",spark:[10,14,12,18,15,20,rows.length]},
          {label:"CLUSTERS",value:clusters?.grouped.length||0,color:D.blue,icon:"🫧",spark:[0,1,2,3,4,clusters?.grouped.length||0]},
          {label:"MENSAGENS",value:messages.length,color:D.purple,icon:"💬",spark:[0,2,4,6,8,messages.length]},
          {label:"IA ANALISOU",value:profiles.length?"SIM":"—",color:profiles.length?D.primary:D.textDim,icon:"🤖",spark:null}].map(k=>(
          <Card key={k.label} glow={k.color} style={{padding:"16px 18px"}}>
            <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:k.color,textTransform:"uppercase",letterSpacing:".1em",marginBottom:8}}>{k.icon} {k.label}</div>
            <div style={{fontFamily:TITLE,fontWeight:900,fontSize:28,color:k.color,marginBottom:8}}>{k.value}</div>
            {k.spark&&<Sparkline values={k.spark} color={k.color}/>}
          </Card>
        ))}
      </div>
      {clusters?(
        <>
          <div style={sectionTitle({marginBottom:12})}>SEGMENTOS IDENTIFICADOS</div>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(2,1fr)",gap:14}}>
            {clusters.grouped.map((g,i)=>{
              const cl=CLUSTER_COLORS[i%4];const p=profiles[i];
              return <Card key={i} glow={cl.color} style={{padding:18}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
                  <div style={{width:10,height:10,borderRadius:"50%",background:cl.color,boxShadow:`0 0 8px ${cl.color}`}}/>
                  <span style={{fontFamily:TITLE,fontWeight:800,fontSize:12,color:cl.color}}>CLUSTER {i+1}</span>
                  <span style={{fontFamily:BODY,fontSize:11,color:D.textSub,marginLeft:"auto"}}>{g.clients.length} clientes</span>
                </div>
                <div style={h3style({color:cl.color,marginBottom:6,fontSize:14})}>{p?p.nome:`${g.clients.length} clientes`}</div>
                <div style={mutedText({lineHeight:1.7})}>{p?p.descricao.slice(0,80)+"…":"Rode a IA para ver o perfil"}</div>
              </Card>;
            })}
          </div>
        </>
      ):(
        <Card style={{textAlign:"center",padding:48}}>
          <div style={{fontSize:44,marginBottom:14}}>📤</div>
          <div style={h2style({fontSize:16,marginBottom:8})}>Nenhum dado carregado</div>
          <div style={mutedText({marginBottom:20})}>Comece importando sua base de clientes</div>
          <Btn onClick={()=>setTab("importar")}>IR PARA IMPORTAR</Btn>
        </Card>
      )}
    </div>
  );

  // ── CLIENTES ────────────────────────────────────────────────
  const PageClientes=()=>(
    <div style={{padding:P}}>
      <div style={{marginBottom:20}}>
        <div style={h1style({fontSize:isMobile?22:28})}>Base de Clientes</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",letterSpacing:".08em",marginTop:6}}>
          {rows.length} REGISTROS{clusters?" COM CLUSTER ATRIBUÍDO":""}
        </div>
      </div>
      <ExportBar data={clientesExport.length?clientesExport:null} basename="clientes" label="clientes"/>
      {!rows.length?(
        <Card style={{textAlign:"center",padding:48}}>
          <div style={{fontSize:40}}>👥</div>
          <div style={h2style({fontSize:16,margin:"12px 0 8px"})}>Nenhum cliente carregado</div>
          <Btn onClick={()=>setTab("importar")} style={{marginTop:12}}>IMPORTAR</Btn>
        </Card>
      ):(
        <Card style={{padding:0,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontFamily:BODY,fontSize:12}}>
              <thead>
                <tr style={{background:D.bg}}>
                  {["nome","idade","cidade","renda_mensal","perfil_risco","engajamento",...(clusters?["cluster"]:[])].map(h=>(
                    <th key={h} style={{padding:"12px 16px",textAlign:"left",color:D.textSub,fontFamily:TITLE,fontWeight:700,fontSize:10,borderBottom:`1px solid ${D.border}`,whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:".08em"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.slice(0,50).map((r,i)=>(
                  <tr key={i} style={{borderBottom:`1px solid ${D.borderSub}`,background:i%2===0?D.card:D.sidebar}}>
                    {["nome","idade","cidade","renda_mensal","perfil_risco","score_engajamento"].map(h=>(
                      <td key={h} style={{padding:"10px 16px",color:D.text,whiteSpace:"nowrap"}}>{r[h]||"—"}</td>
                    ))}
                    {clusters&&(()=>{
                      const ci=clusters.labels[i];const cl=CLUSTER_COLORS[ci%4];
                      return <td style={{padding:"10px 16px"}}><span style={{background:cl.dim,color:cl.color,padding:"3px 10px",borderRadius:20,fontFamily:TITLE,fontWeight:700,fontSize:10,border:`1px solid ${cl.color}40`}}>C{ci+1}</span></td>;
                    })()}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {rows.length>50&&<div style={{padding:"10px 16px",fontFamily:BODY,fontSize:11,color:D.textSub,textAlign:"center",borderTop:`1px solid ${D.border}`}}>Exibindo 50 de {rows.length}</div>}
        </Card>
      )}
    </div>
  );

  // ── CLUSTERS ────────────────────────────────────────────────
  const PageClusters=()=>(
    <div style={{padding:P}}>
      <div style={{marginBottom:20}}>
        <div style={h1style({fontSize:isMobile?22:28})}>Segmentos</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",letterSpacing:".08em",marginTop:6}}>GRUPOS POR COMPORTAMENTO FINANCEIRO E ENGAJAMENTO</div>
      </div>
      <ExportBar data={clustersExport.length?clustersExport:null} basename="clusters" label="segmentos"/>
      {!clusters?(
        <Card style={{textAlign:"center",padding:48}}>
          <div style={{fontSize:44}}>🫧</div>
          <div style={h2style({fontSize:16,margin:"12px 0 8px"})}>Nenhum cluster gerado</div>
          <div style={mutedText({marginBottom:20})}>Rode o K-Means na aba Importar</div>
          <Btn onClick={()=>setTab("importar")}>IR PARA IMPORTAR</Btn>
        </Card>
      ):(
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(2,1fr)",gap:18}}>
          {clusters.grouped.map((g,i)=>{
            const cl=CLUSTER_COLORS[i%4];const p=profiles[i];
            return <Card key={i} glow={cl.color} style={{padding:20}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{width:12,height:12,borderRadius:"50%",background:cl.color,boxShadow:`0 0 10px ${cl.color}`}}/>
                  <span style={{fontFamily:TITLE,fontWeight:800,fontSize:13,color:cl.color}}>CLUSTER {i+1}</span>
                </div>
                <span style={{background:cl.dim,color:cl.color,padding:"3px 12px",borderRadius:20,fontFamily:TITLE,fontWeight:700,fontSize:11,border:`1px solid ${cl.color}40`}}>{g.clients.length} clientes</span>
              </div>
              {p&&<div style={h2style({fontSize:16,color:cl.color,marginBottom:8})}>{p.nome}</div>}
              {p&&<div style={mutedText({lineHeight:1.8,marginBottom:16,fontSize:13})}>{p.descricao}</div>}
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
                {[["IDADE",avf(g,"idade")+" a"],["RENDA","R$"+Number(avf(g,"renda_mensal")||0).toLocaleString("pt-BR")],["ENGAJ.",avf(g,"score_engajamento")]].map(([l,v])=>(
                  <div key={l} style={{background:D.bg,borderRadius:10,padding:"10px 12px",border:`1px solid ${cl.color}30`}}>
                    <div style={{fontFamily:TITLE,fontWeight:700,fontSize:9,color:cl.color,textTransform:"uppercase",letterSpacing:".08em",marginBottom:4}}>{l}</div>
                    <div style={{fontFamily:TITLE,fontWeight:900,fontSize:15,color:cl.color}}>{v}</div>
                  </div>
                ))}
              </div>
            </Card>;
          })}
        </div>
      )}
    </div>
  );

  // ── MENSAGENS ───────────────────────────────────────────────
  const PageMensagens=()=>(
    <div style={{padding:P}}>
      <div style={{marginBottom:20}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <div style={h1style({fontSize:isMobile?22:28})}>
            {messages.length>0?"AI WHATSAPP":"Mensagens WhatsApp"}
          </div>
          {messages.length>0&&<span style={{background:"#25D36620",color:"#25D366",border:"1px solid #25D36650",borderRadius:20,padding:"4px 14px",fontFamily:TITLE,fontWeight:700,fontSize:11,letterSpacing:".06em",animation:"navBlink 2s ease-in-out infinite"}}>● ATIVO</span>}
        </div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",letterSpacing:".08em",marginTop:6}}>PERSONALIZADAS POR SEGMENTO, PRONTAS PARA DISPARAR</div>
      </div>
      <ExportBar data={mensagensExport.length?mensagensExport:null} basename="mensagens" label="mensagens"/>
      {!messages.length?(
        <Card style={{textAlign:"center",padding:48}}>
          <div style={{fontSize:44}}>💬</div>
          <div style={h2style({fontSize:16,margin:"12px 0 8px"})}>Nenhuma mensagem gerada</div>
          <div style={mutedText({marginBottom:20})}>Rode a IA na aba Importar</div>
          <Btn onClick={()=>setTab("importar")}>IR PARA IMPORTAR</Btn>
        </Card>
      ):(
        <>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:20}}>
            {clusters.grouped.map((_,i)=>{
              const cl=CLUSTER_COLORS[i%4];const p=profiles[i];
              return <button key={i} onClick={()=>setActiveCluster(i)} style={{padding:"8px 16px",borderRadius:20,border:`2px solid ${activeCluster===i?cl.color:D.border}`,background:activeCluster===i?cl.dim:D.card,color:activeCluster===i?cl.color:D.textSub,cursor:"pointer",fontFamily:TITLE,fontWeight:700,fontSize:12,boxShadow:activeCluster===i?`0 0 12px ${cl.color}40`:"none",transition:"all .15s"}}>
                CLUSTER {i+1}{p?` · ${p.nome.split(" ")[0]}`:""}
              </button>;
            })}
          </div>
          {messages.filter(m=>m.cluster===activeCluster+1).map((m,idx)=>{
            const cl=CLUSTER_COLORS[activeCluster%4];
            return <Card key={idx} style={{marginBottom:14,padding:18}} glow={cl.color}>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
                <div style={{width:38,height:38,borderRadius:10,background:cl.dim,border:`1px solid ${cl.color}40`,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:TITLE,fontWeight:900,fontSize:16,color:cl.color}}>{m.nome[0]}</div>
                <div>
                  <div style={h3style({fontSize:14})}>{m.nome}</div>
                  {m.cidade&&<div style={mutedText()}>{m.cidade}</div>}
                </div>
                <span style={{marginLeft:"auto",background:cl.dim,color:cl.color,fontSize:10,padding:"3px 10px",borderRadius:10,fontFamily:TITLE,fontWeight:700,border:`1px solid ${cl.color}40`}}>{m.tom}</span>
              </div>
              <div style={{background:"#1A3A2A",border:"1px solid #2D5A3D",borderRadius:"14px 14px 14px 4px",padding:"14px 18px",fontFamily:BODY,fontSize:13,color:"#E8F5E9",lineHeight:1.8,marginBottom:14,maxWidth:"88%",boxShadow:"0 2px 8px rgba(0,0,0,0.3)"}}>
                {m.mensagem}
              </div>
              <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                <button onClick={()=>copyMsg(m.mensagem,idx)} style={{background:D.bg,color:D.primary,border:`1px solid ${D.primary}40`,borderRadius:8,padding:"9px 18px",fontFamily:TITLE,fontWeight:700,fontSize:12,cursor:"pointer",textTransform:"uppercase"}}>
                  {copied===idx?"✅ COPIADO!":"📋 COPIAR"}
                </button>
                <a href={`https://wa.me/?text=${encodeURIComponent(m.mensagem)}`} target="_blank" rel="noreferrer" style={{background:"#25D366",color:"#000",borderRadius:8,padding:"9px 18px",fontFamily:TITLE,fontWeight:700,fontSize:12,textDecoration:"none",display:"flex",alignItems:"center",gap:6,textTransform:"uppercase",boxShadow:"0 4px 12px rgba(37,211,102,0.4)"}}>
                  📱 WHATSAPP
                </a>
              </div>
            </Card>;
          })}
        </>
      )}
    </div>
  );

  return (
    <div style={{fontFamily:BODY,minHeight:"100vh",background:D.bg,display:"flex",flexDirection:"column"}}>
      <style>{GOOGLE}</style>
      {isMobile&&<MobileTopbar onMenu={()=>setSidebarOpen(true)}/>}
      <div style={{display:"flex",flex:1,overflow:"hidden",height:isMobile?"calc(100vh - 56px)":"100vh"}}>
        <Sidebar tab={tab} onChange={setTab} isMobile={isMobile} open={sidebarOpen} onClose={()=>setSidebarOpen(false)} hasMessages={messages.length>0}/>
        <main style={{flex:1,overflowY:"auto",background:D.bg}}>
          {tab==="importar"  &&<PageImportar/>}
          {tab==="dashboard" &&<PageDashboard/>}
          {tab==="clientes"  &&<PageClientes/>}
          {tab==="clusters"  &&<PageClusters/>}
          {tab==="mensagens" &&<PageMensagens/>}
        </main>
      </div>
    </div>
  );
}
