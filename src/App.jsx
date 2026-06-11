import { useState, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

const D = {
  bg:"#0D1117", sidebar:"#161B22", card:"#1C2333", border:"#30363D", borderSub:"#21262D",
  primary:"#00C896", blue:"#3B82F6", purple:"#8B5CF6", orange:"#F59E0B", red:"#EF4444",
  text:"#C9D1D9", textSub:"#8B949E", textDim:"#484F58", white:"#FFFFFF",
};

const CLUSTER_COLORS=[
  {color:"#00C896",dim:"#00C89618"},{color:"#3B82F6",dim:"#3B82F618"},
  {color:"#8B5CF6",dim:"#8B5CF618"},{color:"#F59E0B",dim:"#F59E0B18"},
  {color:"#EF4444",dim:"#EF444418"},{color:"#06B6D4",dim:"#06B6D418"},
];

const DEFAULT_CLUSTERS=[
  {nome:"Conservador",   desc:"Baixo risco, renda estável, pouco engajamento digital. Prefere produtos seguros como CDB e poupança."},
  {nome:"Investidor",    desc:"Alta renda, tolerância média a risco, investe em fundos e ações. Busca rentabilidade."},
  {nome:"Premium",       desc:"Alto saldo, muito engajado, usa múltiplos produtos. Cliente de alto valor estratégico."},
  {nome:"Digital",       desc:"Jovem, alta frequência no app, renda menor mas crescente. Engajamento digital elevado."},
];

// ── DEMO DATA ────────────────────────────────────────────────
const DEMO_ROWS = Array.from({length:40}, (_,i) => {
  const perfis = ["Conservador","Moderado","Arrojado","Agressivo"];
  const cidades = ["São Paulo","Rio de Janeiro","Belo Horizonte","Curitiba","Porto Alegre","Brasília"];
  const nomes = ["Ana Silva","Carlos Souza","Maria Santos","João Oliveira","Fernanda Lima","Ricardo Alves","Beatriz Costa","Paulo Ferreira","Camila Rocha","Diego Martins","Larissa Cardoso","Rafael Mendes","Juliana Neves","Bruno Castro","Tatiana Pires","Marcos Vieira","Priscila Gomes","Leonardo Reis","Amanda Faria","Gabriel Lopes","Natalia Cunha","Felipe Barros","Renata Queiroz","Thiago Carvalho","Isabela Ramos","Guilherme Dias","Patricia Moura","Anderson Sousa","Daniela Borges","Rodrigo Lima","Sabrina Teixeira","Vinícius Melo","Caroline Freitas","Eduardo Ribeiro","Aline Pereira","Gustavo Correia","Mônica Araújo","Leandro Xavier","Cristiane Pinto","Matheus Nunes"];
  const isYoung = i%4===0;
  const isPremium = i%4===1;
  const isInvestor = i%4===2;
  const isDigital = i%4===3;
  return {
    nome: nomes[i%nomes.length],
    idade: isYoung?25+Math.floor(Math.random()*8):isPremium?45+Math.floor(Math.random()*15):isInvestor?38+Math.floor(Math.random()*12):22+Math.floor(Math.random()*10),
    genero: i%2===0?"Feminino":"Masculino",
    cidade: cidades[i%cidades.length],
    renda_mensal: isYoung?8000+Math.floor(Math.random()*7000):isPremium?25000+Math.floor(Math.random()*30000):isInvestor?18000+Math.floor(Math.random()*15000):5000+Math.floor(Math.random()*4000),
    saldo_medio: isYoung?15000+Math.floor(Math.random()*20000):isPremium?150000+Math.floor(Math.random()*200000):isInvestor?80000+Math.floor(Math.random()*100000):3000+Math.floor(Math.random()*8000),
    valor_investido: isYoung?5000+Math.floor(Math.random()*15000):isPremium?200000+Math.floor(Math.random()*300000):isInvestor?100000+Math.floor(Math.random()*150000):0+Math.floor(Math.random()*3000),
    perfil_risco: isPremium?"Agressivo":isInvestor?"Moderado":isYoung?"Moderado":"Conservador",
    tempo_relacionamento: isYoung?1+Math.floor(Math.random()*4):isPremium?8+Math.floor(Math.random()*12):isInvestor?5+Math.floor(Math.random()*8):1+Math.floor(Math.random()*3),
    produtos_contratados: isPremium?6+Math.floor(Math.random()*4):isInvestor?4+Math.floor(Math.random()*3):isYoung?2+Math.floor(Math.random()*3):1+Math.floor(Math.random()*2),
    transacoes_mes: isDigital?80+Math.floor(Math.random()*60):isPremium?40+Math.floor(Math.random()*40):isYoung?50+Math.floor(Math.random()*50):20+Math.floor(Math.random()*30),
    uso_app: isDigital?60+Math.floor(Math.random()*40):isYoung?40+Math.floor(Math.random()*30):isPremium?20+Math.floor(Math.random()*30):10+Math.floor(Math.random()*20),
    score_engajamento: isPremium?75+Math.floor(Math.random()*25):isDigital?70+Math.floor(Math.random()*30):isInvestor?50+Math.floor(Math.random()*30):20+Math.floor(Math.random()*30),
    respondeu_campanha_anterior: i%3===0?"Sim":"Não",
  };
});

function buildDemoState() {
  const NUM_COLS=["idade","renda_mensal","saldo_medio","valor_investido","tempo_relacionamento","produtos_contratados","transacoes_mes","uso_app","score_engajamento"];
  const avail=NUM_COLS.filter(c=>DEMO_ROWS[0][c]!==undefined);
  const norm=normalize(DEMO_ROWS,avail);
  const{labels}=kMeans(norm,4);
  const grouped=Array.from({length:4},(_,ci)=>({id:ci,clients:DEMO_ROWS.filter((_,i)=>labels[i]===ci)}));
  return{labels,grouped,cols:avail};
}

const DEMO_PROFILES = [
  {cluster:1,nome:"Conservador",descricao:"Clientes de baixo risco com renda estável. Preferem produtos seguros como CDB e poupança. Baixo engajamento digital mas alta fidelidade.",tom:"formal",mensagem:"Olá! Temos uma novidade especial em produtos de renda fixa com rentabilidade acima do CDI. Que tal uma conversa para conhecer as opções ideais para o seu perfil?"},
  {cluster:2,nome:"Investidor",descricao:"Alta renda com tolerância moderada a risco. Investe em fundos e ações. Busca rentabilidade e diversificação de portfólio.",tom:"técnico",mensagem:"Boa tarde! Identificamos oportunidades em fundos multimercado alinhados ao seu perfil de investidor. Posso compartilhar uma análise personalizada?"},
  {cluster:3,nome:"Premium",descricao:"Alto patrimônio e muito engajado. Usa múltiplos produtos e é cliente estratégico de alto valor. Merece atendimento exclusivo.",tom:"formal",mensagem:"Prezado cliente, como parte do nosso programa Premium, temos acesso antecipado a uma nova linha de investimentos exclusivos. Gostaria de apresentar as condições especiais?"},
  {cluster:4,nome:"Digital",descricao:"Jovem com alta frequência no app e renda crescente. Perfil digital nativo com alto potencial de crescimento.",tom:"informal",mensagem:"Ei! Vi que você é um super usuário do nosso app. Temos uma oferta especial de cashback para você aproveitar essa semana. Bora ver?"},
];

const DEMO_MESSAGES = DEMO_PROFILES.flatMap(p => {
  const g_clients = DEMO_ROWS.filter((_,i) => i%4 === p.cluster-1).slice(0,3);
  return g_clients.map(client => ({
    cluster:p.cluster, nome:client.nome, cidade:client.cidade,
    mensagem:p.mensagem.replace(/(cliente|Prezado cliente)/gi, client.nome.split(" ")[0]),
    tom:p.tom, perfil:p.nome,
  }));
});



const GOOGLE=`@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;800;900&family=Poppins:wght@400;500;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
::-webkit-scrollbar{width:4px;}
::-webkit-scrollbar-thumb{background:#30363D;border-radius:4px;}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
@keyframes glow{0%,100%{box-shadow:0 0 8px #00C89640}50%{box-shadow:0 0 20px #00C896aa}}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
`;
const BODY="'Poppins',sans-serif";
const TITLE="'Montserrat',sans-serif";

function kMeans(data,k,maxIter=100){
  let centroids=data.slice(0,k).map(p=>[...p]),labels=new Array(data.length).fill(0);
  for(let iter=0;iter<maxIter;iter++){
    const next=data.map(pt=>{let best=0,min=Infinity;centroids.forEach((c,ci)=>{const d=pt.reduce((s,v,i)=>s+(v-c[i])**2,0);if(d<min){min=d;best=ci;}});return best;});
    if(next.every((l,i)=>l===labels[i]))break;
    labels=next;
    centroids=Array.from({length:k},(_,ci)=>{const pts=data.filter((_,i)=>labels[i]===ci);if(!pts.length)return centroids[ci];return Array.from({length:data[0].length},(_,d)=>pts.reduce((s,p)=>s+p[d],0)/pts.length);});
  }
  return{labels,centroids};
}
function normalize(rows,cols){
  const mins=cols.map(c=>Math.min(...rows.map(r=>+r[c]||0)));
  const maxs=cols.map(c=>Math.max(...rows.map(r=>+r[c]||0)));
  return rows.map(r=>cols.map((c,i)=>maxs[i]===mins[i]?0:((+r[c]||0)-mins[i])/(maxs[i]-mins[i])));
}
function parseCsv(text){
  const lines=text.trim().split("\n"),headers=lines[0].split(",").map(h=>h.replace(/^\uFEFF/,"").trim());
  return lines.slice(1).map(line=>{const vals=line.split(","),obj={};headers.forEach((h,i)=>(obj[h]=(vals[i]||"").trim()));return obj;});
}
function exportCSV(data,fn){const keys=Object.keys(data[0]),csv=[keys.join(","),...data.map(r=>keys.map(k=>`"${(r[k]||"").toString().replace(/"/g,'""')}"`).join(","))].join("\n");Object.assign(document.createElement("a"),{href:URL.createObjectURL(new Blob([csv],{type:"text/csv"})),download:fn}).click();}
function exportJSON(data,fn){Object.assign(document.createElement("a"),{href:URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:"application/json"})),download:fn}).click();}
function exportXLSX(data,fn){const ws=XLSX.utils.json_to_sheet(data),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Dados");XLSX.writeFile(wb,fn);}

function useIsMobile(){
  const[m,setM]=useState(window.innerWidth<768);
  useEffect(()=>{const h=()=>setM(window.innerWidth<768);window.addEventListener("resize",h);return()=>window.removeEventListener("resize",h);},[]);
  return m;
}

// ── ATOMS ────────────────────────────────────────────────────
function Card({children,style,glow}){
  return <div style={{background:D.card,border:`1px solid ${glow||D.border}`,borderRadius:14,padding:20,boxShadow:glow?`0 0 20px ${glow}25`:"0 2px 8px rgba(0,0,0,.3)",...style}}>{children}</div>;
}
function Btn({children,onClick,disabled,color,outline,style}){
  const bg=color||D.primary;
  if(outline)return <button onClick={onClick} disabled={disabled} style={{background:"transparent",color:bg,border:`1.5px solid ${bg}`,borderRadius:10,padding:"10px 18px",cursor:disabled?"not-allowed":"pointer",fontFamily:TITLE,fontWeight:700,fontSize:12,textTransform:"uppercase",letterSpacing:".06em",opacity:disabled?.4:1,transition:"all .2s",...style}}>{children}</button>;
  return <button onClick={onClick} disabled={disabled} style={{background:disabled?D.border:bg,color:disabled?D.textDim:"#000",border:"none",borderRadius:10,padding:"12px 20px",cursor:disabled?"not-allowed":"pointer",fontFamily:TITLE,fontWeight:700,fontSize:12,textTransform:"uppercase",letterSpacing:".06em",boxShadow:disabled?"none":`0 4px 14px ${bg}40`,transition:"all .2s",...style}}>{children}</button>;
}
function ExportBar({data,basename,label}){
  if(!data?.length)return null;
  return <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:20,alignItems:"center"}}>
    <span style={{fontFamily:BODY,fontSize:11,color:D.textSub}}>Exportar {label}:</span>
    {[{l:"CSV",f:()=>exportCSV(data,`${basename}.csv`),c:D.primary},{l:"JSON",f:()=>exportJSON(data,`${basename}.json`),c:D.blue},{l:"XLSX",f:()=>exportXLSX(data,`${basename}.xlsx`),c:D.purple}].map(b=>(
      <button key={b.l} onClick={b.f} style={{fontFamily:TITLE,fontWeight:700,fontSize:11,padding:"5px 14px",borderRadius:8,border:`1px solid ${b.c}50`,background:`${b.c}15`,color:b.c,cursor:"pointer"}}>{b.l}</button>
    ))}
  </div>;
}
function Sparkline({values,color}){
  const w=100,h=32;
  if(!values||values.length<2)return null;
  const min=Math.min(...values),max=Math.max(...values);
  const norm=values.map(v=>max===min?.5:(v-min)/(max-min));
  const pts=norm.map((v,i)=>`${(i/(values.length-1))*w},${h-(v*(h-4)+2)}`).join(" ");
  return <svg width={w} height={h}><polyline points={pts} fill="none" stroke={color} strokeWidth="2" strokeLinecap="round"/><circle cx={(norm.length-1)/(norm.length-1)*w} cy={h-(norm[norm.length-1]*(h-4)+2)} r="3" fill={color}/></svg>;
}

// ── CLUSTER LABEL ────────────────────────────────────────────
function ClusterDot({color,size=10}){
  return <div style={{width:size,height:size,borderRadius:"50%",background:color,boxShadow:`0 0 ${size}px ${color}`,flexShrink:0}}/>;
}

// ── SIDEBAR ──────────────────────────────────────────────────
function Sidebar({tab,onChange,isMobile,open,onClose,aiReady}){
  if(isMobile&&!open)return null;
  const items=[
    {id:"importar", label:"Importar",  dot:"1"},
    {id:"dashboard",label:"Dashboard", dot:"2"},
    {id:"clientes", label:"Clientes",  dot:"3"},
    {id:"clusters", label:"Clusters",  dot:"4"},
    {id:"mensagens",label:aiReady?"AI WhatsApp":"Mensagens", pulse:aiReady},
  ];
  return <>
    {isMobile&&<div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:40}}/>}
    <div style={{position:isMobile?"fixed":"relative",top:0,left:0,bottom:0,zIndex:50,width:220,minWidth:220,background:D.sidebar,borderRight:`1px solid ${D.border}`,display:"flex",flexDirection:"column",height:isMobile?"100vh":"auto",overflowY:"auto"}}>
      <div style={{padding:"22px 20px 18px",borderBottom:`1px solid ${D.borderSub}`}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:38,height:38,borderRadius:10,background:D.primary,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:TITLE,fontWeight:900,fontSize:20,color:"#000"}}>C</div>
          <div>
            <div style={{fontFamily:TITLE,fontWeight:900,fontSize:15,color:D.white,letterSpacing:"-.3px"}}>Cluster<span style={{color:D.primary}}>CRM</span></div>
            <div style={{fontFamily:BODY,fontSize:10,color:D.textSub,marginTop:1}}>AI · K-Means · WhatsApp</div>
          </div>
        </div>
      </div>
      <nav style={{padding:"14px 10px",flex:1}}>
        {items.map(item=>{
          const active=tab===item.id;
          return <button key={item.id} onClick={()=>{onChange(item.id);if(isMobile)onClose();}}
            style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:10,marginBottom:2,border:"none",
              background:active?`${D.primary}18`:"transparent",
              color:active?D.primary:item.pulse?D.primary:D.textSub,
              cursor:"pointer",fontFamily:BODY,fontWeight:active||item.pulse?600:400,fontSize:13,
              textAlign:"left",borderLeft:active?`2px solid ${D.primary}`:"2px solid transparent",
              transition:"all .15s",animation:item.pulse&&!active?"pulse 1.5s ease-in-out infinite":undefined}}>
            <div style={{width:22,height:22,borderRadius:6,background:active?`${D.primary}30`:D.bg,border:`1px solid ${active?D.primary:D.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:TITLE,fontWeight:800,fontSize:10,color:active?D.primary:D.textSub,flexShrink:0}}>
              {item.dot||"★"}
            </div>
            <span style={{flex:1}}>{item.label}</span>
            {item.pulse&&!active&&<div style={{width:7,height:7,borderRadius:"50%",background:D.primary,animation:"pulse 1s ease-in-out infinite"}}/>}
          </button>;
        })}
      </nav>
      <div style={{padding:"14px 20px",borderTop:`1px solid ${D.borderSub}`}}>
        <div style={{fontFamily:BODY,fontSize:10,color:D.textSub}}>Dados apenas no browser</div>
      </div>
    </div>
  </>;
}

function MobileTopbar({onMenu,aiReady}){
  return <div style={{height:54,background:D.sidebar,borderBottom:`1px solid ${D.border}`,display:"flex",alignItems:"center",padding:"0 16px",gap:12,position:"sticky",top:0,zIndex:30}}>
    <button onClick={onMenu} style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:8,padding:"6px 10px",color:D.textSub,cursor:"pointer",fontSize:14}}>|||</button>
    <div style={{fontFamily:TITLE,fontWeight:900,fontSize:15,color:D.white}}>Cluster<span style={{color:D.primary}}>CRM</span></div>
    {aiReady&&<div style={{marginLeft:"auto",background:`${D.primary}20`,border:`1px solid ${D.primary}`,borderRadius:20,padding:"3px 10px",fontFamily:TITLE,fontWeight:700,fontSize:10,color:D.primary,animation:"pulse 1.5s ease-in-out infinite"}}>AI WPP</div>}
  </div>;
}

const sT=(e={})=>({fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",letterSpacing:".1em",marginBottom:12,...e});
const h1S=(e={})=>({fontFamily:TITLE,fontWeight:900,color:D.white,...e});
const h2S=(e={})=>({fontFamily:TITLE,fontWeight:800,color:D.white,...e});
const bodyS=(e={})=>({fontFamily:BODY,fontSize:13,color:D.white,...e});
const mutedS=(e={})=>({fontFamily:BODY,fontSize:12,color:D.textSub,...e});

// ── CLUSTER EDITOR ───────────────────────────────────────────
function ClusterEditor({clusters,onChange,k}){
  const[editing,setEditing]=useState(null);
  const update=(i,field,val)=>{const n=[...clusters];n[i]={...n[i],[field]:val};onChange(n);};
  const add=()=>{if(clusters.length>=6)return;onChange([...clusters,{nome:`Cluster ${clusters.length+1}`,desc:"Descreva este perfil de cliente"}]);};
  const remove=(i)=>{if(clusters.length<=2)return;onChange(clusters.filter((_,idx)=>idx!==i));};
  const reset=()=>{onChange(DEFAULT_CLUSTERS.slice(0,k).map(c=>({...c})));setEditing(null);};

  return <div>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
      <div style={sT({marginBottom:0})}>PERFIS DOS CLUSTERS</div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        <Btn onClick={reset} outline color={D.textSub} style={{padding:"6px 12px",fontSize:11}}>RESTAURAR PADRÃO</Btn>
        {clusters.length<6&&<Btn onClick={add} outline color={D.primary} style={{padding:"6px 12px",fontSize:11}}>+ ADICIONAR</Btn>}
      </div>
    </div>
    <p style={mutedS({marginBottom:14,fontSize:11})}>
      Estes são os perfis sugeridos para seus clusters. A IA usará estes nomes e descrições ao gerar os perfis e mensagens. Clique em qualquer card para editar.
    </p>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:10}}>
      {clusters.map((cn,i)=>{
        const cl=CLUSTER_COLORS[i%6];
        const isEditing=editing===i;
        return <div key={i} onClick={()=>!isEditing&&setEditing(i)}
          style={{background:D.bg,borderRadius:12,padding:14,border:`1.5px solid ${isEditing?cl.color:cl.color+"40"}`,cursor:isEditing?"default":"pointer",transition:"all .2s",animation:isEditing?"none":undefined}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <ClusterDot color={cl.color} size={8}/>
            <span style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:cl.color,textTransform:"uppercase",flex:1}}>
              Cluster {i+1}
            </span>
            {clusters.length>2&&<button onClick={e=>{e.stopPropagation();remove(i);}} style={{background:"transparent",border:"none",color:D.textDim,cursor:"pointer",fontSize:14,padding:"0 2px",lineHeight:1}}>×</button>}
          </div>
          {isEditing?(
            <div onClick={e=>e.stopPropagation()}>
              <input value={cn.nome} onChange={e=>update(i,"nome",e.target.value)}
                style={{width:"100%",background:D.card,border:`1px solid ${cl.color}60`,borderRadius:8,padding:"8px 10px",fontFamily:TITLE,fontWeight:700,fontSize:13,color:cl.color,outline:"none",boxSizing:"border-box",marginBottom:8}}/>
              <textarea value={cn.desc} onChange={e=>update(i,"desc",e.target.value)} rows={3}
                style={{width:"100%",background:D.card,border:`1px solid ${D.border}`,borderRadius:8,padding:"8px 10px",fontFamily:BODY,fontSize:11,color:D.text,outline:"none",resize:"vertical",boxSizing:"border-box",marginBottom:8}}/>
              <Btn onClick={()=>setEditing(null)} color={cl.color} style={{width:"100%",padding:"8px 0",fontSize:11}}>SALVAR</Btn>
            </div>
          ):(
            <>
              <div style={{fontFamily:TITLE,fontWeight:700,fontSize:13,color:cl.color,marginBottom:6}}>{cn.nome}</div>
              <div style={mutedS({fontSize:11,lineHeight:1.6})}>{cn.desc}</div>
              <div style={{marginTop:8,fontFamily:BODY,fontSize:10,color:D.textDim}}>Clique para editar</div>
            </>
          )}
        </div>;
      })}
    </div>
  </div>;
}

// ── APP ──────────────────────────────────────────────────────
export default function App(){
  const isMobile=useIsMobile();
  const[onboarding,setOnboarding]=useState(true);
  const[sidebarOpen,setSidebarOpen]=useState(false);
  const[tab,setTab]=useState("dashboard");
  const[rows,setRows]=useState(DEMO_ROWS);
  const[fileName,setFileName]=useState("demo_clientes_40.csv");
  const[k,setK]=useState(4);
  const[clusters,setClusters]=useState(()=>buildDemoState());
  const[profiles,setProfiles]=useState(DEMO_PROFILES);
  const[messages,setMessages]=useState(DEMO_MESSAGES);
  const[apiKey,setApiKey]=useState("");
  const[showKey,setShowKey]=useState(false);
  const[loading,setLoading]=useState(false);
  const[error,setError]=useState("");
  const[activeCluster,setActiveCluster]=useState(0);
  const[copied,setCopied]=useState(null);
  const[dragOver,setDragOver]=useState(false);
  const[clusterDefs,setClusterDefs]=useState(DEFAULT_CLUSTERS.map(c=>({...c})));

  const aiReady=messages.length>0;
  const isDemo=fileName==="demo_clientes_40.csv";
  const NUM_COLS=["idade","renda_mensal","saldo_medio","valor_investido","tempo_relacionamento","produtos_contratados","transacoes_mes","uso_app","score_engajamento"];

  const processFile=file=>{
    setFileName(file.name);
    const ext=file.name.split(".").pop().toLowerCase();
    if(ext==="xlsx"||ext==="xls"){
      const r=new FileReader();
      r.onload=e=>{const wb=XLSX.read(e.target.result,{type:"array"});setRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));setClusters(null);setProfiles([]);setMessages([]);};
      r.readAsArrayBuffer(file);
    }else{
      const r=new FileReader();
      r.onload=e=>{setRows(parseCsv(e.target.result));setClusters(null);setProfiles([]);setMessages([]);};
      r.readAsText(file);
    }
  };

  const runClustering=()=>{
    if(!rows.length)return;
    const avail=NUM_COLS.filter(c=>rows[0][c]!==undefined);
    if(!avail.length){setError("Nenhuma coluna numérica compatível encontrada no CSV.");return;}
    const norm=normalize(rows,avail);
    const kActual=Math.min(k,clusterDefs.length);
    const{labels,centroids}=kMeans(norm,kActual);
    setClusters({labels,grouped:Array.from({length:kActual},(_,ci)=>({id:ci,clients:rows.filter((_,i)=>labels[i]===ci)})),cols:avail});
    setProfiles([]);setMessages([]);setTab("clusters");
  };

  const analyzeWithAI=useCallback(async()=>{
    if(!clusters||!apiKey.trim())return;
    setLoading(true);setError("");
    try{
      const avg=(g,f)=>{const v=g.clients.map(c=>+c[f]||0).filter(x=>x>0);return v.length?(v.reduce((a,b)=>a+b,0)/v.length).toFixed(0):"N/A";};
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
        body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:2000,
          system:`CRM bancário. APENAS JSON válido, sem markdown.\nFormato: [{"cluster":1,"nome":"string","descricao":"string","tom":"formal|informal|técnico","mensagem":"WhatsApp 2-3 linhas com emoji"}]`,
          messages:[{role:"user",content:`Analise e gere perfis + mensagens WhatsApp para estes clusters. Use os nomes sugeridos como referência:\n${clusters.grouped.map((g,i)=>`Cluster ${g.id+1} — Perfil sugerido: "${clusterDefs[i]?.nome||"Cluster "+(i+1)}" (${clusterDefs[i]?.desc||""})\nDados reais: ${g.clients.length} clientes | Idade média: ${avg(g,"idade")} anos | Renda média: R$${avg(g,"renda_mensal")} | Saldo médio: R$${avg(g,"saldo_medio")} | Engajamento: ${avg(g,"score_engajamento")}`).join("\n\n")}`}]}),
      });
      const data=await res.json();
      if(data.error)throw new Error(data.error.message);
      const parsed=JSON.parse(data.content.map(c=>c.text||"").join("").replace(/```json|```/g,"").trim());
      setProfiles(parsed);
      const msgs=[];
      parsed.forEach(p=>{
        const g=clusters.grouped[p.cluster-1];
        if(!g)return;
        g.clients.slice(0,3).forEach(client=>{msgs.push({cluster:p.cluster,nome:client.nome||"Cliente",cidade:client.cidade||"",mensagem:p.mensagem.replace(/\b(Ana|Cliente)\b/gi,(client.nome||"").split(" ")[0]||"Cliente"),tom:p.tom,perfil:p.nome,descricao:p.descricao});});
      });
      setMessages(msgs);setTab("mensagens");
    }catch(e){setError(e.message);}
    finally{setLoading(false);}
  },[clusters,apiKey,clusterDefs]);

  const copyMsg=(text,idx)=>{navigator.clipboard.writeText(text);setCopied(idx);setTimeout(()=>setCopied(null),2000);};
  const avf=(g,f)=>{const v=g.clients.map(r=>+r[f]||0).filter(x=>x>0);return v.length?(v.reduce((a,b)=>a+b,0)/v.length).toFixed(0):"N/A";};
  const clientesExport=clusters?rows.map((r,i)=>({...r,cluster:clusters.labels[i]+1})):rows;
  const clustersExport=clusters?clusters.grouped.flatMap(g=>{const p=profiles[g.id];return g.clients.map(c=>({...c,cluster:g.id+1,perfil:p?.nome||clusterDefs[g.id]?.nome||""}));}):[];
  const mensagensExport=messages.map(m=>({cluster:m.cluster,perfil:m.perfil,nome:m.nome,cidade:m.cidade,tom:m.tom,mensagem:m.mensagem}));
  const P=isMobile?16:28;

  // ── ONBOARDING ──────────────────────────────────────────────
  if(onboarding)return(
    <div style={{minHeight:"100vh",background:D.bg,display:"flex",alignItems:"center",justifyContent:"center",padding:24,fontFamily:BODY}}>
      <style>{GOOGLE}</style>
      <div style={{width:"100%",maxWidth:580,animation:"fadeIn .5s ease"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{width:72,height:72,borderRadius:20,background:`${D.primary}20`,border:`2px solid ${D.primary}50`,margin:"0 auto 16px",display:"flex",alignItems:"center",justifyContent:"center"}}>
            <div style={{fontFamily:TITLE,fontWeight:900,fontSize:32,color:D.primary}}>C</div>
          </div>
          <div style={{fontFamily:TITLE,fontWeight:900,fontSize:isMobile?32:44,color:D.white,letterSpacing:"-2px",lineHeight:1}}>
            Cluster<span style={{color:D.primary}}>CRM</span>
          </div>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.white,marginTop:10,letterSpacing:".08em"}}>
            SEGMENTACAO K-MEANS · PERFIS COM IA · MENSAGENS WHATSAPP
          </div>
        </div>

        <Card style={{marginBottom:14,padding:22}} glow={D.primary}>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.primary,textTransform:"uppercase",letterSpacing:".1em",marginBottom:12}}>O QUE O CLUSTERCRM FAZ</div>
          <p style={{fontFamily:BODY,fontSize:13,color:D.white,lineHeight:1.85,marginBottom:18}}>
            Carregue uma base de clientes em <span style={{color:D.primary,fontWeight:600}}>CSV ou Excel</span>. O app usa <span style={{color:D.blue,fontWeight:600}}>K-Means</span> para agrupar por comportamento financeiro. A <span style={{color:D.purple,fontWeight:600}}>Claude API</span> analisa e gera mensagens prontas para o <span style={{color:D.primary,fontWeight:600}}>WhatsApp</span>.
          </p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
            {[["1","Importe","CSV ou Excel"],[D.blue.replace("#",""),"2","Segmente","K-Means"],["3","Analise","IA + perfis"],["4","Dispare","WhatsApp"]].map(([n,title,desc],idx)=>{
              const colors=[D.primary,D.blue,D.purple,D.orange];
              return <div key={n+title} style={{background:D.bg,borderRadius:10,padding:"12px 8px",textAlign:"center",border:`1px solid ${colors[idx]}30`}}>
                <div style={{fontFamily:TITLE,fontWeight:900,fontSize:20,color:colors[idx],marginBottom:4}}>{idx+1}</div>
                <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.white,textTransform:"uppercase",marginBottom:3}}>{title}</div>
                <div style={{fontFamily:BODY,fontSize:10,color:D.textSub}}>{desc}</div>
              </div>;
            })}
          </div>
        </Card>

        <Card style={{marginBottom:20,padding:18}}>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.primary,textTransform:"uppercase",letterSpacing:".1em",marginBottom:12}}>O QUE VOCE PRECISA</div>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:8}}>
            {[["✓",D.primary,"CSV ou Excel com dados dos clientes"],["✓",D.primary,"Colunas numericas: renda, saldo, engajamento"],["✓",D.blue,"API Key da Anthropic (gratis para testar)"],["✓",D.textSub,"Coluna nome dos clientes (opcional)"]].map(([icon,color,text])=>(
              <div key={text} style={{display:"flex",gap:10,alignItems:"flex-start"}}>
                <span style={{color,fontFamily:TITLE,fontWeight:900,fontSize:16,flexShrink:0,lineHeight:1.4}}>{icon}</span>
                <span style={{fontFamily:BODY,fontSize:12,color:D.white,lineHeight:1.6}}>{text}</span>
              </div>
            ))}
          </div>
        </Card>

        <Btn onClick={()=>setOnboarding(false)} style={{width:"100%",padding:"15px 0",fontSize:13,borderRadius:14,letterSpacing:".1em"}}>COMECAR AGORA</Btn>
        <div style={{fontFamily:BODY,fontSize:10,color:D.textSub,textAlign:"center",marginTop:10}}>Dados ficam apenas no browser</div>
      </div>
    </div>
  );

  // ── IMPORTAR ────────────────────────────────────────────────
  const PageImportar=()=>(
    <div style={{padding:P,animation:"fadeIn .3s ease"}}>
      <div style={{marginBottom:22}}>
        <div style={h1S({fontSize:isMobile?22:26})}>Importar Base</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.textSub,textTransform:"uppercase",letterSpacing:".08em",marginTop:5}}>SIGA OS 4 PASSOS PARA SEGMENTAR E GERAR MENSAGENS</div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:isMobile?"repeat(2,1fr)":"repeat(4,1fr)",gap:8,marginBottom:24}}>
        {[[D.primary,"1","IMPORTE","CSV ou Excel"],[D.blue,"2","SEGMENTE","K-Means"],[D.purple,"3","API KEY","Anthropic"],[D.primary,"4","ANALISE","IA + WhatsApp"]].map(([color,n,title,desc])=>(
          <Card key={title} style={{padding:"14px 12px",borderColor:`${color}30`,textAlign:"center"}}>
            <div style={{fontFamily:TITLE,fontWeight:900,fontSize:24,color,marginBottom:6}}>{n}</div>
            <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.white,textTransform:"uppercase",marginBottom:3}}>{title}</div>
            <div style={{fontFamily:BODY,fontSize:11,color:D.textSub}}>{desc}</div>
          </Card>
        ))}
      </div>

      {/* Arquivo */}
      <Card style={{marginBottom:14,padding:0,overflow:"hidden"}}>
        <div style={{padding:"12px 18px",borderBottom:`1px solid ${D.border}`}}>
          <div style={sT()}>PASSO 1 — ARQUIVO DE CLIENTES</div>
        </div>
        <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
          onDrop={e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)processFile(f);}}
          onClick={()=>document.getElementById("fi").click()}
          style={{margin:14,border:`2px dashed ${dragOver?D.primary:D.border}`,borderRadius:10,padding:isMobile?"22px 16px":"30px 24px",textAlign:"center",cursor:"pointer",background:dragOver?`${D.primary}08`:D.bg,transition:"all .2s"}}>
          <input id="fi" type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&processFile(e.target.files[0])}/>
          <div style={{fontFamily:TITLE,fontWeight:900,fontSize:28,color:fileName?D.primary:D.textDim,marginBottom:10}}>{fileName?"✓":"+"}</div>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.white,textTransform:"uppercase",marginBottom:8}}>{fileName||"ARRASTE OU CLIQUE PARA SELECIONAR"}</div>
          <div style={{display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
            {[[".CSV",D.primary],[".XLSX",D.blue],[".XLS",D.purple]].map(([l,c])=>(
              <span key={l} style={{background:`${c}15`,color:c,border:`1px solid ${c}40`,padding:"3px 10px",borderRadius:6,fontFamily:TITLE,fontWeight:700,fontSize:10}}>{l}</span>
            ))}
          </div>
        </div>
        {rows.length>0&&<div style={{margin:"0 14px 14px",background:`${D.primary}12`,border:`1px solid ${D.primary}40`,borderRadius:8,padding:10,fontFamily:BODY,fontSize:12,color:D.primary,display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontFamily:TITLE,fontWeight:900}}>✓</span> <strong>{rows.length} clientes</strong> carregados — {Object.keys(rows[0]).length} colunas
        </div>}
      </Card>

      {/* Segmentos + Editor */}
      <Card style={{marginBottom:14}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16,flexWrap:"wrap"}}>
          <div style={sT({marginBottom:0})}>PASSO 2 — NUMERO DE SEGMENTOS (K)</div>
        </div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:6}}>
          {[2,3,4,5,6].map(n=>(
            <button key={n} onClick={()=>{setK(n);if(clusterDefs.length<n){const extra=Array.from({length:n-clusterDefs.length},(_,i)=>({nome:`Cluster ${clusterDefs.length+i+1}`,desc:"Descreva este perfil"}));setClusterDefs([...clusterDefs,...extra]);}}}
              style={{width:48,height:48,borderRadius:10,border:`2px solid ${k===n?D.primary:D.border}`,background:k===n?`${D.primary}20`:D.bg,color:k===n?D.primary:D.textSub,cursor:"pointer",fontFamily:TITLE,fontWeight:900,fontSize:18,boxShadow:k===n?`0 0 12px ${D.primary}40`:"none",transition:"all .2s"}}>{n}</button>
          ))}
        </div>
        <div style={mutedS({fontSize:11,marginBottom:20})}>Recomendado: 3–4 para bases de 100–500 clientes</div>
        <ClusterEditor clusters={clusterDefs.slice(0,k)} onChange={defs=>{const full=[...defs,...clusterDefs.slice(defs.length)];setClusterDefs(full);}} k={k}/>
      </Card>

      {/* API Key */}
      <Card style={{marginBottom:16}}>
        <div style={sT()}>PASSO 3 — ANTHROPIC API KEY</div>
        <div style={{position:"relative"}}>
          <input type={showKey?"text":"password"} value={apiKey} onChange={e=>setApiKey(e.target.value)}
            placeholder="sk-ant-..."
            style={{width:"100%",padding:"11px 46px 11px 14px",background:D.bg,border:`1px solid ${D.border}`,borderRadius:10,fontFamily:BODY,fontSize:13,color:D.white,outline:"none",boxSizing:"border-box"}}/>
          <button onClick={()=>setShowKey(!showKey)} style={{position:"absolute",right:14,top:"50%",transform:"translateY(-50%)",background:"transparent",border:"none",cursor:"pointer",fontSize:16,color:D.textSub}}>{showKey?"×":"○"}</button>
        </div>
        <div style={mutedS({marginTop:6,fontSize:11})}>Fica apenas no browser. <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noreferrer" style={{color:D.blue}}>Obter chave</a></div>
      </Card>

      <div style={{display:"flex",gap:10,flexDirection:isMobile?"column":"row"}}>
        <Btn onClick={runClustering} disabled={!rows.length} style={{flex:1,padding:"13px 0"}}>RODAR K-MEANS</Btn>
        <Btn onClick={analyzeWithAI} disabled={!clusters||!apiKey.trim()||loading} color={D.blue} style={{flex:1,padding:"13px 0",color:"#fff"}}>
          {loading?"ANALISANDO...":"ANALISAR COM IA"}
        </Btn>
      </div>

      {aiReady&&<div onClick={()=>setTab("mensagens")} style={{marginTop:14,background:`${D.primary}12`,border:`1px solid ${D.primary}`,borderRadius:12,padding:"14px 18px",cursor:"pointer",display:"flex",alignItems:"center",gap:12,animation:"glow 2s ease-in-out infinite"}}>
        <div style={{fontFamily:TITLE,fontWeight:900,fontSize:22,color:D.primary}}>→</div>
        <div>
          <div style={{fontFamily:TITLE,fontWeight:700,fontSize:12,color:D.primary,textTransform:"uppercase"}}>MENSAGENS PRONTAS!</div>
          <div style={mutedS({fontSize:11,marginTop:2})}>Clique para ver as mensagens WhatsApp geradas pela IA</div>
        </div>
      </div>}
      {error&&<div style={{marginTop:12,background:`${D.red}15`,border:`1px solid ${D.red}40`,borderRadius:10,padding:12,fontFamily:BODY,fontSize:12,color:D.red}}>ERRO: {error}</div>}
    </div>
  );

  // ── DASHBOARD ───────────────────────────────────────────────
  const PageDashboard=()=>(
    <div style={{padding:P,animation:"fadeIn .3s ease"}}>
      <div style={{marginBottom:22}}>
        <div style={h1S({fontSize:isMobile?22:26})}>Dashboard</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.textSub,textTransform:"uppercase",letterSpacing:".08em",marginTop:5}}>VISAO CONSOLIDADA DA BASE SEGMENTADA</div>
      </div>
      <ExportBar data={clientesExport.length?clientesExport:null} basename="dashboard" label="clientes"/>
      <div style={{display:"grid",gridTemplateColumns:`repeat(${isMobile?2:4},1fr)`,gap:12,marginBottom:24}}>
        {[{label:"CLIENTES",value:rows.length||0,color:D.primary,spark:[10,14,12,18,15,20,rows.length||0]},
          {label:"CLUSTERS",value:clusters?.grouped.length||0,color:D.blue,spark:[0,1,2,3,clusters?.grouped.length||0]},
          {label:"MENSAGENS",value:messages.length,color:D.purple,spark:[0,2,4,6,8,messages.length]},
          {label:"IA STATUS",value:profiles.length?"OK":"—",color:profiles.length?D.primary:D.textDim,spark:null}].map(kp=>(
          <Card key={kp.label} glow={kp.color} style={{padding:"14px 16px"}}>
            <div style={{fontFamily:TITLE,fontWeight:700,fontSize:10,color:kp.color,textTransform:"uppercase",letterSpacing:".1em",marginBottom:8}}>{kp.label}</div>
            <div style={{fontFamily:TITLE,fontWeight:900,fontSize:26,color:kp.color,marginBottom:8}}>{kp.value}</div>
            {kp.spark&&<Sparkline values={kp.spark} color={kp.color}/>}
          </Card>
        ))}
      </div>
      {clusters?(
        <>
          <div style={sT({marginBottom:12})}>SEGMENTOS</div>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(2,1fr)",gap:12}}>
            {clusters.grouped.map((g,i)=>{
              const cl=CLUSTER_COLORS[i%6],p=profiles[i],cn=clusterDefs[i];
              return <Card key={i} glow={cl.color} style={{padding:16}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                  <ClusterDot color={cl.color}/>
                  <span style={{fontFamily:TITLE,fontWeight:800,fontSize:11,color:cl.color,textTransform:"uppercase"}}>CLUSTER {i+1}</span>
                  <span style={mutedS({marginLeft:"auto",fontSize:11})}>{g.clients.length} clientes</span>
                </div>
                <div style={h2S({fontSize:14,color:cl.color,marginBottom:5})}>{p?p.nome:cn?.nome||`Cluster ${i+1}`}</div>
                <div style={mutedS({lineHeight:1.7,fontSize:12})}>{p?p.descricao?.slice(0,80)+"…":cn?.desc||"Rode a IA para ver o perfil"}</div>
              </Card>;
            })}
          </div>
        </>
      ):(
        <Card style={{textAlign:"center",padding:48}}>
          <div style={{fontFamily:TITLE,fontWeight:900,fontSize:36,color:D.textDim,marginBottom:12}}>+</div>
          <div style={h2S({fontSize:15,marginBottom:8})}>Nenhum dado carregado</div>
          <div style={mutedS({marginBottom:18})}>Comece importando sua base de clientes</div>
          <Btn onClick={()=>setTab("importar")}>IR PARA IMPORTAR</Btn>
        </Card>
      )}
    </div>
  );

  // ── CLIENTES ────────────────────────────────────────────────
  const PageClientes=()=>(
    <div style={{padding:P,animation:"fadeIn .3s ease"}}>
      <div style={{marginBottom:18}}>
        <div style={h1S({fontSize:isMobile?22:26})}>Base de Clientes</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.textSub,textTransform:"uppercase",letterSpacing:".08em",marginTop:5}}>{rows.length} REGISTROS{clusters?" COM CLUSTER ATRIBUIDO":""}</div>
      </div>
      <ExportBar data={clientesExport.length?clientesExport:null} basename="clientes" label="clientes"/>
      {!rows.length?(
        <Card style={{textAlign:"center",padding:48}}>
          <div style={h2S({fontSize:15,marginBottom:8})}>Nenhum cliente carregado</div>
          <Btn onClick={()=>setTab("importar")} style={{marginTop:12}}>IMPORTAR</Btn>
        </Card>
      ):(
        <Card style={{padding:0,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontFamily:BODY,fontSize:12}}>
              <thead><tr style={{background:D.bg}}>
                {["nome","idade","cidade","renda_mensal","perfil_risco","engajamento",...(clusters?["cluster"]:[])].map(h=>(
                  <th key={h} style={{padding:"10px 14px",textAlign:"left",color:D.textSub,fontFamily:TITLE,fontWeight:700,fontSize:10,borderBottom:`1px solid ${D.border}`,whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:".06em"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {rows.slice(0,50).map((r,i)=>(
                  <tr key={i} style={{borderBottom:`1px solid ${D.borderSub}`,background:i%2===0?D.card:D.sidebar}}>
                    {["nome","idade","cidade","renda_mensal","perfil_risco","score_engajamento"].map(h=>(
                      <td key={h} style={{padding:"9px 14px",color:D.text,whiteSpace:"nowrap"}}>{r[h]||"—"}</td>
                    ))}
                    {clusters&&(()=>{const ci=clusters.labels[i],cl=CLUSTER_COLORS[ci%6];return <td style={{padding:"9px 14px"}}><span style={{background:cl.dim,color:cl.color,padding:"2px 8px",borderRadius:20,fontFamily:TITLE,fontWeight:700,fontSize:10,border:`1px solid ${cl.color}40`}}>C{ci+1}</span></td>;})()}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {rows.length>50&&<div style={{padding:"8px 14px",fontFamily:BODY,fontSize:11,color:D.textSub,textAlign:"center",borderTop:`1px solid ${D.border}`}}>Exibindo 50 de {rows.length}</div>}
        </Card>
      )}
    </div>
  );

  // ── CLUSTERS ────────────────────────────────────────────────
  const PageClusters=()=>(
    <div style={{padding:P,animation:"fadeIn .3s ease"}}>
      <div style={{marginBottom:18}}>
        <div style={h1S({fontSize:isMobile?22:26})}>Segmentos</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.textSub,textTransform:"uppercase",letterSpacing:".08em",marginTop:5}}>GRUPOS POR COMPORTAMENTO FINANCEIRO</div>
      </div>
      <ExportBar data={clustersExport.length?clustersExport:null} basename="clusters" label="segmentos"/>
      {!clusters?(
        <Card style={{textAlign:"center",padding:48}}>
          <div style={h2S({fontSize:15,marginBottom:8})}>Nenhum cluster gerado</div>
          <div style={mutedS({marginBottom:18})}>Rode o K-Means na aba Importar</div>
          <Btn onClick={()=>setTab("importar")}>IMPORTAR</Btn>
        </Card>
      ):(
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(2,1fr)",gap:16}}>
          {clusters.grouped.map((g,i)=>{
            const cl=CLUSTER_COLORS[i%6],p=profiles[i],cn=clusterDefs[i];
            return <Card key={i} glow={cl.color} style={{padding:18}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <ClusterDot color={cl.color} size={12}/>
                  <span style={{fontFamily:TITLE,fontWeight:800,fontSize:12,color:cl.color,textTransform:"uppercase"}}>CLUSTER {i+1}</span>
                </div>
                <span style={{background:cl.dim,color:cl.color,padding:"3px 10px",borderRadius:20,fontFamily:TITLE,fontWeight:700,fontSize:10,border:`1px solid ${cl.color}40`}}>{g.clients.length} clientes</span>
              </div>
              <div style={h2S({fontSize:15,color:cl.color,marginBottom:6})}>{p?p.nome:cn?.nome||`Cluster ${i+1}`}</div>
              <div style={mutedS({lineHeight:1.8,marginBottom:14,fontSize:12})}>{p?p.descricao:cn?.desc||"Rode a IA para ver o perfil completo"}</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
                {[["IDADE",avf(g,"idade")+" a"],["RENDA","R$"+Number(avf(g,"renda_mensal")||0).toLocaleString("pt-BR")],["ENGAJ.",avf(g,"score_engajamento")]].map(([l,v])=>(
                  <div key={l} style={{background:D.bg,borderRadius:8,padding:"8px 10px",border:`1px solid ${cl.color}25`}}>
                    <div style={{fontFamily:TITLE,fontWeight:700,fontSize:9,color:cl.color,textTransform:"uppercase",letterSpacing:".06em",marginBottom:3}}>{l}</div>
                    <div style={{fontFamily:TITLE,fontWeight:900,fontSize:14,color:cl.color}}>{v}</div>
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
    <div style={{padding:P,animation:"fadeIn .3s ease"}}>
      <div style={{marginBottom:18}}>
        <div style={h1S({fontSize:isMobile?22:26})}>AI WhatsApp</div>
        <div style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.textSub,textTransform:"uppercase",letterSpacing:".08em",marginTop:5}}>MENSAGENS PERSONALIZADAS POR SEGMENTO</div>
      </div>
      <ExportBar data={mensagensExport.length?mensagensExport:null} basename="mensagens" label="mensagens"/>
      {!messages.length?(
        <Card style={{textAlign:"center",padding:48}}>
          <div style={h2S({fontSize:15,marginBottom:8})}>Nenhuma mensagem gerada</div>
          <div style={mutedS({marginBottom:18})}>Rode a IA na aba Importar</div>
          <Btn onClick={()=>setTab("importar")}>IMPORTAR</Btn>
        </Card>
      ):(
        <>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:18}}>
            {clusters.grouped.map((_,i)=>{
              const cl=CLUSTER_COLORS[i%6],p=profiles[i],cn=clusterDefs[i];
              return <button key={i} onClick={()=>setActiveCluster(i)}
                style={{padding:"7px 14px",borderRadius:20,border:`2px solid ${activeCluster===i?cl.color:D.border}`,background:activeCluster===i?cl.dim:D.card,color:activeCluster===i?cl.color:D.textSub,cursor:"pointer",fontFamily:TITLE,fontWeight:700,fontSize:11,boxShadow:activeCluster===i?`0 0 10px ${cl.color}40`:"none",transition:"all .15s"}}>
                {p?p.nome:cn?.nome||`Cluster ${i+1}`}
              </button>;
            })}
          </div>
          {messages.filter(m=>m.cluster===activeCluster+1).map((m,idx)=>{
            const cl=CLUSTER_COLORS[activeCluster%6];
            return <Card key={idx} style={{marginBottom:12,padding:16}} glow={cl.color}>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,flexWrap:"wrap"}}>
                <div style={{width:36,height:36,borderRadius:10,background:cl.dim,border:`1px solid ${cl.color}40`,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:TITLE,fontWeight:900,fontSize:15,color:cl.color}}>{m.nome[0]}</div>
                <div>
                  <div style={h2S({fontSize:13})}>{m.nome}</div>
                  {m.cidade&&<div style={mutedS({fontSize:11})}>{m.cidade}</div>}
                </div>
                <span style={{marginLeft:"auto",background:cl.dim,color:cl.color,fontSize:10,padding:"2px 8px",borderRadius:10,fontFamily:TITLE,fontWeight:700,border:`1px solid ${cl.color}40`}}>{m.tom}</span>
              </div>
              <div style={{background:"#1A3A2A",border:"1px solid #2D5A3D",borderRadius:"12px 12px 12px 3px",padding:"12px 16px",fontFamily:BODY,fontSize:13,color:"#E8F5E9",lineHeight:1.8,marginBottom:12,maxWidth:"88%"}}>
                {m.mensagem}
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button onClick={()=>copyMsg(m.mensagem,idx)} style={{background:D.bg,color:D.primary,border:`1px solid ${D.primary}40`,borderRadius:8,padding:"8px 16px",fontFamily:TITLE,fontWeight:700,fontSize:11,cursor:"pointer",textTransform:"uppercase"}}>
                  {copied===idx?"COPIADO":"COPIAR"}
                </button>
                <a href={`https://wa.me/?text=${encodeURIComponent(m.mensagem)}`} target="_blank" rel="noreferrer"
                  style={{background:"#25D366",color:"#000",borderRadius:8,padding:"8px 16px",fontFamily:TITLE,fontWeight:700,fontSize:11,textDecoration:"none",display:"flex",alignItems:"center",gap:6,textTransform:"uppercase",boxShadow:"0 4px 12px rgba(37,211,102,.4)"}}>
                  WHATSAPP
                </a>
              </div>
            </Card>;
          })}
        </>
      )}
    </div>
  );

  return(
    <div style={{fontFamily:BODY,minHeight:"100vh",background:D.bg,display:"flex",flexDirection:"column"}}>
      <style>{GOOGLE}</style>
      {isMobile&&<MobileTopbar onMenu={()=>setSidebarOpen(true)} aiReady={aiReady}/>}
      <div style={{display:"flex",flex:1,overflow:"hidden",height:isMobile?"calc(100vh - 54px)":"100vh"}}>
        <Sidebar tab={tab} onChange={setTab} isMobile={isMobile} open={sidebarOpen} onClose={()=>setSidebarOpen(false)} aiReady={aiReady}/>
        <main style={{flex:1,overflowY:"auto",background:D.bg}}>
          {isDemo&&<div style={{background:`${D.orange}18`,borderBottom:`1px solid ${D.orange}40`,padding:"8px 20px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <span style={{fontFamily:TITLE,fontWeight:700,fontSize:11,color:D.orange,textTransform:"uppercase",letterSpacing:".08em"}}>MODO DEMO</span>
            <span style={{fontFamily:BODY,fontSize:11,color:D.textSub}}>Você está vendo dados fictícios de 40 clientes. Importe seu próprio CSV para começar.</span>
            <button onClick={()=>{setRows([]);setFileName("");setClusters(null);setProfiles([]);setMessages([]);setTab("importar");}} style={{marginLeft:"auto",background:"transparent",border:`1px solid ${D.orange}60`,borderRadius:8,padding:"4px 12px",color:D.orange,cursor:"pointer",fontFamily:TITLE,fontWeight:700,fontSize:10,textTransform:"uppercase"}}>USAR MEUS DADOS</button>
          </div>}
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
