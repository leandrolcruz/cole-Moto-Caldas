// Meta da contagem por locação (Leandro, 12/11/26): todos os endereços em 10 semanas.
// Compartilhado pelo dashboard (aba Locação) e pela tela da equipe (meta.html).
// r = {locs:[endereços do universo, em ordem], feitos:[os já contados], cont:{loc:{recebido_em}}, faltam:n}
// Feriados não contam como dia de contagem.
var META_COB={inicio:'2026-11-16',fim:'2027-01-22',feriados:['2026-11-20','2026-12-25','2027-01-01']};
function cobDia(s){var p=s.split('-');return new Date(+p[0],p[1]-1,+p[2])}
function cobIso(d){return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2)}
function cobN(n){return Number(n).toLocaleString('pt-BR')}
function cobDM(d){return d.toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit'})}
// dias úteis (seg a sex, fora feriado) de a até b, inclusive
function cobUteis(a,b){var n=0,d=new Date(a.getFullYear(),a.getMonth(),a.getDate());for(;d<=b;d.setDate(d.getDate()+1)){var w=d.getDay();if(w>0&&w<6&&META_COB.feriados.indexOf(cobIso(d))<0)n++}return n}
// O plano: base = endereços contados ANTES do início; fila = o que faltava no início, na ordem
// dos endereços (é o roteiro). O ritmo é a fila dividida pelos dias úteis do período.
function cobPlano(r){
  var m=META_COB,ini=cobDia(m.inicio),fim=cobDia(m.fim),t0=ini.getTime();
  var fila=r.locs.filter(function(l){var c=r.cont[l];return !c||new Date(c.recebido_em).getTime()>=t0});
  var tot=cobUteis(ini,fim);
  return {ini:ini,fim:fim,fila:fila,base:r.locs.length-fila.length,tot:tot,porDia:tot?fila.length/tot:0};
}
function cobMeta(r){
  var p=cobPlano(r),hoje=new Date();hoje=new Date(hoje.getFullYear(),hoje.getMonth(),hoje.getDate());
  var nT=r.locs.length,nF=r.feitos.length;
  var ontem=new Date(hoje);ontem.setDate(ontem.getDate()-1);
  var passados=hoje>p.ini?cobUteis(p.ini,ontem<p.fim?ontem:p.fim):0;              // dias úteis já fechados
  var esperado=Math.min(nT,p.base+p.porDia*passados);
  var seg=new Date(hoje);seg.setDate(seg.getDate()-((seg.getDay()+6)%7));var sex=new Date(seg);sex.setDate(sex.getDate()+4);
  var a=seg<p.ini?p.ini:seg,b=sex>p.fim?p.fim:sex,nSem=0;
  r.feitos.forEach(function(l){if(new Date(r.cont[l].recebido_em)>=seg)nSem++});
  var restam=hoje>p.fim?0:cobUteis(hoje<p.ini?p.ini:hoje,p.fim);
  return {ini:p.ini,fim:p.fim,antes:hoje<p.ini,depois:hoje>p.fim,porDia:p.porDia,esperado:esperado,saldo:nF-esperado,
    alvoSem:a<=b?Math.round(p.porDia*cobUteis(a,b)):0,nSem:nSem,restam:restam,precisa:restam?r.faltam/restam:0};
}
// A meta fatiada por semana: meta e contado de cada uma, o acumulado contra o plano e o
// roteiro (rota = os endereços da fila que cabem àquela semana).
// Vale a data da contagem vigente de cada endereço (endereço recontado passa para a semana da recontagem).
function cobSemanas(r){
  var p=cobPlano(r),agora=Date.now(),nT=r.locs.length,out=[],ant=p.base;
  var ts=r.feitos.map(function(l){return new Date(r.cont[l].recebido_em).getTime()});
  for(var k=0;;k++){
    var seg=new Date(p.ini);seg.setDate(seg.getDate()+7*k);if(seg>p.fim)break;
    var sex=new Date(seg);sex.setDate(sex.getDate()+4);if(sex>p.fim)sex=p.fim;
    var prox=new Date(seg);prox.setDate(prox.getDate()+7);
    var plano=Math.min(nT,p.base+p.porDia*cobUteis(p.ini,sex)),a=seg.getTime(),b=prox.getTime();
    var rota=p.fila.slice(Math.round(ant)-p.base,Math.round(plano)-p.base);
    out.push({n:k+1,seg:seg,sex:sex,alvo:rota.length,plano:Math.round(plano),rota:rota,
      rotaOk:rota.filter(function(l){return r.cont[l]}).length,
      cont:ts.filter(function(t){return t>=a&&t<b}).length,acum:ts.filter(function(t){return t<b}).length,
      estado:agora>=b?'fim':agora>=a?'atual':'fut'});
    ant=plano;
  }
  return out;
}
// a semana em andamento; antes do início, a 1ª; depois do fim, a última
function cobSemanaAtual(ss){
  for(var i=0;i<ss.length;i++)if(ss[i].estado==='atual')return ss[i];
  return ss.length?(ss[0].estado==='fut'?ss[0]:ss[ss.length-1]):null;
}
// rota agrupada por coluna (3 primeiros caracteres do endereço): [{k,tot,ok,falta:[...]}]
function cobRotaCols(r,rota){
  var cols=[],by={};
  rota.forEach(function(l){var k=l.substring(0,3);if(!by[k]){by[k]={k:k,tot:0,ok:0,falta:[]};cols.push(by[k])}by[k].tot++;if(r.cont[l])by[k].ok++;else by[k].falta.push(l)});
  return cols;
}
function htmlCobSemanas(r){
  var ss=cobSemanas(r),esc2=1;ss.forEach(function(s){esc2=Math.max(esc2,s.alvo,s.cont)});
  var h='<p class="sub" style="margin:10px 0 0">Meta por semana — a linha tracejada é a meta da semana; a barra é o que foi contado. Embaixo, o acumulado contra o plano.</p><div class="sems">';
  ss.forEach(function(s){
    var bateu=s.cont>=s.alvo,cl=s.estado==='fut'?'s-fu':s.estado==='atual'?(bateu?'s-ok':'s-at'):(bateu?'s-ok':'s-nao');
    var dif=s.acum-s.plano;
    h+='<div class="sem '+cl+'"><div class="n">'+(s.estado==='fut'?'—':s.cont)+' / '+s.alvo+'</div>'+
      '<div class="tr"><i style="height:'+(100*s.cont/esc2).toFixed(1)+'%"></i><u style="bottom:'+(100*s.alvo/esc2).toFixed(1)+'%"></u></div>'+
      '<b>Sem. '+s.n+(s.estado==='atual'?' ◀':'')+'</b>'+cobDM(s.seg)+
      '<small>'+(s.estado==='fut'?'plano '+cobN(s.plano):'<span style="color:'+(dif>=0?'var(--gn)':s.estado==='fim'?'var(--dg)':'var(--tx)')+'">'+cobN(s.acum)+'</span> / '+cobN(s.plano))+'</small></div>';
  });
  return h+'</div>';
}
function htmlCobMeta(r){
  var m=cobMeta(r),dm={day:'2-digit',month:'2-digit',year:'2-digit'},f1=function(n){return n.toLocaleString('pt-BR',{maximumFractionDigits:1})};
  if(!r.faltam)return '<div class="nota">🎯 Meta cumprida: todos os endereços contados. 🎉</div>';
  var s=Math.round(m.saldo),st=m.antes?'começa em '+m.ini.toLocaleDateString('pt-BR',dm):m.depois?'prazo vencido':s>=0?'em dia':'atrasado';
  var cor=m.antes?'var(--tx)':(m.depois||s<0)?'var(--dg)':'var(--gn)';
  return '<div class="kpis" style="margin-top:10px">'+
    '<div class="kpi"><b>🎯 '+m.fim.toLocaleDateString('pt-BR',dm)+'</b><span>meta: todos os endereços</span><small>10 semanas, de '+m.ini.toLocaleDateString('pt-BR',dm)+' em diante</small></div>'+
    '<div class="kpi"><b>'+(m.restam?f1(m.precisa):'—')+'</b><span>endereços por dia útil para fechar</span><small>'+(m.restam?'restam '+m.restam+' dias úteis · plano: '+f1(m.porDia)+'/dia':'sem dias úteis no prazo')+'</small></div>'+
    '<div class="kpi"><b>'+m.nSem+(m.alvoSem?' de '+m.alvoSem:'')+'</b><span>nesta semana</span><small>'+(m.alvoSem?'meta da semana (seg a sex)':'fora do período da meta')+'</small></div>'+
    '<div class="kpi"><b style="color:'+cor+'">'+(m.antes?'—':(s>=0?'+':'')+cobN(s))+'</b><span>'+st+'</span><small>'+(m.antes?'ainda não começou':'contados × o plano até ontem ('+cobN(Math.round(m.esperado))+')')+'</small></div></div>'+htmlCobSemanas(r);
}
// Roteiro: que colunas cabem a cada semana, seguindo a ordem dos endereços.
function htmlCobRoteiro(r){
  var ss=cobSemanas(r);
  var h='<p class="sub">Cada semana recebe os próximos endereços da fila, na ordem das estantes. O número ao lado da coluna é quantos endereços dela entram naquela semana.</p>'+
    '<div class="rolar"><table><tr><th>Semana</th><th>Colunas</th><th>Feito</th></tr>';
  ss.forEach(function(s){
    var cols=cobRotaCols(r,s.rota).map(function(c){return '<span class="rc'+(c.ok===c.tot?' ok':'')+'"><b>'+c.k+'</b> '+(c.ok?c.ok+'/':'')+c.tot+'</span>'}).join(' ');
    h+='<tr'+(s.estado==='atual'?' style="background:rgba(72,149,239,.08)"':'')+'><td style="white-space:nowrap"><b>Sem. '+s.n+'</b><br><span class="sub">'+cobDM(s.seg)+' a '+cobDM(s.sex)+'</span></td><td>'+cols+'</td><td style="white-space:nowrap"><b>'+s.rotaOk+'</b> / '+s.alvo+'</td></tr>';
  });
  return h+'</table></div>';
}
