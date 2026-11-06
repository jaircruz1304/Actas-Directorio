/* FIAS Directorio · Front-End estático para GitHub Pages. Ningún dato del Excel se incrusta en la distribución. */
const SETTINGS={tenantId:'5e23e4af-237d-4d97-bf6e-dca808015787',clientId:'c8a828c1-3a20-4876-96af-b4f30ce4abeb',matrixUrl:'https://fiasec-my.sharepoint.com/:x:/g/personal/dvillacis_fias_org_ec/IQCho4wt1qyXT6DnsDAU-r33AavyNjkIhd9xhXy5QgQvLkk',folderUrl:'https://fiasec-my.sharepoint.com/:f:/g/personal/dvillacis_fias_org_ec/IgAPa6xunze7Rp5Sbb3eoG4bAbnTiJ3TP2iPh5nhaefQklY?e=WH2jtK',scopes:['User.Read','Files.Read.All','Sites.Read.All']};
const $=id=>document.getElementById(id),state={rows:[],acts:0,warnings:[],account:null,client:null,origin:'',lastSync:'',search:'',year:'',fund:'',session:'',visible:12,page:'inicio',semantic:null,vectors:null,semanticBusy:false,pdfChunks:[],pdfVectors:null,pdfCancel:false,repository:[],changeLog:[],baselineStats:null,scanRunning:false,autoTimer:null,privateIndexLoaded:false,baselineDocuments:[]};const num=new Intl.NumberFormat('es-EC');const clean=x=>String(x??'').trim();const fold=x=>clean(x).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLocaleLowerCase('es');const esc=x=>clean(x).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));const snippet=(s,n=250)=>clean(s).slice(0,n)+(clean(s).length>n?'…':'');const funds=['FAP','FEIG','REM','PASF','PASNAP','CONSERVA AVES','BIOECONOMÍA','SOCIO BOSQUE','GALÁPAGOS','KFW'];const toast=(message)=>{const e=$('toast');e.textContent=message;e.classList.add('show');clearTimeout(toast.timer);toast.timer=setTimeout(()=>e.classList.remove('show'),5600)};
function safeUrl(value){try{const url=new URL(value);return url.protocol==='https:'&&(url.hostname==='fiasec-my.sharepoint.com'||url.hostname.endsWith('.sharepoint.com'))?url.href:''}catch{return ''}}
function normLink(value){const s=clean(value);if(s.startsWith('https://'))return safeUrl(s);const m=s.match(/\/:([bfx]):\/g\/personal\/[^\s"<>]+/i);return m?safeUrl('https://fiasec-my.sharepoint.com'+m[0]):''}
function rowTags(text){const s=fold(text);return funds.filter(f=>new RegExp(`(^|[^\\p{L}])${fold(f).replace(/ /g,'\\s+')}($|[^\\p{L}])`,'iu').test(s))}
function dateString(x,XLSX){if(x instanceof Date&&!Number.isNaN(x.getTime()))return x.toISOString().slice(0,10);if(typeof x==='number'&&x>20000&&x<70000){const d=XLSX.SSF.parse_date_code(x);if(d)return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`}return clean(x)}
function actYear(d,sheet){return (clean(d).match(/20\d{2}/)||clean(sheet).match(/20\d{2}/)||[])[0]||''}
function parseExcel(arrayBuffer){const X=window.XLSX;if(!X)throw Error('No está disponible el lector Excel.');const wb=X.read(new Uint8Array(arrayBuffer),{type:'array',cellDates:false,cellFormula:true});const result=[],warnings=[];
 for(const sheet of wb.SheetNames){const ws=wb.Sheets[sheet];const matrix=X.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});const header=matrix.findIndex(r=>fold(r.map(clean).join(' ')).includes('puntos a tratar')&&fold(r.map(clean).join(' ')).includes('resolucion'));if(header<0){warnings.push(`Hoja ${sheet}: no se identificó encabezado de asuntos y resoluciones.`);continue}let current=null,n=0;
 for(let i=header+1;i<matrix.length;i++){const a=matrix[i]||[];const [d,t,p,r,f]=[0,1,2,3,4].map(j=>a[j]);const date=clean(d),title=clean(t),subject=clean(p),resolution=clean(r),filename=clean(f);if(!date&&!title&&!subject&&!resolution&&!filename)continue;
 if(date||title){n++;current={date:dateString(d,X),title,actId:`${sheet}-${n}`,documentUrl:'',documentName:'',members:[]}}if(!current)continue;const cell=ws[`E${i+1}`];const link=normLink(clean(cell?.l?.Target)||clean(cell?.v));if(link)current.documentUrl=link;if(filename)current.documentName=filename;for(const prev of current.members){if(current.documentUrl)prev.documentUrl=current.documentUrl;if(current.documentName)prev.documentName=current.documentName}if(!subject&&!resolution)continue;const year=actYear(current.date,sheet);const program=rowTags([subject,resolution,current.title].join(' '));const amount=([...`${subject} ${resolution}`.matchAll(/(?:USD\s*|US\$\s*|\$\s*)([\d.,]+)/gi)]).map(x=>x[0]).slice(0,3).join(', ');const rec={id:`${sheet}-${i+1}`,year,date:current.date,act:current.title,actId:current.actId,session:/extraordinaria/i.test(current.title)?'Extraordinaria':/ordinaria/i.test(current.title)?'Ordinaria':'',resolution,subject,program,amount,documentUrl:current.documentUrl,documentName:current.documentName,sheet,rowNumber:i+1};result.push(rec);current.members.push(rec)} }
 return{rows:result,warnings,acts:new Set(result.map(r=>r.actId)).size}}
function loadScript(url,name){if(window[name])return Promise.resolve();return new Promise((resolve,reject)=>{const script=document.createElement('script');script.src=url;script.onload=resolve;script.onerror=()=>reject(Error('No se pudo cargar '+name+'. Revisa la conexión o las políticas de red.'));document.head.appendChild(script)})}
async function excelReader(){await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js','XLSX')}
async function msalReader(){if(window.msal)return window.msal;try{await loadScript('https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.13.0/lib/msal-browser.min.js','msal')}catch{ /* Alternate ESM build for hosts that do not carry the UMD build. */ const m=await import('https://esm.sh/@azure/msal-browser@4.13.0');return m}return window.msal}
async function client(){if(state.client)return state.client;const m=await msalReader();const uri=location.origin+location.pathname;const c=new m.PublicClientApplication({auth:{clientId:SETTINGS.clientId,authority:`https://login.microsoftonline.com/${SETTINGS.tenantId}`,redirectUri:uri},cache:{cacheLocation:'sessionStorage'}});await c.initialize();await c.handleRedirectPromise();state.client=c;state.account=c.getAllAccounts()[0]||null;authUi();return c}
async function login(){if(!/^https?:$/.test(location.protocol))throw Error('Abre la página desde HTTPS o localhost para autenticarte.');const c=await client();const response=await c.loginPopup({scopes:SETTINGS.scopes,prompt:'select_account'});state.account=response.account;authUi();toast('Sesión institucional iniciada. Sincronizando matriz…');await sync()}
async function logout(){const c=await client();const a=state.account;if(a)await c.logoutPopup({account:a,mainWindowRedirectUri:location.origin+location.pathname});state.account=null;state.rows=[];state.acts=0;state.origin='';state.vectors=null;state.pdfChunks=[];state.pdfVectors=null;state.repository=[];state.changeLog=[];state.baselineStats=null;state.privateIndexLoaded=false;state.baselineDocuments=[];authUi();updateDashboard();renderResults();toast('Sesión institucional finalizada; datos privados retirados de la memoria de la aplicación.')}
async function accessToken(scopes=SETTINGS.scopes){const c=await client();const account=state.account||c.getAllAccounts()[0];if(!account)throw Error('Primero inicia sesión con tu cuenta Microsoft 365.');try{return(await c.acquireTokenSilent({account,scopes})).accessToken}catch{return(await c.acquireTokenPopup({account,scopes})).accessToken}}
const encodeShare=url=>'u!'+btoa(unescape(encodeURIComponent(url))).replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
async function graph(path){const token=await accessToken();let res=await fetch('https://graph.microsoft.com/v1.0'+path,{headers:{Authorization:'Bearer '+token,Accept:'application/json'}});if([429,503].includes(res.status)){await new Promise(r=>setTimeout(r,Math.min(5000,Number(res.headers.get('Retry-After')||2)*1000)));res=await fetch('https://graph.microsoft.com/v1.0'+path,{headers:{Authorization:'Bearer '+token}})}if(!res.ok){const detail=await res.text();throw Error(`Microsoft Graph ${res.status}: ${detail.slice(0,180)}`)}return res.json()}
async function sync(){if(!state.account)throw Error('Inicia sesión para sincronizar la matriz desde SharePoint.');$('btnSync').disabled=true;connection('Sincronizando…',false);try{await excelReader();const meta=await graph(`/shares/${encodeShare(SETTINGS.matrixUrl)}/driveItem`);let url=meta['@microsoft.graph.downloadUrl'];if(!url&&meta.parentReference?.driveId){const more=await graph(`/drives/${encodeURIComponent(meta.parentReference.driveId)}/items/${encodeURIComponent(meta.id)}`);url=more['@microsoft.graph.downloadUrl']}if(!url)throw Error('Microsoft Graph no devolvió una URL temporal de descarga.');const file=await fetch(url,{cache:'no-store'});if(!file.ok)throw Error('Error al descargar matriz: '+file.status);const arr=await file.arrayBuffer();const bytes=new Uint8Array(arr);if(bytes[0]!==80||bytes[1]!==75)throw Error('La respuesta no parece un archivo XLSX.');const parsed=parseExcel(arr);putData(parsed,'Microsoft SharePoint');$('lastSync').textContent='Actualizado '+new Date().toLocaleString('es-EC',{dateStyle:'short',timeStyle:'short'});connection('SharePoint conectado',true);toast(`Matriz actualizada: ${num.format(parsed.rows.length)} entradas y ${parsed.acts} actas.`);await detectNewActas({automatic:true})}finally{$('btnSync').disabled=false}}
function connection(label,ok){$('connection').innerHTML=`<i></i> ${esc(label)}`;$('connection').classList.toggle('ok',!!ok)}
function authUi(){const logged=!!state.account;$('btnLogin').classList.toggle('hide',logged);$('btnLogout').classList.toggle('hide',!logged);if(logged){$('btnLogout').title='Cerrar sesión de '+(state.account.name||'Microsoft 365')} }
function putData(parsed,origin){state.rows=parsed.rows;state.acts=parsed.acts;state.warnings=parsed.warnings;state.origin=origin;state.visible=12;state.vectors=null;state.pdfVectors=null;updateDashboard();renderResults();renderAnalytics();if(state.semantic)$('semanticStatus').textContent='Modelo listo · reindexar';if(parsed.warnings.length)toast(parsed.warnings.join(' · ').slice(0,350))}
function setPage(page){state.page=page;document.querySelectorAll('.page').forEach(x=>x.classList.toggle('active',x.id===`page-${page}`));document.querySelectorAll('[data-page]').forEach(x=>x.classList.toggle('active',x.dataset.page===page));$('crumbPage').textContent={inicio:'Panorama ejecutivo',buscar:'Explorador de actas',inteligencia:'Asistente inteligente',analitica:'Analítica histórica'}[page];$('sidebar').classList.remove('open');window.scrollTo({top:0,behavior:'smooth'});if(page==='buscar')renderResults();if(page==='analitica')renderAnalytics()}
function updateDashboard(){const rows=state.rows;const years=[...new Set(rows.map(x=>x.year).filter(Boolean))].sort();$('kpiRows').textContent=rows.length?num.format(rows.length):'—';$('kpiActs').textContent=rows.length?num.format(state.acts):'—';$('kpiYears').textContent=rows.length?num.format(years.length):'—';$('kpiLinks').textContent=rows.length?num.format(rows.filter(x=>!!x.documentUrl).length):'—';$('healthTitle').textContent=rows.length?'Catálogo documental disponible':'Esperando datos';$('healthDesc').textContent=rows.length?`${num.format(rows.length)} registros consultables · ${state.warnings.length} alertas`: 'Conecta Microsoft 365 o importa la matriz local.';$('healthDot').classList.toggle('good',!!rows.length);$('originText').textContent=state.origin||'Sin conectar';const select=$('filterYear'),current=select.value;select.innerHTML='<option value="">Todos los años</option>'+years.map(y=>`<option value="${esc(y)}">${esc(y)}</option>`).join('');select.value=current;const fundSet=[...new Set(rows.flatMap(x=>x.program))].sort(),fund=$('filterFund'),prev=fund.value;fund.innerHTML='<option value="">Todos los fondos</option>'+fundSet.map(f=>`<option value="${esc(f)}">${esc(f)}</option>`).join('');fund.value=prev;const recent=rows.slice().sort((a,b)=>b.year.localeCompare(a.year)||b.rowNumber-a.rowNumber).slice(0,5);$('recentRows').innerHTML=recent.length?recent.map(r=>`<div class="recent-row" data-id="${esc(r.id)}"><span class="recent-year">${esc(r.year||'—')}</span><div class="recent-name"><strong>${esc(snippet(r.subject||r.resolution,125))}</strong><small>${esc(r.act||r.sheet)} · ${esc(r.date||r.year)}</small></div><span class="recent-arrow">↗</span></div>`).join(''):'<div class="empty-line">Conecta la matriz para visualizar el historial documental.</div>'}
function matches(r){if(state.year&&r.year!==state.year)return false;if(state.fund&&!r.program.includes(state.fund))return false;if(state.session&&r.session!==state.session)return false;return true}
const wordTokens=q=>fold(q).match(/[a-z0-9ñ]{2,}/g)||[];const STOP=new Set('de la el los las en por para con del al que una un sus su se fue sobre esta este desde como mas entre ha y o a e lo ante cuales cual quienes quien que son es estan fueron cuales cuales encontrar buscar busca dame mostrar muestra informacion antecedentes relacionados relacionadas relacionado relacionada acerca respecto decisiones tomadas existe existen cual cuales directorio actas acta'.split(' '));const expansions={'aprobacion':['aprobar','aprobado','aprobada','aprobacion','autoriza','autorizar'],'presupuesto':['presupuestaria','presupuestario','presupuesto','financiero'],'convenio':['acuerdo','alianza','convenio','cooperacion'],'donacion':['donaciones','donacion','donar','donante'],'contrato':['contratacion','contrataciones','contrato','servicio'],'fondo':['fondos','fondo','fap','feig','rem'],'gestion':['administracion','manejo','gestion'],'proyecto':['programa','iniciativa','proyecto']};
function lexicalRank(row,q){
 if(!q.trim())return 1;
 const terms=[...new Set(wordTokens(q).filter(t=>t.length>2&&!STOP.has(t)))];
 if(!terms.length)return 0;
 const names=[row.subject||'',row.resolution||'',`${row.act} ${row.program.join(' ')} ${row.year}`];
 const bags=names.map(n=>new Set(wordTokens(n)));
 const canon=t=>t.length>=6&&t.endsWith('es')?t.slice(0,-2):t.length>=5&&t.endsWith('s')?t.slice(0,-1):t;
 let score=0,hits=0,anchorMatches=0;
 const anchors=terms.filter(t=>/^(fap|feig|rem|pasf|pasnap|kfw|fias)$/.test(t));
 for(const term of terms){
   const lookup=anchors.includes(term)?null:Object.entries(expansions).find(([k,v])=>k===term||v.includes(term));
   const candidates=new Set([term,...(lookup?[lookup[0],...lookup[1]]:[])]);
   let best=0;
   for(const bagIndex of [1,0,2]){
     const bag=bags[bagIndex];
     for(const cand of candidates){
       if(bag.has(cand)||[...bag].some(w=>canon(w)===canon(cand)))
         best=Math.max(best,(bagIndex===1?5:bagIndex===0?4.7:2.5)*(term===cand?1:0.65));
     }
   }
   if(best){hits++;score+=best;if(anchors.includes(term))anchorMatches++}
 }
 if(anchors.length&&anchorMatches<anchors.length)return 0;
 if(hits<Math.max(1,Math.ceil(terms.length*.42)))return 0;
 return score/(terms.length*5);
}
function ranked(q,filtered=true){const rr=state.rows.filter(r=>!filtered||matches(r));return rr.map(r=>({r,score:lexicalRank(r,q)})).filter(x=>x.score>0).sort((a,b)=>b.score-a.score||b.r.year.localeCompare(a.r.year)||b.r.rowNumber-a.r.rowNumber)}
function renderResults(){state.year=$('filterYear').value;state.fund=$('filterFund').value;state.session=$('filterSession').value;state.search=$('searchQuery').value;const all=ranked(state.search),show=all.slice(0,state.visible);$('searchCount').textContent=all.length?`${num.format(all.length)} coincidencia${all.length===1?'':'s'} · ${num.format(state.rows.length)} entradas disponibles`:(state.rows.length?'Sin coincidencias':'Conecta SharePoint o importa un Excel para comenzar');$('results').innerHTML=show.length?show.map(({r,score})=>`<article class="result-card"><div class="result-top"><span class="result-id">REGISTRO ${esc(r.id)} · ${esc(r.session||'Sesión')}</span><span class="result-year">◷ ${esc(r.date||r.year||'Sin fecha')}</span></div><h3>${esc(snippet(r.subject||r.resolution||'Sin asunto',220))}</h3><p>${esc(snippet(r.resolution||'No se consigna texto resolutivo en este registro. Consulta el acta para verificar el acuerdo.',360))}</p><div class="result-bottom"><span class="chip">${esc(r.year||r.sheet)}</span>${r.program.slice(0,3).map(p=>`<span class="chip">${esc(p)}</span>`).join('')}<div class="result-actions"><button class="link-btn" data-id="${esc(r.id)}">Ver detalle →</button>${r.documentUrl?`<a href="${esc(r.documentUrl)}" target="_blank" rel="noopener noreferrer" class="link-btn">Abrir acta ↗</a>`:''}</div></div></article>`).join(''):'<div class="empty-state"><div class="empty-icon">⌕</div><h3>'+ (state.rows.length?'Sin resultados para esta búsqueda':'Tu conocimiento empieza aquí')+'</h3><p>'+(state.rows.length?'Ajusta los filtros o prueba con otros términos.':'Conecta Microsoft 365 para explorar el historial del Directorio.')+'</p></div>';$('loadMore').classList.toggle('hide',all.length<=state.visible)}
function showRecord(id){const r=state.rows.find(x=>x.id===id);if(!r)return;$('drawerBody').innerHTML=`<div class="drawer-tag">FICHA DOCUMENTAL · ${esc(r.id)}</div><h2 class="drawer-title">${esc(r.subject||'Registro del Directorio')}</h2><div class="drawer-field"><label>IDENTIFICACIÓN</label><p>${esc(r.act||'Acta sin título')}<br>${esc(r.date||r.year||'Fecha no registrada')} · ${esc(r.session||'Tipo no registrado')}</p></div><div class="drawer-field"><label>ASUNTO TRATADO</label><p>${esc(r.subject||'No disponible')}</p></div><div class="drawer-field"><label>RESOLUCIÓN REGISTRADA</label><p>${esc(r.resolution||'La matriz no registra texto resolutivo en esta entrada.')}</p></div><div class="drawer-field"><label>FONDO / INICIATIVA · MENCIONES TEXTUALES</label><p>${esc(r.program.join(', ')||'Sin etiqueta detectada')}</p></div><div class="drawer-field"><label>TRAZABILIDAD EN LA MATRIZ</label><p>Hoja: ${esc(r.sheet)}<br>Fila de origen: ${r.rowNumber}<br>Identificador de acta: ${esc(r.actId)}<br>Documento: ${esc(r.documentName||'No especificado')}</p></div><div class="drawer-source">${r.documentUrl?`<a class="btn btn-primary" target="_blank" rel="noopener noreferrer" href="${esc(r.documentUrl)}">↗ Abrir documento original en SharePoint</a>`:'<div class="api-status">Documento fuente sin enlace validado en la matriz</div>'}</div><div class="drawer-warning">La transcripción del índice no sustituye la lectura del acta firmada. Los permisos de SharePoint se aplican al abrir el origen.</div>`;$('drawerBackdrop').classList.remove('hide');$('recordDrawer').classList.add('open')}
function renderAnalytics(){const rows=state.rows;const years=[...new Set(rows.map(r=>r.year).filter(Boolean))].sort();const max=Math.max(1,...years.map(y=>rows.filter(r=>r.year===y).length));$('yearChart').innerHTML=years.length?years.map(y=>{const n=rows.filter(r=>r.year===y).length;return `<div class="bar-column" title="${esc(y)}: ${n} entradas"><span class="bar-value">${n}</span><div class="bar-shape" style="height:${Math.max(2,Math.round(n/max*159))}px"></div><small>${esc(y)}</small></div>`}).join(''):'<div class="empty-line">Sin registros disponibles.</div>';const counts=funds.map(f=>({name:f,count:rows.filter(r=>r.program.includes(f)).length})).filter(x=>x.count>0).sort((a,b)=>b.count-a.count).slice(0,8),m=Math.max(1,...counts.map(x=>x.count));$('fundChart').innerHTML=counts.length?counts.map(x=>`<div class="fund-row"><strong>${esc(x.name)}</strong><div class="fund-track"><div class="fund-fill" style="width:${x.count/m*100}%"></div></div><span>${x.count}</span></div>`).join(''):'<div class="empty-line">Sin etiquetas detectadas.</div>'}
// Resumen extractivo: siempre construido únicamente de textos concretos indexados y con identificadores comprobables.
function extractiveAnswer(question,hits,mode='contextual'){if(!hits.length)return '<p>No se encontraron antecedentes suficientemente relacionados con la consulta en la matriz disponible. Prueba con otro término o verifica el PDF original.</p>';const sources=hits.slice(0,5);const heading=`<p class="answer-par">Se localizaron <strong>${hits.length}</strong> registros relacionados en el índice. A continuación se presentan los extractos más pertinentes, sin agregar afirmaciones ajenas a la matriz.</p>`;return heading+sources.map(({r},i)=>`<div class="answer-citation"><b>[${i+1}] ${esc(r.year||'Sin año')} · ${esc(snippet(r.act||r.sheet,110))}</b><div><strong>Asunto:</strong> ${esc(snippet(r.subject||'No disponible',260))}</div><div><strong>Texto resolutivo:</strong> ${esc(snippet(r.resolution||'No consta en la matriz',390))}</div><div>${r.documentUrl?`<a href="${esc(r.documentUrl)}" target="_blank" rel="noopener noreferrer">Ver fuente original ↗</a>`:'<span class="muted-source">Sin vínculo de acta validado</span>'} · <button class="link-btn" data-id="${esc(r.id)}">Ficha ${esc(r.id)}</button></div></div>`).join('')+`<div class="answer-footer">Método: ${mode==='semantic'?'similitud neuronal multilingüe + clasificación contextual':'coincidencia contextual y expansión terminológica'}. Este es un extracto documentado, no una conclusión legal ni un resumen generativo. No se analizó el texto completo de los PDF.</div>`}
async function activateSemantic(){if(state.semanticBusy)return;state.semanticBusy=true;$('semanticBtn').disabled=true;$('aiProgress').textContent='Descargando modelo multilingüe para ejecutarlo localmente. La primera activación puede requerir varios minutos…';try{const {pipeline,env}=await import('https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.7.6/+esm');env.allowRemoteModels=true;state.semantic=await pipeline('feature-extraction','Xenova/paraphrase-multilingual-MiniLM-L12-v2',{dtype:'q8'});$('semanticStatus').textContent='Modelo neuronal activo';$('semanticBtn').textContent='✧ Actualizar índice semántico';$('aiMode').textContent='✧ IA semántica activa';toast('Modelo de embeddings multilingüe cargado en el navegador.');await indexSemantic();if(state.pdfChunks.length)await indexPdfVectors()}catch(e){$('aiProgress').textContent='No fue posible cargar el modelo neuronal: '+e.message+'. Puedes seguir usando la búsqueda contextual.';toast('Modelo no disponible en este navegador o red.')}finally{state.semanticBusy=false;$('semanticBtn').disabled=false}}
const embedding=async s=>{const output=await state.semantic(snippet(s,850),{pooling:'mean',normalize:true});return Array.from(output.data)};
async function indexSemantic(){if(!state.semantic){await activateSemantic();return}if(!state.rows.length){$('aiProgress').textContent='Modelo cargado. Conecta primero la matriz para generar vectores.';return}state.semanticBusy=true;$('semanticBtn').disabled=true;try{const vec=[];for(let i=0;i<state.rows.length;i++){const r=state.rows[i];vec.push(await embedding(`${r.subject}. ${r.resolution}. ${r.program.join(' ')}.`));if(i%8===0||i===state.rows.length-1){$('aiProgress').textContent=`Indexación semántica local: ${i+1}/${state.rows.length} registros procesados…`;await new Promise(resolve=>setTimeout(resolve,0))}}state.vectors=vec;$('aiProgress').textContent=`Índice semántico listo: ${vec.length} vectores. Los registros no se enviaron a un servicio de IA externo.`;toast('Índice de similitud semántica generado correctamente.')}catch(e){$('aiProgress').textContent='Error durante indexación: '+e.message;state.vectors=null}finally{state.semanticBusy=false;$('semanticBtn').disabled=false}}
const similarity=(a,b)=>{let n=0;for(let i=0;i<a.length;i++)n+=a[i]*b[i];return n};
async function retrieve(question){let list=ranked(question,false);if(state.semantic&&state.vectors?.length===state.rows.length){const query=await embedding(question);const scored=state.rows.map((r,i)=>({r,score:similarity(query,state.vectors[i])*.8+lexicalRank(r,question)*.2}));list=scored.filter(x=>x.score>.13).sort((a,b)=>b.score-a.score);return{list,mode:'semantic'}}return{list,mode:'contextual'}}
function sendChat(text,html,role){const node=document.createElement('div');node.className='chat-message '+(role==='user'?'chat-user':'chat-assistant');if(role==='user')node.textContent=text;else node.innerHTML='<div class="assistant-title">✳ &nbsp;FIAS · INVESTIGACIÓN DOCUMENTADA</div>'+html;$('chatThread').append(node);$('chatThread').scrollTop=$('chatThread').scrollHeight;return node}
// Indexación de documentos PDF originales desde SharePoint con permisos delegados.
// El navegador mantiene los pasajes exclusivamente en RAM, y no incluye PDFs en el despliegue público.
async function graphPage(url){const token=await accessToken();const res=await fetch(url.startsWith('https://graph.microsoft.com/')?url:'https://graph.microsoft.com/v1.0'+url,{headers:{Authorization:'Bearer '+token,Accept:'application/json'}});if(!res.ok)throw Error('Graph '+res.status+' al recorrer la biblioteca');return res.json()}
async function sharepointPdfs(){
 const root=await graph(`/shares/${encodeShare(SETTINGS.folderUrl)}/driveItem`);
 const driveId=root.parentReference?.driveId;
 if(!driveId)throw Error('No se pudo resolver la biblioteca compartida.');
 const folders=[{id:root.id,year:''}],files=[];
 let visits=0;
 while(folders.length){
  if(state.pdfCancel)break;
  if(++visits>300)throw Error('Se alcanzó el límite de subcarpetas.');
  const folder=folders.shift();
  let next=`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folder.id)}/children?$top=200`;
  while(next){
   const payload=await graphPage(next);
   for(const item of payload.value||[]){
    if(item.folder)folders.push({id:item.id,year:/^20\d{2}$/.test(item.name)?item.name:folder.year});
    else if(/\.pdf$/i.test(item.name||''))files.push({id:item.id,driveId,name:item.name,size:item.size||0,
     etag:item.eTag||item.cTag||'',modified:item.lastModifiedDateTime||'',year:folder.year||(item.name.match(/20\d{2}/)||[])[0]||'',url:safeUrl(item.webUrl||'')});
   }
   next=payload['@odata.nextLink']||null;
  }
 }
 return files.sort((a,b)=>b.year.localeCompare(a.year)||a.name.localeCompare(b.name));
}
async function pdfReader(){if(window.__fiasPdfjs)return window.__fiasPdfjs;const api=await import('https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.mjs');api.GlobalWorkerOptions.workerSrc='https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.mjs';window.__fiasPdfjs=api;return api}
function pdfChunksFromPages(pages,file){const chunks=[];let accumulator='',start=1;for(let i=0;i<pages.length;i++){const text=pages[i].trim();if(!text)continue;if(!accumulator)start=i+1;accumulator+=' '+text;if(accumulator.length>=1400||i===pages.length-1){chunks.push({id:`pdf-${file.id}-${start}`,file:file.name,fileId:file.id,year:file.year||(file.name.match(/20\d{2}/)||[])[0]||'',page:start,lastPage:i+1,text:accumulator.trim().slice(0,5500),documentUrl:file.url});accumulator=''}}if(accumulator.trim())chunks.push({id:`pdf-${file.id}-${start}`,file:file.name,fileId:file.id,year:file.year||'',page:start,lastPage:pages.length,text:accumulator.trim().slice(0,5500),documentUrl:file.url});return chunks}
async function parsePdfFile(file,api){const details=await graph(`/drives/${encodeURIComponent(file.driveId)}/items/${encodeURIComponent(file.id)}`);const url=details['@microsoft.graph.downloadUrl'];if(!url)throw Error('Sin URL temporal de descarga');const response=await fetch(url,{cache:'no-store'});if(!response.ok)throw Error('Descarga '+response.status);const bytes=await response.arrayBuffer();const doc=await api.getDocument({data:new Uint8Array(bytes),useSystemFonts:true}).promise;const pages=[];for(let i=1;i<=Math.min(250,doc.numPages);i++){if(state.pdfCancel)break;const page=await doc.getPage(i);const content=await page.getTextContent();pages.push(content.items.map(it=>it.str||'').join(' '));page.cleanup()}await doc.destroy();return pdfChunksFromPages(pages,file)}
async function indexSharePointPdfs(){return detectNewActas({full:true})}
async function indexPdfVectors(){if(!state.semantic||!state.pdfChunks.length)return;const vectors=[];for(let i=0;i<state.pdfChunks.length;i++){if(state.pdfCancel)break;vectors.push(await embedding(state.pdfChunks[i].text.slice(0,1200)));if(i%10===0){$('pdfIndexProgress').textContent=`Embeddings PDF: ${i+1}/${state.pdfChunks.length} pasajes`;await new Promise(r=>setTimeout(r,0))}}state.pdfVectors=vectors.length===state.pdfChunks.length?vectors:null}
function pdfLexical(question){
 if(!state.pdfChunks.length)return [];
 const tokens=[...new Set(wordTokens(question).filter(t=>t.length>2&&!STOP.has(t)))];
 if(!tokens.length)return [];
 const anchors=tokens.filter(t=>/^(fap|feig|rem|pasf|pasnap|kfw|fias)$/.test(t));
 return state.pdfChunks.map(c=>{
    const bag=new Set(wordTokens(c.text||''));
    if(anchors.length&&anchors.some(t=>!bag.has(t)))return {r:c,score:0};
    const hit=tokens.reduce((sum,t)=>sum+(bag.has(t)?1:0),0);
    return {r:c,score:hit/tokens.length};
 }).filter(x=>x.score>0).sort((a,b)=>b.score-a.score).slice(0,5);
}
async function pdfEvidence(question){let hits=pdfLexical(question);if(state.pdfVectors?.length===state.pdfChunks.length&&state.semantic){const v=await embedding(question);hits=state.pdfChunks.map((r,i)=>({r,score:similarity(v,state.pdfVectors[i])})).filter(x=>x.score>.25).sort((a,b)=>b.score-a.score).slice(0,5)}return hits}
function renderPdfEvidence(hits){if(!hits.length)return '';return `<div class="pdf-evidence-title"><strong>Pasajes de documentos PDF originales (${hits.length})</strong><span>Extracción automática · verificar contra el acta firmada</span></div>`+hits.map(({r},i)=>`<div class="answer-citation"><b>[PDF ${i+1}] ${esc(r.file||r.fileName)} · página ${r.page}${r.lastPage>r.page?'–'+r.lastPage:''}</b><div>${esc(snippet(r.text,520))}</div>${r.documentUrl?`<a target="_blank" rel="noopener noreferrer" href="${esc(r.documentUrl)}">Consultar acta en SharePoint ↗</a>`:'<span class="muted-source">Documento sin enlace confirmado</span>'}</div>`).join('')}
// v07: índice local preparado previamente; nunca se publica en el frontend de GitHub Pages.
function userTrackerKey(){return 'fias-actas-seen-v07-'+(state.account?.homeAccountId||'anónimo')}
function readTracker(){try{const v=JSON.parse(localStorage.getItem(userTrackerKey())||'{}');return v && typeof v==='object'?v:{}}catch{return {}}}
function storedTracker(items,previous={}){const out={...previous};for(const x of items){out[x.id]={etag:x.etag,size:x.size,modified:x.modified}}localStorage.setItem(userTrackerKey(),JSON.stringify(out))}
function stampFile(f){return JSON.stringify([f.etag,f.size,f.modified])}
function reportStatus(line){if(!line)return;const box=$('repoStatus');if(box)box.textContent=line}
function reportRender(){
 const docs=state.repository||[];
 const total=docs.length||state.baselineDocuments.length,news=state.changeLog.filter(x=>x.status==='nuevo').length,updates=state.changeLog.filter(x=>x.status==='modificado').length;
 reportStatus(`${total} ${docs.length?'PDF revisados en SharePoint':'PDF del índice local'} · ${news} nuevos · ${updates} actualizados · ${state.pdfChunks.length} fragmentos en memoria`);
 if($('repoReport'))$('repoReport').innerHTML=state.changeLog.length?
   state.changeLog.slice(0,16).map(x=>`<div class="report-entry"><span class="chip">${esc(x.status)}</span><strong>${esc(x.name)}</strong><small>${esc(x.year||'')}</small>${safeUrl(x.url)?`<a href="${esc(safeUrl(x.url))}" target="_blank" rel="noopener noreferrer">Fuente ↗</a>`:''}<span>${esc(x.note||'')}</span></div>`).join(''):
   '<p class="mini-copy">No se han identificado altas posteriores a la primera inspección de esta cuenta. La primera revisión establece la línea base.</p>';
}
async function detectNewActas({automatic=false,full=false}={}){
 if(!state.account||state.scanRunning)return;
 state.scanRunning=true;state.pdfCancel=false;
 $('pdfIndexBtn').disabled=true;
 $('pdfCancelBtn').classList.remove('hide');
 $('pdfIndexProgress').textContent='Comparando el repositorio SharePoint con el último control…';
 try{
  const files=await sharepointPdfs();
  const previous=readTracker(), first=!Object.keys(previous).length;
  const changes=files.filter(f=>!previous[f.id]||stampFile(previous[f.id])!==stampFile(f));
  state.repository=files;
  // Primera revisión = línea base, no 46 falsas actas "nuevas".
  const candidates=full?files:first?[]:changes;
  const handled=[];
  const successfullyProcessed=new Set();
  if(!first){
    const prevIds=Object.keys(previous);
    for(const id of prevIds)if(!files.some(f=>f.id===id))handled.push({name:'Identificador '+id,status:'no-localizada',note:'El archivo no apareció en esta revisión; verificar permisos/carpeta.'});
  }
  let done=0;
  if(candidates.length){
   const api=await pdfReader();
   for(const file of candidates){
    if(state.pdfCancel)break;
    const isNew=!previous[file.id];
    const entry={id:file.id,name:file.name,url:file.url,year:file.year,status:full&&!isNew?'reprocesada':isNew?'nuevo':'modificado',note:''};
    try{
      if(file.size>45*1024*1024)throw Error('Supera límite de 45 MB de lectura en navegador; se conserva en inventario.');
      const chunks=await parsePdfFile(file,api);
      state.pdfChunks=state.pdfChunks.filter(r=>r.fileId!==file.id && !(r.file===file.name && String(r.year)===String(file.year)));
      state.pdfChunks.push(...chunks);
      state.pdfVectors=null;
      entry.note=chunks.length?`${chunks.length} pasajes extraídos de texto PDF`:'Sin capa de texto: requiere OCR; se conserva el PDF en el inventario.';
      successfullyProcessed.add(file.id);
    }catch(e){entry.status='incidencia';entry.note='Pendiente de reintento: '+e.message}
    handled.push(entry);
    done++;
    $('pdfIndexProgress').textContent=`${done}/${candidates.length} actas revisadas · ${state.pdfChunks.length} pasajes cargados`;
    await new Promise(resolve=>setTimeout(resolve,0));
   }
  }
  // Evitar confirmar como procesadas las entradas no terminadas por una cancelación.
  const candidateIds=new Set(candidates.map(x=>x.id));
  const trackerSave=files.filter(f=>first || !candidateIds.has(f.id) || successfullyProcessed.has(f.id));
  storedTracker(trackerSave,previous);
  state.changeLog=[...handled,...state.changeLog].slice(0,120);
  if(state.semantic&&state.pdfChunks.length && candidates.length)await indexPdfVectors();
  const label=first?`Primera revisión: ${files.length} PDF inventariados. Para usar el índice ya preparado, importa el archivo privado. Nuevos documentos se procesarán en futuras revisiones.`:
   `${files.length} PDF inventariados · ${candidates.length} cambios detectados · ${handled.length} incorporados al reporte. ${state.pdfCancel?'Revisión interrumpida.':''}`;
  $('pdfIndexProgress').textContent=label;
  reportRender();
  if(!automatic||candidates.length)toast(label);
 }catch(e){$('pdfIndexProgress').textContent='No se completó el control de novedades: '+e.message;if(!automatic)toast(e.message)}
 finally{state.scanRunning=false;$('pdfIndexBtn').disabled=false;$('pdfCancelBtn').classList.add('hide')}
}
// Importar un catálogo JSON privado. Nunca se conserva en LocalStorage ni se transmite a GitHub.
async function loadPrivateIndex(file){
 if(file.size>35*1024*1024)throw Error('Índice superior a 35 MB.');
 const data=JSON.parse(await file.text());
 if(data.schema!=='fias-directorio-private-v07'||!Array.isArray(data.rows)||!Array.isArray(data.pdfChunks)||!Array.isArray(data.documents))throw Error('No corresponde al índice privado v07.');
 const rows=data.rows.map((r,i)=>({id:r.id||`privado-${i}`,year:String(r.year||''),date:r.date||'',
  act:r.actTitle||'',actId:r.actId||'',session:/extraordinaria/i.test(r.actTitle||'')?'Extraordinaria':/ordinaria/i.test(r.actTitle||'')?'Ordinaria':'',
  subject:r.subject||'',resolution:r.resolution||'',program:rowTags(`${r.subject} ${r.resolution} ${r.actTitle}`),
  amount:'',documentUrl:safeUrl(r.shareUrl||''),documentName:r.fileName||'',sheet:r.sheet||'',rowNumber:r.rowNumber||0}));
 if(!state.rows.length)putData({rows,acts:new Set(rows.map(x=>x.actId)).size,warnings:[]},'Índice privado local');
 const parsed=data.pdfChunks.map((r,i)=>({id:r.id||`pdf-${i}`,file:r.fileName||r.file||'Documento',fileId:r.fileId||'',year:String(r.year||''),
  page:r.page||0,lastPage:r.page||0,text:r.text||r.content||'',documentUrl:safeUrl(r.shareUrl||r.documentUrl||'')}));
 state.pdfChunks=parsed;state.privateIndexLoaded=true;state.baselineStats=data.stats;state.baselineDocuments=data.documents;
 state.pdfVectors=null;reportRender();
 $('pdfIndexProgress').textContent=`Índice privado importado: ${rows.length} registros de matriz y ${parsed.length} extractos de PDF; ${(data.stats?.pagesNeedingOCR||0)} páginas todavía requieren OCR.`;
 toast('Índice privado cargado únicamente en la memoria del navegador.');
}
function downloadFile(name,mime,text){const u=URL.createObjectURL(new Blob([text],{type:mime}));const a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),3000)}
function exportChangesReport(){
 const head=['Fecha de control','Año','Archivo','Tamaño bytes','Estado en esta revisión','Resultado de procesamiento','Enlace SharePoint','Capacidad de lectura'];
 const now=new Date().toISOString(),statusMap=new Map(state.changeLog.map(e=>[e.id||'',e]));
 const baseline=new Map((state.baselineDocuments||[]).map(x=>[`${x.year}|${x.fileName}`,x]));
 const inventory=state.repository.length?state.repository:(state.baselineDocuments||[]).map(d=>({id:d.path,name:d.fileName,year:d.year,size:d.bytes||'',url:d.shareUrl||''}));
 const rows=inventory.map(f=>{
   const last=statusMap.get(f.id)||{},historical=baseline.get(`${f.year}|${f.name}`),chunks=state.pdfChunks.filter(x=>x.file===f.name&&String(x.year)===String(f.year)).length;
   const lectura=historical?`${historical.readablePages}/${historical.pages} páginas con texto`:(chunks?`${chunks} pasajes indexados`:'Por comprobar / PDF escaneado');
   return [now,f.year,f.name,f.size,last.status||'Existente (sin cambios)',last.note||'Inventariado en SharePoint',f.url||historical?.shareUrl||'',lectura];
 });
 downloadFile('FIAS_Reporte_integral_actas.csv','text/csv;charset=utf-8','\ufeff'+[head,...rows].map(r=>r.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(';')).join('\r\n'));
 toast('Reporte integral exportado: incluye el histórico visible y las actas nuevas o actualizadas.');
}
function exportCurrentIndex(){
 if(!state.account){toast('Conecta Microsoft 365 antes de exportar un índice institucional.');return}
 const data={schema:'fias-directorio-private-v07',generatedAt:new Date().toISOString(),stats:state.baselineStats||{},
 rows:state.rows.map(r=>({id:r.id,year:r.year,date:r.date,actId:r.actId,actTitle:r.act,subject:r.subject,resolution:r.resolution,fileName:r.documentName,shareUrl:r.documentUrl,sheet:r.sheet,rowNumber:r.rowNumber})),
 pdfChunks:state.pdfChunks.map(r=>({id:r.id,fileId:r.fileId,fileName:r.file,year:r.year,page:r.page,text:r.text,shareUrl:r.documentUrl})),
 documents:state.repository.map(r=>({fileName:r.name,year:r.year,shareUrl:r.url,id:r.id,etag:r.etag,size:r.size}))};
 downloadFile('FIAS_INDICE_PRIVADO_ACTUALIZADO.json','application/json;charset=utf-8',JSON.stringify(data));
}
async function ask(question){
 question=String(question||'').trim();if(!question)return;
 if(!state.rows.length){toast('Sincroniza la matriz o importa el índice privado para consultar.');return}
 if(state.semanticBusy){toast('Indexación semántica en curso.');return}
 if($('chatThread').querySelector('.chat-empty'))$('chatThread').innerHTML='';
 sendChat(question,'','user');$('aiQuery').value='';
 const node=sendChat('','<p class="rag-loading">Recuperando antecedentes y fuentes documentales locales…</p>','assistant');
 try{
  const {list,mode}=await retrieve(question),pdfHits=await pdfEvidence(question);
  node.innerHTML='<div class="assistant-title">◈ &nbsp; SÍNTESIS EXTRACTIVA CON REFERENCIAS</div>'+extractiveAnswer(question,list,mode)+renderPdfEvidence(pdfHits)+
   '<div class="answer-footer">Sin servicios Azure AI/OpenAI. La síntesis extrae texto de fuentes, no genera afirmaciones no verificadas. Las páginas escaneadas sin texto siguen pendientes de OCR.</div>';
 }catch(e){node.innerHTML='<div class="assistant-title">Consulta incompleta</div><p class="rag-error">'+esc(e.message)+'</p>'}
 node.scrollIntoView({block:'nearest',behavior:'smooth'});
}
async function runHybridSearch(){
 const q=$('searchQuery').value.trim();if(!q){toast('Escribe tu consulta.');return}
 const box=$('ragRemoteResults');box.innerHTML='<div class="rag-loading">◈ Evaluando índice documental local…</div>';
 try{
  const {list,mode}=await retrieve(q),pdfHits=await pdfEvidence(q);
  box.innerHTML=`<div class="rag-results-summary"><b>Consulta local · ${esc(mode)}</b><span>${list.length} referencias</span><button id="clearHybrid" class="text-btn">Cerrar</button></div>`+extractiveAnswer(q,list,mode)+renderPdfEvidence(pdfHits);
  $('clearHybrid').onclick=()=>box.innerHTML='';
 }catch(e){box.innerHTML='<div class="rag-error">'+esc(e.message)+'</div>'}
}
function downloadResults(){const entries=ranked(state.search).map(({r})=>r);if(!entries.length){toast('No existen resultados para exportar.');return}const cols=['id','year','date','act','subject','resolution','program','documentUrl','sheet','rowNumber'];const csv=[cols.join(';'),...entries.map(r=>cols.map(c=>`"${clean(c==='program'?r[c].join(', '):r[c]).replace(/"/g,'""')}"`).join(';'))].join('\r\n');const url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='FIAS_Directorio_consulta.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}
function events(){document.querySelectorAll('[data-page]').forEach(b=>b.addEventListener('click',()=>setPage(b.dataset.page)));document.querySelectorAll('[data-go]').forEach(b=>b.addEventListener('click',()=>setPage(b.dataset.go)));$('menu').onclick=()=>$('sidebar').classList.toggle('open');$('btnLogin').onclick=()=>login().catch(e=>toast(e.message));$('btnLogout').onclick=()=>logout().catch(e=>toast(e.message));$('btnSync').onclick=()=>sync().catch(e=>{connection('Error de conexión',false);toast(e.message)});$('importBtn').onclick=()=>$('fileInput').click();$('fileInput').onchange=async(e)=>{const file=e.target.files?.[0];if(!file)return;try{await excelReader();const parsed=parseExcel(await file.arrayBuffer());putData(parsed,'Importación local');connection('Excel local',false);toast(`Archivo procesado: ${num.format(parsed.rows.length)} registros.`)}catch(err){toast('Error al importar: '+err.message)}e.target.value=''};$('quickForm').onsubmit=e=>{e.preventDefault();setPage('buscar');$('searchQuery').value=$('quickQuery').value;state.visible=12;renderResults()};document.querySelectorAll('[data-suggest]').forEach(b=>b.onclick=()=>{$('quickQuery').value=b.dataset.suggest;$('quickForm').requestSubmit()});$('searchGo').onclick=()=>{state.visible=12;renderResults()};$('searchQuery').oninput=()=>{state.visible=12;renderResults()};['filterYear','filterFund','filterSession'].forEach(id=>$(id).onchange=()=>{state.visible=12;renderResults()});$('clearFilters').onclick=()=>{$('filterYear').value='';$('filterFund').value='';$('filterSession').value='';$('searchQuery').value='';state.visible=12;renderResults()};$('loadMore').onclick=()=>{state.visible+=12;renderResults()};$('results').addEventListener('click',e=>{const b=e.target.closest('[data-id]');if(b)showRecord(b.dataset.id)});$('recentRows').addEventListener('click',e=>{const b=e.target.closest('[data-id]');if(b)showRecord(b.dataset.id)});$('chatThread').addEventListener('click',e=>{const b=e.target.closest('[data-id]');if(b)showRecord(b.dataset.id)});$('closeDrawer').onclick=closeDrawer;$('drawerBackdrop').onclick=closeDrawer;document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDrawer();if(e.key==='Enter'&&e.ctrlKey&&document.activeElement===$('aiQuery'))$('aiForm').requestSubmit()});$('aiForm').onsubmit=e=>{e.preventDefault();ask($('aiQuery').value)};document.querySelectorAll('[data-prompt]').forEach(b=>b.onclick=()=>{$('aiQuery').value=b.dataset.prompt;$('aiForm').requestSubmit()});$('semanticBtn').onclick=()=>state.semantic?indexSemantic().then(()=>indexPdfVectors()):activateSemantic();$('pdfIndexBtn').onclick=()=>detectNewActas({full:true});$('pdfCancelBtn').onclick=()=>{state.pdfCancel=true;$('pdfIndexProgress').textContent='Solicitando cancelación…'};$('hybridSearchBtn').onclick=runHybridSearch;$('exportBtn').onclick=downloadResults;$('themeBtn').onclick=()=>{const dark=document.documentElement.classList.toggle('dark');localStorage.setItem('fias_theme',dark?'dark':'light')};$('btnChanges').onclick=()=>detectNewActas({});$('btnReport').onclick=exportChangesReport;$('btnIndexImport').onclick=()=>$('indexInput').click();$('indexInput').onchange=async e=>{const f=e.target.files?.[0];if(!f)return;try{await loadPrivateIndex(f)}catch(err){toast(err.message)}e.target.value=''};$('btnIndexExport').onclick=exportCurrentIndex}
function closeDrawer(){$('recordDrawer').classList.remove('open');$('drawerBackdrop').classList.add('hide')}
function init(){
 $('matrixLink').href=SETTINGS.matrixUrl;$('folderLink').href=SETTINGS.folderUrl;
 if(localStorage.getItem('fias_theme')==='dark')document.documentElement.classList.add('dark');
 events();updateDashboard();renderAnalytics();reportRender();
 if(/^https?:$/.test(location.protocol))client().then(async()=>{
   if(state.account){try{await sync()}catch(e){console.warn('La sincronización M365 requiere interacción:',e.message)}}
 }).catch(e=>console.warn('FIAS: No se pudo restaurar la sesión M365.',e));
 state.autoTimer=setInterval(()=>{if(state.account&&!document.hidden)detectNewActas({automatic:true})},15*60*1000);
}
init();
