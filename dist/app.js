/* OrderFlow — browser-local trade order workspace. */
(() => {
  'use strict';
  const $ = (s,root=document) => root.querySelector(s);
  const $$ = (s,root=document) => Array.from(root.querySelectorAll(s));
  const esc = v => String(v ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const norm = v => String(v ?? '').normalize('NFKC').toLowerCase().replace(/[\s\-—_，,。:：;；（）()【】\[\]\/\\]/g,'');
  const uid = () => crypto.randomUUID?.() || Math.random().toString(36).slice(2)+Date.now();
  // 单元格没套日期格式时，Excel 里只剩序列号；这里按日期字段还原。
  function excelSerialText(v){
    const n=Number(String(v).trim());
    if(!Number.isFinite(n)||n<1||n>2958465)return String(v);
    const d=new Date(Date.UTC(1899,11,30)+Math.round(n*86400000));
    if(Number.isNaN(d.getTime()))return String(v);
    const iso=d.toISOString();
    return iso.endsWith('T00:00:00.000Z')?iso.slice(0,10):iso.slice(0,19).replace('T',' ');
  }
  // 只保留 11 位手机号，去掉分机、括号、空格等后缀。
  function normalizePhone(v){
    const raw=String(v==null?'':v).trim();
    if(!raw)return '';
    const digits=raw.replace(/\D/g,'');
    const hit=digits.match(/1[3-9]\d{9}/);
    return hit?hit[0]:raw;
  }
  function nextProductSku(){const used=new Set(state?state.products.map(p=>norm(p.internalSku)):[]);let n=(state?state.products.length:0)+1,code;do{code=`OF-${String(n++).padStart(3,'0')}`;}while(used.has(norm(code)));return code;}
  const copy = obj => JSON.parse(JSON.stringify(obj));
  const stamp = () => new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
  const icons = {
    upload:'<path d="M12 16V3m-5 5 5-5 5 5M4 15v5a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-5"/>',
    file:'<path d="M14 2H5v20h14V7l-5-5Zm0 0v6h5M8 12h8m-8 4h8"/>',
    check:'<path d="m5 12 4 4L19 6"/>',
    box:'<path d="m12 2 9 5v10l-9 5-9-5V7l9-5Zm0 10v10M3 7l9 5 9-5M8 4l9 5"/>',
    clock:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    arrow:'<path d="M4 12h16m-6-6 6 6-6 6"/>',
    download:'<path d="M12 3v12m-5-5 5 5 5-5M4 16v5h16v-5"/>',
    plus:'<path d="M12 4v16M4 12h16"/>',
    shield:'<path d="m12 2 8 4v6c0 5-8 10-8 10S4 17 4 12V6l8-4Z"/><path d="m8 12 3 3 5-6"/>',
    search:'<circle cx="10" cy="10" r="6"/><path d="m15 15 6 6"/>',
    link:'<path d="m10 13 4-4m-7 7-2 2a4 4 0 0 0 6 5l4-4a4 4 0 0 0 0-6m2-5 2-2a4 4 0 0 0-6-5L9 5a4 4 0 0 0 0 6"/>'
  };
  const icon = (name,cls='') => `<svg class="icon ${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name]||icons.file}</svg>`;
  const FIELDS = [
    ['orderNo','订单号',['订单编号','订单号','订单编码','订单ID','外部订单号','订单号码']],
    ['productName','商品名称',['商品名称','产品名称','货品名称','品名','名称','商品','标准商品名','商品标题','宝贝名称']],
    ['quantity','数量',['购买数量','商品数量','数量','件数','订购数量','销售数量','采购数量','发货数量']],
    ['recipient','收货人',['收货人/提货人','收货人','收件人','收货姓名','收件姓名','收货人姓名','收件人姓名']],
    ['phone','收货电话',['收货人手机号/提货人手机号','收货电话','收货人电话','收件人手机','收货人手机号','手机号','联系电话','电话','手机','收货手机','收货人手机','收件人手机号','联系方式']],
    ['address','详细地址',['收货地址/提货地址','收货地址','详细地址','收件地址','地址','收货人地址','收件人地址','地址包含省市区','收货详细地址']],
    ['sourceSku','客户商品编码',['SKU码','SKU编码','商品编码','产品编码','货品编码','商家编码','sku','条码','69码','商品条码']],
    ['sourceSpec','来源规格',['规格属性','商品规格','规格','产品规格','商品规格信息','规格信息','颜色及规格','属性']],
    ['sourceLineId','订单明细ID',['订单商品编号','订单明细编号','子订单编号','明细编号']],
    ['buyerNickname','买家昵称',['买家昵称','买家账号','客户昵称']],
    ['remark','备注',['买家备注','客户备注','订单备注','备注']],
    ['date','下单时间',['下单时间','支付时间','付款完成时间','订单日期','时间','订单时间','创建时间','付款时间']],
    ['carrier','物流公司',['物流公司','快递公司','承运商','配送方式']],
    ['trackingNo','快递单号',['快递单号','运单号','物流单号','快递编号']],
    ['province','省',['省','省份','收货省']],['city','市',['市','城市','收货市']],['district','区',['区','区县','收货区']],
    ['unitPrice','源单价',['成交价','商品单价','单价','销售单价']]
  ];
  const EXPORT_FIELDS = [['orderNo','订单号'],['buyerNickname','买家昵称'],['recipient','收货人'],['address','详细地址'],['phone','收货电话'],['supplierSku','供应商商品编码'],['quantity','数量'],['productName','统一商品名称'],['sourceProductName','来源商品名称'],['internalSku','内部编码'],['spec','标准规格'],['customer','客户'],['remark','备注'],['date','下单时间'],['province','省'],['city','市'],['district','区'],['sourceSku','客户商品编码'],['unitPrice','源单价'],['constant','固定值'],['blank','留空']];
  const UNIFIED = [['orderNo','订单号'],['customer','客户'],['sourceProductName','来源商品名称'],['productName','统一商品名称'],['internalSku','内部编码'],['supplierName','供应商'],['supplierSku','供应商商品编码'],['spec','标准规格'],['quantity','数量'],['recipient','收货人'],['phone','收货电话'],['address','详细地址'],['buyerNickname','买家昵称'],['remark','备注'],['date','下单时间'],['sourceSku','客户商品编码'],['unitPrice','源单价'],['sourceFile','来源文件'],['sourceSheet','来源工作表'],['sourceRow','原始行号']];
  let state, db, view='import', filter='all', query='', page=1, productPage=1, supplierFilter='all', selectedExport='', pending=[], importDraft=null, modalReturn=null, storageFailure=false;
  const seed = () => ({version:1,suppliers:[copy(OrderSeed.supplier)],products:copy(OrderSeed.products),batches:[],activeBatchIds:[],orders:[],customers:[],recharges:[],profiles:{},unified:null,history:[]});
  const pById = id => state.products.find(p=>p.id===id);
  const sById = id => state.suppliers.find(s=>s.id===id);
  const invalidate = () => { state.unified=null; };
  const history = action => {state.history.unshift({time:new Date().toISOString(),action});state.history=state.history.slice(0,50);};
  const activeRows = () => {const ids=new Set(state.activeBatchIds);return state.orders.filter(r=>ids.has(r.batchId));};
  function migrateScope(v){
    if(!Array.isArray(v.activeBatchIds))v.activeBatchIds=v.batches.filter(b=>!b.exportedAt).map(b=>b.id);
    const ids=new Set(v.batches.map(b=>b.id));v.activeBatchIds=[...new Set(v.activeBatchIds.filter(id=>ids.has(id)))];
    // Quote-sheet fields added later: keep older saved workspaces readable.
    for(const p of v.products){p.barcode??='';p.brand??='';p.category??='';p.cost??='';p.groupPrice??='';p.broadcastPrice??='';p.suggestPrice??='';p.marginOwn??='';p.marginBroadcast??='';p.jdPrice??='';p.link??='';p.sourceFile??='';}
    for(const s of v.suppliers)if(s.derived===undefined)s.derived=false;
    if(!Array.isArray(v.customers)){v.customers=[];const seen=new Set();for(const o of v.orders){const n=String(o.customer||'').trim();if(!n)continue;const k=norm(n);if(seen.has(k))continue;seen.add(k);v.customers.push({id:uid(),name:n,contact:'',phone:'',note:'',deposit:'',createdAt:new Date().toISOString()});}}
    for(const c of v.customers)if(c.deposit===undefined)c.deposit='';
    for(const c of v.customers)for(const k of ['companyName','creditCode','companyAddress','contact','phone','bankName','bankBranch','bankCode','bankAccount','license','businessScope'])if(c[k]===undefined)c[k]='';
    for(const o of v.orders){if(o.carrier===undefined)o.carrier='';if(o.trackingNo===undefined)o.trackingNo='';if(o.invoiced===undefined)o.invoiced=false;}
    if(!Array.isArray(v.recharges)){v.recharges=[];for(const c of v.customers){const n=Number(c.deposit);if(Number.isFinite(n)&&n>0)v.recharges.push({id:uid(),customer:c.name,amount:n,date:String(c.createdAt||new Date().toISOString()).slice(0,10),note:'历史充值',createdAt:new Date().toISOString()});}}
    for(const r of v.recharges){if(!r.kind)r.kind='customer';if(!r.name)r.name=r.customer||'';}
    // Correct only the known sample-alias constraint from the initial prototype.
    for(const p of v.products)for(const a of p.aliases||[]){const c=OrderSeed.candidates.find(c=>c.productId===p.id&&norm(c.customer)===norm(a.customer)&&norm(c.name)===norm(a.name));if(c&&c.spec&&norm(a.spec)===norm(c.spec))a.spec='';}
  }
  function refreshBatchExportFlags(){
    for(const b of state.batches){const rows=state.orders.filter(r=>r.batchId===b.id&&!r.excluded);if(rows.length&&rows.every(r=>r.exportedAt))b.exportedAt=rows.map(r=>r.exportedAt).sort().pop();else delete b.exportedAt;}
  }
  function includeNewBatch(id){refreshBatchExportFlags();state.activeBatchIds=state.activeBatchIds.filter(bid=>!state.batches.find(b=>b.id===bid)?.exportedAt);if(!state.activeBatchIds.includes(id))state.activeBatchIds.push(id);}
  function markExported(rows){const ids=new Set(rows.map(r=>r.id)),time=new Date().toISOString();for(const r of state.orders)if(ids.has(r.id))r.exportedAt=time;refreshBatchExportFlags();}
  async function init() {
    try {
      const loaded = await OrderStore.load();
      state = loaded ? loaded.state : null;
    } catch(e) { storageFailure = true; console.error('[OrderFlow] 读取工作区失败', e); }
    if(!state || state.version!==1) state=seed();
    try{migrateScope(state);}catch(e){console.error("[OrderFlow] 数据迁移失败",e);}
    bindShell();
    try{render();}catch(e){
      console.error("[OrderFlow] 页面渲染失败",e);
      const box=document.getElementById("content");
      if(box){box.innerHTML="";const note=document.createElement("div");note.className="alert warning";note.style.margin="24px";note.textContent="页面初始化失败："+(e&&e.message?e.message:e)+"。请在浏览器控制台查看详细错误，或用右上角「恢复备份」还原数据。";box.append(note);}
      toast("页面初始化失败，请查看控制台或恢复备份。","warning");
    }
    if(storageFailure)toast('浏览器未允许本地存储，请在关闭前导出备份。','warning');
  }
  async function save() {
    try { await OrderStore.save(copy(state)); }
    catch(e){
      const msg = e && e.message ? e.message : '保存失败，请立即备份数据。';
      if ((e && e.conflict) || /已被他人修改/.test(msg)) {
        window.__saveFailedAt = Date.now();
        setTimeout(() => openModal('保存失败：数据冲突', `<p>本次写入由于其他人刚刚修改了同一份数据而失败，<strong>你这次的修改没有被保存</strong>。</p><p class="muted">请刷新页面获取最新数据后，再重新操作一次。</p>`, `<button class="btn secondary" data-cancel>关闭</button><button class="btn" id="conflict-reload">刷新页面</button>`, true), 0);
        setTimeout(() => { const btn = $('#conflict-reload'); if (btn) btn.onclick = () => location.reload(); }, 0);
        return;
      }
      storageFailure = true;
      toast(msg, 'warning');
    }
  }
  function toast(msg,type='success') {
    if (type !== 'warning' && Date.now() - (window.__saveFailedAt || 0) < 400) return;
    const el=document.createElement('div');el.className='toast '+type;el.setAttribute('role','status');el.textContent=msg;$('#toast-root').append(el);setTimeout(()=>el.remove(),5500);
  }
  function bindShell() {
    $$('[data-view]').forEach(b=>b.onclick=()=>{view=b.dataset.view;query='';page=1;filter='all';render();});
    $('#backup-btn').onclick=()=>{OrderExcel.downloadBlob(new Blob([JSON.stringify({...state,backupAt:new Date().toISOString()},null,2)],{type:'application/json'}),`序单备份_${stamp()}.json`);toast('备份已下载，包含本机商品映射和订单数据。');};
    $('#restore-btn').onclick=()=>$('#restore-input').click();
    $('#restore-input').onchange=async e=>{try{const f=e.target.files[0];if(!f)return;if(f.size>40*1024*1024)throw Error('备份超过40MB。');const v=JSON.parse(await f.text());validateBackup(v);confirmDialog('恢复数据备份','恢复后将替换当前浏览器的数据。建议先点击右上角“备份数据”。',async()=>{state=v;invalidate();await save();render();toast('已恢复备份，数据可直接用于导出。');});}catch(err){toast(err.message,'warning');}e.target.value='';};
    document.addEventListener('keydown',e=>{if(e.key==='Escape')closeModal();if(e.key==='Tab'&&$('#modal-root .modal')){const a=$$('button,input,select,textarea,[tabindex="0"]',$('#modal-root')).filter(x=>!x.disabled);if(!a.length)return;const first=a[0],last=a[a.length-1];if(e.shiftKey&&document.activeElement===first){last.focus();e.preventDefault();}else if(!e.shiftKey&&document.activeElement===last){first.focus();e.preventDefault();}}});
  }
  function validateBackup(v){
    if(!Array.isArray(v.customers))v.customers=[];
    if(!Array.isArray(v.recharges))v.recharges=[];
    if(v.version!==1||!Array.isArray(v.suppliers)||!Array.isArray(v.products)||!Array.isArray(v.orders)||!Array.isArray(v.batches)||!v.profiles||typeof v.profiles!=='object')throw Error('这不是有效的序单数据备份。');
    if(v.orders.length>30000||v.products.length>30000)throw Error('备份记录过多（最多30,000行）。');
    const ids=a=>a.every(x=>x&&typeof x.id==='string')&&new Set(a.map(x=>x.id)).size===a.length;
    if(!ids(v.suppliers)||!ids(v.products)||!ids(v.orders)||!ids(v.batches)||v.products.some(p=>!v.suppliers.some(s=>s.id===p.supplierId)||!Array.isArray(p.aliases))||v.suppliers.some(s=>s.template&&(!Array.isArray(s.template.columns)||s.template.columns.length>100)))throw Error('备份记录结构或关联关系不完整。');
    v.history=Array.isArray(v.history)?v.history:[];
    if(v.activeBatchIds!==undefined&&(!Array.isArray(v.activeBatchIds)||v.activeBatchIds.some(id=>typeof id!=='string'||!v.batches.some(b=>b.id===id))))throw Error('备份中的本次处理批次无效。');
    migrateScope(v);
  }
  function openModal(title,body,footer='',wide=false) {
    modalReturn=document.activeElement;
    $('#modal-root').innerHTML=`<div class="modal-backdrop"><section class="modal ${wide?'wide':''}" role="dialog" aria-modal="true" aria-labelledby="dialog-title"><div class="modal-header"><h2 id="dialog-title">${esc(title)}</h2><button class="icon-btn" data-close aria-label="关闭对话框">×</button></div><div class="modal-body">${body}</div>${footer?`<div class="modal-footer">${footer}</div>`:''}</section></div>`;
    $('[data-close]').onclick=closeModal;
    $$('.modal [data-cancel]').forEach(b=>b.onclick=closeModal);
    setTimeout(()=>$('input:not([type=hidden]),select,textarea,button',$('#modal-root .modal'))?.focus(),0);
  }
  function closeModal(){$('#modal-root').innerHTML='';modalReturn?.focus?.();}
  function confirmDialog(title,text,action){openModal(title,`<p>${esc(text)}</p>`,`<button class="btn secondary" data-cancel>取消</button><button class="btn ${/删除|清空/.test(title)?'danger':''}" id="confirm-action">${/删除/.test(title)?'确认删除':'确认'}</button>`);$('#confirm-action').onclick=async()=>{closeModal();await action();};}
  function field(label,name,value='',options={}){return `<label class="field ${options.full?'full':''}"><span>${esc(label)}${options.required?' <i class="req" aria-hidden="true">*</i>':''}</span>${options.textarea?`<textarea name="${name}" rows="${options.rows||3}" ${options.required?'required':''}>${esc(value)}</textarea>`:`<input name="${name}" value="${esc(value)}" ${options.required?'required':''} ${options.type?`type="${options.type}"`:''} ${options.placeholder?`placeholder="${esc(options.placeholder)}"`:''}>`}${options.hint?`<small>${esc(options.hint)}</small>`:''}</label>`;}
  function selectField(label,name,options,value='',full=false){return `<label class="field ${full?'full':''}"><span>${esc(label)}</span><select name="${name}">${options.map(([v,l])=>`<option value="${esc(v)}" ${String(v)===String(value)?'selected':''}>${esc(l)}</option>`).join('')}</select></label>`;}
  function findProducts(row) {
    const name=norm(row.productName),code=norm(row.sourceSku),spec=norm(row.sourceSpec);
    return state.products.filter(p=>{
      if(name && norm(p.name)===name && (!spec||!p.spec||!/(\d+)枚/.test(spec)||norm(p.spec)===spec))return true;
      return p.aliases.some(a=>(!a.customer||norm(a.customer)===norm(row.customer))&&(!a.spec||norm(a.spec)===spec)&&((a.code&&code&&norm(a.code)===code)||(!a.code&&a.name&&norm(a.name)===name)));
    });
  }
  function suggestion(row){
    const c=OrderSeed.candidates.find(c=>norm(c.name)===norm(row.productName)&&(!c.sourceSku||norm(c.sourceSku)===norm(row.sourceSku)));
    if(c&&pById(c.productId))return pById(c.productId);
    return null;
  }
  function status(row) {
    if(row.excluded)return {key:'excluded',label:'已排除',type:'neutral',issues:[]};
    if(row.duplicateOf)return {key:'duplicate',label:'重复 / 修订待确认',type:'warning',issues:['与已有订单明细重复，请选择保留版本']};
    const issues=[];
    if(!row.productName)issues.push('缺少商品名称');
    if(!Number.isFinite(Number(row.quantity))||Number(row.quantity)<=0)issues.push('数量必须大于0');
    if(!row.recipient)issues.push('缺少收货人');
    if(!row.phone)issues.push('缺少电话');
    if(!row.address)issues.push('缺少收货地址');
    let products=row.productId&&row.manual?[pById(row.productId)].filter(Boolean):findProducts(row);
    if(products.length>1)issues.push('多个商品匹配，请人工选择');
    if(!products.length)issues.push(row.productId?'映射商品已失效':'尚未关联商品与供应商');
    const product=products.length===1?products[0]:null;
    if(product&&!sById(product.supplierId))issues.push('供应商已失效');
    if(product&&!String(product.supplierSku||'').trim())issues.push('缺少供应商商品编码');
    if(issues.length)return {key:product?'invalid':'unmatched',label:product?'字段待完善':products.length>1?'映射冲突':'待确认映射',type:'warning',issues,product};
    return {key:'ready',label:'匹配完成',type:'success',issues:[],product};
  }
  const summary=()=>activeRows().reduce((a,r)=>{const s=status(r);a[s.key]=(a[s.key]||0)+1;a.total++;if(s.key!=='excluded')a.quantity+=Number(r.quantity)||0;return a;},{total:0,ready:0,unmatched:0,invalid:0,duplicate:0,excluded:0,quantity:0});
  function badge(s){return `<span class="pill ${s.type}" title="${esc(s.issues?.join('；')||'')}">${s.type==='success'?icon('check'):''}${esc(s.label)}</span>`;}
  function workflow(){return `<div class="workflow">${[['suppliers','供应商库','建立产品与供应商关系'],['customers','客户库','给订单的客户'],['import','订单采集与匹配','识别不同客户格式'],['export','供应商格式','一家供应商一套格式'],['shipment','运单回填','补上平台订单号'],['recon','对账','按客户核对订单与金额']].map(([id,t,desc],i)=>`<button class="workflow-step ${view===id?'active':''}" data-go="${id}"><span class="step-number">0${i+1}</span><span><strong>${t}</strong><small>${desc}</small></span>${i<5?icon('arrow','step-arrow'):''}</button>`).join('')}</div>`;}
  function stats(){const s=summary();return `<div class="stats-grid">${[['file','已采集订单',s.total,`本次 ${state.activeBatchIds.length} 个批次 · 历史 ${state.batches.length} 个`,'blue'],['check','匹配完成',s.ready,s.total?`${Math.round(s.ready/s.total*100)}% 的明细已就绪`:'导入订单后自动检查','green'],['clock','待处理',s.unmatched+s.invalid+s.duplicate,'映射与字段异常','amber'],['box','供应商库',state.products.length,`${state.suppliers.length} 家供应商`,'purple']].map(([ic,l,v,n,c])=>`<article class="stat-card"><div class="stat-icon ${c}">${icon(ic)}</div><div><div class="stat-label">${l}</div><div class="stat-value">${v}<span>条</span></div><div class="stat-note">${n}</div></div></article>`).join('')}</div>`;}
  function render(){
    const titles={suppliers:['供应商库','维护供应商、标准商品与客户商品别名'],import:['订单采集与匹配','把不同客户的订单，整理成可交付的数据'],export:['供应商格式','一家供应商一套格式，匹配后直接导出'],customers:['客户库','维护给我们下订单的客户'],shipment:['运单回填','把快递单号表的订单号补上'],recon:['对账','按客户汇总订单、商品与金额']};
    $('#page-title').textContent=titles[view][0];$('#page-subtitle').textContent=titles[view][1];
    $$('[data-view]').forEach(b=>{b.classList.toggle('active',b.dataset.view===view);b.setAttribute('aria-current',b.dataset.view===view?'page':'false');});
    const localBanner = (OrderStore.mode==='cloudbase' && state.orders.length===0 && state.customers.length===0 && !window.__localImportDone) ? `<div class="alert warning import-local-bar"><span>检测到这是空的云端工作区。如果你之前在这台电脑上用过单机版，可以把本机数据导入云端。</span><button class="btn secondary small" id="import-local">导入本机数据</button></div>` : '';
    $('#content').innerHTML=localBanner+workflow()+(view==='import'?renderImport():view==='suppliers'?renderLibrary():view==='customers'?renderCustomers():view==='shipment'?renderShipment():view==='recon'?renderRecon():renderExport());
    $$('[data-go]').forEach(b=>b.onclick=()=>{view=b.dataset.go;query='';page=1;render();});
    bindView();
    if (view === 'recon') maybeShowBalanceNotice();
  }
  function renderImport(){const s=summary();return `<div class="page-intro"><div><div class="eyebrow">DAILY ORDER WORKSPACE</div><h1>今天的订单，从这里开始</h1><p>批量导入客户表格，确认商品映射，再按供应商格式导出。</p></div><div class="intro-actions"><button class="btn secondary" id="demo-btn">体验示例</button><button class="btn" data-go="export" ${!s.ready?'disabled':''}>按供应商导出 ${icon('arrow')}</button></div></div>${stats()}
      <div class="workspace-grid"><section class="card"><div class="card-header"><h2>${icon('upload')} 导入客户订单</h2><span class="pill neutral">支持批量导入</span></div><div class="card-body"><div class="dropzone" id="dropzone" tabindex="0" role="button" aria-label="选择或拖入客户订单文件"><div class="upload-icon">${icon('upload')}</div><h3>拖入订单表格，或 <span class="text-link">选择文件</span></h3><p>支持 .xlsx / UTF-8 CSV，单文件不超过 15 MB</p><span class="drop-hint">自动识别表头 · 支持手动调整字段</span></div><input id="order-files" type="file" accept=".xlsx,.csv" multiple hidden></div></section>
      <aside class="card quick-guide"><div class="card-header"><h2>已适配你的样表</h2><span class="pill success">2 种订单格式</span></div><div class="card-body"><div class="profile-row">${icon('file')}<div><strong>兮迪订单</strong><small>订单编号 / 购买数量 / 收货人</small></div><span class="profile-mark">29 列</span></div><div class="profile-row">${icon('file')}<div><strong>GBH 订单</strong><small>订单号 / SKU码 / 数量</small></div><span class="profile-mark">18 列</span></div><div class="guide-note">${icon('shield')}<p>文件在本浏览器中处理。真实订单不会被预置到演示页面。</p></div></div></aside></div>
      <section class="card order-panel"><div class="card-header"><div><h2>订单明细 <span class="count">${s.total}</span></h2><p class="muted">本次 ${state.activeBatchIds.length} 个批次 · ${s.total?`共 ${s.quantity} 件，${s.excluded} 条已排除。未处理明细会阻止生成。`:'导入后可逐行核对匹配结果和来源。'}</p></div>${state.orders.length?`<button class="btn ghost small" id="export-unified">导出统一订单</button><button class="btn ghost small danger-ghost" id="clear-orders">清空订单</button>`:''}${state.batches.length?'<button class="btn ghost small" id="batches-btn">选择本次批次 / 历史</button>':''}</div><div class="toolbar"><div class="tabs">${[['all','全部',s.total],['ready','已匹配',s.ready],['pending','待处理',s.unmatched+s.invalid+s.duplicate],['excluded','已排除',s.excluded]].map(([k,t,n])=>`<button class="tab ${filter===k?'active':''}" data-filter="${k}">${t}<span>${n}</span></button>`).join('')}</div><label class="search-wrap">${icon('search')}<input class="search-input" id="order-search" value="${esc(query)}" placeholder="搜索订单号、商品或客户" aria-label="搜索订单"></label></div><div id="orders-table">${orderTable()}</div></section>
      <p class="page-footnote">${icon('link')} 商品别名确认后可复用；不使用订单明细 ID 作为商品编码。</p>`;}
  function visibleOrders(){return activeRows().filter(r=>{const s=status(r);return (filter==='all'||filter==='pending'&&!['ready','excluded'].includes(s.key)||filter===s.key)&&(!query||norm([r.orderNo,r.customer,r.productName,r.recipient].join(' ')).includes(norm(query)));});}
  function orderTable() {
    const rows = visibleOrders();
    const pages = Math.max(1, Math.ceil(rows.length / 30));
    if (page > pages) page = pages;
    const slice = rows.slice((page - 1) * 30, page * 30);
    const offset = (page - 1) * 30;
    const cell = v => v === '' || v == null ? '<span class="muted">—</span>' : esc(v);
    return slice.length ? `<div class="table-wrap"><table class="data-table"><thead><tr><th>序号</th><th>供应商名称</th><th>订单号</th><th>商品名称</th><th>69码</th><th>数量</th><th>收件人</th><th>电话</th><th>详细地址</th><th>物流公司</th><th>快递单号</th><th>状态</th><th class="sticky-action">操作</th></tr></thead><tbody>${slice.map((r, i) => {
      const s = status(r), p = s.product, c = suggestion(r);
      const supplier = p ? (sById(p.supplierId)?.name || '') : '';
      return `<tr class="${s.key === 'excluded' ? 'row-excluded' : ''}"><td class="numeric">${offset + i + 1}</td><td>${cell(supplier)}</td><td class="mono">${cell(r.orderNo)}<small>${esc(r.customer)} · ${esc(r.sourceSheet)} 第 ${r.sourceRow} 行</small></td><td class="product-cell">${esc(p ? p.name : (r.productName || '未填写'))}<small>${esc(r.sourceSpec || r.sourceSku || '')}${p ? '' : c ? ' · 建议：' + esc(c.name) : ''}</small></td><td class="mono">${cell(p ? p.barcode : '')}</td><td class="numeric">${esc(r.quantity)}</td><td>${cell(r.recipient)}</td><td class="mono">${cell(r.phone)}</td><td>${cell(r.address)}</td><td>${cell(r.carrier)}</td><td class="mono">${cell(r.trackingNo)}</td><td>${badge(s)}</td><td class="sticky-action"><button class="text-link" data-edit-order="${esc(r.id)}">${s.key === 'ready' ? '查看 / 编辑' : '处理'}</button> <button class="text-link danger" data-del-order="${esc(r.id)}">删除</button></td></tr>`;
    }).join('')}</tbody></table></div>${pagination(rows.length)}` : `<div class="empty-state">${icon('file')}<h3>${activeRows().length ? '没有符合条件的明细' : '本次还没有订单'}</h3><p>${activeRows().length ? '尝试调整筛选条件或搜索内容。' : '导入新订单，或在“选择本次批次 / 历史”中勾选历史批次。'}</p></div>`;
  }
  async function exportUnifiedOrders() {
    const rows = activeRows().filter(r => !r.excluded);
    if (!rows.length) { toast('本次没有可导出的订单。', 'warning'); return; }
    const headers = ['序号', '供应商名称', '订单号', '商品名称', '69码', '数量', '收件人', '电话', '详细地址', '物流公司', '快递单号'];
    const data = rows.map((r, i) => {
      const st = status(r), p = st.product;
      return [i + 1, p ? (sById(p.supplierId)?.name || '') : '', r.orderNo || '', p ? p.name : (r.productName || ''), p ? (p.barcode || '') : '', Number(r.quantity) || r.quantity, r.recipient || '', r.phone || '', r.address || '', r.carrier || '', r.trackingNo || ''];
    });
    await OrderExcel.downloadWorkbook({ headers, rows: data, sheetName: '统一订单', filename: `统一订单_${stamp()}.xlsx` });
    toast(`已导出 ${data.length} 条统一订单。`);
  }
  function pagination(total){return `<div class="pagination"><span>共 ${total} 条 · 每页 30 条</span><div><button class="btn ghost small" data-page="${page-1}" ${page<=1?'disabled':''}>上一页</button><span>${page} / ${Math.max(1,Math.ceil(total/30))}</span><button class="btn ghost small" data-page="${page+1}" ${page*30>=total?'disabled':''}>下一页</button></div></div>`;}
  function supplierMetrics() {
    const m = new Map(state.suppliers.map(s => [s.id, { count: 0, brands: new Map(), min: null, max: null }]));
    for (const p of state.products) {
      const st = m.get(p.supplierId); if (!st) continue;
      st.count++;
      if (p.brand) st.brands.set(p.brand, (st.brands.get(p.brand) || 0) + 1);
      const c = Number(p.cost);
      if (Number.isFinite(c) && c > 0) { st.min = st.min == null ? c : Math.min(st.min, c); st.max = st.max == null ? c : Math.max(st.max, c); }
    }
    return m;
  }
  function supplierRail() {
    const metrics = supplierMetrics();
    const tiles = [{ id: 'all', name: '全部供应商', count: state.products.length }, ...state.suppliers.map(s => ({ id: s.id, name: s.name, derived: s.derived, metrics: metrics.get(s.id), count: metrics.get(s.id)?.count || 0 }))];
    return `<div class="supplier-rail" role="tablist" aria-label="按供应商查看商品">${tiles.map(t => {
      const active = supplierFilter === t.id;
      const brands = t.metrics ? [...t.metrics.brands.entries()].sort((a, b) => b[1] - a[1]).slice(0, 3).map(([b]) => b) : [];
      const range = t.metrics && t.metrics.min != null ? `供价 ¥${t.metrics.min}${t.metrics.max > t.metrics.min ? '–' + t.metrics.max : ''}` : (t.id === 'all' ? `${state.suppliers.length} 家供应商` : '暂无供价');
      return `<article class="supplier-tile ${active ? 'active' : ''}" data-supplier-tile="${esc(t.id)}" role="tab" tabindex="0" aria-selected="${active}" title="点击查看该供应商的商品"><span class="tile-avatar">${esc(t.name.slice(0, 1))}</span><span class="tile-body"><strong>${esc(t.name)}</strong><small>${t.count} 个商品 · ${esc(range)}</small>${brands.length ? `<span class="tile-brands">${brands.map(b => `<i>${esc(b)}</i>`).join('')}</span>` : ''}</span>${t.derived ? '<span class="tile-flag">品牌推断</span>' : ''}${t.id !== 'all' ? `<button type="button" class="tile-edit" data-edit-supplier="${esc(t.id)}">编辑</button>` : ''}</article>`;
    }).join('')}</div>`;
  }
  let customerQuery = '';
  const CUSTOMER_FIELDS = [
    ['name', '公司简称', { hint: '订单上显示的名字，例如「成都丰涵」' }],
    ['companyName', '公司名称', {}],
    ['creditCode', '社会信用代码', {}],
    ['companyAddress', '公司地址', { full: true }],
    ['contact', '联系人姓名', {}],
    ['phone', '联系人电话', {}],
    ['bankName', '开户银行', {}],
    ['bankBranch', '所在支行', {}],
    ['bankCode', '银行联行号', {}],
    ['bankAccount', '银行账号', {}],
    ['license', '营业执照', {}],
    ['businessScope', '经营范围', { textarea: true, full: true }]
  ];
  const customerByName = name => state.customers.find(c => norm(c.name) === norm(name));
  function ensureCustomer(name) {
    const n = String(name || '').trim();
    if (!n) return null;
    let c = customerByName(n);
    if (!c) {
      c = { id: uid(), createdAt: new Date().toISOString(), note: '', deposit: '' };
      for (const [k] of CUSTOMER_FIELDS) c[k] = '';
      c.name = n;
      state.customers.push(c);
    }
    return c;
  }
  function customerOrderCount(name) {
    const key = norm(name);
    return state.orders.filter(o => norm(o.customer) === key).length;
  }
  function customerModal(id) {
    const editing = !!id;
    const base = state.customers.find(x => x.id === id) || { id: uid(), createdAt: new Date().toISOString(), note: '', deposit: '' };
    const c = copy(base);
    for (const [k] of CUSTOMER_FIELDS) if (c[k] === undefined) c[k] = '';
    openModal(editing ? '编辑客户' : '新增客户', `<form id="customer-form"><div class="form-grid">${CUSTOMER_FIELDS.map(([k, label, opts]) => field(label, k, c[k] || '', { ...opts, required: k === 'name' })).join('')}${field('备注', 'note', c.note, { textarea: true, full: true })}${editing ? `<div class="field full"><span>总充值金额</span><div class="read-count">¥${reconMoney(reconDeposit(c.name))} · 在对账页充值与查看记录</div></div>` : ''}</div><p class="form-error" id="customer-error"></p></form>`, `<button class="btn secondary" data-cancel>取消</button>${editing ? '<button class="btn danger" id="delete-customer">删除客户</button>' : ''}<button class="btn" id="save-customer">保存客户</button>`, true);
    $('#save-customer').onclick = async () => {
      const f = new FormData($('#customer-form'));
      for (const [k] of CUSTOMER_FIELDS) c[k] = String(f.get(k) || '').trim();
      c.note = String(f.get('note') || '').trim();
      if (!c.name) { $('#customer-error').textContent = '请填写公司简称。'; return; }
      if (state.customers.some(x => x.id !== c.id && norm(x.name) === norm(c.name))) { $('#customer-error').textContent = '公司简称已存在。'; return; }
      const idx = state.customers.findIndex(x => x.id === c.id);
      idx < 0 ? state.customers.push(c) : state.customers.splice(idx, 1, c);
      await save(); closeModal(); render(); toast('客户已保存。');
    };
    $('#delete-customer')?.addEventListener('click', () => {
      const used = customerOrderCount(c.name);
      confirmDialog('删除客户', used ? `「${c.name}」还有 ${used} 条订单记录，删除客户不会删除订单，订单上的客户名称会保留。` : `确认删除「${c.name}」？`, async () => {
        state.customers = state.customers.filter(x => x.id !== c.id);
        state.recharges = state.recharges.filter(r => norm(r.customer) !== norm(c.name));
        await save(); render(); toast('客户已删除。');
      });
    });
  }
  function renderCustomers() {
    const q = norm(customerQuery);
    const rows = state.customers
      .filter(c => !q || norm(CUSTOMER_FIELDS.map(([k]) => c[k]).concat([c.note]).join(' ')).includes(q))
      .map(c => ({ ...c, orders: customerOrderCount(c.name) }))
      .sort((a, b) => b.orders - a.orders || a.name.localeCompare(b.name, 'zh-CN'));
    const cell = v => v === '' || v == null ? '<span class="muted">—</span>' : esc(v);
    return `<div class="page-intro"><div><div class="eyebrow">CUSTOMER LIBRARY</div><h1>谁给我们下的单</h1><p>按「客户信息表」的口径维护客户主体、联系人与银行资料；导入订单时按公司简称自动挂上。</p></div><div class="intro-actions"><button class="btn secondary" id="export-customers">${icon('download')} 导出客户信息表</button><button class="btn secondary" id="import-customers">${icon('upload')} 导入客户信息表</button><input type="file" id="customer-file" accept=".xlsx,.csv" hidden><button class="btn" id="new-customer">${icon('plus')} 新增客户</button></div></div>
      ${rows.length ? `<div class="stats-grid">${[['box', '客户数', state.customers.length, '已登记客户', '个'], ['file', '有订单的客户', state.customers.filter(c => customerOrderCount(c.name) > 0).length, '至少下过一单', '个'], ['check', '订单总数', state.orders.length, '含历史批次', '条']].map(([ic, l, v, n, u]) => `<article class="stat-card"><div class="stat-icon blue">${icon(ic)}</div><div><div class="stat-label">${l}</div><div class="stat-value">${v}<span>${u}</span></div><div class="stat-note">${n}</div></div></article>`).join('')}</div>` : ''}
      <section class="card"><div class="card-header"><div><h2>客户库 <span class="count">${state.customers.length}</span></h2><p class="muted">列与「客户信息表」表头一致：简称、公司名称、信用代码、联系人、电话、订单数；其余字段在编辑里。</p></div><div class="intro-actions"><button class="btn ghost small" id="sync-customers">从历史订单补登记</button></div></div>
      <div class="toolbar toolbar-end"><label class="search-wrap">${icon('search')}<input id="customer-search" class="search-input" value="${esc(customerQuery)}" placeholder="搜索简称、公司、信用代码、联系人" aria-label="搜索客户"></label></div>
      ${rows.length ? `<div class="table-wrap"><table class="data-table pool-table"><thead><tr><th>供应商名称（公司简称）</th><th>公司名称</th><th>社会信用代码</th><th>公司地址</th><th>联系人姓名</th><th>联系人电话</th><th>开户银行</th><th>所在支行</th><th>银行联行号</th><th>银行账号</th><th>营业执照</th><th>经营范围</th><th>订单数</th><th>操作</th></tr></thead><tbody>${rows.map(c => `<tr><td class="product-cell"><strong>${esc(c.name)}</strong><small>${esc((c.createdAt || '').slice(0, 10))}</small></td><td>${cell(c.companyName)}</td><td class="mono">${cell(c.creditCode)}</td><td>${cell(c.companyAddress)}</td><td>${cell(c.contact)}</td><td class="mono">${cell(c.phone)}</td><td>${cell(c.bankName)}</td><td>${cell(c.bankBranch)}</td><td class="mono">${cell(c.bankCode)}</td><td class="mono">${cell(c.bankAccount)}</td><td>${cell(c.license)}</td><td>${cell(c.businessScope)}</td><td class="numeric">${c.orders}</td><td><button class="text-link" data-edit-customer="${esc(c.id)}">编辑</button></td></tr>`).join('')}</tbody></table></div>` : `<div class="empty-state">${icon('box')}<h3>${state.customers.length ? '没有匹配的客户' : '还没有客户'}</h3><p>${state.customers.length ? '换个关键词试试。' : '导入订单时会按客户名自动登记；也可以点右上角「新增客户」先建好。'}</p></div>`}</section>`;
  }
  async function importCustomers(file) {
    if (!file) return;
    await guardAsync(async () => {
      const book = await OrderExcel.read(file);
      let si = book.sheets.findIndex(s => s.rows.some(r => (r || []).some(c => poolNorm(c).includes('公司'))));
      if (si < 0) si = book.sheets.findIndex(s => s.rows.length > 2);
      if (si < 0) si = 0;
      const rows = book.sheets[si].rows;
      const header = OrderExcel.detectHeader(rows);
      const headers = rows[header] || [];
      const map = {};
      for (const [k, label] of CUSTOMER_FIELDS) {
        const target = poolNorm(label);
        let idx = headers.findIndex(h => poolNorm(h) === target);
        if (idx < 0) idx = headers.findIndex(h => poolNorm(h) && (poolNorm(h).includes(target) || target.includes(poolNorm(h))));
        map[k] = idx;
      }
      if (map.name < 0) { toast('没有找到「公司简称」列，请确认使用客户信息表模板。', 'warning'); return; }
      const get = (r, k) => { const i = map[k]; if (i < 0) return ''; const v = r[i]; return v == null ? '' : String(v).trim(); };
      const items = [];
      for (let i = header + 1; i < rows.length; i++) {
        const r = rows[i] || [];
        if (!get(r, 'name')) continue;
        items.push({ row: i + 1, raw: r, name: get(r, 'name') });
      }
      if (!items.length) { toast('表里没有可导入的客户行。', 'warning'); return; }
      const updates = items.filter(it => customerByName(it.name)).length;
      openModal('导入客户信息表', `
        <div class="alert info">${icon('file')}<span>已读取 <strong>${esc(file.name)}</strong> · 识别 <strong>${items.length}</strong> 个客户。</span></div>
        <div class="stats-grid compact">${[['新增客户', items.length - updates, '公司简称不存在'], ['更新客户', updates, '按公司简称匹配，覆盖非空字段']].map(([l, v, n]) => `<article class="stat-card"><div class="stat-label">${l}</div><div class="stat-value">${v}<span>个</span></div><div class="stat-note">${esc(n)}</div></article>`).join('')}</div>
        <p class="muted">按「公司简称」匹配；空单元格不会覆盖已有值。</p>
        <div class="table-wrap"><table class="data-table"><thead><tr><th>行</th><th>公司简称</th><th>公司名称</th><th>联系人</th><th>电话</th></tr></thead><tbody>${items.slice(0, 6).map(it => `<tr><td class="numeric">${it.row}</td><td>${esc(it.name)}</td><td>${esc(get(it.raw, 'companyName') || '—')}</td><td>${esc(get(it.raw, 'contact') || '—')}</td><td class="mono">${esc(get(it.raw, 'phone') || '—')}</td></tr>`).join('')}</tbody></table></div>
      `, `<button class="btn secondary" data-cancel>取消</button><button class="btn" id="customer-import-confirm">确认导入</button>`, true);
      $('#customer-import-confirm').onclick = async () => {
        let added = 0, updated = 0;
        for (const it of items) {
          const r = it.raw;
          let c = customerByName(it.name);
          if (!c) { c = { id: uid(), createdAt: new Date().toISOString(), note: '', deposit: '' }; for (const [k] of CUSTOMER_FIELDS) c[k] = ''; c.name = it.name; state.customers.push(c); added++; } else updated++;
          for (const [k] of CUSTOMER_FIELDS) {
            const v = get(r, k);
            if (v !== '') c[k] = v;
          }
        }
        await save(); closeModal(); render();
        toast(`客户信息表导入完成：新增 ${added} 个、更新 ${updated} 个。`);
      };
    });
  }
  async function exportCustomers() {
    const headers = ['供应商名称（公司简称）', '公司名称', '社会信用代码', '公司地址', '联系人姓名', '联系人电话', '开户银行', '所在支行', '银行联行号', '银行账号', '营业执照', '经营范围'];
    const rows = state.customers.map(c => CUSTOMER_FIELDS.map(([k]) => c[k] ?? ''));
    if (!rows.length) { toast('还没有客户可导出。', 'warning'); return; }
    await OrderExcel.downloadWorkbook({ headers, rows, sheetName: '客户信息表', filename: `客户信息表_${stamp()}.xlsx` });
    toast(`已导出 ${rows.length} 个客户。`);
  }
  function bindCustomers() {
    $('#new-customer')?.addEventListener('click', () => customerModal());
    $$('[data-edit-customer]').forEach(b => b.onclick = () => customerModal(b.dataset.editCustomer));
    $('#export-customers')?.addEventListener('click', () => guardAsync(exportCustomers));
    $('#import-customers')?.addEventListener('click', () => $('#customer-file').click());
    $('#customer-file')?.addEventListener('change', e => { const fl = e.target.files && e.target.files[0]; e.target.value = ''; if (fl) importCustomers(fl); });
    $('#customer-search')?.addEventListener('input', e => { customerQuery = e.target.value; render(); const el = $('#customer-search'); if (el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length); } });
    $('#sync-customers')?.addEventListener('click', async () => {
      let added = 0;
      const before = state.customers.length;
      for (const o of state.orders) if (String(o.customer || '').trim() && !customerByName(o.customer)) ensureCustomer(o.customer);
      added = state.customers.length - before;
      if (added) { await save(); render(); toast(`已从历史订单补登记 ${added} 个客户。`); }
      else toast('历史订单里的客户都已登记。');
    });
  }
  let reconGroup = '';
  let reconScope = 'all';
  let reconMode = 'customer';
  function reconBase() {
    return (reconScope === 'active' ? activeRows() : state.orders).filter(o => !o.excluded);
  }
  function reconLine(o, mode) {
    const st = status(o), p = st.product || null;
    const qty = Number(o.quantity) || 0;
    const customer = String(o.customer || '').trim() || '未命名客户';
    const supplier = p ? (sById(p.supplierId)?.name || '') : '';
    let price = null;
    if (mode === 'supplier') {
      const cost = p && p.cost !== '' && p.cost != null ? Number(p.cost) : NaN;
      price = Number.isFinite(cost) ? cost : null;
    } else {
      const raw = o.unitPrice === '' || o.unitPrice == null ? '' : String(o.unitPrice).trim();
      price = raw === '' ? null : (Number(raw) || 0);
    }
    return {
      customer, supplier: supplier || '未分配供应商', supplierKey: supplier || '未分配供应商',
      orderNo: String(o.orderNo || ''), date: String(o.date || ''),
      product: p ? p.name : String(o.productName || ''), spec: p ? (p.spec || '') : String(o.sourceSpec || ''),
      sku: p ? (p.supplierSku || '') : '', barcode: p ? (p.barcode || '') : '', orderId: o.id, invoiced: !!o.invoiced, qty, price,
      carrier: o.carrier || '', trackingNo: o.trackingNo || '', remark: o.remark || '', freight: Number(o.shippingFee) || 0, taxCode: p ? (p.taxCode || '') : '', taxRate: p ? (p.taxRate || '') : '',
      amount: price == null ? null : Math.round(price * qty * 100) / 100,
      recipient: o.recipient || '', phone: o.phone || '', address: o.address || '', status: st.label
    };
  }
  const rechargeKind = r => r.kind || 'customer';
  const rechargeName = r => r.name || r.customer || '';
  const rechargeLabel = kind => kind === 'supplier' ? '供应商' : '客户';
  function rechargeRecords(name, kind = 'customer') {
    const key = norm(name);
    return state.recharges
      .filter(r => rechargeKind(r) === kind && norm(rechargeName(r)) === key)
      .sort((a, b) => String(b.date).localeCompare(String(a.date)) || String(b.createdAt).localeCompare(String(a.createdAt)));
  }
  function reconDeposit(name, kind = 'customer') {
    return rechargeRecords(name, kind).reduce((n, r) => n + (Number(r.amount) || 0), 0);
  }
  function partyExists(kind, name) {
    return kind === 'supplier'
      ? state.suppliers.some(s => norm(s.name) === norm(name))
      : !!customerByName(name);
  }
  function rechargeModal(name, kind = 'customer') {
    if (!partyExists(kind, name)) { toast(`先选择${rechargeLabel(kind)}。`, 'warning'); return; }
    const hint = kind === 'supplier' ? '本次实际付给该供应商的金额' : '本次实际到账金额';
    openModal(`充值 · ${name}`, `<form id="recharge-form"><div class="form-grid">${field('充值金额', 'amount', '', { type: 'number', required: true, hint })}${field('充值日期', 'date', stamp(), { type: 'date', required: true, hint: '默认今天，可改' })}${field('备注', 'note', '', { textarea: true, full: true })}</div><p class="form-error" id="recharge-error"></p></form>`, `<button class="btn secondary" data-cancel>取消</button><button class="btn" id="save-recharge">保存充值</button>`);
    $('#save-recharge').onclick = async () => {
      const form = new FormData($('#recharge-form'));
      const amount = Number(form.get('amount'));
      const date = String(form.get('date') || '').trim();
      if (!Number.isFinite(amount) || amount <= 0) { $('#recharge-error').textContent = '请输入大于 0 的充值金额。'; return; }
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) { $('#recharge-error').textContent = '请选择充值日期。'; return; }
      state.recharges.push({ id: uid(), kind, name, customer: kind === 'customer' ? name : '', amount: Math.round(amount * 100) / 100, date, note: String(form.get('note') || '').trim(), createdAt: new Date().toISOString() });
      await save(); closeModal(); render();
      toast(`已记录「${name}」充值 ¥${reconMoney(amount)}，总充值 ¥${reconMoney(reconDeposit(name, kind))}。`);
    };
  }
  function rechargeListModal(name, kind = 'customer') {
    const list = rechargeRecords(name, kind);
    const total = list.reduce((n, r) => n + (Number(r.amount) || 0), 0);
    openModal(`充值记录 · ${name}`, `<p class="muted">共 ${list.length} 笔，累计 ¥${reconMoney(total)}。</p>${list.length ? `<div class="table-wrap"><table class="data-table"><thead><tr><th>日期</th><th>金额</th><th>备注</th><th>操作</th></tr></thead><tbody>${list.map(r => `<tr><td>${esc(r.date)}</td><td class="numeric">¥${reconMoney(r.amount)}</td><td>${esc(r.note || '—')}</td><td><button class="text-link danger" data-del-recharge="${esc(r.id)}">删除</button></td></tr>`).join('')}</tbody></table></div>` : '<div class="empty-state"><h3>还没有充值记录</h3><p>点下方「新增充值」记录第一笔。</p></div>'}`, `<button class="btn secondary" data-cancel>关闭</button><button class="btn" id="add-recharge">新增充值</button>`, true);
    $('#add-recharge').onclick = () => { closeModal(); rechargeModal(name, kind); };
    $$('[data-del-recharge]').forEach(b => b.onclick = () => {
      const rec = state.recharges.find(r => r.id === b.dataset.delRecharge);
      confirmDialog('删除充值记录', `确认删除 ${rec.date} 的 ¥${reconMoney(rec.amount)} 充值记录？`, async () => {
        state.recharges = state.recharges.filter(r => r.id !== rec.id);
        await save(); render(); rechargeListModal(name, kind); toast('充值记录已删除。');
      });
    });
  }
  function reconAccount(name, kind = 'customer') {
    let total = 0, noPrice = 0, count = 0;
    for (const o of state.orders) {
      if (o.excluded) continue;
      const line = reconLine(o, kind === 'supplier' ? 'supplier' : 'customer');
      const owner = kind === 'supplier' ? line.supplierKey : line.customer;
      if (owner !== name) continue;
      count++;
      if (line.amount == null) noPrice++;
      else total += line.amount;
    }
    const deposit = reconDeposit(name, kind);
    return { deposit, total: Math.round(total * 100) / 100, balance: Math.round((deposit - total) * 100) / 100, noPrice, count };
  }
  const RECON_BILL_HEADERS = ['订单编号', '收件人', '收件人电话', '收件人地址', '发货状态', '发货时间', '物流公司', '物流单号', '客户名称', '商品名称', '商品单价', '商品数量', '商品金额（商品单价*商品数量）', '订单金额（含运费）', '运费', '退款说明', '售后凭证', '订单状态', '下单时间', '付款时间', '账期类型', '对账状态', '收款状态', '收款人', '收款方式', '收款时间', '税收编码', '税率', '是否开票', '是否 完成对账备注'];
  function reconBillValues(l, mode) {
    const money = v => v == null ? '' : Math.round(v * 100) / 100;
    const freight = Number(l.freight) || 0;
    const total = l.amount == null ? '' : Math.round((l.amount + freight) * 100) / 100;
    return [
      l.orderNo || '', l.recipient, l.phone, l.address,
      l.trackingNo ? '已发货' : '待发货', '', l.carrier, l.trackingNo,
      mode === 'supplier' ? l.supplierKey : l.customer, l.product,
      money(l.price), l.qty, money(l.amount), total, freight || 0,
      '', '', l.trackingNo ? '已发货' : '待发货', l.date, '', '',
      '', '', '', '', '', '', l.taxCode, l.taxRate, l.invoiced ? '是' : '否', l.remark || ''
    ];
  }
  function reconGroups(rows, mode) {
    const map = new Map();
    for (const o of rows) {
      const line = reconLine(o, mode);
      const key = mode === 'supplier' ? line.supplierKey : line.customer;
      if (!map.has(key)) map.set(key, { name: key, count: 0, amount: 0, noPrice: 0 });
      const item = map.get(key);
      item.count++;
      if (line.amount == null) item.noPrice++;
      else item.amount += line.amount;
    }
    const list = [...map.values()].sort((a, b) => b.count - a.count);
    const kind = mode === 'supplier' ? 'supplier' : 'customer';
    for (const g of list) { const acc = reconAccount(g.name, kind); g.deposit = acc.deposit; g.balance = acc.balance; g.total = acc.total; }
    return list;
  }
  const reconMoney = v => v == null ? '' : (Math.round(v * 100) / 100).toFixed(2);
  let reconNoticeShown = false;
  function balanceNoticeItems() {
    const out = [];
    for (const c of state.customers) {
      const acc = reconAccount(c.name, 'customer');
      if ((acc.count > 0 || acc.deposit > 0) && acc.balance <= 0) out.push({ name: c.name, deposit: acc.deposit, total: acc.total, balance: acc.balance });
    }
    return out.sort((a, b) => a.balance - b.balance);
  }
  function maybeShowBalanceNotice() {
    if (reconNoticeShown) return;
    reconNoticeShown = true;
    if (state.reconNoticeDate === stamp()) return;
    const items = balanceNoticeItems();
    if (!items.length) return;
    openModal('客户余额提醒', `<p class="muted">以下客户的账户余额已用完或欠款，建议先充值再继续发货：</p><div class="table-wrap"><table class="data-table"><thead><tr><th>客户</th><th>总充值</th><th>应收</th><th>余额</th></tr></thead><tbody>${items.map(i => `<tr><td>${esc(i.name)}</td><td class="numeric">¥${reconMoney(i.deposit)}</td><td class="numeric">¥${reconMoney(i.total)}</td><td class="numeric"><strong class="neg">¥${reconMoney(i.balance)}</strong></td></tr>`).join('')}</tbody></table></div>`, `<button class="btn secondary" data-cancel>知道了</button><button class="btn" id="mute-balance-notice">今天不再提示</button>`);
    $('#mute-balance-notice').onclick = async () => { state.reconNoticeDate = stamp(); await save(); closeModal(); toast('今天不再提醒客户余额。'); };
  }
  function renderRecon() {
    const mode = reconMode;
    const rows = reconBase();
    const groups = reconGroups(rows, mode);
    const pick = reconGroup && groups.some(g => g.name === reconGroup) ? reconGroup : '';
    const keyOf = o => mode === 'supplier' ? reconLine(o, mode).supplierKey : reconLine(o, mode).customer;
    const own = pick ? rows.filter(o => keyOf(o) === pick) : rows;
    const lines = own.map(o => reconLine(o, mode)).sort((a, b) => String(a.date).localeCompare(String(b.date)) || a.orderNo.localeCompare(b.orderNo));
    const units = lines.reduce((n, l) => n + l.qty, 0);
    const amount = lines.reduce((n, l) => n + (l.amount || 0), 0);
    const noPrice = lines.filter(l => l.amount == null).length;
    const isSup = mode === 'supplier';
    const groupLabel = isSup ? '供应商' : '客户';
    const accKind = isSup ? 'supplier' : 'customer';
    const account = pick ? reconAccount(pick, accKind) : {
      deposit: groups.reduce((n, g) => n + (g.deposit || 0), 0),
      total: groups.reduce((n, g) => n + (g.total || 0), 0),
      balance: groups.reduce((n, g) => n + (g.balance || 0), 0),
      noPrice: groups.reduce((n, g) => n + (g.noPrice || 0), 0)
    };
    return `<div class="page-intro"><div><div class="eyebrow">RECONCILIATION</div><h1>谁的单，多少钱，一页看清</h1><p>${isSup ? '按供应商汇总应付款：商品、数量、供价、金额。' : '按客户汇总应收款：商品、数量、单价、金额；填了充值金额还能看账户余额。'}已排除的明细不计入。</p></div><div class="intro-actions"><div class="tabs"><button class="tab ${isSup ? '' : 'active'}" data-recon-mode="customer">客户对账</button><button class="tab ${isSup ? 'active' : ''}" data-recon-mode="supplier">供应商对账</button></div>${rows.length ? `${isSup ? '' : '<button class="btn secondary" id="recon-import">' + icon('upload') + ' 导入对账</button>'}<button class="btn secondary" id="recon-export">${icon('download')} 导出${groupLabel}对账单</button>` : ''}<input type="file" id="recon-file" accept=".xlsx,.csv" hidden></div></div>
      ${!rows.length ? `<section class="card"><div class="empty-state">${icon('file')}<h3>还没有可对账的订单</h3><p>先到「订单采集与匹配」导入并匹配订单。</p><button class="btn" data-go="import">前往订单采集与匹配</button></div></section>` : `
      <div class="stats-grid">${[['box', groupLabel + '数', groups.length, `有订单的${groupLabel}`, '家'], ['file', '订单明细', lines.length, pick || `全部${groupLabel}`, '条'], ['check', '商品件数', units, '按订单数量累计', '件'], ['clock', isSup ? '应付合计' : '应收合计', reconMoney(amount), noPrice ? `还有 ${noPrice} 条缺${isSup ? '供价' : '单价'}` : `按${isSup ? '供价' : '订单单价'}计算`, '元']].map(([ic, l, v, n, u]) => `<article class="stat-card"><div class="stat-icon blue">${icon(ic)}</div><div><div class="stat-label">${l}</div><div class="stat-value">${v}<span>${u}</span></div><div class="stat-note">${n}</div></div></article>`).join('')}</div>
      <section class="card recon-account"><div class="recon-account-row">
        <div class="recon-field"><span>${pick ? esc(pick) + (isSup ? ' 已充值金额' : ' 总充值金额') : `全部${groupLabel}${isSup ? '已充值' : '充值'}合计`}</span><strong>¥${reconMoney(account.deposit)}</strong></div>
        ${pick ? `<div class="recon-actions"><button class="btn secondary small" id="recon-recharge">${icon('plus')} 充值</button><button class="btn ghost small" id="recon-recharge-list">充值记录</button></div>` : `<div class="recon-hint">在左侧选一个${groupLabel}后可充值、查看充值记录</div>`}
        <div class="recon-field"><span>全部订单${isSup ? '应付' : '应收'}</span><strong>¥${reconMoney(account.total)}</strong></div>
        <div class="recon-field"><span>账户余额</span><strong class="${account.balance < 0 ? 'neg' : 'pos'}">¥${reconMoney(account.balance)}</strong></div>
        ${account.noPrice ? `<div class="recon-hint">${account.noPrice} 条明细缺${isSup ? '供价' : '单价'}，未计入${isSup ? '应付' : '应收'}</div>` : ''}
      </div></section>
      <div class="export-layout"><aside class="card export-list"><div class="card-header"><h2>${groupLabel} <span class="count">${groups.length}</span></h2></div>
        <button class="export-group ${pick === '' ? 'active' : ''}" data-recon-group=""><span class="supplier-avatar">全</span><span><strong>全部${groupLabel}</strong><small>${rows.length} 条明细</small></span>${icon('arrow')}</button>
        ${groups.map(g => `<button class="export-group ${pick === g.name ? 'active' : ''}" data-recon-group="${esc(g.name)}"><span class="supplier-avatar">${esc(g.name.slice(0, 1))}</span><span><strong>${esc(g.name)}</strong><small>${g.count} 条明细 · ¥${reconMoney(g.amount)}${!isSup && g.deposit ? ` · 余额 ¥${reconMoney(g.balance)}` : ''}</small></span>${icon('arrow')}</button>`).join('')}
      </aside><section class="card export-preview"><div class="card-header"><div><h2>${pick ? esc(pick) : `全部${groupLabel}`} · 订单明细</h2><p class="muted">${reconScope === 'active' ? '仅本次勾选的批次' : '含历史批次'} · ${lines.length} 条 · ${isSup ? '应付' : '应收'}合计 ¥${reconMoney(amount)}${noPrice ? ` · ${noPrice} 条缺${isSup ? '供价' : '单价'}` : ''}</p></div><div class="intro-actions"><div class="tabs"><button class="tab ${reconScope === 'all' ? 'active' : ''}" data-recon-scope="all">全部订单</button><button class="tab ${reconScope === 'active' ? 'active' : ''}" data-recon-scope="active">仅本次批次</button></div></div></div>
        ${noPrice ? `<div class="alert warning">${icon('shield')}<span>${noPrice} 条明细没有${isSup ? '供价（商品档案未填）' : '单价（源表未提供）'}，金额不计入合计；导出后可在 Excel 里补。</span></div>` : ''}
        <div class="table-wrap"><table class="data-table pool-table"><thead><tr>${RECON_BILL_HEADERS.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${lines.slice(0, 200).map(l => `<tr>${reconBillValues(l, isSup ? 'supplier' : 'customer').map((v, i) => i === 28 ? `<td><label class="checkbox-row"><input type="checkbox" data-invoice="${esc(l.orderId)}" ${l.invoiced ? 'checked' : ''}><span>${l.invoiced ? '已开票' : '未开票'}</span></label></td>` : `<td class="${typeof v === 'number' ? 'numeric' : ''}">${v === '' ? '<span class="muted">—</span>' : esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
        <div class="pagination"><span>预览前 ${Math.min(lines.length, 200)} 条，导出包含全部 ${lines.length} 条。</span></div></section></div>`}`;
  }
  async function importRecon(file) {
    if (!file) return;
    await guardAsync(async () => {
      const book = await OrderExcel.read(file);
      let si = book.sheets.findIndex(s => s.rows.some(r => (r || []).some(c => poolNorm(c) === '是否开票')));
      if (si < 0) si = 0;
      const rows = book.sheets[si].rows;
      const header = OrderExcel.detectHeader(rows);
      const headers = rows[header] || [];
      const find = label => {
        const target = poolNorm(label);
        let idx = headers.findIndex(h => poolNorm(h) === target);
        if (idx < 0) idx = headers.findIndex(h => poolNorm(h) && poolNorm(h).includes(target));
        return idx;
      };
      const productCol = find('商品名称') >= 0 ? find('商品名称') : find('产品名称');
      const cols = { orderNo: find('订单编号'), year: find('年份'), date: find('下单时间') >= 0 ? find('下单时间') : find('日期'), customer: find('客户名称'), barcode: find('条码'), product: productCol, qty: find('商品数量') >= 0 ? find('商品数量') : find('数量'), price: find('商品单价') >= 0 ? find('商品单价') : find('单价'), invoice: find('是否开票') };
      if (cols.product < 0 || cols.invoice < 0) { toast('没有找到「商品名称」或「是否开票」列，请使用对账格式表头。', 'warning'); return; }
      const get = (r, i) => { if (i < 0) return ''; const v = r[i]; return v == null ? '' : String(v).trim(); };
      const truthy = v => /^(是|已开票|开票|y|yes|true|1)$/i.test(v);
      const falsy = v => /^(否|未开票|不开票|n|no|false|0)$/i.test(v);
      const lines = reconBase().map(o => reconLine(o, 'customer'));
      const byNo = new Map();
      for (const l of lines) if (l.orderNo) byNo.set(norm(l.orderNo), (byNo.get(norm(l.orderNo)) || []).concat(l));
      const key = l => [norm(l.customer), String(l.date).slice(0, 10), norm(l.product), String(l.qty), String(l.price == null ? '' : l.price)].join('|');
      const byKey = new Map();
      for (const l of lines) { const k = key(l); byKey.set(k, (byKey.get(k) || []).concat(l)); }
      const items = [];
      for (let i = header + 1; i < rows.length; i++) {
        const r = rows[i] || [];
        const product = get(r, cols.product);
        const mark = get(r, cols.invoice);
        if (!product || (!mark)) continue;
        const want = truthy(mark) ? true : falsy(mark) ? false : null;
        if (want === null) continue;
        const date = String(get(r, cols.date)).slice(0, 10);
        const no = get(r, cols.orderNo);
        const k = [norm(get(r, cols.customer)), date, norm(product), String(get(r, cols.qty)), String(get(r, cols.price))].join('|');
        const cands = (no && byNo.get(norm(no))) || byKey.get(k) || [];
        items.push({ row: i + 1, product, customer: get(r, cols.customer), date, mark, want, orderId: cands.length === 1 ? cands[0].orderId : '', reason: cands.length === 1 ? '' : (cands.length ? '匹配到多条明细，已跳过' : '未找到对应订单明细') });
      }
      if (!items.length) { toast('表里没有可回填的「是否开票」数据。', 'warning'); return; }
      const ok = items.filter(it => it.orderId);
      const bad = items.filter(it => !it.orderId);
      const changes = ok.filter(it => { const o = state.orders.find(x => x.id === it.orderId); return o && !!o.invoiced !== it.want; });
      openModal('导入对账（回填开票状态）', `
        <div class="alert info">${icon('file')}<span>已读取 <strong>${esc(file.name)}</strong> · ${items.length} 行带「是否开票」标记。</span></div>
        <div class="stats-grid compact">${[['可回填', ok.length, '优先按订单编号匹配，其次按客户+日期+商品+数量+单价'], ['将变更', changes.length, '与当前开票标记不同'], ['未匹配', bad.length, '明细对不上，跳过']].map(([l, v, n]) => `<article class="stat-card"><div class="stat-label">${l}</div><div class="stat-value">${v}<span>条</span></div><div class="stat-note">${esc(n)}</div></article>`).join('')}</div>
        <p class="muted">只回填「是否开票」，不改单价与金额。订单编号匹配不到时才用五要素，且必须唯一，避免改错行。</p>
        ${bad.length ? `<div class="table-wrap"><table class="data-table"><thead><tr><th>行</th><th>客户</th><th>日期</th><th>产品</th><th>原因</th></tr></thead><tbody>${bad.slice(0, 5).map(it => `<tr><td class="numeric">${it.row}</td><td>${esc(it.customer)}</td><td>${esc(it.date)}</td><td>${esc(it.product)}</td><td>${esc(it.reason)}</td></tr>`).join('')}</tbody></table></div>` : ''}
      `, `<button class="btn secondary" data-cancel>取消</button><button class="btn" id="recon-import-confirm" ${ok.length ? '' : 'disabled'}>确认回填</button>`, true);
      const btn = $('#recon-import-confirm');
      if (!ok.length) return;
      btn.onclick = async () => {
        let changed = 0;
        for (const it of ok) {
          const o = state.orders.find(x => x.id === it.orderId);
          if (!o) continue;
          if (!!o.invoiced !== it.want) { o.invoiced = it.want; changed++; }
        }
        await save(); closeModal(); render();
        toast(`已回填 ${changed} 条开票状态${bad.length ? `，${bad.length} 行未匹配跳过` : ''}。`);
      };
    });
  }
  async function exportRecon() {
    const mode = reconMode;
    const rows = reconBase();
    const groups = reconGroups(rows, mode);
    const pick = reconGroup && groups.some(g => g.name === reconGroup) ? reconGroup : '';
    const keyOf = o => mode === 'supplier' ? reconLine(o, mode).supplierKey : reconLine(o, mode).customer;
    const own = pick ? rows.filter(o => keyOf(o) === pick) : rows;
    const lines = own.map(o => reconLine(o, mode));
    const isSup = mode === 'supplier';
    const headers = RECON_BILL_HEADERS;
    const data = lines.map(l => reconBillValues(l, mode));
    const name = pick ? `${pick}_${isSup ? '供应商' : '客户'}对账` : `全部${isSup ? '供应商' : '客户'}_对账`;
    await OrderExcel.downloadWorkbook({ headers, rows: data, sheetName: '对账单明细', filename: `${name}_${stamp()}.xlsx` });
    if (!isSup) {
      const acc = pick ? reconAccount(pick) : groups.reduce((a, g) => ({ deposit: a.deposit + (g.deposit || 0), total: a.total + (g.total || 0), balance: a.balance + (g.balance || 0) }), { deposit: 0, total: 0, balance: 0 });
      toast(`已导出 ${data.length} 条客户对账；充值 ¥${reconMoney(acc.deposit)} · 应收 ¥${reconMoney(acc.total)} · 余额 ¥${reconMoney(acc.balance)}。`);
    } else {
      toast(`已导出 ${data.length} 条供应商对账明细。`);
    }
  }
  function bindRecon() {
    $$('[data-recon-mode]').forEach(b => b.onclick = () => { if (reconMode !== b.dataset.reconMode) { reconMode = b.dataset.reconMode; reconGroup = ''; render(); } });
    $$('[data-recon-group]').forEach(b => b.onclick = () => { reconGroup = b.dataset.reconGroup; render(); });
    $$('[data-recon-scope]').forEach(b => b.onclick = () => { reconScope = b.dataset.reconScope; render(); });
    $('#recon-export')?.addEventListener('click', () => guardAsync(exportRecon));
    $('#recon-import')?.addEventListener('click', () => $('#recon-file').click());
    $('#recon-file')?.addEventListener('change', e => { const fl = e.target.files && e.target.files[0]; e.target.value = ''; if (fl) importRecon(fl); });
    $$('[data-invoice]').forEach(cb => cb.onchange = async () => {
      const o = state.orders.find(x => x.id === cb.dataset.invoice);
      if (!o) return;
      o.invoiced = cb.checked;
      await save(); render(); toast(cb.checked ? '已标记为已开票。' : '已标记为未开票。');
    });
    const rk = () => reconMode === 'supplier' ? 'supplier' : 'customer';
    $('#recon-recharge')?.addEventListener('click', () => rechargeModal(reconGroup, rk()));
    $('#recon-recharge-list')?.addEventListener('click', () => rechargeListModal(reconGroup, rk()));
  }
  const POOL_FIELDS = [
    ['category', '类别'], ['supplier', '供应商名称'], ['brand', '品牌'], ['name', '品名'], ['image', '图片'],
    ['spec', '内容物配置'], ['netContent', '净含量'], ['groupPrice', '集采价'], ['minQty', '集采起订量'],
    ['cost', '一件代发价格'], ['suggestPrice', '零售价'], ['shipping', '快递'], ['taxRate', '税率'],
    ['taxCode', '税收编码'], ['internalSku', '内部编码'], ['supplierSku', '供应商编码'], ['barcode', '条码'],
    ['shelfLife', '保质期'], ['origin', '产地'], ['shipFrom', '发货地'], ['size', '产品尺寸'],
    ['remoteExcluded', '偏远不发'], ['link', '京东/天猫链接'], ['payAccount', '供应商付款账号'],
    ['bankCode', '供应商开户行行号'], ['bankAccount', '供应商银行账号'], ['bankName', '供应商银行名称'],
    ['createdAt', '建档时间'], ['status', '合作状态']
  ];
  const poolNorm = v => String(v == null ? '' : v).replace(/\s+/g, '').replace(/[（）()]/g, '');
  function poolHeaderMap(headers) {
    const map = {};
    for (const [k, label] of POOL_FIELDS) {
      const target = poolNorm(label);
      let idx = headers.findIndex(h => poolNorm(h) === target);
      if (idx < 0) idx = headers.findIndex(h => poolNorm(h) && (poolNorm(h).includes(target) || target.includes(poolNorm(h))));
      map[k] = idx;
    }
    return map;
  }
  function supplierByName(name) {
    return state.suppliers.find(s => norm(s.name) === norm(name));
  }
  async function poolImport(file) {
    if (!file) return;
    await guardAsync(async () => {
      const book = await OrderExcel.read(file);
      let si = book.sheets.findIndex(s => s.rows.some(r => (r || []).some(c => poolNorm(c) === '品名')));
      if (si < 0) si = book.sheets.findIndex(s => s.rows.length > 2);
      if (si < 0) si = 0;
      const rows = book.sheets[si].rows;
      const header = OrderExcel.detectHeader(rows);
      const headers = rows[header] || [];
      const map = poolHeaderMap(headers);
      if (map.name < 0) { toast('没有找到「品名」列，请确认使用的是商品池模板。', 'warning'); return; }
      const get = (r, k) => { const i = map[k]; if (i < 0) return ''; const v = r[i]; return v == null ? '' : String(v).trim(); };
      const items = [];
      for (let i = header + 1; i < rows.length; i++) {
        const r = rows[i] || [];
        if (!get(r, 'name')) continue;
        items.push({ row: i + 1, raw: r });
      }
      if (!items.length) { toast('表里没有可导入的商品行。', 'warning'); return; }
      const existing = items.filter(it => {
        const sku = get(it.raw, 'internalSku');
        const supName = get(it.raw, 'supplier');
        if (sku && state.products.some(p => norm(p.internalSku) === norm(sku))) return true;
        const sup = supName ? supplierByName(supName) : null;
        const supSku = get(it.raw, 'supplierSku');
        return !!(sup && supSku && state.products.some(p => p.supplierId === sup.id && norm(p.supplierSku) === norm(supSku)));
      }).length;
      const newSupNames = [...new Set(items.map(it => get(it.raw, 'supplier')).filter(Boolean))]
        .filter(name => !supplierByName(name));
      openModal('导入商品表', `
        <div class="alert info">${icon('file')}<span>已读取 <strong>${esc(file.name)}</strong> · 工作表「${esc(book.sheets[si].name)}」· 识别 <strong>${items.length}</strong> 条商品。</span></div>
        <div class="stats-grid compact">${[['新增商品', items.length - existing, '按内部编码 / 供应商编码匹配不到'], ['更新商品', existing, '编码已存在，按表内数据覆盖'], ['新建供应商', newSupNames.length, newSupNames.slice(0, 3).join('、') || '无需新建']].map(([l, v, n]) => `<article class="stat-card"><div class="stat-label">${l}</div><div class="stat-value">${v}<span>条</span></div><div class="stat-note">${esc(n)}</div></article>`).join('')}</div>
        <p class="muted">匹配顺序：先按「内部编码」，再按同一供应商下的「供应商编码」；都匹配不到就新增。空单元格不会覆盖已有值。</p>
        <div class="table-wrap"><table class="data-table"><thead><tr><th>行</th><th>品名</th><th>供应商</th><th>内部编码</th><th>一件代发价格</th></tr></thead><tbody>${items.slice(0, 6).map(it => `<tr><td class="numeric">${it.row}</td><td>${esc(get(it.raw, 'name'))}</td><td>${esc(get(it.raw, 'supplier') || '—')}</td><td class="mono">${esc(get(it.raw, 'internalSku') || '自动生成')}</td><td class="numeric">${esc(get(it.raw, 'cost') || '—')}</td></tr>`).join('')}</tbody></table></div>
        ${items.length > 6 ? `<p class="muted">仅预览前 6 条，实际导入 ${items.length} 条。</p>` : ''}
      `, `<button class="btn secondary" data-cancel>取消</button><button class="btn" id="pool-confirm">确认导入</button>`, true);
      $('#pool-confirm').onclick = async () => {
        let added = 0, updated = 0, supAdded = 0;
        for (const it of items) {
          const r = it.raw;
          const supName = get(r, 'supplier');
          let sup = supName ? supplierByName(supName) : null;
          if (supName && !sup) {
            sup = { id: uid(), name: supName, contact: '', phone: '', note: '', template: null, derived: false, bankName: '', bankCode: '', bankAccount: '', payAccount: '' };
            state.suppliers.push(sup);
            supAdded++;
          }
          if (sup) {
            for (const k of ['bankName', 'bankCode', 'bankAccount', 'payAccount']) {
              const v = get(r, k === 'bankName' ? 'bankName' : k);
              if (v) sup[k] = v;
            }
          }
          const sku = get(r, 'internalSku');
          const supSku = get(r, 'supplierSku');
          let prod = sku ? state.products.find(p => norm(p.internalSku) === norm(sku)) : null;
          if (!prod && sup && supSku) prod = state.products.find(p => p.supplierId === sup.id && norm(p.supplierSku) === norm(supSku));
          if (!prod) {
            prod = { id: uid(), internalSku: sku || nextProductSku(), name: '', spec: '', netContent: '', image: '', unit: '盒', supplierId: sup ? sup.id : '', supplierSku: supSku || sku || '', barcode: '', brand: '', category: '', cost: '', groupPrice: '', broadcastPrice: '', suggestPrice: '', marginOwn: '', marginBroadcast: '', jdPrice: '', link: '', aliases: [], note: '', shipping: '', taxRate: '', taxCode: '', shelfLife: '', origin: '', shipFrom: '', size: '', remoteExcluded: '否', status: '合作中', minQty: '', createdAt: '' };
            state.products.push(prod);
            added++;
          } else updated++;
          for (const [k] of POOL_FIELDS) {
            if (['supplier', 'payAccount', 'bankCode', 'bankAccount', 'bankName'].includes(k)) continue;
            const v = get(r, k);
            if (v !== '') prod[k] = v;
          }
          if (sup) prod.supplierId = sup.id;
          if (!prod.supplierSku) prod.supplierSku = prod.internalSku;
          if (!prod.createdAt) prod.createdAt = get(r, 'createdAt') || new Date().toISOString();
          if (!prod.status) prod.status = '合作中';
        }
        await save(); closeModal(); render();
        toast(`商品表导入完成：新增 ${added} 个商品、更新 ${updated} 个${supAdded ? `，新建 ${supAdded} 个供应商` : ''}。`);
      };
    });
  }
  async function exportProductPool() {
    const list = visibleProducts();
    if (!list.length) { toast('还没有商品可导出。', 'warning'); return; }
    const headers = ['序号', '类别', '供应商名称', '品牌', '品名', '图片', '内容物配置', '净含量', '集采价', '集采起订量', '一件代发价格', '零售价', '快递', '税率', '税收编码', '内部编码', '供应商编码', '条码', '保质期', '产地', '发货地', '产品尺寸', '偏远不发', '京东/天猫链接', '供应商付款账号', '供应商开户行行号', '供应商银行账号', '供应商银行名称', '建档时间', '合作状态'];
    const rows = list.map((p, i) => {
      const s = sById(p.supplierId) || {};
      return [
        i + 1, p.category || '', s.name || '', p.brand || '', p.name || '', p.image || '', p.spec || '', p.netContent || '',
        p.groupPrice || '', p.minQty || '', p.cost || '', p.suggestPrice || '', p.shipping || '', p.taxRate || '', p.taxCode || '',
        p.internalSku || '', p.supplierSku || '', p.barcode || '', p.shelfLife || '', p.origin || '', p.shipFrom || '', p.size || '',
        p.remoteExcluded || '否', p.link || '', s.payAccount || s.bankAccount || '', s.bankCode || '', s.bankAccount || '', s.bankName || '',
        (p.createdAt || '').slice(0, 10), p.status || '合作中'
      ];
    });
    const scope = supplierFilter === 'all' ? '全部供应商' : (sById(supplierFilter)?.name || '供应商');
    await OrderExcel.downloadWorkbook({ headers, rows, sheetName: '商品池', filename: `商品表_${scope}_${stamp()}.xlsx` });
    toast(`已导出 ${rows.length} 个商品（${scope}）。`);
  }
  function renderLibrary() {
    if (supplierFilter !== 'all' && !sById(supplierFilter)) supplierFilter = 'all';
    const sel = supplierFilter === 'all' ? null : sById(supplierFilter);
    const count = visibleProducts().length;
    return `<div class="page-intro"><div><div class="eyebrow">SUPPLIER LIBRARY</div><h1>一次建立，日常复用</h1><p>导入商品表（30 列标准模板）维护商品池；旧格式报价单也可导入，点击上方供应商可切换查看其商品。</p></div><div class="intro-actions"><button class="btn secondary" id="new-supplier">${icon('plus')} 新增供应商</button><button class="btn secondary" id="export-pool-top">${icon('download')} 导出商品表</button><button class="btn" id="import-pool-top">${icon('upload')} 导入商品表</button></div></div>
      ${supplierRail()}
      ${OrderSeed.candidates.some(c => !state.products.find(p => p.id === c.productId)?.aliases.some(a => norm(a.name) === norm(c.name))) ? `<div class="alert info candidate-alert"><div>${icon('link')}<span><strong>样表中发现 3 条商品别名建议</strong><small>已区分儿童款与普通款。确认后才会用于订单匹配。</small></span></div><button class="btn secondary small" id="review-candidates">查看并确认</button></div>` : ''}
      <section class="card"><div class="card-header"><div><h2>${sel ? esc(sel.name) : '供应商库'} <span class="count" id="product-count">${count}</span></h2><p class="muted">${sel ? (sel.derived ? '该供应商由报价单品牌自动推断，请核对实际供货主体。' : `点击其他供应商可切换视图。`) : `共 ${state.suppliers.length} 家供应商`}</p></div><div class="intro-actions"><button class="btn ghost small" id="pool-cols">${poolColumnLabel()}</button><input type="file" id="pool-file" accept=".xlsx,.csv" hidden><button class="btn ghost small" id="import-quote">导入报价单（旧格式）</button><button class="btn ghost small" id="import-catalog">简易导入（编码+名称）</button><input type="file" id="catalog-file" accept=".xlsx,.csv" hidden><input type="file" id="quote-file" accept=".xlsx,.csv" hidden><button class="btn small" id="new-product">${icon('plus')} 新增商品</button></div></div>
      <div class="toolbar toolbar-end"><label class="search-wrap">${icon('search')}<input id="product-search" class="search-input" value="${esc(query)}" placeholder="搜索商品名称、品牌、69码、编码" aria-label="搜索商品"></label></div><div id="products-table">${productsTable()}</div></section>`;
  }
  function visibleProducts() {
    const q = norm(query);
    return state.products.filter(p => (supplierFilter === 'all' || p.supplierId === supplierFilter) && (!q || norm([p.name, p.supplierSku, p.internalSku, p.brand, p.barcode, p.category, ...p.aliases.map(a => a.name || a.code)].join(' ')).includes(q)));
  }
  const POOL_COMMON = ['no', 'supplier', 'name', 'spec', 'netContent', 'cost', 'suggestPrice', 'internalSku', 'supplierSku', 'barcode', 'status'];
  const poolCell = v => v === '' || v == null ? '<span class="muted">—</span>' : esc(v);
  const poolLink = v => v ? `<a class="text-link" href="${esc(v)}" target="_blank" rel="noopener">链接</a>` : '<span class="muted">—</span>';
  const POOL_COLUMNS = [
    { key: 'no', label: '序号', cell: (p, i) => String(i + 1), cls: 'numeric' },
    { key: 'category', label: '类别', cell: p => poolCell(p.category) },
    { key: 'supplier', label: '供应商名称', cell: (p, i, s) => poolCell(s.name) },
    { key: 'brand', label: '品牌', cell: p => poolCell(p.brand) },
    { key: 'name', label: '品名', cell: p => `<strong>${esc(p.name)}</strong>`, cls: 'product-cell' },
    { key: 'image', label: '图片', cell: p => poolLink(p.image) },
    { key: 'spec', label: '内容物配置', cell: p => poolCell(p.spec) },
    { key: 'netContent', label: '净含量', cell: p => poolCell(p.netContent) },
    { key: 'groupPrice', label: '集采价', cell: p => poolCell(p.groupPrice), cls: 'numeric' },
    { key: 'minQty', label: '集采起订量', cell: p => poolCell(p.minQty), cls: 'numeric' },
    { key: 'cost', label: '一件代发价格', cell: p => poolCell(p.cost), cls: 'numeric' },
    { key: 'suggestPrice', label: '零售价', cell: p => poolCell(p.suggestPrice), cls: 'numeric' },
    { key: 'shipping', label: '快递', cell: p => poolCell(p.shipping) },
    { key: 'taxRate', label: '税率', cell: p => poolCell(p.taxRate) },
    { key: 'taxCode', label: '税收编码', cell: p => poolCell(p.taxCode), cls: 'mono' },
    { key: 'internalSku', label: '内部编码', cell: p => poolCell(p.internalSku), cls: 'mono' },
    { key: 'supplierSku', label: '供应商编码', cell: p => poolCell(p.supplierSku), cls: 'mono' },
    { key: 'barcode', label: '条码', cell: p => poolCell(p.barcode), cls: 'mono' },
    { key: 'shelfLife', label: '保质期', cell: p => poolCell(p.shelfLife) },
    { key: 'origin', label: '产地', cell: p => poolCell(p.origin) },
    { key: 'shipFrom', label: '发货地', cell: p => poolCell(p.shipFrom) },
    { key: 'size', label: '产品尺寸', cell: p => poolCell(p.size) },
    { key: 'remoteExcluded', label: '偏远不发', cell: p => poolCell(p.remoteExcluded || '否') },
    { key: 'link', label: '京东/天猫链接', cell: p => poolLink(p.link) },
    { key: 'payAccount', label: '供应商付款账号', cell: (p, i, s) => poolCell(s.payAccount || s.bankAccount), cls: 'mono' },
    { key: 'bankCode', label: '供应商开户行行号', cell: (p, i, s) => poolCell(s.bankCode), cls: 'mono' },
    { key: 'bankAccount', label: '供应商银行账号', cell: (p, i, s) => poolCell(s.bankAccount), cls: 'mono' },
    { key: 'bankName', label: '供应商银行名称', cell: (p, i, s) => poolCell(s.bankName) },
    { key: 'createdAt', label: '建档时间', cell: p => poolCell((p.createdAt || '').slice(0, 10)) },
    { key: 'status', label: '合作状态', cell: p => poolCell(p.status || '合作中') }
  ];
  function poolColumnKeys() {
    try {
      const saved = localStorage.getItem('orderflow.poolColumns');
      if (saved === 'all') return POOL_COLUMNS.map(c => c.key);
      if (saved) {
        const arr = JSON.parse(saved);
        const valid = arr.filter(k => POOL_COLUMNS.some(c => c.key === k));
        if (valid.length) return valid;
      }
    } catch (_) { }
    return POOL_COMMON.slice();
  }
  function poolColumnLabel() {
    const keys = poolColumnKeys();
    if (keys.length === POOL_COLUMNS.length) return '列：全部';
    if (keys.length === POOL_COMMON.length && POOL_COMMON.every(k => keys.includes(k))) return '列：常用';
    return `列：自定义(${keys.length})`;
  }
  function openColumnPicker() {
    const current = new Set(poolColumnKeys());
    openModal('显示哪些列', `<p class="muted">勾选要在商品表里显示的列；表格列再多都只是展示，导出始终是完整 30 列。</p>`
      + `<div class="col-grid">${POOL_COLUMNS.map(c => `<label class="col-item"><input type="checkbox" value="${c.key}" ${current.has(c.key) ? 'checked' : ''}><span>${c.label}</span></label>`).join('')}</div>`,
      `<button class="btn secondary" data-cancel>取消</button><button class="btn secondary" id="cols-common">常用列</button><button class="btn secondary" id="cols-all">全部列</button><button class="btn" id="cols-apply">应用</button>`, true);
    const boxes = () => [...document.querySelectorAll('#modal-root .col-item input')];
    $('#cols-common').onclick = () => boxes().forEach(b => { b.checked = POOL_COMMON.includes(b.value); });
    $('#cols-all').onclick = () => boxes().forEach(b => { b.checked = true; });
    $('#cols-apply').onclick = () => {
      const keys = boxes().filter(b => b.checked).map(b => b.value);
      if (!keys.length) { toast('至少保留一列。', 'warning'); return; }
      localStorage.setItem('orderflow.poolColumns', keys.length === POOL_COLUMNS.length ? 'all' : JSON.stringify(keys));
      closeModal();
      const box = $('#products-table');
      if (box) box.innerHTML = productsTable();
      bindProductActions();
      const btn = $('#pool-cols');
      if (btn) btn.textContent = poolColumnLabel();
    };
  }
  function productsTable() {
    const ps = visibleProducts();
    const pages = Math.max(1, Math.ceil(ps.length / 40));
    if (productPage > pages) productPage = pages;
    const slice = ps.slice((productPage - 1) * 40, productPage * 40);
    const offset = (productPage - 1) * 40;
    const cols = poolColumnKeys().map(k => POOL_COLUMNS.find(c => c.key === k)).filter(Boolean);
    return `<div class="table-wrap"><table class="data-table pool-table"><thead><tr>${cols.map(c => `<th>${c.label}</th>`).join('')}<th>操作</th></tr></thead><tbody>${slice.map((p, i) => {
      const s = sById(p.supplierId) || {};
      return `<tr>${cols.map(c => `<td class="${c.cls || ''}">${c.cell(p, offset + i, s)}</td>`).join('')}<td><button class="text-link" data-edit-product="${esc(p.id)}">编辑</button></td></tr>`;
    }).join('') || `<tr><td colspan="${cols.length + 1}" class="empty-cell">没有找到商品，请导入报价单或调整筛选。</td></tr>`}</tbody></table></div><div class="pagination"><span>共 ${ps.length} 个商品 · 每页 40 个</span><div><button class="btn ghost small" data-ppage="${productPage - 1}" ${productPage <= 1 ? 'disabled' : ''}>上一页</button><span>${productPage} / ${pages}</span><button class="btn ghost small" data-ppage="${productPage + 1}" ${productPage >= pages ? 'disabled' : ''}>下一页</button></div></div>`;
  }
  function exportRow(order, product) {
    const sup = sById(product.supplierId);
    return {
      ...order,
      sourceProductName: order.productName,
      productName: product.name,
      internalSku: product.internalSku,
      spec: product.spec,
      unit: product.unit,
      supplierId: product.supplierId,
      supplierName: sup ? sup.name : '',
      supplierSku: product.supplierSku,
      quantity: Number(order.quantity) || order.quantity,
      unitPrice: order.unitPrice === '' ? '' : (Number(order.unitPrice) || order.unitPrice)
    };
  }
  function exportGroups() {
    const buckets = new Map();
    for (const order of activeRows()) {
      const st = status(order);
      if (st.key !== 'ready' || !st.product) continue;
      const id = st.product.supplierId;
      if (!buckets.has(id)) buckets.set(id, []);
      buckets.get(id).push(exportRow(order, st.product));
    }
    return state.suppliers.map(s => ({ supplier: s, rows: buckets.get(s.id) || [] })).filter(g => g.rows.length);
  }
  function pendingExportCount() {
    return activeRows().filter(r => { const k = status(r).key; return k !== 'ready' && k !== 'excluded'; }).length;
  }
  function templateValue(row,col){if(col.field==='constant')return col.fallback||'';if(col.field==='blank')return '';const v=row[col.field];return v===undefined||v===null||v===''?col.fallback||'':v;}
  function exportIssues(g){if(!g.supplier.template?.columns.length)return ['尚未设置格式'];const cols=g.supplier.template.columns;const errs=[];for(const c of cols){if(c.field==='orderNo')continue;const n=c.required?g.rows.filter(r=>templateValue(r,c)==='').length:0;if(n)errs.push(`“${c.header}”有 ${n} 条为空`);}return errs;}
  function renderExport(){const gs=exportGroups();if(!gs.some(g=>g.supplier.id===selectedExport))selectedExport=gs[0]?.supplier.id||'';const g=gs.find(g=>g.supplier.id===selectedExport);const pending=pendingExportCount();const bare=state.suppliers.filter(s=>!s.template?.columns?.length);const blankOrder=g?g.rows.filter(r=>!String(r.orderNo||'').trim()).length:0;return `<div class="page-intro"><div><div class="eyebrow">SUPPLIER FORMAT</div><h1>一家供应商，一套格式</h1><p>先给每家供应商配好格式表，再把匹配好的订单按供应商输出。</p></div><div class="intro-actions"><button class="btn secondary" data-go="import">返回匹配订单</button><button class="btn" id="export-all" ${!gs.length||gs.some(g=>exportIssues(g).length)?'disabled':''}>${icon('download')} 批量导出 ZIP</button></div></div>${bare.length?`<div class="alert warning">${icon('shield')}<span>${esc(bare.map(s=>s.name).join('、'))} 还没有格式模板，先点下方「设置格式」配好列顺序。</span></div>`:''}${pending?`<div class="alert info">${icon('clock')}<span>还有 ${pending} 条明细未完成匹配，本次不导出；已匹配的 ${gs.reduce((n,x)=>n+x.rows.length,0)} 条按供应商输出。</span></div>`:''}${!gs.length?`<section class="card"><div class="empty-state">${icon('box')}<h3>还没有可导出的订单</h3><p>导入订单并完成商品匹配后，这里会按供应商分组。</p><button class="btn" data-go="import">前往订单采集与匹配</button></div></section>`:`<div class="export-layout"><aside class="card export-list"><div class="card-header"><h2>供应商分组 <span class="count">${gs.length}</span></h2></div>${gs.map(g=>`<button class="export-group ${g.supplier.id===selectedExport?'active':''}" data-export-group="${esc(g.supplier.id)}"><span class="supplier-avatar">${esc(g.supplier.name[0])}</span><span><strong>${esc(g.supplier.name)}</strong><small>${g.rows.length} 条明细 · ${g.rows.reduce((n,r)=>n+Number(r.quantity),0)} 件</small></span>${icon('arrow')}</button>`).join('')}</aside><section class="card export-preview"><div class="card-header"><div><h2>${esc(g.supplier.name)} · 格式预览</h2><p class="muted">${esc(g.supplier.template?.name||'尚未设置导单格式')}</p></div><div class="intro-actions"><button class="btn secondary small" data-template="${esc(g.supplier.id)}">设置格式</button><button class="btn small" id="export-current" ${exportIssues(g).length?'disabled':''}>${icon('download')} 导出 Excel</button></div></div>${exportIssues(g).length?`<div class="alert warning">${esc(exportIssues(g).join('；'))}</div>`:g.supplier.template.columns.some(c=>c.field==='buyerNickname'&&!c.required)?'<div class="alert info">源表缺少买家昵称时留空。可在格式设置中选择固定值或明确取收货人。</div>':''}${blankOrder?`<div class="alert info">${icon('clock')}<span>有 ${blankOrder} 条明细没有订单号，该列将留空导出。</span></div>`:''}${templatePreview(g)}<div class="card-body export-summary"><span>${icon('check')} 仅导出本次选中批次的明细，样表原有示例订单不会混入。</span><span>预览前 10 行，完整导出 ${g.rows.length} 条。</span></div></section></div>`}<section class="card template-management"><div class="card-header"><h2>供应商格式管理</h2><span class="muted">支持上传样表并逐列配置</span></div><div class="card-body template-list">${state.suppliers.map(s=>`<div><span>${icon('file')}<strong>${esc(s.name)}</strong><small>${esc(s.template?.name||'待配置')}</small></span><button class="btn secondary small" data-template="${esc(s.id)}">设置格式</button></div>`).join('')}</div></section>`;}
  function templatePreview(g){const cols=g.supplier.template?.columns||[];return cols.length?`<div class="table-wrap"><table class="data-table"><thead><tr>${cols.map(c=>`<th>${esc(c.header)}</th>`).join('')}</tr></thead><tbody>${g.rows.slice(0,10).map(r=>`<tr>${cols.map(c=>`<td>${esc(templateValue(r,c))}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`:'';}
  function bindView(){
    $('#demo-btn')?.addEventListener('click',loadDemo);
    $('#dropzone')?.addEventListener('click',()=>$('#order-files').click());
    $('#dropzone')?.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('#order-files').click();}});
    $('#order-files')?.addEventListener('change',e=>queueFiles(e.target.files));
    const drop=$('#dropzone');if(drop){drop.ondragover=e=>{e.preventDefault();drop.classList.add('dragover');};drop.ondragleave=()=>drop.classList.remove('dragover');drop.ondrop=e=>{e.preventDefault();drop.classList.remove('dragover');queueFiles(e.dataTransfer.files);};}
    $$('[data-filter]').forEach(b=>b.onclick=()=>{filter=b.dataset.filter;page=1;render();});
    $('#order-search')?.addEventListener('input',e=>{query=e.target.value;page=1;$('#orders-table').innerHTML=orderTable();bindOrderActions();});
    bindOrderActions();
    $('#batches-btn')?.addEventListener('click',manageBatches);
    $('#new-supplier')?.addEventListener('click',()=>supplierModal());
    $('#new-product')?.addEventListener('click',()=>productModal());
    $$('[data-edit-supplier]').forEach(b=>b.onclick=()=>supplierModal(b.dataset.editSupplier));
    bindProductActions();
    $('#product-search')?.addEventListener('input',e=>{query=e.target.value;productPage=1;$('#products-table').innerHTML=productsTable();const badge=$('#product-count');if(badge)badge.textContent=String(visibleProducts().length);bindProductActions();});
    $('#review-candidates')?.addEventListener('click',reviewCandidates);
    bindShipment();
    bindCustomers();
    bindRecon();
    $('#import-quote')?.addEventListener('click',()=>$('#quote-file').click());
    $('#quote-file')?.addEventListener('change',e=>{const f=e.target.files?.[0];e.target.value='';quoteImport(f);});
    $$('[data-supplier-tile]').forEach(t=>{const pick=()=>{supplierFilter=t.dataset.supplierTile;productPage=1;render();};t.addEventListener('click',e=>{if(e.target.closest('[data-edit-supplier]'))return;pick();});t.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();pick();}});});
    $('#pool-cols')?.addEventListener('click',openColumnPicker);
    $('#import-pool-top')?.addEventListener('click',()=>$('#pool-file').click());
    $('#export-pool-top')?.addEventListener('click',()=>guardAsync(exportProductPool));
    $('#export-pool')?.addEventListener('click',()=>guardAsync(exportProductPool));
    $('#import-pool')?.addEventListener('click',()=>$('#pool-file').click());
    $('#pool-file')?.addEventListener('change',e=>{const f=e.target.files&&e.target.files[0];e.target.value='';if(f)poolImport(f);});
    $('#import-catalog')?.addEventListener('click',()=>$('#catalog-file').click());
    $('#catalog-file')?.addEventListener('change',e=>catalogImport(e.target.files[0]));
    $$('[data-template]').forEach(b=>b.onclick=()=>templateModal(b.dataset.template));
    $('#import-local')?.addEventListener('click', () => confirmDialog('导入本机数据', '将把本机浏览器里保存的工作区写入云端，覆盖当前云端空数据。继续？', async () => {
      const local = await OrderStore.readLocal();
      if (!local || !local.state) { toast('本机没有找到单机版数据。', 'warning'); return; }
      state = local.state;
      window.__localImportDone = true;
      try { await OrderStore.save(copy(state)); } catch (e) { toast(e && e.message ? e.message : '导入失败。', 'warning'); return; }
      render(); toast('本机数据已导入云端。');
    }));
    $('#export-unified')?.addEventListener('click', () => guardAsync(exportUnifiedOrders));
    $('#clear-orders')?.addEventListener('click',()=>confirmDialog('清空全部订单',`将删除 ${state.orders.length} 条明细和 ${state.batches.length} 个批次，供应商库不受影响。`,async()=>{await deleteOrders(state.orders.map(r=>r.id),'订单已清空。');}));
    $$('[data-export-group]').forEach(b=>b.onclick=()=>{selectedExport=b.dataset.exportGroup;render();});
    $('#export-current')?.addEventListener('click',()=>guardAsync(()=>exportOne(exportGroups().find(g=>g.supplier.id===selectedExport))));
    $('#export-all')?.addEventListener('click',()=>guardAsync(exportAll));
  }
  const guardAsync=async fn=>{try{await fn();}catch(e){toast(e.message||'操作失败，请重试。','warning');}};
  function bindOrderActions(){$$('[data-del-order]').forEach(b=>b.onclick=()=>confirmDialog('删除订单明细','确认删除这一条明细？删除后不可恢复。',async()=>{await deleteOrders([b.dataset.delOrder],'已删除 1 条明细。');}));$$('[data-edit-order]').forEach(b=>b.onclick=()=>orderModal(b.dataset.editOrder));$$('[data-page]').forEach(b=>b.onclick=()=>{page=Number(b.dataset.page);$('#orders-table').innerHTML=orderTable();bindOrderActions();});}
  function bindProductActions(){$$('[data-edit-product]').forEach(b=>b.onclick=()=>productModal(b.dataset.editProduct));$$('[data-ppage]').forEach(b=>b.onclick=()=>{productPage=Number(b.dataset.ppage);$('#products-table').innerHTML=productsTable();bindProductActions();});}
  async function queueFiles(files){if(!files?.length)return;const list=Array.from(files);if(list.length>30){toast('每次最多导入30个文件。','warning');return;}let failed=0;for(const f of list){try{if(f.size>15*1024*1024)throw Error('文件超过15MB');const ab=await f.arrayBuffer();const hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',ab))).map(v=>v.toString(16).padStart(2,'0')).join('');if(state.batches.some(b=>b.hash===hash)||pending.some(p=>p.hash===hash)){toast(`已跳过重复文件：${f.name}`,'warning');continue;}const book=await OrderExcel.read(f);pending.push({book,hash});}catch(e){failed++;toast(`${f.name}：${e.message}`,'warning');}}if(pending.length)nextImport();else if(!failed)toast('所选文件已导入过。','warning');}
  function inferCustomer(name){if(name.includes('兮迪'))return '兮迪';if(/GBH/i.test(name))return 'GBH';return name.replace(/\.(xlsx|csv)$/i,'').replace(/\s+/g,' ').trim().slice(0,50);}
  function autoMap(headers){
    const out={};
    for(const [key,,names]of FIELDS){
      let best=-1,bestScore=0;
      (headers||[]).forEach((h,i)=>{
        const t=norm(h);if(!t)return;
        for(const n of names){
          const na=norm(n);if(!na)continue;
          const single=[...na].length===1; // 单字别名（省/市/区）只接受完全相同，避免吞掉整段地址
          const score=t===na?3:(!single&&t.startsWith(na)?2:(!single&&t.includes(na)?1:0));
          if(score>bestScore){bestScore=score;best=i;}
        }
      });
      out[key]=best<0?'':String(best);
    }
    return out;
  }
  function nextImport(){if(!pending.length){render();return;}const p=pending[0];const s=p.book.sheets.find(s=>s.rows.length>1)||p.book.sheets[0];const idx=p.book.sheets.indexOf(s);const header=OrderExcel.detectHeader(s.rows);importDraft={...p,sheet:idx,header,customer:inferCustomer(p.book.name),map:{}};showImportModal(true);}
  function showImportModal(reset=false){const d=importDraft,s=d.book.sheets[d.sheet],headers=s.rows[d.header]||[];if(reset){const signature=headers.map(norm).join('|');const auto=autoMap(headers),saved=state.profiles[signature]||{};d.map={};for(const [k] of FIELDS)d.map[k]=(saved[k]===undefined||saved[k]==='')?auto[k]:saved[k];}
    openModal(`确认导入字段 · ${pending.length} 个文件待处理`,`<div class="alert info">${icon('file')} ${esc(d.book.name)} · 请选择客户及正确的表头行。文件中的说明文字不会执行。</div><form id="import-form"><div class="form-grid">${field('客户名称','customer',d.customer,{required:true})}${selectField('工作表','sheet',d.book.sheets.map((s,i)=>[i,`${s.name}（${Math.max(0,s.rows.length-1)} 行）`]),d.sheet)}${field('表头在第几行','header',d.header+1,{type:'number',hint:'选择前20行内的表头'})}<div class="field"><span>预计读取</span><div class="read-count">${Math.max(0,s.rows.length-d.header-1)} 行数据</div></div></div><h3 class="section-label">来源字段对应关系</h3><div class="form-grid">${FIELDS.map(([key,label])=>selectField(label,'map-'+key,[['','不读取 / 未提供'],...headers.map((h,i)=>[String(i),`${String.fromCharCode(65+Math.floor(i/26)-1).replace('@','')}${String.fromCharCode(65+i%26)} · ${h||'空白列'}`])],d.map[key]??'')).join('')}</div><div class="alert warning">“订单商品编号”是订单明细 ID。商品映射请使用“客户商品编码”或商品名称与规格。</div><h3 class="section-label">来源预览（前2行）</h3><div class="table-wrap compact"><table class="data-table"><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${s.rows.slice(d.header+1,d.header+3).map(r=>`<tr>${headers.map((_,i)=>`<td>${esc(r[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div><p id="import-error" class="form-error" role="alert"></p></form>`,`<button class="btn secondary" id="skip-import">跳过此文件</button><button class="btn" id="confirm-import">确认导入并匹配</button>`,true);
    const capture=()=>{d.customer=$('[name=customer]').value.trim();FIELDS.forEach(([k])=>d.map[k]=$(`[name="map-${k}"]`).value);};
    $('[name=sheet]').onchange=e=>{capture();d.sheet=Number(e.target.value);d.header=OrderExcel.detectHeader(d.book.sheets[d.sheet].rows);showImportModal(true);};
    $('[name=header]').onchange=e=>{capture();const n=Number(e.target.value);if(!Number.isInteger(n)||n<1||n>Math.min(20,s.rows.length)){toast('请输入有效表头行（1至20）。','warning');e.target.value=d.header+1;return;}d.header=n-1;showImportModal(true);};
    $('#skip-import').onclick=()=>{pending.shift();closeModal();nextImport();};
    $('#confirm-import').onclick=async()=>{capture();if(!d.customer){$('#import-error').textContent='请填写客户名称。';return;}if(d.map.productName===''){ $('#import-error').textContent='请至少指定商品名称字段。';return;}if(d.map.sourceSku!==''&&norm(headers[Number(d.map.sourceSku)])===norm('订单商品编号')){$('#import-error').textContent='订单商品编号不能作为稳定商品编码，请改选或留空。';return;}
      const batchId=uid(),orders=[],get=(r,k)=>d.map[k]===''?'':String(r[Number(d.map[k])]??'').trim();
      s.rows.slice(d.header+1).forEach((raw,idx)=>{if(!raw.some(v=>String(v??'').trim()))return;if(headers.every((h,i)=>String(h??'')===String(raw[i]??'')))return;const row={id:uid(),batchId,customer:d.customer,sourceFile:d.book.name,sourceSheet:s.name,sourceRow:d.header+idx+2,excluded:false,manual:false,productId:''};for(const[k]of FIELDS)row[k]=get(raw,k);row.phone=normalizePhone(row.phone);row.address=[row.province,row.city,row.district,row.address].filter(Boolean).join('');row.quantity=row.quantity===''?'':Number(row.quantity.replace(/,/g,''));if(!Number.isFinite(row.quantity))row.quantity='';if(/^\d{5}(\.\d+)?$/.test(String(row.date||'').trim()))row.date=excelSerialText(row.date);row.raw=Object.fromEntries(headers.map((h,i)=>[String(h||`列${i+1}`),String(raw[i]??'')]));const existing=[...state.orders,...orders].find(o=>!o.excluded&&duplicateKey(o)===duplicateKey(row));if(existing&&row.orderNo)row.duplicateOf=existing.id;orders.push(row);});
      if(!orders.length){$('#import-error').textContent='所选表头下没有可导入的数据。';return;}if(state.orders.length+orders.length>30000){$('#import-error').textContent='工作区最多保留30,000行，请先备份并清理旧批次。';return;}
      ensureCustomer(d.customer);state.orders.push(...orders);state.batches.push({id:batchId,name:d.book.name,hash:d.hash,customer:d.customer,createdAt:new Date().toISOString(),rowCount:orders.length});includeNewBatch(batchId);state.profiles[headers.map(norm).join('|')]=copy(d.map);invalidate();history(`导入 ${orders.length} 行：${d.book.name}`);await save();pending.shift();closeModal();toast(`已导入 ${orders.length} 条订单明细。`);nextImport();};
  }
  function duplicateKey(r){return [norm(r.customer),String(r.orderNo||''),r.sourceLineId||[norm(r.sourceSku),norm(r.productName),norm(r.sourceSpec),norm(r.recipient),norm(r.phone),norm(r.address)].join('|'),r.duplicateDiscriminator||''].join('::');}
  function resolveDuplicateChoice(row,independent){
    const wasDuplicate=!!row.duplicateOf;
    if(independent)row.duplicateDiscriminator=row.duplicateDiscriminator||uid();else delete row.duplicateDiscriminator;
    if(row.excluded)return;
    const key=duplicateKey(row);
    if(wasDuplicate&&!independent){for(const other of state.orders)if(other.id!==row.id&&duplicateKey(other)===key)other.excluded=true;delete row.duplicateOf;}
    else if(independent)delete row.duplicateOf;
    else{const other=state.orders.find(r=>r.id!==row.id&&!r.excluded&&duplicateKey(r)===key);if(other)row.duplicateOf=other.id;}
  }
  function orderModal(id){const row=state.orders.find(r=>r.id===id);if(!row)return;const s=status(row),p=s.product||suggestion(row);openModal('核对订单与商品映射',`<div class="order-source"><span class="pill neutral">${esc(row.customer)}</span><span>${esc(row.sourceFile)} · ${esc(row.sourceSheet)} 第 ${row.sourceRow} 行</span></div>${s.issues.length?`<div class="alert warning">${esc(s.issues.join('；'))}</div>`:''}<form id="order-form"><div class="form-grid">${field('订单号（可空）','orderNo',row.orderNo)}${field('客户名称','customer',row.customer,{required:true})}${field('来源商品名称','productName',row.productName,{required:true,full:true})}${field('来源规格','sourceSpec',row.sourceSpec)}${field('客户商品编码','sourceSku',row.sourceSku)}${selectField('确认对应的标准商品及供应商','productId',[['','请选择商品'],...state.products.map(p=>[p.id,`${sById(p.supplierId)?.name} · ${p.supplierSku} · ${p.name}`])],p?.id||'',true)}${p&&!s.product?`<div class="field full"><small class="amber-text">根据样表给出候选，请核对款式、包装规格和供应商后确认。</small></div>`:''}<label class="checkbox-row field full"><input type="checkbox" name="remember" checked><span>将此客户的商品名称与规格保存为映射，下次复用</span></label>${field('数量','quantity',row.quantity,{type:'number',required:true})}${field('收货人','recipient',row.recipient,{required:true})}${field('收货电话','phone',row.phone,{required:true})}${field('买家昵称（可空）','buyerNickname',row.buyerNickname)}${field('完整收货地址','address',row.address,{required:true,full:true})}${field('备注','remark',row.remark,{full:true})}<label class="checkbox-row field full"><input type="checkbox" name="excluded" ${row.excluded?'checked':''}><span>排除此行，不参与导出</span></label>${row.duplicateOf||row.duplicateDiscriminator?`<label class="checkbox-row field full"><input type="checkbox" name="independent" ${row.duplicateDiscriminator?'checked':''}><span>作为独立明细保留（确认这不是修订版本）</span></label>`:''}</div>${row.duplicateOf?'<div class="alert warning">这是重复或修订版本。默认保留本行并排除同组其他版本；真实独立订单行请勾选“作为独立明细保留”。</div>':''}<p class="form-error" id="order-error"></p><details><summary>查看原始字段</summary><div class="raw-grid">${Object.entries(row.raw||{}).map(([k,v])=>`<div><strong>${esc(k)}</strong><span>${esc(v)}</span></div>`).join('')}</div></details></form>`,`<button class="btn secondary" data-cancel>取消</button><button class="btn danger" id="delete-order">删除明细</button><button class="btn" id="save-order">保存并重新检查</button>`,true);
        $('#delete-order')?.addEventListener('click',()=>confirmDialog('删除订单明细','确认删除这一条明细？删除后不可恢复。',async()=>{closeModal();await deleteOrders([row.id],'已删除 1 条明细。');}));
$('#save-order').onclick=async()=>{const f=new FormData($('#order-form'));const product=pById(f.get('productId'));const excluded=f.has('excluded');if(!excluded&&!product){$('#order-error').textContent='请选择标准商品，或将本行排除。';return;}const keys=['orderNo','customer','productName','sourceSpec','sourceSku','recipient','phone','buyerNickname','address','remark'];keys.forEach(k=>row[k]=String(f.get(k)||'').trim());row.phone=normalizePhone(row.phone);ensureCustomer(row.customer);row.quantity=Number(f.get('quantity'));row.excluded=excluded;row.productId=product?.id||'';row.manual=!!product;delete row.exportedAt;
      resolveDuplicateChoice(row,f.has('independent'));
      if(product&&f.has('remember')&&!excluded){const alias={customer:row.customer,name:row.productName,spec:row.sourceSpec,code:''};if(!product.aliases.some(a=>JSON.stringify(a)===JSON.stringify(alias)))product.aliases.push(alias);}
      refreshBatchExportFlags();invalidate();history('确认订单映射与字段');await save();closeModal();render();toast('已保存，匹配状态已更新。');};
  }
  function supplierModal(id){const s=copy(sById(id)||{id:uid(),name:'',contact:'',phone:'',note:'',template:null,bankName:'',bankCode:'',bankAccount:'',payAccount:''});openModal(id?'编辑供应商':'新增供应商',`<form id="supplier-form"><div class="form-grid">${field('供应商名称','name',s.name,{required:true,full:true})}${field('联系人','contact',s.contact)}${field('联系电话','phone',s.phone)}${field('供应商银行名称','bankName',s.bankName||'')}${field('供应商开户行行号','bankCode',s.bankCode||'')}${field('供应商银行账号','bankAccount',s.bankAccount||'')}${field('供应商付款账号','payAccount',s.payAccount||'',{hint:'默认与银行账号相同'})}${field('备注 / 供应商主体信息','note',s.note,{textarea:true,full:true})}</div><p id="supplier-error" class="form-error"></p></form>`,`<button class="btn secondary" data-cancel>取消</button>${id?'<button class="btn danger" id="delete-supplier">删除供应商</button>':''}<button class="btn" id="save-supplier">保存供应商</button>`);
    $('#save-supplier').onclick=async()=>{const f=new FormData($('#supplier-form'));for(const[k,v]of f)s[k]=v.trim();if(!s.name){$('#supplier-error').textContent='请输入供应商名称。';return;}if(state.suppliers.some(x=>x.id!==s.id&&norm(x.name)===norm(s.name))){$('#supplier-error').textContent='供应商名称已存在。';return;}const idx=state.suppliers.findIndex(x=>x.id===s.id);idx<0?state.suppliers.push(s):state.suppliers.splice(idx,1,s);invalidate();await save();closeModal();render();toast('供应商已保存。');};
    $('#delete-supplier')?.addEventListener('click',()=>{const owned=state.products.filter(p=>p.supplierId===s.id),ids=new Set(owned.map(p=>p.id));confirmDialog('删除供应商',owned.length?`“${s.name}”下有 ${owned.length} 个商品，会一并删除；已引用这些商品的订单明细需要重新映射。`:`确认删除“${s.name}”？`,async()=>{state.suppliers=state.suppliers.filter(x=>x.id!==s.id);state.products=state.products.filter(p=>p.supplierId!==s.id);state.recharges=state.recharges.filter(r=>(r.kind||'customer')!=='supplier'||norm(r.name||r.customer||'')!==norm(s.name));state.orders.forEach(r=>{if(ids.has(r.productId)){r.productId='';r.manual=false;}});if(supplierFilter===s.id)supplierFilter='all';if(selectedExport===s.id)selectedExport='';invalidate();await save();render();toast(owned.length?`已删除供应商及其 ${owned.length} 个商品。`:'已删除供应商。');});});
  }
  function productModal(id){if(!state.suppliers.length){toast('请先新增供应商。','warning');return;}const p=copy(pById(id)||{id:uid(),internalSku:nextProductSku(),name:'',spec:'',netContent:'',image:'',unit:'盒',supplierId:supplierFilter==='all'?state.suppliers[0].id:supplierFilter,supplierSku:'',barcode:'',brand:'',category:'',cost:'',groupPrice:'',broadcastPrice:'',suggestPrice:'',marginOwn:'',marginBroadcast:'',jdPrice:'',link:'',aliases:[],note:'',shipping:'',taxRate:'',taxCode:'',shelfLife:'',origin:'',shipFrom:'',size:'',remoteExcluded:'否',status:'合作中',minQty:'',createdAt:''});openModal(id?'编辑标准商品':'新增标准商品',`<form id="product-form"><div class="form-grid">${field('标准商品名称','name',p.name,{required:true,full:true})}${field('内部商品编码','internalSku',p.internalSku,{hint:'留空自动生成，例如 OF-012'})}${field('内容物配置','spec',p.spec,{hint:'例如 40枚/盒'})}${field('净含量','netContent',p.netContent||'',{hint:'例如 200g'})}${field('计量单位','unit',p.unit)}${selectField('供应商','supplierId',state.suppliers.map(s=>[s.id,s.name]),p.supplierId)}${field('供应商商品编码','supplierSku',p.supplierSku,{hint:'留空则与内部编码相同'})}${field('69码 / 条码','barcode',p.barcode||'')}${field('品牌','brand',p.brand||'')}${field('商品品类','category',p.category||'')}${field('保质期','shelfLife',p.shelfLife||'',{hint:'例如 12 个月'})}${field('产地','origin',p.origin||'')}${field('发货地','shipFrom',p.shipFrom||'')}${field('产品尺寸','size',p.size||'',{hint:'例如 20×15×8cm'})}${field('图片链接','image',p.image||'',{full:true})}${field('集采价','groupPrice',p.groupPrice||'',{type:'number'})}${field('集采起订量','minQty',p.minQty||'',{type:'number'})}${field('一件代发价格','cost',p.cost||'',{type:'number',hint:'供应商单件代发价，用于供应商对账'})}${field('广电供价','broadcastPrice',p.broadcastPrice||'',{type:'number'})}${field('零售价','suggestPrice',p.suggestPrice||'',{type:'number'})}${field('定宣毛利（0.2 = 20%）','marginOwn',p.marginOwn||'',{type:'number'})}${field('广电毛利（0.2 = 20%）','marginBroadcast',p.marginBroadcast||'',{type:'number'})}${field('京东价','jdPrice',p.jdPrice||'',{type:'number'})}${field('快递','shipping',p.shipping||'',{hint:'例如 中通/圆通'})}${field('税率','taxRate',p.taxRate||'',{hint:'例如 13%'})}${field('税收编码','taxCode',p.taxCode||'')}${selectField('偏远不发','remoteExcluded',[['否','否'],['是','是']],p.remoteExcluded||'否')}${selectField('合作状态','status',[['合作中','合作中'],['暂停','暂停'],['终止','终止']],p.status||'合作中')}${field('京东/天猫链接','link',p.link||'',{full:true})}${field('新增通用商品别名（每行一个）','newAliases','',{textarea:true,full:true,hint:'精确匹配；不同包装规格请分别建商品。客户专属映射在订单处理时确认。'})}</div><h3 class="section-label">已保存的客户别名</h3><div id="alias-list">${p.aliases.map((a,i)=>`<div class="alias-item"><span><strong>${esc(a.customer||'通用')}</strong> ${esc(a.name||a.code)}<small>${esc(a.spec||'')}</small></span><button type="button" class="text-link" data-remove-alias="${i}">移除</button></div>`).join('')||'<p class="muted">尚无客户别名</p>'}</div><p class="form-error" id="product-error"></p></form>`,`<button class="btn secondary" data-cancel>取消</button>${id?'<button class="btn danger" id="delete-product">删除商品</button>':''}<button class="btn" id="save-product">保存商品</button>`,true);
    $$('[data-remove-alias]').forEach(b=>b.onclick=()=>{p.aliases[Number(b.dataset.removeAlias)]=null;b.closest('.alias-item').remove();});
    $('#save-product').onclick=async()=>{const f=new FormData($('#product-form'));['name','internalSku','spec','unit','supplierId','supplierSku','barcode','brand','category','cost','groupPrice','broadcastPrice','suggestPrice','marginOwn','marginBroadcast','jdPrice','link','netContent','image','shipping','taxRate','taxCode','shelfLife','origin','shipFrom','size','remoteExcluded','status','minQty'].forEach(k=>p[k]=String(f.get(k)||'').trim());if(!p.name){$('#product-error').textContent='请填写商品名称。';return;}if(!p.createdAt)p.createdAt=new Date().toISOString();p.internalSku=p.internalSku||nextProductSku();p.supplierSku=p.supplierSku||p.internalSku;if(state.products.some(x=>x.id!==p.id&&(norm(x.internalSku)===norm(p.internalSku)||x.supplierId===p.supplierId&&x.supplierSku===p.supplierSku))){$('#product-error').textContent='内部编码或同供应商的商品编码已存在。';return;}p.aliases=p.aliases.filter(Boolean);String(f.get('newAliases')||'').split('\n').map(x=>x.trim()).filter(Boolean).forEach(name=>{if(!p.aliases.some(a=>!a.customer&&a.name===name))p.aliases.push({customer:'',name,spec:'',code:''});});const idx=state.products.findIndex(x=>x.id===p.id);idx<0?state.products.push(p):state.products.splice(idx,1,p);invalidate();await save();closeModal();render();toast('商品与映射已保存。');};
    $('#delete-product')?.addEventListener('click',()=>confirmDialog('删除商品',`删除“${p.name}”后，相关订单需要重新映射。`,async()=>{state.products=state.products.filter(x=>x.id!==p.id);invalidate();await save();render();}));
  }
  function reviewCandidates(){const cs=OrderSeed.candidates.filter(c=>pById(c.productId));openModal('确认样表中的商品映射',`<p class="muted">供应商名称取自友本模板。请核对实际供货方及商品款式；勾选后会保存为对应客户的别名。</p>${cs.map((c,i)=>{const p=pById(c.productId);return `<label class="candidate-row"><input type="checkbox" name="candidate" value="${i}" checked><span><small>${esc(c.customer)} · 客户商品名称</small><strong>${esc(c.name)}</strong><span class="mapping-arrow">${icon('arrow')} ${esc(p.name)}</span><small>${esc(sById(p.supplierId)?.name)} · ${esc(p.supplierSku)} · ${esc(c.reason)}</small></span></label>`;}).join('')}`,`<button class="btn secondary" data-cancel>暂不确认</button><button class="btn" id="confirm-candidates">确认所选映射</button>`,true);$('#confirm-candidates').onclick=async()=>{let n=0;$$('[name=candidate]:checked').forEach(el=>{const c=cs[Number(el.value)],p=pById(c.productId);const a={customer:c.customer,name:c.name,spec:'',code:''};const existing=p.aliases.find(x=>x.customer===a.customer&&x.name===a.name);if(existing){existing.spec='';}else{p.aliases.push(a);n++;}});invalidate();await save();closeModal();render();toast(`已确认 ${n} 条新别名映射。`);};}
  async function catalogImport(file){if(!file)return;await guardAsync(async()=>{const book=await OrderExcel.read(file);let selected=book.sheets.findIndex(s=>s.name.includes('商品编码'));if(selected<0)selected=0;let d={sheet:selected,header:OrderExcel.detectHeader(book.sheets[selected].rows)};const show=()=>{const s=book.sheets[d.sheet],hs=s.rows[d.header]||[],choose=terms=>String(hs.findIndex(h=>terms.includes(norm(h))));openModal('导入供应商商品资料',`<div class="alert info">仅导入商品编码与名称，不导入模板内的示例订单。</div><form id="catalog-form"><div class="form-grid">${selectField('供应商','supplier',state.suppliers.map(s=>[s.id,s.name]),state.suppliers[0]?.id)}${selectField('工作表','sheet',book.sheets.map((s,i)=>[i,s.name]),d.sheet)}${field('表头所在行','header',d.header+1,{type:'number'})}${selectField('供应商商品编码列','sku',hs.map((h,i)=>[i,h||`列${i+1}`]),choose(['商品编码','供应商商品编码','sku码','sku']))}${selectField('商品名称列','name',hs.map((h,i)=>[i,h||`列${i+1}`]),choose(['名称','商品名称','产品名称']))}</div><p class="muted">同一供应商已有的商品编码将跳过。预计 ${Math.max(0,s.rows.length-d.header-1)} 行。</p><p class="form-error" id="catalog-error"></p></form>`,`<button class="btn secondary" data-cancel>取消</button><button class="btn" id="save-catalog">导入商品资料</button>`);$('[name=sheet]').onchange=e=>{d.sheet=Number(e.target.value);d.header=OrderExcel.detectHeader(book.sheets[d.sheet].rows);show();};$('[name=header]').onchange=e=>{const n=Number(e.target.value);if(Number.isInteger(n)&&n>0&&n<=Math.min(s.rows.length,20)){d.header=n-1;show();}else toast('请输入1到20之间的有效表头行。','warning');};$('#save-catalog').onclick=async()=>{const f=new FormData($('#catalog-form')),sid=f.get('supplier');let count=0;for(const r of s.rows.slice(d.header+1)){const sku=String(r[Number(f.get('sku'))]||'').trim(),name=String(r[Number(f.get('name'))]||'').trim();if(!sku||!name||state.products.some(p=>p.supplierId===sid&&p.supplierSku===sku))continue;state.products.push({id:uid(),supplierId:sid,supplierSku:sku,name,spec:name.match(/\d+枚(?:\/盒)?/)?.[0]||'',unit:'',internalSku:`OF-${String(state.products.length+1).padStart(3,'0')}`,aliases:[]});count++;}invalidate();await save();closeModal();render();toast(`已导入 ${count} 个新商品，重复编码已跳过。`);};};show();});}
  const QUOTE_FIELDS = [
    ['category', '商品品类', ['商品品类', '品类', '分类', '类别', '类目', '商品分类']],
    ['brand', '品牌', ['品牌', '品牌名称', '商标']],
    ['barcode', '69码 / 条码', ['69码', '条码', '条形码', '商品条码', '国际条码', 'ean13', 'ean', '69条码']],
    ['name', '商品名称', ['商品名称', '产品名称', '品名', '名称', '货品名称', '商品']],
    ['supplierSku', '商品编码', ['商品编码', '供应商商品编码', '商品编号', '货号', 'sku码', 'sku', '编码']],
    ['groupPrice', '集采价', ['集采价', '集采价格', '集采']],
    ['cost', '供价 / 进货价', ['供价', '供货价', '进货价', '成本价', '采购价', '底价', '一件代发价']],
    ['broadcastPrice', '广电供价', ['广电供价', '广电供货价']],
    ['suggestPrice', '建议售价', ['建议售价', '建议零售价', '零售价', '售价', '指导价']],
    ['marginOwn', '定宣毛利%', ['定宣毛利', '定宣毛利率']],
    ['marginBroadcast', '广电毛利%', ['广电毛利', '广电毛利率']],
    ['jdPrice', '京东价 / 市场价', ['京东价', '京东售价', '市场价', '吊牌价', '划线价']],
    ['link', '商品链接', ['商品京东链接', '商品链接', '链接', '网址', 'url', '天猫链接']],
    ['supplier', '供应商', ['供应商', '供货商', '厂家', '供应方', '供应商名称']],
    ['spec', '规格', ['规格', '规格型号', '包装规格', '净含量']],
    ['unit', '单位', ['单位', '计量单位']]
  ];
  function matchColumn(headers, aliases) {
    let best = -1, bestScore = 0;
    headers.forEach((h, i) => {
      const t = norm(h); if (!t) return;
      for (const a of aliases) {
        const na = norm(a); let s = 0;
        if (t === na) s = 3; else if (t.startsWith(na)) s = 2; else if (t.includes(na)) s = 1;
        if (s > bestScore) { bestScore = s; best = i; }
      }
    });
    return best;
  }
  function quoteAutoMap(headers) {
    const out = {};
    for (const [key, , aliases] of QUOTE_FIELDS) out[key] = matchColumn(headers || [], aliases);
    return out;
  }
  function moneyValue(v) {
    const t = String(v == null ? '' : v).replace(/[¥￥,，\s]/g, '');
    if (!t) return '';
    const m = t.match(/-?\d+(?:\.\d+)?/);
    return m ? String(Number(m[0])) : '';
  }
  function ratioValue(v) {
    const t = String(v == null ? '' : v).replace(/[%％,，\s]/g, '');
    if (!t) return '';
    const n = Number(t);
    if (!Number.isFinite(n)) return '';
    const ratio = Math.abs(n) > 1 ? n / 100 : n;
    return String(Math.round(ratio * 1e6) / 1e6);
  }
  function percentText(v) {
    if (v === '' || v == null) return '';
    const n = Number(v);
    return Number.isFinite(n) ? (n * 100).toFixed(n * 100 % 1 ? 1 : 0) + '%' : String(v);
  }
  function barcodeValue(v) {
    const t = String(v == null ? '' : v).replace(/[\s'"]/g, '').replace(/\.0+$/, '');
    return /^\d{8,14}$/.test(t) ? t : '';
  }
  function quoteRecords(book, sheetIndex, headerIndex, map) {
    const rows = book.sheets[sheetIndex]?.rows || [];
    const get = (row, key) => { const i = map[key]; return i == null || i < 0 ? '' : String(row?.[i] == null ? '' : row[i]).replace(/\s+/g, ' ').trim(); };
    const records = [];
    for (let r = headerIndex + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const name = get(row, 'name'); if (!name) continue;
      records.push({
        row: r + 1, name,
        category: get(row, 'category'), brand: get(row, 'brand'),
        barcode: barcodeValue(get(row, 'barcode')),
        supplierSku: get(row, 'supplierSku').replace(/\.0+$/, ''),
        groupPrice: moneyValue(get(row, 'groupPrice')), cost: moneyValue(get(row, 'cost')),
        broadcastPrice: moneyValue(get(row, 'broadcastPrice')),
        suggestPrice: moneyValue(get(row, 'suggestPrice')),
        marginOwn: ratioValue(get(row, 'marginOwn')), marginBroadcast: ratioValue(get(row, 'marginBroadcast')),
        jdPrice: moneyValue(get(row, 'jdPrice')), link: get(row, 'link'),
        supplier: get(row, 'supplier'), spec: get(row, 'spec'), unit: get(row, 'unit')
      });
    }
    return records;
  }
  function quoteGroups(records, fallbackName, fallbackDerived) {
    const map = new Map();
    for (const r of records) {
      const named = r.supplier.trim();
      const name = named || fallbackName || '未标注供应商';
      const key = norm(name) || '未标注供应商';
      if (!map.has(key)) map.set(key, { name, derived: !named && !!fallbackDerived, manual: !named && !fallbackDerived, rows: [] });
      map.get(key).rows.push(r);
    }
    return [...map.values()];
  }
  function quoteGroupCards(groups) {
    return groups.map(g => `<div class="quote-group"><span class="supplier-avatar">${esc(g.name.slice(0, 1))}</span><span class="quote-group-body"><strong>${esc(g.name)}${g.derived ? '<i class="tile-flag">品牌推断</i>' : ''}${g.manual ? '<i class="tile-flag">手动指定</i>' : ''}</strong><small>${g.rows.length} 个商品${g.brands.length ? ' · ' + esc(g.brands.slice(0, 4).join('、')) : ''}${g.min != null ? ` · 供价 ¥${g.min}${g.max > g.min ? '–' + g.max : ''}` : ''}</small></span></div>`).join('') || '<p class="muted">没有可导入的商品行。</p>';
  }
  function quoteSummary(groups) {
    return groups.map(g => {
      const brands = [...new Set(g.rows.map(r => r.brand).filter(Boolean))];
      const costs = g.rows.map(r => Number(r.cost)).filter(n => Number.isFinite(n) && n > 0);
      return { ...g, brands, min: costs.length ? Math.min(...costs) : null, max: costs.length ? Math.max(...costs) : null };
    }).sort((a, b) => b.rows.length - a.rows.length);
  }
  async function applyQuote(groups, fileName) {
    let newSuppliers = 0, added = 0, updated = 0;
    const byName = new Map(state.suppliers.map(s => [norm(s.name), s]));
    for (const g of groups) {
      let sup = byName.get(norm(g.name));
      if (!sup) {
        sup = { id: uid(), name: g.name, contact: '', phone: '', note: g.manual ? '报价单供应商列为空，导入时手动指定，请核对供货主体与联系人。' : (g.derived ? '报价单未填写供应商，请补充实际供货主体与联系人。' : '由报价单导入，请补充联系人与结算信息。'), template: null, derived: !!g.derived };
        state.suppliers.push(sup); byName.set(norm(g.name), sup); newSuppliers++;
      } else if (g.derived && !sup.derived) { /* keep an explicitly named supplier as-is */ }
      const owned = state.products.filter(p => p.supplierId === sup.id);
      const usedCodes = new Set(owned.map(p => p.supplierSku).filter(Boolean));
      for (const r of g.rows) {
        const key = r.barcode || r.supplierSku;
        const existing = owned.find(p => norm(p.name) === norm(r.name) && (!key || p.barcode === key || p.supplierSku === key));
        const internalSku = existing?.internalSku || nextProductSku();
        let code = r.supplierSku || r.barcode || existing?.supplierSku || internalSku;
        // One 69码 can legitimately cover a set and a single unit; keep the code unique per supplier.
        if (!existing && usedCodes.has(code)) code = internalSku;
        usedCodes.add(code);
        const payload = {
          name: r.name,
          brand: r.brand || existing?.brand || '',
          category: r.category || existing?.category || '',
          barcode: r.barcode || existing?.barcode || '',
          supplierSku: code,
          groupPrice: r.groupPrice !== '' ? r.groupPrice : (existing?.groupPrice || ''),
          cost: r.cost !== '' ? r.cost : (existing?.cost || ''),
          broadcastPrice: r.broadcastPrice !== '' ? r.broadcastPrice : (existing?.broadcastPrice || ''),
          marginOwn: r.marginOwn !== '' ? r.marginOwn : (existing?.marginOwn || ''),
          marginBroadcast: r.marginBroadcast !== '' ? r.marginBroadcast : (existing?.marginBroadcast || ''),
          suggestPrice: r.suggestPrice !== '' ? r.suggestPrice : (existing?.suggestPrice || ''),
          jdPrice: r.jdPrice !== '' ? r.jdPrice : (existing?.jdPrice || ''),
          link: r.link || existing?.link || '',
          spec: r.spec || existing?.spec || '',
          unit: r.unit || existing?.unit || '',
          sourceFile: fileName
        };
        if (existing) { Object.assign(existing, payload); updated++; }
        else {
          const created = { id: uid(), internalSku, supplierId: sup.id, aliases: [], note: r.barcode ? '' : '报价单未提供条码，编码为内部占位，请向供应商补齐。', ...payload };
          state.products.push(created); owned.push(created);
          added++;
        }
      }
    }
    invalidate(); history(`导入报价单：${fileName}`); await save();
    return { newSuppliers, added, updated };
  }
  async function quoteImport(file) {
    if (!file) return;
    await guardAsync(async () => {
      const book = await OrderExcel.read(file);
      let si = book.sheets.findIndex(s => s.rows.length > 2);
      if (si < 0) si = 0;
      const d = { sheet: si, header: OrderExcel.detectHeader(book.sheets[si].rows), map: {}, fallback: 'skip', manualName: '' };
      d.map = quoteAutoMap(book.sheets[si].rows[d.header]);
      const draw = () => {
        const sheet = book.sheets[d.sheet];
        const headers = sheet.rows[d.header] || [];
        const all = quoteRecords(book, d.sheet, d.header, d.map);
        const blankRows = all.filter(r => !r.supplier.trim()).length;
        const manualName = d.fallback === 'manual' ? d.manualName.trim() : '';
        const skipped = d.fallback === 'skip' ? blankRows : 0;
        const records = d.fallback === 'skip' || (d.fallback === 'manual' && !manualName) ? all.filter(r => r.supplier.trim()) : all;
        const groups = quoteSummary(quoteGroups(records, d.fallback === 'manual' ? manualName : '', d.fallback === 'none'));
        const missing = QUOTE_FIELDS.filter(([k]) => d.map[k] < 0 && ['name', 'supplier'].includes(k)).map(([, label]) => label);
        const unnamed = groups.filter(g => g.derived).reduce((n, g) => n + g.rows.length, 0);
        openModal('导入供应商报价单', `
          <div class="alert info">${icon('file')}<span id="quote-summary-text">已读取 <strong>${esc(file.name)}</strong> · ${book.sheets.length} 个工作表 · 识别 ${records.length} 条商品、${groups.length} 个供应商分组${skipped ? `，跳过 ${skipped} 条供应商为空` : ''}。</span></div>
          <div class="form-grid">
            ${selectField('工作表', 'q-sheet', book.sheets.map((s, i) => [i, `${s.name}（${s.rows.length} 行）`]), d.sheet)}
            ${field('表头所在行', 'q-header', d.header + 1, { type: 'number', hint: '第 1 行起算，通常为 1' })}
            ${selectField('供应商列为空的行', 'q-fallback', [['skip', '跳过，不导入（默认）'], ['manual', '手动指定供应商'], ['none', '归入「未标注供应商」']], d.fallback, true)}${d.fallback==='manual'?field('供应商名称','q-manual',d.manualName,{required:true,hint:'供应商列为空的明细都会归到这一家',full:true}):''}
          </div>
          <h3 class="section-label">列对应关系</h3>
          <div class="map-grid">${QUOTE_FIELDS.map(([k, label]) => `<label class="map-item"><span>${esc(label)}${k === 'name' ? ' *' : ''}</span><select name="q-map-${k}"><option value="-1">不导入</option>${headers.map((h, i) => `<option value="${i}" ${d.map[k] === i ? 'selected' : ''}>${esc(String(h == null || h === '' ? `列${i + 1}` : h).replace(/\s+/g, ' ').slice(0, 24))}</option>`).join('')}</select></label>`).join('')}</div>
          ${missing.length ? `<div class="alert warning">${icon('shield')}<span>未识别到这些列：${esc(missing.join('、'))}。可手动指定，或直接忽略。</span></div>` : ''}
          ${skipped ? `<div class="alert warning">${icon('shield')}<span>有 ${skipped} 条明细的「供应商」为空，按当前设置不会导入；把该列补齐后重新导入即可。</span></div>` : ''}
          ${!skipped && unnamed ? `<div class="alert info">${icon('link')}<span>有 ${unnamed} 条明细的供应商为空，将归入「未标注供应商」；导入后可在供应商卡片上改名或编辑。</span></div>` : ''}${!skipped && d.fallback === 'manual' ? `<div class="alert info">${icon('link')}<span>有 ${blankRows} 条明细的供应商为空，将归入下方手动指定的供应商。</span></div>` : ''}
          <h3 class="section-label" id="quote-group-title">将导入的供应商（${groups.length}）</h3>
          <div class="quote-groups" id="quote-groups">${quoteGroupCards(groups)}</div>
          <p class="muted">同名供应商自动合并；同供应商、同 69码 或同名称的商品会更新价格与链接，不会重复新建。</p>
          <p class="form-error" id="quote-error"></p>
        `, `<button class="btn secondary" data-cancel>取消</button><button class="btn" id="run-quote-import" ${records.length ? '' : 'disabled'}>导入 ${records.length} 个商品 / ${groups.length} 家供应商${skipped ? `（跳过 ${skipped} 行）` : ''}</button>`, true);
        $('[name="q-sheet"]').onchange = e => { d.sheet = Number(e.target.value); d.header = OrderExcel.detectHeader(book.sheets[d.sheet].rows); d.map = quoteAutoMap(book.sheets[d.sheet].rows[d.header]); draw(); };
        $('[name="q-header"]').onchange = e => { const n = Number(e.target.value); if (Number.isInteger(n) && n >= 1 && n <= Math.min(book.sheets[d.sheet].rows.length, 20)) { d.header = n - 1; d.map = quoteAutoMap(book.sheets[d.sheet].rows[d.header]); draw(); } else toast('表头行需在 1 到 20 之间。', 'warning'); };
        $('[name="q-fallback"]').onchange = e => { d.fallback = e.target.value; draw(); };
        $('[name="q-manual"]')?.addEventListener('input', e => {
          d.manualName = e.target.value;
          const all2 = quoteRecords(book, d.sheet, d.header, d.map);
          const manual = d.manualName.trim();
          const recs = manual ? all2 : all2.filter(r => r.supplier.trim());
          const gs = quoteSummary(quoteGroups(recs, manual, false));
          const title = $('#quote-group-title'); if (title) title.textContent = `将导入的供应商（${gs.length}）`;
          const box = $('#quote-groups'); if (box) box.innerHTML = quoteGroupCards(gs);
          const sum = $('#quote-summary-text');
          if (sum) sum.innerHTML = `已读取 <strong>${esc(file.name)}</strong> · ${book.sheets.length} 个工作表 · 识别 ${recs.length} 条商品、${gs.length} 个供应商分组。`;
          const btn = $('#run-quote-import');
          if (btn) { btn.textContent = `导入 ${recs.length} 个商品 / ${gs.length} 家供应商`; btn.disabled = !recs.length; }
        });
        for (const [k] of QUOTE_FIELDS) { const el = $(`[name="q-map-${k}"]`); if (el) el.onchange = e => { d.map[k] = Number(e.target.value); draw(); }; }
        $('#run-quote-import').onclick = () => guardAsync(async () => {
          const all = quoteRecords(book, d.sheet, d.header, d.map);
          const manualEl = $('[name="q-manual"]');
          if (manualEl) d.manualName = manualEl.value.trim();
          if (d.fallback === 'manual' && !d.manualName) { $('#quote-error').textContent = '请填写要归入的供应商名称，或把「供应商列为空的行」改回跳过。'; return; }
          if (d.fallback === 'manual') d.manualName = d.manualName;
          const rows = d.fallback === 'skip' || (d.fallback === 'manual' && !d.manualName) ? all.filter(r => r.supplier.trim()) : all;
          if (!rows.length) { $('#quote-error').textContent = '没有可导入的商品行，请检查列对应关系、表头行或供应商列。'; return; }
          const stat = await applyQuote(quoteSummary(quoteGroups(rows, d.fallback === 'manual' ? d.manualName : '', d.fallback === 'none')), file.name);
          closeModal(); render();
          toast(`已导入 ${stat.added} 个商品${stat.updated ? `，更新 ${stat.updated} 个` : ''}${stat.newSuppliers ? `，新增 ${stat.newSuppliers} 家供应商` : ''}${skipped ? `，跳过 ${skipped} 行` : ''}。`);
        });
      };
      draw();
    });
  }
  const SHIPMENT_FIELDS = [
    ['orderNo', '订单号列', ['订单编号', '订单号', '平台订单号', '订单编码', '订单号码']],
    ['recipient', '收货人', ['收货人', '收件人', '收货人姓名', '收件人姓名', '收货姓名']],
    ['phone', '收货电话', ['收货人号码', '收货人电话', '收货人手机', '收件人电话', '收件人手机', '手机', '手机号', '联系电话', '电话']],
    ['address', '收货地址', ['收货人地址', '收货地址', '详细地址', '收件地址', '地址']],
    ['product', '商品名称', ['商品名称', '产品名称', '品名', '货品名称', '商品']],
    ['carrier', '物流公司', ['物流公司', '快递公司', '承运商', '配送方式']],
    ['tracking', '快递单号', ['快递单号', '运单号', '物流单号', '快递编号', 'tracking', '追踪号']]
  ];
  let shipment = null;
  function shipmentAutoMap(headers) {
    const out = {};
    for (const [key, , aliases] of SHIPMENT_FIELDS) out[key] = matchColumn(headers || [], aliases);
    return out;
  }
  function shipmentOrders() {
    return activeRows().filter(r => String(r.orderNo || '').trim());
  }
  function shipmentMatch(rows, map) {
    const cell = (row, key) => { const i = map[key]; return i == null || i < 0 ? '' : String(row?.[i] ?? '').trim(); };
    const orders = shipmentOrders();
    const byPhoneName = new Map(), byPhone = new Map(), byName = new Map();
    const push = (m, k, v) => { if (!k) return; if (!m.has(k)) m.set(k, []); m.get(k).push(v); };
    for (const o of orders) {
      const phone = normalizePhone(o.phone), name = norm(o.recipient);
      if (phone && name) push(byPhoneName, phone + '|' + name, o);
      if (phone) push(byPhone, phone, o);
      if (name) push(byName, name, o);
    }
    const uniq = list => { if (!list || !list.length) return null; return new Set(list.map(o => o.id)).size === 1 ? list[0] : null; };
    return rows.map((row, index) => {
      const phone = normalizePhone(cell(row, 'phone'));
      const name = norm(cell(row, 'recipient'));
      const addr = norm(cell(row, 'address'));
      const nearAddr = list => (list || []).filter(o => { const a = norm(o.address); return a && addr && (a.includes(addr) || addr.includes(a)); });
      let hit = null, how = '';
      const tiers = [
        [phone && name ? byPhoneName.get(phone + '|' + name) : null, '电话+收货人'],
        [phone ? byPhone.get(phone) : null, '电话'],
        [name && addr ? nearAddr(byName.get(name)) : null, '收货人+地址'],
        [name ? byName.get(name) : null, '收货人']
      ];
      for (const [list, label] of tiers) { const u = uniq(list); if (u) { hit = u; how = label; break; } }
      return {
        row, index,
        recipient: cell(row, 'recipient'), phone: cell(row, 'phone'), address: cell(row, 'address'),
        product: cell(row, 'product'), tracking: cell(row, 'tracking'), carrier: cell(row, 'carrier'),
        orderNo: hit ? hit.orderNo : '', how: hit ? how : '', orderId: hit ? hit.id : ''
      };
    });
  }
  async function openShipment(file) {
    if (!file) return;
    await guardAsync(async () => {
      const book = await OrderExcel.read(file);
      let si = book.sheets.findIndex(s => s.rows.length > 2);
      if (si < 0) si = 0;
      const d = { sheet: si, header: OrderExcel.detectHeader(book.sheets[si].rows), map: {}, target: 'auto' };
      d.map = shipmentAutoMap(book.sheets[si].rows[d.header]);
      const draw = () => {
        const sheet = book.sheets[d.sheet], headers = sheet.rows[d.header] || [];
        const configured = { ...d.map };
        const targetIndex = configured.orderNo >= 0 ? configured.orderNo : -1;
        openModal('运单回填 · 选择列', `
          <div class="alert info">${icon('file')}<span>已读取 <strong>${esc(file.name)}</strong> · ${book.sheets.length} 个工作表 · ${Math.max(0, sheet.rows.length - d.header - 1)} 行。</span></div>
          <div class="form-grid">
            ${selectField('工作表', 's-sheet', book.sheets.map((s, i) => [i, `${s.name}（${s.rows.length} 行）`]), d.sheet)}
            ${field('表头所在行', 's-header', d.header + 1, { type: 'number', hint: '第 1 行起算' })}
            ${selectField('订单号写到哪一列', 's-target', [...headers.map((h, i) => [String(i), `${String(h || `列${i + 1}`).replace(/\s+/g, ' ').slice(0, 20)}（覆盖该列）`]), ['new', '追加新列「订单号」']], targetIndex >= 0 ? String(targetIndex) : 'new', true)}
          </div>
          <h3 class="section-label">匹配用的列</h3>
          <div class="map-grid">${SHIPMENT_FIELDS.filter(([k]) => k !== 'orderNo').map(([k, label]) => `<label class="map-item"><span>${esc(label)}</span><select name="s-map-${k}"><option value="-1">没有此列</option>${headers.map((h, i) => `<option value="${i}" ${d.map[k] === i ? 'selected' : ''}>${esc(String(h == null || h === '' ? `列${i + 1}` : h).replace(/\s+/g, ' ').slice(0, 24))}</option>`).join('')}</select></label>`).join('')}</div>
          <p class="muted">按「电话+收货人」优先匹配本次批次里已填订单号的明细，其次电话、收货人+地址、收货人；一条对多单时不回填，留给人工核对。</p>
          <p class="form-error" id="shipment-error"></p>
        `, `<button class="btn secondary" data-cancel>取消</button><button class="btn" id="run-shipment">开始匹配</button>`, true);
        $('[name="s-sheet"]').onchange = e => { d.sheet = Number(e.target.value); d.header = OrderExcel.detectHeader(book.sheets[d.sheet].rows); d.map = shipmentAutoMap(book.sheets[d.sheet].rows[d.header]); draw(); };
        $('[name="s-header"]').onchange = e => { const n = Number(e.target.value); if (Number.isInteger(n) && n >= 1 && n <= Math.min(sheet.rows.length, 20)) { d.header = n - 1; d.map = shipmentAutoMap(book.sheets[d.sheet].rows[d.header]); draw(); } };
        for (const [k] of SHIPMENT_FIELDS) { if (k === 'orderNo') continue; const el = $(`[name="s-map-${k}"]`); if (el) el.onchange = e => { d.map[k] = Number(e.target.value); }; }
        $('#run-shipment').onclick = () => guardAsync(async () => {
          const targetVal = $('[name="s-target"]').value;
          const rows = sheet.rows.slice(d.header + 1).filter(r => r && r.some(v => String(v ?? '').trim()));
          if (!rows.length) { $('#shipment-error').textContent = '表头下没有数据行。'; return; }
          const matched = shipmentMatch(rows, d.map);
          shipment = {
            name: file.name, sheetName: sheet.name, header: d.header,
            headers: [...headers],
            rows, matched,
            map: { ...d.map },
            target: targetVal === 'new' ? 'new' : Number(targetVal),
            widths: sheet.widths || [],
            createdAt: new Date().toISOString()
          };
          let filled = 0;
          for (const m of matched) {
            if (!m.orderId) continue;
            const order = state.orders.find(o => o.id === m.orderId);
            if (!order) continue;
            if (m.tracking && order.trackingNo !== m.tracking) { order.trackingNo = m.tracking; filled++; }
            if (m.carrier) order.carrier = m.carrier;
          }
          if (filled) await save();
          closeModal(); render();
          toast(`已匹配 ${matched.filter(m => m.orderNo).length} / ${matched.length} 行${filled ? `，${filled} 条快递单号已写回订单` : ''}。`);
        });
      };
      draw();
    });
  }
  function shipmentStats() {
    if (!shipment) return { total: 0, matched: 0, missing: 0 };
    const matched = shipment.matched.filter(m => m.orderNo).length;
    return { total: shipment.matched.length, matched, missing: shipment.matched.length - matched };
  }
  function shipmentPreviewRows() {
    return shipment.matched.slice(0, 20);
  }
  function renderShipment() {
    const st = shipmentStats();
    const hasOrders = shipmentOrders().length > 0;
    return `<div class="page-intro"><div><div class="eyebrow">TRACKING BACKFILL</div><h1>快递回来了，订单号补上</h1><p>上传带快递单号的表，按收货人和电话匹配上一步的订单，把订单号写回同一张表再下载。</p></div><div class="intro-actions">${shipment ? `<button class="btn secondary" id="shipment-reset">重新选择文件</button><button class="btn" id="shipment-download">${icon('download')} 下载回填表</button>` : ''}</div></div>
      ${!hasOrders ? `<div class="alert warning">${icon('shield')}<span>本次批次里还没有带订单号的明细，先去「订单采集与匹配」完成匹配，再回来回填。</span></div>` : ''}
      <section class="card"><div class="card-header"><div><h2>上传快递单号表</h2><p class="muted">支持 .xlsx / UTF-8 CSV；原表列结构保持不动，只在订单号列写入值。</p></div></div><div class="card-body"><div class="dropzone" id="shipment-drop" tabindex="0" role="button" aria-label="选择或拖入快递单号表"><div class="upload-icon">${icon('upload')}</div><h3>拖入快递单号表，或 <span class="text-link">选择文件</span></h3><p>支持 .xlsx / UTF-8 CSV，单文件不超过 15 MB</p><span class="drop-hint">自动识别收货人、电话、地址、快递单号</span></div><input id="shipment-file" type="file" accept=".xlsx,.csv" hidden></div></section>
      ${shipment ? `<section class="card"><div class="card-header"><div><h2>${esc(shipment.name)}</h2><p class="muted">${esc(shipment.sheetName)} · 共 ${st.total} 行 · 匹配 ${st.matched} · 未匹配 ${st.missing}</p></div><div class="intro-actions"><span class="pill ${st.missing ? 'warning' : 'success'}">${st.missing ? `${st.missing} 行未匹配` : '全部匹配'}</span><button class="btn ghost small" id="shipment-remap">调整列对应</button></div></div>
        ${st.missing ? `<div class="alert warning">${icon('shield')}<span>${st.missing} 行没有匹配到订单号，下载后这些行的订单号列会是空的；可在表格里按收货人/电话人工补。</span></div>` : ''}
        <div class="table-wrap"><table class="data-table"><thead><tr><th>收货人</th><th>电话</th><th>快递单号</th><th>商品</th><th>订单号</th><th>匹配依据</th></tr></thead><tbody>${shipmentPreviewRows().map(m => `<tr><td>${esc(m.recipient || '—')}</td><td class="mono">${esc(m.phone || '—')}</td><td class="mono">${esc(m.tracking || '—')}</td><td class="product-cell">${esc(String(m.product || '').slice(0, 30) || '—')}</td><td class="mono">${m.orderNo ? esc(m.orderNo) : '<span class="muted">未匹配</span>'}</td><td>${m.how ? esc(m.how) : '<span class="muted">—</span>'}</td></tr>`).join('')}</tbody></table></div>
        <div class="pagination"><span>预览前 ${shipmentPreviewRows().length} 行，下载包含全部 ${st.total} 行。</span></div></section>` : ''}`;
  }
  async function exportShipment() {
    if (!shipment) return;
    const headers = [...shipment.headers];
    const append = shipment.target === 'new';
    if (append) headers.push('订单号');
    const rows = shipment.matched.map(m => {
      const cells = shipment.headers.map((h, i) => m.row[i] ?? '');
      if (append) cells.push(m.orderNo || '');
      else cells[shipment.target] = m.orderNo || '';
      return cells;
    });
    const bytes = await OrderExcel.buildWorkbook({
      headers: headers.map((h, i) => String(h || `列${i + 1}`)),
      rows,
      sheetName: shipment.sheetName || 'Sheet1',
      widths: shipment.widths
    });
    OrderExcel.downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${shipment.name.replace(/\.(xlsx|csv)$/i, '')}_回填订单号_${stamp()}.xlsx`);
    toast(`已导出 ${rows.length} 行，其中 ${shipment.matched.filter(m => m.orderNo).length} 行带订单号。`);
  }
  function bindShipment() {
    const drop = $('#shipment-drop');
    $('#shipment-drop')?.addEventListener('click', () => $('#shipment-file').click());
    $('#shipment-drop')?.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('#shipment-file').click(); } });
    if (drop) {
      drop.ondragover = e => { e.preventDefault(); drop.classList.add('dragover'); };
      drop.ondragleave = () => drop.classList.remove('dragover');
      drop.ondrop = e => { e.preventDefault(); drop.classList.remove('dragover'); const f = e.dataTransfer.files?.[0]; if (f) openShipment(f); };
    }
    $('#shipment-file')?.addEventListener('change', e => { const f = e.target.files?.[0]; e.target.value = ''; openShipment(f); });
    $('#shipment-download')?.addEventListener('click', () => guardAsync(exportShipment));
    $('#shipment-reset')?.addEventListener('click', () => { shipment = null; render(); });
    $('#shipment-remap')?.addEventListener('click', () => { toast('重新上传同一份表即可重选列。', 'warning'); $('#shipment-file').click(); });
  }
  function templateModal(sid,override=null){const supplier=sById(sid);if(!supplier)return;const t=copy(override||supplier.template||{name:'自定义格式',sheetName:'Sheet1',columns:copy(OrderSeed.supplier.template.columns)});openModal(`${supplier.name} · 格式设置`,`<div class="template-top"><div><strong>${esc(t.name)}</strong><p class="muted">逐列指定导出内容。只读取上传模板的表头，不带入示例订单。</p></div><button class="btn secondary small" id="upload-template">上传格式样表</button><input id="template-file" type="file" accept=".xlsx,.csv" hidden></div><form id="template-form"><div class="form-grid">${field('格式名称','templateName',t.name,{required:true})}${field('导出工作表名称','sheetName',t.sheetName||'Sheet1',{required:true})}</div><div class="table-wrap"><table class="data-table template-table"><thead><tr><th>输出列名</th><th>取值来源</th><th>缺失时默认值 / 固定值</th><th>必填</th></tr></thead><tbody>${t.columns.map((c,i)=>`<tr><td><input name="header-${i}" value="${esc(c.header)}" aria-label="第${i+1}列名称"></td><td><select name="field-${i}" aria-label="${esc(c.header)}取值来源">${EXPORT_FIELDS.map(([k,l])=>`<option value="${k}" ${k===c.field?'selected':''}>${l}</option>`).join('')}</select></td><td><input name="fallback-${i}" value="${esc(c.fallback||'')}" aria-label="${esc(c.header)}默认值"></td><td><input type="checkbox" name="required-${i}" ${c.required?'checked':''} aria-label="${esc(c.header)}为必填"></td></tr>`).join('')}</tbody></table></div><p class="muted">友本要求的列顺序为：订单号、买家昵称、收货人、详细地址、收货人电话、商品编码、数量。未提供买家昵称时，默认留空。</p><p class="muted">「订单号」允许为空：即使勾选必填，也不会阻止导出，只是该格留空。</p><p class="form-error" id="template-error"></p></form>`,`<button class="btn secondary" data-cancel>取消</button><button class="btn" id="save-template">保存格式</button>`,true);
    $('#upload-template').onclick=()=>$('#template-file').click();$('#template-file').onchange=e=>guardAsync(async()=>{const file=e.target.files[0];if(!file)return;const book=await OrderExcel.read(file);let si=0,hi=OrderExcel.detectHeader(book.sheets[0].rows);const showPicker=()=>{const sh=book.sheets[si],headers=sh.rows[hi]||[];openModal('选择供应商格式表头',`<div class="form-grid">${selectField('模板工作表','templateSheet',book.sheets.map((s,i)=>[i,s.name]),si)}${field('表头行','templateHeader',hi+1,{type:'number'})}</div><p class="muted">将使用下列 ${headers.length} 列生成新订单文件，表头上方的说明及原有数据行不会复制。</p><div class="template-header-chips">${headers.map(h=>`<span class="pill neutral">${esc(h||'空白列')}</span>`).join('')}</div>`,`<button class="btn secondary" data-cancel>取消</button><button class="btn" id="apply-template">使用这些列</button>`);$('[name=templateSheet]').onchange=e=>{si=Number(e.target.value);hi=OrderExcel.detectHeader(book.sheets[si].rows);showPicker();};$('[name=templateHeader]').onchange=e=>{const n=Number(e.target.value);if(Number.isInteger(n)&&n>0&&n<=Math.min(20,sh.rows.length)){hi=n-1;showPicker();}else toast('请输入有效表头行。','warning');};$('#apply-template').onclick=()=>{if(!headers.length||headers.some(h=>!String(h??'').trim())){toast('表头包含空白列，请选择完整表头或整理模板后再上传。','warning');return;}const cols=headers.map(h=>{const n=norm(h);let field='blank';if(['商品编码','sku码','sku','供应商商品编码'].map(norm).includes(n))field='supplierSku';else{const f=FIELDS.find(([,l,names])=>norm(l)===n||names.some(x=>norm(x)===n));field=f?.[0]||'blank';if(field==='sourceSku')field='supplierSku';}return {header:String(h),field,required:['orderNo','recipient','phone','address','supplierSku','quantity'].includes(field),fallback:''};});templateModal(sid,{name:file.name,sheetName:sh.name,columns:cols,widths:sh.widths||[]});};};showPicker();});
    $('#save-template').onclick=async()=>{const f=new FormData($('#template-form'));const columns=t.columns.map((c,i)=>({header:String(f.get(`header-${i}`)||'').trim(),field:f.get(`field-${i}`),fallback:String(f.get(`fallback-${i}`)||''),required:f.has(`required-${i}`)}));if(columns.some(c=>!c.header)){ $('#template-error').textContent='所有输出列名都不能为空。';return;}if(new Set(columns.map(c=>c.header)).size!==columns.length){$('#template-error').textContent='输出列名不能重复。';return;}if(!f.get('templateName')?.trim()||!f.get('sheetName')?.trim()){$('#template-error').textContent='请填写格式和工作表名称。';return;}if(columns.some(c=>c.required&&c.field==='blank'&&!c.fallback)){$('#template-error').textContent='必填列不能设为留空，请选择取值来源。';return;}supplier.template={name:f.get('templateName').trim(),sheetName:f.get('sheetName').trim(),columns,widths:t.widths};history('更新供应商模板');await save();closeModal();render();toast('格式已保存。');};
  }
  const safeName=s=>s.replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').slice(0,70)||'供应商';
  async function exportBytes(g){const errs=exportIssues(g);if(errs.length)throw Error(errs.join('；'));const t=g.supplier.template;return OrderExcel.buildWorkbook({headers:t.columns.map(c=>c.header),rows:g.rows.map(r=>t.columns.map(c=>templateValue(r,c))),sheetName:t.sheetName||'Sheet1',widths:t.widths});}
  async function exportOne(g){if(!g)return;const bytes=await exportBytes(g);OrderExcel.downloadBlob(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`${safeName(g.supplier.name)}_导单_${stamp()}.xlsx`);markExported(g.rows);history(`导出 ${g.supplier.name} ${g.rows.length} 行`);await save();toast('供应商订单已导出。');}
  async function exportAll(){const zip=new JSZip(),gs=exportGroups();if(!gs.length)return;for(let i=0;i<gs.length;i++){const g=gs[i],bytes=await exportBytes(g);zip.file(`${String(i+1).padStart(2,'0')}_${safeName(g.supplier.name)}_导单_${stamp()}.xlsx`,bytes);}const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE'});OrderExcel.downloadBlob(blob,`供应商订单_${stamp()}.zip`);markExported(gs.flatMap(g=>g.rows));history(`批量导出 ${gs.length} 家供应商`);await save();toast(`已打包 ${gs.length} 家供应商的订单。`);}
  async function deleteOrders(ids, message) {
    const set = new Set(ids);
    if (!set.size) return;
    state.orders = state.orders.filter(r => !set.has(r.id));
    state.orders.forEach(r => { if (set.has(r.duplicateOf)) delete r.duplicateOf; });
    const alive = new Set(state.orders.map(r => r.batchId));
    state.batches = state.batches.filter(b => alive.has(b.id));
    state.activeBatchIds = state.activeBatchIds.filter(id => alive.has(id));
    invalidate();
    await save();
    render();
    toast(message || '已删除。');
  }
  function manageBatches(){
    openModal('本次处理批次与历史',`<p class="muted">仅勾选的批次参与本次订单检查与供应商导出。导出记录会保留；下次导入新文件时，已完整导出的旧批次会自动移出本次范围。</p>${state.batches.map(b=>`<div class="batch-row"><label class="checkbox-row"><input type="checkbox" data-scope-batch="${esc(b.id)}" ${state.activeBatchIds.includes(b.id)?'checked':''}><span><strong>${esc(b.name)}</strong><small>${esc(b.customer)} · ${state.orders.filter(r=>r.batchId===b.id).length} 行 · ${b.exportedAt?'已完整导出':'尚未完整导出'}</small></span></label><button class="btn danger small" data-delete-batch="${esc(b.id)}">删除</button></div>`).join('')}`,`<button class="btn secondary" data-cancel>关闭</button>`,true);
    $$('[data-scope-batch]').forEach(el=>el.onchange=async()=>{const id=el.dataset.scopeBatch;state.activeBatchIds=state.activeBatchIds.filter(bid=>bid!==id);if(el.checked)state.activeBatchIds.push(id);page=1;invalidate();await save();render();});
    $$('[data-delete-batch]').forEach(b=>b.onclick=()=>{const batch=state.batches.find(x=>x.id===b.dataset.deleteBatch);confirmDialog('删除订单批次',`确认删除“${batch.name}”及其订单明细？`,async()=>{const ids=new Set(state.orders.filter(r=>r.batchId===batch.id).map(r=>r.id));state.orders=state.orders.filter(r=>r.batchId!==batch.id);state.orders.forEach(r=>{if(ids.has(r.duplicateOf))delete r.duplicateOf;});state.batches=state.batches.filter(x=>x.id!==batch.id);state.activeBatchIds=state.activeBatchIds.filter(id=>id!==batch.id);invalidate();await save();render();toast('批次已删除。');});});
  }
  async function loadDemo(){if(state.batches.some(b=>b.demo)){toast('演示数据已经载入，可在“管理导入批次”中删除。','warning');return;}const add=async()=>{const id=uid();state.batches.push({id,name:'演示订单（模拟数据）.xlsx',customer:'演示客户',hash:'demo-'+id,demo:true,createdAt:new Date().toISOString()});for(let i=0;i<6;i++){const c=OrderSeed.candidates[i<2?0:i<5?1:2];state.orders.push({id:uid(),batchId:id,customer:c.customer,orderNo:'DEMO-'+String(i+1).padStart(4,'0'),productName:c.name,sourceSku:c.sourceSku||'',sourceSpec:c.customer==='GBH'?'早餐代餐':'',sourceLineId:'demo-'+i,quantity:1,recipient:'模拟收件人'+(i+1),phone:'1380000000'+i,address:'示例省示例市示例区演示路'+(i+1)+'号（模拟地址）',buyerNickname:'',remark:'模拟数据，请勿实际发货',date:stamp(),unitPrice:'',sourceFile:'演示订单（模拟数据）.xlsx',sourceSheet:'模拟数据',sourceRow:i+2,excluded:false,manual:false,raw:{说明:'模拟数据，请勿实际发货'}});}includeNewBatch(id);invalidate();await save();render();toast('已载入6条模拟订单，请确认商品映射。');};if(activeRows().length)confirmDialog('加入演示订单','将新增6条模拟订单，可在批次管理中删除。本次选中的批次会一同参与后续生成。',add);else await add();}
  window.OrderFlow={getState:()=>copy(state),status,summary,findProducts,validateBackup,exportIssues,templateValue,UNIFIED,autoMap,activeRows,duplicateKey};
  const boot=()=>{ if (window.OrderAuth && window.OrderAuth.ready) window.OrderAuth.ready.then(() => init()); else init(); };
  document.readyState==='loading'?document.addEventListener('DOMContentLoaded',boot):boot();
})();
