/** Pure annual economy. All changes stay in memory until the engine saves. */
const ECONOMY_RULES = {
  rangeName: 'Игровые_данные',
  headers: { buildings: 'Здания', templates: 'Шаблоны зданий', goods: 'Товары',
    countries: 'Страны', markets: 'Рынки', diplomacy: 'Торговые отношения' },
  salaryStep: 0.05, priceSpeed: 0.25, minPriceFactor: 0.25, maxPriceFactor: 4,
  epsilon: 1e-8,
};

function processEconomy(data, ctx) {
  const columns = data[ECONOMY_RULES.rangeName];
  if (!columns || !Object.prototype.hasOwnProperty.call(columns, ECONOMY_RULES.headers.buildings)) return data;
  const state = GameEconomy.load(columns);
  if (!state.buildings.length) return data;
  GameEconomy.run(state);
  GameEconomy.save(state, columns, ctx.config.maxJsonCellCharacters);
  if(ctx.log)state.buildings.forEach(b=>ctx.log({type:'GAME',category:'ECONOMY',country:state.provinces[b.provinceId].ownerId,visibility:'COUNTRY',ttl:1,message:(state.templatesById[b.templateId].name||b.templateId)+' ['+b.id+']: мощность '+Math.round(b.lastTurn.utilization*100)+'%, доход '+b.lastTurn.revenue.toFixed(2)+', расходы '+b.lastTurn.expenses.toFixed(2)+(b.lastTurn.reasons.length?'. '+b.lastTurn.reasons.join('; '):'')}));
  return data;
}

function validateEconomyData(data, ctx) {
  const c = data[ECONOMY_RULES.rangeName];
  if (!c || !Object.prototype.hasOwnProperty.call(c, ECONOMY_RULES.headers.buildings)) return true;
  const s = GameEconomy.load(c);
  if (s.buildings.length) GameEconomy.save(s, c, ctx.config.maxJsonCellCharacters);
  return true;
}

const GameEconomy = {
  fail: function (message) { throw new GameTurnDataError('Экономика: ' + message); },
  number: function (n, label, integer) {
    if (typeof n !== 'number' || !Number.isFinite(n) || n < 0 || n > Number.MAX_SAFE_INTEGER ||
        (integer && !Number.isSafeInteger(n))) this.fail(label + ': нужно неотрицательное ' + (integer ? 'целое ' : '') + 'число.');
    return n;
  },
  id: function (id, label) {
    if (typeof id !== 'string' || !id.trim() || ['__proto__', 'constructor', 'prototype'].includes(id)) this.fail(label + ': недопустимый id.');
  },
  read: function (cells, name) {
    if (!Array.isArray(cells)) this.fail('нужен столбец «' + name + '».');
    const records = [];
    cells.forEach(cell => {
      if (cell === null || cell === '' || cell === undefined) return;
      if (!Array.isArray(cell)) this.fail('«' + name + '»: нужен массив объектов в каждой ячейке.');
      cell.forEach(item => {
        if (!item || typeof item !== 'object' || Array.isArray(item)) this.fail('«' + name + '»: неверная запись.');
        records.push(JSON.parse(JSON.stringify(item)));
      });
    });
    return records;
  },
  index: function (list, name) {
    const result = Object.create(null);
    list.forEach(item => { this.id(item.id, name); if (result[item.id]) this.fail('повтор id ' + item.id); result[item.id] = item; });
    return result;
  },
  pack: function (items, capacity, limit, name) {
    try { return GamePopulation.pack(items, capacity, limit); }
    catch (e) { this.fail('столбец «' + name + '»: ' + e.message); }
  },
  load: function (columns) {
    const s = { lists: {}, provinces: Object.create(null), resources: Object.create(null) };
    s.buildings = this.read(columns[ECONOMY_RULES.headers.buildings], ECONOMY_RULES.headers.buildings);
    if (!s.buildings.length) return s;
    Object.keys(ECONOMY_RULES.headers).forEach(key => {
      s[key] = this.read(columns[ECONOMY_RULES.headers[key]], ECONOMY_RULES.headers[key]);
      s.lists[key] = s[key];
    });
    ['buildings','templates','goods','countries','markets'].forEach(key => { s[key + 'ById'] = this.index(s[key], key); });
    const provinceCells = columns[GAME_RULES.population.provinceHeader];
    const provinceIds = GamePopulation.provinces(provinceCells);
    s.provinceCells = JSON.parse(JSON.stringify(provinceCells));
    s.provinceCells.forEach(p => { if (p) s.provinces[p.id] = p; });
    s.population = GamePopulation.read(columns[GAME_RULES.population.header], provinceIds).map(g => JSON.parse(JSON.stringify(g)));
    s.populationById = this.index(s.population, 'население');
    s.goods.forEach(g => { this.number(g.basePrice, g.id + '.basePrice'); this.number(g.transportWeight, g.id + '.transportWeight'); if (!g.basePrice || !g.transportWeight) this.fail('цена и вес должны быть больше нуля: ' + g.id); });
    s.markets.forEach(m => { m.prices = m.prices || {}; Object.keys(m.prices).forEach(id => { if (!s.goodsById[id]) this.fail('неизвестный товар ' + id); this.number(m.prices[id], 'цена'); if (!m.prices[id]) this.fail('нулевая цена'); }); });
    s.countries.forEach(c => {
      if (!s.marketsById[c.marketId]) this.fail('страна ' + c.id + ': неизвестный рынок');
      this.number(c.treasury, c.id + '.treasury');
      ['men','women'].forEach(sex => { this.number(c.workerCoefficients && c.workerCoefficients[sex], 'коэффициент ' + sex); if (c.workerCoefficients[sex] > 1) this.fail('коэффициент рабочих больше 1'); });
      c.subsidies = c.subsidies || {budget:0, templateIds:[]};
      this.number(c.subsidies.budget, 'бюджет субсидий');
      if (!Array.isArray(c.subsidies.templateIds) || c.subsidies.templateIds.some(id => !s.templatesById[id])) this.fail('неизвестный тип субсидируемого здания');
    });
    Object.values(s.provinces).forEach(p => {
      const active = s.buildings.some(b => b.provinceId === p.id);
      if (active && !s.countriesById[p.ownerId]) this.fail('провинция ' + p.id + ': неизвестный владелец');
      p.transport = p.transport || {baseCapacity:0, tariff:0, buildingCapacity:0};
      ['baseCapacity','tariff','buildingCapacity'].forEach(k => { if (p.transport[k] === undefined && k === 'buildingCapacity') p.transport[k] = 0; this.number(p.transport[k], 'транспорт ' + p.id + '.' + k); });
      if (p.neighbors !== undefined && (!Array.isArray(p.neighbors) || p.neighbors.some(id => !Number.isSafeInteger(id)))) this.fail('неверные соседи провинции ' + p.id);
      const map = Object.create(null);
      (p.resources || []).forEach(r => {
        // Preserve old province resources: type serves as ID; absent mode means finite.
        const id = r.id || r.type; this.id(id, 'ресурс'); if (map[id]) this.fail('повтор ресурса ' + id);
        const mode = r.mode || 'finite'; if (!['finite','renewable','annual'].includes(mode)) this.fail('неизвестный тип ресурса');
        if (mode === 'annual') this.number(r.annualCapacity, 'годовой лимит');
        else this.number(r.amount, 'запас ресурса');
        if (mode === 'renewable') { this.number(r.maxAmount, 'предел восстановления'); this.number(r.regeneration, 'восстановление'); if (r.amount > r.maxAmount) this.fail('ресурс выше максимального запаса'); }
        map[id] = r;
      });
      s.resources[p.id] = map;
    });
    s.templates.forEach(t => {
      ['workers','inputs','outputs','extraction'].forEach(k => { if (t[k] === undefined) t[k] = []; if (!Array.isArray(t[k])) this.fail(t.id + '.' + k + ': нужен массив'); });
      this.number(t.storageCapacity, 'вместимость ' + t.id);
      const professions = new Set();
      t.workers.forEach(w => { this.id(w.profession, 'профессия'); if (professions.has(w.profession)) this.fail('повтор профессии'); professions.add(w.profession); this.number(w.count, 'рабочие', true); this.number(w.baseSalary, 'зарплата'); });
      ['inputs','outputs'].forEach(k => { const seen = new Set(); t[k].forEach(x => { if (!s.goodsById[x.goodId] || seen.has(x.goodId)) this.fail('неизвестный или повторный товар в ' + t.id); seen.add(x.goodId); this.number(x.amount, 'норма товара'); }); });
      const resourceIds = new Set(); t.extraction.forEach(x => { this.id(x.resourceId,'источник'); if (resourceIds.has(x.resourceId)) this.fail('повтор источника'); resourceIds.add(x.resourceId); this.number(x.amount,'норма добычи'); });
      if (t.transportCapacity !== undefined) this.number(t.transportCapacity, 'транспортная мощность здания');
      this.validateCondition(t.requirements);
      (t.efficiencyModifiers || []).forEach(x => { this.validateCondition(x.when); this.number(x.multiplier, 'модификатор эффективности'); });
    });
    const provinceTemplates = new Set();
    s.buildings.forEach(b => {
      if (!s.templatesById[b.templateId] || !s.provinces[b.provinceId]) this.fail('здание ' + b.id + ': неизвестный шаблон или провинция');
      const key = b.provinceId + ':' + b.templateId; if (provinceTemplates.has(key)) this.fail('повтор типа здания в провинции'); provinceTemplates.add(key);
      this.number(b.level, 'уровень', true); if (!b.level) this.fail('нулевой уровень');
      if (b.enabled !== undefined && typeof b.enabled !== 'boolean') this.fail('enabled должен быть true/false');
      this.number(b.priority === undefined ? 0 : b.priority, 'приоритет'); this.number(b.money, 'деньги здания');
      b.inventory = b.inventory || {}; b.salaries = b.salaries || {}; b.employment = b.employment || [];
      Object.keys(b.inventory).forEach(id => { if (!s.goodsById[id]) this.fail('неизвестный товар на складе'); this.number(b.inventory[id], 'запас товара'); });
      Object.keys(b.salaries).forEach(id => { const w = s.templatesById[b.templateId].workers.find(w => w.profession === id); if (!w) this.fail('неизвестная профессия зарплаты'); this.number(b.salaries[id], 'зарплата'); if (b.salaries[id] < w.baseSalary) this.fail('зарплата ниже базовой'); });
      if (!Array.isArray(b.employment)) this.fail('employment должен быть массивом');
      b.employment.forEach(e => { const g = s.populationById[e.populationId]; if (!g || g.provinceId !== b.provinceId || !s.templatesById[b.templateId].workers.some(w=>w.profession===e.profession)) this.fail('неверная занятость ' + b.id); this.number(e.count, 'занятость', true); });
      if (this.weight(s,b) > s.templatesById[b.templateId].storageCapacity*b.level + ECONOMY_RULES.epsilon) this.fail('склад здания ' + b.id + ' переполнен');
    });
    s.population.forEach(g=>{ if(g.money!==undefined)this.number(g.money,'деньги населения'); });
    s.diplomacy.forEach(d => {
      if (!s.countriesById[d.from] || !s.countriesById[d.to]) this.fail('неизвестная страна торговых отношений');
      ['tradeAllowed','embargo','transitAllowed'].forEach(k=>{if(typeof d[k]!=='boolean')this.fail('торговые отношения: нужен ' + k);});
      ['importTariff','exportTariff'].forEach(k=>this.number(d[k],k));
    });
    const pairs = new Set(); s.diplomacy.forEach(d=>{const k=d.from+':'+d.to;if(pairs.has(k))this.fail('повтор торговых отношений');pairs.add(k);});
    return s;
  },
  validateCondition: function (c) {
    if (c === undefined || c === null) return;
    if (!c || typeof c !== 'object' || Array.isArray(c)) this.fail('неверное условие');
    const keys = ['all','any','not','field','resourceId'].filter(k=>Object.prototype.hasOwnProperty.call(c,k));
    if (keys.length !== 1) this.fail('условие должно иметь один вид: all/any/not/field/resourceId');
    if (c.all || c.any) { const a=c.all||c.any;if(!Array.isArray(a)||!a.length)this.fail('пустая группа условий');a.forEach(x=>this.validateCondition(x)); }
    else if(c.not) this.validateCondition(c.not);
    else if(c.resourceId) {this.id(c.resourceId,'условие ресурса');if(c.operator!=='available')this.fail('для ресурса нужен available');}
    else { if(typeof c.field!=='string'||!c.field||c.field.split('.').some(k=>['__proto__','constructor','prototype'].includes(k)))this.fail('неверное поле условия');if(!['eq','ne','in','gte','lte','gt','lt','between'].includes(c.operator))this.fail('неизвестный оператор условия');if(c.operator==='in'&&!Array.isArray(c.value))this.fail('in требует массив');if(c.operator==='between'&&(!Array.isArray(c.value)||c.value.length!==2))this.fail('between требует две границы'); }
  },
  condition: function(c,p,resources) {
    if(!c)return true;
    if(c.all)return c.all.every(x=>this.condition(x,p,resources));
    if(c.any)return c.any.some(x=>this.condition(x,p,resources));
    if(c.not)return !this.condition(c.not,p,resources);
    if(c.resourceId){const r=resources[c.resourceId];return !!r&&((r.mode==='annual'?r.annualCapacity:r.amount)>0);}
    const v=c.field.split('.').reduce((o,k)=>o==null?undefined:o[k],p);
    switch(c.operator){case 'eq':return v===c.value;case 'ne':return v!==c.value;case 'in':return c.value.includes(v);case 'gte':return v>=c.value;case 'lte':return v<=c.value;case 'gt':return v>c.value;case 'lt':return v<c.value;case 'between':return v>=c.value[0]&&v<=c.value[1];}
    return false;
  },
  weight:function(s,b){return Object.keys(b.inventory).reduce((n,id)=>n+b.inventory[id]*s.goodsById[id].transportWeight,0);},
  price:function(s,country,id){const m=s.marketsById[s.countriesById[country].marketId];return m.prices[id]||s.goodsById[id].basePrice;},
  relation:function(s,from,to){return s.diplomacy.find(d=>d.from===from&&d.to===to);},
  allowed:function(s,from,to){if(from===to)return true;const a=this.relation(s,from,to),b=this.relation(s,to,from);if((a&&a.embargo)||(b&&b.embargo))return false;return s.countriesById[from].marketId===s.countriesById[to].marketId||!!(a&&b&&a.tradeAllowed&&b.tradeAllowed);},
  route:function(s,from,to,buyer,weight){
    const dist=Object.create(null),paths=Object.create(null),visited=new Set();
    const permitted=p=>{if(!s.countriesById[p.ownerId])return false;if(p.ownerId!==buyer&&p.id!==from&&p.id!==to){const d=this.relation(s,p.ownerId,buyer);if(!d||!d.transitAllowed||d.embargo)return false;}return s.transport[p.id]>ECONOMY_RULES.epsilon&&s.transport[p.id]>=weight-ECONOMY_RULES.epsilon;};
    if(!permitted(s.provinces[from])||!permitted(s.provinces[to]))return null;
    dist[from]=s.provinces[from].transport.tariff;paths[from]=[from];
    while(true){const next=Object.keys(dist).filter(id=>!visited.has(+id)).sort((a,b)=>dist[a]-dist[b]||+a-+b)[0];if(next===undefined)return null;const id=+next;if(id===to)return {path:paths[id],tariff:dist[id]};visited.add(id);(s.provinces[id].neighbors||[]).forEach(n=>{const p=s.provinces[n];if(!p||visited.has(n)||!permitted(p))return;const cost=dist[id]+p.transport.tariff;if(dist[n]===undefined||cost<dist[n]){dist[n]=cost;paths[n]=paths[id].concat(n);}});}
  },
  allocateLabor:function(plans,s,capField){
    s.population.forEach(g=>{g.workforce.free=g.workforce.available;});
    plans.forEach(x=>{x.hires=[];});
    const jobs=[];plans.filter(x=>x.active).forEach(x=>x.t.workers.forEach(w=>jobs.push({x,w,wanted:Math.floor(w.count*x.b.level*(capField?x[capField]:1)+ECONOMY_RULES.epsilon),salary:x.b.salaries[w.profession]})));
    jobs.sort((a,b)=>(b.x.b.priority||0)-(a.x.b.priority||0)||b.salary-a.salary||a.x.b.id.localeCompare(b.x.b.id));
    while(jobs.length){const first=jobs.shift(),batch=[first];for(let i=jobs.length-1;i>=0;i--){const j=jobs[i];if(j.x.p.id===first.x.p.id&&(j.x.b.priority||0)===(first.x.b.priority||0)&&j.salary===first.salary){batch.push(j);jobs.splice(i,1);}}batch.sort((a,b)=>a.x.b.id.localeCompare(b.x.b.id)||a.w.profession.localeCompare(b.w.profession));const groups=s.population.filter(g=>g.provinceId===first.x.p.id).sort((a,b)=>a.id.localeCompare(b.id));const free=groups.reduce((n,g)=>n+g.workforce.free,0),wanted=batch.reduce((n,j)=>n+j.wanted,0);let allocations=batch.map(j=>Math.floor(j.wanted*Math.min(1,free/(wanted||1))));let remainder=Math.min(free,wanted)-allocations.reduce((a,b)=>a+b,0);for(let i=0;remainder>0;i=(i+1)%batch.length){if(allocations[i]<batch[i].wanted){allocations[i]++;remainder--;}}batch.forEach((j,i)=>{let count=allocations[i];groups.forEach(g=>{const n=Math.min(count,g.workforce.free);if(n){j.x.hires.push({populationId:g.id,profession:j.w.profession,count:n});g.workforce.free-=n;count-=n;}});});}
  },
  run:function(s){
    const eps=ECONOMY_RULES.epsilon;
    s.transport=Object.create(null);s.remainingResources=Object.create(null);s.demand=Object.create(null);s.supply=Object.create(null);
    Object.values(s.provinces).forEach(p=>{s.transport[p.id]=p.transport.baseCapacity+p.transport.buildingCapacity;s.remainingResources[p.id]=Object.create(null);Object.keys(s.resources[p.id]).forEach(id=>{const r=s.resources[p.id][id];s.remainingResources[p.id][id]=r.mode==='annual'?r.annualCapacity:r.amount;});});
    s.population.forEach(g=>{const c=s.countriesById[s.provinces[g.provinceId].ownerId];if(!c)this.fail('нет коэффициентов страны для населения '+g.id);g.money=g.money||0;g.workforce={available:Math.floor(g.men*c.workerCoefficients.men)+Math.floor(g.women*c.workerCoefficients.women),free:0,byProfession:{}};g.workforce.free=g.workforce.available;});
    const plans=s.buildings.map(b=>{
      const t=s.templatesById[b.templateId],p=s.provinces[b.provinceId];
      const plan={b,t,p,c:s.countriesById[p.ownerId],efficiency:1,active:b.enabled!==false&&this.condition(t.requirements,p,s.resources[p.id]),hires:[],reasons:[],start:Object.assign({},b.inventory),salaryCost:0};
      (t.efficiencyModifiers||[]).forEach(x=>{if(this.condition(x.when,p,s.resources[p.id]))plan.efficiency*=x.multiplier;});
      this.number(plan.efficiency,'эффективность');if(!plan.efficiency)plan.active=false;
      b.lastTurn={utilization:0,efficiency:plan.efficiency,revenue:0,expenses:0,subsidy:0,reasons:plan.reasons};b.employment=[];
      t.workers.forEach(w=>{if(b.salaries[w.profession]===undefined)b.salaries[w.profession]=w.baseSalary;plan.salaryCost+=w.count*b.level*b.salaries[w.profession];});
      if(!plan.active)plan.reasons.push('Отключено или не выполнены условия работы');
      return plan;
    });
    // Snapshot sale offers: no newly produced or purchased goods can be resold this turn.
    const offers=[];
    plans.forEach(x=>{Object.keys(x.start).forEach(id=>{const input=x.t.inputs.find(i=>i.goodId===id);const amount=Math.max(0,x.start[id]-(input?input.amount*x.b.level:0));if(amount>eps){offers.push({x,id,remaining:amount});const k=x.c.marketId+':'+id;s.supply[k]=(s.supply[k]||0)+amount;}});x.t.inputs.forEach(i=>{if(x.active){const k=x.c.marketId+':'+i.goodId;s.demand[k]=(s.demand[k]||0)+Math.max(0,i.amount*x.b.level-(x.b.inventory[i.goodId]||0));}});});
    // Subsidies are grants capped by both annual policy budget and treasury.
    s.countries.forEach(c=>{let budget=Math.min(c.treasury,c.subsidies.budget);const eligible=plans.filter(x=>x.active&&x.c===c&&c.subsidies.templateIds.includes(x.t.id)).sort((a,b)=>(b.b.priority||0)-(a.b.priority||0)||a.b.id.localeCompare(b.b.id));const tiers=[...new Set(eligible.map(x=>x.b.priority||0))];tiers.forEach(priority=>{const tier=eligible.filter(x=>(x.b.priority||0)===priority);const needs=tier.map(x=>Math.max(0,x.salaryCost+x.t.inputs.reduce((n,i)=>n+Math.max(0,i.amount*x.b.level-(x.b.inventory[i.goodId]||0))*this.price(s,c.id,i.goodId),0)-x.b.money));const total=needs.reduce((a,b)=>a+b,0),fraction=total?Math.min(1,budget/total):0;tier.forEach((x,i)=>{const grant=needs[i]*fraction;x.b.money+=grant;x.b.lastTurn.subsidy+=grant;c.treasury-=grant;budget-=grant;});});});
    // Labor allocated by priority, then offered salary; equal bids share proportionally.
    this.allocateLabor(plans,s);
    plans.forEach(x=>{x.labor=1;x.t.workers.forEach(w=>{if(w.count){const n=x.hires.filter(e=>e.profession===w.profession).reduce((n,e)=>n+e.count,0);x.labor=Math.min(x.labor,n/(w.count*x.b.level));}});x.labor=x.active?x.labor:0;x.workerShortage=x.active&&x.labor<1-eps;if(x.active&&x.labor<1-eps)x.reasons.push('Нехватка работников');});
    // Buyers with equal priority pay the same market quote. Fair-share rounds prevent one buyer taking all stock.
    plans.forEach(x=>{const inputsCost=x.t.inputs.reduce((n,i)=>n+Math.max(0,i.amount*x.b.level-(x.b.inventory[i.goodId]||0))*this.price(s,x.c.id,i.goodId),0);if(x.salaryCost+inputsCost>0)x.labor=Math.min(x.labor,x.b.money/(x.salaryCost+inputsCost));});
    const buyers=plans.filter(x=>x.active).sort((a,b)=>(b.b.priority||0)-(a.b.priority||0)||a.b.id.localeCompare(b.b.id));
    const tiers=[...new Set(buyers.map(x=>x.b.priority||0))];
    tiers.forEach(priority=>{
      const tier=buyers.filter(x=>(x.b.priority||0)===priority);
      const quotas=new Map();
      // Bounded rounds redistribute offers left unused by insolvent buyers.
      for(let round=0;round<32;round++){
      offers.forEach(o=>{const rivals=tier.filter(y=>y!==o.x&&this.allowed(s,o.x.c.id,y.c.id)&&this.route(s,o.x.p.id,y.p.id,y.c.id,eps)&&y.t.inputs.some(i=>i.goodId===o.id));const requests=rivals.map(y=>Math.max(0,Math.min(y.t.inputs.find(i=>i.goodId===o.id).amount*y.b.level-(y.b.inventory[o.id]||0),(y.b.money-y.salaryCost*y.labor)/this.price(s,o.x.c.id,o.id),(y.t.storageCapacity*y.b.level-this.weight(s,y.b))/s.goodsById[o.id].transportWeight)));const total=requests.reduce((a,b)=>a+b,0);rivals.forEach((y,i)=>quotas.set(y.b.id+':'+o.x.b.id+':'+o.id,o.remaining*(total?requests[i]/total:0)));});
        let changed=false;
        tier.forEach(x=>x.t.inputs.forEach(input=>{
          let need=Math.max(0,input.amount*x.b.level-(x.b.inventory[input.goodId]||0));if(need<=eps)return;
          let candidates=offers.filter(o=>o.id===input.goodId&&o.x!==x&&o.remaining>eps&&this.allowed(s,o.x.c.id,x.c.id)).map(o=>{
            const route=this.route(s,o.x.p.id,x.p.id,x.c.id,eps);if(!route)return null;
            const price=this.price(s,o.x.c.id,input.goodId),same=o.x.c.marketId===x.c.marketId;
            const imp=same?0:this.relation(s,x.c.id,o.x.c.id).importTariff,exp=same?0:this.relation(s,o.x.c.id,x.c.id).exportTariff;
            const cost=price*(1+imp+exp)+route.tariff*s.goodsById[input.goodId].transportWeight;
            return {o,route,price,imp,exp,cost};
          }).filter(Boolean).sort((a,b)=>a.cost-b.cost||a.o.x.b.id.localeCompare(b.o.x.b.id));
          candidates.forEach(q=>{
            const weight=s.goodsById[input.goodId].transportWeight;
            const reservedSalary=x.salaryCost*x.labor;
            const amount=Math.max(0,Math.min(need,Math.min(q.o.remaining,quotas.get(x.b.id+':'+q.o.x.b.id+':'+input.goodId)||0),(x.b.money-reservedSalary)/q.cost,(x.t.storageCapacity*x.b.level-this.weight(s,x.b))/weight,...q.route.path.map(id=>s.transport[id]/weight)));
            if(amount<=eps)return;
            const quotaKey=x.b.id+':'+q.o.x.b.id+':'+input.goodId;quotas.set(quotaKey,quotas.get(quotaKey)-amount);q.o.remaining-=amount;q.o.x.b.inventory[input.goodId]-=amount;x.b.inventory[input.goodId]=(x.b.inventory[input.goodId]||0)+amount;
            const paid=amount*q.cost;x.b.money-=paid;x.b.lastTurn.expenses+=paid;q.o.x.b.money+=amount*q.price;q.o.x.b.lastTurn.revenue+=amount*q.price;
            x.c.treasury+=amount*q.price*q.imp;q.o.x.c.treasury+=amount*q.price*q.exp;
            q.route.path.forEach(id=>{s.transport[id]-=amount*weight;s.countriesById[s.provinces[id].ownerId].treasury+=amount*weight*s.provinces[id].transport.tariff;});
            need-=amount;changed=true;
          });
        }));if(!changed)break;
      }
    });
    // Produce after trade. Shared resources are proportionally reserved within priority tiers.
    plans.forEach(x=>{
      let u=x.active?1:0;
      x.t.inputs.forEach(i=>{if(i.amount)u=Math.min(u,(x.b.inventory[i.goodId]||0)/(i.amount*x.b.level));});
      if(x.salaryCost)u=Math.min(u,x.b.money/x.salaryCost);
      const netWeight=x.t.outputs.reduce((n,o)=>n+o.amount*s.goodsById[o.goodId].transportWeight,0)*x.efficiency-x.t.inputs.reduce((n,i)=>n+i.amount*s.goodsById[i.goodId].transportWeight,0);
      if(netWeight>0)u=Math.min(u,Math.max(0,x.t.storageCapacity*x.b.level-this.weight(s,x.b))/(netWeight*x.b.level));
      x.target=Math.max(0,Math.min(1,u));
    });
    plans.forEach(x=>{x.u=x.target;});
    tiers.forEach(priority=>{
      const tier=plans.filter(x=>(x.b.priority||0)===priority);
      Object.keys(s.provinces).forEach(pid=>{
        Object.keys(s.resources[pid]).forEach(rid=>{
          const users=tier.filter(x=>x.p.id===+pid&&x.t.extraction.some(e=>e.resourceId===rid));const requested=users.reduce((n,x)=>n+x.t.extraction.find(e=>e.resourceId===rid).amount*x.b.level*x.u*x.efficiency,0);const available=s.remainingResources[pid][rid];if(requested>available&&requested>0)users.forEach(x=>{x.u*=available/requested;});
        });
      });
      tier.forEach(x=>{x.t.extraction.forEach(e=>{if(!s.resources[x.p.id][e.resourceId]){x.u=0;x.reasons.push('Нет источника '+e.resourceId);}});x.t.extraction.forEach(e=>{s.remainingResources[x.p.id][e.resourceId]-=e.amount*x.b.level*x.u*x.efficiency;});});
    });
    // Reassign workers freed by material or financial constraints within this same turn.
    plans.forEach(x=>{x.target=x.u;});
    this.allocateLabor(plans,s,'target');
    plans.forEach(x=>{x.u=x.target;x.t.workers.forEach(w=>{if(w.count)x.u=Math.min(x.u,x.hires.filter(e=>e.profession===w.profession).reduce((n,e)=>n+e.count,0)/(w.count*x.b.level));});});
    // Rebuild employment from actual production. Whole workers only; freed workers remain available next turn.
    s.population.forEach(g=>{g.workforce.free=g.workforce.available;g.workforce.byProfession={};});
    Object.values(s.provinces).forEach(p=>{p.transport.buildingCapacity=0;});
    plans.forEach(x=>{
      const actual=[];let u=x.u;
      x.t.workers.forEach(w=>{if(!w.count)return;let need=Math.floor(w.count*x.b.level*u+eps);x.hires.filter(e=>e.profession===w.profession).forEach(e=>{const n=Math.min(need,e.count);if(n){actual.push({populationId:e.populationId,profession:e.profession,count:n});need-=n;}});const count=actual.filter(e=>e.profession===w.profession).reduce((n,e)=>n+e.count,0);u=Math.min(u,count/(w.count*x.b.level));});
      // Trim all professions to the final common utilization.
      const final=[];x.t.workers.forEach(w=>{let need=Math.floor(w.count*x.b.level*u+eps);actual.filter(e=>e.profession===w.profession).forEach(e=>{const n=Math.min(need,e.count);if(n){final.push(Object.assign({},e,{count:n}));need-=n;}});});
      x.b.employment=final;
      final.forEach(e=>{const g=s.populationById[e.populationId],pay=e.count*x.b.salaries[e.profession];g.workforce.free-=e.count;g.workforce.byProfession[e.profession]=(g.workforce.byProfession[e.profession]||0)+e.count;g.money+=pay;x.b.money-=pay;x.b.lastTurn.expenses+=pay;});
      x.t.inputs.forEach(i=>{x.b.inventory[i.goodId]=(x.b.inventory[i.goodId]||0)-i.amount*x.b.level*u;});
      x.t.outputs.forEach(o=>{x.b.inventory[o.goodId]=(x.b.inventory[o.goodId]||0)+o.amount*x.b.level*u*x.efficiency;});
      x.t.extraction.forEach(e=>{const r=s.resources[x.p.id][e.resourceId];if(r&&r.mode!=='annual')r.amount-=e.amount*x.b.level*u*x.efficiency;});
      x.p.transport.buildingCapacity+=(x.t.transportCapacity||0)*x.b.level*u*x.efficiency;
      x.b.lastTurn.utilization=u;if(x.active&&u<x.labor-eps)x.reasons.push('Ограничение материалов, денег, источника или склада');
      const profit=x.b.lastTurn.revenue-x.b.lastTurn.expenses;
      x.t.workers.forEach(w=>{const old=x.b.salaries[w.profession];if(profit<0)x.b.salaries[w.profession]=Math.max(w.baseSalary,old*(1-ECONOMY_RULES.salaryStep));else if(x.workerShortage&&profit>0)x.b.salaries[w.profession]=old*(1+ECONOMY_RULES.salaryStep);});
      Object.keys(x.b.inventory).forEach(id=>{if(Math.abs(x.b.inventory[id])<eps)x.b.inventory[id]=0;});if(Math.abs(x.b.money)<eps)x.b.money=0;
    });
    Object.values(s.resources).forEach(map=>Object.values(map).forEach(r=>{if(r.mode==='renewable')r.amount=Math.min(r.maxAmount,r.amount+r.regeneration);}));
    s.markets.forEach(m=>{m.lastTurn={demand:{},supply:{}};s.goods.forEach(g=>{const k=m.id+':'+g.id,d=s.demand[k]||0,o=s.supply[k]||0;m.lastTurn.demand[g.id]=d;m.lastTurn.supply[g.id]=o;const imbalance=(d-o)/Math.max(d,o,1);const old=m.prices[g.id]||g.basePrice;m.prices[g.id]=Math.max(g.basePrice*ECONOMY_RULES.minPriceFactor,Math.min(g.basePrice*ECONOMY_RULES.maxPriceFactor,old*(1+ECONOMY_RULES.priceSpeed*imbalance)));});});
  },
  save:function(s,columns,limit){
    if(!s.buildings.length)return;
    // Validate calculated finances, populations and inventories before packing.
    s.buildings.forEach(b=>{this.number(b.money,'баланс '+b.id);Object.keys(b.inventory).forEach(id=>this.number(b.inventory[id],'склад '+b.id));if(this.weight(s,b)>s.templatesById[b.templateId].storageCapacity*b.level+ECONOMY_RULES.epsilon)this.fail('переполнен склад '+b.id);});
    s.population.forEach(g=>{if(g.workforce){this.number(g.workforce.free,'свободные рабочие',true);const employed=Object.values(g.workforce.byProfession).reduce((a,b)=>a+b,0);if(employed+g.workforce.free!==g.workforce.available)this.fail('нарушен баланс рабочих');}if(g.money!==undefined)this.number(g.money,'деньги населения');});
    s.countries.forEach(c=>this.number(c.treasury,'казна'));Object.values(s.resources).forEach(map=>Object.values(map).forEach(r=>{if(r.mode!=='annual')this.number(r.amount,'остаток ресурса');}));
    const packed={};Object.keys(s.lists).forEach(key=>{const h=ECONOMY_RULES.headers[key];packed[h]=this.pack(s[key],columns[h].length,limit,h);});
    packed[GAME_RULES.population.header]=this.pack(s.population,columns[GAME_RULES.population.header].length,limit,GAME_RULES.population.header);
    Object.keys(packed).forEach(h=>{columns[h]=packed[h];});columns[GAME_RULES.population.provinceHeader]=s.provinceCells;
  },
};
