const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),path=require('node:path');
const context={console,assert};vm.createContext(context);
vm.runInContext(['Game_Rules.gs','Population.gs','Economy.gs'].map(f=>fs.readFileSync(path.join(__dirname,'..',f),'utf8')).join('\n')+'\nfunction GameTurnDataError(m){this.message=m;} GameTurnDataError.prototype=Object.create(Error.prototype);',context);
vm.runInContext(`
const clone=x=>JSON.parse(JSON.stringify(x));
const fixture=()=>({Провинции:[{id:1,ownerId:'A',neighbors:[2],transport:{baseCapacity:1000,tariff:1,buildingCapacity:0},resources:[{id:'ore',mode:'finite',amount:1000}]},{id:2,ownerId:'B',neighbors:[1],transport:{baseCapacity:1000,tariff:1,buildingCapacity:0},resources:[]}],
Население:[[{id:'p1',provinceId:1,race:'Люди',culture:'c',religion:'r',men:100,women:100,total:200,birthRate:0,deathRate:0},{id:'p2',provinceId:2,race:'Люди',culture:'c',religion:'r',men:100,women:100,total:200,birthRate:0,deathRate:0}]],
Здания:[[{id:'mine',provinceId:1,templateId:'mine',level:1,money:1000,inventory:{}},{id:'supplier',provinceId:2,templateId:'supplier',level:1,money:100,inventory:{tools:10}}]],
'Шаблоны зданий':[[{id:'mine',workers:[{profession:'worker',count:10,baseSalary:2}],inputs:[{goodId:'tools',amount:10}],outputs:[{goodId:'iron',amount:100}],extraction:[{resourceId:'ore',amount:100}],storageCapacity:1000},{id:'supplier',workers:[],inputs:[],outputs:[{goodId:'tools',amount:10}],storageCapacity:1000}]],
Товары:[[{id:'tools',basePrice:5,transportWeight:1},{id:'iron',basePrice:2,transportWeight:1}]],
Страны:[[{id:'A',marketId:'A',treasury:0,workerCoefficients:{men:.5,women:.5}},{id:'B',marketId:'B',treasury:0,workerCoefficients:{men:.5,women:.5}}]],
Рынки:[[{id:'A',prices:{}},{id:'B',prices:{}}]],
'Торговые отношения':[[{from:'A',to:'B',tradeAllowed:true,embargo:false,transitAllowed:true,importTariff:.1,exportTariff:0},{from:'B',to:'A',tradeAllowed:true,embargo:false,transitAllowed:true,importTariff:0,exportTariff:.2}]]});
const run=c=>{const s=GameEconomy.load(c);GameEconomy.run(s);GameEconomy.save(s,c,48000);return s;};
const cash=s=>s.buildings.reduce((n,b)=>n+b.money,0)+s.countries.reduce((n,c)=>n+c.treasury,0)+s.population.reduce((n,g)=>n+(g.money||0),0);
let c=fixture(),s=run(c);
assert.equal(s.buildingsById.mine.inventory.iron,100);assert.equal(s.buildingsById.mine.inventory.tools,0);assert.equal(s.resources[1].ore.amount,900);
assert.equal(s.buildingsById.mine.money,895);assert.equal(s.buildingsById.supplier.money,150);assert.equal(s.countriesById.A.treasury,15);assert.equal(s.countriesById.B.treasury,20);assert.equal(s.populationById.p1.money,20);assert.equal(cash(s),1100);
assert.equal(s.buildingsById.supplier.inventory.tools,10);assert.equal(s.populationById.p1.workforce.free,90);
// New output cannot be bought this turn.
c=fixture();c.Здания[0][1].inventory.tools=0;s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,0);assert.equal(s.buildingsById.supplier.inventory.tools,10);
// Capacity applies at each endpoint; partial workers are whole people.
c=fixture();c.Провинции[1].transport.baseCapacity=5;s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,50);assert.equal(s.buildingsById.mine.employment[0].count,5);assert.equal(s.transport[2],0);assert.equal(cash(s),1100);
// Embargo overrides common market membership.
c=fixture();c.Страны[0][1].marketId='A';c['Торговые отношения'][0][0].embargo=true;s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,0);
// Finite, renewable, annual; extraction never exceeds available quantity.
c=fixture();c.Здания[0][0].inventory.tools=10;c.Провинции[0].resources[0].amount=50;s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,50);assert.equal(s.resources[1].ore.amount,0);
c=fixture();c.Здания[0][0].inventory.tools=10;c.Провинции[0].resources[0]={id:'ore',mode:'renewable',amount:100,maxAmount:200,regeneration:30};s=run(c);assert.equal(s.resources[1].ore.amount,30);
c=fixture();c.Здания[0][0].inventory.tools=10;c.Провинции[0].resources[0]={id:'ore',mode:'annual',annualCapacity:50};s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,50);assert.equal(s.resources[1].ore.annualCapacity,50);
// Storage accounts for simultaneous input consumption.
c=fixture();c['Шаблоны зданий'][0][0].storageCapacity=55;c.Здания[0][0].inventory.tools=10;s=run(c);assert(GameEconomy.weight(s,s.buildingsById.mine)<=55+1e-8);
// No debts; grants capped by treasury and policy budget.
c=fixture();c.Здания[0][0].money=0;c.Страны[0][0].treasury=30;c.Страны[0][0].subsidies={budget:100,templateIds:['mine']};s=run(c);assert(s.buildingsById.mine.money>=0);assert(s.countriesById.A.treasury>=0);assert.equal(s.buildingsById.mine.lastTurn.subsidy,30);assert(Math.abs(cash(s)-130)<1e-7);
// Nested requirements and modifiers.
c=fixture();c.Здания[0][0].inventory.tools=10;c['Шаблоны зданий'][0][0].requirements={all:[{field:'ownerId',operator:'eq',value:'A'},{not:{field:'id',operator:'eq',value:99}}]};c['Шаблоны зданий'][0][0].efficiencyModifiers=[{when:{field:'id',operator:'eq',value:1},multiplier:.5}];s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,50);assert.equal(s.resources[1].ore.amount,950);
// Missing setup and duplicate records fail explicitly.
c=fixture();delete c.Товары;assert.throws(()=>GameEconomy.load(c),/Товары/);
c=fixture();c.Здания[0].push(clone(c.Здания[0][0]));assert.throws(()=>GameEconomy.load(c),/повтор/);
c=fixture();c['Шаблоны зданий'][0][0].requirements={field:'id',operator:'bogus',value:1};assert.throws(()=>GameEconomy.load(c),/оператор/);
// Overflow does not mutate original loaded columns.
c=fixture();const before=JSON.stringify(c);s=GameEconomy.load(c);GameEconomy.run(s);assert.throws(()=>GameEconomy.save(s,c,10),/столбец/);assert.equal(JSON.stringify(c),before);
// Old games without buildings remain valid.
const old={Игровые_данные:{Провинции:[]}};assert.strictEqual(processEconomy(old,{}),old);
// Transit requires explicit permission and follows declared neighbors.
c=fixture();c.Провинции[0].neighbors=[3];c.Провинции[1].neighbors=[3];c.Провинции.push({id:3,ownerId:'C',neighbors:[1,2],transport:{baseCapacity:1000,tariff:2,buildingCapacity:0},resources:[]});c.Страны[0].push({id:'C',marketId:'A',treasury:0,workerCoefficients:{men:.5,women:.5}});s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,0);
c['Торговые отношения'][0].push({from:'C',to:'A',tradeAllowed:false,embargo:false,transitAllowed:true,importTariff:0,exportTariff:0});s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,100);assert.equal(s.countriesById.C.treasury,20);
// Equal-priority buyers split scarce goods and finite extraction proportionally.
c=fixture();const second=clone(c.Здания[0][0]);second.id='mine2';second.templateId='mine2';c.Здания[0].push(second);const t2=clone(c['Шаблоны зданий'][0][0]);t2.id='mine2';c['Шаблоны зданий'][0].push(t2);s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,50);assert.equal(s.buildingsById.mine2.inventory.iron,50);assert.equal(cash(s),2100);
// A higher-priority buyer gets first access.
c=fixture();c.Здания[0].push(clone(second));c['Шаблоны зданий'][0].push(clone(t2));c.Здания[0][0].priority=10;s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,100);assert.equal(s.buildingsById.mine2.inventory.iron,0);
// Released workers can be hired elsewhere during the same annual calculation.
c=fixture();c.Страны[0][0].workerCoefficients={men:.05,women:.05};c.Здания[0][0].priority=10;c.Здания[0][1].inventory.tools=0;const b2=clone(second);b2.inventory={};c.Здания[0].push(b2);const independent=clone(t2);independent.inputs=[];independent.extraction=[];c['Шаблоны зданий'][0].push(independent);s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,0);assert.equal(s.buildingsById.mine2.inventory.iron,100);assert.equal(s.populationById.p1.workforce.free,0);
// Prices react to declared demand/supply and stay bounded.
c=fixture();s=run(c);assert.equal(s.marketsById.A.prices.tools,6.25);assert.equal(s.marketsById.B.prices.tools,3.75);
// Profitable understaffed buildings raise pay; losses cannot lower it below base.
c=fixture();c.Страны[0][1].workerCoefficients={men:.025,women:.025};c['Шаблоны зданий'][0][1].workers=[{profession:'worker',count:10,baseSalary:2}];s=run(c);assert.equal(s.buildingsById.supplier.salaries.worker,2.1);assert.equal(s.buildingsById.mine.salaries.worker,2);
// Transport output is available only next turn.
c=fixture();c.Провинции[1].transport.baseCapacity=0;c['Шаблоны зданий'][0][1].transportCapacity=1000;s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,0);assert.equal(s.provinces[2].transport.buildingCapacity,1000);s=run(c);assert.equal(s.buildingsById.mine.inventory.iron,100);

console.log('PASS: trade snapshots, money conservation, tariffs, transit, capacity, population pay, extraction modes, storage, subsidies, conditions, invalid setup and packing rollback');
`,context);
