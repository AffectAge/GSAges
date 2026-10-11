/** Editable starter catalog. Initialization fills only entirely empty columns. */
const ECONOMY_STARTER_CATALOG = {
  goods: [
    {id:'tools',name:'Инструменты',basePrice:20,transportWeight:1},
    {id:'iron_ore',name:'Железная руда',basePrice:10,transportWeight:1},
    {id:'wood',name:'Древесина',basePrice:5,transportWeight:2},
    {id:'grain',name:'Зерно',basePrice:3,transportWeight:1},
  ],
  templates: [
    {id:'iron_mine',name:'Железная шахта',workers:[{profession:'worker',count:90,baseSalary:10},{profession:'engineer',count:10,baseSalary:20}],inputs:[{goodId:'tools',amount:20}],outputs:[{goodId:'iron_ore',amount:1000}],extraction:[{resourceId:'Железная руда',amount:1000}],storageCapacity:5000,requirements:{all:[{field:'terrainType',operator:'in',value:['Суша','Остров']},{resourceId:'Железная руда',operator:'available'}]},efficiencyModifiers:[{when:{field:'landscapeType',operator:'eq',value:'Горы'},multiplier:.8}]},
    {id:'logging',name:'Лесозаготовка',workers:[{profession:'worker',count:100,baseSalary:10}],inputs:[{goodId:'tools',amount:10}],outputs:[{goodId:'wood',amount:800}],extraction:[{resourceId:'forest',amount:800}],storageCapacity:5000},
    {id:'farm',name:'Ферма',workers:[{profession:'farmer',count:100,baseSalary:10}],inputs:[{goodId:'tools',amount:10}],outputs:[{goodId:'grain',amount:1000}],extraction:[{resourceId:'farmland',amount:1000}],storageCapacity:5000,efficiencyModifiers:[{when:{field:'fertility',operator:'lt',value:.3},multiplier:.5}]},
    {id:'tool_workshop',name:'Мастерская инструментов',workers:[{profession:'worker',count:20,baseSalary:10}],inputs:[{goodId:'iron_ore',amount:100},{goodId:'wood',amount:50}],outputs:[{goodId:'tools',amount:200}],storageCapacity:2000},
    {id:'transport_hub',name:'Транспортный узел',workers:[{profession:'worker',count:100,baseSalary:10}],inputs:[{goodId:'tools',amount:10}],outputs:[],storageCapacity:1000,transportCapacity:10000},
  ],
};

/** Adds headers to the existing named range, without running a turn or creating buildings. */
function INITIALIZE_BUILDING_ECONOMY() {
  const lock=LockService.getDocumentLock();lock.waitLock(GAME_TURN_CONFIG.lockTimeoutMs);
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet(),range=ss.getRangeByName(ECONOMY_RULES.rangeName);
    if(!range)GameEconomy.fail('не найден диапазон «'+ECONOMY_RULES.rangeName+'».');
    const rows=range.getNumRows(),sheet=range.getSheet(),matrix=range.getValues(),headers=matrix[0].map(String);
    if(rows<2)GameEconomy.fail('нужны заголовки и строки данных');
    const missing=Object.values(ECONOMY_RULES.headers).filter(h=>!headers.includes(h));
    // Pre-serialize catalogs before any spreadsheet mutation.
    const catalog={};Object.keys(ECONOMY_STARTER_CATALOG).forEach(k=>{const h=ECONOMY_RULES.headers[k];const index=headers.indexOf(h);if(index<0||matrix.slice(1).every(r=>r[index]===''))catalog[h]=GameEconomy.pack(ECONOMY_STARTER_CATALOG[k],rows-1,GAME_TURN_CONFIG.maxJsonCellCharacters,h).map(x=>[x?JSON.stringify(x):'']);});
    Object.keys(catalog).forEach(h=>{const i=headers.indexOf(h);if(i>=0&&range.offset(1,i,rows-1,1).getFormulas().some(r=>r[0]))GameEconomy.fail('формулы в столбце '+h);});
    if(missing.length){
      // Inserting dedicated columns avoids overwriting neighboring user data.
      const last=range.getColumn()+range.getNumColumns()-1;
      sheet.insertColumnsAfter(last,missing.length);
      sheet.getRange(range.getRow(),last+1,1,missing.length).setValues([missing]);
      ss.setNamedRange(ECONOMY_RULES.rangeName,sheet.getRange(range.getRow(),range.getColumn(),rows,range.getNumColumns()+missing.length));
    }
    const target=ss.getRangeByName(ECONOMY_RULES.rangeName),current=target.getValues()[0].map(String);
    Object.keys(catalog).forEach(h=>{const col=target.offset(1,current.indexOf(h),rows-1,1);if(col.getFormulas().some(r=>r[0]))GameEconomy.fail('формулы в столбце '+h);col.setValues(catalog[h]);});
    return {addedHeaders:missing,catalogColumns:Object.keys(catalog),message:'Укажите страны, рынки, транспорт и здания. Ход не запускался.'};
  }finally{lock.releaseLock();}
}
