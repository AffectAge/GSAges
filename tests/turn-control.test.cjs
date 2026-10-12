const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),path=require('node:path');
let value=true,formula='',status='',runs=0,locked=false,available=true,fail=false,properties={},triggers=[],creates=0;
const request={getValue:()=>value,setValue:v=>{value=v;return request;},getFormula:()=>formula,getSheet:()=>sheet,getNumRows:()=>1,getNumColumns:()=>1,getA1Notation:()=> 'B2'};
const cell=kind=>({setValue:v=>{if(kind==='B3')status=v;return cell(kind);},clearContent(){},setNumberFormat(){}});
const sheet={getName:()=> 'Управление',getRange:addr=>addr==='B2'?request:cell(addr)};
const spreadsheet={getId:()=> 'sheet',getSheetByName:()=>sheet};
const lock={tryLock:()=>{if(locked||!available)return false;locked=true;return true;},waitLock:()=>{assert(!locked);locked=true;},releaseLock:()=>{locked=false;}};
const trigger=id=>({getHandlerFunction:()=> 'HANDLE_GAME_TURN_EDIT',getEventType:()=> 'ON_EDIT',getTriggerSourceId:()=> 'sheet',getUniqueId:()=>id});
const context={assert,console,PropertiesService:{getScriptProperties:()=>({getProperty:k=>properties[k],setProperty:(k,v)=>{properties[k]=v;}})},LockService:{getDocumentLock:()=>lock},SpreadsheetApp:{getActiveSpreadsheet:()=>spreadsheet,flush(){}},GameTurnEngine:{_processUnlocked:ss=>{assert(locked);assert.equal(value,false);assert.strictEqual(ss,spreadsheet);runs++;if(fail)throw Error('broken');return {nextTurn:runs+1};}},ScriptApp:{EventType:{ON_EDIT:'ON_EDIT'},getProjectTriggers:()=>triggers,newTrigger:handler=>{assert.equal(handler,'HANDLE_GAME_TURN_EDIT');return {forSpreadsheet:ss=>{assert.strictEqual(ss,spreadsheet);return {onEdit:()=>({create:()=>{const t=trigger('new'+ ++creates);triggers.push(t);return t;}})};}};},deleteTrigger:t=>{triggers=triggers.filter(x=>x!==t);}}};
vm.createContext(context);vm.runInContext(fs.readFileSync(path.join(__dirname,'..','GameTurnControl.gs'),'utf8'),context);
const event=(extras={})=>Object.assign({source:spreadsheet,range:request,value:'TRUE'},extras);
// Simple edits never start the engine, whether no trigger or a stale property.
context.onEdit(event());assert.equal(runs,0);assert.equal(value,true);assert.match(status,/Нужен установленный/);
properties.GAME_TURN_EDIT_TRIGGER_sheet='deleted';context.onEdit(event());assert.equal(runs,0);assert.match(status,/Ожидается/);
// Installer repairs stale markers and is idempotent.
context.INSTALL_GAME_TURN_EDIT_TRIGGER();assert.equal(creates,1);assert.equal(properties.GAME_TURN_EDIT_TRIGGER_sheet,'new1');
context.INSTALL_GAME_TURN_EDIT_TRIGGER();assert.equal(creates,1);assert.equal(runs,0);
// Authorized edit processes exactly one turn, resets request, reports completion.
context.HANDLE_GAME_TURN_EDIT(event({triggerUid:'new1'}));assert.equal(runs,1);assert.equal(value,false);assert.match(status,/Ход завершён/);assert.equal(locked,false);
context.HANDLE_GAME_TURN_EDIT(event({triggerUid:'new1'}));assert.equal(runs,1);
// Installer attached to onEdit itself also works without the disabled-simple branch.
value=true;context.onEdit(event({triggerUid:'manual'}));assert.equal(runs,2);
// Localized and whitespace values, and events missing value, use the actual cell.
value=' ИСТИНА ';context.HANDLE_GAME_TURN_EDIT(event({value:' ИСТИНА ',triggerUid:'new1'}));assert.equal(runs,3);
value=true;context.HANDLE_GAME_TURN_EDIT(event({value:undefined,triggerUid:'new1'}));assert.equal(runs,4);
// Reject formulas, FALSE, unrelated/multicell events, and requests during another turn.
value=true;formula='=TRUE()';context.HANDLE_GAME_TURN_EDIT(event());assert.equal(runs,4);formula='';
context.HANDLE_GAME_TURN_EDIT(event({value:'FALSE'}));assert.equal(runs,4);
context.HANDLE_GAME_TURN_EDIT(event({range:Object.assign({},request,{getA1Notation:()=> 'B5'})}));assert.equal(runs,4);
context.HANDLE_GAME_TURN_EDIT(event({range:Object.assign({},request,{getNumRows:()=>2})}));assert.equal(runs,4);
available=false;context.HANDLE_GAME_TURN_EDIT(event());context.onEdit(event());assert.equal(runs,4);available=true;
// Late simple-trigger delivery cannot overwrite the installed handler result.
value=false;const completed=status;context.onEdit(event());assert.equal(status,completed);
// Errors reach both caller and sheet; lock and request are reset.
value=true;fail=true;assert.throws(()=>context.HANDLE_GAME_TURN_EDIT(event({triggerUid:'new1'})),/broken/);assert.equal(value,false);assert.equal(locked,false);assert.match(status,/Ошибка: broken/);
// Visible duplicate triggers are removed.
fail=false;triggers.push(trigger('duplicate'));context.INSTALL_GAME_TURN_EDIT_TRIGGER();assert.equal(triggers.length,1);assert.equal(creates,1);
console.log('PASS: installed-only execution, stale-trigger repair, idempotence, duplicate protection, localized values, error handling and lock release');
