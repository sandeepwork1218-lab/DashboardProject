// ============================================================
var SHEET_NAME_OFFICE     = "Office_Laptops";
var SHEET_NAME_CLIENT     = "Client_Laptops";
var SHEET_NAME_OTHER      = "Other_Assets_IT";
var SHEET_NAME_MOBILE     = "Mobile";
var SHEET_NAME_STUDIO     = "Studio_Items";
var SHEET_NAME_COURIER    = "Courier_Data";
var SHEET_NAME_MATERIALIN = "Material_IN";
var SHEET_NAME_ASSET_DETAILS_LAPTOP = "Laptop";
var SHEET_NAME_ASSET_DETAILS_DESKTOP = "Desktop";
var ASSET_DETAILS_DESKTOP_SHEET_ID = 1746060339;
var ASSET_DETAILS_SPREADSHEET_ID = "1JMexH_OIRWI7tjAiV3xvP4euJoMN0n_ZBIUbCY-Wb5o";
var SHEET_NAME_MAC_OFFICE   = "MAC";
var SHEET_NAME_MAC_CLIENT   = "Client-Laptop";
var SHEET_NAME_MAC_FIREWALL = "Adobe User On Firewall";
var MAC_SPREADSHEET_ID = "1GBX4nc1Db9Tso6FEjTAwwLZCM_44VLl412LH2XgpnPU";
var DASHBOARD_ADMIN_EMAIL='sandeep01@cyntexa.com';
// Dashboard access approval removed.
// Dashboard web-app links now open the dashboard directly.
// Keep admin identity for the separate activity-log administration controls.
function getCurrentDashboardUserEmail_(){try{return String(Session.getActiveUser().getEmail()||'').trim().toLowerCase();}catch(e){return '';}}
function isDashboardAdmin_(email){return String(email||'').trim().toLowerCase()===DASHBOARD_ADMIN_EMAIL.toLowerCase();}
var DASHBOARD_ACCESS_USERS_KEY='DASHBOARD_AUTHORIZED_USERS';
// ============================================================
// HARDENED DEPLOYMENT / PRIVATE DATA SOURCE
// ============================================================
// IMPORTANT: This project is intended to run as a STANDALONE Web App,
// not as a spreadsheet-bound script. Set the main spreadsheet ID once
// with configureMainSpreadsheet(spreadsheetId), then deploy the Web App
// to execute as the owner. Keep this Apps Script project private.
var MAIN_SPREADSHEET_PROPERTY_KEY_ = 'ASSET_DASHBOARD_MAIN_SPREADSHEET_ID';
var WRITE_RETRY_COUNT_ = 3;
var WRITE_RETRY_SLEEP_MS_ = 250;

function getMainSpreadsheet_(){
  var id=PropertiesService.getScriptProperties().getProperty(MAIN_SPREADSHEET_PROPERTY_KEY_);
  if(!id) throw new Error('Main Asset Register spreadsheet is not configured. Admin must run configureMainSpreadsheet(spreadsheetId) once.');
  try{return SpreadsheetApp.openById(id);}catch(e){throw new Error('Configured main spreadsheet could not be opened. Check the Spreadsheet ID and script permissions.');}
}
//Only admin can congigure........./
function configureMainSpreadsheet(spreadsheetId){
  var email=getCurrentDashboardUserEmail_();
  if(email && !isDashboardAdmin_(email)) throw new Error('Only administrator can configure the main spreadsheet.');
  spreadsheetId=String(spreadsheetId||'').trim();
  if(!/^[A-Za-z0-9_-]{20,}$/.test(spreadsheetId)) throw new Error('Please provide a valid Google Spreadsheet ID.');
  var ss=SpreadsheetApp.openById(spreadsheetId);
  PropertiesService.getScriptProperties().setProperty(MAIN_SPREADSHEET_PROPERTY_KEY_,ss.getId());
  PropertiesService.getScriptProperties().setProperty(ACTIVITY_LOG_MAIN_SPREADSHEET_KEY,ss.getId());
  return {ok:true,spreadsheetId:ss.getId(),name:ss.getName(),url:ss.getUrl()};
}
// checking the verification......../
function verifyWrittenRow_(sheet,row,expected){
  if(!sheet || !row || row<2) throw new Error('Write verification failed: invalid target row.');
  SpreadsheetApp.flush();
  var values=sheet.getRange(row,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  expected=expected||{};
  Object.keys(expected).forEach(function(key){
    var col=Number(key),want=String(expected[key]===undefined||expected[key]===null?'':expected[key]).trim();
    if(!col || col>values.length) throw new Error('Write verification failed: invalid verification column.');
    var got=String(values[col-1]||'').trim();
    if(got!==want) throw new Error('Write verification failed. The sheet did not retain the submitted value.');
  });
  return true;
}
function withWriteRetry_(fn){
  var lastError=null;
  for(var attempt=1;attempt<=WRITE_RETRY_COUNT_;attempt++){
    try{return fn(attempt);}catch(e){lastError=e;if(attempt<WRITE_RETRY_COUNT_)Utilities.sleep(WRITE_RETRY_SLEEP_MS_*attempt);}
  }
  throw lastError||new Error('Write failed.');
}

function dashboardAuthorizedUsers_(){var r=PropertiesService.getScriptProperties().getProperty(DASHBOARD_ACCESS_USERS_KEY)||'[]';try{var a=JSON.parse(r);return Array.isArray(a)?a:[];}catch(e){return [];}}
function setDashboardAuthorizedUsers_(a){PropertiesService.getScriptProperties().setProperty(DASHBOARD_ACCESS_USERS_KEY,JSON.stringify(a||[]));}
function isDashboardUserAuthorized_(email){email=String(email||'').trim().toLowerCase();return !!email&&(isDashboardAdmin_(email)||dashboardAuthorizedUsers_().indexOf(email)>=0);}
function requireDashboardAccess_(){var e=getCurrentDashboardUserEmail_();if(!e)throw new Error('Google account email could not be identified. Please sign in with your Google account.');if(!isDashboardUserAuthorized_(e))throw new Error('Dashboard access is not enabled for '+e+'.');return e;}
function getDashboardAccessStatus(){
  requireDashboardAccess_();var e=getCurrentDashboardUserEmail_();return{email:e,admin:isDashboardAdmin_(e),authorized:isDashboardUserAuthorized_(e)};}
function getDashboardAccessData(){
  requireDashboardAccess_();
  if(!isDashboardAdmin_(getCurrentDashboardUserEmail_()))throw new Error('Only administrator can view access settings.');
  var users=dashboardAuthorizedUsers_();
  return {total:users.length,counts:{AUTHORIZED:users.length},records:users.map(function(email){return {email:email,status:'AUTHORIZED',code:email};}),updated:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy, hh:mm a')};
}
function addDashboardAuthorizedEmail(email){
  requireDashboardAccess_();
  if(!isDashboardAdmin_(getCurrentDashboardUserEmail_()))throw new Error('Only administrator can manage access.');
  email=String(email||'').trim().toLowerCase();
  if(!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Please enter a valid email address.');
  if(email===DASHBOARD_ADMIN_EMAIL.toLowerCase())return{ok:true,exists:true,message:'Admin email is already allowed.'};
  var a=dashboardAuthorizedUsers_();
  if(a.indexOf(email)<0)a.push(email);
  setDashboardAuthorizedUsers_(a);
  return{ok:true,exists:a.indexOf(email)>=0,email:email};
}
//Access Details (Functionality)......../
function removeDashboardAuthorizedEmail(email){
  requireDashboardAccess_();
  if(!isDashboardAdmin_(getCurrentDashboardUserEmail_()))throw new Error('Only administrator can manage access.');
  email=String(email||'').trim().toLowerCase();
  if(email===DASHBOARD_ADMIN_EMAIL.toLowerCase())throw new Error('Administrator access cannot be removed.');
  setDashboardAuthorizedUsers_(dashboardAuthorizedUsers_().filter(function(x){return x!==email;}));
  return{ok:true};
}
function listDashboardAuthorizedEmails(){
  requireDashboardAccess_();
  if(!isDashboardAdmin_(getCurrentDashboardUserEmail_()))throw new Error('Only administrator can manage access.');
  return{users:dashboardAuthorizedUsers_(),admin:DASHBOARD_ADMIN_EMAIL};
}

function onOpen(){SpreadsheetApp.getUi().createMenu("📊 Dashboard").addItem("Open Asset Dashboard","showDashboard").addSeparator().addItem("Enable change logging","setupActivityAuditLog").addToUi();}
function showDashboard(){
  requireDashboardAccess_();requireDashboardAccess_();var html=HtmlService.createHtmlOutputFromFile("dashboard").setWidth(1200).setHeight(780);SpreadsheetApp.getUi().showModalDialog(html,"Asset Dashboard");}
function dashboardHtml_(){
  return HtmlService.createHtmlOutputFromFile("dashboard")
    .setTitle("Asset Dashboard")
    .addMetaTag("viewport","width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doGet(e){
  var email=getCurrentDashboardUserEmail_();
  if(isDashboardUserAuthorized_(email))return dashboardHtml_();
  if(!email)return accessPage_('Google account required','Please open this dashboard while signed in to your company Google account.',false);
  return accessPage_('Access denied','Your Google account <b>'+esc_(email)+'</b> is not authorized to access this dashboard. Please contact the administrator.',false);
}
function accessPage_(title,msg,request){return HtmlService.createHtmlOutput('<style>body{font-family:Arial;background:#f5f7fb;display:flex;justify-content:center;align-items:center;height:100vh}.c{background:white;padding:30px;border-radius:15px;box-shadow:0 5px 25px #ccc;max-width:560px}h2{margin-top:0}</style><div class="c"><h2>'+title+'</h2><p>'+msg+'</p></div>').setTitle('Dashboard Access');}
function esc_(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

// ============================================================
// CENTRAL ACTIVITY AUDIT LOG — SHARED MAIN WORKBOOK
//
// Activity_Log lives in the main Asset Register spreadsheet so every
// dashboard user reads and writes the same log. No per-user/per-workbook
// audit log is created. The main spreadsheet ID is remembered once by the
// admin when Enable change logging is run.
// ============================================================
var ACTIVITY_LOG_SHEET_NAME = 'Activity_Log';
var ACTIVITY_LOG_MAIN_SPREADSHEET_KEY = 'ASSET_DASHBOARD_MAIN_SPREADSHEET_ID';
var ACTIVITY_LOG_PROPERTY_KEY = 'ASSET_DASHBOARD_AUDIT_LOG_WORKBOOK_ID';
var ACTIVITY_LOG_DOCUMENT_PROPERTY_KEY = 'ASSET_DASHBOARD_AUDIT_LOG_DOCUMENT_ID';
var ACTIVITY_LOG_WORKBOOK_NAME = 'Asset Dashboard - Audit Logs';
var ACTIVITY_LOG_HEADERS_ = ['Timestamp','User','Source Workbook','Workspace','Tab','Action','Record ID','Employee ID','Employee Name','Field','Old Value','New Value','Row'];

function rememberActivityLogSpreadsheet_(id){
  if(!id)return;
  PropertiesService.getScriptProperties().setProperty(ACTIVITY_LOG_PROPERTY_KEY,id);
  try{PropertiesService.getDocumentProperties().setProperty(ACTIVITY_LOG_DOCUMENT_PROPERTY_KEY,id);}catch(ignore){}
}
function rememberedActivityLogSpreadsheetId_(){
  var id=PropertiesService.getScriptProperties().getProperty(ACTIVITY_LOG_PROPERTY_KEY);
  if(id)return id;
  try{return PropertiesService.getDocumentProperties().getProperty(ACTIVITY_LOG_DOCUMENT_PROPERTY_KEY)||'';}catch(ignore){return '';}
}
function findExistingActivityLogSpreadsheet_(){
  var files=DriveApp.getFilesByName(ACTIVITY_LOG_WORKBOOK_NAME),best=null,bestRows=-1,bestUpdated=0;
  while(files.hasNext()){
    var file=files.next();
    try{
      var ss=SpreadsheetApp.openById(file.getId()),sheet=ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME),rows=sheet?sheet.getLastRow():0,updated=file.getLastUpdated().getTime();
      if(sheet && (rows>bestRows || (rows===bestRows && updated>=bestUpdated))){best=ss;bestRows=rows;bestUpdated=updated;}
    }catch(ignore){}
  }
  return best;
}

function activityMainSpreadsheet_(){
  var props=PropertiesService.getScriptProperties();
  var id=props.getProperty(ACTIVITY_LOG_MAIN_SPREADSHEET_KEY);
  if(id){
    try{return SpreadsheetApp.openById(id);}catch(ignore){props.deleteProperty(ACTIVITY_LOG_MAIN_SPREADSHEET_KEY);}
  }
  var configured=PropertiesService.getScriptProperties().getProperty(MAIN_SPREADSHEET_PROPERTY_KEY_);
  if(configured){
    try{return SpreadsheetApp.openById(configured);}catch(ignore2){props.deleteProperty(MAIN_SPREADSHEET_PROPERTY_KEY_);}
  }
  throw new Error('Main Asset Register spreadsheet is not configured. Admin must run configureMainSpreadsheet(spreadsheetId) once.');
}

function activityLogSpreadsheet_(){
  // IMPORTANT: always use the same main Asset Register workbook.
  // Do not fall back to a per-user Audit Logs workbook here.
  var ss=activityMainSpreadsheet_();
  rememberActivityLogSpreadsheet_(ss.getId());
  return ss;
}

function getActivityLogSheet_(){
  var ss=activityLogSpreadsheet_(),sh=ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
  if(!sh){
    sh=ss.insertSheet(ACTIVITY_LOG_SHEET_NAME);
  }
  if(sh.getLastRow()===0){
    sh.getRange(1,1,1,ACTIVITY_LOG_HEADERS_.length).setValues([ACTIVITY_LOG_HEADERS_]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function migrateLegacyActivityLogs_(mainSs,legacyId){
  var props=PropertiesService.getScriptProperties();

  var legacy=null;
  // Prefer the previously remembered separate audit workbook, then search by name.
  var oldId=String(legacyId||'').trim();
  if(oldId && oldId!==mainSs.getId()){
    try{legacy=SpreadsheetApp.openById(oldId);}catch(ignore){}
  }
  if(!legacy)legacy=findExistingActivityLogSpreadsheet_();
  if(!legacy || legacy.getId()===mainSs.getId()){
    props.setProperty('ACTIVITY_LOG_LEGACY_MIGRATED','1');
    return;
  }

  var oldSheet=legacy.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
  var newSheet=mainSs.getSheetByName(ACTIVITY_LOG_SHEET_NAME)||mainSs.insertSheet(ACTIVITY_LOG_SHEET_NAME);
  if(newSheet.getLastRow()===0){
    newSheet.getRange(1,1,1,ACTIVITY_LOG_HEADERS_.length).setValues([ACTIVITY_LOG_HEADERS_]);
    newSheet.setFrozenRows(1);
  }
  if(!oldSheet || oldSheet.getLastRow()<2){
    props.setProperty('ACTIVITY_LOG_LEGACY_MIGRATED','1');
    return;
  }

  var oldLastCol=Math.min(oldSheet.getLastColumn(),ACTIVITY_LOG_HEADERS_.length);
  if(oldLastCol<1){props.setProperty('ACTIVITY_LOG_LEGACY_MIGRATED','1');return;}
  var oldRows=oldSheet.getRange(2,1,oldSheet.getLastRow()-1,oldLastCol).getValues();
  var existing=newSheet.getLastRow()>1?newSheet.getRange(2,1,newSheet.getLastRow()-1,oldLastCol).getValues():[];
  var seen={};
  existing.forEach(function(row){seen[JSON.stringify(row)]=true;});
  var toAppend=[];
  oldRows.forEach(function(row){
    var full=row.slice();
    while(full.length<ACTIVITY_LOG_HEADERS_.length)full.push('');
    var key=JSON.stringify(full);
    if(!seen[key]){seen[key]=true;toAppend.push(full);}
  });
  if(toAppend.length){
    newSheet.getRange(newSheet.getLastRow()+1,1,toAppend.length,ACTIVITY_LOG_HEADERS_.length).setValues(toAppend);
    SpreadsheetApp.flush();
  }
  props.setProperty('ACTIVITY_LOG_LEGACY_MIGRATED','1');
}

function activityTrackedWorkbooks_(){
  var main=activityMainSpreadsheet_();
  var list=[{ss:main,workspace:'Asset Register'}],seen={};
  [{id:MAC_SPREADSHEET_ID,workspace:'MAC Address'},{id:ASSET_DETAILS_SPREADSHEET_ID,workspace:'Asset Details'}].forEach(function(item){
    if(item.id) list.push({ss:SpreadsheetApp.openById(item.id),workspace:item.workspace});
  });
  return list.filter(function(item){var id=item.ss.getId();if(seen[id])return false;seen[id]=true;return true;});
}
function setupActivityAuditLog(){
  requireDashboardAccess_();
  if(!isDashboardAdmin_(getCurrentDashboardUserEmail_()))throw new Error('Only administrator can enable change logging.');

  // Capture the old separate-log ID BEFORE switching the log pointer to the main workbook.
  var legacyId=rememberedActivityLogSpreadsheetId_();
  // Capture the ONE main workbook while the admin is running setup from it.
  var mainSs=getMainSpreadsheet_();
  if(!mainSs)throw new Error('Open the main Asset Register spreadsheet and run Enable change logging from there.');
  PropertiesService.getScriptProperties().setProperty(ACTIVITY_LOG_MAIN_SPREADSHEET_KEY,mainSs.getId());

  // Create the shared Activity_Log tab and migrate any history from the old
  // separate Audit Log workbook, without changing any other dashboard data.
  var logSheet=getActivityLogSheet_();
  migrateLegacyActivityLogs_(mainSs,legacyId);
  var logId=mainSs.getId(),tracked=activityTrackedWorkbooks_();

  installActivityAuditTriggers_(tracked);
  return {ok:true,workbooks:tracked.length,logUrl:mainSs.getUrl(),message:'Shared Activity_Log enabled. Manual Sheet edits will also be logged.'};
}

function installActivityAuditTriggers_(tracked){
  var seen={};
  // Remove ONLY our audit triggers. Other project triggers are left untouched.
  ScriptApp.getProjectTriggers().forEach(function(trigger){
    if(trigger.getHandlerFunction()==='auditSheetEdit'){
      try{ScriptApp.deleteTrigger(trigger);}catch(ignore){}
    }
  });
  (tracked||[]).forEach(function(item){
    if(!item||!item.ss)return;
    var id=item.ss.getId();
    if(seen[id])return;
    seen[id]=true;
    // Installable onEdit trigger: this is what captures direct/manual edits
    // made in Google Sheets, including edits by other users.
    ScriptApp.newTrigger('auditSheetEdit').forSpreadsheet(item.ss).onEdit().create();
  });
}

function repairActivityAuditTriggers(){
  requireDashboardAccess_();
  if(!isDashboardAdmin_(getCurrentDashboardUserEmail_()))throw new Error('Only administrator can repair change-logging triggers.');
  var mainSs=activityMainSpreadsheet_();
  var tracked=activityTrackedWorkbooks_();
  installActivityAuditTriggers_(tracked);
  return {ok:true,workbooks:tracked.length,triggerCount:ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==='auditSheetEdit';}).length,logUrl:mainSs.getUrl(),message:'Manual Sheet change logging triggers repaired.'};
}

function getActivityLogSetup(){
  requireDashboardAccess_();
  var log=activityLogSpreadsheet_(),count=0;
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='auditSheetEdit')count++;});
  return {enabled:count>0,triggerCount:count,logUrl:log.getUrl(),spreadsheetId:log.getId(),sheetName:ACTIVITY_LOG_SHEET_NAME};
}
function activityWorkspace_(ss){
  var id=ss.getId();
  if(id===MAC_SPREADSHEET_ID)return 'MAC Address';
  if(id===ASSET_DETAILS_SPREADSHEET_ID)return 'Asset Details';
  return 'Asset Register';
}
function activityEditAction_(field,oldValue,newValue,isSingle){
  var header=String(field||'').toUpperCase(),value=String(newValue||'').toUpperCase();
  if(/ASSET CODE|^CODE$|SERIAL NO|SERIAL NUMBER/.test(header) && !oldValue)return 'RECORD CREATED';
  if(isSingle && /STATUS/.test(header) && (value==='SUBMITTED'||value==='NOT ASSIGNED'))return 'ASSET SUBMITTED';
  if(isSingle && /STATUS/.test(header) && value==='ASSIGNED')return 'ASSET ASSIGNED';
  if(isSingle && /EMP ID|EMPLOYEE ID|EMP NAME|EMPLOYEE NAME|USER NAME|ASSIGNED TO/.test(header))return 'ASSIGNMENT UPDATED';
  return isSingle?'UPDATED':'BULK UPDATED';
}
function activityValue_(headers,values,names){for(var i=0;i<headers.length;i++){var h=String(headers[i]||'').toUpperCase();for(var j=0;j<names.length;j++)if(h.indexOf(names[j])>-1)return String(values[i]||'');}return '';}
function activitySnapshot_(sheet,row){var lastCol=Math.max(1,sheet.getLastColumn());return{headers:sheet.getRange(1,1,1,lastCol).getDisplayValues()[0],values:sheet.getRange(row,1,1,lastCol).getDisplayValues()[0],row:row};}
function activityUser_(event){try{if(event&&event.user&&event.user.getEmail())return event.user.getEmail();}catch(ignore){}try{return Session.getActiveUser().getEmail()||'Dashboard';}catch(ignore2){return 'Dashboard';}}
function activityItemsForRow_(sheet,row,action,before,user){
  var after=activitySnapshot_(sheet,row),headers=after.headers,oldValues=before&&before.values?before.values:[],items=[];
  for(var c=0;c<after.values.length;c++){
    var oldValue=before?String(oldValues[c]||''):'' ,newValue=String(after.values[c]||'');
    if(before && oldValue===newValue)continue;
    if(!before && !newValue)continue;
    items.push({user:user||activityUser_(),sourceWorkbook:sheet.getParent().getName(),workspace:activityWorkspace_(sheet.getParent()),tab:sheet.getName(),action:action,recordId:activityValue_(headers,after.values,['ASSET CODE','CODE','SERIAL NO','SERIAL NUMBER','ITEM NAME','CONSIGNMENT','MAC ADDRESS','EMP ID','EMPLOYEE ID']),empId:activityValue_(headers,after.values,['EMP ID','EMPLOYEE ID','EMP CODE','EMPLOYEE CODE']),empName:activityValue_(headers,after.values,['EMP NAME','EMPLOYEE NAME','USER NAME','ASSIGNED TO','CHECKED BY','NAME']),field:headers[c]||('Column '+(c+1)),oldValue:oldValue,newValue:newValue,row:row});
  }
  return items;
}
function appendActivities_(items){
  if(!items||!items.length)return;
  var lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    var sheet=getActivityLogSheet_(),rows=items.map(function(item){return[new Date(),item.user||activityUser_(),item.sourceWorkbook||'',item.workspace||'',item.tab||'',item.action||'UPDATED',item.recordId||'',item.empId||'',item.empName||'',item.field||'',item.oldValue||'',item.newValue||'',item.row||''];});
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,ACTIVITY_LOG_HEADERS_.length).setValues(rows);
    SpreadsheetApp.flush();
  }finally{lock.releaseLock();}
}
// A dashboard form submission is one business process, even when it updates
// several cells. Keep it as one searchable log record instead of one record
// per field (for example: Assign laptop, Submit laptop, Add new asset).
function activityProcessItem_(sheet,row,action,before){
  var after=activitySnapshot_(sheet,row),changed=[],oldParts=[],newParts=[];
  for(var c=0;c<after.values.length;c++){
    var oldValue=before?String((before.values[c]===undefined||before.values[c]===null)?'':before.values[c]):'',newValue=String((after.values[c]===undefined||after.values[c]===null)?'':after.values[c]);
    if(before && oldValue===newValue)continue;
    if(!before && !newValue)continue;
    changed.push(c);
    var label=after.headers[c]||('Column '+(c+1));
    if(before && oldValue)oldParts.push(label+': '+oldValue);
    if(newValue)newParts.push(label+': '+newValue);
  }
  if(before && !changed.length)return null;
  var recordId=activityValue_(after.headers,after.values,['ASSET CODE','CODE','SERIAL NO','SERIAL NUMBER','ITEM NAME','CONSIGNMENT','MAC ADDRESS','EMP ID','EMPLOYEE ID']);
  return {user:activityUser_(),sourceWorkbook:sheet.getParent().getName(),workspace:activityWorkspace_(sheet.getParent()),tab:sheet.getName(),action:action,recordId:recordId,empId:activityValue_(after.headers,after.values,['EMP ID','EMPLOYEE ID','EMP CODE','EMPLOYEE CODE']),empName:activityValue_(after.headers,after.values,['EMP NAME','EMPLOYEE NAME','USER NAME','ASSIGNED TO','CHECKED BY','NAME']),field:'Process ('+changed.length+' field'+(changed.length===1?'':'s')+')',oldValue:before?(oldParts.join(' · ')||'No previous value'):'New record',newValue:newParts.join(' · ')||action,row:row};
}
function auditDashboardRow_(sheet,row,action,before){var item=activityProcessItem_(sheet,row,action,before);if(item)appendActivities_([item]);}
function auditSheetEdit(e){
  if(!e||!e.range)return;
  var range=e.range,sh=range.getSheet(),logId=activityLogSpreadsheet_().getId();
  // Only ignore edits made INSIDE the Activity_Log tab itself. Other tabs in
  // the main workbook must continue to be audited.
  if(sh.getParent().getId()===logId && sh.getName()===ACTIVITY_LOG_SHEET_NAME)return;
  if(range.getRow()===1)return;
  var startRow=Math.max(2,range.getRow()),endRow=range.getLastRow(),startCol=range.getColumn(),endCol=range.getLastColumn(),snapshotHeaders=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],items=[];
  for(var row=startRow;row<=endRow;row++){
    var snapshot=activitySnapshot_(sh,row);
    for(var col=startCol;col<=endCol;col++){
      var isSingle=range.getNumRows()===1&&range.getNumColumns()===1;
      var field=snapshotHeaders[col-1]||('Column '+col),oldValue=isSingle&&e.oldValue!==undefined?String(e.oldValue):'',newValue=String(snapshot.values[col-1]||'');
      var action=activityEditAction_(field,oldValue,newValue,isSingle);
      items.push({user:activityUser_(e),sourceWorkbook:sh.getParent().getName(),workspace:activityWorkspace_(sh.getParent()),tab:sh.getName(),action:action,recordId:activityValue_(snapshot.headers,snapshot.values,['ASSET CODE','CODE','SERIAL NO','SERIAL NUMBER','ITEM NAME','CONSIGNMENT','MAC ADDRESS','EMP ID','EMPLOYEE ID']),empId:activityValue_(snapshot.headers,snapshot.values,['EMP ID','EMPLOYEE ID','EMP CODE','EMPLOYEE CODE']),empName:activityValue_(snapshot.headers,snapshot.values,['EMP NAME','EMPLOYEE NAME','USER NAME','ASSIGNED TO','CHECKED BY','NAME']),field:field,oldValue:oldValue,newValue:newValue,row:row});
    }
  }
  appendActivities_(items);
}
function isActivityProcessAction_(action){return /^(ASSET|CLIENT ASSET|MAC RECORD|COURIER RECORD)/.test(String(action||'').toUpperCase());}
function activityProcessKey_(record){return String(record.recordId||record.empId||record.empName||record.row||'').trim();}
function groupActivityProcessRecords_(records){
  var groups=[];
  records.slice().sort(function(a,b){return a.timestampMs-b.timestampMs;}).forEach(function(record){
    var previous=groups.length?groups[groups.length-1]:null;
    var processKey=activityProcessKey_(record),sameProcess=previous && processKey && isActivityProcessAction_(record.action) && record.action===previous.action && record.sourceWorkbook===previous.sourceWorkbook && record.tab===previous.tab && processKey===previous.processKey && (record.timestampMs-previous.lastTimestampMs)<=15000;
    if(sameProcess){previous.records.push(record);previous.lastTimestampMs=record.timestampMs;}
    else groups.push({records:[record],lastTimestampMs:record.timestampMs,action:record.action,sourceWorkbook:record.sourceWorkbook,tab:record.tab,recordId:record.recordId,processKey:processKey});
  });
  return groups.map(function(group){
    if(group.records.length===1)return group.records[0];
    var first=group.records[0],last=group.records[group.records.length-1],oldParts=[],newParts=[];
    group.records.forEach(function(record){if(record.oldValue)oldParts.push(record.field+': '+record.oldValue);if(record.newValue)newParts.push(record.field+': '+record.newValue);});
    return {timestamp:first.timestamp,timestampMs:first.timestampMs,user:last.user||first.user,sourceWorkbook:first.sourceWorkbook,workspace:first.workspace,tab:first.tab,action:first.action,recordId:first.recordId,empId:last.empId||first.empId,empName:last.empName||first.empName,field:'Process ('+group.records.length+' field changes)',oldValue:oldParts.join(' · ')||'New record',newValue:newParts.join(' · '),row:last.row||first.row};
  });
}
function getActivityLogVersion(){
  requireDashboardAccess_();
  requireDashboardAccess_();
  var sh=getActivityLogSheet_(),lastRow=sh.getLastRow();
  if(lastRow<2)return {lastRow:lastRow,lastTimestampMs:0};
  var v=sh.getRange(lastRow,1).getValue();
  var d=v instanceof Date?v:new Date(v);
  return {lastRow:lastRow,lastTimestampMs:isNaN(d.getTime())?0:d.getTime()};
}

function getAssetTimeline(assetCode){
  requireDashboardAccess_();
  requireDashboardAccess_();
  var code=String(assetCode||'').trim();
  if(!code)return{ok:true,records:[]};
  var sheet=getActivityLogSheet_(),rows=sheet.getDataRange().getValues();
  if(rows.length<2)return{ok:true,records:[]};
  var headers=rows[0].map(function(h){return String(h||'').trim().toUpperCase();});
  var at=function(name,fallback){var i=headers.indexOf(name);return i>-1?i:fallback;};
  var idx={timestamp:at('TIMESTAMP',0),user:at('USER',1),workspace:at('WORKSPACE',3),tab:at('TAB',4),action:at('ACTION',5),recordId:at('RECORD ID',6),empId:at('EMPLOYEE ID',7),empName:at('EMPLOYEE NAME',8),field:at('FIELD',9),oldValue:at('OLD VALUE',10),newValue:at('NEW VALUE',11),row:at('ROW',12)};
  var target=code.toUpperCase(),records=[];
  rows.slice(1).forEach(function(r){
    var recordId=String(r[idx.recordId]||'').trim();
    if(recordId.toUpperCase()!==target)return;
    var d=r[idx.timestamp] instanceof Date?r[idx.timestamp]:new Date(r[idx.timestamp]);
    if(isNaN(d.getTime()))return;
    records.push({
      timestamp:Utilities.formatDate(d,Session.getScriptTimeZone(),'dd-MMM-yyyy hh:mm a'),
      timestampMs:d.getTime(),
      user:String(r[idx.user]||''),workspace:String(r[idx.workspace]||''),tab:String(r[idx.tab]||''),
      action:String(r[idx.action]||''),recordId:recordId,empId:String(r[idx.empId]||''),empName:String(r[idx.empName]||''),
      field:String(r[idx.field]||''),oldValue:String(r[idx.oldValue]||''),newValue:String(r[idx.newValue]||''),row:String(r[idx.row]||'')
    });
  });
  records.sort(function(a,b){return b.timestampMs-a.timestampMs;});
  return{ok:true,assetCode:code,records:records.slice(0,200)};
}

function getActivityLog(filters){
  requireDashboardAccess_();requireDashboardAccess_();
  filters=filters||{};
  var rows=getActivityLogSheet_().getDataRange().getValues();
  if(rows.length<2)return{records:[],total:0,counts:{},updated:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy, hh:mm a')};
  var headers=rows[0].map(function(h){return String(h||'').trim().toUpperCase();}),at=function(name,fallback){var i=headers.indexOf(name);return i>-1?i:fallback;};
  var idx={timestamp:at('TIMESTAMP',0),user:at('USER',1),sourceWorkbook:at('SOURCE WORKBOOK',2),workspace:at('WORKSPACE',3),tab:at('TAB',4),action:at('ACTION',5),recordId:at('RECORD ID',6),empId:at('EMPLOYEE ID',7),empName:at('EMPLOYEE NAME',8),field:at('FIELD',9),oldValue:at('OLD VALUE',10),newValue:at('NEW VALUE',11),row:at('ROW',12)};
  var from=filters.from?new Date(filters.from+'T00:00:00'):null,to=filters.to?new Date(filters.to+'T23:59:59.999'):null,q=String(filters.query||'').toUpperCase(),workspace=String(filters.workspace||''),tab=String(filters.tab||'');
  var records=rows.slice(1).map(function(r){var d=r[idx.timestamp] instanceof Date?r[idx.timestamp]:new Date(r[idx.timestamp]);return{timestamp:Utilities.formatDate(d,Session.getScriptTimeZone(),'dd-MMM-yyyy hh:mm a'),timestampMs:d.getTime(),user:String(r[idx.user]||''),sourceWorkbook:String(r[idx.sourceWorkbook]||''),workspace:String(r[idx.workspace]||''),tab:String(r[idx.tab]||''),action:String(r[idx.action]||''),recordId:String(r[idx.recordId]||''),empId:String(r[idx.empId]||''),empName:String(r[idx.empName]||''),field:String(r[idx.field]||''),oldValue:String(r[idx.oldValue]||''),newValue:String(r[idx.newValue]||''),row:String(r[idx.row]||'')};}).filter(function(r){if(from&&r.timestampMs<from.getTime())return false;if(to&&r.timestampMs>to.getTime())return false;if(workspace&&r.workspace!==workspace)return false;if(tab&&r.tab!==tab)return false;return !q||Object.keys(r).some(function(k){return String(r[k]||'').toUpperCase().indexOf(q)>-1;});});
  records=groupActivityProcessRecords_(records).sort(function(a,b){return b.timestampMs-a.timestampMs;});
  var counts={};records.forEach(function(r){counts[r.action]=(counts[r.action]||0)+1;});
  return{records:records,total:records.length,availableTotal:records.length,counts:counts,updated:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy, hh:mm a')};
}


// ============================================================
// ACTIVITY LOG REPORT DOWNLOAD
// Creates a temporary spreadsheet containing only the selected
// Activity Log period, exports it as a real .xlsx file, and then
// removes the temporary spreadsheet. Existing Activity_Log data
// and logging behavior are not modified.
// ============================================================
function downloadActivityLogReport(filters){
  requireDashboardAccess_();
  requireDashboardAccess_();
  filters=filters||{};
  var mode=String(filters.mode||'daily').toLowerCase();
  var selected=String(filters.value||'').trim();
  if(mode!=='daily' && mode!=='monthly') throw new Error('Invalid report type.');
  if(!selected) throw new Error(mode==='daily'?'Please select a date.':'Please select a month.');

  var tz=Session.getScriptTimeZone();
  var from,to,filenameLabel;
  if(mode==='daily'){
    if(!/^\d{4}-\d{2}-\d{2}$/.test(selected)) throw new Error('Please select a valid date.');
    from=new Date(selected+'T00:00:00');
    to=new Date(selected+'T23:59:59.999');
    filenameLabel=selected;
  }else{
    if(!/^\d{4}-\d{2}$/.test(selected)) throw new Error('Please select a valid month.');
    var parts=selected.split('-'),year=Number(parts[0]),month=Number(parts[1])-1;
    from=new Date(year,month,1,0,0,0,0);
    to=new Date(year,month+1,0,23,59,59,999);
    filenameLabel=selected;
  }

  var sh=getActivityLogSheet_();
  var values=sh.getDataRange().getValues();
  if(values.length<2) throw new Error('No Activity Logs are available for the selected period.');

  var headers=values[0].map(function(h){return String(h||'').trim().toUpperCase();});
  function col(name,fallback){var i=headers.indexOf(name);return i>-1?i:fallback;}
  var idx={timestamp:col('TIMESTAMP',0),user:col('USER',1),sourceWorkbook:col('SOURCE WORKBOOK',2),workspace:col('WORKSPACE',3),tab:col('TAB',4),action:col('ACTION',5),recordId:col('RECORD ID',6),empId:col('EMPLOYEE ID',7),empName:col('EMPLOYEE NAME',8),field:col('FIELD',9),oldValue:col('OLD VALUE',10),newValue:col('NEW VALUE',11),row:col('ROW',12)};

  var records=values.slice(1).map(function(r){
    var raw=r[idx.timestamp],d=raw instanceof Date?raw:new Date(raw);
    return {
      timestamp:d,
      timestampMs:d.getTime(),
      user:String(r[idx.user]||''),
      sourceWorkbook:String(r[idx.sourceWorkbook]||''),
      workspace:String(r[idx.workspace]||''),
      tab:String(r[idx.tab]||''),
      action:String(r[idx.action]||''),
      recordId:String(r[idx.recordId]||''),
      empId:String(r[idx.empId]||''),
      empName:String(r[idx.empName]||''),
      field:String(r[idx.field]||''),
      oldValue:String(r[idx.oldValue]||''),
      newValue:String(r[idx.newValue]||''),
      row:String(r[idx.row]||'')
    };
  }).filter(function(r){
    return !isNaN(r.timestampMs) && r.timestampMs>=from.getTime() && r.timestampMs<=to.getTime();
  });

  records=groupActivityProcessRecords_(records).sort(function(a,b){return b.timestampMs-a.timestampMs;});
  if(!records.length) throw new Error('No Activity Logs found for '+(mode==='daily'?'the selected date.':'the selected month.'));

  var temp=null;
  try{
    temp=SpreadsheetApp.create('Activity Log Report - '+filenameLabel);
    var out=temp.getSheets()[0];
    out.setName('Activity_Log_Report');
    var exportHeaders=['Timestamp','User','Source Workbook','Workspace','Tab','Action','Record ID','Employee ID','Employee Name','Field','Old Value','New Value','Row'];
    var outputRows=records.map(function(r){
      return [r.timestamp,r.user,r.sourceWorkbook,r.workspace,r.tab,r.action,r.recordId,r.empId,r.empName,r.field,r.oldValue,r.newValue,r.row];
    });
    out.getRange(1,1,1,exportHeaders.length).setValues([exportHeaders]);
    out.getRange(2,1,outputRows.length,exportHeaders.length).setValues(outputRows);
    out.getRange(2,1,outputRows.length,1).setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');
    // Make the Timestamp column wide enough for Excel to display the full date/time.
    out.setColumnWidth(1, 190);
    out.setFrozenRows(1);
    out.getRange(1,1,1,exportHeaders.length).setFontWeight('bold');
    out.autoResizeColumns(1,exportHeaders.length);
    SpreadsheetApp.flush();

    var exportUrl='https://docs.google.com/spreadsheets/d/'+temp.getId()+'/export?format=xlsx';
    var response=UrlFetchApp.fetch(exportUrl,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
    if(response.getResponseCode()!==200) throw new Error('Could not create the Excel report. Please try again.');
    var blob=response.getBlob().setName('Activity_Logs_'+mode+'_'+filenameLabel+'.xlsx');
    return {
      ok:true,
      filename:blob.getName(),
      mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64:Utilities.base64Encode(blob.getBytes()),
      count:records.length
    };
  }finally{
    if(temp){try{DriveApp.getFileById(temp.getId()).setTrashed(true);}catch(ignore){}}
  }
}

function findSheetLoose_(ss,sheetName){var direct=ss.getSheetByName(sheetName);if(direct)return direct;var target=String(sheetName||"").trim().toLowerCase(),all=ss.getSheets();for(var i=0;i<all.length;i++)if(all[i].getName().trim().toLowerCase()===target)return all[i];return null;}
function readSheetData_(sheetName,colMap,rowProcessor,spreadsheetId){requireDashboardAccess_();try{var cache=CacheService.getScriptCache(),cacheKey="asset:v2:"+(spreadsheetId||"active")+":"+sheetName,cached=cache.get(cacheKey);if(cached){try{return JSON.parse(cached);}catch(e){cache.remove(cacheKey);}}var ss;if(spreadsheetId){if(spreadsheetId.indexOf("PASTE_")===0)return{error:"Set MAC_SPREADSHEET_ID in Code.gs to the other workbook's Spreadsheet ID first (see the comment above it)."};ss=SpreadsheetApp.openById(spreadsheetId);}else ss=getMainSpreadsheet_();var sheet=findSheetLoose_(ss,sheetName);if(!sheet){var availableNames=ss.getSheets().map(function(sh){return sh.getName();}).join(", ");return{error:"Sheet '"+sheetName+"' not found. Tabs available in that workbook: "+availableNames};}var lastRow=sheet.getLastRow(),lastCol=sheet.getLastColumn();if(lastRow<2)return{error:"No data in '"+sheetName+"'."};if(lastRow>5000)lastRow=5000;if(lastCol>30)lastCol=30;var headerRow=sheet.getRange(1,1,1,lastCol).getValues()[0].map(function(h){return String(h||"").trim().toUpperCase();});function findCol(patterns){for(var i=0;i<headerRow.length;i++)for(var p=0;p<patterns.length;p++)if(headerRow[i].indexOf(patterns[p])>-1)return i;return -1;}var detected={};Object.keys(colMap).forEach(function(key){if(Array.isArray(colMap[key])){var found=findCol(colMap[key]);if(found>-1)detected[key]=found;}});var COL={};Object.keys(colMap).forEach(function(key){COL[key]=detected[key]!==undefined?detected[key]:colMap[key][1];});var values=sheet.getRange(2,1,lastRow-1,lastCol).getValues(),counts={},records=[];values.forEach(function(row,idx){var result=rowProcessor(row,idx,COL);if(result){if(result.skip)return;var key=result.status||"UNKNOWN";counts[key]=(counts[key]||0)+1;records.push(result.record);}});var result={total:records.length,counts:counts,records:records,updated:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy, hh:mm a")};try{var packed=JSON.stringify(result);if(packed.length<95000)cache.put(cacheKey,packed,10);}catch(ignore){}return result;}catch(e){return{error:e.message};}}
function formatDate_(val){if(Object.prototype.toString.call(val)==="[object Date]")return Utilities.formatDate(val,Session.getScriptTimeZone(),"dd-MMM-yyyy");return String(val);}
function isMacLike_(val){return/^([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}$/.test(String(val||"").trim());}
function isPcAssetCode_(code){return/^\d{4}\s*PC\s*\d+$/i.test(String(code||"").trim());}

function getDashboardData(){
  requireDashboardAccess_();return readSheetData_(SHEET_NAME_OFFICE,{code:[["CODE","ASSET CODE"],0],status:[["STATUS","ASSET STATUS"],1],empId:[["EMP ID","EMPLOYEE ID"],3],empName:[["EMP NAME","EMPLOYEE NAME"],4],date:[["DATE","ASSIGNED DATE"],5],remarks:[["REMARKS"],7]},function(row,idx,COL){var code=row[COL.code];if(!code)return{skip:true};var status=String(row[COL.status]||"UNKNOWN").trim().toUpperCase();return{status:status,record:{code:String(code),status:status,category:isPcAssetCode_(code)?"PC / Server":"Laptop",empId:row[COL.empId]?String(row[COL.empId]):"",empName:row[COL.empName]?String(row[COL.empName]):"",date:row[COL.date]?formatDate_(row[COL.date]):"",remarks:row[COL.remarks]?String(row[COL.remarks]):""}};});}

function getClientLaptopsData(){
  requireDashboardAccess_();return readSheetData_(SHEET_NAME_CLIENT,{clientStatus:[["CLIENT STATUS","STATUS"],0],clientName:[["CLIENT NAME"],1],clientCode:[["CLIENT CODE"],2],project:[["PROJECT"],3],assignDate:[["ASSIGN DATE","DATE"],4],userName:[["USER NAME","EMP NAME"],5],empId:[["EMP ID"],6],cyntexaLoc:[["LOCATION","CYNTExA LOC"],7],assetCode:[["ASSET CODE"],8],status:[["STATUS"],9],assetType:[["ASSET TYPE"],10],macAddress:[["MAC ADDRESS"],11],makeModel:[["MAKE MODEL","MODEL"],12],serialNumber:[["SERIAL NUMBER","SERIAL NO"],13],processor:[["PROCESSOR"],14],ram:[["RAM"],15],ssd:[["SSD"],16],itemsReceived:[["ITEMS RECEIVED"],17],otherAssetCode:[["OTHER ASSET CODE"],18],laptopSubmitDate:[["SUBMIT DATE"],19],dispatchedDate:[["DISPATCHED DATE"],20],courierName:[["COURIER NAME"],21],dispatchedToPerson:[["DISPATCHED TO"],22],dispatchedItems:[["DISPATCHED ITEMS"],23],remarks:[["REMARKS"],-1]},function(row,idx,COL){var code=row[COL.assetCode];if(!code)return{skip:true};var badgeStatus=String(row[COL.clientStatus]||"UNKNOWN").trim().toUpperCase();if(!badgeStatus)badgeStatus="UNKNOWN";return{status:badgeStatus,record:{code:String(code),status:badgeStatus,location:String(row[COL.cyntexaLoc]||"").trim(),clientName:row[COL.clientName]?String(row[COL.clientName]):"",clientCode:row[COL.clientCode]?String(row[COL.clientCode]):"",project:row[COL.project]?String(row[COL.project]):"",userName:row[COL.userName]?String(row[COL.userName]):"",empId:row[COL.empId]?String(row[COL.empId]):"",cyntexaLoc:String(row[COL.cyntexaLoc]||"").trim(),assignDate:row[COL.assignDate]?formatDate_(row[COL.assignDate]):"",assetType:row[COL.assetType]?String(row[COL.assetType]):"",macAddress:row[COL.macAddress]?String(row[COL.macAddress]):"",makeModel:row[COL.makeModel]?String(row[COL.makeModel]):"",serialNumber:row[COL.serialNumber]?String(row[COL.serialNumber]):"",processor:row[COL.processor]?String(row[COL.processor]):"",ram:row[COL.ram]?String(row[COL.ram]):"",ssd:row[COL.ssd]?String(row[COL.ssd]):"",itemsReceived:row[COL.itemsReceived]?String(row[COL.itemsReceived]):"",otherAssetCode:row[COL.otherAssetCode]?String(row[COL.otherAssetCode]):"",laptopSubmitDate:row[COL.laptopSubmitDate]?formatDate_(row[COL.laptopSubmitDate]):"",dispatchedDate:row[COL.dispatchedDate]?formatDate_(row[COL.dispatchedDate]):"",courierName:row[COL.courierName]?String(row[COL.courierName]):"",dispatchedToPerson:row[COL.dispatchedToPerson]?String(row[COL.dispatchedToPerson]):"",dispatchedItems:row[COL.dispatchedItems]?String(row[COL.dispatchedItems]):"",remarks:row[COL.remarks]>-1?String(row[COL.remarks]):""}};});}

var OTHER_ASSET_CATEGORY_RULES_=[{label:"TV",keywords:["MI-TV","LEDTV","LED-TV"," TV-"," TV "]},{label:"Monitor",keywords:["MONITOR","SCREEN"]},{label:"UPS / Inverter",keywords:["UPS","INVERTOR","INVERTER"]},{label:"Printer",keywords:["PRINTER","SPLITTER"]},{label:"Writing Pad",keywords:["WRITING PAD","PEN TABLET","CPTABLET"]},{label:"Wheel Stand",keywords:["WHEEL STAND","TROLLEY"]},{label:"Mic",keywords:["MIC","MICROPHONE","RODE"]},{label:"Camera & Lens",keywords:["CAMERA","LENS","GIMBAL","CAM-LINK","CAM LINK"]},{label:"Webcam",keywords:["WEBCAM","WEB CAM","LWC0"]},{label:"Face Terminal",keywords:["FACE TERMINAL","FACETERMINAL","HIK-FT"]},{label:"WiFi Access Point",keywords:["ACCESS POINT","UAP-","U6-LR","U6-PRO","UBNT-AP"]},{label:"WiFi Router",keywords:["WIFI ROUTER","WI-FI ROUTER","ROUTER"]},{label:"Network Switch",keywords:["SWITCH","PSWITCH","POE"]},{label:"Firewall",keywords:["FIREWALL","FORTIGATE","UDM"]},{label:"NVR",keywords:["NVR"]},{label:"NAS",keywords:["NAS ","DISKSTATION"]},{label:"Speaker",keywords:["SPEAKER"]},{label:"Gaming",keywords:["PS4","PS5","GAME SET","GAME"]},{label:"Coffee Machine",keywords:["COFFEE MACHINE"]},{label:"Keyboard & Mouse",keywords:["KEYBOARD","MOUSE","-KM0"]},{label:"Dock / Stand",keywords:["DOCK","LAPTOP STAND"]},{label:"Headphones",keywords:["HEADPHONE","H390"]},{label:"Cables & Accessories",keywords:["CABLE","USB HUB","USB-C","HDMI"]}];
function normalizeOtherAssetCategory_(rawCategory,assetName,assetCode){var text=(String(rawCategory||"")+" "+String(assetName||"")+" "+String(assetCode||"")).toUpperCase();for(var i=0;i<OTHER_ASSET_CATEGORY_RULES_.length;i++){var rule=OTHER_ASSET_CATEGORY_RULES_[i];for(var j=0;j<rule.keywords.length;j++)if(text.indexOf(rule.keywords[j])>-1)return rule.label;}var base=String(assetName||assetCode||"").trim();base=base.replace(/['’]s?\s*$/i,"").replace(/[\s\-_#]*\d+\s*$/,"").trim();return base?base:"Other";}
function getOtherAssetsData(){
  requireDashboardAccess_();return readSheetData_(SHEET_NAME_OTHER,{assetName:[["ASSET NAME"],1],status:[["ASSET STATUS","STATUS"],2],assetCode:[["ASSET CODE"],3],category:[["CATEGORY"],-1],configuration:[["CONFIGURATION"],4],empId:[["EMPLOYEE ID","EMP ID","EMP.ID"],5],empName:[["EMPLOYEE NAME","EMP NAME"],6],assignedDate:[["ASSIGNED DATE"],7],serialNo:[["SERIAL NO","SERIAL NUMBER"],8],onCRM:[["ON CRM"],9],purchaseDate:[["PURCHASE DATE"],10],physicallyVerified:[["PHYSICALLY VERIFIED","VERIFIED"],11],assetCondition:[["ASSET CONDITION","CONDITION"],12],location:[["LOCATION"],13],remarks:[["REMARKS"],-1]},function(row,idx,COL){var code=row[COL.assetCode];if(!code)return{skip:true};var status=String(row[COL.status]||"UNKNOWN").trim().toUpperCase(),assetName=row[COL.assetName]?String(row[COL.assetName]):"",rawCategory=COL.category>-1&&row[COL.category]?String(row[COL.category]):"";return{status:status,record:{code:String(code),status:status,category:normalizeOtherAssetCategory_(rawCategory,assetName,code),assetName:assetName,configuration:row[COL.configuration]?String(row[COL.configuration]):"",empId:row[COL.empId]?String(row[COL.empId]):"",empName:row[COL.empName]?String(row[COL.empName]):"",assignedDate:row[COL.assignedDate]?formatDate_(row[COL.assignedDate]):"",serialNo:row[COL.serialNo]?String(row[COL.serialNo]):"",onCRM:row[COL.onCRM]?String(row[COL.onCRM]):"",purchaseDate:row[COL.purchaseDate]?formatDate_(row[COL.purchaseDate]):"",physicallyVerified:row[COL.physicallyVerified]?String(row[COL.physicallyVerified]):"",assetCondition:row[COL.assetCondition]?String(row[COL.assetCondition]):"",location:row[COL.location]?String(row[COL.location]):"",remarks:COL.remarks>-1?String(row[COL.remarks]):""}};});}

// ============================================================
// OTHER ASSETS (IT) — ADD / EDIT
// All fields in Other_Assets_IT are editable from the dashboard.
// ============================================================
function otherAssetColumns_(sheet){
  var lastCol=sheet.getLastColumn();
  if(lastCol<1) throw new Error("Sheet '"+SHEET_NAME_OTHER+"' needs a header row.");
  var headers=sheet.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(v){return String(v||'').trim().toUpperCase();});
  function col(patterns,fallback){
    for(var p=0;p<patterns.length;p++) for(var h=0;h<headers.length;h++) if(headers[h]===patterns[p]) return h+1;
    for(var p2=0;p2<patterns.length;p2++) for(var h2=0;h2<headers.length;h2++) if(headers[h2].indexOf(patterns[p2])>-1) return h2+1;
    return fallback;
  }
  return {
    lastCol:lastCol,
    assetName:col(['ASSET NAME'],2),
    status:col(['ASSET STATUS','STATUS'],3),
    assetCode:col(['ASSET CODE'],4),
    configuration:col(['CONFIGURATION'],5),
    empId:col(['EMPLOYEE ID','EMP ID','EMP.ID'],6),
    empName:col(['EMPLOYEE NAME','EMP NAME'],7),
    assignedDate:col(['ASSIGNED DATE'],8),
    serialNo:col(['SERIAL NO','SERIAL NUMBER'],9),
    onCRM:col(['ON CRM'],10),
    purchaseDate:col(['PURCHASE DATE'],11),
    physicallyVerified:col(['PHYSICALLY VERIFIED','VERIFIED'],12),
    assetCondition:col(['ASSET CONDITION','CONDITION'],13),
    location:col(['LOCATION'],14),
    remarks:col(['REMARKS'],Math.max(lastCol,15))
  };
}
function otherAssetDate_(value,label){
  var text=String(value||'').trim();
  if(!text) return '';
  var d=parseOfficeAssignmentDate_(text);
  if(!d) throw new Error('Please select a valid '+label+'.');
  return d;
}
function clearOtherAssetCaches_(){
  var c=CacheService.getScriptCache();
  c.removeAll(['asset:v2:active:'+SHEET_NAME_OTHER,'asset:v3:active:'+SHEET_NAME_OTHER,'asset:v3:active:'+SHEET_NAME_OTHER+':gz']);
}
function otherAssetPayloadValue_(payload,key){return String(payload[key]===undefined||payload[key]===null?'':payload[key]).trim();}
function addOtherAsset(payload){
  requireDashboardAccess_();
  payload=payload||{};
  var assetCode=otherAssetPayloadValue_(payload,'assetCode');
  if(!assetCode) throw new Error('Asset Code is required.');
  var lock=LockService.getDocumentLock(); lock.waitLock(30000);
  try{
    var sheet=findSheetLoose_(getMainSpreadsheet_(),SHEET_NAME_OTHER);
    if(!sheet) throw new Error("Sheet '"+SHEET_NAME_OTHER+"' was not found.");
    var cols=otherAssetColumns_(sheet), lastRow=sheet.getLastRow();
    if(lastRow>=2){
      var codes=sheet.getRange(2,cols.assetCode,lastRow-1,1).getDisplayValues();
      for(var i=0;i<codes.length;i++) if(String(codes[i][0]||'').trim().toUpperCase()===assetCode.toUpperCase()) throw new Error('Asset Code "'+assetCode+'" already exists in '+SHEET_NAME_OTHER+'.');
    }
    var values={
      assetName:otherAssetPayloadValue_(payload,'assetName'), status:otherAssetPayloadValue_(payload,'status')||'ACTIVE', assetCode:assetCode,
      configuration:otherAssetPayloadValue_(payload,'configuration'), empId:otherAssetPayloadValue_(payload,'empId'), empName:otherAssetPayloadValue_(payload,'empName'),
      assignedDate:otherAssetDate_(payload.assignedDate,'Assigned Date'), serialNo:otherAssetPayloadValue_(payload,'serialNo'), onCRM:otherAssetPayloadValue_(payload,'onCRM'),
      purchaseDate:otherAssetDate_(payload.purchaseDate,'Purchase Date'), physicallyVerified:otherAssetPayloadValue_(payload,'physicallyVerified'),
      assetCondition:otherAssetPayloadValue_(payload,'assetCondition'), location:otherAssetPayloadValue_(payload,'location'), remarks:otherAssetPayloadValue_(payload,'remarks')
    };
    var width=cols.lastCol; Object.keys(cols).forEach(function(k){if(k!=='lastCol') width=Math.max(width,cols[k]);});
    var row=[]; for(var c=0;c<width;c++) row.push('');
    Object.keys(values).forEach(function(k){row[cols[k]-1]=values[k];});
    var newRow=lastRow+1; sheet.getRange(newRow,1,1,width).setValues([row]);
    [cols.assignedDate,cols.purchaseDate].forEach(function(c){if(sheet.getRange(newRow,c).getValue() instanceof Date) sheet.getRange(newRow,c).setNumberFormat('dd-mmm-yyyy');});
    auditDashboardRow_(sheet,newRow,'ASSET CREATED');
    clearOtherAssetCaches_();
    return {ok:true,assetCode:assetCode,status:values.status};
  } finally {lock.releaseLock();}
}
function updateOtherAsset(payload){
  requireDashboardAccess_();
  payload=payload||{};
  var originalCode=otherAssetPayloadValue_(payload,'originalAssetCode')||otherAssetPayloadValue_(payload,'assetCode');
  var assetCode=otherAssetPayloadValue_(payload,'assetCode');
  if(!originalCode) throw new Error('Original Asset Code is required.');
  if(!assetCode) throw new Error('Asset Code is required.');
  var lock=LockService.getDocumentLock(); lock.waitLock(30000);
  try{
    var sheet=findSheetLoose_(getMainSpreadsheet_(),SHEET_NAME_OTHER);
    if(!sheet) throw new Error("Sheet '"+SHEET_NAME_OTHER+"' was not found.");
    var cols=otherAssetColumns_(sheet), lastRow=sheet.getLastRow(); if(lastRow<2) throw new Error("Sheet '"+SHEET_NAME_OTHER+"' has no asset rows.");
    var codes=sheet.getRange(2,cols.assetCode,lastRow-1,1).getDisplayValues(), target=0;
    for(var i=0;i<codes.length;i++) if(String(codes[i][0]||'').trim().toUpperCase()===originalCode.toUpperCase()){target=i+2;break;}
    if(!target) throw new Error('Asset Code "'+originalCode+'" was not found in '+SHEET_NAME_OTHER+'.');
    for(var j=0;j<codes.length;j++) if(j+2!==target && String(codes[j][0]||'').trim().toUpperCase()===assetCode.toUpperCase()) throw new Error('Asset Code "'+assetCode+'" already exists in '+SHEET_NAME_OTHER+'.');
    var before=activitySnapshot_(sheet,target);
    var values={
      assetName:otherAssetPayloadValue_(payload,'assetName'), status:otherAssetPayloadValue_(payload,'status')||'ACTIVE', assetCode:assetCode,
      configuration:otherAssetPayloadValue_(payload,'configuration'), empId:otherAssetPayloadValue_(payload,'empId'), empName:otherAssetPayloadValue_(payload,'empName'),
      assignedDate:otherAssetDate_(payload.assignedDate,'Assigned Date'), serialNo:otherAssetPayloadValue_(payload,'serialNo'), onCRM:otherAssetPayloadValue_(payload,'onCRM'),
      purchaseDate:otherAssetDate_(payload.purchaseDate,'Purchase Date'), physicallyVerified:otherAssetPayloadValue_(payload,'physicallyVerified'),
      assetCondition:otherAssetPayloadValue_(payload,'assetCondition'), location:otherAssetPayloadValue_(payload,'location'), remarks:otherAssetPayloadValue_(payload,'remarks')
    };
    Object.keys(values).forEach(function(k){
      var v=values[k]; if(k==='assignedDate'||k==='purchaseDate'){if(v) sheet.getRange(target,cols[k]).setValue(v).setNumberFormat('dd-mmm-yyyy'); else sheet.getRange(target,cols[k]).clearContent();}
      else sheet.getRange(target,cols[k]).setValue(v);
    });
    auditDashboardRow_(sheet,target,'ASSET UPDATED',before);
    clearOtherAssetCaches_();
    return {ok:true,assetCode:assetCode,status:values.status};
  } finally {lock.releaseLock();}
}


var MOBILE_CATEGORY_RULES_=[{label:"Tablet / iPad",keywords:["TABLET","I-PAD","IPAD"]},{label:"Mobile Phone",keywords:["MOBILE","ONEPLUS","REDMI","REALME","REAL-","MI-A2","IPHONE","I-PHONE","I PHONE"]},{label:"Buds / Earbuds",keywords:["BUDS","EARBUDS"]},{label:"Headphones / Earphones",keywords:["HEADPHONE","EARPHONE","H/E-"]},{label:"Landline Phone",keywords:["LANDLINE"]}];
function normalizeMobileCategory_(assetTypeRaw,configuration,assetCode){var text=(String(assetTypeRaw||"")+" "+String(configuration||"")+" "+String(assetCode||"")).toUpperCase();for(var i=0;i<MOBILE_CATEGORY_RULES_.length;i++){var rule=MOBILE_CATEGORY_RULES_[i];for(var j=0;j<rule.keywords.length;j++)if(text.indexOf(rule.keywords[j])>-1)return rule.label;}var base=String(assetTypeRaw||"").trim().replace(/['’]s?\s*$/i,"").replace(/[\s\-_#]*\d+\s*$/,"").trim();return base?base:"Other";}

function mobileAssetColumns_(sheet){
  var lastCol=sheet.getLastColumn();
  if(lastCol<1) throw new Error("Sheet '"+SHEET_NAME_MOBILE+"' needs a header row.");
  var headers=sheet.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(v){return String(v||'').trim().toUpperCase();});
  function col(patterns,fallback){for(var i=0;i<patterns.length;i++)for(var j=0;j<headers.length;j++)if(headers[j]===patterns[i])return j+1;return fallback;}
  return {lastCol:lastCol,assetType:col(['ASSET NAME','ASSET TYPE'],1),status:col(['ASSET STATUS','STATUS'],2),assetCode:col(['ASSET CODE','CODE'],3),configuration:col(['CONFIGURATION','MODEL','MODEL / CONFIGURATION'],4),empId:col(['EMPLOYEE ID','EMP ID','EMP.ID'],5),empName:col(['EMPLOYEE NAME','EMP NAME'],6),assignedDate:col(['ASSIGNED DATE'],7),serialNo:col(['SERIAL NO','SERIAL NUMBER'],8),imeiNo:col(['IMEI','IMEI NO','IMEI NUMBER'],9),onCRM:col(['ON CRM'],10),remarks:col(['REMARKS'],-1)};
}
function clearMobileAssetCaches_(){CacheService.getScriptCache().removeAll(['asset:v2:active:'+SHEET_NAME_MOBILE,'asset:v3:active:'+SHEET_NAME_MOBILE,'asset:v3:active:'+SHEET_NAME_MOBILE+':gz']);}
function mobileAssetValue_(payload,key){return String(payload && payload[key]!==undefined && payload[key]!==null ? payload[key] : '').trim();}
function mobileAssetDate_(value,label){var text=String(value||'').trim();if(!text)return '';var d=parseOfficeAssignmentDate_(text);if(!d)throw new Error('Please select a valid '+label+'.');return d;}
function addMobileAccessory(payload){
  requireDashboardAccess_();
  payload=payload||{}; var assetCode=mobileAssetValue_(payload,'assetCode'); if(!assetCode)throw new Error('Asset Code is required.');
  var lock=LockService.getDocumentLock();lock.waitLock(30000);try{
    var sheet=findSheetLoose_(getMainSpreadsheet_(),SHEET_NAME_MOBILE);if(!sheet)throw new Error("Sheet '"+SHEET_NAME_MOBILE+"' was not found.");
    var cols=mobileAssetColumns_(sheet),lastRow=sheet.getLastRow();
    if(lastRow>=2){var codes=sheet.getRange(2,cols.assetCode,lastRow-1,1).getDisplayValues();for(var i=0;i<codes.length;i++)if(String(codes[i][0]||'').trim().toUpperCase()===assetCode.toUpperCase())throw new Error('Asset Code "'+assetCode+'" already exists in '+SHEET_NAME_MOBILE+'.');}
    var values={assetType:mobileAssetValue_(payload,'assetType'),status:mobileAssetValue_(payload,'status')||'ACTIVE',assetCode:assetCode,configuration:mobileAssetValue_(payload,'configuration'),empId:mobileAssetValue_(payload,'empId'),empName:mobileAssetValue_(payload,'empName'),assignedDate:mobileAssetDate_(payload.assignedDate,'Assigned Date'),serialNo:mobileAssetValue_(payload,'serialNo'),imeiNo:mobileAssetValue_(payload,'imeiNo'),onCRM:mobileAssetValue_(payload,'onCRM'),remarks:mobileAssetValue_(payload,'remarks')};
    var width=cols.lastCol;Object.keys(cols).forEach(function(k){if(k!=='lastCol'&&cols[k]>width)width=cols[k];});var row=[];for(var c=0;c<width;c++)row.push('');Object.keys(values).forEach(function(k){if(cols[k]>0)row[cols[k]-1]=values[k];});
    var nr=lastRow+1;sheet.getRange(nr,1,1,width).setValues([row]);if(values.assignedDate)sheet.getRange(nr,cols.assignedDate).setNumberFormat('dd-mmm-yyyy');auditDashboardRow_(sheet,nr,'ASSET CREATED');clearMobileAssetCaches_();return {ok:true,assetCode:assetCode,status:values.status};
  }finally{lock.releaseLock();}
}
function updateMobileAccessory(payload){
  requireDashboardAccess_();
  payload=payload||{};var originalCode=mobileAssetValue_(payload,'originalAssetCode')||mobileAssetValue_(payload,'assetCode'),assetCode=mobileAssetValue_(payload,'assetCode');if(!originalCode)throw new Error('Original Asset Code is required.');if(!assetCode)throw new Error('Asset Code is required.');
  var lock=LockService.getDocumentLock();lock.waitLock(30000);try{
    var sheet=findSheetLoose_(getMainSpreadsheet_(),SHEET_NAME_MOBILE);if(!sheet)throw new Error("Sheet '"+SHEET_NAME_MOBILE+"' was not found.");var cols=mobileAssetColumns_(sheet),lastRow=sheet.getLastRow();if(lastRow<2)throw new Error("Sheet '"+SHEET_NAME_MOBILE+"' has no asset rows.");
    var codes=sheet.getRange(2,cols.assetCode,lastRow-1,1).getDisplayValues(),target=0;for(var i=0;i<codes.length;i++)if(String(codes[i][0]||'').trim().toUpperCase()===originalCode.toUpperCase()){target=i+2;break;}if(!target)throw new Error('Asset Code "'+originalCode+'" was not found in '+SHEET_NAME_MOBILE+'.');
    for(var j=0;j<codes.length;j++)if(j+2!==target&&String(codes[j][0]||'').trim().toUpperCase()===assetCode.toUpperCase())throw new Error('Asset Code "'+assetCode+'" already exists in '+SHEET_NAME_MOBILE+'.');
    var before=activitySnapshot_(sheet,target);
    var values={assetType:mobileAssetValue_(payload,'assetType'),status:mobileAssetValue_(payload,'status')||'ACTIVE',assetCode:assetCode,configuration:mobileAssetValue_(payload,'configuration'),empId:mobileAssetValue_(payload,'empId'),empName:mobileAssetValue_(payload,'empName'),assignedDate:mobileAssetDate_(payload.assignedDate,'Assigned Date'),serialNo:mobileAssetValue_(payload,'serialNo'),imeiNo:mobileAssetValue_(payload,'imeiNo'),onCRM:mobileAssetValue_(payload,'onCRM'),remarks:mobileAssetValue_(payload,'remarks')};
    Object.keys(values).forEach(function(k){if(cols[k]<1)return;var v=values[k];if(k==='assignedDate'){if(v)sheet.getRange(target,cols[k]).setValue(v).setNumberFormat('dd-mmm-yyyy');else sheet.getRange(target,cols[k]).clearContent();}else sheet.getRange(target,cols[k]).setValue(v);});
    auditDashboardRow_(sheet,target,'ASSET UPDATED',before);clearMobileAssetCaches_();return {ok:true,assetCode:assetCode,status:values.status};
  }finally{lock.releaseLock();}
}
function getMobileAccessoriesData(){
  requireDashboardAccess_();return readSheetData_(SHEET_NAME_MOBILE,{assetType:[["ASSET NAME"],0],status:[["ASSET STATUS","STATUS"],1],assetCode:[["ASSET CODE"],2],configuration:[["CONFIGURATION"],3],empId:[["EMPLOYEE ID","EMP ID","EMP.ID"],4],empName:[["EMPLOYEE NAME","EMP NAME"],5],assignedDate:[["ASSIGNED DATE"],6],serialNo:[["SERIAL NO","SERIAL NUMBER"],7],imeiNo:[["IMEI"],8],onCRM:[["ON CRM"],9],remarks:[["REMARKS"],-1]},function(row,idx,COL){var code=row[COL.assetCode];if(!code)return{skip:true};var status=String(row[COL.status]||"UNKNOWN").trim().toUpperCase(),assetTypeRaw=row[COL.assetType]?String(row[COL.assetType]):"",configuration=row[COL.configuration]?String(row[COL.configuration]):"";return{status:status,record:{code:String(code),status:status,category:normalizeMobileCategory_(assetTypeRaw,configuration,code),assetType:assetTypeRaw,configuration:configuration,empId:row[COL.empId]?String(row[COL.empId]):"",empName:row[COL.empName]?String(row[COL.empName]):"",assignedDate:row[COL.assignedDate]?formatDate_(row[COL.assignedDate]):"",serialNo:row[COL.serialNo]?String(row[COL.serialNo]):"",imeiNo:row[COL.imeiNo]?String(row[COL.imeiNo]):"",onCRM:row[COL.onCRM]?String(row[COL.onCRM]):"",remarks:COL.remarks>-1?String(row[COL.remarks]):""}};});}

var STUDIO_CATEGORY_RULES_=[{label:"Camera & Lens",keywords:["CAMERA","LENS","CAM CORDER","CAMCORDER","CINEMA LINE","GO PRO","GOPRO"]},{label:"Microphone",keywords:["MIC","MICROPHONE","RODE","PODMIC","LAPLE","LAVALIER","CLOUD LIFTER","CASTER"]},{label:"Light",keywords:["LIGHT","LED","TUBELIGHT","RING LIGHT","SL60","SL 60","SL-60"]},{label:"Soft Box / Reflector",keywords:["SOFT BOX","SOFTBOX","REFLECTOR","GREEN SCREEN"]},{label:"Stand / Tripod",keywords:["STAND","TRIPOD","TRIPODE","BOOM","CLAMP","ZIGZAG"]},{label:"Gimbal / Stabilizer",keywords:["MOZA","GIMBAL","STABILIZER"]},{label:"Speaker",keywords:["SPEAKER","JBL"]},{label:"Cable & Adapter",keywords:["CABLE","HDMI","SPLITTER","CAPTURE"]},{label:"Memory Card",keywords:["MEMORY CARD","SD CARD"]},{label:"Battery & Charger",keywords:["BATTERY","CHARGER"]},{label:"Curtain / Backdrop",keywords:["CURTON","CURTAIN","BACKDROP"]},{label:"Teleprompter",keywords:["TELE PROMTER","TELEPROMPTER","TELE PROMPTER"]},{label:"Tablet / iPad",keywords:["IPAD","I-PAD","TABLET"]},{label:"Keyboard & Mouse",keywords:["MICE","MOUSE","KEYBOARD"]}];
function normalizeStudioCategory_(itemName){var text=String(itemName||"").toUpperCase();for(var i=0;i<STUDIO_CATEGORY_RULES_.length;i++){var rule=STUDIO_CATEGORY_RULES_[i];for(var j=0;j<rule.keywords.length;j++)if(text.indexOf(rule.keywords[j])>-1)return rule.label;}return"Other";}
function normalizeStudioStatus_(rawStatus){var s=String(rawStatus||"").trim().toUpperCase();if(!s)return"UNKNOWN";if(s.indexOf("NEW")>-1||s.indexOf("PURCHASE")>-1)return"NEW / PURCHASED";if(s.indexOf("NOT WORKING")>-1||s.indexOf("NOT PROPERLY")>-1)return"NOT WORKING";if(s.indexOf("DAMAGE")>-1)return"DAMAGED";if(s.indexOf("FAULTY")>-1)return"FAULTY";if(s==="OK"||s.indexOf("WORKING")>-1)return"WORKING";return s;}
function getStudioItemsData(){
  requireDashboardAccess_();return readSheetData_(SHEET_NAME_STUDIO,{name:[["NAME"],0],qty:[["QTY","QUANTITY"],1],status:[["STATUS"],2],purchaseDate:[["PURCHASE DATE"],3],purchaseFrom:[["PURCHASE FROM"],4],modelNumber:[["MODEL NUMBER","MODEL NO"],5],serialNumber:[["SERIAL NUMBER","SERIAL NO"],6],checkedBy:[["CHECKED BY"],7],checkedDate:[["ON DATE","CHECKED DATE"],8],location:[["LOCATION"],9],remarks:[["REMARKS"],-1]},function(row,idx,COL){var name=row[COL.name]?String(row[COL.name]).trim():"";if(!name)return{skip:true};var serial=row[COL.serialNumber]?String(row[COL.serialNumber]).trim():"",model=row[COL.modelNumber]?String(row[COL.modelNumber]).trim():"",code=serial||model||(name+"-"+(idx+1)),status=normalizeStudioStatus_(row[COL.status]);return{status:status,record:{code:code,status:status,category:normalizeStudioCategory_(name),name:name,qty:row[COL.qty]!==""&&row[COL.qty]!==undefined&&row[COL.qty]!==null?String(row[COL.qty]):"",purchaseDate:row[COL.purchaseDate]?formatDate_(row[COL.purchaseDate]):"",purchaseFrom:row[COL.purchaseFrom]?String(row[COL.purchaseFrom]):"",modelNumber:model,serialNumber:serial,checkedBy:row[COL.checkedBy]?String(row[COL.checkedBy]):"",checkedDate:row[COL.checkedDate]?formatDate_(row[COL.checkedDate]):"",location:row[COL.location]?String(row[COL.location]):"",remarks:COL.remarks>-1?String(row[COL.remarks]):""}};});}

function normalizeCourierStatus_(rawStatus){var s=String(rawStatus||"").trim().toUpperCase();if(!s)return"PENDING";if(s.indexOf("DELIVER")>-1)return"DELIVERED";if(s.indexOf("RETURN")>-1)return"RETURNED";if(s.indexOf("TRANSIT")>-1||s.indexOf("PENDING")>-1)return"PENDING";return s;}
function getCourierData(){
  requireDashboardAccess_();return readSheetData_(SHEET_NAME_COURIER,{assetCode:[["CLIENT ASSET CODE","ASSET CODE"],0],empName:[["EMPLOYEE NAME","EMP NAME"],1],clientName:[["CLIENT NAME"],2],clientAddress:[["CLIENT ADDRESS","ADDRESS"],3],dispatchDate:[["DISPATCH DATE"],4],courierName:[["COURIER NAME"],5],consignmentNo:[["CONSIGNMENT NO","CONSIGNMENT"],6],weightKg:[["WEIGHT"],7],totalAmount:[["TOTAL AMOUNT","AMOUNT"],8],itemSent:[["ITEM SENT","ITEM"],9],attachment:[["ATTACHMENT"],10],finalStatus:[["FINAL STATUS","STATUS"],11],remarks:[["REMARKS"],-1]},function(row,idx,COL){var assetCode=row[COL.assetCode]?String(row[COL.assetCode]).trim():"",empName=row[COL.empName]?String(row[COL.empName]).trim():"",consignmentNo=row[COL.consignmentNo]?String(row[COL.consignmentNo]).trim():"";if(!assetCode&&!empName&&!consignmentNo)return{skip:true};var code=assetCode||consignmentNo||(empName?(empName+"-"+(idx+2)):("ROW-"+(idx+2))),status=normalizeCourierStatus_(row[COL.finalStatus]);return{status:status,record:{rowNumber:idx+2,code:code,status:status,empName:empName,clientName:row[COL.clientName]?String(row[COL.clientName]):"",clientAddress:row[COL.clientAddress]?String(row[COL.clientAddress]):"",dispatchDate:row[COL.dispatchDate]?formatDate_(row[COL.dispatchDate]):"",courierName:row[COL.courierName]?String(row[COL.courierName]):"",consignmentNo:consignmentNo,weightKg:row[COL.weightKg]!==""&&row[COL.weightKg]!==undefined&&row[COL.weightKg]!==null?String(row[COL.weightKg]):"",totalAmount:row[COL.totalAmount]!==""&&row[COL.totalAmount]!==undefined&&row[COL.totalAmount]!==null?String(row[COL.totalAmount]):"",itemSent:row[COL.itemSent]?String(row[COL.itemSent]):"",attachment:row[COL.attachment]?String(row[COL.attachment]):"",finalStatusRaw:row[COL.finalStatus]?String(row[COL.finalStatus]):"Pending",remarks:COL.remarks>-1?String(row[COL.remarks]):""}};});}

// Courier_Data — add and edit every dashboard field. Attachment is stored unchanged as its Drive URL.
function courierColumns_(sheet){var h=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v||'').toUpperCase();});function c(names,f){for(var i=0;i<h.length;i++)for(var j=0;j<names.length;j++)if(h[i].indexOf(names[j])>-1)return i+1;return f;}return{lastCol:sheet.getLastColumn(),assetCode:c(['CLIENT ASSET CODE','ASSET CODE'],1),empName:c(['EMPLOYEE NAME','EMP NAME'],2),clientName:c(['CLIENT NAME'],3),clientAddress:c(['CLIENT ADDRESS','ADDRESS'],4),dispatchDate:c(['DISPATCH DATE'],5),courierName:c(['COURIER NAME'],6),consignmentNo:c(['CONSIGNMENT NO','CONSIGNMENT'],7),weightKg:c(['WEIGHT'],8),totalAmount:c(['TOTAL AMOUNT','AMOUNT'],9),itemSent:c(['ITEM SENT','ITEM'],10),attachment:c(['ATTACHMENT'],11),finalStatus:c(['FINAL STATUS','STATUS'],12),remarks:c(['REMARKS'],13)};}
function courierPayload_(p){p=p||{};var out={};['assetCode','empName','clientName','clientAddress','courierName','consignmentNo','weightKg','totalAmount','itemSent','attachment','finalStatus','remarks'].forEach(function(k){out[k]=String(p[k]||'').trim();});out.dispatchDate=p.dispatchDate?clientLaptopDate_(p.dispatchDate,'Dispatch Date'):'';return out;}
function writeCourier_(sheet,row,p){var c=courierColumns_(sheet);Object.keys(p).forEach(function(k){var cell=sheet.getRange(row,c[k]);if(k==='dispatchDate'){if(p[k])cell.setValue(p[k]).setNumberFormat('dd-mmm-yyyy');else cell.clearContent();}else cell.setValue(p[k]);});}
function addCourierData(payload){
  requireDashboardAccess_();var p=courierPayload_(payload),sheet=findSheetLoose_(getMainSpreadsheet_(),SHEET_NAME_COURIER);if(!sheet)throw new Error("Sheet '"+SHEET_NAME_COURIER+"' was not found.");var row=sheet.getLastRow()+1;writeCourier_(sheet,row,p);auditDashboardRow_(sheet,row,'COURIER RECORD CREATED');CacheService.getScriptCache().remove('asset:v2:active:'+SHEET_NAME_COURIER);return{ok:true,rowNumber:row};}
function updateCourierData(payload){
  requireDashboardAccess_();var row=Number(payload&&payload.rowNumber),sheet=findSheetLoose_(getMainSpreadsheet_(),SHEET_NAME_COURIER);if(!sheet)throw new Error("Sheet '"+SHEET_NAME_COURIER+"' was not found.");if(!row||row<2||row>sheet.getLastRow())throw new Error('Courier record was not found. Refresh and try again.');var before=activitySnapshot_(sheet,row);writeCourier_(sheet,row,courierPayload_(payload));auditDashboardRow_(sheet,row,'COURIER RECORD UPDATED',before);CacheService.getScriptCache().remove('asset:v2:active:'+SHEET_NAME_COURIER);return{ok:true,rowNumber:row};}

var MATERIALIN_CATEGORY_RULES_=[{label:"Laptop",keywords:["LAPTOP"]},{label:"CCTV Camera",keywords:["CAM DOM","CAMERA","DOME","PRAMA","HIKVISION CAM"]},{label:"NVR",keywords:["NVR"]},{label:"Network Switch",keywords:["SWITCH","POE","PORT"]},{label:"Rack",keywords:["RACK"]},{label:"Monitor",keywords:["MONITOR","SCREEN"]},{label:"UPS / Inverter",keywords:["UPS","INVERTOR","INVERTER"]},{label:"Printer",keywords:["PRINTER"]},{label:"WiFi Router / AP",keywords:["ROUTER","ACCESS POINT","UAP-"]},{label:"Firewall",keywords:["FIREWALL","FORTIGATE","UDM"]},{label:"Cables & Accessories",keywords:["CABLE","HDMI","USB"]}];
function normalizeMaterialInCategory_(item,materialDescription){var text=(String(item||"")+" "+String(materialDescription||"")).toUpperCase();for(var i=0;i<MATERIALIN_CATEGORY_RULES_.length;i++){var rule=MATERIALIN_CATEGORY_RULES_[i];for(var j=0;j<rule.keywords.length;j++)if(text.indexOf(rule.keywords[j])>-1)return rule.label;}var base=String(item||"").trim().replace(/^\d+\s*/,"").replace(/[\s\-_#]*\d+\s*$/,"").trim();return base?base:"Other";}
function normalizeMaterialInStatus_(rawRemarks){var s=String(rawRemarks||"").trim().toUpperCase();if(!s)return"PENDING";if(s.indexOf("DONE")>-1||s.indexOf("RECEIVED")>-1||s.indexOf("COMPLETE")>-1)return"RECEIVED";if(s.indexOf("PENDING")>-1)return"PENDING";return s;}
function getMaterialInData(){
  requireDashboardAccess_();try{var ss=getMainSpreadsheet_(),sheet=ss.getSheetByName(SHEET_NAME_MATERIALIN);if(!sheet)return{error:"Sheet '"+SHEET_NAME_MATERIALIN+"' not found."};var lastRow=sheet.getLastRow(),lastCol=sheet.getLastColumn();if(lastRow<2)return{error:"No data in '"+SHEET_NAME_MATERIALIN+"'."};if(lastCol<1)return{error:"'"+SHEET_NAME_MATERIALIN+"' has no columns."};if(lastRow>3000)lastRow=3000;if(lastCol>30)lastCol=30;var headerRow=sheet.getRange(1,1,1,lastCol).getValues()[0].map(function(h){return String(h||"").trim().toUpperCase();});function findCol(patterns){for(var i=0;i<headerRow.length;i++)for(var p=0;p<patterns.length;p++)if(headerRow[i].indexOf(patterns[p])>-1)return i;return -1;}var COL={sr:0,item:1,fromVendor:2,date:3,purpose:4,materialDescription:5,qty:6,receivedBy:7,checkedBy:8,location:9,remarks:lastCol-1},detected={sr:findCol(["S.R","SR NO","S NO","SR."]),item:findCol(["ITEM"]),fromVendor:findCol(["FROM"]),date:findCol(["DATE"]),purpose:findCol(["PURPOSE"]),materialDescription:findCol(["MATERIAL DESCRIPTION","DESCRIPTION"]),qty:findCol(["QTY","QUANTITY"]),receivedBy:findCol(["RECEIVED BY"]),checkedBy:findCol(["CHECKED BY"]),location:findCol(["LOCATION"])};Object.keys(detected).forEach(function(key){if(detected[key]>-1)COL[key]=detected[key];});var values=sheet.getRange(2,1,lastRow-1,lastCol).getValues(),counts={},records=[];values.forEach(function(row,idx){var item=row[COL.item]?String(row[COL.item]).trim():"",sr=row[COL.sr]!==""&&row[COL.sr]!==undefined&&row[COL.sr]!==null?String(row[COL.sr]).trim():"";if(!item&&!sr)return;var code=sr?"MI-"+sr:(item+"-"+(idx+2)),rawRemarks=row[COL.remarks]?String(row[COL.remarks]):"",status=normalizeMaterialInStatus_(rawRemarks);counts[status]=(counts[status]||0)+1;var materialDescription=row[COL.materialDescription]?String(row[COL.materialDescription]):"";records.push({code:code,status:status,category:normalizeMaterialInCategory_(item,materialDescription),sr:sr,item:item,fromVendor:row[COL.fromVendor]?String(row[COL.fromVendor]):"",date:row[COL.date]?formatDate_(row[COL.date]):"",purpose:row[COL.purpose]?String(row[COL.purpose]):"",materialDescription:materialDescription,qty:row[COL.qty]!==""&&row[COL.qty]!==undefined&&row[COL.qty]!==null?String(row[COL.qty]):"",receivedBy:row[COL.receivedBy]?String(row[COL.receivedBy]):"",checkedBy:row[COL.checkedBy]?String(row[COL.checkedBy]):"",location:row[COL.location]?String(row[COL.location]):"",remarks:rawRemarks});});return{total:records.length,counts:counts,records:records,updated:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy, hh:mm a")};}catch(e){return{error:"getMaterialInData failed: "+e.message};}}

function buildMacToAssetCodeMap_(){var mc=CacheService.getScriptCache(),mk="asset:v2:macmap",hit=mc.get(mk);if(hit){try{return JSON.parse(hit);}catch(e){mc.remove(mk);}}var ss=getMainSpreadsheet_(),map={};function indexSheet(sheetName,macPatterns,codePatterns){var sheet=findSheetLoose_(ss,sheetName);if(!sheet)return;var lastRow=sheet.getLastRow(),lastCol=sheet.getLastColumn();if(lastRow<2)return;if(lastRow>5000)lastRow=5000;if(lastCol>30)lastCol=30;var headerRow=sheet.getRange(1,1,1,lastCol).getValues()[0].map(function(h){return String(h||"").trim().toUpperCase();});function findCol(patterns){for(var i=0;i<headerRow.length;i++)for(var p=0;p<patterns.length;p++)if(headerRow[i].indexOf(patterns[p])>-1)return i;return-1;}var macCol=findCol(macPatterns),codeCol=findCol(codePatterns);if(macCol===-1||codeCol===-1)return;sheet.getRange(2,1,lastRow-1,lastCol).getValues().forEach(function(row){var mac=String(row[macCol]||"").trim().toUpperCase(),code=row[codeCol]?String(row[codeCol]).trim():"",macKey=mac.replace(/[^A-Z0-9]/g,"");if(macKey&&code)map[macKey]=code;});}indexSheet(SHEET_NAME_CLIENT,["MAC ADDRESS"],["ASSET CODE"]);indexSheet(SHEET_NAME_OFFICE,["MAC ADDRESS"],["CODE","ASSET CODE"]);try{var packed=JSON.stringify(map);if(packed.length<95000)mc.put(mk,packed,30);}catch(ignore){}return map;}
function buildEmpIdToNameMap_(){var mc=CacheService.getScriptCache(),mk="asset:v2:empmap",hit=mc.get(mk);if(hit){try{return JSON.parse(hit);}catch(e){mc.remove(mk);}}var ss=getMainSpreadsheet_(),map={};function indexSheet(sheetName,idPatterns,namePatterns){var sheet=findSheetLoose_(ss,sheetName);if(!sheet)return;var lastRow=sheet.getLastRow(),lastCol=sheet.getLastColumn();if(lastRow<2)return;if(lastRow>5000)lastRow=5000;if(lastCol>30)lastCol=30;var headerRow=sheet.getRange(1,1,1,lastCol).getValues()[0].map(function(h){return String(h||"").trim().toUpperCase();});function findCol(patterns){for(var i=0;i<headerRow.length;i++)for(var p=0;p<patterns.length;p++)if(headerRow[i].indexOf(patterns[p])>-1)return i;return-1;}var idCol=findCol(idPatterns),nameCol=findCol(namePatterns);if(idCol===-1||nameCol===-1)return;sheet.getRange(2,1,lastRow-1,lastCol).getValues().forEach(function(row){var id=String(row[idCol]||"").trim().toUpperCase(),name=row[nameCol]?String(row[nameCol]).trim():"";if(id&&name&&!isMacLike_(name)&&!map[id])map[id]=name;});}indexSheet(SHEET_NAME_CLIENT,["EMP ID"],["USER NAME","EMP NAME"]);indexSheet(SHEET_NAME_OFFICE,["EMP ID","EMPLOYEE ID"],["EMP NAME","EMPLOYEE NAME"]);try{var packed=JSON.stringify(map);if(packed.length<95000)mc.put(mk,packed,30);}catch(ignore){}return map;}
function resolveEmpName_(empIdToNameMap,empId,empName){if(empName)return empName;var key=String(empId||"").trim().toUpperCase();return key&&empIdToNameMap[key]?empIdToNameMap[key]:"";}
function resolveMacAssetCode_(macToAssetCode,macAddress,otherMac){var candidates=[macAddress,otherMac];for(var i=0;i<candidates.length;i++){var key=String(candidates[i]||"").trim().toUpperCase().replace(/[^A-Z0-9]/g,"");if(key&&macToAssetCode[key])return macToAssetCode[key];}return"";}
function normalizeMacAssetType_(raw){var s=String(raw||"").trim().toUpperCase();if(!s)return"UNKNOWN";s=s.replace(/N[\s\-\.\/]*A$/,"N/A");if(s.indexOf("OFFICE")>-1)return"OFFICE ASSET";if(s.indexOf("PERSONAL")>-1)return"PERSONAL";if(s.indexOf("DESKTOP")>-1)return"DESKTOP";if(s.indexOf("CLIENT")>-1)return"USING CLIENT ASSET";if(s.indexOf("GUEST")>-1)return"GUEST";return s;}

function getMacOfficeAssetData(){
  requireDashboardAccess_();var macToAssetCode=buildMacToAssetCodeMap_(),empIdToName=buildEmpIdToNameMap_();return readSheetData_(SHEET_NAME_MAC_OFFICE,{empId:[["EMPLOYEE ID","EMP ID","EMPLOYEE CODE"],0],empName:[["EMPLOYEE NAME","EMP NAME"],1],assetType:[["ASSET TYPE"],2],macAddress:[["MAC ADDRESS"],3],otherMac:[["OTHER MAC","MAC ADDRESS 2"],4],location:[["LOCATION"],5],remarks:[["REMARKS"],-1]},function(row,idx,COL){var empId=row[COL.empId]?String(row[COL.empId]).trim():"",empName=row[COL.empName]?String(row[COL.empName]).trim():"";if(!empId&&!empName)return{skip:true};var assetTypeRaw=row[COL.assetType]?String(row[COL.assetType]).trim():"",status=normalizeMacAssetType_(assetTypeRaw),location=row[COL.location]?String(row[COL.location]).trim():"",code=empId||(empName+"-"+(idx+2)),macAddress=row[COL.macAddress]?String(row[COL.macAddress]).trim():"",otherMac=row[COL.otherMac]?String(row[COL.otherMac]).trim():"";if(isMacLike_(empName)&&!macAddress){macAddress=empName;empName="";}empName=resolveEmpName_(empIdToName,empId,empName);return{status:status,record:{code:code,status:status,empId:empId,empName:empName,assetType:assetTypeRaw,macAddress:macAddress,otherMac:otherMac,assetCode:resolveMacAssetCode_(macToAssetCode,macAddress,otherMac),location:location,remarks:COL.remarks>-1?String(row[COL.remarks]||"").trim():""}};},MAC_SPREADSHEET_ID);}
function getMacClientLaptopData(){
  requireDashboardAccess_();var macToAssetCode=buildMacToAssetCodeMap_(),empIdToName=buildEmpIdToNameMap_();return readSheetData_(SHEET_NAME_MAC_CLIENT,{empId:[["EMPLOYEE ID","EMP ID","EMPLOYEE CODE"],0],empName:[["EMPLOYEE NAME","EMP NAME"],1],macAddress:[["MAC ADDRESS 1","MAC ADDRESS"],2],otherMac:[["MAC ADDRESS 2","OTHER MAC"],3],clientName:[["CLIENT NAME"],4],remarks:[["REMARKS"],-1]},function(row,idx,COL){var empId=row[COL.empId]?String(row[COL.empId]).trim():"",empName=row[COL.empName]?String(row[COL.empName]).trim():"";if(!empId&&!empName)return{skip:true};var clientName=row[COL.clientName]?String(row[COL.clientName]).trim():"",status="CLIENT LAPTOP",code=empId||(empName+"-"+(idx+2)),macAddress=row[COL.macAddress]?String(row[COL.macAddress]).trim():"",otherMac=row[COL.otherMac]?String(row[COL.otherMac]).trim():"";if(isMacLike_(empName)&&!macAddress){macAddress=empName;empName="";}if(!macAddress&&isMacLike_(otherMac)){macAddress=otherMac;otherMac="";}empName=resolveEmpName_(empIdToName,empId,empName);return{status:status,record:{code:code,status:status,empId:empId,empName:empName,macAddress:macAddress,otherMac:otherMac,assetCode:resolveMacAssetCode_(macToAssetCode,macAddress,otherMac),clientName:clientName,remarks:COL.remarks>-1?String(row[COL.remarks]||"").trim():""}};},MAC_SPREADSHEET_ID);}
function getMacFirewallData(){
  requireDashboardAccess_();var macToAssetCode=buildMacToAssetCodeMap_(),empIdToName=buildEmpIdToNameMap_();return readSheetData_(SHEET_NAME_MAC_FIREWALL,{empId:[["EMP CODE","EMPLOYEE CODE","EMP ID","EMPLOYEE ID"],0],empName:[["NAME","EMPLOYEE NAME","EMP NAME"],1],macAddress:[["MAC ADDRESS","MAC"],2],remarks:[["REMARKS","ON FIREWALL"],-1]},function(row,idx,COL){var empId=row[COL.empId]?String(row[COL.empId]).trim():"",empName=row[COL.empName]?String(row[COL.empName]).trim():"",macAddress=row[COL.macAddress]?String(row[COL.macAddress]).trim():"";if(!empId&&!empName&&!macAddress)return{skip:true};if(isMacLike_(empName)&&!macAddress){macAddress=empName;empName="";}empName=resolveEmpName_(empIdToName,empId,empName);var status="ALLOWED ON FIREWALL",code=empId||(empName?(empName+"-"+(idx+2)):("ROW-"+(idx+2)));return{status:status,record:{code:code,status:status,empId:empId,empName:empName,macAddress:macAddress,assetCode:resolveMacAssetCode_(macToAssetCode,macAddress,""),remarks:COL.remarks>-1?String(row[COL.remarks]||"").trim():""}};},MAC_SPREADSHEET_ID);}

// MAC workbook — Add and edit records for All Devices, Client Laptop and Firewall.
function macConfig_(kind){if(kind==='mac_office')return{sheet:SHEET_NAME_MAC_OFFICE,fields:['empId','empName','assetType','macAddress','otherMac','location','remarks']};if(kind==='mac_client')return{sheet:SHEET_NAME_MAC_CLIENT,fields:['empId','empName','macAddress','otherMac','clientName','remarks']};if(kind==='mac_firewall')return{sheet:SHEET_NAME_MAC_FIREWALL,fields:['empId','empName','macAddress','remarks']};throw new Error('Unknown MAC tab.');}
function macColumns_(sheet,kind){var h=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v||'').toUpperCase();}), cfg=macConfig_(kind);function c(names,f){for(var i=0;i<h.length;i++)for(var j=0;j<names.length;j++)if(h[i].indexOf(names[j])>-1)return i+1;return f;}return{lastCol:sheet.getLastColumn(),empId:c(['EMPLOYEE ID','EMP ID','EMPLOYEE CODE','EMP CODE'],1),empName:c(['EMPLOYEE NAME','EMP NAME','NAME'],2),assetType:c(['ASSET TYPE'],3),macAddress:c(['MAC ADDRESS 1','MAC ADDRESS','MAC'],4),otherMac:c(['MAC ADDRESS 2','OTHER MAC'],5),location:c(['LOCATION'],6),clientName:c(['CLIENT NAME'],5),remarks:c(['REMARKS','ON FIREWALL'],7),fields:cfg.fields};}
function macText_(p,k){return String((p||{})[k]||'').trim();}
function writeMacRecord_(sheet,kind,row,p){var c=macColumns_(sheet,kind);c.fields.forEach(function(k){sheet.getRange(row,c[k]).setValue(macText_(p,k));});}
function addMacRecord(kind,payload){
  requireDashboardAccess_();var cfg=macConfig_(kind),ss=SpreadsheetApp.openById(MAC_SPREADSHEET_ID),sheet=findSheetLoose_(ss,cfg.sheet);if(!sheet)throw new Error("Sheet '"+cfg.sheet+"' was not found.");var row=sheet.getLastRow()+1;writeMacRecord_(sheet,kind,row,payload);auditDashboardRow_(sheet,row,'MAC RECORD CREATED');CacheService.getScriptCache().remove('asset:v2:'+MAC_SPREADSHEET_ID+':'+cfg.sheet);return{ok:true,rowNumber:row};}
function updateMacRecord(kind,payload){
  requireDashboardAccess_();var cfg=macConfig_(kind),ss=SpreadsheetApp.openById(MAC_SPREADSHEET_ID),sheet=findSheetLoose_(ss,cfg.sheet),c;if(!sheet)throw new Error("Sheet '"+cfg.sheet+"' was not found.");c=macColumns_(sheet,kind);var last=sheet.getLastRow(),row=Number(payload&&payload.rowNumber);if(!row||row<2||row>last){var id=macText_(payload,'originalEmpId'),mac=macText_(payload,'originalMac'),name=macText_(payload,'originalEmpName'),values=sheet.getRange(2,1,last-1,Math.max(c.empId,c.empName,c.macAddress)).getDisplayValues();for(var i=0;i<values.length;i++)if((id&&String(values[i][c.empId-1]).trim()===id)||(mac&&String(values[i][c.macAddress-1]).trim()===mac)||(!id&&!mac&&name&&String(values[i][c.empName-1]).trim()===name)){row=i+2;break;}}if(!row||row<2||row>last)throw new Error('MAC record was not found. Refresh and try again.');var before=activitySnapshot_(sheet,row);writeMacRecord_(sheet,kind,row,payload);auditDashboardRow_(sheet,row,'MAC RECORD UPDATED',before);CacheService.getScriptCache().remove('asset:v2:'+MAC_SPREADSHEET_ID+':'+cfg.sheet);return{ok:true,rowNumber:row};}

function normalizeAssetDetailsStatus_(rawStatus){var s=String(rawStatus||"").trim().toUpperCase();if(!s)return"NOT SET";return s;}

// ============================================================
// OFFICE LAPTOP ASSIGNMENT / SUBMISSION — called from dashboard.html
// Status, Employee Code, Employee Name, Assign Date, Remarks and Notes are editable.
// The selected status is always written to Office_Laptops column B.
// ============================================================
var OFFICE_LAPTOP_ALLOWED_STATUSES_ = [
  'ASSIGNED', 'NOT ASSIGNED', 'IN STOCK', 'GIVEN', 'ON_REPAIR', 'MISSING',
  'SOLD', 'DEAD STOCK', 'FAULTY', 'TEMP ASSIGN', 'INTERVIEW'
];

function findOfficeAssignmentColumn_(headers, patterns, fallbackColumn) {
  for (var i = 0; i < headers.length; i++) {
    for (var p = 0; p < patterns.length; p++) {
      if (headers[i].indexOf(patterns[p]) > -1) return i + 1;
    }
  }
  return fallbackColumn;
}

function parseOfficeAssignmentDate_(value) {
  var text = String(value || '').trim();
  var match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  var year = Number(match[1]);
  var month = Number(match[2]);
  var day = Number(match[3]);
  var date = new Date(year, month - 1, day, 12, 0, 0);
  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day ? date : null;
}

function setOfficeAssignmentCell_(sheet, row, column, value) {
  var cell = sheet.getRange(row, column);
  if (value === '' || value === null || value === undefined) cell.clearContent();
  else cell.setValue(value);
  return cell;
}

// Keep the dashboard dropdown and the Google Sheets dropdown in sync.  This
// validation is added whenever a status is created or changed, so the new row
// does not become a free-text status cell in the register.
function applyStatusDropdown_(sheet, row, column, allowedStatuses) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allowedStatuses, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, column).setDataValidation(rule);
}

function clearOfficeAssignmentCaches_() {
  var cache = CacheService.getScriptCache();
  var keys = [
    'asset:v2:active:' + SHEET_NAME_OFFICE,
    'asset:v3:active:' + SHEET_NAME_OFFICE,
    'asset:v3:active:' + SHEET_NAME_OFFICE + ':gz',
    'asset:v2:macmap',
    'asset:v2:empmap'
  ];
  cache.removeAll(keys);
}

// Creates a new Office_Laptops row from the dashboard's Add New Laptop form.
// Field columns are detected by their headers; Asset Status is explicitly kept
// in column B and Notes in column G to match the existing register layout.
function addOfficeLaptop(payload) {
  requireDashboardAccess_();
  payload = payload || {};
  var assetCode = String(payload.assetCode || '').trim();
  var status = String(payload.status || 'IN STOCK').trim().toUpperCase();
  var empId = String(payload.empId || '').trim();
  var empName = String(payload.empName || '').trim();
  var remarks = String(payload.remarks || '').trim();
  var notes = String(payload.notes || '').trim();
  var assignDateText = String(payload.assignDate || '').trim();
  var assignDate = assignDateText ? parseOfficeAssignmentDate_(assignDateText) : null;

  if (!assetCode) throw new Error('Asset Code is required.');
  if (OFFICE_LAPTOP_ALLOWED_STATUSES_.indexOf(status) === -1) {
    throw new Error('Please select a valid Asset Status.');
  }

  // Keep new assets consistent with the rules already used by the update form.
  if (status === 'NOT ASSIGNED') {
    empId = '';
    empName = '';
    assignDateText = '';
    assignDate = null;
    remarks = 'Ready to Assign';
    notes = '';
  } else if (status === 'ASSIGNED') {
    if (!assignDate) assignDate = new Date();
    remarks = 'DONE';
  }
  if (assignDateText && !assignDate) throw new Error('Please select a valid Assign Date.');

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var ss = getMainSpreadsheet_();
    var sheet = findSheetLoose_(ss, SHEET_NAME_OFFICE);
    if (!sheet) throw new Error("Sheet '" + SHEET_NAME_OFFICE + "' was not found.");

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 1) throw new Error("Sheet '" + SHEET_NAME_OFFICE + "' needs a header row before adding an asset.");

    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (value) {
      return String(value || '').trim().toUpperCase();
    });
    var codeColumn = findOfficeAssignmentColumn_(headers, ['CODE', 'ASSET CODE'], 1);
    var statusColumn = 2; // Office_Laptops column B
    var empIdColumn = findOfficeAssignmentColumn_(headers, ['EMP ID', 'EMPLOYEE ID'], 4);
    var empNameColumn = findOfficeAssignmentColumn_(headers, ['EMP NAME', 'EMPLOYEE NAME'], 5);
    var dateColumn = findOfficeAssignmentColumn_(headers, ['ASSIGNED DATE', 'DATE'], 6);
    var notesColumn = 7; // Office_Laptops column G
    var remarksColumn = findOfficeAssignmentColumn_(headers, ['REMARKS'], 8);

    if (lastRow >= 2) {
      var existingCodes = sheet.getRange(2, codeColumn, lastRow - 1, 1).getDisplayValues();
      for (var i = 0; i < existingCodes.length; i++) {
        if (String(existingCodes[i][0] || '').trim().toUpperCase() === assetCode.toUpperCase()) {
          throw new Error('Asset Code "' + assetCode + '" already exists in ' + SHEET_NAME_OFFICE + '.');
        }
      }
    }

    var rowWidth = Math.max(lastCol, codeColumn, statusColumn, empIdColumn, empNameColumn, dateColumn, notesColumn, remarksColumn);
    var newRowValues = [];
    for (var col = 0; col < rowWidth; col++) newRowValues.push('');
    newRowValues[codeColumn - 1] = assetCode;
    newRowValues[statusColumn - 1] = status;
    newRowValues[empIdColumn - 1] = empId;
    newRowValues[empNameColumn - 1] = empName;
    newRowValues[dateColumn - 1] = assignDate || '';
    newRowValues[notesColumn - 1] = notes;
    newRowValues[remarksColumn - 1] = remarks;

    var newRow = lastRow + 1;
    sheet.getRange(newRow, 1, 1, rowWidth).setValues([newRowValues]);
    applyStatusDropdown_(sheet, newRow, statusColumn, OFFICE_LAPTOP_ALLOWED_STATUSES_);
    if (assignDate) sheet.getRange(newRow, dateColumn).setNumberFormat('dd-mmm-yyyy');

    auditDashboardRow_(sheet, newRow, 'ASSET CREATED');
    clearOfficeAssignmentCaches_();
    return { ok: true, assetCode: assetCode, status: status };
  } finally {
    lock.releaseLock();
  }
}

function assignOfficeLaptop(payload) {
  payload = payload || {};
  var assetCode = String(payload.assetCode || '').trim();
  var empId = String(payload.empId || '').trim();
  var empName = String(payload.empName || '').trim();
  var remarks = String(payload.remarks || '').trim();
  var notes = String(payload.notes || '').trim();
  var requestedStatus = String(payload.status || '').trim().toUpperCase();
  var assignDateText = String(payload.assignDate || '').trim();
  var assignDate = assignDateText ? parseOfficeAssignmentDate_(assignDateText) : null;

  if (!assetCode) throw new Error('Asset Code is required.');
  if (requestedStatus && OFFICE_LAPTOP_ALLOWED_STATUSES_.indexOf(requestedStatus) === -1) {
    throw new Error('Please select a valid Asset Status.');
  }

  // These values are also enforced on the server so they remain correct even
  // if a stale browser tab sends an older form payload.
  if (requestedStatus === 'NOT ASSIGNED') {
    empId = '';
    empName = '';
    assignDateText = '';
    assignDate = null;
    remarks = 'Ready to Assign';
    notes = '';
  } else if (requestedStatus === 'ASSIGNED') {
    if (!assignDate) assignDate = new Date();
    remarks = 'DONE';
  }

  if (assignDateText && !assignDate) throw new Error('Please select a valid Assign Date.');

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var ss = getMainSpreadsheet_();
    var sheet = findSheetLoose_(ss, SHEET_NAME_OFFICE);
    if (!sheet) throw new Error("Sheet '" + SHEET_NAME_OFFICE + "' was not found.");

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) throw new Error("Sheet '" + SHEET_NAME_OFFICE + "' has no asset rows.");

    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (value) {
      return String(value || '').trim().toUpperCase();
    });
    var codeColumn = findOfficeAssignmentColumn_(headers, ['CODE', 'ASSET CODE'], 1);
    var statusColumn = 2; // Office_Laptops column B — Asset Status
    var empIdColumn = findOfficeAssignmentColumn_(headers, ['EMP ID', 'EMPLOYEE ID'], 4);
    var empNameColumn = findOfficeAssignmentColumn_(headers, ['EMP NAME', 'EMPLOYEE NAME'], 5);
    var dateColumn = findOfficeAssignmentColumn_(headers, ['ASSIGNED DATE', 'DATE'], 6);
    var notesColumn = 7; // User-requested G column
    var remarksColumn = findOfficeAssignmentColumn_(headers, ['REMARKS'], 8);

    var codes = sheet.getRange(2, codeColumn, lastRow - 1, 1).getDisplayValues();
    var targetRow = 0;
    for (var i = 0; i < codes.length; i++) {
      if (String(codes[i][0] || '').trim().toUpperCase() === assetCode.toUpperCase()) {
        targetRow = i + 2;
        break;
      }
    }
    if (!targetRow) throw new Error('Asset Code "' + assetCode + '" was not found in ' + SHEET_NAME_OFFICE + '.');

    var before = activitySnapshot_(sheet, targetRow);
    setOfficeAssignmentCell_(sheet, targetRow, empIdColumn, empId);
    setOfficeAssignmentCell_(sheet, targetRow, empNameColumn, empName);
    if (assignDate) {
      sheet.getRange(targetRow, dateColumn).setValue(assignDate).setNumberFormat('dd-mmm-yyyy');
    } else {
      sheet.getRange(targetRow, dateColumn).clearContent();
    }
    setOfficeAssignmentCell_(sheet, targetRow, notesColumn, notes);
    setOfficeAssignmentCell_(sheet, targetRow, remarksColumn, remarks);

    // Old dashboard versions do not send status, so retain their behaviour as
    // a safe fallback. The current dashboard always sends the selected value.
    var resultingStatus = requestedStatus || ((empId || empName || assignDate) ? 'ASSIGNED' : 'NOT ASSIGNED');
    sheet.getRange(targetRow, statusColumn).setValue(resultingStatus);
    applyStatusDropdown_(sheet, targetRow, statusColumn, OFFICE_LAPTOP_ALLOWED_STATUSES_);
    auditDashboardRow_(sheet, targetRow, resultingStatus === 'NOT ASSIGNED' ? 'ASSET SUBMITTED' : (resultingStatus === 'ASSIGNED' ? 'ASSET ASSIGNED' : 'ASSIGNMENT UPDATED'), before);

    // Returning the complete dashboard data here made the Save button wait for
    // a full Office_Laptops read. The browser now does that refresh separately.
    // No flush is needed: Apps Script commits pending sheet changes on completion.
    clearOfficeAssignmentCaches_();
    return {
      ok: true,
      assetCode: assetCode,
      status: resultingStatus,
      submitted: resultingStatus === 'NOT ASSIGNED'
    };
  } finally {
    lock.releaseLock();
  }
}

// Include the user-requested Notes field from column G in the Office Laptops
// dashboard data, while keeping all existing fields and matching logic.
function getDashboardData() {
  return readSheetData_(SHEET_NAME_OFFICE, {
    code: [['CODE', 'ASSET CODE'], 0],
    status: [['STATUS', 'ASSET STATUS'], 1],
    empId: [['EMP ID', 'EMPLOYEE ID'], 3],
    empName: [['EMP NAME', 'EMPLOYEE NAME'], 4],
    date: [['DATE', 'ASSIGNED DATE'], 5],
    notes: [['NOTES', 'NOTE'], 6],
    remarks: [['REMARKS'], 7]
  }, function (row, idx, COL) {
    var code = row[COL.code];
    if (!code) return { skip: true };
    var status = String(row[COL.status] || 'UNKNOWN').trim().toUpperCase();
    return {
      status: status,
      record: {
        code: String(code),
        status: status,
        category: isPcAssetCode_(code) ? 'PC / Server' : 'Laptop',
        empId: row[COL.empId] ? String(row[COL.empId]) : '',
        empName: row[COL.empName] ? String(row[COL.empName]) : '',
        date: row[COL.date] ? formatDate_(row[COL.date]) : '',
        notes: row[COL.notes] ? String(row[COL.notes]) : '',
        remarks: row[COL.remarks] ? String(row[COL.remarks]) : ''
      }
    };
  });
}

// ============================================================
// CLIENT LAPTOPS — dashboard edit/add flow and sheet dropdown.
// Client Status is intentionally separate from the register's other Status
// column: ACTIVE, SUBMITTED and DISPATCHED always write to Client Status.
// ============================================================
var CLIENT_LAPTOP_ALLOWED_STATUSES_ = ['ACTIVE', 'SUBMITTED', 'DISPATCHED'];

function clearClientLaptopCaches_() {
  CacheService.getScriptCache().removeAll([
    'asset:v2:active:' + SHEET_NAME_CLIENT,
    'asset:v3:active:' + SHEET_NAME_CLIENT,
    'asset:v3:active:' + SHEET_NAME_CLIENT + ':gz',
    'asset:v2:macmap',
    'asset:v2:empmap'
  ]);
}

function clientLaptopColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error("Sheet '" + SHEET_NAME_CLIENT + "' needs a header row.");
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (value) {
    return String(value || '').trim().toUpperCase();
  });
  function column(patterns, fallback) {
    // Prefer exact header names.  In particular, a generic "CODE" match must
    // never select the Client Code column instead of the Asset Code column.
    for (var p = 0; p < patterns.length; p++) {
      for (var h = 0; h < headers.length; h++) {
        if (headers[h] === patterns[p]) return h + 1;
      }
    }
    return findOfficeAssignmentColumn_(headers, patterns, fallback);
  }
  return {
    lastCol: lastCol,
    clientStatus: column(['CLIENT STATUS', 'STATUS'], 1),
    clientName: column(['CLIENT NAME'], 2),
    clientCode: column(['CLIENT CODE'], 3),
    project: column(['PROJECT'], 4),
    assignDate: column(['ASSIGN DATE', 'ASSIGNED DATE'], 5),
    userName: column(['USER NAME', 'EMP NAME'], 6),
    empId: column(['EMP ID', 'EMPLOYEE ID'], 7),
    cyntexaLoc: column(['CYNTEXA LOC', 'LOCATION'], 8),
    assetCode: column(['ASSET CODE', 'CODE'], 9),
    assetStatus: column(['ASSET STATUS'], 10),
    assetType: column(['ASSET TYPE'], 11),
    macAddress: column(['MAC ADDRESS'], 12),
    makeModel: column(['MAKE MODEL', 'MODEL'], 13),
    serialNumber: column(['SERIAL NUMBER', 'SERIAL NO'], 14),
    processor: column(['PROCESSOR'], 15),
    ram: column(['RAM'], 16),
    ssd: column(['SSD'], 17),
    itemsReceived: column(['ITEMS RECEIVED'], 18),
    otherAssetCode: column(['OTHER ASSET CODE'], 19),
    laptopSubmitDate: column(['LAPTOP SUBMIT DATE', 'SUBMIT DATE'], 20),
    dispatchedDate: column(['DISPATCHED DATE'], 21),
    courierName: column(['COURIER NAME'], 22),
    dispatchedToPerson: column(['DISPATCHED TO'], 23),
    dispatchedItems: column(['DISPATCHED ITEMS'], 24),
    remarks: column(['REMARKS'], Math.max(lastCol, 25))
  };
}

function clientLaptopDate_(value, label) {
  var text = String(value || '').trim();
  if (!text) return '';
  var date = parseOfficeAssignmentDate_(text);
  if (!date) throw new Error('Please select a valid ' + label + '.');
  return date;
}

function clientLaptopText_(payload, key) {
  return String(payload[key] || '').trim();
}

function addClientLaptop(payload) {
  requireDashboardAccess_();
  payload = payload || {};
  var assetCode = clientLaptopText_(payload, 'assetCode');
  var clientStatus = clientLaptopText_(payload, 'clientStatus').toUpperCase() || 'ACTIVE';
  if (!assetCode) throw new Error('Asset Code is required.');
  if (CLIENT_LAPTOP_ALLOWED_STATUSES_.indexOf(clientStatus) === -1) {
    throw new Error('Please select ACTIVE, SUBMITTED, or DISPATCHED.');
  }

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sheet = findSheetLoose_(getMainSpreadsheet_(), SHEET_NAME_CLIENT);
    if (!sheet) throw new Error("Sheet '" + SHEET_NAME_CLIENT + "' was not found.");
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) throw new Error("Sheet '" + SHEET_NAME_CLIENT + "' needs a header row before adding an asset.");

    var columns = clientLaptopColumns_(sheet);
    if (lastRow >= 2) {
      var existingCodes = sheet.getRange(2, columns.assetCode, lastRow - 1, 1).getDisplayValues();
      for (var i = 0; i < existingCodes.length; i++) {
        if (String(existingCodes[i][0] || '').trim().toUpperCase() === assetCode.toUpperCase()) {
          throw new Error('Asset Code "' + assetCode + '" already exists in ' + SHEET_NAME_CLIENT + '.');
        }
      }
    }

    var valuesByField = {
      clientStatus: clientStatus,
      clientName: clientLaptopText_(payload, 'clientName'),
      clientCode: clientLaptopText_(payload, 'clientCode'),
      project: clientLaptopText_(payload, 'project'),
      assignDate: clientLaptopDate_(payload.assignDate, 'Assign Date'),
      userName: clientLaptopText_(payload, 'userName'),
      empId: clientLaptopText_(payload, 'empId'),
      cyntexaLoc: clientLaptopText_(payload, 'cyntexaLoc'),
      assetCode: assetCode,
      assetType: clientLaptopText_(payload, 'assetType'),
      macAddress: clientLaptopText_(payload, 'macAddress'),
      makeModel: clientLaptopText_(payload, 'makeModel'),
      serialNumber: clientLaptopText_(payload, 'serialNumber'),
      processor: clientLaptopText_(payload, 'processor'),
      ram: clientLaptopText_(payload, 'ram'),
      ssd: clientLaptopText_(payload, 'ssd'),
      itemsReceived: clientLaptopText_(payload, 'itemsReceived'),
      otherAssetCode: clientLaptopText_(payload, 'otherAssetCode'),
      laptopSubmitDate: clientLaptopDate_(payload.laptopSubmitDate, 'Laptop Submit Date'),
      dispatchedDate: clientLaptopDate_(payload.dispatchedDate, 'Dispatched Date'),
      courierName: clientLaptopText_(payload, 'courierName'),
      dispatchedToPerson: clientLaptopText_(payload, 'dispatchedToPerson'),
      dispatchedItems: clientLaptopText_(payload, 'dispatchedItems'),
      remarks: clientLaptopText_(payload, 'remarks')
    };
    var rowWidth = columns.lastCol;
    Object.keys(columns).forEach(function (key) {
      if (key !== 'lastCol') rowWidth = Math.max(rowWidth, columns[key]);
    });
    var newValues = [];
    for (var c = 0; c < rowWidth; c++) newValues.push('');
    Object.keys(valuesByField).forEach(function (key) {
      newValues[columns[key] - 1] = valuesByField[key];
    });

    var newRow = lastRow + 1;
    sheet.getRange(newRow, 1, 1, rowWidth).setValues([newValues]);
    applyStatusDropdown_(sheet, newRow, columns.clientStatus, CLIENT_LAPTOP_ALLOWED_STATUSES_);
    [columns.assignDate, columns.laptopSubmitDate, columns.dispatchedDate].forEach(function (column) {
      if (sheet.getRange(newRow, column).getValue() instanceof Date) {
        sheet.getRange(newRow, column).setNumberFormat('dd-mmm-yyyy');
      }
    });
    auditDashboardRow_(sheet, newRow, 'CLIENT ASSET CREATED');
    clearClientLaptopCaches_();
    return { ok: true, assetCode: assetCode, status: clientStatus };
  } finally {
    lock.releaseLock();
  }
}

function updateClientLaptopStatus(payload) {
  requireDashboardAccess_();
  payload = payload || {};
  var assetCode = clientLaptopText_(payload, 'assetCode');
  var clientStatus = clientLaptopText_(payload, 'clientStatus').toUpperCase();
  if (!assetCode) throw new Error('Asset Code is required.');
  if (CLIENT_LAPTOP_ALLOWED_STATUSES_.indexOf(clientStatus) === -1) {
    throw new Error('Please select ACTIVE, SUBMITTED, or DISPATCHED.');
  }

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sheet = findSheetLoose_(getMainSpreadsheet_(), SHEET_NAME_CLIENT);
    if (!sheet) throw new Error("Sheet '" + SHEET_NAME_CLIENT + "' was not found.");
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Sheet '" + SHEET_NAME_CLIENT + "' has no asset rows.");
    var columns = clientLaptopColumns_(sheet);
    var codes = sheet.getRange(2, columns.assetCode, lastRow - 1, 1).getDisplayValues();
    var targetRow = 0;
    for (var i = 0; i < codes.length; i++) {
      if (String(codes[i][0] || '').trim().toUpperCase() === assetCode.toUpperCase()) {
        targetRow = i + 2;
        break;
      }
    }
    if (!targetRow) throw new Error('Asset Code "' + assetCode + '" was not found in ' + SHEET_NAME_CLIENT + '.');

    var before = activitySnapshot_(sheet, targetRow);
    sheet.getRange(targetRow, columns.clientStatus).setValue(clientStatus);
    applyStatusDropdown_(sheet, targetRow, columns.clientStatus, CLIENT_LAPTOP_ALLOWED_STATUSES_);
    auditDashboardRow_(sheet, targetRow, clientStatus === 'SUBMITTED' ? 'CLIENT ASSET SUBMITTED' : 'CLIENT STATUS UPDATED', before);
    clearClientLaptopCaches_();
    return { ok: true, assetCode: assetCode, status: clientStatus };
  } finally {
    lock.releaseLock();
  }
}


// Edit every dashboard-managed field of an existing Client_Laptops row.
// originalAssetCode is used to locate the row, so Asset Code itself can also be changed.
function updateClientLaptop(payload) {
  requireDashboardAccess_();
  payload = payload || {};
  var originalAssetCode = clientLaptopText_(payload, 'originalAssetCode') || clientLaptopText_(payload, 'assetCode');
  var assetCode = clientLaptopText_(payload, 'assetCode');
  var clientStatus = clientLaptopText_(payload, 'clientStatus').toUpperCase() || 'ACTIVE';

  if (!originalAssetCode) throw new Error('Original Asset Code is required.');
  if (!assetCode) throw new Error('Asset Code is required.');
  if (CLIENT_LAPTOP_ALLOWED_STATUSES_.indexOf(clientStatus) === -1) {
    throw new Error('Please select ACTIVE, SUBMITTED, or DISPATCHED.');
  }

  var assignDate = clientLaptopDate_(payload.assignDate, 'Assign Date');
  var laptopSubmitDate = clientLaptopDate_(payload.laptopSubmitDate, 'Laptop Submit Date');
  var dispatchedDate = clientLaptopDate_(payload.dispatchedDate, 'Dispatched Date');

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sheet = findSheetLoose_(getMainSpreadsheet_(), SHEET_NAME_CLIENT);
    if (!sheet) throw new Error("Sheet '" + SHEET_NAME_CLIENT + "' was not found.");

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Sheet '" + SHEET_NAME_CLIENT + "' has no asset rows.");

    var columns = clientLaptopColumns_(sheet);
    var codes = sheet.getRange(2, columns.assetCode, lastRow - 1, 1).getDisplayValues();
    var targetRow = 0;

    for (var i = 0; i < codes.length; i++) {
      if (String(codes[i][0] || '').trim().toUpperCase() === originalAssetCode.toUpperCase()) {
        targetRow = i + 2;
        break;
      }
    }
    if (!targetRow) {
      throw new Error('Asset Code "' + originalAssetCode + '" was not found in ' + SHEET_NAME_CLIENT + '.');
    }

    var before = activitySnapshot_(sheet, targetRow);

    // If Asset Code is changed, do not allow a duplicate.
    if (assetCode.toUpperCase() !== originalAssetCode.toUpperCase()) {
      for (var j = 0; j < codes.length; j++) {
        if (j + 2 === targetRow) continue;
        if (String(codes[j][0] || '').trim().toUpperCase() === assetCode.toUpperCase()) {
          throw new Error('Asset Code "' + assetCode + '" already exists in ' + SHEET_NAME_CLIENT + '.');
        }
      }
    }

    var valuesByField = {
      clientStatus: clientStatus,
      clientName: clientLaptopText_(payload, 'clientName'),
      clientCode: clientLaptopText_(payload, 'clientCode'),
      project: clientLaptopText_(payload, 'project'),
      assignDate: assignDate || '',
      userName: clientLaptopText_(payload, 'userName'),
      empId: clientLaptopText_(payload, 'empId'),
      cyntexaLoc: clientLaptopText_(payload, 'cyntexaLoc'),
      assetCode: assetCode,
      assetType: clientLaptopText_(payload, 'assetType'),
      macAddress: clientLaptopText_(payload, 'macAddress'),
      makeModel: clientLaptopText_(payload, 'makeModel'),
      serialNumber: clientLaptopText_(payload, 'serialNumber'),
      processor: clientLaptopText_(payload, 'processor'),
      ram: clientLaptopText_(payload, 'ram'),
      ssd: clientLaptopText_(payload, 'ssd'),
      itemsReceived: clientLaptopText_(payload, 'itemsReceived'),
      otherAssetCode: clientLaptopText_(payload, 'otherAssetCode'),
      laptopSubmitDate: laptopSubmitDate || '',
      dispatchedDate: dispatchedDate || '',
      courierName: clientLaptopText_(payload, 'courierName'),
      dispatchedToPerson: clientLaptopText_(payload, 'dispatchedToPerson'),
      dispatchedItems: clientLaptopText_(payload, 'dispatchedItems'),
      remarks: clientLaptopText_(payload, 'remarks')
    };

    Object.keys(valuesByField).forEach(function (key) {
      var col = columns[key];
      if (!col) return;
      sheet.getRange(targetRow, col).setValue(valuesByField[key]);
    });

    [columns.assignDate, columns.laptopSubmitDate, columns.dispatchedDate].forEach(function (column) {
      if (column && sheet.getRange(targetRow, column).getValue() instanceof Date) {
        sheet.getRange(targetRow, column).setNumberFormat('dd-mmm-yyyy');
      }
    });

    applyStatusDropdown_(sheet, targetRow, columns.clientStatus, CLIENT_LAPTOP_ALLOWED_STATUSES_);
    auditDashboardRow_(sheet, targetRow, clientStatus === 'SUBMITTED' ? 'CLIENT ASSET SUBMITTED' : 'CLIENT ASSET UPDATED', before);
    clearClientLaptopCaches_();

    return {
      ok: true,
      assetCode: assetCode,
      originalAssetCode: originalAssetCode,
      status: clientStatus
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// PERFORMANCE LAYER — cached reads, compressed cache values,
// spreadsheet-handle reuse and date-format memoization.
// ============================================================
var DASHBOARD_CACHE_PREFIX_ = 'asset:v3:';
var DASHBOARD_CACHE_TTL_SECONDS_ = 300; // 5 minutes: dashboard opens are served from cache; writes clear the affected cache.
var DASHBOARD_SPREADSHEET_CACHE_ = {};
var DASHBOARD_DATE_CACHE_ = {};
var DASHBOARD_DATE_CACHE_SIZE_ = 0;

function getDashboardSpreadsheet_(spreadsheetId) {
  var key = spreadsheetId || '__ACTIVE__';
  if (DASHBOARD_SPREADSHEET_CACHE_[key]) return DASHBOARD_SPREADSHEET_CACHE_[key];
  var ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : getMainSpreadsheet_();
  DASHBOARD_SPREADSHEET_CACHE_[key] = ss;
  return ss;
}

function getDashboardCache_(cache, key) {
  var raw = cache.get(key);
  if (raw) {
    try { return JSON.parse(raw); }
    catch (e) { cache.remove(key); }
  }

  var packed = cache.get(key + ':gz');
  if (!packed) return null;
  try {
    var json = Utilities.ungzip(Utilities.newBlob(Utilities.base64DecodeWebSafe(packed))).getDataAsString();
    return JSON.parse(json);
  } catch (e2) {
    cache.remove(key + ':gz');
    return null;
  }
}

function putDashboardCache_(cache, key, value) {
  try {
    var json = JSON.stringify(value);
    if (json.length <= 95000) {
      cache.put(key, json, DASHBOARD_CACHE_TTL_SECONDS_);
      return;
    }

    // CacheService has a 100 KB per-value limit. Gzip lets sizeable repeated
    // sheet data (which is common in asset registers) fit without dropping
    // any columns or records.
    var packed = Utilities.base64EncodeWebSafe(Utilities.gzip(Utilities.newBlob(json)).getBytes());
    if (packed.length <= 95000) cache.put(key + ':gz', packed, DASHBOARD_CACHE_TTL_SECONDS_);
  } catch (ignore) {}
}

function readSheetData_(sheetName, colMap, rowProcessor, spreadsheetId, maxRows) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = DASHBOARD_CACHE_PREFIX_ + (spreadsheetId || 'active') + ':' + sheetName;
    var cached = getDashboardCache_(cache, cacheKey);
    if (cached) return cached;

    if (spreadsheetId && spreadsheetId.indexOf('PASTE_') === 0) {
      return { error: "Set MAC_SPREADSHEET_ID in Code.gs to the other workbook's Spreadsheet ID first (see the comment above it)." };
    }

    var ss = getDashboardSpreadsheet_(spreadsheetId);
    var sheet = findSheetLoose_(ss, sheetName);
    if (!sheet) {
      var availableNames = ss.getSheets().map(function (sh) { return sh.getName(); }).join(', ');
      return { error: "Sheet '" + sheetName + "' not found. Tabs available in that workbook: " + availableNames };
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { error: "No data in '" + sheetName + "'." };
    if (lastRow > (maxRows || 5000)) lastRow = maxRows || 5000;
    if (lastCol > 30) lastCol = 30;

    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
      return String(h || '').trim().toUpperCase();
    });
    function findCol(patterns) {
      for (var i = 0; i < headerRow.length; i++) {
        for (var p = 0; p < patterns.length; p++) {
          if (headerRow[i].indexOf(patterns[p]) > -1) return i;
        }
      }
      return -1;
    }

    var detected = {};
    Object.keys(colMap).forEach(function (key) {
      var found = findCol(colMap[key][0]);
      if (found > -1) detected[key] = found;
    });

    var COL = {};
    Object.keys(colMap).forEach(function (key) {
      var fallback = colMap[key][1];
      COL[key] = detected.hasOwnProperty(key) ? detected[key] : (fallback === '__LAST__' ? lastCol - 1 : fallback);
    });

    var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var counts = {};
    var records = [];
    for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
      var processed = rowProcessor(values[rowIndex], rowIndex, COL);
      if (!processed || processed.skip) continue;
      var status = processed.status || 'UNKNOWN';
      counts[status] = (counts[status] || 0) + 1;
      records.push(processed.record);
    }

    var result = {
      total: records.length,
      counts: counts,
      records: records,
      updated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy, hh:mm a')
    };
    putDashboardCache_(cache, cacheKey, result);
    return result;
  } catch (e) {
    return { error: e.message };
  }
}

function formatDate_(val) {
  if (Object.prototype.toString.call(val) !== '[object Date]') return String(val);
  var key = String(val.getTime());
  if (DASHBOARD_DATE_CACHE_.hasOwnProperty(key)) return DASHBOARD_DATE_CACHE_[key];
  if (DASHBOARD_DATE_CACHE_SIZE_ >= 512) {
    DASHBOARD_DATE_CACHE_ = {};
    DASHBOARD_DATE_CACHE_SIZE_ = 0;
  }
  var formatted = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  DASHBOARD_DATE_CACHE_[key] = formatted;
  DASHBOARD_DATE_CACHE_SIZE_++;
  return formatted;
}

function getMaterialInData() {
  return readSheetData_(SHEET_NAME_MATERIALIN, {
    sr: [["S.R", "SR NO", "S NO", "SR."], 0],
    item: [["ITEM"], 1],
    fromVendor: [["FROM"], 2],
    date: [["DATE"], 3],
    purpose: [["PURPOSE"], 4],
    materialDescription: [["MATERIAL DESCRIPTION", "DESCRIPTION"], 5],
    qty: [["QTY", "QUANTITY"], 6],
    receivedBy: [["RECEIVED BY"], 7],
    checkedBy: [["CHECKED BY"], 8],
    location: [["LOCATION"], 9],
    remarks: [["__NO_MATCH__"], "__LAST__"]
  }, function (row, idx, COL) {
    var item = row[COL.item] ? String(row[COL.item]).trim() : '';
    var sr = row[COL.sr] !== '' && row[COL.sr] !== undefined && row[COL.sr] !== null ? String(row[COL.sr]).trim() : '';
    if (!item && !sr) return { skip: true };

    var code = sr ? 'MI-' + sr : (item + '-' + (idx + 2));
    var rawRemarks = row[COL.remarks] ? String(row[COL.remarks]) : '';
    var status = normalizeMaterialInStatus_(rawRemarks);
    var materialDescription = row[COL.materialDescription] ? String(row[COL.materialDescription]) : '';

    return {
      status: status,
      record: {
        code: code,
        status: status,
        category: normalizeMaterialInCategory_(item, materialDescription),
        sr: sr,
        item: item,
        fromVendor: row[COL.fromVendor] ? String(row[COL.fromVendor]) : '',
        date: row[COL.date] ? formatDate_(row[COL.date]) : '',
        purpose: row[COL.purpose] ? String(row[COL.purpose]) : '',
        materialDescription: materialDescription,
        qty: row[COL.qty] !== '' && row[COL.qty] !== undefined && row[COL.qty] !== null ? String(row[COL.qty]) : '',
        receivedBy: row[COL.receivedBy] ? String(row[COL.receivedBy]) : '',
        checkedBy: row[COL.checkedBy] ? String(row[COL.checkedBy]) : '',
        location: row[COL.location] ? String(row[COL.location]) : '',
        remarks: rawRemarks
      }
    };
  }, null, 3000);
}
var ASSET_DETAILS_BRAND_RULES_=[{label:"HP",keywords:["HP "]},{label:"Lenovo",keywords:["LENOVO"]},{label:"Dell",keywords:["DELL"]},{label:"Apple / MacBook",keywords:["MACBOOK","MAC "]},{label:"Microsoft Surface",keywords:["MICROSOFT","SURFACE"]}];
function normalizeAssetDetailsBrand_(assetType){var text=(" "+String(assetType||"").toUpperCase()+" ");for(var i=0;i<ASSET_DETAILS_BRAND_RULES_.length;i++){var rule=ASSET_DETAILS_BRAND_RULES_[i];for(var j=0;j<rule.keywords.length;j++)if(text.indexOf(rule.keywords[j])>-1)return rule.label;}return"Other";}
function getAssetDetailsLaptopData(){return readSheetData_(SHEET_NAME_ASSET_DETAILS_LAPTOP,{assetCode:[["ASSET CODE"],0],assetType:[["ASSET TYPE"],1],serialNumber:[["SERIAL NUMBER","SERIAL NO"],2],macAddress:[["MAC ADDRESS"],3],purchaseDate:[["PURCHASE DATE"],4],purchasedFrom:[["PURCHASED FROM"],5],assetStatus:[["ASSET STATUS"],6],crmStatus:[["CRM STATUS"],7]},function(row,idx,COL){var code=row[COL.assetCode];if(!code)return{skip:true};var status=normalizeAssetDetailsStatus_(row[COL.assetStatus]),assetType=row[COL.assetType]?String(row[COL.assetType]):"";return{status:status,record:{code:String(code),status:status,category:normalizeAssetDetailsBrand_(assetType),assetType:assetType,serialNumber:row[COL.serialNumber]?String(row[COL.serialNumber]):"",macAddress:row[COL.macAddress]?String(row[COL.macAddress]):"",purchaseDate:row[COL.purchaseDate]?formatDate_(row[COL.purchaseDate]):"",purchasedFrom:row[COL.purchasedFrom]?String(row[COL.purchasedFrom]):"",crmStatus:row[COL.crmStatus]?String(row[COL.crmStatus]):""}};},ASSET_DETAILS_SPREADSHEET_ID);}


// ============================================================
// ASSET DETAILS — LAPTOP + DESKTOP ADD / UPDATE
// Both tabs live in the Asset Details workbook. Laptop keeps its
// existing 8 fields; Desktop follows the supplied Desktop sheet
// structure: Asset Code, Asset Type, Serial Number, MAC Address,
// Purchase Date, Purchased From, Asset Status and Emp Id.
// ============================================================
var ASSET_DETAILS_STATUS_OPTIONS_ = ['ACTIVE','ASSIGNED','NOT ASSIGNED','IN STOCK','GIVEN','ON_REPAIR','MISSING','SOLD','FAULTY','DAMAGED','WORKING','NOT WORKING','NEW / PURCHASED','DECOMMISSIONED'];

function assetDetailsDesktopSheet_() {
  var ss = SpreadsheetApp.openById(ASSET_DETAILS_SPREADSHEET_ID);
  var sheet = ss.getSheetById(ASSET_DETAILS_DESKTOP_SHEET_ID);
  if (!sheet) throw new Error('Desktop tab (gid ' + ASSET_DETAILS_DESKTOP_SHEET_ID + ') was not found in the Asset Details workbook.');
  return sheet;
}

function assetDetailsSheet_(sheetName, headersToCreate) {
  var ss = SpreadsheetApp.openById(ASSET_DETAILS_SPREADSHEET_ID);
  var sheet;
  // Desktop must ALWAYS use the existing tab supplied by the user. Never create a new tab.
  if (sheetName === SHEET_NAME_ASSET_DETAILS_DESKTOP) {
    sheet = assetDetailsDesktopSheet_();
  } else {
    sheet = findSheetLoose_(ss, sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1,1,1,headersToCreate.length).setValues([headersToCreate]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function assetDetailsColumns_(sheet, desktop) {
  var lastCol = Math.max(sheet.getLastColumn(), desktop ? 8 : 8);
  var headers = sheet.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(v){return String(v||'').trim().toUpperCase();});
  function col(patterns, fallback) {
    for (var p=0;p<patterns.length;p++) for (var h=0;h<headers.length;h++) if (headers[h]===patterns[p]) return h+1;
    for (var p2=0;p2<patterns.length;p2++) for (var h2=0;h2<headers.length;h2++) if (headers[h2].indexOf(patterns[p2])>-1) return h2+1;
    return fallback;
  }
  var c = {
    lastCol:lastCol,
    assetCode:col(['ASSET CODE'],1),
    assetType:col(['ASSET TYPE'],2),
    serialNumber:col(['SERIAL NUMBER','SERIAL NO'],3),
    macAddress:col(['MAC ADDRESS'],4),
    purchaseDate:col(['PURCHASE DATE'],5),
    purchasedFrom:col(['PURCHASED FROM'],6),
    assetStatus:col(['ASSET STATUS','STATUS'],7)
  };
  if (desktop) c.empId=col(['EMP ID','EMPLOYEE ID','EMPLOYEE CODE'],8);
  else c.crmStatus=col(['CRM STATUS'],8);
  return c;
}

function assetDetailsValue_(payload,key){
  return String(payload && payload[key] !== undefined && payload[key] !== null ? payload[key] : '').trim();
}
function assetDetailsDate_(value,label){
  var text=String(value||'').trim();
  if(!text) return '';
  var d=parseOfficeAssignmentDate_(text);
  if(!d) throw new Error('Please select a valid '+label+'.');
  return d;
}
function clearAssetDetailsCache_(sheetName){
  var c=CacheService.getScriptCache();
  var actualName = sheetName === SHEET_NAME_ASSET_DETAILS_DESKTOP ? assetDetailsDesktopSheet_().getName() : sheetName;
  c.remove('asset:v2:'+ASSET_DETAILS_SPREADSHEET_ID+':'+actualName);
  // Also clear the legacy Desktop cache key used by older dashboard versions.
  if (sheetName === SHEET_NAME_ASSET_DETAILS_DESKTOP) c.remove('asset:v2:'+ASSET_DETAILS_SPREADSHEET_ID+':'+SHEET_NAME_ASSET_DETAILS_DESKTOP);
}
function assetDetailsWrite_(sheet,row,c,values){
  Object.keys(values).forEach(function(k){
    if(c[k]===undefined || c[k]<1) return;
    var cell=sheet.getRange(row,c[k]),v=values[k];
    if(k==='purchaseDate'){
      if(v) cell.setValue(v).setNumberFormat('dd-mmm-yyyy');
      else cell.clearContent();
    } else {
      cell.setValue(v);
    }
  });
}
function assetDetailsDuplicateCheck_(sheet,c,assetCode,targetRow){
  var last=sheet.getLastRow();
  if(last<2) return;
  var vals=sheet.getRange(2,c.assetCode,last-1,1).getDisplayValues();
  for(var i=0;i<vals.length;i++){
    var row=i+2,existing=String(vals[i][0]||'').trim();
    if(row!==targetRow && existing && existing.toUpperCase()===assetCode.toUpperCase())
      throw new Error('Asset Code "'+assetCode+'" already exists in '+sheet.getName()+'.');
  }
}
function assetDetailsPayload_(payload,desktop){
  payload=payload||{};
  var p={
    assetCode:assetDetailsValue_(payload,'assetCode'),
    assetType:assetDetailsValue_(payload,'assetType'),
    serialNumber:assetDetailsValue_(payload,'serialNumber'),
    macAddress:assetDetailsValue_(payload,'macAddress').toUpperCase(),
    purchaseDate:assetDetailsDate_(payload.purchaseDate,'Purchase Date'),
    purchasedFrom:assetDetailsValue_(payload,'purchasedFrom'),
    assetStatus:assetDetailsValue_(payload,'assetStatus')||'ACTIVE'
  };
  if(desktop) p.empId=assetDetailsValue_(payload,'empId');
  else p.crmStatus=assetDetailsValue_(payload,'crmStatus');
  return p;
}
function addAssetDetails_(payload,desktop){
  var sheetName=desktop?SHEET_NAME_ASSET_DETAILS_DESKTOP:SHEET_NAME_ASSET_DETAILS_LAPTOP;
  var headers=desktop?['Asset Code','Asset type','Serial Number','MAC Address','Purchase Date','Purchased From','Asset Status','Emp Id']:['Asset Code','Asset type','Serial Number','MAC Address','Purchase Date','Purchased From','Asset Status','CRM Status'];
  var sheet=assetDetailsSheet_(sheetName,headers),c=assetDetailsColumns_(sheet,desktop),p=assetDetailsPayload_(payload,desktop),lock=LockService.getDocumentLock();
  if(!p.assetCode) throw new Error('Asset Code is required.');
  lock.waitLock(30000);
  try{
    assetDetailsDuplicateCheck_(sheet,c,p.assetCode,0);
    var row=sheet.getLastRow()+1;
    // Keep the existing Desktop tab's row formatting/data validation.
    // Copy the previous data row's formatting only, then write the new values.
    if (row > 2 && sheet.getLastColumn() > 0) {
      var sourceRow = row - 1;
      sheet.getRange(sourceRow,1,1,sheet.getLastColumn()).copyTo(
        sheet.getRange(row,1,1,sheet.getLastColumn()),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false
      );
      sheet.getRange(sourceRow,1,1,sheet.getLastColumn()).copyTo(
        sheet.getRange(row,1,1,sheet.getLastColumn()),
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false
      );
    }
    assetDetailsWrite_(sheet,row,c,p);
    auditDashboardRow_(sheet,row,'ASSET DETAILS CREATED');
    clearAssetDetailsCache_(sheetName);
    return {ok:true,assetCode:p.assetCode,status:p.assetStatus};
  }finally{lock.releaseLock();}
}
function updateAssetDetails_(payload,desktop){
  var sheetName=desktop?SHEET_NAME_ASSET_DETAILS_DESKTOP:SHEET_NAME_ASSET_DETAILS_LAPTOP;
  var sheet=assetDetailsSheet_(sheetName,desktop?['Asset Code','Asset type','Serial Number','MAC Address','Purchase Date','Purchased From','Asset Status','Emp Id']:['Asset Code','Asset type','Serial Number','MAC Address','Purchase Date','Purchased From','Asset Status','CRM Status']);
  var c=assetDetailsColumns_(sheet,desktop),p=assetDetailsPayload_(payload,desktop),original=assetDetailsValue_(payload,'originalAssetCode')||p.assetCode;
  if(!original) throw new Error('Original Asset Code is required.');
  if(!p.assetCode) throw new Error('Asset Code is required.');
  var lock=LockService.getDocumentLock();lock.waitLock(30000);
  try{
    var last=sheet.getLastRow(),target=0;
    if(last>=2){var vals=sheet.getRange(2,c.assetCode,last-1,1).getDisplayValues();for(var i=0;i<vals.length;i++)if(String(vals[i][0]||'').trim().toUpperCase()===original.toUpperCase()){target=i+2;break;}}
    if(!target) throw new Error('Asset Code "'+original+'" was not found in '+sheetName+'.');
    assetDetailsDuplicateCheck_(sheet,c,p.assetCode,target);
    var before=activitySnapshot_(sheet,target);
    assetDetailsWrite_(sheet,target,c,p);
    auditDashboardRow_(sheet,target,'ASSET DETAILS UPDATED',before);
    clearAssetDetailsCache_(sheetName);
    return {ok:true,assetCode:p.assetCode,status:p.assetStatus};
  }finally{lock.releaseLock();}
}
function addAssetDetailsLaptop(payload){return addAssetDetails_(payload,false);}
function updateAssetDetailsLaptop(payload){return updateAssetDetails_(payload,false);}
function addAssetDetailsDesktop(payload){return addAssetDetails_(payload,true);}
function updateAssetDetailsDesktop(payload){return updateAssetDetails_(payload,true);}

var ASSET_DETAILS_DESKTOP_SEED_ = [
  ['2018PC001','DESKTOP-I7-10TH-16GB-700GB-4TB-2TB-512GB-4GB','DC:1B:A1:A0:E1:F0','', '', 'SMART TECHNO','Akshat Saini (Maharani Farm)','CE/2024/02/714'],
  ['2021PC002','HP-24EA-I7-10TH 32GB-2TB SSD 500GB4.,4GB GPU','MS-7C83','DC:1B:A1:9F:C1:8A','', 'SMART TECHNO','Nonit Agrawal (Maharani Farm)','CE/2021/06/179'],
  ['2021PC003','DESKTOP-I7-10TH-16GB,,-4TB,,4TB,, -512GB-4GB','E8:84:A5:CC:9A:7A','', '', 'SMART TECHNO','TEENA KHANDELWAL','CE/2024/01/633'],
  ['2023PC009','DESKTOP-PCi7-12TH-16GB-2TB-512GBSSD,6GB','A09C755A4E3F','F0:A6:54:C5:00:A5','08-02-2023','SMART TECHNO','PUNIT PAREEK','CE/2022/12/474'],
  ['2024PC010','DESKTOP-PC-i7-13TH-32GB-1TBSSD,4GB','PB4H273500786','30:F6:EF:19:85:39','12-01-2024','SMART TECHNO','Akshat Saini (IT Park)','CE/2024/02/714'],
  ['2024PC011','Desktop-PC-i7-13TH,,32GB-1TB-8GB','30f6ef189084','30:f6:ef:18:90:84','12-01-2024','Smart Techno','Nonit Agrawal (IT Park)','CE/2021/06/179'],
  ['2019PC004','PC-ViewSonic-i5-10G-64GB-512GB-2GB','18CC182D5D61','18:CC:18:2D:5D:61','19-03-2019','SMART TECHNO','PROXMOX SERVER','IT'],
  ['2019PC005','DESKTOP-I5-8G-16GB-256-GB','5294AD331CB1','0A:00:27:00:00:17','19-03-2019','SMART TECHNO','DEEPAK KUMAR JANGID (Server Room)','IT'],
  ['2019PC006','DESKTOP-I5-10G-16GB-512GB-2GB','5294AD331CC3','E8:65:D4:18:B9:A6','19-03-2019','SMART TECHNO','2nd Floor Cabin (EXAM CERTI.)','IT'],
  ['2022PC007','DESKTOP-ANTESPORT-I5-11G-16GB-512GB','879D7DA2','C4:BD:E5:33:66:CA','07-06-2022','SMART TECHNO','TALLY SERVER (JASMEET SETHY)','CE/2022/06/412'],
  ['2022PC008','DESKTOP-i9-12TH-32GB-1TB-8GB','652B8050551F','2C:0D:A7:26:70:28','19-10-2022','SMART TECHNO','Studio IT Park (WEBINAR PC)','IT']
];

function seedAssetDetailsDesktopIfEmpty_(sheet){
  if(sheet.getLastRow()>1) return;
  var rows=ASSET_DETAILS_DESKTOP_SEED_.map(function(r){
    var d=r[4] ? parseOfficeAssignmentDate_(r[4]) : '';
    return [r[0],r[1],r[2],String(r[3]||'').toUpperCase(),d,r[5], 'ACTIVE', r[7]];
  });
  sheet.getRange(2,1,rows.length,8).setValues(rows);
  sheet.getRange(2,5,rows.length,1).setNumberFormat('dd-mmm-yyyy');
}

function getAssetDetailsDesktopData(){
  var sheet=assetDetailsSheet_(SHEET_NAME_ASSET_DETAILS_DESKTOP,['Asset Code','Asset type','Serial Number','MAC Address','Purchase Date','Purchased From','Asset Status','Emp Id']);
  return readSheetData_(assetDetailsDesktopSheet_().getName(),{assetCode:[['ASSET CODE'],0],assetType:[['ASSET TYPE'],1],serialNumber:[['SERIAL NUMBER','SERIAL NO'],2],macAddress:[['MAC ADDRESS'],3],purchaseDate:[['PURCHASE DATE'],4],purchasedFrom:[['PURCHASED FROM'],5],assetStatus:[['ASSET STATUS','STATUS'],6],empId:[['EMP ID','EMPLOYEE ID','EMPLOYEE CODE'],7]},function(row,idx,COL){
    var code=row[COL.assetCode]; if(!code)return{skip:true};
    var assetType=row[COL.assetType]?String(row[COL.assetType]):'';
    var status=normalizeAssetDetailsStatus_(row[COL.assetStatus]);
    return {status:status,record:{code:String(code),status:status,category:'Desktop',assetType:assetType,serialNumber:row[COL.serialNumber]?String(row[COL.serialNumber]):'',macAddress:row[COL.macAddress]?String(row[COL.macAddress]):'',purchaseDate:row[COL.purchaseDate]?formatDate_(row[COL.purchaseDate]):'',purchasedFrom:row[COL.purchasedFrom]?String(row[COL.purchasedFrom]):'',empId:row[COL.empId]?String(row[COL.empId]):''}};
  },ASSET_DETAILS_SPREADSHEET_ID);
}


// ============================================================
// AI-POWERED IT ASSET ASSISTANT — V3
// Deterministic asset intelligence first; AI is used only to explain
// verified results. Current data + complete Activity_Log history are
// searched separately so historical questions never depend on the
// current employee value alone.
// ============================================================
var ASSET_ASSISTANT_OPENAI_KEY='ASSET_ASSISTANT_OPENAI_API_KEY';
var ASSET_ASSISTANT_MODEL='gpt-5.6-luna';

function assetAssistantSetApiKey(key){
  if(!isDashboardAdmin_(getCurrentDashboardUserEmail_()))throw new Error('Only administrator can configure the AI assistant.');
  key=String(key||'').trim(); if(!key)throw new Error('Please enter a valid OpenAI API key.');
  PropertiesService.getScriptProperties().setProperty(ASSET_ASSISTANT_OPENAI_KEY,key); return{ok:true,configured:true};
}
function assetAssistantHasApiKey(){
  if(!isDashboardAdmin_(getCurrentDashboardUserEmail_()))throw new Error('Only administrator can check AI configuration.');
  return{configured:!!PropertiesService.getScriptProperties().getProperty(ASSET_ASSISTANT_OPENAI_KEY)};
}
function assetAssistantApiKey_(){return PropertiesService.getScriptProperties().getProperty(ASSET_ASSISTANT_OPENAI_KEY)||'';}
function assetAssistantText_(v){return String(v===undefined||v===null?'':v).trim();}
function assetAssistantNorm_(v){return assetAssistantText_(v).toUpperCase().replace(/[^A-Z0-9]/g,'');}
function assetAssistantDateText_(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd-MMM-yyyy hh:mm a');}
function assetAssistantIsUnassigned_(r){
  var status=assetAssistantNorm_(r.status),emp=assetAssistantText_(r.employee),id=assetAssistantText_(r.empId);
  return (!emp&&!id)||/NOTASSIGNED|UNASSIGNED|UNALLOCATED|AVAILABLE|INSTOCK|IN STOCK|FREE/.test(status);
}
function assetAssistantLaptop_(r){
  var t=(r.type+' '+r.category+' '+r.section+' '+r.assetType).toUpperCase();
  return /LAPTOP|MACBOOK|NOTEBOOK/.test(t);
}
function assetAssistantPushRecords_(all,data,source,section){
  if(!data||!Array.isArray(data.records))return;
  data.records.forEach(function(r){
    var employee=r.empName||r.userName||r.employeeName||'';
    var code=r.assetCode||r.code||'';
    all.push({source:source,section:section,code:assetAssistantText_(code),type:assetAssistantText_(r.assetType||r.category||r.assetName||r.name||section),category:assetAssistantText_(r.category),employee:assetAssistantText_(employee),empId:assetAssistantText_(r.empId||r.employeeId||r.employeeCode),status:assetAssistantText_(r.status||r.assetStatus),mac:assetAssistantText_(r.macAddress||r.mac||''),otherMac:assetAssistantText_(r.otherMac||''),serial:assetAssistantText_(r.serialNumber||r.serialNo||''),date:assetAssistantText_(r.date||r.assignedDate||r.assignDate||r.purchaseDate||r.dispatchDate||r.checkedDate||''),text:assetAssistantText_([r.clientName,r.project,r.configuration,r.makeModel,r.remarks,r.location,r.purchasedFrom,r.assetName].join(' '))});
  });
}
function getAssetAssistantSnapshot_(){
  requireDashboardAccess_(); var all=[];
  assetAssistantPushRecords_(all,getDashboardData(),'Asset Register','Office Laptops');
  assetAssistantPushRecords_(all,getClientLaptopsData(),'Asset Register','Client Laptops');
  assetAssistantPushRecords_(all,getOtherAssetsData(),'Asset Register','Other Assets (IT)');
  assetAssistantPushRecords_(all,getMobileAccessoriesData(),'Asset Register','Mobile & Accessories');
  assetAssistantPushRecords_(all,getStudioItemsData(),'Asset Register','Studio Items');
  assetAssistantPushRecords_(all,getCourierData(),'Asset Register','Courier Data');
  assetAssistantPushRecords_(all,getMaterialInData(),'Asset Register','Material IN');
  assetAssistantPushRecords_(all,getMacOfficeAssetData(),'MAC Address','MAC — All Devices');
  assetAssistantPushRecords_(all,getMacClientLaptopData(),'MAC Address','Client Laptop');
  assetAssistantPushRecords_(all,getMacFirewallData(),'MAC Address','Adobe User on Firewall');
  assetAssistantPushRecords_(all,getAssetDetailsLaptopData(),'Asset Details','Laptops');
  assetAssistantPushRecords_(all,getAssetDetailsDesktopData(),'Asset Details','Desktop');
  return all;
}
function assetAssistantRow_(r){return{code:r.code||'',type:r.type||'',employee:r.employee||'',empId:r.empId||'',status:r.status||'',mac:r.mac||r.otherMac||'',serial:r.serial||'',source:r.source||'',section:r.section||'',date:r.date||''};}
function assetAssistantNameKey_(v){return assetAssistantText_(v).toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}
function assetAssistantFindPerson_(q,all){
  var text=assetAssistantText_(q),low=text.toLowerCase();
  var m=low.match(/(?:assigned\s+(?:to|for)|show\s+(?:the\s+)?mac(?:\s+address(?:es)?)?\s+(?:to|for)|mac(?:\s+address)?\s+(?:to|for)|for|to|with|used\s+by|belonging\s+to|of)\s+([a-z][a-z .'-]{1,60}?)(?=\s+(?:how|what|which|who|when|and|that|this|has|have|did|was|were|is|are|change|changed|update|updated|asset|assets|mac|address|history|so|far|ab|tak|kub|kisko|kisne|kitne)\b|[?.!,]|$)/i);
  if(m&&m[1])return m[1].trim();
  var people={};all.forEach(function(r){if(r.employee)people[r.employee.trim()]=true;});
  var names=Object.keys(people).filter(function(n){return n.length>1;}).sort(function(a,b){return b.length-a.length;});
  for(var i=0;i<names.length;i++)if(low.indexOf(names[i].toLowerCase())>-1)return names[i];
  // Also allow a first/last-name fragment from a full employee name.
  for(var j=0;j<names.length;j++){
    var parts=names[j].split(/\s+/).filter(function(x){return x.length>=3;});
    for(var k=0;k<parts.length;k++)if(low.indexOf(parts[k].toLowerCase())>-1)return names[j];
  }
  return '';
}
function assetAssistantPersonMatch_(r,person){
  var p=assetAssistantNameKey_(person);if(!p)return false;
  var h=assetAssistantNameKey_((r.employee||'')+' '+(r.empId||''));
  if(h.indexOf(p)>-1)return true;
  var words=p.split(/\s+/).filter(function(x){return x.length>=3;});
  return words.length>0&&words.every(function(w){return h.indexOf(w)>-1;});
}
function assetAssistantAssetCode_(q){
  var m=String(q||'').match(/(?:asset|code|record|laptop|desktop|pc)\s*#?\s*([a-z0-9][a-z0-9_-]*)/i);
  return m?m[1].trim():'';
}
function assetAssistantExactAsset_(all,code){
  if(!code)return[];var n=assetAssistantNorm_(code);
  return all.filter(function(r){return assetAssistantNorm_(r.code)===n||assetAssistantNorm_(r.code).indexOf(n)>-1;});
}
function assetAssistantActivity_(){
  var sheet=getActivityLogSheet_(),values=sheet.getDataRange().getValues();if(values.length<2)return[];
  var h=values[0].map(function(v){return assetAssistantText_(v).toUpperCase();}),idx={};
  ['TIMESTAMP','USER','SOURCE WORKBOOK','WORKSPACE','TAB','ACTION','RECORD ID','EMPLOYEE ID','EMPLOYEE NAME','FIELD','OLD VALUE','NEW VALUE','ROW'].forEach(function(n,i){var x=h.indexOf(n);idx[n]=x>-1?x:i;});
  return values.slice(1).map(function(r){var d=r[idx.TIMESTAMP] instanceof Date?r[idx.TIMESTAMP]:new Date(r[idx.TIMESTAMP]);return{date:d,ts:d.getTime(),user:assetAssistantText_(r[idx.USER]),sourceWorkbook:assetAssistantText_(r[idx['SOURCE WORKBOOK']]),workspace:assetAssistantText_(r[idx.WORKSPACE]),tab:assetAssistantText_(r[idx.TAB]),action:assetAssistantText_(r[idx.ACTION]),recordId:assetAssistantText_(r[idx['RECORD ID']]),empId:assetAssistantText_(r[idx['EMPLOYEE ID']]),empName:assetAssistantText_(r[idx['EMPLOYEE NAME']]),field:assetAssistantText_(r[idx.FIELD]),oldValue:assetAssistantText_(r[idx['OLD VALUE']]),newValue:assetAssistantText_(r[idx['NEW VALUE']]),row:assetAssistantText_(r[idx.ROW])};}).filter(function(x){return !isNaN(x.ts);});
}
function assetAssistantActivityRecord_(a){return{date:assetAssistantDateText_(a.date),timestampMs:a.ts,user:a.user,workspace:a.workspace,tab:a.tab,action:a.action,recordId:a.recordId,empId:a.empId,empName:a.empName,field:a.field,oldValue:a.oldValue,newValue:a.newValue,row:a.row};}
function assetAssistantThisWeekStart_(){var now=new Date(),day=now.getDay(),diff=day===0?6:day-1,start=new Date(now);start.setDate(now.getDate()-diff);start.setHours(0,0,0,0);return start.getTime();}
function assetAssistantMonthRange_(){var now=new Date(),start=new Date(now.getFullYear(),now.getMonth(),1),end=new Date(now.getFullYear(),now.getMonth()+1,0,23,59,59,999);return{from:start.getTime(),to:end.getTime(),label:Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM')};}
function assetAssistantExtractMonth_(low){
  var months=['january','february','march','april','may','june','july','august','september','october','november','december'];
  var m=low.match(/(january|february|march|april|may|june|july|august|september|october|november|december)\s+(20\d{2})/);
  if(!m)return'';return m[2]+'-'+String(months.indexOf(m[1])+1).padStart(2,'0');
}
function assetAssistantEventText_(a){return(a.action+' '+a.field+' '+a.oldValue+' '+a.newValue).toUpperCase();}
function assetAssistantIsMacEvent_(a){return /MAC\s*(ADDRESS)?/.test(assetAssistantEventText_(a));}
function assetAssistantIsAssignmentEvent_(a){
  var z=assetAssistantEventText_(a);
  return /ASSIGN|ALLOCAT|HANDOVER|HAND[- ]?OVER|EMPLOYEE\s*(NAME|ID)|EMP\s*(NAME|ID)|ASSIGNED\s*TO|USER\s*NAME/.test(z);
}
function assetAssistantIsTransferEvent_(a){
  var z=assetAssistantEventText_(a);if(/TRANSFER|MOVED|MOVEMENT|HANDOVER|HAND[- ]?OVER/.test(z))return true;
  if(!assetAssistantIsAssignmentEvent_(a))return false;
  // A change from one employee value to another is a movement/transfer.
  return assetAssistantText_(a.oldValue)!==''&&assetAssistantText_(a.newValue)!==''&&assetAssistantNameKey_(a.oldValue)!==assetAssistantNameKey_(a.newValue);
}
function assetAssistantHistoryForPerson_(all,activity,person){
  var codes={};all.forEach(function(r){if(assetAssistantPersonMatch_(r,person)&&r.code)codes[assetAssistantNorm_(r.code)]=r.code;});
  return activity.filter(function(a){
    if(!assetAssistantIsAssignmentEvent_(a))return false;
    if(a.empName&&assetAssistantPersonMatch_({employee:a.empName,empId:a.empId},person))return true;
    if(assetAssistantNameKey_(a.oldValue+' '+a.newValue).indexOf(assetAssistantNameKey_(person))>-1)return true;
    return !!codes[assetAssistantNorm_(a.recordId)];
  });
}
function assetAssistantFacts_(question,all){
  var q=assetAssistantText_(question),low=q.toLowerCase(),facts={intent:'general',question:q,rows:[],activity:[],summary:'',notes:[]};
  var codeOnly=(q.match(/\b\d{3,8}\b/)||[])[0]||'';
  var explicitEmp=/employee\s*(?:code|id)|emp\s*(?:code|id)|employee\s*number|emp\s*number/.test(low);

  // Reports are commands, not searches.
  if(/\b(download|export|generate|create|save)\b/.test(low)&&/(log|logs|activity|audit|report)/.test(low)){
    facts.intent='download_activity';var range=assetAssistantMonthRange_();
    if(/this\s+week|weekly|week|is\s+week/.test(low)){facts.reportMode='range';facts.reportFrom=assetAssistantThisWeekStart_();facts.reportTo=Date.now();facts.reportValue=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');facts.summary='Download this week Activity Logs report.';}
    else if(/today|daily|day|aaj/.test(low)){facts.reportMode='daily';facts.reportValue=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');facts.summary='Download today Activity Logs report.';}
    else {var month=assetAssistantExtractMonth_(low);facts.reportMode='monthly';facts.reportValue=month||range.label;facts.summary='Download Activity Logs for '+facts.reportValue+'.';}
    return facts;
  }
  // Explicit employee code must win over generic numeric asset searches.
  if(explicitEmp){
    facts.intent='employee_code';facts.employeeCode=codeOnly;facts.rows=all.filter(function(r){return assetAssistantText_(r.empId)===codeOnly;}).map(assetAssistantRow_);facts.summary=facts.rows.length+' current record(s) found for employee code '+codeOnly+'.';return facts;
  }

  if(/warranty/.test(low)){facts.intent='warranty';facts.summary='Warranty expiry cannot be determined because the current dashboard data does not expose a warranty-expiry field.';return facts;}

  var person=assetAssistantFindPerson_(q,all),assetCode=assetAssistantAssetCode_(q)||'';
  var asksHistory=/(history|historical|ever|so far|till now|until now|ab tak|aaj tak|kub|kab|when|who|whom|kisko|kisne|change|changed|update|updated|modified|badla|assign|assigned|assignment|allocate|allocated|handover|transferred|transfer)/.test(low);

  // MAC questions: current MAC and MAC-change history are completely separate.
  if(/mac/.test(low)&&asksHistory){
    facts.intent='mac_history';facts.person=person;facts.assetCode=assetCode;
    var macAct=assetAssistantActivity_().filter(assetAssistantIsMacEvent_);
    if(assetCode)macAct=macAct.filter(function(a){return assetAssistantNorm_(a.recordId)===assetAssistantNorm_(assetCode)||assetAssistantNorm_(a.recordId).indexOf(assetAssistantNorm_(assetCode))>-1;});
    else if(person){
      var personAssets={};all.forEach(function(r){if(assetAssistantPersonMatch_(r,person)&&r.code)personAssets[assetAssistantNorm_(r.code)]=true;});
      macAct=macAct.filter(function(a){return (a.empName&&assetAssistantPersonMatch_({employee:a.empName,empId:a.empId},person))||personAssets[assetAssistantNorm_(a.recordId)]||assetAssistantNameKey_(a.oldValue+' '+a.newValue).indexOf(assetAssistantNameKey_(person))>-1;});
    }
    facts.activity=macAct.sort(function(a,b){return b.ts-a.ts;}).map(assetAssistantActivityRecord_);facts.summary=facts.activity.length+' MAC history record(s) found.';return facts;
  }
  if(/mac/.test(low)&&person){
    facts.intent='mac_for_person';facts.person=person;facts.rows=all.filter(function(r){return assetAssistantPersonMatch_(r,person)&&(r.mac||r.otherMac);}).map(assetAssistantRow_);facts.summary=facts.rows.length+' current record(s) with MAC address found for '+person+'.';return facts;
  }
  if(/mac/.test(low)&&assetCode&&!asksHistory){
    facts.intent='mac_for_asset';facts.assetCode=assetCode;facts.rows=assetAssistantExactAsset_(all,assetCode).filter(function(r){return r.mac||r.otherMac;}).map(assetAssistantRow_);facts.summary=facts.rows.length+' current MAC record(s) found for asset '+assetCode+'.';return facts;
  }

  // Assignment questions. Current assignment and historical assignment are separate.
  if(/assign|assigned|assignment|allocate|allocated|handover|handed over|kisko|kisne|ab tak|aaj tak|so far|till now|until now/.test(low)&&asksHistory){
    var act=assetAssistantActivity_();facts.activity=act.filter(assetAssistantIsAssignmentEvent_);
    if(assetCode)facts.activity=facts.activity.filter(function(a){return assetAssistantNorm_(a.recordId)===assetAssistantNorm_(assetCode)||assetAssistantNorm_(a.recordId).indexOf(assetAssistantNorm_(assetCode))>-1;});
    else if(person)facts.activity=assetAssistantHistoryForPerson_(all,act,person);
    facts.activity=facts.activity.sort(function(a,b){return b.ts-a.ts;}).map(assetAssistantActivityRecord_);
    facts.intent=assetCode?'assignment_history_asset':(person?'assignment_history_person':'assignment_history');facts.person=person;facts.assetCode=assetCode;
    facts.summary=facts.activity.length+' historical assignment record(s) found.';return facts;
  }

  if(/duplicate\s*mac|mac.*duplicate|same\s+mac/.test(low)){
    facts.intent='duplicate_mac';var groups={};
    all.forEach(function(r){[r.mac,r.otherMac].forEach(function(mac){var key=assetAssistantNorm_(mac);if(!key||key.length<6)return;if(!groups[key])groups[key]=[];var sig=(r.source+'|'+r.section+'|'+r.code+'|'+r.employee).toUpperCase();if(!groups[key].some(function(x){return x.sig===sig;}))groups[key].push({sig:sig,record:assetAssistantRow_(r)});});});
    Object.keys(groups).forEach(function(mac){var distinct={};groups[mac].forEach(function(x){distinct[x.record.code+'|'+x.record.employee+'|'+x.record.source+'|'+x.record.section]=true;});if(Object.keys(distinct).length>1)facts.rows.push({mac:mac,entries:groups[mac].map(function(x){return x.record;})});});
    facts.summary=facts.rows.length+' duplicate MAC group(s) found.';return facts;
  }
  if(/transfer|transferred|movement|moved|handover|handed over/.test(low)){
    facts.intent='transfers';var st=assetAssistantThisWeekStart_(),now=Date.now();facts.activity=assetAssistantActivity_().filter(function(a){return a.ts>=st&&a.ts<=now&&assetAssistantIsTransferEvent_(a);}).sort(function(a,b){return b.ts-a.ts;}).map(assetAssistantActivityRecord_);facts.summary=facts.activity.length+' transfer/movement record(s) found this week.';return facts;
  }
  if(/unassign|not assigned|unallocated|available laptop|free laptop/.test(low)){
    facts.intent='unassigned_laptops';var laptops=all.filter(assetAssistantLaptop_);facts.rows=laptops.filter(assetAssistantIsUnassigned_).map(assetAssistantRow_);facts.summary=facts.rows.length+' laptop record(s) currently appear unassigned.';return facts;
  }
  if(person&&/(assigned to|assigned for|assets?\s+(?:of|for)|with|belonging|used by|using|show.*mac|mac)/.test(low)){
    facts.intent='assigned_to';facts.person=person;facts.rows=all.filter(function(r){return assetAssistantPersonMatch_(r,person);}).map(assetAssistantRow_);facts.summary=facts.rows.length+' current asset record(s) found for '+person+'.';return facts;
  }
  if(assetCode){
    facts.intent='asset_lookup';facts.assetCode=assetCode;facts.rows=assetAssistantExactAsset_(all,assetCode).map(assetAssistantRow_);facts.summary=facts.rows.length+' current record(s) found for asset '+assetCode+'.';return facts;
  }

  // Broad search: any meaningful token may match; then rank by exact code/name/MAC.
  facts.intent='search';var stop=/^(show|find|give|tell|what|which|how|many|are|the|for|from|this|that|with|and|all|assets?|data|please|current|currently|can|you|me|is|to|of|a|an|dashboard|address|details?)$/;
  var tokens=low.replace(/[^a-z0-9:.-]+/g,' ').split(/\s+/).filter(function(t){return t.length>=2&&!stop.test(t);});
  var scored=all.map(function(r){var hay=(r.code+' '+r.type+' '+r.category+' '+r.employee+' '+r.empId+' '+r.status+' '+r.mac+' '+r.otherMac+' '+r.serial+' '+r.source+' '+r.section+' '+r.text).toLowerCase(),score=0;tokens.forEach(function(t){if(hay.indexOf(t)>-1)score+=1;if(String(r.code).toLowerCase()===t)score+=10;});return{r:r,score:score};}).filter(function(x){return x.score>0;}).sort(function(a,b){return b.score-a.score;});
  facts.rows=scored.slice(0,100).map(function(x){return assetAssistantRow_(x.r);});facts.summary=facts.rows.length+' matching record(s) found.';return facts;
}
function assetAssistantFallbackAnswer_(facts){
  if(facts.intent==='warranty')return 'I cannot determine warranty expiry because the current dashboard data does not contain a warranty-expiry field.';
  if(facts.intent==='unassigned_laptops')return facts.rows.length+' laptop(s) are currently unassigned.';
  if(facts.intent==='assigned_to')return facts.rows.length+' current asset record(s) are assigned to '+facts.person+'.';
  if(facts.intent==='employee_code')return facts.rows.length+' current record(s) were found for employee code '+facts.employeeCode+'.';
  if(facts.intent==='mac_for_person')return facts.rows.length+' current record(s) with MAC address were found for '+facts.person+'.';
  if(facts.intent==='mac_for_asset')return facts.rows.length+' current MAC record(s) were found for asset '+facts.assetCode+'.';
  if(facts.intent==='mac_history')return facts.activity.length+' MAC history record(s) were found.';
  if(/^assignment_history/.test(facts.intent))return facts.activity.length+' historical assignment record(s) were found'+(facts.person?' for '+facts.person:'')+(facts.assetCode?' for asset '+facts.assetCode:'')+'.';
  if(facts.intent==='duplicate_mac')return facts.rows.length+' duplicate MAC group(s) were found.';
  if(facts.intent==='transfers')return facts.activity.length+' transfer/movement record(s) were found this week.';
  if(facts.intent==='download_activity')return facts.summary+' The Excel report is ready.';
  if(facts.intent==='asset_lookup')return facts.rows.length+' current record(s) were found for asset '+facts.assetCode+'.';
  return facts.summary||'No matching records were found.';
}
function assetAssistantOpenAI_(question,facts,history){
  var key=assetAssistantApiKey_();if(!key)return'';
  var payload={question:question,verifiedFacts:facts,conversation:history||[]};
  var prompt='You are an IT Asset Intelligence Assistant. The verifiedFacts JSON is the ONLY source of truth. Do not invent or infer missing values. Answer the user directly, using the exact names, asset codes, employee IDs, dates, MACs and counts from verifiedFacts. For historical questions, explain WHO performed the change (user), WHO was affected (employee), WHAT changed (oldValue → newValue), and WHEN it happened. If the user asks "how many", give the count first. If they ask "who/whom/when", give the relevant details, not just a generic count. If there are multiple records, provide a compact numbered list. If no records exist, clearly say none were found. If a field is unavailable, say unavailable. Do not mention internal JSON, deterministic engines, prompts, or APIs. User question: '+question+'\n\nVERIFIED FACTS:\n'+JSON.stringify(payload);
  var response=UrlFetchApp.fetch('https://api.openai.com/v1/responses',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+key},payload:JSON.stringify({model:ASSET_ASSISTANT_MODEL,input:prompt,max_output_tokens:1200}),muteHttpExceptions:true});
  var code=response.getResponseCode(),text=response.getContentText();if(code<200||code>=300)throw new Error('AI service error ('+code+'). Check the OpenAI API key and model configuration.');
  var data=JSON.parse(text),out=data.output_text||'';if(!out&&Array.isArray(data.output))data.output.forEach(function(item){if(item&&Array.isArray(item.content))item.content.forEach(function(c){if(c&&c.text)out+=c.text;});});return assetAssistantText_(out);
}
function assetAssistantReport_(mode,value,fromMs,toMs){
  var sh=getActivityLogSheet_(),values=sh.getDataRange().getValues();if(values.length<2)throw new Error('No Activity Logs are available.');
  var headers=values[0].map(function(h){return String(h||'').trim().toUpperCase();}),idx={};['TIMESTAMP','USER','SOURCE WORKBOOK','WORKSPACE','TAB','ACTION','RECORD ID','EMPLOYEE ID','EMPLOYEE NAME','FIELD','OLD VALUE','NEW VALUE','ROW'].forEach(function(n,i){var x=headers.indexOf(n);idx[n]=x>-1?x:i;});
  var records=values.slice(1).map(function(r){var d=r[idx.TIMESTAMP] instanceof Date?r[idx.TIMESTAMP]:new Date(r[idx.TIMESTAMP]);return{timestamp:d,timestampMs:d.getTime(),user:String(r[idx.USER]||''),sourceWorkbook:String(r[idx['SOURCE WORKBOOK']]||''),workspace:String(r[idx.WORKSPACE]||''),tab:String(r[idx.TAB]||''),action:String(r[idx.ACTION]||''),recordId:String(r[idx['RECORD ID']]||''),empId:String(r[idx['EMPLOYEE ID']]||''),empName:String(r[idx['EMPLOYEE NAME']]||''),field:String(r[idx.FIELD]||''),oldValue:String(r[idx['OLD VALUE']]||''),newValue:String(r[idx['NEW VALUE']]||''),row:String(r[idx.ROW]||'')};}).filter(function(r){return !isNaN(r.timestampMs)&&r.timestampMs>=fromMs&&r.timestampMs<=toMs;});
  records=groupActivityProcessRecords_(records).sort(function(a,b){return b.timestampMs-a.timestampMs;});if(!records.length)throw new Error('No Activity Logs found for the requested period.');
  var temp=null;try{temp=SpreadsheetApp.create('Activity Log Report - '+value);var out=temp.getSheets()[0],hs=['Timestamp','User','Source Workbook','Workspace','Tab','Action','Record ID','Employee ID','Employee Name','Field','Old Value','New Value','Row'],rows=records.map(function(r){return[r.timestamp,r.user,r.sourceWorkbook,r.workspace,r.tab,r.action,r.recordId,r.empId,r.empName,r.field,r.oldValue,r.newValue,r.row];});out.setName('Activity_Log_Report');out.getRange(1,1,1,hs.length).setValues([hs]);out.getRange(2,1,rows.length,hs.length).setValues(rows);out.getRange(2,1,rows.length,1).setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');out.setColumnWidth(1,190);out.setFrozenRows(1);out.getRange(1,1,1,hs.length).setFontWeight('bold');out.autoResizeColumns(1,hs.length);SpreadsheetApp.flush();var url='https://docs.google.com/spreadsheets/d/'+temp.getId()+'/export?format=xlsx',resp=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});if(resp.getResponseCode()!==200)throw new Error('Could not create the Excel report.');var blob=resp.getBlob().setName('Activity_Logs_'+mode+'_'+value+'.xlsx');return{filename:blob.getName(),mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',base64:Utilities.base64Encode(blob.getBytes()),count:records.length};}finally{if(temp){try{DriveApp.getFileById(temp.getId()).setTrashed(true);}catch(ignore){}}}
}
function askAssetAssistant(question,history){
  requireDashboardAccess_();question=assetAssistantText_(question);if(!question)throw new Error('Please enter a question.');
  var all=getAssetAssistantSnapshot_(),facts=assetAssistantFacts_(question,all),answer='',download=null;
  try{
    if(facts.intent==='download_activity'&&facts.reportMode==='daily')download=downloadActivityLogReport({mode:'daily',value:facts.reportValue});
    else if(facts.intent==='download_activity'&&facts.reportMode==='monthly')download=downloadActivityLogReport({mode:'monthly',value:facts.reportValue});
    else if(facts.intent==='download_activity'&&facts.reportMode==='range')download=assetAssistantReport_('week',facts.reportValue,facts.reportFrom,facts.reportTo);
    answer=assetAssistantOpenAI_(question,facts,history||[]);
  }catch(e){answer='';}
  if(!answer)answer=assetAssistantFallbackAnswer_(facts);
  var rows=[];
  if(facts.intent==='duplicate_mac')facts.rows.slice(0,50).forEach(function(g){g.entries.forEach(function(r){rows.push({code:r.code,type:r.type,employee:r.employee,status:r.status,mac:g.mac||r.mac});});});
  else if(facts.rows)rows=facts.rows.slice(0,50);
  else if(facts.activity)rows=facts.activity.slice(0,50).map(function(a){return{code:a.recordId,type:a.action,employee:a.empName,status:a.workspace,mac:a.newValue,date:a.date,field:a.field,user:a.user,oldValue:a.oldValue,newValue:a.newValue};});
  if(download)answer+='\n\n📥 Report ready: '+download.filename+' ('+download.count+' log(s)).';
  return{ok:true,answer:answer,rows:rows,matchedCount:(facts.rows||facts.activity||[]).length,intent:facts.intent,download:download||null};
}
