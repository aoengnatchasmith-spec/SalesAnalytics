// ============================================================
// BeGifted Sales Dashboard — Google Apps Script
// Code.gs  |  v4.5  |  April 2026
// ============================================================
// FIX v4.1: Payment Amount resolved by column name (not hdr.length-1)
// FIX v4.2: Cache includes totalTxn + totalAddTxn counts
// FIX v4.3: normalDay + addDay store .count per day
// ============================================================
// RUN ORDER:
//   runStep1_Extract()  → 8 staging sheets
//   runStep2_Build()    → 2 master sheets
//   runStep3_Analyze()  → Enrollment Type + Program (Wise Name) + Package Hours (Clean)
//   Deploy as Web App   → serves Dashboard.html
// ============================================================

var FILES = [
  { id: '1z9LAQbZ-V2GYLm_NA5lkkhR8fdXqyiUzW9EuiHJyeJM', mm: '01', yyyy: '2026', label: '2026-01 Jan' },
  { id: '1dRZjgRP3f0isr-ssZxobwhlsw1v8WWzR0v4zMR82o3k',  mm: '02', yyyy: '2026', label: '2026-02 Feb' },
  { id: '1G3wgBV9KnSyqNiSwHKULmbtgEbJnnLTCR-zDBqalS4w',  mm: '03', yyyy: '2026', label: '2026-03 Mar' },
  { id: '1HHtZ6YYCqK8wI6nYvVXpwgHSrqoFzcPOD7mMz8hQVJk',  mm: '04', yyyy: '2026', label: '2026-04 Apr' },
];

var SRC_PACKAGE    = '(1)PackageSales';
var SRC_ADDITIONAL = '(2)AdditionalSales';
var HEADER_ROW     = 3;
var MASTER_NORMAL     = 'MasterNormalized_NormalSales';
var MASTER_ADDITIONAL = 'MasterNormalized_AdditionalSales';

var PAID_PACKAGES = [
  '10-hr','20-hr','30-hr',
  '30-hr (free extra 1 hr)',
  '40-hr (free extra 1 hr)',
  '60-hr (free extra 3 hrs)',
  'Drop-in',
];

var PROGRAM_MAP = {
  'School Curriculum':'Y2-8 / G1-7 (Int.)','School Curriculum (2 STU)':'(2-STU) Y2-8 / G1-7 (Int.)','School Curriculum (3 STU)':'(3-STU) Y2-8 / G1-7 (Int.)','School Curriculum Master':'Y2-8 / G1-7 (Int.) Master','School Curriculum Master (2 STU)':'(2-STU) Y2-8 / G1-7 (Int.) Master','School Curriculum Master (3 STU)':'(3-STU) Y2-8 / G1-7 (Int.) Master',
  'Admission Exam Prep 11+/13+':'11+/13+','Admission Exam Prep 11+/13+ (2 STU)':'(2-STU) 11+/13+','Admission Exam Prep 11+/13+ (3 STU)':'(3-STU) 11+/13+','Admission Exam Prep 11+/13+ Master':'11+/13+ Master',
  'Admission Exam Prep 16+':'16+','Admission Exam Prep 16+ (2 STU)':'(2-STU) 16+','Admission Exam Prep 16+ (3 STU)':'(3-STU) 16+',
  'IGCSE':'Y9-11 / G8-10 (Int.)','IGCSE (2 STU)':'(2-STU) Y9-11 / G8-10 (Int.)','IGCSE (3 STU)':'(3-STU) Y9-11 / G8-10 (Int.)','IGCSE Master':'Y9-11 / G8-10 (Int.) Master','IGCSE Master (2 STU)':'(2-STU) Y9-11 / G8-10 (Int.) Master',
  'A-level OR IB Diploma':'Y12-13 / G11-12 (Int.)','A-level':'Y12-13 / G11-12 (Int.)','IB Diploma':'Y12-13 / G11-12 (Int.)',
  'A-level (2 STU) OR IB Diploma (2 STU)':'(2-STU) Y12-13 / G11-12 (Int.)','A-level (2 STU)':'(2-STU) Y12-13 / G11-12 (Int.)','IB Diploma (2 STU)':'(2-STU) Y12-13 / G11-12 (Int.)',
  'A-level (3 STU) OR IB Diploma (3 STU)':'(3-STU) Y12-13 / G11-12 (Int.)','A-level (3 STU)':'(3-STU) Y12-13 / G11-12 (Int.)','IB Diploma (3 STU)':'(3-STU) Y12-13 / G11-12 (Int.)',
  'A-Level Master OR IB Diploma Master':'Y12-13 / G11-12 (Int.) Master','A-Level Master':'Y12-13 / G11-12 (Int.) Master','IB Diploma Master':'Y12-13 / G11-12 (Int.) Master',
  'A-Level Master (2 STU) OR IB Diploma Master (2 STU)':'(2-STU) Y12-13 / G11-12 (Int.) Master','A-Level Master (2 STU)':'(2-STU) Y12-13 / G11-12 (Int.) Master','IB Diploma Master (2 STU)':'(2-STU) Y12-13 / G11-12 (Int.) Master',
  'A-Level Master (3 STU) OR IB Diploma Master (3 STU)':'(3-STU) Y12-13 / G11-12 (Int.) Master','A-Level Master (3 STU)':'(3-STU) Y12-13 / G11-12 (Int.) Master','IB Diploma Master (3 STU)':'(3-STU) Y12-13 / G11-12 (Int.) Master',
  'GED':'GED','GED (2 STU)':'(2-STU) GED','GED (3 STU)':'(3-STU) GED',
  'SAT':'SAT','SAT (2 STU)':'(2-STU) SAT','SAT (3 STU)':'(3-STU) SAT','SAT Master':'SAT Master','SAT Master (2 STU)':'(2-STU) SAT Master','SAT Master (3 STU)':'(3-STU) SAT Master',
  'IELTS/TOEFL':'IELTS','IELTS/TOEFL (2 STU)':'(2-STU) IELTS','IELTS/TOEFL (3 STU)':'(3-STU) IELTS','IELTS/TOEFL Master':'IELTS Master','IELTS/TOEFL Master (2 STU)':'(2-STU) IELTS Master',
  'University':'University','University (2 STU)':'(2-STU) University','University (3 STU)':'(3-STU) University','University Master':'University Master','University Master (2 STU)':'(2-STU) University Master','University Master (3 STU)':'(3-STU) University Master',
  'English Master Class':'English Masterclass','Interview Prep':'Interview Prep',
};


// ── STEP 1 ───────────────────────────────────────────────────
function runStep1_Extract() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), log = [];
  FILES.forEach(function(f) {
    var srcSS;
    try { srcSS = SpreadsheetApp.openById(f.id); }
    catch(e) { log.push('FAIL ' + f.label + ': ' + e); return; }
    var pkgSh = srcSS.getSheetByName(SRC_PACKAGE);
    if (!pkgSh) { log.push('FAIL: ' + SRC_PACKAGE + ' not found in ' + f.label); }
    else {
      var r = extractNormalSales(pkgSh, f.label);
      writeStaging(ss, 'NormalSales_' + f.mm + f.yyyy, r.headers, r.rows);
      log.push('OK NormalSales_' + f.mm + f.yyyy + ' — ' + r.rows.length + ' rows');
    }
    var addSh = srcSS.getSheetByName(SRC_ADDITIONAL);
    if (!addSh) { log.push('FAIL: ' + SRC_ADDITIONAL + ' not found in ' + f.label); }
    else {
      var r2 = extractAdditionalSales(addSh, f.label);
      writeStaging(ss, 'AdditionalSales_' + f.mm + f.yyyy, r2.headers, r2.rows);
      log.push('OK AdditionalSales_' + f.mm + f.yyyy + ' — ' + r2.rows.length + ' rows');
    }
  });
  showAlert('Step 1 complete.\n\n' + log.join('\n'));
}

function extractNormalSales(sheet, sourceMonth) {
  var all = sheet.getDataRange().getValues();
  var hdr = all[HEADER_ROW - 1].map(function(h) { return String(h).trim(); });
  var c   = colMap(hdr);
  var headers = ['Source Month','Transaction Date','ผู้ขาย',"Student's Nickname",'Program','Package','No. of Student','ยอดชำระสุทธิ'];
  var rows = [];
  all.slice(HEADER_ROW).forEach(function(row) {
    var nick = str(row, c["Student's Nickname"]);
    if (!nick) return;
    rows.push([sourceMonth, val(row,c['Transaction Date']), str(row,c['ผู้ขาย']), nick,
      str(row,c['Program']), str(row,c['Package']), str(row,c['No. of Student']), num(row,c['ยอดชำระสุทธิ'])]);
  });
  return { headers: headers, rows: rows };
}

function extractAdditionalSales(sheet, sourceMonth) {
  var all = sheet.getDataRange().getValues();
  var hdr = all[HEADER_ROW - 1].map(function(h) { return String(h).trim(); });
  var c   = colMap(hdr);
  var headers = ['Source Month','Transaction Date',"Student's Nickname",'Sales Type','Package','ยอดชำระสุทธิ'];
  var rows = [];
  all.slice(HEADER_ROW).forEach(function(row) {
    var nick = str(row, c["Student's Nickname"]);
    if (!nick) return;
    rows.push([sourceMonth, val(row,c['Transaction Date']), nick,
      str(row,c['Sales Type']), str(row,c['Package']), num(row,c['ยอดชำระสุทธิ'])]);
  });
  return { headers: headers, rows: rows };
}


// ── STEP 2 ───────────────────────────────────────────────────
function runStep2_Build() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), log = [];

  var normalHdrs = ["Student's Nickname",'Program','Package Hours','No. of Student','Payment Amount','Sales Representative','Transaction Date','Source Month','Enrollment Type','Program (Wise Name)','Package Hours (Clean)'];
  var normalRows = [];
  FILES.forEach(function(f) {
    var stageSh = ss.getSheetByName('NormalSales_' + f.mm + f.yyyy);
    if (!stageSh) { log.push('MISSING: NormalSales_' + f.mm + f.yyyy); return; }
    var data = stageSh.getDataRange().getValues();
    if (data.length < 2) return;
    var hdr = data[0].map(function(h){return String(h).trim();}), c = colMap(hdr);
    var before = normalRows.length;
    data.slice(1).forEach(function(row) {
      var nick = str(row, c["Student's Nickname"]);
      if (!nick) return;
      normalRows.push([nick, str(row,c['Program']), str(row,c['Package']),
        str(row,c['No. of Student']),
        num(row, c['ยอดชำระสุทธิ']),
        str(row,c['ผู้ขาย']), val(row,c['Transaction Date']), f.label, '', '', '']);
    });
    log.push('OK NormalSales_' + f.mm + f.yyyy + ' — ' + (normalRows.length-before) + ' rows');
  });
  writeMaster(ss, MASTER_NORMAL, normalHdrs, normalRows);
  log.push('Written: ' + MASTER_NORMAL + ' (' + normalRows.length + ' rows)');

  var addHdrs = ["Student's Nickname",'Sales Type','Package','Payment Amount','Transaction Date','Source Month'];
  var addRows = [];
  FILES.forEach(function(f) {
    var stageSh = ss.getSheetByName('AdditionalSales_' + f.mm + f.yyyy);
    if (!stageSh) { log.push('MISSING: AdditionalSales_' + f.mm + f.yyyy); return; }
    var data = stageSh.getDataRange().getValues();
    if (data.length < 2) return;
    var hdr = data[0].map(function(h){return String(h).trim();}), c = colMap(hdr);
    var before = addRows.length;
    data.slice(1).forEach(function(row) {
      var nick = str(row, c["Student's Nickname"]);
      if (!nick) return;
      addRows.push([nick, str(row,c['Sales Type']), str(row,c['Package']),
        num(row, c['ยอดชำระสุทธิ']),
        val(row,c['Transaction Date']), f.label]);
    });
    log.push('OK AdditionalSales_' + f.mm + f.yyyy + ' — ' + (addRows.length-before) + ' rows');
  });
  writeMaster(ss, MASTER_ADDITIONAL, addHdrs, addRows);
  log.push('Written: ' + MASTER_ADDITIONAL + ' (' + addRows.length + ' rows)');

  showAlert('Step 2 complete.\n\n' + log.join('\n'));
}


// ── STEP 3 ───────────────────────────────────────────────────
function runStep3_Analyze() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MASTER_NORMAL);
  if (!sh) { showAlert('Run Step 2 first.'); return; }

  var data = sh.getDataRange().getValues();
  var hdr  = data[0].map(function(h){return String(h).trim();});
  var rows = data.slice(1);
  var c    = colMap(hdr);
  var iNick=c["Student's Nickname"],iPkg=c['Package Hours'],iDate=c['Transaction Date'];
  var iEnroll=c['Enrollment Type'],iWise=c['Program (Wise Name)'],iProgram=c['Program'],iClean=c['Package Hours (Clean)'];

  rows.sort(function(a,b){
    var da=new Date(a[iDate]),db=new Date(b[iDate]);
    return (isNaN(da)?0:da.getTime())-(isNaN(db)?0:db.getTime());
  });

  var groups={};
  rows.forEach(function(row,idx){
    var key=String(row[iNick]||'').toLowerCase().trim();
    if (!key) return;
    if (!groups[key]) groups[key]=[];
    groups[key].push({row:row,idx:idx});
  });

  // ── Enrollment Type Classification (v4.5) ────────────────
  // Rule 1: Package == 'Trial' (case-insensitive) → always "Trial"
  //         validity hint: Payment Amount < 1000
  // Rule 2: Determine New Student vs Renewal using ALL rows per nickname
  //         sorted by Transaction Date (already sorted above)
  //
  //   Find the student's FIRST non-Trial paid row:
  //   - If there is exactly ONE paid row AND the row before it (chronologically)
  //     was a Trial row → that paid row = "New Student"
  //   - Otherwise all paid rows = "Renewal"
  //
  //   Edge cases handled:
  //   - 1 row, not Trial → Renewal  (e.g. Custard.Pa with 20-hr)
  //   - 1 row, Trial     → Trial
  //   - 2 rows: Trial → paid  → Trial + New Student
  //   - 2 rows: paid → paid   → Renewal + Renewal
  //   - 3+ rows: Trial → paid → paid... → Trial + New Student + Renewal...
  //   - 3+ rows: paid → paid → paid... → Renewal + Renewal + Renewal...

  Object.keys(groups).forEach(function(key){
    var items=groups[key]; // already sorted by date

    // Separate into Trial rows and paid rows
    var trialIndices=[];
    var paidIndices=[];
    items.forEach(function(item,i){
      var pkg=String(item.row[iPkg]||'').trim().toLowerCase();
      if(pkg==='trial') trialIndices.push(i);
      else paidIndices.push(i);
    });

    // Step 1: label all Trial rows
    trialIndices.forEach(function(i){
      rows[items[i].idx][iEnroll]='Trial';
    });

    // Step 2: label paid rows
    // A paid row is "New Student" only if:
    //   - it is the FIRST paid row for this student, AND
    //   - the row immediately before it (by position in sorted items) is a Trial row
    paidIndices.forEach(function(pi, arrPos){
      var idx=items[pi].idx;
      if(arrPos===0){
        // First paid row — check if the previous item (pi-1) was a Trial
        var prevIsTrial=(pi>0 && String(items[pi-1].row[iPkg]||'').trim().toLowerCase()==='trial');
        rows[idx][iEnroll]=prevIsTrial?'New Student':'Renewal';
      } else {
        // All subsequent paid rows are Renewal
        rows[idx][iEnroll]='Renewal';
      }
    });
  });

  rows.forEach(function(row){
    var prog=String(row[iProgram]||'').trim();
    if (iWise>=0)  row[iWise]=PROGRAM_MAP[prog]||prog;
    if (iClean>=0) row[iClean]=String(row[iPkg]||'').trim().replace(/\s*\(.*?\)/g,'').trim();
  });

  var out=[hdr].concat(rows);
  sh.clearContents();sh.clearFormats();
  sh.getRange(1,1,out.length,hdr.length).setValues(out);
  sh.getRange(1,1,1,hdr.length).setBackground('#003087').setFontColor('#FFFFFF').setFontWeight('bold');
  for (var r=2;r<=rows.length+1;r++) sh.getRange(r,1,1,hdr.length).setBackground(r%2===0?'#F0F4FF':'#FFFFFF');
  if (iEnroll>=0&&rows.length>0) rows.forEach(function(row,i){
    var et=row[iEnroll];
    sh.getRange(i+2,iEnroll+1).setBackground(et==='Trial'?'#DBEAFE':et==='New Student'?'#D1FAE5':et==='Renewal'?'#FEF3C7':'#FFFFFF');
  });
  if (iDate>=0&&rows.length>0) sh.getRange(2,iDate+1,rows.length,1).setNumberFormat('dd/mm/yyyy');
  var iPay=c['Payment Amount'];
  if (iPay>=0&&rows.length>0) sh.getRange(2,iPay+1,rows.length,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(1);sh.autoResizeColumns(1,hdr.length);

  var counts={};
  rows.forEach(function(r){var et=r[iEnroll]||'';counts[et]=(counts[et]||0)+1;});
  var summary=Object.keys(counts).map(function(k){return k+': '+counts[k];}).join(' | ');

  buildDashboardCache(ss);
  showAlert('Step 3 complete.\n\n'+rows.length+' rows processed.\n'+summary+'\nDashboard cache built.');
}


// ── BUILD DASHBOARD CACHE ────────────────────────────────────
function buildDashboardCache(ss) {
  var normalSh=ss.getSheetByName(MASTER_NORMAL);
  var addSh=ss.getSheetByName(MASTER_ADDITIONAL);
  if (!normalSh||!addSh){Logger.log('Cannot build cache — master sheets missing');return;}

  var DAY_NAMES=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  var nData=normalSh.getDataRange().getValues();
  var nHdr=nData[0].map(function(h){return String(h).trim();}),nc=colMap(nHdr);
  var byDay={},pkgCount={},progCount={},repRev={},repCnt={};
  var dayCount={Mon:0,Tue:0,Wed:0,Thu:0,Fri:0,Sat:0,Sun:0};
  var totalTxn=0;

  nData.slice(1).forEach(function(r){
    if (!String(r[0]).trim()) return;
    totalTxn++;
    var d=formatDate(r[nc['Transaction Date']]);
    var mon=String(r[nc['Source Month']]||'').trim();
    var e=String(r[nc['Enrollment Type']]||'').trim();
    var pay=parseFloat(r[nc['Payment Amount']])||0;
    var pkg=String(r[nc['Package Hours (Clean)']]||r[nc['Package Hours']]||'').trim();
    var prg=String(r[nc['Program (Wise Name)']]||'').trim();
    var rep=String(r[nc['Sales Representative']]||'').trim();
    if (!byDay[d]) byDay[d]={d:d,m:mon,rev:0,trial:0,newS:0,renew:0,count:0};
    byDay[d].rev+=pay;
    byDay[d].count++;
    if (e==='Trial')       byDay[d].trial++;
    if (e==='New Student') byDay[d].newS++;
    if (e==='Renewal')     byDay[d].renew++;
    if (pkg) pkgCount[pkg]=(pkgCount[pkg]||0)+1;
    if (prg) progCount[prg]=(progCount[prg]||0)+1;
    if (rep){repRev[rep]=(repRev[rep]||0)+pay;repCnt[rep]=(repCnt[rep]||0)+1;}
    if (d){var dt=new Date(d);if(!isNaN(dt))dayCount[DAY_NAMES[dt.getDay()]]++;}
  });

  var aData=addSh.getDataRange().getValues();
  var aHdr=aData[0].map(function(h){return String(h).trim();}),ac=colMap(aHdr);
  var addByDay={},addPkgCount={};
  var totalAddTxn=0;

  aData.slice(1).forEach(function(r){
    if (!String(r[0]).trim()) return;
    totalAddTxn++;
    var d=formatDate(r[ac['Transaction Date']]);
    var mon=String(r[ac['Source Month']]||'').trim();
    var pay=parseFloat(r[ac['Payment Amount']])||0;
    var pkg=String(r[ac['Package']]||'').trim();
    if (!addByDay[d]) addByDay[d]={d:d,m:mon,rev:0,count:0};
    addByDay[d].rev+=pay;
    addByDay[d].count++;
    if (pkg) addPkgCount[pkg]=(addPkgCount[pkg]||0)+1;
  });

  var repArr=Object.keys(repRev).map(function(k){
    return{name:k,revenue:Math.round(repRev[k]),count:repCnt[k]};
  }).sort(function(a,b){return b.revenue-a.revenue;});

  var payload={
    normalDays:  Object.values(byDay).sort(function(a,b){return a.d<b.d?-1:1;}),
    addDays:     Object.values(addByDay).sort(function(a,b){return a.d<b.d?-1:1;}),
    pkgCount:    pkgCount,
    progCount:   progCount,
    addPkgCount: addPkgCount,
    repArr:      repArr,
    dayCount:    dayCount,
    totalTxn:    totalTxn,
    totalAddTxn: totalAddTxn,
    lastUpdated: new Date().toISOString(),
  };

  var json=JSON.stringify(payload);
  var cacheSh=ss.getSheetByName('Dashboard_Cache');
  if (!cacheSh) cacheSh=ss.insertSheet('Dashboard_Cache');
  cacheSh.clearContents();
  cacheSh.getRange(1,1).setValue(json);
  cacheSh.hideSheet();
  Logger.log('Cache built: '+Math.round(json.length/1024)+' KB | normalTxn='+totalTxn+' addTxn='+totalAddTxn);
}


// ── WEB APP ──────────────────────────────────────────────────
function doGet() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cacheSh=ss.getSheetByName('Dashboard_Cache');
  var dataJson='null';
  if (cacheSh){try{var v=cacheSh.getRange(1,1).getValue();if(v)dataJson=v;}catch(e){}}
  var template=HtmlService.createTemplateFromFile('Dashboard');
  template.embeddedData=dataJson;
  return template.evaluate().setTitle('BeGifted Sales Dashboard').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cacheSh=ss.getSheetByName('Dashboard_Cache');
  if (!cacheSh) return{error:'Run Step 3 first to build cache.'};
  try{var v=cacheSh.getRange(1,1).getValue();if(v)return JSON.parse(v);}catch(e){}
  return{error:'Cache error. Re-run Step 3.'};
}


// ── TRIGGERS ─────────────────────────────────────────────────
// Run setupTriggers() ONCE from Apps Script editor.
// It will schedule a daily full refresh (Step1+2+3) at 1:00 AM Bangkok time.
// Run removeTriggers() to cancel all triggers.

function setupTriggers() {
  // Remove any existing triggers first to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // Daily trigger at 01:00 AM
  ScriptApp.newTrigger('dailyRefresh')
    .timeBased()
    .atHour(1)        // 1 AM
    .everyDays(1)
    .inTimezone('Asia/Bangkok')
    .create();

  showAlert('✅ Trigger set!\n\ndailyRefresh() will run every day at 1:00 AM (Bangkok time).\n\nDashboard will auto-update overnight — no manual steps needed.');
}

function removeTriggers() {
  var count = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
    count++;
  });
  showAlert('Removed ' + count + ' trigger(s).');
}

// Called automatically by the daily trigger
// Runs full pipeline: Extract → Build → Analyze → Cache
function dailyRefresh() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ['dailyRefresh started: ' + new Date().toISOString()];

  try {
    // Step 1 — Extract from source sheets
    var step1Log = [];
    FILES.forEach(function(f) {
      var srcSS;
      try { srcSS = SpreadsheetApp.openById(f.id); }
      catch(e) { step1Log.push('SKIP ' + f.label + ': ' + e); return; }

      var pkgSh = srcSS.getSheetByName(SRC_PACKAGE);
      if (pkgSh) {
        var r = extractNormalSales(pkgSh, f.label);
        writeStaging(ss, 'NormalSales_' + f.mm + f.yyyy, r.headers, r.rows);
        step1Log.push('OK NormalSales_' + f.mm + f.yyyy + ' (' + r.rows.length + ' rows)');
      }
      var addSh = srcSS.getSheetByName(SRC_ADDITIONAL);
      if (addSh) {
        var r2 = extractAdditionalSales(addSh, f.label);
        writeStaging(ss, 'AdditionalSales_' + f.mm + f.yyyy, r2.headers, r2.rows);
        step1Log.push('OK AdditionalSales_' + f.mm + f.yyyy + ' (' + r2.rows.length + ' rows)');
      }
    });
    log.push('Step 1 done:\n  ' + step1Log.join('\n  '));

    // Step 2 — Build master sheets
    var normalHdrs = ["Student's Nickname",'Program','Package Hours','No. of Student','Payment Amount','Sales Representative','Transaction Date','Source Month','Enrollment Type','Program (Wise Name)','Package Hours (Clean)'];
    var normalRows = [];
    FILES.forEach(function(f) {
      var stageSh = ss.getSheetByName('NormalSales_' + f.mm + f.yyyy);
      if (!stageSh) return;
      var data = stageSh.getDataRange().getValues();
      if (data.length < 2) return;
      var hdr = data[0].map(function(h){return String(h).trim();}), c = colMap(hdr);
      data.slice(1).forEach(function(row) {
        var nick = str(row, c["Student's Nickname"]);
        if (!nick) return;
        normalRows.push([nick, str(row,c['Program']), str(row,c['Package']),
          str(row,c['No. of Student']), num(row, c['ยอดชำระสุทธิ']),
          str(row,c['ผู้ขาย']), val(row,c['Transaction Date']), f.label, '', '', '']);
      });
    });
    writeMaster(ss, MASTER_NORMAL, normalHdrs, normalRows);

    var addHdrs = ["Student's Nickname",'Sales Type','Package','Payment Amount','Transaction Date','Source Month'];
    var addRows = [];
    FILES.forEach(function(f) {
      var stageSh = ss.getSheetByName('AdditionalSales_' + f.mm + f.yyyy);
      if (!stageSh) return;
      var data = stageSh.getDataRange().getValues();
      if (data.length < 2) return;
      var hdr = data[0].map(function(h){return String(h).trim();}), c = colMap(hdr);
      data.slice(1).forEach(function(row) {
        var nick = str(row, c["Student's Nickname"]);
        if (!nick) return;
        addRows.push([nick, str(row,c['Sales Type']), str(row,c['Package']),
          num(row, c['ยอดชำระสุทธิ']), val(row,c['Transaction Date']), f.label]);
      });
    });
    writeMaster(ss, MASTER_ADDITIONAL, addHdrs, addRows);
    log.push('Step 2 done: ' + normalRows.length + ' normal, ' + addRows.length + ' additional rows');

    // Step 3 — Analyze + build cache (reuse existing function)
    runStep3_Analyze_silent(ss);
    log.push('Step 3 done: cache rebuilt');

  } catch(e) {
    log.push('ERROR: ' + e.toString());
  }

  log.push('dailyRefresh finished: ' + new Date().toISOString());
  Logger.log(log.join('\n'));
}

// Silent version of Step 3 — no UI alerts, accepts ss parameter
function runStep3_Analyze_silent(ss) {
  var sh = ss.getSheetByName(MASTER_NORMAL);
  if (!sh) { Logger.log('runStep3_Analyze_silent: MASTER_NORMAL not found'); return; }

  var data = sh.getDataRange().getValues();
  var hdr  = data[0].map(function(h){return String(h).trim();});
  var rows = data.slice(1);
  var c    = colMap(hdr);
  var iNick=c["Student's Nickname"],iPkg=c['Package Hours'],iDate=c['Transaction Date'];
  var iEnroll=c['Enrollment Type'],iWise=c['Program (Wise Name)'],iProgram=c['Program'],iClean=c['Package Hours (Clean)'];

  rows.sort(function(a,b){
    var da=new Date(a[iDate]),db=new Date(b[iDate]);
    return (isNaN(da)?0:da.getTime())-(isNaN(db)?0:db.getTime());
  });

  var groups={};
  rows.forEach(function(row,idx){
    var key=String(row[iNick]||'').toLowerCase().trim();
    if (!key) return;
    if (!groups[key]) groups[key]=[];
    groups[key].push({row:row,idx:idx});
  });

  // Enrollment Type — v4.5 logic
  Object.keys(groups).forEach(function(key){
    var items=groups[key];
    var trialIndices=[],paidIndices=[];
    items.forEach(function(item,i){
      var pkg=String(item.row[iPkg]||'').trim().toLowerCase();
      if(pkg==='trial') trialIndices.push(i);
      else paidIndices.push(i);
    });
    trialIndices.forEach(function(i){ rows[items[i].idx][iEnroll]='Trial'; });
    paidIndices.forEach(function(pi,arrPos){
      var idx=items[pi].idx;
      if(arrPos===0){
        var prevIsTrial=(pi>0&&String(items[pi-1].row[iPkg]||'').trim().toLowerCase()==='trial');
        rows[idx][iEnroll]=prevIsTrial?'New Student':'Renewal';
      } else {
        rows[idx][iEnroll]='Renewal';
      }
    });
  });

  rows.forEach(function(row){
    var prog=String(row[iProgram]||'').trim();
    if (iWise>=0)  row[iWise]=PROGRAM_MAP[prog]||prog;
    if (iClean>=0) row[iClean]=String(row[iPkg]||'').trim().replace(/\s*\(.*?\)/g,'').trim();
  });

  var out=[hdr].concat(rows);
  sh.clearContents();sh.clearFormats();
  sh.getRange(1,1,out.length,hdr.length).setValues(out);
  sh.getRange(1,1,1,hdr.length).setBackground('#003087').setFontColor('#FFFFFF').setFontWeight('bold');
  for (var r=2;r<=rows.length+1;r++) sh.getRange(r,1,1,hdr.length).setBackground(r%2===0?'#F0F4FF':'#FFFFFF');
  if (iEnroll>=0&&rows.length>0) rows.forEach(function(row,i){
    var et=row[iEnroll];
    sh.getRange(i+2,iEnroll+1).setBackground(et==='Trial'?'#DBEAFE':et==='New Student'?'#D1FAE5':et==='Renewal'?'#FEF3C7':'#FFFFFF');
  });
  if (iDate>=0&&rows.length>0) sh.getRange(2,iDate+1,rows.length,1).setNumberFormat('dd/mm/yyyy');
  var iPay=c['Payment Amount'];
  if (iPay>=0&&rows.length>0) sh.getRange(2,iPay+1,rows.length,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(1);sh.autoResizeColumns(1,hdr.length);

  buildDashboardCache(ss);
}


// ── DEBUG ─────────────────────────────────────────────────────
function debugUnmappedPrograms() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(MASTER_NORMAL);
  if (!sh){showAlert('Run Step 2 first.');return;}
  var data=sh.getDataRange().getValues();
  var hdr=data[0].map(function(h){return String(h).trim();}),c=colMap(hdr);
  var iProgram=c['Program'],iWise=c['Program (Wise Name)'];
  var unmappedRows=[],unmappedUniq={};
  data.slice(1).forEach(function(row,i){
    var prog=String(row[iProgram]||'').trim();
    if (!prog||PROGRAM_MAP[prog]) return;
    unmappedRows.push({rowNum:i+2,program:prog});
    unmappedUniq[prog]=(unmappedUniq[prog]||0)+1;
    sh.getRange(i+2,iProgram+1).setBackground('#FED7AA');
    if (iWise>=0) sh.getRange(i+2,iWise+1).setBackground('#FEE2E2');
  });
  var list=Object.keys(unmappedUniq).sort().map(function(k){return '"'+k+'" ('+unmappedUniq[k]+' rows)';});
  var logSh=ss.getSheetByName('UnmappedPrograms_Log');
  if (logSh){logSh.clearContents();logSh.clearFormats();}else{logSh=ss.insertSheet('UnmappedPrograms_Log');}
  logSh.getRange(1,1,1,3).setValues([['Program','Row','Action']]).setBackground('#C05621').setFontColor('#FFFFFF').setFontWeight('bold');
  if (unmappedRows.length>0) logSh.getRange(2,1,unmappedRows.length,3).setValues(unmappedRows.map(function(r){return[r.program,r.rowNum,'Add to PROGRAM_MAP'];}));
  logSh.autoResizeColumns(1,3);
  showAlert('Found '+unmappedRows.length+' unmapped rows / '+list.length+' unique:\n\n'+list.join('\n'));
}


// ── HELPERS ──────────────────────────────────────────────────
function showAlert(msg){
  Logger.log(msg);
  try{SpreadsheetApp.getUi().alert(msg);}catch(e){}
}
function colMap(headers){var m={};headers.forEach(function(h,i){m[h]=i;});return m;}
function str(row,idx){return(idx>=0&&row[idx]!=null)?String(row[idx]).trim():'';}
function num(row,idx){if(idx==null||idx<0||row[idx]==null)return'';var n=parseFloat(row[idx]);return isNaN(n)?'':n;}
function val(row,idx){return(idx>=0&&row[idx]!=null)?row[idx]:'';}

function formatDate(d){
  if (!d) return'';
  var dt=(d instanceof Date)?d:new Date(d);
  if (isNaN(dt)||dt.getFullYear()<2000) return'';
  var m=dt.getMonth()+1,dd=dt.getDate();
  return dt.getFullYear()+'-'+(m<10?'0'+m:m)+'-'+(dd<10?'0'+dd:dd);
}

function writeStaging(ss,sheetName,headers,rows){
  var sh=ss.getSheetByName(sheetName);
  if (sh){sh.clearContents();sh.clearFormats();}else{sh=ss.insertSheet(sheetName);}
  var all=[headers].concat(rows);
  sh.getRange(1,1,all.length,headers.length).setValues(all);
  sh.getRange(1,1,1,headers.length).setBackground('#4A4A6A').setFontColor('#FFFFFF').setFontWeight('bold');
  if (rows.length>0) sh.getRange(2,headers.length,rows.length,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(1);sh.autoResizeColumns(1,headers.length);
}

function writeMaster(ss,sheetName,headers,rows){
  var sh=ss.getSheetByName(sheetName);
  if (sh){sh.clearContents();sh.clearFormats();}else{sh=ss.insertSheet(sheetName);}
  var all=[headers].concat(rows);
  sh.getRange(1,1,all.length,headers.length).setValues(all);
  sh.getRange(1,1,1,headers.length).setBackground('#003087').setFontColor('#FFFFFF').setFontWeight('bold');
  for (var r=2;r<=rows.length+1;r++) sh.getRange(r,1,1,headers.length).setBackground(r%2===0?'#F0F4FF':'#FFFFFF');
  var dc=headers.indexOf('Transaction Date');
  if (dc>=0&&rows.length>0) sh.getRange(2,dc+1,rows.length,1).setNumberFormat('dd/mm/yyyy');
  var pc=headers.indexOf('Payment Amount');
  if (pc>=0&&rows.length>0) sh.getRange(2,pc+1,rows.length,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(1);sh.autoResizeColumns(1,headers.length);
}
