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
  var headers = ['Source Month','วันที่ชำระเงิน','ผู้ขาย',"Student's Nickname",'Program','Package','No. of Student','ยอดชำระสุทธิ','Valid Until'];
  var rows = [];
  all.slice(HEADER_ROW).forEach(function(row) {
    var nick = str(row, c["Student's Nickname"]);
    if (!nick) return;
    // ✅ Skip rows where วันที่ชำระเงิน is empty — payment not received yet
    var payDate = val(row, c['วันที่ชำระเงิน']);
    if (!payDate || String(payDate).trim()==='' || (payDate instanceof Date && isNaN(payDate))) return;
    rows.push([sourceMonth, payDate, str(row,c['ผู้ขาย']), nick,
      str(row,c['Program']), str(row,c['Package']), str(row,c['No. of Student']),
      num(row,c['ยอดชำระสุทธิ']), val(row,c['Valid Until'])]);
  });
  return { headers: headers, rows: rows };
}

function extractAdditionalSales(sheet, sourceMonth) {
  var all = sheet.getDataRange().getValues();
  var hdr = all[HEADER_ROW - 1].map(function(h) { return String(h).trim(); });
  var c   = colMap(hdr);
  var headers = ['Source Month','วันที่ชำระเงิน',"Student's Nickname",'Sales Type','Package','ยอดชำระสุทธิ'];
  var rows = [];
  all.slice(HEADER_ROW).forEach(function(row) {
    var nick = str(row, c["Student's Nickname"]);
    if (!nick) return;
    // ✅ Skip rows where วันที่ชำระเงิน is empty
    var payDate = val(row, c['วันที่ชำระเงิน']);
    if (!payDate || String(payDate).trim()==='' || (payDate instanceof Date && isNaN(payDate))) return;
    rows.push([sourceMonth, payDate, nick,
      str(row,c['Sales Type']), str(row,c['Package']), num(row,c['ยอดชำระสุทธิ'])]);
  });
  return { headers: headers, rows: rows };
}


// ── STEP 2 ───────────────────────────────────────────────────
function runStep2_Build() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), log = [];

  var normalHdrs = ["Student's Nickname",'Program','Package Hours','No. of Student','Payment Amount','Sales Representative','Payment Date','Source Month','Enrollment Type','Program (Wise Name)','Package Hours (Clean)','Valid Until'];
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
        str(row,c['ผู้ขาย']), val(row,c['วันที่ชำระเงิน']), f.label, '', '', '',
        val(row,c['Valid Until'])]);
    });
    log.push('OK NormalSales_' + f.mm + f.yyyy + ' — ' + (normalRows.length-before) + ' rows');
  });
  writeMaster(ss, MASTER_NORMAL, normalHdrs, normalRows);
  log.push('Written: ' + MASTER_NORMAL + ' (' + normalRows.length + ' rows)');

  var addHdrs = ["Student's Nickname",'Sales Type','Package','Payment Amount','Payment Date','Source Month'];
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
        val(row,c['วันที่ชำระเงิน']), f.label]);
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
  var iNick=c["Student's Nickname"],iPkg=c['Package Hours'],iDate=c['Payment Date'];
  var iEnroll=c['Enrollment Type'],iWise=c['Program (Wise Name)'],iProgram=c['Program'],iClean=c['Package Hours (Clean)'];
  var iValid=c['Valid Until'];

  // ── Add Churn Status column if not present ────────────────
  var iChurn=c['Churn Status'];
  if (iChurn===undefined) {
    hdr.push('Churn Status');
    iChurn=hdr.length-1;
    rows.forEach(function(row){ row.push(''); });
  }

  // Sort by Payment Date (oldest → newest)
  rows.sort(function(a,b){
    var da=new Date(a[iDate]),db=new Date(b[iDate]);
    return (isNaN(da)?0:da.getTime())-(isNaN(db)?0:db.getTime());
  });

  // Group rows by student nickname
  var groups={};
  rows.forEach(function(row,idx){
    var key=String(row[iNick]||'').toLowerCase().trim();
    if (!key) return;
    if (!groups[key]) groups[key]=[];
    groups[key].push({row:row,idx:idx});
  });

  // ── Enrollment Type Classification (v4.5) ────────────────
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

  // ── Program (Wise Name) + Package Hours (Clean) ───────────
  rows.forEach(function(row){
    var prog=String(row[iProgram]||'').trim();
    if (iWise>=0)  row[iWise]=PROGRAM_MAP[prog]||prog;
    if (iClean>=0) row[iClean]=String(row[iPkg]||'').trim().replace(/\s*\(.*?\)/g,'').trim();
  });

  // ── Churn Status per student (latest row only) ────────────
  //
  // Values:
  //   'Active'   — Valid Until + 14d >= Today (still in grace or future)
  //   'Retained' — Valid Until + 14d < Today, but has a later payment
  //   'Churned'  — Valid Until + 14d < Today, no later payment
  //   'N/A'      — Trial-only student OR no Valid Until
  //
  var GRACE_DAYS=14, MS_PER_DAY=86400000;
  var today=new Date(); today.setHours(0,0,0,0);

  function parseValidUntil(v){
    if (!v||v==='') return null;
    if (v instanceof Date) return isNaN(v)?null:v;
    if (typeof v==='number'&&v>1000) return new Date((v-25569)*MS_PER_DAY);
    var d=new Date(v); return isNaN(d)?null:d;
  }

  // First pass: reset all Churn Status to '—' (non-latest rows)
  rows.forEach(function(row){ row[iChurn]='—'; });

  // Second pass: compute status for latest row of each student
  Object.keys(groups).forEach(function(key){
    var items=groups[key];
    // latest item = last in sorted array
    var latestItem=items[items.length-1];
    var latestRow=rows[latestItem.idx];
    var enrollType=String(latestRow[iEnroll]||'').trim();

    // Trial-only: check if ALL rows are Trial
    var allTrial=items.every(function(item){
      return String(item.row[iPkg]||'').trim().toLowerCase()==='trial';
    });
    if (allTrial){ latestRow[iChurn]='N/A'; return; }

    // Get Valid Until from latest NON-trial row
    var latestPaidItem=null;
    for (var i=items.length-1;i>=0;i--){
      var pkg=String(items[i].row[iPkg]||'').trim().toLowerCase();
      if (pkg!=='trial'){ latestPaidItem=items[i]; break; }
    }
    if (!latestPaidItem){ latestRow[iChurn]='N/A'; return; }

    var validUntil=parseValidUntil(latestPaidItem.row[iValid]);
    if (!validUntil){ latestRow[iChurn]='N/A'; return; }

    var graceDeadline=new Date(validUntil.getTime()+GRACE_DAYS*MS_PER_DAY);
    graceDeadline.setHours(0,0,0,0);

    // Still active (within grace period)
    if (graceDeadline>=today){ latestRow[iChurn]='Active'; return; }

    // Past grace — check if any payment after grace deadline
    var allDates=items.map(function(item){
      var d=new Date(item.row[iDate]);
      return isNaN(d)?0:d.getTime();
    });
    var renewed=allDates.some(function(ts){ return ts>graceDeadline.getTime(); });
    latestRow[iChurn]=renewed?'Retained':'Churned';
  });

  // ── Write back to sheet ───────────────────────────────────
  var out=[hdr].concat(rows);
  sh.clearContents();sh.clearFormats();
  sh.getRange(1,1,out.length,hdr.length).setValues(out);
  sh.getRange(1,1,1,hdr.length).setBackground('#003087').setFontColor('#FFFFFF').setFontWeight('bold');
  for (var r=2;r<=rows.length+1;r++) sh.getRange(r,1,1,hdr.length).setBackground(r%2===0?'#F0F4FF':'#FFFFFF');

  // Enrollment Type colors
  if (iEnroll>=0&&rows.length>0) rows.forEach(function(row,i){
    var et=row[iEnroll];
    sh.getRange(i+2,iEnroll+1).setBackground(et==='Trial'?'#DBEAFE':et==='New Student'?'#D1FAE5':et==='Renewal'?'#FEF3C7':'#FFFFFF');
  });

  // Churn Status colors
  if (iChurn>=0&&rows.length>0) rows.forEach(function(row,i){
    var cs=row[iChurn];
    var bg=cs==='Churned'?'#FEE2E2':cs==='Active'?'#D1FAE5':cs==='Retained'?'#FEF3C7':cs==='N/A'?'#F3F4F6':'#FFFFFF';
    sh.getRange(i+2,iChurn+1).setBackground(bg);
  });

  // Date / number formats
  if (iDate>=0&&rows.length>0) sh.getRange(2,iDate+1,rows.length,1).setNumberFormat('dd/mm/yyyy');
  var iPay=c['Payment Amount'];
  if (iPay>=0&&rows.length>0) sh.getRange(2,iPay+1,rows.length,1).setNumberFormat('#,##0.00');
  var iValidCol=hdr.indexOf('Valid Until');
  if (iValidCol>=0&&rows.length>0) sh.getRange(2,iValidCol+1,rows.length,1).setNumberFormat('dd/mm/yyyy');

  sh.setFrozenRows(1);sh.autoResizeColumns(1,hdr.length);

  // ── Summary ───────────────────────────────────────────────
  var enrollCounts={},churnCounts={};
  rows.forEach(function(r){
    var et=r[iEnroll]||''; enrollCounts[et]=(enrollCounts[et]||0)+1;
    var cs=r[iChurn]||''; if(cs!=='—') churnCounts[cs]=(churnCounts[cs]||0)+1;
  });
  var enrollSummary=Object.keys(enrollCounts).map(function(k){return k+': '+enrollCounts[k];}).join(' | ');
  var churnSummary=Object.keys(churnCounts).map(function(k){return k+': '+churnCounts[k];}).join(' | ');

  SpreadsheetApp.flush(); // commit all sheet writes before reading for cache
  buildDashboardCache(ss);
  showAlert('Step 3 complete.\n\n'+rows.length+' rows processed.\n\nEnrollment: '+enrollSummary+'\nChurn: '+churnSummary+'\n\nDashboard cache built.');
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

  // ✅ unique student sets for correct conversion rate
  var trialStudents={};    // nick → earliest trial date
  var newStudents={};      // nick → date of New Student row
  var renewStudents={};    // nick → true

  nData.slice(1).forEach(function(r){
    if (!String(r[0]).trim()) return;
    totalTxn++;
    var d=formatDate(r[nc['Payment Date']]);
    var mon=String(r[nc['Source Month']]||'').trim();
    var e=String(r[nc['Enrollment Type']]||'').trim();
    var pay=parseFloat(r[nc['Payment Amount']])||0;
    var pkg=String(r[nc['Package Hours (Clean)']]||r[nc['Package Hours']]||'').trim();
    var prg=String(r[nc['Program (Wise Name)']]||'').trim();
    var rep=String(r[nc['Sales Representative']]||'').trim();
    var nick=String(r[nc["Student's Nickname"]]||'').trim().toLowerCase();
    var dow=d?(function(){var dt=new Date(d);return isNaN(dt)?'':DAY_NAMES[dt.getDay()];}()):'';

    if (!byDay[d]) byDay[d]={d:d,m:mon,rev:0,trial:0,newS:0,renew:0,count:0,revT:0,revN:0,revR:0,pkgs:{},prgs:{},reps:{},dow:dow};
    byDay[d].rev+=pay;
    byDay[d].count++;
    if (e==='Trial')       { byDay[d].trial++; byDay[d].revT+=pay; if(nick&&!trialStudents[nick]) trialStudents[nick]=d; }
    if (e==='New Student') { byDay[d].newS++;  byDay[d].revN+=pay; if(nick) newStudents[nick]=d; }
    if (e==='Renewal')     { byDay[d].renew++; byDay[d].revR+=pay; if(nick) renewStudents[nick]=true; }
    // per-day aggregates for filter-aware charts
    if (pkg) byDay[d].pkgs[pkg]=(byDay[d].pkgs[pkg]||0)+1;
    if (prg) byDay[d].prgs[prg]=(byDay[d].prgs[prg]||0)+1;
    if (rep){
      if(!byDay[d].reps[rep]) byDay[d].reps[rep]={rev:0,count:0,revT:0,revN:0,revR:0,cntT:0,cntN:0,cntR:0};
      byDay[d].reps[rep].rev+=pay;
      byDay[d].reps[rep].count++;
      if(e==='Trial')       { byDay[d].reps[rep].revT+=pay; byDay[d].reps[rep].cntT++; }
      if(e==='New Student') { byDay[d].reps[rep].revN+=pay; byDay[d].reps[rep].cntN++; }
      if(e==='Renewal')     { byDay[d].reps[rep].revR+=pay; byDay[d].reps[rep].cntR++; }
    }
    // global counts (kept for backward compat)
    if (pkg) pkgCount[pkg]=(pkgCount[pkg]||0)+1;
    if (prg) progCount[prg]=(progCount[prg]||0)+1;
    if (rep){repRev[rep]=(repRev[rep]||0)+pay;repCnt[rep]=(repCnt[rep]||0)+1;}
    if (dow) dayCount[dow]=(dayCount[dow]||0)+1;
  });

  // Count unique students per category
  var uniqueTrials      = Object.keys(trialStudents).length;
  var uniqueNewStudents = Object.keys(newStudents).length;
  var uniqueRenewals    = Object.keys(renewStudents).length;

  // ── Churn Stats — read from Churn Status column ──────────
  // Build churnList: one entry per Churned/Retained student
  // with validUntilDate so dashboard can filter by period
  var iChurnCol=nc['Churn Status'];
  var iValidCol=nc['Valid Until'];
  var iNickCol =nc["Student's Nickname"];
  var churnList=[];  // [{nick, validUntilDate, status:'Churned'|'Retained'}]
  var MS_PER_DAY_C=86400000;

  if(iChurnCol!==undefined){
    // Collect only latest-row entries (those with Churned/Retained/Active)
    nData.slice(1).forEach(function(r){
      var cs=String(r[iChurnCol]||'').trim();
      if(cs!=='Churned'&&cs!=='Retained'&&cs!=='Active') return; // '—' and 'N/A' skipped
      var nick=String(r[iNickCol]||'').trim().toLowerCase();
      var validRaw=r[iValidCol];
      var validD;
      if(validRaw instanceof Date)                               validD=validRaw;
      else if(typeof validRaw==='number'&&validRaw>1000)         validD=new Date((validRaw-25569)*MS_PER_DAY_C);
      else                                                        validD=new Date(validRaw);
      if(isNaN(validD)) validD=null;
      churnList.push({
        nick:nick,
        validUntil: validD ? formatDate(validD) : '',  // "YYYY-MM-DD"
        status:cs
      });
    });
  }

  // Global totals (for YTD display)
  var churnedStudents=churnList.filter(function(x){return x.status==='Churned';}).length;
  var eligibleStudents=churnList.filter(function(x){return x.status==='Churned'||x.status==='Retained';}).length;

  var aData=addSh.getDataRange().getValues();
  var aHdr=aData[0].map(function(h){return String(h).trim();}),ac=colMap(aHdr);
  var addByDay={},addPkgCount={};
  var totalAddTxn=0;

  aData.slice(1).forEach(function(r){
    if (!String(r[0]).trim()) return;
    totalAddTxn++;
    var d=formatDate(r[ac['Payment Date']]);
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

  // Split payload into parts to avoid 50000 char cell limit
  // Part 1: normalDays (largest - has per-day detail)
  // Part 2: addDays + aggregates
  // Part 3: churnList

  var normalDaysArr=Object.values(byDay).sort(function(a,b){return a.d<b.d?-1:1;});
  var addDaysArr=Object.values(addByDay).sort(function(a,b){return a.d<b.d?-1:1;});

  var part1=JSON.stringify(normalDaysArr);
  var part2=JSON.stringify({
    addDays:      addDaysArr,
    pkgCount:     pkgCount,
    progCount:    progCount,
    addPkgCount:  addPkgCount,
    repArr:       repArr,
    dayCount:     dayCount,
    totalTxn:     totalTxn,
    totalAddTxn:  totalAddTxn,
    uniqueTrials:     uniqueTrials,
    uniqueNewStudents:uniqueNewStudents,
    uniqueRenewals:   uniqueRenewals,
    churnedStudents:  churnedStudents,
    eligibleStudents: eligibleStudents,
    lastUpdated:  new Date().toISOString(),
  });
  var part3=JSON.stringify(churnList);

  var cacheSh=ss.getSheetByName('Dashboard_Cache');
  if (!cacheSh) cacheSh=ss.insertSheet('Dashboard_Cache');
  cacheSh.clearContents();
  cacheSh.getRange(1,1).setValue(part1);  // normalDays
  cacheSh.getRange(1,2).setValue(part2);  // aggregates
  cacheSh.getRange(1,3).setValue(part3);  // churnList
  cacheSh.hideSheet();
  Logger.log('Cache built: normalDays='+Math.round(part1.length/1024)+'KB aggregates='+Math.round(part2.length/1024)+'KB churn='+Math.round(part3.length/1024)+'KB');
}


// ── WEB APP ──────────────────────────────────────────────────
function doGet() {
  var template=HtmlService.createTemplateFromFile('Dashboard');
  template.embeddedData='null';  // always load fresh from getDashboardData()
  return template.evaluate().setTitle('BeGifted Sales Dashboard').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cacheSh=ss.getSheetByName('Dashboard_Cache');
  if (!cacheSh) return{error:'Run Step 3 first to build cache.'};
  try{
    var merged=mergeCacheParts_(cacheSh);
    if(merged) return merged;
  }catch(e){}
  return{error:'Cache error. Re-run Step 3.'};
}

// Read 3 cache cells and merge into single payload object
function mergeCacheParts_(cacheSh){
  var p1=cacheSh.getRange(1,1).getValue();  // normalDays array
  var p2=cacheSh.getRange(1,2).getValue();  // aggregates
  var p3=cacheSh.getRange(1,3).getValue();  // churnList
  if(!p1||!p2) return null;
  var normalDays=JSON.parse(p1);
  var agg=JSON.parse(p2);
  agg.normalDays=normalDays;
  if(p3) agg.churnList=JSON.parse(p3);
  return agg;
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
    var normalHdrs = ["Student's Nickname",'Program','Package Hours','No. of Student','Payment Amount','Sales Representative','Payment Date','Source Month','Enrollment Type','Program (Wise Name)','Package Hours (Clean)','Valid Until'];
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
          str(row,c['ผู้ขาย']), val(row,c['วันที่ชำระเงิน']), f.label, '', '', '',
          val(row,c['Valid Until'])]);
      });
    });
    writeMaster(ss, MASTER_NORMAL, normalHdrs, normalRows);

    var addHdrs = ["Student's Nickname",'Sales Type','Package','Payment Amount','Payment Date','Source Month'];
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
          num(row, c['ยอดชำระสุทธิ']), val(row,c['วันที่ชำระเงิน']), f.label]);
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
  return 'ok';
}

// Silent version of Step 3 — no UI alerts, accepts ss parameter
function runStep3_Analyze_silent(ss) {
  var sh = ss.getSheetByName(MASTER_NORMAL);
  if (!sh) { Logger.log('runStep3_Analyze_silent: MASTER_NORMAL not found'); return; }

  var data = sh.getDataRange().getValues();
  var hdr  = data[0].map(function(h){return String(h).trim();});
  var rows = data.slice(1);
  var c    = colMap(hdr);
  var iNick=c["Student's Nickname"],iPkg=c['Package Hours'],iDate=c['Payment Date'];
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
  if (rows.length>0){
    // Payment amount column
    var payCol=headers.indexOf('ยอดชำระสุทธิ');
    if(payCol>=0) sh.getRange(2,payCol+1,rows.length,1).setNumberFormat('#,##0.00');
    // วันที่ชำระเงิน column
    var pdCol=headers.indexOf('วันที่ชำระเงิน');
    if(pdCol>=0) sh.getRange(2,pdCol+1,rows.length,1).setNumberFormat('dd/mm/yyyy');
    // Valid Until column
    var validCol=headers.indexOf('Valid Until');
    if(validCol>=0) sh.getRange(2,validCol+1,rows.length,1).setNumberFormat('dd/mm/yyyy');
  }
  sh.setFrozenRows(1);sh.autoResizeColumns(1,headers.length);
}

function writeMaster(ss,sheetName,headers,rows){
  var sh=ss.getSheetByName(sheetName);
  if (sh){sh.clearContents();sh.clearFormats();}else{sh=ss.insertSheet(sheetName);}
  var all=[headers].concat(rows);
  sh.getRange(1,1,all.length,headers.length).setValues(all);
  sh.getRange(1,1,1,headers.length).setBackground('#003087').setFontColor('#FFFFFF').setFontWeight('bold');
  for (var r=2;r<=rows.length+1;r++) sh.getRange(r,1,1,headers.length).setBackground(r%2===0?'#F0F4FF':'#FFFFFF');
  // Payment Date
  var dc=headers.indexOf('Payment Date');
  if (dc>=0&&rows.length>0) sh.getRange(2,dc+1,rows.length,1).setNumberFormat('dd/mm/yyyy');
  // Payment Amount
  var pc=headers.indexOf('Payment Amount');
  if (pc>=0&&rows.length>0) sh.getRange(2,pc+1,rows.length,1).setNumberFormat('#,##0.00');
  // Valid Until
  var vc=headers.indexOf('Valid Until');
  if (vc>=0&&rows.length>0) sh.getRange(2,vc+1,rows.length,1).setNumberFormat('dd/mm/yyyy');
  sh.setFrozenRows(1);sh.autoResizeColumns(1,headers.length);
}
