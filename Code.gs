/**
 * ══════════════════════════════════════════════════════════════
 *  BAR ASRA — Review Gate & Competition System
 *  Google Apps Script Backend  v3.0
 * ══════════════════════════════════════════════════════════════
 *
 *  HOW IT WORKS:
 *    Customer scans QR → rates 1-5 stars
 *    ★★★★★  → Enters name + email for "Win Dinner for Two"
 *             → Logged as competition entry → Redirected to Google Reviews
 *    ★★★★☆  → Private feedback form (name, contact, message) — LOCKED
 *    ★★★☆☆     Feedback goes to Sheet + instant email alert
 *    ★★☆☆☆     Bad reviews NEVER reach Google.
 *    ★☆☆☆☆
 *
 *  SETUP:
 *    1. Create a new Google Sheet
 *    2. Extensions → Apps Script → paste this file
 *    3. Select "initializeSheet" → Run (grant permissions)
 *    4. Deploy → New Deployment → Web App
 *         Execute as: Me  |  Who has access: Anyone
 *    5. Copy Web App URL into review.html + dashboard.html
 *    6. (Optional) Triggers → Add Trigger → sendDailyDigest → Day timer
 *
 *  SHEETS:
 *    📊 Raw Data       — every star tap
 *    💬 Feedback        — private feedback (1-4 stars)
 *    🏆 Competition     — "dinner for two" entries (5 stars)
 *    ⚙️ Settings        — configurable values
 * ══════════════════════════════════════════════════════════════
 */

const CONFIG = {
  managerEmail:     "lbarwe1@gmail.com",
  adminPassword:    "barasra2024",
  defaultGoogleUrl: "https://search.google.com/local/writereview?placeid=YOUR_PLACE_ID",
  defaultThreshold: 5,
  competitionPrize: "Dinner for Two",
};


// ═════════════════════════════════════════════════════════════
//  INITIALISATION
// ═════════════════════════════════════════════════════════════

function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupRaw(_sheet(ss, "Raw Data"));
  _setupFeedback(_sheet(ss, "Feedback"));
  _setupCompetition(_sheet(ss, "Competition"));
  _setupSettings(_sheet(ss, "Settings"));
  ss.setActiveSheet(ss.getSheetByName("Raw Data"));
  // Only show UI alert when running from editor (not from web app)
  try {
    SpreadsheetApp.getUi().alert(
      "✅ Bar Asra Review System Ready!\n\n" +
      "📊 Raw Data — all star taps\n💬 Feedback — private feedback\n" +
      "🏆 Competition — dinner-for-two entries\n⚙️ Settings\n\nNext → Deploy as Web App"
    );
  } catch (e) {
    // Silently ignore UI errors when running from web app context
    Logger.log("initializeSheet completed (no UI available in this context)");
  }
}

function _sheet(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function _header(s, n, bg, fg) {
  s.getRange(1,1,1,n).setFontWeight("bold").setBackground(bg).setFontColor(fg).setHorizontalAlignment("center");
  s.setFrozenRows(1);
}

function _setupRaw(s) {
  if (s.getLastRow() > 0 && s.getRange(1,1).getValue() === "Timestamp") return;
  s.clear(); s.appendRow(["Timestamp","Rating","Source","Date","Time","Day"]);
  _header(s, 6, "#1c1a18", "#C7A27A");
}

function _setupFeedback(s) {
  if (s.getLastRow() > 0 && s.getRange(1,1).getValue() === "Timestamp") return;
  s.clear(); s.appendRow(["Timestamp","Rating","Name","Contact","Message","Date","Status"]);
  _header(s, 7, "#C7A27A", "#1c1a18"); s.setColumnWidth(5, 400);
}

function _setupCompetition(s) {
  if (s.getLastRow() > 0 && s.getRange(1,1).getValue() === "Timestamp") return;
  s.clear(); s.appendRow(["Timestamp","Name","Email","Date","Google Click"]);
  _header(s, 5, "#2d5a27", "#ffffff"); s.setColumnWidth(2, 180); s.setColumnWidth(3, 250);
}

function _setupSettings(s) {
  if (s.getLastRow() > 0 && s.getRange(1,1).getValue() === "Key") return;
  s.clear(); s.appendRow(["Key","Value"]); _header(s, 2, "#333", "#fff");
  s.appendRow(["googleReviewUrl", CONFIG.defaultGoogleUrl]);
  s.appendRow(["googleThreshold", CONFIG.defaultThreshold]);
  s.appendRow(["managerEmail", CONFIG.managerEmail]);
  s.appendRow(["competitionPrize", CONFIG.competitionPrize]);
}


// ═════════════════════════════════════════════════════════════
//  REQUEST HANDLERS
// ═════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const a = (e.parameter.action || "").trim();
    let r;
    switch (a) {
      case "submitRating":
        _recordRating(int(e.parameter.rating), e.parameter.source || "website"); r = {status:"success"}; break;
      case "submitFeedback":
        _recordFeedback({rating:int(e.parameter.rating), name:e.parameter.name||"", contact:e.parameter.contact||"", message:e.parameter.message||""});
        r = {status:"success"}; break;
      case "submitCompetition":
        _recordCompetition({name:e.parameter.name||"", email:e.parameter.email||""});
        r = {status:"success"}; break;
      case "getData":
        r = _buildDashboard(); break;
      case "getSettings":
        r = _publicSettings(); break;
      case "updateSettings":
        r = _handleSettings(e.parameter); break;
      default:
        r = {status:"running", version:"3.0"};
    }
    return _json(r);
  } catch (err) {
    Logger.log("doGet: " + err); return _json({status:"error", message:err.toString()});
  }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type === "feedback")         _recordFeedback(d);
    else if (d.type === "rating")      _recordRating(d.rating, d.source||"website");
    else if (d.type === "competition") _recordCompetition(d);
    return _json({status:"success"});
  } catch (err) {
    Logger.log("doPost: " + err); return _json({status:"error", message:err.toString()});
  }
}

function _json(d) {
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}
function int(v) { return parseInt(v) || 0; }


// ═════════════════════════════════════════════════════════════
//  DATA RECORDING
// ═════════════════════════════════════════════════════════════

function _recordRating(rating, source) {
  if (rating < 1 || rating > 5) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = _sheet(ss, "Raw Data"); _setupRaw(s);
  const now = new Date(), tz = Session.getScriptTimeZone();
  s.appendRow([now, rating, source,
    Utilities.formatDate(now, tz, "yyyy-MM-dd"),
    Utilities.formatDate(now, tz, "HH:mm:ss"),
    Utilities.formatDate(now, tz, "EEEE")]);
}

function _recordFeedback(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = _sheet(ss, "Feedback"); _setupFeedback(s);
  const now = new Date(), tz = Session.getScriptTimeZone();
  s.appendRow([now, d.rating, d.name||"", d.contact||"", d.message||"",
    Utilities.formatDate(now, tz, "yyyy-MM-dd"), "New"]);
  _recordRating(d.rating, "feedback_form");
  _sendFeedbackEmail(d);
}

function _recordCompetition(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = _sheet(ss, "Competition"); _setupCompetition(s);
  const now = new Date(), tz = Session.getScriptTimeZone();
  s.appendRow([now, d.name||"", d.email||"",
    Utilities.formatDate(now, tz, "yyyy-MM-dd"), "Yes"]);
  _recordRating(5, "competition_google");
  _sendCompetitionEmail(d);
}


// ═════════════════════════════════════════════════════════════
//  EMAILS
// ═════════════════════════════════════════════════════════════

function _sendFeedbackEmail(d) {
  const to = _getSetting("managerEmail") || CONFIG.managerEmail;
  const urgent = d.rating <= 2;
  const subject = urgent
    ? "🚨 " + d.rating + "-Star Review — Action Needed"
    : "📝 " + d.rating + "-Star Feedback";
  const body =
    "BAR ASRA — Private Feedback\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
    "Rating:   " + "★".repeat(d.rating) + "☆".repeat(5-d.rating) + "  (" + d.rating + "/5)\n" +
    "Customer: " + (d.name||"Anonymous") + "\n" +
    "Contact:  " + (d.contact||"Not provided") + "\n\n" +
    "Message:\n" + (d.message||"(none)") + "\n\n" +
    (urgent ? "⚠️  LOW RATING — reach out today.\n\n" : "") +
    "✅ Captured privately. Did NOT go to Google.\n\n" +
    SpreadsheetApp.getActiveSpreadsheet().getUrl();
  try { MailApp.sendEmail({to, subject, body}); } catch(e) { Logger.log("Email: "+e); }
}

function _sendCompetitionEmail(d) {
  const to = _getSetting("managerEmail") || CONFIG.managerEmail;
  const prize = _getSetting("competitionPrize") || CONFIG.competitionPrize;
  const body =
    "BAR ASRA — New Competition Entry 🏆\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
    "A 5★ customer entered the \"" + prize + "\" draw\n" +
    "and was sent to leave a Google review.\n\n" +
    "Name:  " + d.name + "\nEmail: " + d.email + "\n\n" +
    SpreadsheetApp.getActiveSpreadsheet().getUrl();
  try { MailApp.sendEmail({to, subject: "🏆 Entry: " + d.name, body}); } catch(e) { Logger.log("Email: "+e); }
}

function sendDailyDigest() {
  const data = _buildDashboard(), s = data.stats;
  const to = _getSetting("managerEmail") || CONFIG.managerEmail;
  if (s.todayRatings === 0 && data.todayCompEntries === 0) return;
  const body =
    "BAR ASRA — Daily Digest\n━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
    "Today: " + s.todayRatings + " ratings, avg " + s.avgRating + "★\n" +
    "Competition entries today: " + data.todayCompEntries + "\n" +
    "Week: " + s.weekRatings + " (" + (s.weekChange>0?"+":"") + s.weekChange + "% WoW)\n" +
    "All time: " + s.totalRatings + " ratings, " + data.totalCompEntries + " entries\n\n" +
    SpreadsheetApp.getActiveSpreadsheet().getUrl();
  try { MailApp.sendEmail({to, subject: "📊 Digest: " + s.todayRatings + " ratings", body}); } catch(e) {}
}


// ═════════════════════════════════════════════════════════════
//  DASHBOARD DATA
// ═════════════════════════════════════════════════════════════

function _buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName("Raw Data"), fb = ss.getSheetByName("Feedback"), comp = ss.getSheetByName("Competition");
  const now = new Date(), tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const weekAgo = new Date(now.getTime()-7*86400000), monthAgo = new Date(now.getTime()-30*86400000), prevWeek = new Date(now.getTime()-14*86400000);

  const rawRows = raw && raw.getLastRow()>1 ? raw.getRange(2,1,raw.getLastRow()-1,6).getValues() : [];
  let total=0, today=0, week=0, month=0, prevWk=0, sumAll=0, sumWeek=0, sumPrev=0, googleRouted=0;
  const dist={1:0,2:0,3:0,4:0,5:0}, dailyC={}, dailyS={}, hourly=new Array(24).fill(0);
  const dayNames=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const dow={}; dayNames.forEach(d=>dow[d]=0);

  for (const row of rawRows) {
    const r=parseInt(row[1])||0; if(r<1||r>5) continue;
    const rd=new Date(row[0]), ds=row[3]||Utilities.formatDate(rd,tz,"yyyy-MM-dd"), src=row[2]||"", dn=row[5]||dayNames[rd.getDay()];
    total++; sumAll+=r; dist[r]++;
    if(ds===todayStr) today++;
    if(rd>=weekAgo){week++;sumWeek+=r;} if(rd>=monthAgo) month++;
    if(rd>=prevWeek&&rd<weekAgo){prevWk++;sumPrev+=r;}
    if(src==="competition_google") googleRouted++;
    if(rd>=monthAgo){dailyC[ds]=(dailyC[ds]||0)+1;dailyS[ds]=(dailyS[ds]||0)+r;}
    const h=rd.getHours(); if(!isNaN(h)) hourly[h]++;
    if(dow[dn]!==undefined) dow[dn]++;
  }

  const trend=[];
  for(let i=29;i>=0;i--){const d=new Date(now.getTime()-i*86400000),k=Utilities.formatDate(d,tz,"yyyy-MM-dd");
    trend.push({date:k,label:Utilities.formatDate(d,tz,"MMM dd"),count:dailyC[k]||0,avg:dailyC[k]?((dailyS[k]||0)/dailyC[k]).toFixed(1):null});}

  const weekChange=prevWk>0?Math.round(((week-prevWk)/prevWk)*100):(week>0?100:0);
  const avgWk=week>0?(sumWeek/week).toFixed(2):0, avgPv=prevWk>0?(sumPrev/prevWk).toFixed(2):0;
  const avgChange=prevWk>0?Math.round((avgWk-avgPv)/avgPv*100):0;
  const fiveTotal=dist[5]||0, convRate=fiveTotal>0?Math.round((googleRouted/fiveTotal)*100):0;

  const fbRows=fb&&fb.getLastRow()>1?fb.getRange(2,1,fb.getLastRow()-1,7).getValues():[];
  const recentFb=[]; for(let i=fbRows.length-1;i>=Math.max(0,fbRows.length-30);i--){
    const r=fbRows[i]; recentFb.push({rating:r[1],name:r[2]||"Anonymous",contact:r[3]||"",message:r[4]||"",date:r[5]||"",status:r[6]||"New"});}

  const compRows=comp&&comp.getLastRow()>1?comp.getRange(2,1,comp.getLastRow()-1,5).getValues():[];
  const compEntries=[]; let totalComp=0, todayComp=0;
  for(let i=compRows.length-1;i>=0;i--){const r=compRows[i],d=r[3]||""; totalComp++; if(d===todayStr) todayComp++;
    if(compEntries.length<50) compEntries.push({name:r[1]||"",email:r[2]||"",date:d,clicked:r[4]||""});}

  return {
    stats:{totalRatings:total,todayRatings:today,weekRatings:week,monthRatings:month,
      avgRating:total>0?(sumAll/total).toFixed(2):"0",avgWeekRating:avgWk,weekChange,avgChange,
      distribution:dist,googleRouted,fiveStarCount:fiveTotal,googleConversion:convRate},
    dailyTrend:trend, hourly, dayOfWeek:dow, recentFeedback:recentFb,
    compEntries, totalCompEntries:totalComp, todayCompEntries:todayComp,
    lastUpdated:now.toISOString(),
  };
}


// ═════════════════════════════════════════════════════════════
//  SETTINGS
// ═════════════════════════════════════════════════════════════

function _getSetting(k){const ss=SpreadsheetApp.getActiveSpreadsheet(),s=ss.getSheetByName("Settings");
  if(!s||s.getLastRow()<2) return null; const d=s.getDataRange().getValues();
  for(let i=1;i<d.length;i++){if(d[i][0]===k) return d[i][1];} return null;}

function _setSetting(k,v){const ss=SpreadsheetApp.getActiveSpreadsheet(); let s=ss.getSheetByName("Settings");
  if(!s){s=ss.insertSheet("Settings");_setupSettings(s);} const d=s.getDataRange().getValues();
  for(let i=1;i<d.length;i++){if(d[i][0]===k){s.getRange(i+1,2).setValue(v);return;}} s.appendRow([k,v]);}

function _publicSettings(){return {
  googleReviewUrl:_getSetting("googleReviewUrl")||CONFIG.defaultGoogleUrl,
  googleThreshold:parseInt(_getSetting("googleThreshold"))||CONFIG.defaultThreshold,
  competitionPrize:_getSetting("competitionPrize")||CONFIG.competitionPrize};}

function _handleSettings(p){
  if(p.password!==CONFIG.adminPassword) return {success:false,message:"Invalid password"};
  if(p.googleUrl) _setSetting("googleReviewUrl",p.googleUrl);
  if(p.threshold) _setSetting("googleThreshold",parseInt(p.threshold));
  if(p.email) _setSetting("managerEmail",p.email);
  if(p.prize) _setSetting("competitionPrize",p.prize);
  return {success:true,message:"Settings updated"};}
