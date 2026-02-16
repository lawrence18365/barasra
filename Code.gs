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
//  DASHBOARD DATA — GM Focused
// ═════════════════════════════════════════════════════════════

function _buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fb = ss.getSheetByName("Feedback"), comp = ss.getSheetByName("Competition"), raw = ss.getSheetByName("Raw Data");
  const now = new Date(), tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const weekAgo = new Date(now.getTime()-7*86400000), monthAgo = new Date(now.getTime()-30*86400000);

  // Feedback data (actual customer submissions, not every star tap)
  const fbRows = fb && fb.getLastRow()>1 ? fb.getRange(2,1,fb.getLastRow()-1,7).getValues() : [];
  let fbTotal=0, fbToday=0, fbWeek=0, fbMonth=0, fbSum=0, fbWeekSum=0, unreadCount=0, needsAttention=0;
  const dist={1:0,2:0,3:0,4:0,5:0}, recentFb=[];
  
  for(let i=fbRows.length-1;i>=0;i--){
    const r=fbRows[i];
    const rating=parseInt(r[1])||0;
    const dateStr=r[5]||"";
    const status=r[6]||"New";
    const rd=new Date(r[0]);
    
    if(rating<1||rating>5) continue;
    fbTotal++; fbSum+=rating; dist[rating]++;
    if(dateStr===todayStr) fbToday++;
    if(rd>=weekAgo){fbWeek++; fbWeekSum+=rating;}
    if(rd>=monthAgo) fbMonth++;
    if(status==="New") unreadCount++;
    if(rating<=2 && status==="New") needsAttention++;
    
    if(recentFb.length<30){
      recentFb.push({rating:rating,name:r[2]||"Anonymous",contact:r[3]||"",message:r[4]||"",date:dateStr,status:status});
    }
  }

  // Competition entries
  const compRows = comp && comp.getLastRow()>1 ? comp.getRange(2,1,comp.getLastRow()-1,5).getValues() : [];
  const compEntries=[]; let totalComp=0, todayComp=0, compWeek=0, compMonth=0;
  for(let i=compRows.length-1;i>=0;i--){
    const r=compRows[i];
    const d=r[3]||"";
    let rd;
    try { rd = new Date(r[0]); } catch(e) { rd = new Date(); }
    if(isNaN(rd.getTime())) rd = new Date();
    totalComp++; 
    if(d===todayStr) todayComp++;
    if(rd>=weekAgo) compWeek++;
    if(rd>=monthAgo) compMonth++;
    if(compEntries.length<50) compEntries.push({name:String(r[1]||""),email:String(r[2]||""),date:String(d),clicked:String(r[4]||"")});
  }

  // Raw data for Google conversion tracking only
  const rawRows = raw && raw.getLastRow()>1 ? raw.getRange(2,1,raw.getLastRow()-1,6).getValues() : [];
  let googleRouted=0, fiveStarTaps=0;
  for(const row of rawRows){
    const src=row[2]||"", rating=parseInt(row[1])||0;
    if(src==="competition_google") googleRouted++;
    if(rating===5) fiveStarTaps++;
  }

  // GM Key Metrics
  const avgRating = fbTotal>0 ? (fbSum/fbTotal).toFixed(2) : "0";
  const avgWeek = fbWeek>0 ? (fbWeekSum/fbWeek).toFixed(2) : "0";
  const googleConvRate = fiveStarTaps>0 ? Math.round((googleRouted/fiveStarTaps)*100) : 0;
  const totalInteractions = fbTotal + totalComp;
  
  return {
    stats:{
      // Key GM Metrics
      totalFeedback: fbTotal,
      todayFeedback: fbToday,
      weekFeedback: fbWeek,
      monthFeedback: fbMonth,
      avgRating: avgRating,
      avgWeekRating: avgWeek,
      
      // Action items
      unreadCount: unreadCount,
      needsAttention: needsAttention, // 1-2 star + unread
      
      // Competition
      totalCompEntries: totalComp,
      todayCompEntries: todayComp,
      weekCompEntries: compWeek,
      monthCompEntries: compMonth,
      
      // Google conversion
      fiveStarTaps: fiveStarTaps,
      googleRouted: googleRouted,
      googleConversion: googleConvRate,
      
      // Distribution (from actual feedback, not taps)
      distribution: dist,
    },
    recentFeedback: recentFb,
    compEntries: compEntries,
    lastUpdated: now.toISOString(),
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
