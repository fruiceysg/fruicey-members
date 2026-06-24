// ============================================================
// FRUICEY STAMP CARD — Google Apps Script Backend
// Add this to your existing Fruicey Members Apps Script
// as a NEW FILE called StampCard.gs
// ============================================================

const STAMP_SH  = 'StampCards';
const STAMP_LOG_SH = 'StampLog';
const STAMP_HEADERS = ['id','first','last','phone','email','bday','since','stamps','cardsCompleted','lastStamp','stampLog','redeemLog'];
const STAMPLOG_HEADERS = ['token','phone','created','used','expired'];

// ── Add to doGet routing ──────────────────────────────────────
// if (action === 'getStampCard')     return cors(getStampCard(e.parameter.phone));
// if (action === 'getAllStampCards')  return cors(getAllStampCards());

// ── Add to doPost routing ─────────────────────────────────────
// if (action === 'createStampCard')   return cors(createStampCard(data));
// if (action === 'generateStampToken')return cors(generateStampToken(data));
// if (action === 'redeemStampToken')  return cors(redeemStampToken(data));
// if (action === 'redeemFreeCard')    return cors(redeemFreeCard(data));
// if (action === 'editStampCard')     return cors(editStampCard(data));

// ── READ ──────────────────────────────────────────────────────
function getStampCard(phone) {
  const members = sheetRows(STAMP_SH);
  const np = p => (p||'').replace(/\D/g,'').replace(/^65(?=\d{8}$)/,'');
  const member = members.find(m => np(m.phone) === np(phone||''));
  if (!member) return {error:'not_found'};
  return {member};
}

function getAllStampCards() {
  return {members: sheetRows(STAMP_SH)};
}

// ── WRITE ─────────────────────────────────────────────────────
function createStampCard(data) {
  const members = sheetRows(STAMP_SH);
  const np = p => (p||'').replace(/\D/g,'').replace(/^65(?=\d{8}$)/,'');
  if (members.find(m => np(m.phone) === np(data.phone))) return {error:'phone_exists'};
  if (members.find(m => m.email.toLowerCase() === (data.email||'').toLowerCase())) return {error:'email_exists'};

  ensureSheet(STAMP_SH, STAMP_HEADERS);
  const member = {
    id: 'sc_'+Date.now(), first:data.first, last:data.last,
    phone:data.phone, email:data.email, bday:data.bday||'',
    since:today(), stamps:'0', cardsCompleted:'0',
    lastStamp:'', stampLog:'', redeemLog:''
  };
  appendRow(STAMP_SH, member, STAMP_HEADERS);
  return {success:true, member};
}

function generateStampToken(data) {
  const np = p => (p||'').replace(/\D/g,'').replace(/^65(?=\d{8}$)/,'');
  const members = sheetRows(STAMP_SH);
  const member = members.find(m => np(m.phone) === np(data.phone||''));
  if (!member) return {error:'not_found'};
  if (parseInt(member.stamps||0) >= 9) return {error:'already_complete'};

  ensureSheet(STAMP_LOG_SH, STAMPLOG_HEADERS);

  // Generate secure token
  const token = Utilities.getUuid().replace(/-/g,'').substring(0,16).toUpperCase();
  const created = new Date().toISOString();

  appendRow(STAMP_LOG_SH, {token, phone:data.phone, created, used:'false', expired:'false'}, STAMPLOG_HEADERS);

  // Auto-expire after 90 seconds via timestamp check (done at redemption time)
  return {success:true, token, phone:data.phone};
}

function redeemStampToken(data) {
  const token = data.token;
  const np = p => (p||'').replace(/\D/g,'').replace(/^65(?=\d{8}$)/,'');

  // Find token in log
  const logs = sheetRows(STAMP_LOG_SH);
  const log  = logs.find(l => l.token === token);
  if (!log) return {error:'invalid'};
  if (log.used === 'true') return {error:'already_used'};

  // Check expiry (90 seconds)
  const created = new Date(log.created);
  const now     = new Date();
  if ((now - created) > 90000) {
    updateRowById(STAMP_LOG_SH, token, {expired:'true'});
    return {error:'expired'};
  }

  // Check phone matches
  if (np(log.phone) !== np(data.phone||'')) return {error:'invalid'};

  // Mark token as used
  updateRowByField(STAMP_LOG_SH, 'token', token, {used:'true'});

  // Add stamp to member
  const members = sheetRows(STAMP_SH);
  const member  = members.find(m => np(m.phone) === np(data.phone||''));
  if (!member) return {error:'member_not_found'};

  let stamps = parseInt(member.stamps||0) + 1;
  const dateStr = today();
  const newLog  = (member.stampLog||'') + dateStr + ','+token+'|';
  let cardsCompleted = parseInt(member.cardsCompleted||0);
  let redeemLog = member.redeemLog||'';

  // If card complete, auto-reset
  if (stamps >= 9) {
    stamps = 0;
    cardsCompleted++;
    redeemLog += dateStr+'|';
  }

  updateRowByField(STAMP_SH, 'phone', np(member.phone), {
    stamps: String(stamps),
    cardsCompleted: String(cardsCompleted),
    lastStamp: dateStr,
    stampLog: newLog,
    redeemLog: redeemLog
  });

  return {success:true, stamps, cardsCompleted};
}

// Helper: update by any field value
function updateRowByField(name, field, value, updates) {
  const s       = sh(name);
  if (!s) return false;
  const data    = s.getDataRange().getValues();
  const headers = data[0];
  const col     = headers.indexOf(field);
  const np = p => (p||'').replace(/\D/g,'').replace(/^65(?=\d{8}$)/,'');
  for (let i = 1; i < data.length; i++) {
    if (np(String(data[i][col])) === np(String(value))) {
      Object.keys(updates).forEach(key => {
        const c = headers.indexOf(key);
        if (c !== -1) s.getRange(i+1, c+1).setValue(updates[key]);
      });
      return true;
    }
  }
  return false;
}

// ── EDIT STAMPS (admin correction) ───────────────────────────
function editStampCard(data) {
  const np = p => (p||'').replace(/\D/g,'').replace(/^65(?=\d{8}$)/,'');
  const members = sheetRows(STAMP_SH);
  const member  = members.find(m => np(m.phone) === np(data.phone||''));
  if (!member) return {error:'not_found'};

  const updates = {
    stamps:         String(Math.max(0, Math.min(9, parseInt(data.stamps)||0))),
    cardsCompleted: String(Math.max(0, parseInt(data.cardsCompleted)||0))
  };

  // Append edit note to stampLog
  const note = '\n[EDIT '+today()+': stamps='+updates.stamps+', completed='+updates.cardsCompleted+(data.reason?', reason: '+data.reason:'')+']';
  updates.stampLog = (member.stampLog||'') + note;

  updateRowByField(STAMP_SH, 'phone', np(member.phone), updates);
  return {success: true};
}

// ── Setup ─────────────────────────────────────────────────────
function setupStampSheets() {
  ensureSheet(STAMP_SH,     STAMP_HEADERS);
  ensureSheet(STAMP_LOG_SH, STAMPLOG_HEADERS);
  Logger.log('Stamp sheets ready!');
}
