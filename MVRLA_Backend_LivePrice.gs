/**
 * MVRLA-AI Material Search Backend (Live Price Edition)
 * ------------------------------------------------------------
 * Google Sheet tabs created/used:
 *   Materials, Price_History
 *
 * Materials columns:
 *   Part_Number | Item_Description | Machine | Location | Stock_Qty |
 *   Price | Category | Reference_Image | Uploaded_By | Uploaded_Date |
 *   Last_Modified
 *
 * Price_History columns:
 *   Part_Number | Old_Price | New_Price | Changed_By | Changed_Date |
 *   Stock_Value_After
 *
 * Actions:
 *   GET  ?action=ping
 *   GET  ?action=getAll                      -> all materials + live price & stock value
 *   GET  ?action=getPrice&part=PART_NO       -> live price for one part
 *   POST action=updatePrice                  -> update price (admin), single or bulk
 *   POST action=uploadMaterialImage
 *   POST action=deleteMaterialImage
 *   POST action=sendOtp          {email, purpose:'signup'|'reset'} -> emails 6-digit OTP
 *   POST action=register         {username,email,mobile,password,otp}
 *   POST action=login            {username,password}
 *   POST action=resetPassword    {email,otp,newPassword}
 *   POST action=listUsers        {adminCode}
 *   POST action=updateUserStatus {adminCode,username,status}
 *   POST action=deleteUser       {adminCode,username}
 */

const IMAGE_UPLOAD_ADMIN  = '100561611';
const PRICE_UPDATE_ADMIN  = '100561611';          // change if a different employee manages prices
const IMAGE_FOLDER_NAME   = 'MVRLA_Material_Images';
const PRICE_CURRENCY      = 'INR';

// OPTIONAL: point to a live price API. Use {part} as placeholder.
// API must return JSON like: {"price": 450}
// Example: 'https://your-supplier-api.com/price?part={part}&key=YOUR_KEY'
const EXTERNAL_PRICE_API  = '';

const MATERIAL_HEADERS = ['Part_Number','Item_Description','Machine','Location','Stock_Qty','Price','Category','Reference_Image','Uploaded_By','Uploaded_Date','Last_Modified'];
const PRICE_HISTORY_HEADERS = ['Part_Number','Old_Price','New_Price','Changed_By','Changed_Date','Stock_Value_After'];
const USER_HEADERS = ['Username','Email','Mobile','Salt','Password_Hash','Role','Created_Date','Last_Login','Status'];
const OTP_HEADERS  = ['Email','Purpose','OTP','Expires_At','Verified','Created_At'];
const PHOTO_HEADERS = ['Part_Number','Image_URL','Location','Updated_Date'];   // 🔑 photo + location NEVER lost
const OTP_EXPIRY_MIN = 5;          // OTP valid for 5 minutes

/* ============================ ROUTERS ============================ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || '').toLowerCase();
  try {
    if (action === 'ping')     return json({ ok: true, time: new Date().toISOString() });
    if (action === 'getall')   return json(getAllData());
    if (action === 'getprice') return json(getLivePrice(p.part || p.partNumber || ''));
    if (action === 'getimg')   return json(getImageBase64(p.id || ''));   // 📸 Material Lens
    return json({ ok: false, error: 'Unknown GET action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  let data = {};
  try { data = JSON.parse(e.postData.contents); } catch (err) {}
  const action = String(data.action || '').toLowerCase();
  try {
    if (action === 'uploadmaterialimage') return json(uploadMaterialImage(data));
    if (action === 'deletematerialimage') return json(deleteMaterialImage(data));
    if (action === 'updateprice')         return json(updatePrice(data));
    // ── USER AUTH + OTP ──
    if (action === 'sendotp')             return json(sendOtp(data));
    if (action === 'register')            return json(registerUser(data));
    if (action === 'login')               return json(loginUser(data));
    if (action === 'resetpassword')       return json(resetPassword(data));
    // ── ADMIN USER MANAGEMENT ──
    if (action === 'listusers')           return json(adminListUsers(data));
    if (action === 'updateuserstatus')    return json(adminUpdateUserStatus(data));
    if (action === 'deleteuser')          return json(adminDeleteUser(data));
    return json({ ok: false, error: 'Unknown POST action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ============================ SETUP ============================ */

function setupSheet() {
  const sh = sheet('Materials', MATERIAL_HEADERS);
  if (sh.getLastRow() === 1) {
    sh.appendRow(['262-030-00010','Cam Follower Bearing KR22','BM Stacking','Bay 1',14,450,'Bearing','','','','']);
    sh.appendRow(['262-030-00015','Pneumatic Cylinder 40x100mm','BM Stacking','Bay 1',6,2850,'Pneumatic','','','','']);
    sh.appendRow(['262-030-00022','Proximity Sensor PNP 5mm','BM Stacking','Bay 1',18,340,'Sensor','','','','']);
    sh.appendRow(['262-030-00045','Timing Belt 5M-600','BM Stacking','Bay 1',4,1280,'Belt','','','','']);
    sh.appendRow(['252-070-00020','EP2 Grease Cartridge 400g','Maintenance','Store',20,180,'Lubricant','','','','']);
  }
  sheet('Price_History', PRICE_HISTORY_HEADERS);
  sheet('Users', USER_HEADERS);
  sheet('OTP_Log', OTP_HEADERS);
  sheet('Photos', PHOTO_HEADERS);
  getOrCreateImagesFolder_();
  Logger.log('Setup complete: Materials, Price_History sheets and Drive folder ready.');
}

/** ⚠️ IMPORTANT — run this ONCE from the editor (Run ▶) to authorize OTP emails.
 *  It sends a test mail to your own Gmail and grants MailApp permission.
 *  Without this, OTP emails will FAIL when called from the web app. */
function authorizeEmail_() {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'MVRLA-AI Email Ready ✅', 'OTP email service is now active. You can send OTP from the app.');
  Logger.log('Email authorized — OTP ready!');
}

/* ============================ LIVE PRICE ============================ */

function getAllData() {
  const photos = getPhotoMap_();           // 🔑 separate safe photos map
  const materials = rows('Materials').map(r => mapMaterial(r, photos));
  const totalValue = materials.reduce((s, x) => s + (Number(x.v) || 0), 0);
  return {
    ok: true,
    time: new Date().toISOString(),
    currency: PRICE_CURRENCY,
    count: materials.length,
    totalStockValue: totalValue,          // live total cost of all stock
    materials: materials
  };
}

function mapMaterial(r, photos) {
  const qty   = Number(r.Stock_Qty) || 0;
  const price = Number(r.Price)     || 0;
  const p     = r.Part_Number || '';
  // Sheet image first; if the Materials data was re-entered/overwritten and the
  // Reference_Image cell got wiped, fall back to the permanent Photos tab.
  const sheetImg = normalizeImageUrl_(r.Reference_Image || '');
  const ph = photos && photos[p];
  const img = sheetImg || (ph ? ph.img : '');
  // 📍 Location ఖాళీగా ఉంటే (row re-add ayinappudu) Photos tab location ni vaadutam
  const loc = String(r.Location || '').trim() || (ph ? ph.loc : '');
  return {
    p:   p,
    n:   r.Item_Description || '',
    m:   r.Machine || '',
    l:   loc,
    q:   qty,
    pr:  price,                           // live unit price
    v:   qty * price,                     // live stock value for this part
    c:   r.Category || '',
    img: img,
    lm:  (r.Last_Modified instanceof Date) ? r.Last_Modified.toISOString() : (r.Last_Modified || '')
  };
}

/** GET live price for a single part. Uses external API if configured, else the sheet. */
function getLivePrice(partNumber) {
  const part = String(partNumber || '').trim();
  if (!part) return { ok: false, error: 'Missing part parameter' };

  const sh = sheet('Materials', MATERIAL_HEADERS);
  const head = ensureMaterialHeaders_(sh);
  const rowIndex = findMaterialRow_(sh, head, part);
  if (rowIndex < 2) return { ok: false, error: 'Part_Number not found: ' + part };

  const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  const qty   = Number(row[head.Stock_Qty]) || 0;
  const sheetPrice = Number(row[head.Price]) || 0;

  const ext = fetchExternalLivePrice_(part);   // optional external market price
  const livePrice = (ext && isFinite(ext.price)) ? ext.price : sheetPrice;

  return {
    ok: true,
    time: new Date().toISOString(),
    currency: PRICE_CURRENCY,
    source: (ext && isFinite(ext.price)) ? 'external' : 'sheet',
    part: {
      p:  part,
      n:  row[head.Item_Description] || '',
      q:  qty,
      pr: livePrice,                  // live unit price
      v:  qty * livePrice,            // live stock value
      sheetPrice: sheetPrice,
      lm: (row[head.Last_Modified] instanceof Date) ? row[head.Last_Modified].toISOString() : (row[head.Last_Modified] || '')
    }
  };
}

/**
 * POST action=updatePrice  (admin only)
 * Single: { action:'updatePrice', adminCode, partNumber, price, updatedBy }
 * Bulk:   { action:'updatePrice', adminCode, updatedBy, updates:[{partNumber, price}, ...] }
 */
function updatePrice(d) {
  if (String(d.adminCode || '') !== PRICE_UPDATE_ADMIN) {
    return { ok: false, error: 'Only employee ' + PRICE_UPDATE_ADMIN + ' can update prices.' };
  }

  const list = Array.isArray(d.updates) && d.updates.length
    ? d.updates
    : [{ partNumber: d.partNumber, price: d.price }];

  const sh = sheet('Materials', MATERIAL_HEADERS);
  const head = ensureMaterialHeaders_(sh);
  const hist = sheet('Price_History', PRICE_HISTORY_HEADERS);
  const now = new Date();
  const updatedBy = d.updatedBy || PRICE_UPDATE_ADMIN;

  const results = [];
  let okCount = 0;

  list.forEach(u => {
    const part = String(u.partNumber || '').trim();
    const newPrice = Number(u.price);

    if (!part)                       { results.push({ partNumber: part, ok: false, error: 'Missing Part_Number' }); return; }
    if (!isFinite(newPrice) || newPrice < 0) { results.push({ partNumber: part, ok: false, error: 'Invalid price' }); return; }

    const rowIndex = findMaterialRow_(sh, head, part);
    if (rowIndex < 2)                { results.push({ partNumber: part, ok: false, error: 'Part_Number not found' }); return; }

    const oldPrice = Number(sh.getRange(rowIndex, head.Price + 1).getValue()) || 0;
    const qty      = Number(sh.getRange(rowIndex, head.Stock_Qty + 1).getValue()) || 0;

    sh.getRange(rowIndex, head.Price + 1).setValue(newPrice);
    sh.getRange(rowIndex, head.Last_Modified + 1).setValue(now);

    hist.appendRow([part, oldPrice, newPrice, updatedBy, now, qty * newPrice]);

    results.push({
      partNumber: part,
      ok: true,
      oldPrice: oldPrice,
      newPrice: newPrice,
      stockValue: qty * newPrice,     // live stock value after update
      currency: PRICE_CURRENCY
    });
    okCount++;
  });

  const totalValue = rows('Materials').map(mapMaterial).reduce((s, x) => s + (Number(x.v) || 0), 0);
  return {
    ok: okCount > 0,
    updated: okCount,
    failed: results.length - okCount,
    time: now.toISOString(),
    currency: PRICE_CURRENCY,
    totalStockValue: totalValue,      // live total cost of all stock after update
    results: results.length === 1 && !Array.isArray(d.updates) ? results[0] : results
  };
}

/** GET ?action=getimg&id=DRIVE_FILE_ID — image ni base64 ga istundi (Lens matching cross-browser safe) */
function getImageBase64(id) {
  id = String(id || '').trim();
  if (!id) return { ok: false, error: 'Missing id' };
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    const mime = blob.getContentType() || 'image/jpeg';
    return { ok: true, mime: mime, dataUrl: 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  } catch (err) {
    return { ok: false, error: 'Image not found: ' + id };
  }
}

/** Optional external live-price lookup (returns null unless EXTERNAL_PRICE_API is set). */
function fetchExternalLivePrice_(partNumber) {
  if (!EXTERNAL_PRICE_API) return null;
  try {
    const url = EXTERNAL_PRICE_API.replace('{part}', encodeURIComponent(partNumber));
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const j = JSON.parse(res.getContentText());
    const price = Number(j.price);
    return isFinite(price) ? { price: price } : null;
  } catch (err) {
    return null;
  }
}

/* ============================ USER AUTH + OTP ============================ */

function sha256_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function ensureCols_(sh, required) {
  let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  required.forEach(h => {
    if (headers.indexOf(h) === -1) {
      sh.getRange(1, headers.length + 1).setValue(h);
      headers.push(h);
    }
  });
  const out = {};
  headers.forEach((h, i) => out[h] = i);
  return out;
}

function findUserRow_(sh, head, username) {
  const values = sh.getDataRange().getValues();
  const want = String(username || '').trim().toLowerCase();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][head.Username]).trim().toLowerCase() === want) return r + 1;
  }
  return -1;
}

function findUserByEmail_(sh, head, email) {
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][head.Email]).trim().toLowerCase() === email) return r + 1;
  }
  return -1;
}

/** POST action=sendOtp  {email, purpose:'signup'|'reset'} — emails a 6-digit OTP (free via MailApp) */
function sendOtp(d) {
  const email = String(d.email || '').trim().toLowerCase();
  const purpose = String(d.purpose || 'signup').toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'Enter a valid email address' };

  const users = sheet('Users', USER_HEADERS);
  const head = ensureCols_(users, USER_HEADERS);
  const rowByEmail = findUserByEmail_(users, head, email);

  if (purpose === 'signup' && rowByEmail >= 2) return { ok: false, error: 'This email is already registered. Please login.' };
  if (purpose === 'reset'  && rowByEmail < 2)  return { ok: false, error: 'No account found with this email.' };

  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const now = new Date();
  const expires = new Date(now.getTime() + OTP_EXPIRY_MIN * 60000);
  sheet('OTP_Log', OTP_HEADERS).appendRow([email, purpose, otp, expires, '', now]);

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'MVRLA-AI ' + (purpose === 'reset' ? 'Password Reset' : 'Signup') + ' OTP: ' + otp,
      body: 'Hello,\n\nYour MVRLA-AI verification code is:\n\n    ' + otp + '\n\nIt is valid for ' + OTP_EXPIRY_MIN + ' minutes. Do not share it with anyone.\n\n— MVRLA-AI Material Search (Amara Raja Engineering Stores)'
    });
  } catch (err) {
    return { ok: false, error: 'Email sending failed: ' + err };
  }
  return { ok: true, message: 'OTP sent to ' + email, expiresInSec: OTP_EXPIRY_MIN * 60 };
}

/** Verify latest matching, unused, unexpired OTP. Marks it used on success. */
function checkOtp_(email, purpose, otp) {
  const log = sheet('OTP_Log', OTP_HEADERS);
  const values = log.getDataRange().getValues();
  if (values.length < 2) return false;
  const now = new Date();
  for (let r = values.length - 1; r >= 1; r--) {
    const row = values[r];
    if (String(row[0]).trim().toLowerCase() === email &&
        String(row[1]).toLowerCase() === purpose &&
        String(row[2]).trim() === String(otp).trim() &&
        !row[4] &&
        new Date(row[3]) > now) {
      log.getRange(r + 1, 5).setValue('YES');   // one-time use
      return true;
    }
  }
  return false;
}

/** POST action=register  {username, email, mobile, password, otp} */
function registerUser(d) {
  const username = String(d.username || '').trim();
  const email    = String(d.email || '').trim().toLowerCase();
  const mobile   = String(d.mobile || '').trim();
  const password = String(d.password || '');
  if (username.length < 3) return { ok: false, error: 'Username must be at least 3 characters' };
  if (password.length < 4) return { ok: false, error: 'Password must be at least 4 characters' };
  if (!/^[0-9]{10}$/.test(mobile)) return { ok: false, error: 'Enter a valid 10-digit mobile number' };
  if (!checkOtp_(email, 'signup', d.otp)) return { ok: false, error: 'Invalid or expired OTP' };

  const sh = sheet('Users', USER_HEADERS);
  const head = ensureCols_(sh, USER_HEADERS);
  if (findUserRow_(sh, head, username) >= 2) return { ok: false, error: 'Username already taken. Choose another.' };
  if (findUserByEmail_(sh, head, email) >= 2) return { ok: false, error: 'Email already registered. Please login.' };

  const salt = Utilities.getUuid();
  const hash = sha256_(salt + password);
  const now = new Date();
  sh.appendRow([username, email, mobile, salt, hash, 'user', now, now, 'Active']);
  return { ok: true, message: 'Account created', user: { name: username, email: email, role: 'user' } };
}

/** POST action=login  {username, password} */
function loginUser(d) {
  const username = String(d.username || '').trim();
  const password = String(d.password || '');
  const sh = sheet('Users', USER_HEADERS);
  const head = ensureCols_(sh, USER_HEADERS);
  const rowIndex = findUserRow_(sh, head, username);
  if (rowIndex < 2) return { ok: false, error: 'Username not found. Please sign up first.' };

  const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  if (String(row[head.Status]).toLowerCase() === 'blocked') {
    return { ok: false, error: 'Account blocked. Contact admin (100561611).' };
  }
  const hash = sha256_(String(row[head.Salt]) + password);
  if (hash !== String(row[head.Password_Hash])) return { ok: false, error: 'Wrong password' };

  sh.getRange(rowIndex, head.Last_Login + 1).setValue(new Date());   // login data saved to sheet
  return { ok: true, user: { name: row[head.Username], email: row[head.Email], mobile: row[head.Mobile], role: row[head.Role] } };
}

/** POST action=resetPassword  {email, otp, newPassword} */
function resetPassword(d) {
  const email = String(d.email || '').trim().toLowerCase();
  const password = String(d.newPassword || '');
  if (password.length < 4) return { ok: false, error: 'Password must be at least 4 characters' };
  if (!checkOtp_(email, 'reset', d.otp)) return { ok: false, error: 'Invalid or expired OTP' };

  const sh = sheet('Users', USER_HEADERS);
  const head = ensureCols_(sh, USER_HEADERS);
  const rowIndex = findUserByEmail_(sh, head, email);
  if (rowIndex < 2) return { ok: false, error: 'No account with this email' };

  const salt = Utilities.getUuid();
  sh.getRange(rowIndex, head.Salt + 1).setValue(salt);
  sh.getRange(rowIndex, head.Password_Hash + 1).setValue(sha256_(salt + password));
  return { ok: true, message: 'Password reset successful. Please login.' };
}

/* ============================ ADMIN USER MANAGEMENT ============================ */

function adminCheck_(d) { return String(d.adminCode || '') === PRICE_UPDATE_ADMIN; }

/** POST action=listUsers  {adminCode} */
function adminListUsers(d) {
  if (!adminCheck_(d)) return { ok: false, error: 'Admin only' };
  return {
    ok: true,
    users: rows('Users').map(r => ({
      name: r.Username || '', email: r.Email || '', mobile: r.Mobile || '',
      role: r.Role || 'user', status: r.Status || 'Active',
      created: (r.Created_Date instanceof Date) ? r.Created_Date.toISOString() : (r.Created_Date || ''),
      lastLogin: (r.Last_Login instanceof Date) ? r.Last_Login.toISOString() : (r.Last_Login || '')
    }))
  };
}

/** POST action=updateUserStatus  {adminCode, username, status:'Active'|'Blocked'} */
function adminUpdateUserStatus(d) {
  if (!adminCheck_(d)) return { ok: false, error: 'Admin only' };
  const sh = sheet('Users', USER_HEADERS);
  const head = ensureCols_(sh, USER_HEADERS);
  const rowIndex = findUserRow_(sh, head, d.username);
  if (rowIndex < 2) return { ok: false, error: 'User not found' };
  const status = String(d.status) === 'Blocked' ? 'Blocked' : 'Active';
  sh.getRange(rowIndex, head.Status + 1).setValue(status);
  return { ok: true, username: String(d.username), status: status };
}

/** POST action=deleteUser  {adminCode, username} */
function adminDeleteUser(d) {
  if (!adminCheck_(d)) return { ok: false, error: 'Admin only' };
  const sh = sheet('Users', USER_HEADERS);
  const head = ensureCols_(sh, USER_HEADERS);
  const rowIndex = findUserRow_(sh, head, d.username);
  if (rowIndex < 2) return { ok: false, error: 'User not found' };
  sh.deleteRow(rowIndex);
  return { ok: true, message: 'User deleted' };
}

/* ============================ IMAGES (unchanged) ============================ */

function uploadMaterialImage(d) {
  if (String(d.adminCode || '') !== IMAGE_UPLOAD_ADMIN) {
    return { ok: false, error: 'Only employee ' + IMAGE_UPLOAD_ADMIN + ' can upload images.' };
  }
  const partNumber = String(d.partNumber || '').trim();
  if (!partNumber) return { ok: false, error: 'Missing Part_Number' };
  if (!d.image) return { ok: false, error: 'Missing image data' };

  const sh = sheet('Materials', MATERIAL_HEADERS);
  const head = ensureMaterialHeaders_(sh);
  const rowIndex = findMaterialRow_(sh, head, partNumber);
  if (rowIndex < 2) return { ok: false, error: 'Part_Number not found: ' + partNumber };

  const oldUrl = sh.getRange(rowIndex, head.Reference_Image + 1).getValue();
  const blob = dataUrlToBlob_(d.image);
  if (!blob) return { ok: false, error: 'Invalid image format' };

  const folder = getOrCreateImagesFolder_();
  const file = folder.createFile(blob);
  file.setName(partNumber + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';

  const now = new Date();
  // 🔑 FIX: అదే Part_Number ఉన్న అన్ని rows (duplicates) కి photo set చేయి
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][head.Part_Number]).trim() === partNumber) {
      sh.getRange(r + 1, head.Reference_Image + 1).setValue(url);
    }
  }
  // 🔑 FIX: ప్రత్యేక Photos tab లో కూడా save — photo + location రెండూ safe
  setPhoto_(partNumber, url, String(d.location || '').trim());
  sh.getRange(rowIndex, head.Uploaded_By + 1).setValue(d.uploadedBy || IMAGE_UPLOAD_ADMIN);
  sh.getRange(rowIndex, head.Uploaded_Date + 1).setValue(now);
  sh.getRange(rowIndex, head.Last_Modified + 1).setValue(now);

  trashDriveFileFromUrl_(oldUrl);

  return { ok: true, url: url, fileId: file.getId() };
}

function deleteMaterialImage(d) {
  if (String(d.adminCode || '') !== IMAGE_UPLOAD_ADMIN) {
    return { ok: false, error: 'Only employee ' + IMAGE_UPLOAD_ADMIN + ' can delete images.' };
  }
  const partNumber = String(d.partNumber || '').trim();
  if (!partNumber) return { ok: false, error: 'Missing Part_Number' };

  const sh = sheet('Materials', MATERIAL_HEADERS);
  const head = ensureMaterialHeaders_(sh);
  const rowIndex = findMaterialRow_(sh, head, partNumber);
  if (rowIndex < 2) return { ok: false, error: 'Part_Number not found: ' + partNumber };

  const oldUrl = sh.getRange(rowIndex, head.Reference_Image + 1).getValue();
  trashDriveFileFromUrl_(oldUrl);
  // 🔑 అన్ని duplicate rows నుండి క్లియర్ + Photos tab నుండి కూడా తీసేయి
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][head.Part_Number]).trim() === partNumber) {
      sh.getRange(r + 1, head.Reference_Image + 1).setValue('');
    }
  }
  setPhoto_(partNumber, '');
  sh.getRange(rowIndex, head.Last_Modified + 1).setValue(new Date());
  return { ok: true };
}

/* ============================ HELPERS ============================ */

/** 🔑 Photos tab — Part_Number → {img, loc}. Materials sheet reset ayina photo + location safe. */
function getPhotoMap_() {
  const map = {};
  rows('Photos').forEach(r => {
    const p = String(r.Part_Number || '').trim();
    const u = String(r.Image_URL || '').trim();
    if (p && u) map[p] = { img: normalizeImageUrl_(u), loc: String(r.Location || '').trim() };
  });
  return map;
}

/** Save (or clear, if url='') a part's photo (+ optional location) in the Photos tab. */
function setPhoto_(partNumber, url, location) {
  const sh = sheet('Photos', PHOTO_HEADERS);
  const head = ensureCols_(sh, PHOTO_HEADERS);   // old Photos tab ki Location column auto-add avutundi
  const values = sh.getDataRange().getValues();
  const now = new Date();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][head.Part_Number]).trim() === partNumber) {
      if (url) {
        sh.getRange(r + 1, head.Image_URL + 1).setValue(url);
        if (location) sh.getRange(r + 1, head.Location + 1).setValue(location);
        sh.getRange(r + 1, head.Updated_Date + 1).setValue(now);
      } else {
        sh.deleteRow(r + 1);   // explicit admin delete only
      }
      return;
    }
  }
  if (url) {
    const row = new Array(Object.keys(head).length).fill('');
    row[head.Part_Number] = partNumber;
    row[head.Image_URL] = url;
    row[head.Location]  = location || '';
    row[head.Updated_Date] = now;
    sh.appendRow(row);
  }
}

function dataUrlToBlob_(dataUrl) {
  const match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9+.-]+);base64,(.*)$/);
  if (!match) return null;
  const mime = match[1];
  const bytes = Utilities.base64Decode(match[2].replace(/\s/g,''));
  return Utilities.newBlob(bytes, mime, 'material_image.jpg');
}

function getOrCreateImagesFolder_() {
  const folders = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

function normalizeImageUrl_(url) {
  if (!url) return '';
  const s = String(url).trim();
  const match = s.match(/[-\w]{25,}/);
  if (match && s.indexOf('drive.google') !== -1) {
    return 'https://drive.google.com/thumbnail?id=' + match[0] + '&sz=w1000';
  }
  return s;
}

function trashDriveFileFromUrl_(url) {
  try {
    const s = String(url || '');
    const match = s.match(/[-\w]{25,}/);
    if (match && s.indexOf('drive.google') !== -1) DriveApp.getFileById(match[0]).setTrashed(true);
  } catch (err) {}
}

function ensureMaterialHeaders_(sh) {
  const required = MATERIAL_HEADERS;
  let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  required.forEach(h => {
    if (headers.indexOf(h) === -1) {
      sh.getRange(1, headers.length + 1).setValue(h);
      headers.push(h);
    }
  });
  const out = {};
  headers.forEach((h, i) => out[h] = i);
  return out;
}

function findMaterialRow_(sh, head, partNumber) {
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][head.Part_Number]).trim() === partNumber) return r + 1;
  }
  return -1;
}

function sheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
