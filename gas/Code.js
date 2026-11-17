/**
 * 성장일기 웹앱 - 서버 코드 (Apps Script 프로젝트의 Code.gs)
 * VERSION: V3.4
 * 버전: v3.4
 * ※ GitHub 저장소 myungseok83/Familydiary 의 gas/Code.js 가 바뀌면 자동으로 Apps Script 에 배포됨
 * ※ 비밀 값(가족 비밀번호, GIPHY 키)은 코드에 쓰지 않고 '스크립트 속성'에 넣음 (저장소가 공개라서)
 *    Apps Script 편집기 → ⚙️ 프로젝트 설정 → 스크립트 속성
 *      FAMILY_PASSWORD : 가족 비밀번호 (없으면 비밀번호 없이 입장)
 *      GIPHY_API_KEY   : GIPHY API 키 (없으면 GIF 검색 안 됨)
 * 데이터: 기존 '성장일기' 구글 시트를 그대로 사용 (앱시트와 동시 사용 가능)
 */
const APP_VERSION = '3.4';
const SPREADSHEET_ID = '1cdRe4DVc41KBOTPiOTdqeMqnua6k8BGJUEsysKZuEEw';

// 스크립트 속성 읽기 — 가족 비밀번호(FAMILY_PASSWORD)는 여기서 가져옴 (없으면 비밀번호 없이 입장)
function prop_(k) { try { return PropertiesService.getScriptProperties().getProperty(k) || ''; } catch (e) { return ''; } }
function familyPw_() { return prop_('FAMILY_PASSWORD'); }

// 아이들 생년월일 (D+ 계산용), 성별 (성장도표 비교용: F 여아, M 남아)
const CHILDREN = [
  { name: '서아', birth: '2020-02-21', sex: 'F' },
  { name: '서준', birth: '2025-01-14', sex: 'M' }
];

const TZ = 'Asia/Seoul';

// 채팅 읽음 숫자에서 뺄 사람 (아직 채팅을 하지 않는 아이 등) — 이메일을 넣으면 '안 읽은 사람 수'에서 빠짐
const CHAT_EXCLUDE = ['seojoon@gmail.com'];

// GIF 검색 (GIPHY) — API 키는 스크립트 속성 GIPHY_API_KEY 에 넣기 (developers.giphy.com 에서 무료 발급)
function giphyKey_() { return prop_('GIPHY_API_KEY'); }

// 폰 알림 (ntfy 앱) 설정
const NTFY_ENABLED = true;              // false 로 바꾸면 알림 끔
const NTFY_SERVER = 'https://ntfy.sh';
const NTFY_SHOW_CONTENT = true;         // false 면 내용 없이 '새 소식이 있어요'만 표시
const NTFY_FROM_SERVER = false;         // 구글 서버에서는 ntfy 연결이 막혀 있어서, 알림은 화면(폰)에서 직접 보냄
const SHEET = { POST: '게시판', COMMENT: '댓글', USER: '유저', GROWTH: '성장곡선', CHAT: '챗', MEDIA: 'Multi file upload' };
const PATH_COLS = ['이미지', '이미지2', '이미지3', '이미지4', '이미지5'];
const URL_COLS = ['Image URL', 'Image2 URL', 'Image3 URL', 'Image4 URL', 'Image5 URL'];

/* ================= 웹앱 진입 ================= */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('성장일기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= 데이터 API (GitHub 화면용) =================
   GitHub Pages 의 index.html 이 이 주소(/exec)로 POST 요청을 보내면
   아래 목록의 함수를 실행하고 결과를 JSON 으로 돌려줌.
   ※ 배포 설정: 실행 = 나, 액세스 = '모든 사용자' (로그인 필요 없는 쪽) */
const API_FUNCS = {
  hello, init, reload, getImage,
  addPost, updatePost, deletePost, likePost, addComment, deleteComment,
  getMedia, addMedia, registerMedia, notifyMedia, deleteMedia, scanOrphans, cleanOrphans,
  addGrowth, deleteGrowth,
  getChats, sendChat, startUpload, uploadChunk, sendChatFile, getFileData, deleteChat,
  getNotifyInfo, testNotify, getNotifyTargets, sendChatGif, getGifKey
};
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API_FUNCS[req.fn];
    if (!fn) throw new Error('알 수 없는 요청: ' + req.fn);
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= 공통 도구 ================= */
function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('시트를 찾을 수 없습니다: ' + name);
  return sh;
}
function check_(pw) {
  const fp = familyPw_();
  if (fp && String(pw || '') !== fp) throw new Error('PW_ERROR');
}
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function lock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
// 날짜를 'yyyy-MM-dd HH:mm:ss' 로 통일
function time_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  const s = String(v || '');
  const m = s.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D+?(오전|오후|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/i);
  if (!m) return s;
  let h = Number(m[5]);
  if (/오후|PM/i.test(m[4] || '') && h < 12) h += 12;
  if (/오전|AM/i.test(m[4] || '') && h === 12) h = 0;
  const p = n => ('0' + (n || 0)).slice(-2);
  return m[1] + '-' + p(m[2]) + '-' + p(m[3]) + ' ' + p(h) + ':' + p(m[6]) + ':' + p(m[7]);
}
function date_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = String(v || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : String(v || '');
}
// 화면에서 받은 'yyyy-MM-ddTHH:mm' (한국시간) → Date
function parseLocal_(v) {
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/);
  if (!m) return null;
  return new Date(m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] + ':00+09:00');
}
function daysSince_(birth, dateStr) {
  return Math.round((Date.parse(dateStr.slice(0, 10)) - Date.parse(birth)) / 86400000);
}
function driveId_(s) {
  s = String(s || '');
  const m = s.match(/[?&]id=([\w-]{20,})/) || s.match(/\/d\/([\w-]{20,})/);
  return m ? m[1] : '';
}
function imgRef_(pathVal, urlVal) {
  const id = driveId_(urlVal) || driveId_(pathVal);
  if (id) return { id: id };
  if (pathVal) return { path: String(pathVal) };
  return null;
}

// 시트 → 객체 배열 (헤더 이름 기준)
function table_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(h => String(h).trim());
  const rows = values.map((r, i) => {
    const o = { _row: i + 2 };
    headers.forEach((h, j) => { if (h) o[h] = r[j]; });
    return o;
  });
  return { sh: sh, headers: headers, rows: rows };
}
function lastRowA_(sh) {
  const n = Math.max(sh.getLastRow(), 1);
  const v = sh.getRange(1, 1, n, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '') return i + 1;
  return 1;
}
// 헤더 이름으로 값 쓰기 (헤더에 수식(ArrayFormula)이 있는 열은 건드리지 않음)
function writeRow_(sh, row, obj) {
  const lastCol = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, lastCol);
  const headers = head.getValues()[0].map(h => String(h).trim());
  const formulas = head.getFormulas()[0];
  headers.forEach((h, j) => {
    if (h && Object.prototype.hasOwnProperty.call(obj, h) && !formulas[j]) {
      sh.getRange(row, j + 1).setValue(obj[h]);
    }
  });
}
function append_(name, obj) {
  const sh = sheet_(name);
  writeRow_(sh, lastRowA_(sh) + 1, obj);
}
function findRow_(name, id) {
  const sh = sheet_(name);
  const n = lastRowA_(sh);
  const ids = sh.getRange(1, 1, n, 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return { sh: sh, row: i + 1 };
  throw new Error('항목을 찾을 수 없습니다.');
}

/* ================= 이미지 ================= */
function md5_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
// 앱시트 경로(예: 게시판_Images/xxx.jpg) → 드라이브 파일 ID
function resolvePath_(path) {
  if (!path) return '';
  const cache = CacheService.getScriptCache();
  const key = 'p:' + md5_(path);
  const c = cache.get(key);
  if (c) return c === '-' ? '' : c;
  const it = DriveApp.getFilesByName(String(path).split('/').pop());
  const id = it.hasNext() ? it.next().getId() : '';
  cache.put(key, id || '-', 21600);
  return id;
}
// 화면에 보여줄 이미지 (작게 줄인 썸네일 → base64)
function getImage(pw, ref, size) {
  check_(pw);
  const id = (ref && ref.id) || resolvePath_(ref && ref.path);
  if (!id) return '';
  const cache = CacheService.getScriptCache();
  const sz = Math.min(Math.max(Number(size) || 1000, 200), 2400);   // 화면용 1000, 크게 보기 2000
  const ck = 'img:' + id + ':' + sz;
  const c = cache.get(ck);
  if (c) return c;
  let data = '';
  try {
    const token = ScriptApp.getOAuthToken();
    const opt = { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
    const meta = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink', opt).getContentText());
    if (meta.thumbnailLink) {
      const blob = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s' + sz), opt).getBlob();
      data = 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
    }
  } catch (e) { /* 아래 원본으로 대체 */ }
  if (!data) {
    const file = DriveApp.getFileById(id);
    if (/^image\//.test(file.getMimeType()) && file.getSize() < 5e6) {
      const blob = file.getBlob();
      const bytes = blob.getBytes();
      data = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(bytes);
    }
  }
  if (data && data.length < 95000) cache.put(ck, data, 21600);
  return data;
}
// 사진 저장 폴더 찾기 (기존 사진이 있는 폴더 = 앱시트 폴더)
function folderFor_(propKey, sheetName, cols, defName) {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty(propKey);
  if (saved) { try { return DriveApp.getFolderById(saved); } catch (e) {} }
  const t = table_(sheetName);
  for (let i = t.rows.length - 1; i >= 0; i--) {
    for (const c of cols) {
      const v = t.rows[i][c];
      const id = driveId_(v) || (v && !/^https?:/.test(String(v)) ? resolvePath_(v) : '');
      if (!id) continue;
      try {
        const ps = DriveApp.getFileById(id).getParents();
        if (ps.hasNext()) { const f = ps.next(); props.setProperty(propKey, f.getId()); return f; }
      } catch (e) {}
    }
  }
  const it = DriveApp.getFoldersByName(defName);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(defName);
  props.setProperty(propKey, f.getId());
  return f;
}
function saveImage_(folder, dataUrl, name) {
  const m = String(dataUrl).match(/^data:(.+?);base64,(.*)$/);
  if (!m) throw new Error('이미지 형식 오류');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  return folder.createFile(blob).getId();
}

/* ================= 폰 알림 (ntfy) ================= */
// 가족마다 알림 주제(채널)를 하나씩 자동으로 만듦 (남이 추측할 수 없는 긴 이름)
function ntfyTopic_(email) {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('NTFY_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '').slice(0, 16); props.setProperty('NTFY_KEY', key); }
  return 'gd-' + key + '-' + md5_(String(email).toLowerCase()).slice(0, 8);
}
function appUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }
function userName_(email) {
  try {
    const u = table_(SHEET.USER).rows.find(r => String(r['이메일주소']).trim() === email);
    if (u && u['이름']) return String(u['이름']);
  } catch (e) {}
  return String(email || '').split('@')[0];
}
// ntfy 요청 만들기: https://ntfy.sh/주제?title=..&tags=..&click=..  (본문 = 메시지, 한글 OK)
function ntfyReq_(topic, title, message, tag) {
  const q = ['title=' + encodeURIComponent(title)];
  if (tag) q.push('tags=' + encodeURIComponent(tag));
  const click = appUrl_();
  if (click) q.push('click=' + encodeURIComponent(click));
  return {
    url: NTFY_SERVER.replace(/\/+$/, '') + '/' + encodeURIComponent(topic) + '?' + q.join('&'),
    method: 'post', contentType: 'text/plain; charset=utf-8',
    payload: String(message || ' '), muteHttpExceptions: true
  };
}
// 보낸 사람을 뺀 가족 모두에게 알림
function notify_(senderEmail, title, message, tag) {
  if (!NTFY_ENABLED || !NTFY_FROM_SERVER) return;
  try {
    const users = table_(SHEET.USER).rows.map(r => String(r['이메일주소'] || '').trim()).filter(Boolean);
    const reqs = users.filter(e => e !== senderEmail).map(e => ntfyReq_(ntfyTopic_(e),
      NTFY_SHOW_CONTENT ? title : '성장일기',
      NTFY_SHOW_CONTENT ? String(message || '').slice(0, 300) : '새 소식이 있어요', tag));
    if (reqs.length) UrlFetchApp.fetchAll(reqs);
  } catch (e) { console.log('알림 실패: ' + e); }   // 알림이 실패해도 저장은 정상 처리
}
// 화면이 알림을 직접 보낼 가족 주제 목록 (보낸 사람 제외)
function getNotifyTargets(pw, me) {
  check_(pw);
  if (!NTFY_ENABLED) return { enabled: false, topics: [] };
  const users = table_(SHEET.USER).rows.map(r => String(r['이메일주소'] || '').trim()).filter(Boolean);
  return { enabled: true, server: NTFY_SERVER, showContent: NTFY_SHOW_CONTENT,
           topics: users.filter(e => e !== me).map(e => ntfyTopic_(e)) };
}
// 화면의 🔔 버튼: 내 알림 주제 알려주기
function getNotifyInfo(pw, me) {
  check_(pw);
  return { enabled: NTFY_ENABLED, server: NTFY_SERVER, topic: ntfyTopic_(me) };
}
// 테스트 알림 (나에게만)
function testNotify(pw, me) {
  check_(pw);
  const req = ntfyReq_(ntfyTopic_(me), '🌱 성장일기 테스트', '알림이 잘 와요!', 'tada');
  const r = UrlFetchApp.fetch(req.url, req);
  if (r.getResponseCode() !== 200) throw new Error('알림 전송 실패 (' + r.getResponseCode() + '): ' + r.getContentText().slice(0, 200));
  return true;
}

/* ================= 초기 데이터 ================= */
function hello() { return { version: APP_VERSION, needPw: !!familyPw_() }; }

function init(pw) {
  check_(pw);
  const errors = [];
  const safe = (name, fn, def) => { try { return fn(); } catch (e) { errors.push(name + ': ' + e.message); return def; } };
  const users = safe('유저 시트', () => table_(SHEET.USER).rows.filter(r => r['이메일주소']).map(r => ({
    name: String(r['이름'] || ''),
    email: String(r['이메일주소']).trim(),
    avatar: imgRef_(r['이미지'], '')
  })), []);
  return {
    version: APP_VERSION,
    children: CHILDREN,
    users: users,
    posts: safe('게시판 시트', getPosts_, []),
    comments: safe('댓글 시트', getComments_, []),
    growth: safe('성장곡선 시트', getGrowth_, []),
    errors: errors
  };
}
function reload(pw) {
  check_(pw);
  const errors = [];
  const safe = (name, fn, def) => { try { return fn(); } catch (e) { errors.push(name + ': ' + e.message); return def; } };
  return { posts: safe('게시판 시트', getPosts_, []), comments: safe('댓글 시트', getComments_, []),
           growth: safe('성장곡선 시트', getGrowth_, []), errors: errors };
}

/**
 * 진단: Apps Script 편집기 위쪽에서 'diagnose' 함수를 고르고 ▶실행 → 아래 '실행 로그'를 캡처해서 보내주세요.
 */
function diagnose() {
  Logger.log('버전 v' + APP_VERSION);
  const run = (name, fn) => {
    const t = Date.now();
    try {
      const r = fn();
      const n = Array.isArray(r) ? r.length + '건' : (r && typeof r === 'object' ? Object.keys(r).join(',') : String(r));
      Logger.log('✅ ' + name + ' : ' + n + ' (' + (Date.now() - t) + 'ms)');
    } catch (e) { Logger.log('❌ ' + name + ' : ' + e.message + '\n' + (e.stack || '')); }
  };
  run('시트 열기', () => ss_().getSheets().map(s => s.getName()).join(' / '));
  run('유저', () => table_(SHEET.USER).rows);
  run('게시판', getPosts_);
  run('댓글', getComments_);
  run('성장곡선', getGrowth_);
  run('채팅', () => getChats(familyPw_()));
  run('앨범', () => getMedia(familyPw_()));
  run('초기 데이터 전송 크기', () => Math.round(JSON.stringify(init(familyPw_())).length / 1024) + 'KB');
}

/* ================= 게시판 ================= */
function getPosts_() {
  return table_(SHEET.POST).rows.filter(r => r['게시글ID']).map(r => ({
    id: String(r['게시글ID']),
    title: String(r['제목'] || ''),
    content: String(r['내용'] || ''),
    author: String(r['작성자'] || '').trim(),
    time: time_(r['작성일시']),
    likes: Number(r['좋아요']) || 0,
    target: String(r['확인1'] || ''),
    link: String(r['링크'] || ''),
    file: String(r['자료'] || ''),
    video: String(r['영상'] || ''),
    images: PATH_COLS.map((p, i) => imgRef_(r[p], r[URL_COLS[i]])).filter(Boolean)
  })).sort((a, b) => b.time.localeCompare(a.time));
}
function getComments_() {
  return table_(SHEET.COMMENT).rows.filter(r => r['댓글ID']).map(r => ({
    id: String(r['댓글ID']),
    postId: String(r['해당게시글ID'] || ''),
    text: String(r['댓글내용'] || ''),
    author: String(r['작성자'] || '').trim(),
    time: time_(r['작성일시'])
  })).sort((a, b) => a.time.localeCompare(b.time));
}
function addPost(pw, me, d) {
  check_(pw);
  if (!d || !String(d.title || '').trim()) throw new Error('제목을 입력해 주세요.');
  const id = uid_();
  const now = new Date();
  const when = parseLocal_(d.time) || now;          // 화면에서 고른 등록 날짜
  const nowStr = time_(when);
  const obj = {
    '게시글ID': id, '제목': d.title, '내용': d.content || '', '작성자': me,
    '작성일시': when, '좋아요': 0, '확인1': d.target || '모두'
  };
  const imgs = (d.images || []).slice(0, 5);
  if (imgs.length) {
    const folder = folderFor_('BOARD_FOLDER', SHEET.POST, ['Image URL', '이미지'], '게시판_Images');
    const hhmmss = Utilities.formatDate(now, TZ, 'HHmmss');
    imgs.forEach((data, i) => {
      const col = PATH_COLS[i];
      const fname = id + '.' + col + '.' + hhmmss + '.jpg';
      const fid = saveImage_(folder, data, fname);
      obj[col] = folder.getName() + '/' + fname;
      obj[URL_COLS[i]] = 'https://drive.google.com/uc?export=view&id=' + fid;
    });
  }
  CHILDREN.forEach(c => { obj[c.name + ' 생후'] = daysSince_(c.birth, nowStr); });
  lock_(() => append_(SHEET.POST, obj));
  const nP = Math.max(imgs.length, Number(d.totalPhotos) || 0), nV = Number(d.totalVideos) || 0;
  const extra = [nP ? '사진 ' + nP + '장' : '', nV ? '동영상 ' + nV + '개' : ''].filter(Boolean).join(', ');
  notify_(me, '📖 ' + userName_(me) + '님이 새 일기를 올렸어요', d.title + (extra ? ' (' + extra + ')' : ''), 'notebook');
  return { posts: getPosts_(), id: id };
}
function updatePost(pw, me, id, d) {
  check_(pw);
  return lock_(() => {
    const f = findRow_(SHEET.POST, id);
    const t = table_(SHEET.POST);
    const r = t.rows.find(x => x._row === f.row);
    if (String(r['작성자']).trim() !== me) throw new Error('작성자만 수정할 수 있습니다.');
    const upd = { '제목': d.title, '내용': d.content || '', '확인1': d.target || '모두' };
    const when = parseLocal_(d.time);
    if (when) {
      upd['작성일시'] = when;
      CHILDREN.forEach(c => { upd[c.name + ' 생후'] = daysSince_(c.birth, time_(when)); });
    }
    writeRow_(f.sh, f.row, upd);
    return getPosts_();
  });
}
function deletePost(pw, me, id) {
  check_(pw);
  return lock_(() => {
    const f = findRow_(SHEET.POST, id);
    const author = String(f.sh.getRange(f.row, table_(SHEET.POST).headers.indexOf('작성자') + 1).getValue()).trim();
    if (author !== me) throw new Error('작성자만 삭제할 수 있습니다.');
    const row = table_(SHEET.POST).rows.find(x => x._row === f.row);
    const files = row ? postFileIds_(row) : [];
    f.sh.deleteRow(f.row);
    deleteRowsWhere_(SHEET.MEDIA, r => String(r['Link_ID'] || '').trim() === String(id)).forEach(r => mediaRowIds_(r).forEach(x => files.push(x)));
    deleteRowsWhere_(SHEET.COMMENT, r => String(r['해당게시글ID'] || '').trim() === String(id));
    trashUnused_(files);
    return getPosts_();
  });
}
function likePost(pw, id) {
  check_(pw);
  return lock_(() => {
    const f = findRow_(SHEET.POST, id);
    const col = table_(SHEET.POST).headers.indexOf('좋아요') + 1;
    const cell = f.sh.getRange(f.row, col);
    const n = (Number(cell.getValue()) || 0) + 1;
    cell.setValue(n);
    return n;
  });
}
function addComment(pw, me, postId, text) {
  check_(pw);
  if (!String(text || '').trim()) throw new Error('댓글을 입력해 주세요.');
  lock_(() => append_(SHEET.COMMENT, {
    '댓글ID': uid_(), '해당게시글ID': postId, '댓글내용': text, '작성자': me, '작성일시': new Date()
  }));
  try {
    const post = getPosts_().find(x => x.id === String(postId));
    notify_(me, '💬 ' + userName_(me) + '님의 댓글' + (post ? ' · ' + post.title : ''), text, 'speech_balloon');
  } catch (e) {}
  return getComments_();
}
function deleteComment(pw, me, id) {
  check_(pw);
  return lock_(() => {
    const f = findRow_(SHEET.COMMENT, id);
    const col = table_(SHEET.COMMENT).headers.indexOf('작성자') + 1;
    if (String(f.sh.getRange(f.row, col).getValue()).trim() !== me) throw new Error('작성자만 삭제할 수 있습니다.');
    f.sh.deleteRow(f.row);
    return getComments_();
  });
}

/* ================= 앨범 ('Multi file upload' 시트) ================= */
// Link_ID = 게시판 게시글ID. 한 칸에 여러 링크가 있어도 모두 분리.
function getMedia(pw) {
  check_(pw);
  const seen = {}, list = [];
  table_(SHEET.MEDIA).rows.forEach(r => {
    const postId = String(r['Link_ID'] || '').trim();
    const time = time_(r['타임스탬프']);
    const ids = [];
    const f = String(r['fileID'] || '').trim();
    if (/^[\w-]{20,}$/.test(f)) ids.push(f);
    String(r['Photo OR Video'] || '').split(/[,\s]+/).forEach(s => { const id = driveId_(s); if (id) ids.push(id); });
    const e = driveId_(r['Photo_adress_copy']); if (e) ids.push(e);
    ids.forEach(id => { if (!seen[id]) { seen[id] = 1; list.push({ id: id, postId: postId, time: time }); } });
  });
  const types = mimeTypes_(list.map(m => m.id));
  return list.filter(m => types[m.id] && types[m.id] !== 'x')
    .map(m => ({ id: m.id, postId: m.postId, time: m.time, video: /^video\//.test(types[m.id]) }))
    .sort((a, b) => b.time.localeCompare(a.time));
}
// 일기에 사진·동영상 추가 ('Multi file upload' 시트에 앱시트와 같은 형식으로 기록)
function mediaFolder_() {
  return folderFor_('MEDIA_FOLDER', SHEET.MEDIA, ['Photo_adress_copy', 'Photo OR Video'], 'Multi file upload_Images');
}
function appendMedia_(postId, ids) {
  const now = new Date();
  lock_(() => ids.forEach(id => append_(SHEET.MEDIA, {
    '타임스탬프': now, 'Link_ID': postId,
    'Photo OR Video': 'https://drive.google.com/thumbnail?sz=s0&id=' + id,
    'ID': uid_().toUpperCase(),
    'Photo_adress_copy': 'https://drive.google.com/uc?export=view&id=' + id,
    'Description': 'https://drive.google.com/file/d/' + id + '/view?usp=drivesdk',
    'fileID': id
  })));
}
// 사진 (화면에서 줄여서 5장씩 보냄)
function addMedia(pw, me, postId, images) {
  check_(pw);
  if (!postId) throw new Error('일기를 찾을 수 없어요.');
  const folder = mediaFolder_();
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
  const ids = (images || []).slice(0, 10).map((data, i) => saveImage_(folder, data, postId + '_' + stamp + '_' + (i + 1) + '.jpg'));
  appendMedia_(postId, ids);
  return ids.length;
}
// 동영상 (분할 업로드가 끝난 파일 등록)
function registerMedia(pw, me, postId, fileIds) {
  check_(pw);
  if (!postId) throw new Error('일기를 찾을 수 없어요.');
  appendMedia_(postId, (fileIds || []).filter(Boolean));
  return true;
}
// 기존 일기에 추가했을 때 알림
function notifyMedia(pw, me, postId, nPhotos, nVideos) {
  check_(pw);
  const post = getPosts_().find(x => x.id === String(postId));
  const what = [nPhotos ? '사진 ' + nPhotos + '장' : '', nVideos ? '동영상 ' + nVideos + '개' : ''].filter(Boolean).join(', ');
  if (what) notify_(me, '🖼️ ' + userName_(me) + '님이 ' + what + '을 추가했어요', post ? post.title : '', 'frame_with_picture');
  return true;
}
// 파일 종류(사진/영상) 확인 — 한 번 확인한 건 6시간 캐시
function mimeTypes_(ids) {
  const cache = CacheService.getScriptCache();
  const out = {};
  for (let i = 0; i < ids.length; i += 100) {
    const got = cache.getAll(ids.slice(i, i + 100).map(id => 'mt:' + id));
    Object.keys(got).forEach(k => { out[k.slice(3)] = got[k]; });
  }
  const need = ids.filter(id => !out[id]);
  const token = ScriptApp.getOAuthToken();
  for (let i = 0; i < need.length; i += 50) {
    const part = need.slice(i, i + 50);
    const res = UrlFetchApp.fetchAll(part.map(id => ({
      url: 'https://www.googleapis.com/drive/v3/files/' + id + '?fields=mimeType,trashed',
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    })));
    const put = {};
    res.forEach((r, j) => {
      const code = r.getResponseCode();
      if (code === 200) {
        const m = JSON.parse(r.getContentText());
        const v = m.trashed ? 'x' : (m.mimeType || 'image/unknown');
        out[part[j]] = v; put['mt:' + part[j]] = v;
      } else if (code === 404) {
        out[part[j]] = 'x'; put['mt:' + part[j]] = 'x';      // 삭제된 파일
      } else {
        out[part[j]] = 'image/unknown';                       // 일시 오류: 일단 사진으로 표시, 캐시 안 함
      }
    });
    cache.putAll(put, 21600);
  }
  return out;
}

/* ================= 저장공간 관리 (드라이브 파일 정리) ================= */
// 시트 어디에서든 쓰이고 있는 파일 (ID / 파일 이름)
function referenced_() {
  const ids = {}, names = {};
  const addId = v => { const id = driveId_(v) || (/^[\w-]{25,}$/.test(String(v || '').trim()) ? String(v).trim() : ''); if (id) ids[id] = 1; };
  const addPath = v => { const s = String(v || ''); if (s && !/^https?:/.test(s)) names[s.split('/').pop()] = 1; };
  const scan = (sheet, idCols, pathCols) => {
    let t; try { t = table_(sheet); } catch (e) { return; }
    t.rows.forEach(r => {
      idCols.forEach(c => String(r[c] || '').split(/[,\s]+/).forEach(addId));
      pathCols.forEach(c => addPath(r[c]));
    });
  };
  scan(SHEET.POST, URL_COLS.concat(PATH_COLS), PATH_COLS);
  scan(SHEET.MEDIA, ['fileID', 'Photo OR Video', 'Photo_adress_copy', 'Description'], []);
  scan(SHEET.CHAT, ['Photo', '첨부ID'], ['Photo']);
  scan(SHEET.USER, ['이미지'], ['이미지']);
  return { ids: ids, names: names };
}
// 다른 곳에서 안 쓰는 파일만 휴지통으로 (30일 안에 복구 가능)
function trashUnused_(fileIds) {
  const ref = referenced_();
  let n = 0;
  (fileIds || []).filter(Boolean).forEach(id => {
    if (ref.ids[id]) return;
    try {
      const f = DriveApp.getFileById(id);
      if (ref.names[f.getName()]) return;
      f.setTrashed(true); n++;
    } catch (e) { console.log('휴지통 실패 ' + id + ': ' + e); }
  });
  return n;
}
function postFileIds_(r) {
  return PATH_COLS.map((p, i) => driveId_(r[URL_COLS[i]]) || driveId_(r[p]) || (r[p] ? resolvePath_(r[p]) : '')).filter(Boolean);
}
function mediaRowIds_(r) {
  const ids = [];
  const f = String(r['fileID'] || '').trim(); if (/^[\w-]{20,}$/.test(f)) ids.push(f);
  String(r['Photo OR Video'] || '').split(/[,\s]+/).forEach(s => { const id = driveId_(s); if (id) ids.push(id); });
  const e = driveId_(r['Photo_adress_copy']); if (e) ids.push(e);
  return ids.filter((x, i, a) => a.indexOf(x) === i);
}
// 시트에서 조건에 맞는 행 삭제 (아래쪽부터)
function deleteRowsWhere_(sheetName, test) {
  const t = table_(sheetName);
  const hit = t.rows.filter(test);
  hit.map(r => r._row).sort((a, b) => b - a).forEach(row => t.sh.deleteRow(row));
  return hit;
}
// 사진·영상 1개 삭제 (앨범·일기 상세에서)
function deleteMedia(pw, me, fileId) {
  check_(pw);
  const removed = lock_(() => deleteRowsWhere_(SHEET.MEDIA, r => mediaRowIds_(r).indexOf(fileId) >= 0));
  if (!removed.length) throw new Error('이 사진·영상은 일기 본문 사진이라 여기서 지울 수 없어요.');
  trashUnused_([fileId]);
  return true;
}
function owner_() { try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } }
function storageFolders_() {
  const list = [];
  // 안전장치: 내 드라이브 최상위 폴더나, 앱 전용 이름이 아닌 폴더는 절대 검사하지 않음
  const rootId = DriveApp.getRootFolder().getId();
  const safeName = n => /_Images|File responses|파일 응답/i.test(n);
  const add = (label, fn) => {
    try {
      const f = fn();
      if (!f || f.getId() === rootId || !safeName(f.getName())) return;
      if (!list.some(x => x.f.getId() === f.getId())) list.push({ label: label, f: f });
    } catch (e) {}
  };
  add('일기 사진', () => folderFor_('BOARD_FOLDER', SHEET.POST, ['Image URL', '이미지'], '게시판_Images'));
  add('사진·영상', () => mediaFolder_());
  add('채팅', () => folderFor_('CHAT_FOLDER', SHEET.CHAT, ['Photo'], '챗_Images'));
  return list;
}
function findOrphans_() {
  const ref = referenced_();
  const cutoff = Date.now() - 30 * 60 * 1000;     // 30분 안에 올린 파일은 제외 (업로드 중일 수 있음)
  const out = [];
  storageFolders_().forEach(x => {
    const it = x.f.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (ref.ids[f.getId()] || ref.names[f.getName()]) continue;
      if (f.getDateCreated().getTime() > cutoff) continue;
      out.push({ id: f.getId(), name: f.getName(), size: f.getSize(), folder: x.label, date: date_(f.getDateCreated()) });
    }
  });
  return out;
}
// 검사: 시트와 연결이 끊긴 파일 목록 (아빠 계정만)
function scanOrphans(pw, me) {
  check_(pw);
  const owner = owner_();
  if (owner && me !== owner) throw new Error('저장공간 정리는 ' + userName_(owner) + ' 계정에서만 할 수 있어요.');
  const list = findOrphans_();
  let used = 0, files = 0;
  storageFolders_().forEach(x => { const it = x.f.getFiles(); while (it.hasNext()) { used += it.next().getSize(); files++; } });
  return {
    count: list.length, size: list.reduce((a, b) => a + b.size, 0),
    used: used, files: files, folders: storageFolders_().map(x => x.label + ' (' + x.f.getName() + ')'),
    list: list.sort((a, b) => b.size - a.size).slice(0, 40)
  };
}
// 정리: 다시 검사해서 휴지통으로
function cleanOrphans(pw, me) {
  check_(pw);
  const owner = owner_();
  if (owner && me !== owner) throw new Error('저장공간 정리는 ' + userName_(owner) + ' 계정에서만 할 수 있어요.');
  let n = 0, size = 0;
  findOrphans_().forEach(o => { try { DriveApp.getFileById(o.id).setTrashed(true); n++; size += o.size; } catch (e) {} });
  return { count: n, size: size };
}

/* ================= 성장곡선 ================= */
function getGrowth_() {
  return table_(SHEET.GROWTH).rows.filter(r => r['성장곡선_ID'] && r['이름']).map(r => ({
    id: String(r['성장곡선_ID']),
    name: String(r['이름']),
    date: date_(r['작성일']),
    height: r['키'] === '' ? null : Number(r['키']),
    weight: r['체중'] === '' ? null : Number(r['체중']),
    note: String(r['비고'] || '')
  })).sort((a, b) => a.date.localeCompare(b.date));
}
function addGrowth(pw, me, d) {
  check_(pw);
  if (!d || !d.name || !d.date) throw new Error('이름과 날짜를 입력해 주세요.');
  const obj = {
    '성장곡선_ID': uid_(), '이름': d.name,
    '작성일': new Date(d.date + 'T00:00:00+09:00'),
    '키': d.height === '' || d.height == null ? '' : Number(d.height),
    '체중': d.weight === '' || d.weight == null ? '' : Number(d.weight),
    '비고': d.note || ''
  };
  CHILDREN.forEach(c => { obj[c.name + ' 생후'] = daysSince_(c.birth, d.date); });
  lock_(() => append_(SHEET.GROWTH, obj));
  return getGrowth_();
}
function deleteGrowth(pw, me, id) {
  check_(pw);
  return lock_(() => {
    const f = findRow_(SHEET.GROWTH, id);
    f.sh.deleteRow(f.row);
    return getGrowth_();
  });
}

/* ================= 가족 채팅 ('챗' 시트) ================= */
// 동영상·파일 정보는 '챗' 시트 오른쪽 끝에 열을 자동 추가해서 저장
const CHAT_ATTACH_COLS = ['첨부ID', '첨부명', '첨부종류', '첨부크기'];
const MAX_UPLOAD = 100 * 1024 * 1024;   // 동영상·파일 최대 100MB

function ensureCols_(sh, names) {
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = names.filter(n => headers.indexOf(n) < 0);
  if (missing.length) sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
}
function autoLabel_(kind, name) { return kind === 'video' ? '🎬 동영상' : kind === 'gif' ? '🎞️ GIF' : '📎 ' + name; }

function chatList_() {
  const list = table_(SHEET.CHAT).rows
    .filter(r => r['ID'] && (r['Message'] || r['Photo'] || r['첨부ID']))
    .map(r => {
      const m = {
        id: String(r['ID']),
        sender: String(r['작성자'] || r['이름'] || '').trim(),
        text: String(r['Message'] || ''),
        time: time_(r['Timestamp']),
        photo: imgRef_(r['Photo'], r['Photo'])
      };
      if (r['첨부ID']) {
        m.attach = {
          id: String(r['첨부ID']), name: String(r['첨부명'] || '파일'),
          kind: String(r['첨부종류'] || 'file'), size: Number(r['첨부크기']) || 0,
          url: String(r['첨부종류']) === 'gif' ? String(r['첨부ID']) : ''
        };
        if (m.text === autoLabel_(m.attach.kind, m.attach.name)) m.text = '';
      }
      return m;
    })
    .sort((a, b) => a.time.localeCompare(b.time));
  return list.slice(-300);
}

// 읽음 기록: 사람마다 '마지막으로 읽은 시각' (스크립트 속성에 저장, 시트는 건드리지 않음)
function chatReads_() {
  const all = PropertiesService.getScriptProperties().getProperties(), out = {};
  Object.keys(all).forEach(k => { if (k.indexOf('read:') === 0) out[k.slice(5)] = all[k]; });
  return out;
}
function chatMembers_() {
  return table_(SHEET.USER).rows.map(r => String(r['이메일주소'] || '').trim()).filter(e => e && CHAT_EXCLUDE.indexOf(e) < 0);
}
function markRead_(me, t) {
  if (!me || !t) return;
  const p = PropertiesService.getScriptProperties(), k = 'read:' + me;
  if (t > (p.getProperty(k) || '')) p.setProperty(k, t);
}
// 채팅 목록 + 읽음 정보. mark=true 면 '나'는 최신 메시지까지 읽은 것으로 기록
function chatPayload_(me, mark) {
  const list = chatList_();
  if (mark && me && list.length) markRead_(me, list[list.length - 1].time);
  return { chats: list, reads: chatReads_(), members: chatMembers_() };
}
function getChats(pw, me, mark) {
  check_(pw);
  return chatPayload_(me, !!mark);
}

// 텍스트 + 사진 여러 장 (사진은 화면에서 줄여서 보냄)
function sendChat(pw, me, myName, text, images) {
  check_(pw);
  if (images && !Array.isArray(images)) images = [images];
  images = (images || []).slice(0, 30);
  if (!String(text || '').trim() && !images.length) throw new Error('메시지를 입력해 주세요.');
  const now = new Date();
  const rows = [];
  if (images.length) {
    const folder = folderFor_('CHAT_FOLDER', SHEET.CHAT, ['Photo'], '챗_Images');
    const hhmmss = Utilities.formatDate(now, TZ, 'HHmmss');
    images.forEach(data => {
      const id = uid_();
      const fname = id + '.Photo.' + hhmmss + '.jpg';
      saveImage_(folder, data, fname);
      rows.push({ 'ID': id, 'Link_ID': myName || '', '이름': me, 'Message': '', 'Timestamp': now, '작성자': me,
                  'Photo': folder.getName() + '/' + fname });
    });
  }
  if (String(text || '').trim()) {
    rows.push({ 'ID': uid_(), 'Link_ID': myName || '', '이름': me, 'Message': text, 'Timestamp': now, '작성자': me });
  }
  lock_(() => rows.forEach(o => append_(SHEET.CHAT, o)));
  const msg = [images.length ? '📷 사진 ' + images.length + '장' : '', String(text || '').trim()].filter(Boolean).join('\n');
  notify_(me, '💬 ' + (myName || userName_(me)), msg, 'speech_balloon');
  return chatPayload_(me, true);
}

// 큰 파일 업로드 ① 시작: 드라이브 업로드 주소 받기
// origin 을 주면 '폰 → 드라이브 직접 업로드'용 주소를 만듦 (빠름). 안 주면 서버 경유용.
function startUpload(pw, name, mime, size, target, origin) {
  check_(pw);
  if (!size || size > MAX_UPLOAD) throw new Error('100MB 이하 파일만 올릴 수 있어요.');
  const folder = target === 'media' ? mediaFolder_() : folderFor_('CHAT_FOLDER', SHEET.CHAT, ['Photo'], '챗_Images');
  const safe = String(name || 'file').replace(/[\\/:*?"<>|]/g, '_');
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
    method: 'post', contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify({ name: uid_() + '_' + safe, parents: [folder.getId()], mimeType: mime || 'application/octet-stream' }),
    headers: Object.assign({ Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
               'X-Upload-Content-Type': mime || 'application/octet-stream', 'X-Upload-Content-Length': String(size) },
               (origin && /^https:\/\/[\w.-]+(\.googleusercontent\.com|\.github\.io)$/.test(origin)) ? { 'Origin': origin } : {}),
    muteHttpExceptions: true
  });
  const h = res.getHeaders();
  const url = h['Location'] || h['location'];
  if (!url) throw new Error('업로드 시작 실패 (' + res.getResponseCode() + ')');
  return url;
}
// 큰 파일 업로드 ② 조각 보내기 (4MB씩)
function uploadChunk(pw, url, b64, start, total) {
  check_(pw);
  const bytes = Utilities.base64Decode(b64);
  const end = start + bytes.length - 1;
  const res = UrlFetchApp.fetch(url, {
    method: 'put', payload: bytes, contentType: 'application/octet-stream',
    headers: { 'Content-Range': 'bytes ' + start + '-' + end + '/' + total },
    muteHttpExceptions: true, followRedirects: false
  });
  const code = res.getResponseCode();
  if (code === 308) return { done: false };
  if (code === 200 || code === 201) return { done: true, id: JSON.parse(res.getContentText()).id };
  throw new Error('업로드 실패 (' + code + ')');
}
// 큰 파일 업로드 ③ 채팅에 등록
function sendChatFile(pw, me, myName, f) {
  check_(pw);
  const kind = f.kind === 'video' ? 'video' : 'file';
  lock_(() => {
    const sh = sheet_(SHEET.CHAT);
    ensureCols_(sh, CHAT_ATTACH_COLS);
    append_(SHEET.CHAT, {
      'ID': uid_(), 'Link_ID': myName || '', '이름': me, 'Message': autoLabel_(kind, f.name),
      'Timestamp': new Date(), '작성자': me,
      '첨부ID': f.id, '첨부명': f.name, '첨부종류': kind, '첨부크기': f.size || 0
    });
  });
  notify_(me, '💬 ' + (myName || userName_(me)), autoLabel_(kind, f.name), 'speech_balloon');
  return chatPayload_(me, true);
}
// GIF 보내기 (GIPHY 주소를 그대로 저장)
function sendChatGif(pw, me, myName, url, title) {
  check_(pw);
  if (!/^https:\/\/[\w.-]*giphy\.com\//.test(String(url || ''))) throw new Error('GIF 주소가 올바르지 않아요.');
  lock_(() => {
    ensureCols_(sheet_(SHEET.CHAT), CHAT_ATTACH_COLS);
    append_(SHEET.CHAT, {
      'ID': uid_(), 'Link_ID': myName || '', '이름': me, 'Message': autoLabel_('gif', ''),
      'Timestamp': new Date(), '작성자': me,
      '첨부ID': url, '첨부명': String(title || 'GIF').slice(0, 100), '첨부종류': 'gif', '첨부크기': 0
    });
  });
  return chatPayload_(me, true);
}
function getGifKey(pw) { check_(pw); return giphyKey_(); }
// 파일 받기 (25MB 이하는 앱에서 바로 저장)
function getFileData(pw, id) {
  check_(pw);
  const file = DriveApp.getFileById(id);
  if (file.getSize() > 25 * 1024 * 1024) return { big: true };
  const blob = file.getBlob();
  return { name: file.getName().replace(/^[0-9a-f]{8}_/, ''), mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}
// 내 메시지 삭제
function deleteChat(pw, me, id) {
  check_(pw);
  return lock_(() => {
    const f = findRow_(SHEET.CHAT, id);
    const col = table_(SHEET.CHAT).headers.indexOf('작성자') + 1;
    if (String(f.sh.getRange(f.row, col).getValue()).trim() !== me) throw new Error('내 메시지만 삭제할 수 있어요.');
    const row = table_(SHEET.CHAT).rows.find(x => x._row === f.row) || {};
    const files = [driveId_(row['Photo']) || (row['Photo'] ? resolvePath_(row['Photo']) : ''), String(row['첨부ID'] || '').trim()];
    f.sh.deleteRow(f.row);
    trashUnused_(files);
    return chatPayload_(me, true);
  });
}
