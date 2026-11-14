/**
 * 2026 코딩앤메이커 with AI캠프 - 가족 앱 부스 투표
 * 구글 스프레드시트에 붙여 쓰는 Apps Script 웹앱 (시트가 DB 역할)
 */
const TAB = { fam: '가족', vote: '투표', judge: '심사', photo: '사진', conf: '설정' };
const HEAD = {
  '가족': ['ID', '뒷자리', '가운데', '이름', '모둠이름', '앱이름', '사진버전', '설정시각'],
  '투표': ['시각', '투표한가족', '받은가족', '응원의말'],
  '심사': ['ID', '심사1_문제해결', '심사1_작동', '심사1_학생설명', '심사2_문제해결', '심사2_작동', '심사2_학생설명'],
  '사진': ['ID', '버전', '조각1', '조각2', '조각3'],
  '설정': ['항목', '값']
};
const STATES = ['준비', '진행', '마감'];
const CHUNK = 45000;
const PHOTO_MAX = CHUNK * 3;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('가족 앱 부스 투표')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 시트 메뉴 ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('캠프')
    .addItem('1. 초기 설정 (탭 만들기)', 'setup')
    .addItem('2. 등록부에서 가족 불러오기', 'importRoster')
    .addToUi();
}

function setup() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEAD).forEach(function (name) {
    let s = book.getSheetByName(name);
    if (!s) {
      s = book.insertSheet(name);
      s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).setNumberFormat('@');
      s.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]);
      s.setFrozenRows(1);
    }
  });
  const c = conf_();
  const s = book.getSheetByName(TAB.conf);
  if (!c['관리자비밀번호']) s.appendRow(['관리자비밀번호', String(Math.floor(100000 + Math.random() * 900000))]);
  if (!c['투표상태']) s.appendRow(['투표상태', '준비']);
  try {
    SpreadsheetApp.getUi().alert('초기 설정 완료.\n관리자 비밀번호는 [설정] 탭에서 확인하고 바꿀 수 있습니다.');
  } catch (e) {}
}

/** 등록부 탭(우리가 만든 탭이 아닌 것)에서 보호자 전화번호를 찾아 가족 계정을 만든다. 다시 실행해도 새 가족만 추가된다. */
function importRoster() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const ours = Object.keys(HEAD);
  let src = book.getActiveSheet();
  if (ours.indexOf(src.getName()) >= 0) {
    src = book.getSheets().filter(function (s) { return ours.indexOf(s.getName()) < 0; })[0];
  }
  if (!src) throw new Error('등록부 탭을 찾지 못했습니다.');
  setupQuiet_();
  const v = src.getDataRange().getDisplayValues();
  let hr = -1;
  for (let i = 0; i < Math.min(v.length, 10) && hr < 0; i++) {
    if (v[i].some(function (c) { return /이름|성명/.test(c); })) hr = i;
  }
  const head = hr >= 0 ? v[hr] : [];
  let nameCol = -1, phoneCol = -1;
  head.forEach(function (h, i) {
    if (/이름|성명/.test(h) && (nameCol < 0 || /학생/.test(h))) nameCol = i;
    if (/전화|연락처|휴대|핸드폰/.test(h) && (phoneCol < 0 || /보호자|학부모|부모/.test(h))) phoneCol = i;
  });
  const have = {};
  fams_().forEach(function (f) { have[f.mid4 + f.last4] = true; });
  const groups = {}, order = [];
  for (let r = hr + 1; r < v.length; r++) {
    const cells = phoneCol >= 0 ? [v[r][phoneCol]].concat(v[r]) : v[r];
    let d = null;
    for (let k = 0; k < cells.length && !d; k++) d = phoneDigits_(cells[k]);
    if (!d) continue;
    const key = d.slice(-8);
    if (!groups[key]) { groups[key] = []; order.push(key); }
    const nm = nameCol >= 0 ? String(v[r][nameCol]).trim() : '';
    if (nm && groups[key].indexOf(nm) < 0) groups[key].push(nm);
  }
  const fs = book.getSheetByName(TAB.fam);
  let n = fs.getLastRow() - 1, added = 0;
  order.forEach(function (key) {
    if (have[key]) return;
    n++; added++;
    fs.appendRow(['F' + ('0' + n).slice(-2), key.slice(4), key.slice(0, 4), groups[key].join(', '), '', '', '', '']);
  });
  SpreadsheetApp.getUi().alert('[' + src.getName() + '] 탭에서 ' + added + '가족을 추가했습니다. (전체 ' + n + '가족)\n[가족] 탭에서 확인하세요.');
}

function setupQuiet_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!Object.keys(HEAD).every(function (n) { return book.getSheetByName(n); })) setup();
}

function phoneDigits_(cell) {
  const m = String(cell).match(/0?1[016789][-\s.)]*\d{4}[-\s.]*\d{4}/);
  return m ? m[0].replace(/\D/g, '') : null;
}

/* ---------- 공통 ---------- */

function sh_(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('[' + name + '] 탭이 없습니다. 시트 메뉴 [캠프 > 초기 설정]을 먼저 실행하세요.');
  return s;
}
function rows_(name) {
  const v = sh_(name).getDataRange().getValues();
  v.shift();
  return v;
}
function str_(v) { return String(v == null ? '' : v).replace(/^'/, '').trim(); }
function pad4_(v) { return ('0000' + String(v).replace(/\D/g, '')).slice(-4); }
function safe_(s, max) {
  s = String(s == null ? '' : s).replace(/[\u0000-\u0008\u000b-\u001f]/g, '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function conf_() {
  const o = {};
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.conf);
  if (!s) return o;
  s.getDataRange().getValues().forEach(function (r) { o[str_(r[0])] = str_(r[1]); });
  return o;
}
function state_() {
  const s = conf_()['투표상태'];
  return STATES.indexOf(s) >= 0 ? s : '준비';
}
function fams_() {
  return rows_(TAB.fam).map(function (r, i) {
    return {
      row: i + 2, id: str_(r[0]), last4: pad4_(r[1]), mid4: pad4_(r[2]), name: str_(r[3]),
      team: str_(r[4]), app: str_(r[5]), photoVer: str_(r[6])
    };
  }).filter(function (f) { return f.id; });
}
function votes_() {
  return rows_(TAB.vote).map(function (r, i) {
    return { row: i + 2, from: str_(r[1]), to: str_(r[2]), msg: str_(r[3]) };
  }).filter(function (v) { return v.from && v.to; });
}
function auth_(a) {
  a = a || {};
  if (a.admin != null) {
    const pw = conf_()['관리자비밀번호'];
    if (!pw || String(a.admin) !== pw) throw new Error('관리자 비밀번호가 맞지 않습니다.');
    return { admin: true };
  }
  const l = String(a.last4 || ''), m = String(a.mid4 || '');
  if (!/^\d{4}$/.test(l) || !/^\d{4}$/.test(m)) throw new Error('숫자 4자리씩 입력해 주세요.');
  const f = fams_().filter(function (x) { return x.last4 === l && x.mid4 === m; })[0];
  if (!f) throw new Error('등록된 번호가 아닙니다. 보호자 전화번호를 다시 확인해 주세요.');
  return { fam: f };
}
function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function pub_(f) { return { id: f.id, team: f.team, app: f.app, photoVer: f.photoVer }; }

/* ---------- 참가 가족 API ---------- */

function api_state(a) {
  const me = auth_(a).fam;
  if (!me) throw new Error('가족 계정으로 로그인해 주세요.');
  return {
    me: { id: me.id, name: me.name, team: me.team, app: me.app, photoVer: me.photoVer },
    state: state_(),
    booths: fams_().filter(function (f) { return f.id !== me.id; }).map(pub_),
    myVotes: votes_().filter(function (v) { return v.from === me.id; })
      .map(function (v) { return { to: v.to, msg: v.msg }; })
  };
}

function api_saveProfile(a, team, app, photo) {
  return locked_(function () {
    const me = auth_(a).fam;
    if (!me) throw new Error('가족 계정으로 로그인해 주세요.');
    if (state_() === '마감') throw new Error('투표가 마감되어 더 이상 바꿀 수 없습니다.');
    team = safe_(team, 20); app = safe_(app, 30);
    if (!team || !app) throw new Error('가족 모둠 이름과 앱 이름을 모두 적어 주세요.');
    let ver = me.photoVer;
    if (photo) {
      photo = String(photo);
      if (photo.indexOf('data:image/jpeg;base64,') !== 0 || photo.length > PHOTO_MAX) throw new Error('사진을 올리지 못했습니다. 다른 사진으로 다시 시도해 주세요.');
      ver = String(Date.now());
      const row = [me.id, ver];
      for (let i = 0; i < 3; i++) row.push('|' + photo.slice(i * CHUNK, (i + 1) * CHUNK));
      const ps = sh_(TAB.photo);
      const idx = rows_(TAB.photo).map(function (r) { return str_(r[0]); }).indexOf(me.id);
      if (idx >= 0) ps.getRange(idx + 2, 1, 1, 5).setValues([row]); else ps.appendRow(row);
    }
    if (!ver) throw new Error('앱 대표 사진을 올려 주세요.');
    sh_(TAB.fam).getRange(me.row, 5, 1, 4).setValues([[team, app, ver, new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })]]);
    return api_state(a);
  });
}

function api_vote(a, picks) {
  return locked_(function () {
    const me = auth_(a).fam;
    if (!me) throw new Error('가족 계정으로 로그인해 주세요.');
    const st = state_();
    if (st !== '진행') throw new Error(st === '마감' ? '투표가 마감되었습니다.' : '아직 투표 시간이 아닙니다.');
    if (!me.app) throw new Error('먼저 우리 가족 설정을 마쳐 주세요.');
    if (!Array.isArray(picks) || picks.length !== 3) throw new Error('부스 3곳을 골라 주세요.');
    const all = {};
    fams_().forEach(function (f) { all[f.id] = f; });
    const seen = {};
    picks = picks.map(function (p) {
      const to = String(p && p.to || '');
      if (to === me.id) throw new Error('우리 가족 부스에는 투표할 수 없습니다.');
      if (!all[to] || !all[to].app) throw new Error('투표할 수 없는 부스가 있습니다.');
      if (seen[to]) throw new Error('한 부스에는 한 표만 줄 수 있습니다.');
      seen[to] = true;
      return { to: to, msg: safe_(p.msg, 200) };
    });
    const s = sh_(TAB.vote);
    votes_().filter(function (v) { return v.from === me.id; }).map(function (v) { return v.row; })
      .sort(function (x, y) { return y - x; }).forEach(function (r) { s.deleteRow(r); });
    const now = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
    picks.forEach(function (p) { s.appendRow([now, me.id, p.to, p.msg]); });
    return api_state(a);
  });
}

function api_photos(a, ids) {
  auth_(a);
  const want = {};
  (ids || []).slice(0, 8).forEach(function (i) { want[String(i)] = true; });
  const out = {};
  rows_(TAB.photo).forEach(function (r) {
    const id = str_(r[0]);
    if (!want[id]) return;
    out[id] = { ver: str_(r[1]), data: [r[2], r[3], r[4]].map(function (c) { return String(c || '').replace(/^\|/, ''); }).join('') };
  });
  return out;
}

/* ---------- 관리자 API ---------- */

function admin_(a) {
  if (!auth_(a).admin) throw new Error('관리자만 사용할 수 있습니다.');
}

function api_admin(a) {
  admin_(a);
  const fams = fams_(), votes = votes_(), byId = {}, judge = {};
  fams.forEach(function (f) { byId[f.id] = f; });
  rows_(TAB.judge).forEach(function (r) { judge[str_(r[0])] = r.slice(1, 7).map(str_); });
  return {
    state: state_(),
    fams: fams.map(function (f) {
      const got = votes.filter(function (v) { return v.to === f.id; });
      return {
        id: f.id, name: f.name, team: f.team, app: f.app, photoVer: f.photoVer,
        voted: votes.some(function (v) { return v.from === f.id; }),
        votes: got.length,
        msgs: got.filter(function (v) { return v.msg; }).map(function (v) {
          return { from: (byId[v.from] && byId[v.from].team) || v.from, msg: v.msg };
        }),
        judge: judge[f.id] || ['', '', '', '', '', '']
      };
    })
  };
}

function api_adminState(a, state) {
  return locked_(function () {
    admin_(a);
    if (STATES.indexOf(state) < 0) throw new Error('알 수 없는 상태입니다.');
    const s = sh_(TAB.conf);
    const idx = s.getDataRange().getValues().map(function (r) { return str_(r[0]); }).indexOf('투표상태');
    if (idx >= 0) s.getRange(idx + 1, 2, 1, 1).setValues([[state]]); else s.appendRow(['투표상태', state]);
    return api_admin(a);
  });
}

function api_adminScore(a, id, col, val) {
  return locked_(function () {
    admin_(a);
    col = Number(col);
    if (!(col >= 0 && col < 6) || ['', '상', '중', '하'].indexOf(val) < 0) throw new Error('잘못된 점수입니다.');
    id = String(id);
    if (!fams_().some(function (f) { return f.id === id; })) throw new Error('없는 가족입니다.');
    const s = sh_(TAB.judge);
    const idx = rows_(TAB.judge).map(function (r) { return str_(r[0]); }).indexOf(id);
    if (idx >= 0) s.getRange(idx + 2, col + 2, 1, 1).setValues([[val]]);
    else { const row = [id, '', '', '', '', '', '']; row[col + 1] = val; s.appendRow(row); }
    return api_admin(a);
  });
}
