/**
 * 送信機チェック（テレメータ送信機の点検・所在管理）GAS
 *
 * 画面（GitHub Pages：telemetry/index.html）から呼ばれる API です。
 * データを入れるスプレッドシートに「拡張機能 → Apps Script」で貼り付けて使います。
 * 手順は telemetry/README.md を見てください。
 *
 * シート
 *   送信機マスタ：CH / 種別 / 登録部署 / 所属 / 備考 / 更新日時
 *   部署マスタ　：部署 / 表示順 / 受信側の枠数（セントラルモニタに登録できるCHの数）/ 備考
 *   点検履歴　　：ID / 日時 / 日付 / CH / 所在 / 管理状況 / 点検結果1 / 点検結果2 / コメント / 点検者 / 記録種別 / 登録
 *   最新状態　　：CH ごとの今の所在と最後の点検（自動で更新。手で直さない）
 *   スタッフ　　：名前 / 番号（5桁）/ 権限（管理者 または 空欄）
 *   設定　　　　：未点検とみなす日数、ダッシュボード用のAPIキー
 */

const SH = { MASTER: '送信機マスタ', DEPT: '部署マスタ', LOG: '点検履歴', LAST: '最新状態', STAFF: 'スタッフ', CFG: '設定' };
const HEAD = {
  [SH.MASTER]: ['CH', '種別', '登録部署', '所属', '備考', '更新日時'],
  [SH.DEPT]: ['部署', '表示順', '受信側の枠数', '備考'],
  [SH.LOG]: ['ID', '日時', '日付', 'CH', '所在', '管理状況', '点検結果1', '点検結果2', 'コメント', '点検者', '記録種別', '登録'],
  [SH.LAST]: ['CH', '所在', '所在の更新', '最終点検', '管理状況', '点検結果1', '点検結果2', '点検者', '貸出日時'],
  [SH.STAFF]: ['名前', '番号', '権限'],
  [SH.CFG]: ['項目', '値', '説明']
};
const ST_JP = { use: '使用中', keep: '保管中', repair: '修理中', lost: '紛失', other: 'その他', lend: '貸出', back: '返却' };
const JP_ST = Object.keys(ST_JP).reduce((o, k) => (o[ST_JP[k]] = k, o), {});
const MOVE = { lend: 1, back: 1 };
const ME_ROOM = 'ME室';
const TZ = 'Asia/Tokyo';

/* ---------- 入口 ---------- */
function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'JSONを読めません' }); }
  return out_(api(req));
}
function doGet() { return ContentService.createTextOutput('送信機チェック API は動いています'); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function api(req) {
  try {
    const a = req.action;
    if (a === 'summary') {
      if (!req.key || String(req.key) !== cfg_().apiKey) return { ok: false, error: 'APIキーが違います' };
      return { ok: true, data: summary_() };
    }
    const me = who_(req.pin);
    if (a === 'login') return me ? { ok: true, data: me } : { ok: false, error: '番号が見つかりません', auth: true };
    if (!me) return { ok: false, error: 'ログインし直してください', auth: true };
    if (a === 'init') return { ok: true, data: init_(me) };
    if (a === 'push') return { ok: true, data: push_(req.ops || []) };
    if (a === 'range') return { ok: true, data: { rows: scanLog_(String(req.from), String(req.to)) } };
    if (a === 'stats') return { ok: true, data: stats_(String(req.from), String(req.to)) };
    return { ok: false, error: '不明な操作です：' + a };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

/* ---------- 共通 ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function rows_(name) {
  const s = sh_(name), n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, HEAD[name].length).getValues();
}
const dkey_ = t => Utilities.formatDate(new Date(t), TZ, 'yyyy-MM-dd');
const str_ = v => v === null || v === undefined ? '' : String(v).trim();
const ms_ = v => v instanceof Date ? v.getTime() : (v ? new Date(v).getTime() || 0 : 0);
const dt_ = t => t ? new Date(t) : '';
function dateStr_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : str_(v); }
function chSort_(a, b) { return (Number(a) - Number(b)) || String(a).localeCompare(String(b)); }
function writeAll_(name, rows) {
  const s = sh_(name), W = HEAD[name].length, n = s.getLastRow() - 1;
  s.getRange(1, 1, 1, W).setValues([HEAD[name]]);
  if (n > 0) s.getRange(2, 1, n, W).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, W).setValues(rows);
}
function entryOf_(r) {
  return { id: str_(r[0]), t: ms_(r[1]), ch: str_(r[3]), dept: str_(r[4]), st: JP_ST[str_(r[5])] || 'other',
    r1: str_(r[6]), r2: str_(r[7]), memo: str_(r[8]), by: str_(r[9]), k: str_(r[10]) || '点検', u: str_(r[11]) === '未登録' ? 1 : 0 };
}

/* ---------- スタッフ・設定 ---------- */
function staff_() {
  const c = CacheService.getScriptCache(), hit = c.get('staff');
  if (hit) return JSON.parse(hit);
  const list = rows_(SH.STAFF).filter(r => str_(r[0]) && str_(r[1]))
    .map(r => ({ name: str_(r[0]), pin: str_(r[1]).padStart(5, '0'), admin: /管理/.test(str_(r[2])) }));
  c.put('staff', JSON.stringify(list), 300);
  return list;
}
function who_(pin) {
  const p = str_(pin);
  if (!/^\d{5}$/.test(p)) return null;
  const s = staff_().find(x => x.pin === p);
  return s ? { name: s.name, admin: s.admin } : null;
}
function cfg_() {
  const o = {};
  rows_(SH.CFG).forEach(r => { o[str_(r[0])] = str_(r[1]); });
  return { overdueDays: Number(o['未点検とみなす日数']) || 1, apiKey: o['APIキー'] || '' };
}
function setCfg_(key, val) {
  const s = sh_(SH.CFG), v = rows_(SH.CFG), i = v.findIndex(r => str_(r[0]) === key);
  if (i >= 0) s.getRange(i + 2, 2).setValues([[val]]); else s.getRange(s.getLastRow() + 1, 1, 1, 2).setValues([[key, val]]);
}

/* ---------- 読み込み ---------- */
function lastObj_(r) {
  return { dept: str_(r[1]), mt: ms_(r[2]), t: ms_(r[3]), st: JP_ST[str_(r[4])] || '', r1: str_(r[5]), r2: str_(r[6]), by: str_(r[7]), lt: ms_(r[8]) };
}
function init_(me) {
  const dr = rows_(SH.DEPT).filter(r => str_(r[0])).sort((a, b) => (Number(a[1]) || 999) - (Number(b[1]) || 999));
  const depts = dr.map(r => str_(r[0])), info = {};
  dr.forEach(r => { info[str_(r[0])] = { slots: r[2] === '' ? '' : Number(r[2]) || 0, note: str_(r[3]) }; });
  const ch = {};
  rows_(SH.MASTER).forEach(r => { const c = str_(r[0]); if (c) ch[c] = { dept: str_(r[2]), kind: str_(r[1]) || '送信機', home: str_(r[3]), note: str_(r[4]) }; });
  const last = {};
  rows_(SH.LAST).forEach(r => { const c = str_(r[0]); if (c) last[c] = lastObj_(r); });
  const c = cfg_();
  return { me, depts, info, ch, last, cfg: { overdueDays: c.overdueDays },
    recent: scanLog_(dkey_(Date.now() - 864e5), dkey_(Date.now() + 864e5)), now: Date.now() };
}
/* 点検履歴を下（新しい方）から読み、期間内の行を集める。期間より古い行だけのかたまりになったら止める */
function scanLog_(from, to) {
  const s = sh_(SH.LOG), W = HEAD[SH.LOG].length, CH = 3000, out = [];
  let end = s.getLastRow();
  while (end >= 2) {
    const start = Math.max(2, end - CH + 1);
    const vals = s.getRange(start, 1, end - start + 1, W).getValues();
    let anyNewer = false;
    for (let i = vals.length - 1; i >= 0; i--) {
      const d = dateStr_(vals[i][2]);
      if (d >= from) anyNewer = true;
      if (d >= from && d <= to && str_(vals[i][0])) out.push(entryOf_(vals[i]));
    }
    if (!anyNewer) break;
    end = start - 1;
  }
  return out.sort((a, b) => a.t - b.t);
}

/* ---------- 最新状態の更新 ---------- */
/* row：[CH, 所在, 所在の更新, 最終点検, 管理状況, 結果1, 結果2, 点検者, 貸出日時] */
function applyEntryToLast_(row, e) {
  const t = Number(e.t);
  if (!row) row = [String(e.ch), '', '', '', '', '', '', '', ''];
  if (t >= ms_(row[2])) { row[1] = e.dept; row[2] = new Date(t); }
  if (MOVE[e.st]) {
    if (t >= ms_(row[2]) - 1) row[8] = e.st === 'lend' ? new Date(t) : '';
  } else if (t >= ms_(row[3])) {
    row[3] = new Date(t); row[4] = ST_JP[e.st] || 'その他'; row[5] = e.r1 || ''; row[6] = e.r2 || ''; row[7] = e.by || '';
  }
  return row;
}
function recomputeLast_(log, ch) {
  const W = HEAD[SH.LOG].length, hits = [];
  let end = log.getLastRow();
  while (end >= 2) {
    const start = Math.max(2, end - 3000 + 1), vals = log.getRange(start, 1, end - start + 1, W).getValues();
    for (let i = vals.length - 1; i >= 0; i--) if (str_(vals[i][3]) === ch) hits.push(entryOf_(vals[i]));
    const insp = hits.some(h => !MOVE[h.st]), move = hits.some(h => MOVE[h.st]);
    if (insp && move) break;
    end = start - 1;
  }
  if (!hits.length) return null;
  let row = null;
  hits.sort((a, b) => a.t - b.t).forEach(h => { row = applyEntryToLast_(row, h); });
  return row;
}

/* ---------- 書き込み（送信箱からまとめて届く。操作ごとに番号を確認） ---------- */
function push_(ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const log = sh_(SH.LOG), W = HEAD[SH.LOG].length;
    const lastRow = log.getLastRow();
    const tailN = Math.min(Math.max(lastRow - 1, 0), 5000);
    const tail = tailN ? log.getRange(lastRow - tailN + 1, 1, tailN, W).getValues() : [];
    const seen = new Set(tail.map(r => str_(r[0])));
    const tailStart = lastRow - tailN + 1, tailIdx = {};
    tail.forEach((r, i) => { tailIdx[str_(r[0])] = i; });
    const dirtyTail = new Set();
    /* 記録を直せるのは、記録した本人か管理者だけ */
    const canEdit = (me, row) => me.admin || str_(row[9]).replace(/（一括）$/, '') === me.name;
    const findRow = id => { const a = add.find(r => r[0] === id); if (a) return { row: a }; const i = tailIdx[id]; return i === undefined ? null : { row: tail[i], i }; };
    const lastMap = {};
    rows_(SH.LAST).forEach(r => { const c = str_(r[0]); if (c) lastMap[c] = r; });
    const master = rows_(SH.MASTER), mIdx = {};
    master.forEach((r, i) => { mIdx[str_(r[0])] = i; });
    let depts = rows_(SH.DEPT).filter(r => str_(r[0])).sort((a, b) => (Number(a[1]) || 999) - (Number(b[1]) || 999));
    let masterDirty = false, deptDirty = false, lastDirty = false;
    const add = [], done = [], rejected = [], delIds = new Set(), recompute = new Set();

    ops.forEach(op => {
      const me = who_(op.pin);
      if (!me) { rejected.push({ id: op.id, error: '番号が確認できません' }); return; }
      const adminOnly = ['chdel', 'dept', 'deptdel', 'cfg'].indexOf(op.k) >= 0 || (op.k === 'ch' && String(op.ch) in mIdx) || (op.k === 'deptset' && op.slots !== undefined);
      if (adminOnly && !me.admin) { rejected.push({ id: op.id, error: '管理者だけができる操作です' }); return; }

      if (op.k === 'e') {
        (op.es || []).forEach(e => {
          if (!e || !e.id || seen.has(String(e.id))) return;
          seen.add(String(e.id));
          const kind = MOVE[e.st] ? ST_JP[e.st] : (e.k === '一括' ? '一括' : '点検');
          const by = me.name + (kind === '一括' ? '（一括）' : '');
          const t = new Date(Number(e.t) || Date.now()), c = String(e.ch);
          add.push([String(e.id), t, dkey_(t), c, str_(e.dept), ST_JP[e.st] || 'その他', str_(e.r1), str_(e.r2), str_(e.memo), by, kind,
            (c in mIdx) ? '登録済' : '未登録']);
          lastMap[c] = applyEntryToLast_(lastMap[c], { t: t.getTime(), ch: c, dept: str_(e.dept), st: e.st, r1: str_(e.r1), r2: str_(e.r2), by });
          lastDirty = true;
        });
      } else if (op.k === 'd' && op.e && op.e.id) {
        const f = findRow(String(op.e.id));
        if (f && !canEdit(me, f.row)) { rejected.push({ id: op.id, error: '自分の記録だけ削除できます' }); return; }
        if (f && f.i === undefined) add.splice(add.indexOf(f.row), 1); else delIds.add(String(op.e.id));
        recompute.add(String(op.e.ch));
      } else if (op.k === 'upd' && op.e && op.e.id) {
        const f = findRow(String(op.e.id));
        if (!f) { rejected.push({ id: op.id, error: '直す記録が見つかりません（古い記録はスプレッドシートで直してください）' }); return; }
        if (!canEdit(me, f.row)) { rejected.push({ id: op.id, error: '自分の記録だけ編集できます' }); return; }
        const r = f.row, e = op.e, isMove = r[10] === '貸出' || r[10] === '返却';
        r[4] = str_(e.dept);
        if (!isMove) { r[5] = ST_JP[e.st] || 'その他'; r[6] = str_(e.r1); r[7] = str_(e.r2); }
        r[8] = str_(e.memo);
        if (f.i !== undefined) dirtyTail.add(f.i);
        recompute.add(str_(r[3]));
      } else if (op.k === 'ch' && op.ch) {
        const c = String(op.ch), now = new Date();
        if (c in mIdx) { const r = master[mIdx[c]]; if (op.dept !== undefined) r[2] = op.dept; if (op.kind) r[1] = op.kind; if (op.home !== undefined) r[3] = op.home; r[5] = now; }
        else { mIdx[c] = master.length; master.push([c, op.kind || '送信機', op.dept || '', op.home || '', '', now]); }
        masterDirty = true;
      } else if (op.k === 'chdel' && op.ch) {
        const c = String(op.ch);
        if (c in mIdx) { master[mIdx[c]] = null; delete mIdx[c]; masterDirty = true; }
      } else if (op.k === 'dept' && op.name) {
        if (!depts.some(r => str_(r[0]) === op.name)) { depts.push([op.name, depts.length + 1, op.slots === undefined ? '' : op.slots, '']); deptDirty = true; }
      } else if (op.k === 'deptset' && op.name) {
        const r = depts.find(x => str_(x[0]) === op.name);
        if (r) {
          if (op.slots !== undefined) r[2] = op.slots === '' ? '' : Math.max(0, Math.min(999, Number(op.slots) || 0));
          if (op.note !== undefined) r[3] = String(op.note).slice(0, 300);
          deptDirty = true;
        }
      } else if (op.k === 'deptdel' && op.name) {
        depts = depts.filter(r => str_(r[0]) !== op.name); deptDirty = true;
      } else if (op.k === 'cfg' && op.key === 'overdueDays') {
        setCfg_('未点検とみなす日数', Math.max(1, Math.min(90, Number(op.value) || 1)));
      }
      done.push(op.id);
    });

    dirtyTail.forEach(i => log.getRange(tailStart + i, 1, 1, W).setValues([tail[i]]));
    if (add.length) log.getRange(log.getLastRow() + 1, 1, add.length, W).setValues(add);
    if (delIds.size) deleteLogRows_(log, delIds);
    recompute.forEach(c => {
      const row = recomputeLast_(log, c);
      if (row) lastMap[c] = row; else delete lastMap[c];
      lastDirty = true;
    });
    if (lastDirty) writeAll_(SH.LAST, Object.keys(lastMap).sort(chSort_).map(c => lastMap[c].map((v, i) => (i >= 2 && i !== 4 && i !== 5 && i !== 6 && i !== 7) ? dt_(ms_(v)) : v)));
    if (masterDirty) writeAll_(SH.MASTER, master.filter(Boolean).sort((a, b) => chSort_(str_(a[0]), str_(b[0]))));
    if (deptDirty) writeAll_(SH.DEPT, depts.map((r, i) => [str_(r[0]), i + 1, r[2] === undefined ? '' : r[2], str_(r[3])]));
    return { done, rejected };
  } finally {
    lock.releaseLock();
  }
}
function deleteLogRows_(log, ids) {
  const last = log.getLastRow(), n = Math.min(last - 1, 10000);
  if (n < 1) return;
  const start = last - n + 1, col = log.getRange(start, 1, n, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) if (ids.has(str_(col[i][0]))) log.deleteRow(start + i);
}

/* ---------- 稼働率などの集計 ----------
 * 各日の状態は「その日までの最後の記録」で決める（点検のない日は前の状態を引き継ぐ）。
 * 稼働率＝使用中だったCH・日数 ÷（CH数×日数）。所在も同じ考え方で日ごとに決める。 */
function stats_(from, to) {
  const today = dkey_(Date.now());
  if (to > today) to = today;
  if (!/^\d{4}-\d\d-\d\d$/.test(from) || !/^\d{4}-\d\d-\d\d$/.test(to) || from > to) return { from, to, days: [], depts: {}, chs: {}, total: {} };
  const master = {}; rows_(SH.MASTER).forEach(r => { const c = str_(r[0]); if (c) master[c] = { dept: str_(r[2]), home: str_(r[3]) }; });
  const slots = {}; rows_(SH.DEPT).forEach(r => { const d = str_(r[0]); if (d) slots[d] = r[2] === '' ? null : Number(r[2]) || 0; });
  const log = sh_(SH.LOG), n = log.getLastRow() - 1, ev = {};
  if (n > 0) log.getRange(2, 1, n, HEAD[SH.LOG].length).getValues().forEach(r => {
    if (!str_(r[0]) || dateStr_(r[2]) > to) return;
    const e = entryOf_(r); e.d = dateStr_(r[2]); (ev[e.ch] || (ev[e.ch] = [])).push(e);
  });
  const days = []; for (let t = new Date(from + 'T00:00:00+09:00').getTime(); dkey_(t) <= to; t += 864e5) days.push(dkey_(t));
  const D = days.length, chs = {}, depts = {}, total = { chDays: 0, use: 0, repair: 0, lost: 0, insp: 0, bulk: 0, abn: 0, dirty: 0, lend: 0, back: 0 };
  const dept = name => depts[name] || (depts[name] = { slots: slots[name] === undefined ? null : slots[name], present: new Array(D).fill(0), use: new Array(D).fill(0), chDays: 0, useDays: 0, inspDays: 0, insp: 0, bulk: 0, abn: 0, dirty: 0, alarm: 0, lend: 0 });
  const isAbn = e => e.st === 'lost' || (e.st === 'use' && e.r1 && e.r1 !== '問題なし') || (e.st === 'keep' && ((e.r1 && e.r1 !== '正常') || e.r2 === '未実施'));
  const useDaily = new Array(D).fill(0);
  Object.keys(master).concat(Object.keys(ev).filter(c => !master[c])).forEach(c => {
    const list = (ev[c] || []).sort((a, b) => a.t - b.t);
    const st = { dept: master[c] ? master[c].dept : '', s: '', lt: 0 };
    let i = 0;
    const o = chs[c] = { days: 0, use: 0, keep: 0, repair: 0, lost: 0, unknown: 0, lendDays: 0, inspDays: 0, insp: 0, bulk: 0, abn: 0, dirty: 0 };
    days.forEach((d, k) => {
      let inspToday = false;
      while (i < list.length && list[i].d <= d) {
        const e = list[i++];
        st.dept = e.dept || st.dept;
        if (MOVE[e.st]) {
          st.lt = e.st === 'lend' ? e.t : 0;
          if (e.d >= from) { if (e.st === 'lend') { total.lend++; dept(e.dept).lend++; } else total.back++; }
        }
        else {
          st.s = e.st;
          if (e.d >= from) {
            const dp = dept(e.dept || st.dept);
            if (e.k === '一括') { o.bulk++; dp.bulk++; total.bulk++; } else { o.insp++; dp.insp++; total.insp++; }
            if (isAbn(e)) { o.abn++; dp.abn++; total.abn++; }
            if (e.r1 === '汚染') { o.dirty++; dp.dirty++; total.dirty++; }
            if (e.st === 'use' && e.r1 && e.r1 !== '問題なし') dp.alarm++;
          }
          if (e.d === d) inspToday = true;
        }
      }
      const dp = dept(st.dept || '（所在不明）');
      o.days++; dp.present[k]++; dp.chDays++; total.chDays++;
      if (inspToday) { o.inspDays++; dp.inspDays++; }
      if (st.lt && st.dept !== ME_ROOM) o.lendDays++;
      if (!st.s) o.unknown++;
      else if (st.s === 'use') { o.use++; dp.use[k]++; dp.useDays++; total.use++; useDaily[k]++; }
      else if (st.s === 'keep') o.keep++;
      else if (st.s === 'repair') { o.repair++; total.repair++; }
      else if (st.s === 'lost') { o.lost++; total.lost++; }
    });
  });
  return { from, to, days, depts, chs, total, useDaily, chCount: Object.keys(chs).length };
}

/* ---------- ME業務ダッシュボード用の要約 ---------- */
function summary_() {
  const c = cfg_(), now = Date.now(), DAY = 864e5, today = dkey_(now);
  const master = {}; rows_(SH.MASTER).forEach(r => { const k = str_(r[0]); if (k) master[k] = { dept: str_(r[2]), home: str_(r[3]) }; });
  const last = {}; rows_(SH.LAST).forEach(r => { const k = str_(r[0]); if (k) last[k] = lastObj_(r); });
  const items = [], days = t => Math.floor((now - t) / DAY);
  let doneToday = 0, overdue = 0;
  Object.keys(master).sort(chSort_).forEach(ch => {
    const l = last[ch];
    if (l && l.t && dkey_(l.t) === today) doneToday++;
    const od = !l || !l.t || days(l.t) >= c.overdueDays;
    if (od && !(l && l.st === 'repair')) overdue++;
    if (l && l.st === 'lost') items.push({ ch, kind: '紛失', where: l.dept, days: days(l.t) });
    else if (l && l.st === 'repair') items.push({ ch, kind: '修理中', where: l.dept, days: days(l.t) });
    else if (l && l.st === 'keep' && (l.r1 === '破損' || l.r1 === '汚染')) items.push({ ch, kind: l.r1, where: l.dept, days: days(l.t) });
    if (l && l.lt && master[ch].home === 'ME在庫' && l.dept !== ME_ROOM) items.push({ ch, kind: '貸出中', where: l.dept, days: days(l.lt) });
  });
  return { at: Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'), total: Object.keys(master).length, doneToday, overdue,
    overdueDays: c.overdueDays, items };
}

/* ---------- はじめに一度だけ実行 ---------- */
function 初期設定() {
  Object.keys(HEAD).forEach(sh_);
  sh_(SH.LOG).getRange('C:D').setNumberFormat('@');
  sh_(SH.LOG).getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  sh_(SH.MASTER).getRange('A:A').setNumberFormat('@');
  sh_(SH.LAST).getRange('A:A').setNumberFormat('@');
  sh_(SH.STAFF).getRange('B:B').setNumberFormat('@');
  const ss = ss_();
  const src = ss.getSheetByName('チャネル台帳');   // 「送信機チャネル台帳.xlsx」を取り込んだ場合
  if (src && sh_(SH.MASTER).getLastRow() < 2) {
    const v = src.getDataRange().getValues(), h = v[0].map(String), ix = n => h.indexOf(n);
    writeAll_(SH.MASTER, v.slice(1).filter(r => str_(r[ix('CH')])).map(r => [
      str_(r[ix('CH')]), str_(r[ix('種別')]) || '送信機', str_(r[ix('所在部署')]), str_(r[ix('所属')]), str_(r[ix('要確認')]), new Date()]));
  }
  const dsrc = ss.getSheetByName('部署');
  if (dsrc && sh_(SH.DEPT).getLastRow() < 2) {
    const v = dsrc.getDataRange().getValues().slice(1).map(r => str_(r[0])).filter(n => n && n !== '合計');
    writeAll_(SH.DEPT, v.map((n, i) => [n, i + 1, '', '']));
  }
  writeAll_(SH.DEPT, rows_(SH.DEPT).filter(r => str_(r[0])));   // 部署マスタの見出しを最新の列に合わせる
  if (rows_(SH.CFG).length === 0) {
    writeAll_(SH.CFG, [
      ['未点検とみなす日数', 1, '最後の点検からこの日数がたつと「未点検」として赤く表示します。アプリの設定画面からも変えられます'],
      ['APIキー', Utilities.getUuid().replace(/-/g, ''), 'ME業務ダッシュボードから要約を読むときの合言葉。外に出さない']]);
  }
  CacheService.getScriptCache().remove('staff');
  ss.toast('初期設定が終わりました。「スタッフ」シートに名前・5桁の番号・権限を入れてください。', '送信機チェック', 8);
}

/* スタッフシートを直したらすぐ反映したいときに実行（普段は5分で自動反映） */
function スタッフの変更を反映() { CacheService.getScriptCache().remove('staff'); }

/* ---------- 年1回くらい：前年までの点検履歴を別ファイルへ移して軽くする ---------- */
function 前年までをアーカイブ() {
  const log = sh_(SH.LOG), W = HEAD[SH.LOG].length, n = log.getLastRow() - 1;
  if (n < 1) return;
  const y = Number(Utilities.formatDate(new Date(), TZ, 'yyyy')), cut = y + '-01-01';
  const vals = log.getRange(2, 1, n, W).getValues();
  const old = vals.filter(r => dateStr_(r[2]) < cut), keep = vals.filter(r => dateStr_(r[2]) >= cut);
  if (!old.length) return;
  const file = SpreadsheetApp.create('送信機点検履歴_' + (y - 1) + '年まで');
  const s = file.getSheets()[0];
  s.setName(SH.LOG);
  s.getRange(1, 1, 1, W).setValues([HEAD[SH.LOG]]);
  s.getRange(2, 1, old.length, W).setValues(old);
  log.getRange(2, 1, n, W).clearContent();
  if (keep.length) log.getRange(2, 1, keep.length, W).setValues(keep);
  ss_().toast(old.length + '行を「' + file.getName() + '」へ移しました', '送信機チェック', 8);
}
