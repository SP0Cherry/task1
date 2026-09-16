/* 课程表在线转换
 * 学生课表(网格) + 总课表(课程目录) -> 课表模板(课程名称/星期/开始节数/结束节数/老师/地点/周数)
 */
'use strict';

const CN_NUM = { '一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10,'十一':11,'十二':12 };
const DAY_NAME = { 1:'周一',2:'周二',3:'周三',4:'周四',5:'周五',6:'周六',7:'周日' };

function txt(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' && Number.isFinite(v) && Number.isInteger(v)) return String(v);
  return String(v).trim();
}

/* 中文数字节次 -> 序号，如 第一节/第1节/十一节 */
function periodIndex(label) {
  const m = String(label).match(/第?\s*([一二三四五六七八九十]+|\d{1,2})\s*[节]/);
  if (!m) return null;
  const t = m[1];
  if (/^\d+$/.test(t)) return parseInt(t, 10);
  if (t === '十') return 10;
  if (t.startsWith('十')) return 10 + (CN_NUM[t.slice(1)] || 0);
  if (t.endsWith('十')) return CN_NUM[t[0]] * 10;
  let n = 0;
  for (const ch of t) n = n * 10 + (CN_NUM[ch] || 0);
  return n || null;
}

function dayIndex(label) {
  const m = String(label).match(/(?:周|星期)\s*([一二三四五六七日天])/);
  if (!m) return null;
  const ch = m[1];
  return ch === '日' || ch === '天' ? 7 : CN_NUM[ch];
}

/* 课程核心名：去掉 -01班 / （1班） 等班级后缀与空白 */
function coreName(name) {
  let t = String(name).replace(/[\s\u3000]+/g, '');
  t = t.replace(/[（(][^（）()]*班[）)]$/, '');
  t = t.replace(/[-－—]\s*\d+\s*班$/, '');
  return t;
}

function nameMatches(a, b) {
  a = coreName(a); b = coreName(b);
  if (!a || !b) return false;
  return a === b || (a.length >= 3 && (a.includes(b) || b.includes(a)));
}

function periodsToTod(ps, pe) {
  const res = new Set();
  for (let p = ps; p <= pe; p++) {
    if (p <= 4) res.add('上午');
    else if (p <= 8) res.add('下午');
    else res.add('晚上');
  }
  return res;
}

function todCompat(studentTods, masterTod) {
  const t = String(masterTod || '').replace(/\s/g, '');
  if (!t || t.includes('待定')) return true;
  const mt = new Set();
  if (t.includes('上午')) mt.add('上午');
  if (t.includes('下午')) mt.add('下午');
  if (t.includes('晚上')) mt.add('晚上');
  if (mt.size === 0) return true;
  for (const x of studentTods) if (mt.has(x)) return true;
  return false;
}

function roomOf(dayCell) {
  const s = String(dayCell);
  const i = s.indexOf('/');
  if (i < 0) return '';
  return s.slice(i + 1).replace(/\s+/g, '').replace(/^[、，,]+|[、，,]+$/g, '');
}

/* ---------- 解析学生课表（节数 × 星期 网格） ---------- */
function parseStudent(rows) {
  let hr = -1, dc = -1;
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < (rows[r] || []).length; c++) {
      const v = txt(rows[r][c]);
      if (v.includes('节') && (v.includes('次') || v.includes('数'))) { hr = r; dc = c; break; }
    }
    if (hr >= 0) break;
  }
  if (hr < 0) throw new Error('未在个人课表中找到“节数/节次”表头，请确认文件格式与示例 schedule 一致。');

  const dayCols = new Map();
  for (let c = 0; c < (rows[hr] || []).length; c++) {
    const d = dayIndex(txt(rows[hr][c]));
    if (d) dayCols.set(d, c);
  }
  if (dayCols.size === 0) throw new Error('未在个人课表中找到“星期一~星期日”列。');

  const blocks = [];
  for (const [day, col] of dayCols) {
    let r = hr + 1;
    while (r < rows.length) {
      const pidx = periodIndex(txt((rows[r] || [])[dc]));
      const cell = txt((rows[r] || [])[col]);
      if (pidx !== null && cell) {
        const name = cell.split(/[（(]/)[0].trim();
        let remark = '';
        const rm = cell.match(/[（(]\s*备注[：:]\s*([\s\S]*?)\s*[）)]/);
        if (rm) remark = rm[1].trim();
        let flag = '';
        const fm = cell.match(/单双周|单周|双周|每周/);
        if (fm) flag = fm[0];
        let ps = pidx, pe = pidx;
        let r2 = r + 1;
        while (r2 < rows.length) {
          const sameText = txt((rows[r2] || [])[col]) === cell;
          const nextP = periodIndex(txt((rows[r2] || [])[dc]));
          if (sameText && nextP === pe + 1) { pe++; r2++; } else break;
        }
        if (name) blocks.push({ name, day, ps, pe, remark, flag });
        r = r2;
      } else {
        r++;
      }
    }
  }
  blocks.sort((a, b) => a.day - b.day || a.ps - b.ps);
  return blocks;
}

/* ---------- 解析总课表（自动识别每个 sheet 的表头与星期列） ---------- */
function parseMaster(wb) {
  const records = [];
  const defaultWeeks = 18;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    if (!rows.length) continue;

    // sheet 级标题（学期起始日期 / 总周数）
    let sheetStart = null, sheetWeeks = defaultWeeks;
    const headText = rows.slice(0, 3).map(row => row.map(txt).join(' ')).join(' ');
    let md = headText.match(/(20\d{2})[.\-/年](\d{1,2})[.\-/月](\d{1,2})/);
    if (md) {
      const d = new Date(+md[1], +md[2] - 1, +md[3]);
      if (!isNaN(d)) sheetStart = d;
    }
    const mw = headText.match(/共\s*(\d+)\s*周/);
    if (mw) sheetWeeks = parseInt(mw[1], 10);

    for (let hr = 0; hr < rows.length; hr++) {
      const cols = {};
      for (let c = 0; c < (rows[hr] || []).length; c++) {
        const v = txt(rows[hr][c]);
        if (v === '课程名称' && !('name' in cols)) cols.name = c;
        else if ((v === '课程编号' || v === '序号') && !('code' in cols)) cols.code = c;
        else if (v.includes('教师') && !('teacher' in cols)) cols.teacher = c;
        else if (v === '时间' && !('tod' in cols)) cols.tod = c;
        else if (v === '学时' && !('hours' in cols)) cols.hours = c;
        else if (v === '备注' && !('remark' in cols)) cols.remark = c;
      }
      if (!('name' in cols)) continue;

      // 表头下方 1~2 行内寻找 周X 子表头
      const dayCols = new Map();
      for (let dr = hr + 1; dr < Math.min(hr + 3, rows.length); dr++) {
        for (let c = 0; c < (rows[dr] || []).length; c++) {
          const v = txt(rows[dr][c]).replace(/\s/g, '');
          const m = v.match(/^(?:周|星期)([一二三四五六七日天])$/);
          if (m) {
            const ch = m[1];
            dayCols.set(ch === '日' || ch === '天' ? 7 : CN_NUM[ch], c);
          }
        }
      }
      if (dayCols.size === 0) continue;

      // 本表标题（同一 sheet 内可能有多个排课表）
      let start = sheetStart, weeks = sheetWeeks;
      const tblText = rows.slice(Math.max(0, hr - 3), hr).map(row => row.map(txt).join(' ')).join(' ');
      md = tblText.match(/(20\d{2})[.\-/年](\d{1,2})[.\-/月](\d{1,2})/);
      if (md) {
        const d = new Date(+md[1], +md[2] - 1, +md[3]);
        if (!isNaN(d)) start = d;
      }
      const mw2 = tblText.match(/共\s*(\d+)\s*周/);
      if (mw2) weeks = parseInt(mw2[1], 10);

      for (let r = hr + 1; r < rows.length; r++) {
        const name = txt((rows[r] || [])[cols.name]);
        const codev = 'code' in cols ? txt((rows[r] || [])[cols.code]) : '';
        if (!name) continue;
        if (name.includes('上课时间') || name === '课程名称' || codev === '序号' || codev === '课程编号') break;

        const daymap = {};
        for (const [day, c] of dayCols) {
          const v = txt((rows[r] || [])[c]);
          if (v) daymap[day] = v;
        }
        if (Object.keys(daymap).length === 0) continue;

        const rec = {
          sheet: sheetName, name,
          teacher: 'teacher' in cols ? txt((rows[r] || [])[cols.teacher]) : '',
          tod: 'tod' in cols ? txt((rows[r] || [])[cols.tod]) : '',
          remark: 'remark' in cols ? txt((rows[r] || [])[cols.remark]) : '',
          code: 'code' in cols ? txt((rows[r] || [])[cols.code]) : '',
          hours: null, daymap, start, weeks,
        };
        if ('hours' in cols) {
          const hv = txt((rows[r] || [])[cols.hours]);
          if (/^\d{1,3}$/.test(hv)) rec.hours = parseInt(hv, 10);
        }
        records.push(rec);
      }
    }
  }
  return records;
}

/* ---------- 匹配学生课程块 -> 总课表记录 ---------- */
function matchBlock(b, records) {
  const stods = periodsToTod(b.ps, b.pe);
  const cands = [];
  for (const rec of records) {
    if (!nameMatches(b.name, rec.name)) continue;
    if (!(b.day in rec.daymap)) continue;
    if (!todCompat(stods, rec.tod)) continue;

    let score = 0;
    if (coreName(b.name) === coreName(rec.name)) score += 5;
    const bc = new Set((b.remark.match(/\d+(?=\s*班)/g) || []));
    const rc = new Set(((rec.name + rec.remark).match(/\d+(?=\s*班)/g) || []));
    if (bc.size && rc.size) for (const x of bc) if (rc.has(x)) score += 6;
    const rk = new Set(b.remark.match(/[一-龥]{2,}/g) || []);
    const mk = new Set(rec.remark.match(/[一-龥]{2,}/g) || []);
    let overlap = 0;
    for (const x of rk) if (mk.has(x)) overlap++;
    score += 2 * overlap;
    if (roomOf(rec.daymap[b.day])) score += 1;
    if (rec.teacher) score += 0.5;
    cands.push({ score, rec });
  }
  cands.sort((a, b2) => b2.score - a.score);
  return cands;
}

/* ---------- 周数推断 ---------- */
function weekText(b, rec) {
  const total = rec.weeks || 18;
  const remark = `${rec.remark || ''} ${b.remark || ''}`;
  const single = remark.includes('单周');
  const double = remark.includes('双周');

  const mr = remark.match(/第?\s*(\d+)\s*[-－—~～]\s*(\d+)\s*周/);
  if (mr) return `${mr[1]}-${mr[2]}${single ? '单' : double ? '双' : ''}`;

  const me = remark.match(/第?\s*([\d、,\s]+?)\s*周/);
  if (me && me[1].includes('、')) {
    const nums = (me[1].match(/\d+/g) || []).map(Number);
    if (nums.length) {
      const segs = [];
      let st = nums[0], prev = nums[0];
      for (let i = 1; i <= nums.length; i++) {
        const n = nums[i];
        if (n === prev + 1) { prev = n; continue; }
        segs.push(st === prev ? String(st) : `${st}-${prev}`);
        if (n !== undefined) st = prev = n;
      }
      return segs.join('、');
    }
  }

  if (single) return `1-${total}单`;
  if (double) return `1-${total}双`;
  if (b.flag === '单周') return `1-${total}单`;
  if (b.flag === '双周') return `1-${total}双`;

  // “X月X日开课”：结合学期起始日与学时推断起止周
  const ms = remark.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*日[^，,。；;]*?开课/);
  if (ms && rec.start) {
    const st = rec.start;
    let y = st.getFullYear();
    const mon = +ms[1];
    if (mon < (st.getMonth() + 1) && st.getMonth() + 1 >= 7) y += 1;
    const d0 = new Date(y, mon - 1, +ms[2]);
    const wk = Math.floor((d0 - st) / 86400000 / 7) + 1;
    if (wk >= 1 && wk <= total) {
      const blockLen = b.pe - b.ps + 1;
      let end = total;
      if (rec.hours && rec.hours % blockLen === 0) {
        const n = rec.hours / blockLen;
        if (n >= 1 && n <= total) end = Math.min(total, wk + n - 1);
      }
      return { text: `${wk}-${end}`, inferred: true };
    }
  }
  return `1-${total}`;
}

/* ---------- 主转换 ---------- */
function convertAll(studentRows, masterWB) {
  const blocks = parseStudent(studentRows);
  const records = parseMaster(masterWB);
  const warnings = [];
  const results = [];

  for (const b of blocks) {
    const cands = matchBlock(b, records);
    if (!cands.length) {
      warnings.push({ level: 'error', text: `「${b.name}」（${DAY_NAME[b.day]} 第${b.ps}-${b.pe}节）在总课表中未找到，请手动补充老师与地点。` });
      results.push({
        课程名称: b.name, 星期: String(b.day), 开始节数: String(b.ps), 结束节数: String(b.pe),
        老师: '', 地点: '', 周数: '1-18',
      });
      continue;
    }
    const tied = cands.length > 1 && cands[0].score === cands[1].score;
    const { rec } = cands[0];
    const wk = weekText(b, rec);
    const week = typeof wk === 'string' ? wk : wk.text;
    if (typeof wk === 'object' && wk.inferred) {
      warnings.push({ level: 'info', text: `「${b.name}」周数根据开课日期与学时自动推断为 ${week}，请核对。` });
    }
    if (tied) {
      const names = cands.slice(0, 3).map(c => c.rec.name).join(' / ');
      warnings.push({ level: 'warn', text: `「${b.name}」在总课表中有多个相同条件的班级（${names}），已默认选第一个，请核对老师/地点。` });
    }
    results.push({
      课程名称: b.name,
      星期: String(b.day),
      开始节数: String(b.ps),
      结束节数: String(b.pe),
      老师: rec.teacher,
      地点: roomOf(rec.daymap[b.day]),
      周数: week,
    });
  }
  return { results, warnings };
}

/* ---------- 读取 Excel ---------- */
function readWorkbook(data) {
  return XLSX.read(data, { type: 'array' });
}
function b64ToBytes(b64) {
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}
