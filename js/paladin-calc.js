/*!
 * 圣骑士环破伤害计算器 · 网页版
 * 逐条复刻 Excel「圣骑士环破伤害计算器 v1.5」（Wings 圣骑士四转环破木桩伤害）
 * 口径说明：Excel INT() 对正数等价于 Math.trunc()；每分钟攻击次数 = 60000 / 延迟(ms)
 * 剑 / 钝器合并为一张表：每列用「武器类型」下拉切换，系数、熟练度、优劣比例随武器带出
 */
(function () {
  'use strict';

  /* ============================ 模型 ============================ */

  // 原表 G3:K10 攻速表。2 档最快、7 档最慢（9 行延迟为空，原表 J 列为 #DIV/0!，故不提供）
  var SPEED_TABLE = [
    { level: 2, delay: 630 },
    { level: 3, delay: 690 },
    { level: 4, delay: 750 },
    { level: 5, delay: 810 },
    { level: 6, delay: 840 },
    { level: 7, delay: 900 }
  ];

  // 武器类型 → 优势 / 劣势主属性系数、熟练度(%)、优势 / 劣势比例(%)
  var WEAPONS = [
    { text: '单手剑',   advCoef: 1,    disCoef: 1,    mastery: 60, advRatio: 100, disRatio: 0 },
    { text: '双手剑',   advCoef: 1.15, disCoef: 1.15, mastery: 60, advRatio: 100, disRatio: 0 },
    { text: '单手钝器', advCoef: 1.1,  disCoef: 0.8,  mastery: 75, advRatio: 60,  disRatio: 40 },
    { text: '双手钝器', advCoef: 1.2,  disCoef: 0.85, mastery: 75, advRatio: 60,  disRatio: 40 }
  ];

  // I12 是否计算暴击增益 → I13 暴击率 / I14 暴击伤害（暴击伤害以「加点」形式加到技能百分比上）
  var CRIT_OPTIONS = [
    { text: '不计算', rate: 0, dmg: 0 },
    { text: '计算火眼晶晶', rate: 0.15, dmg: 1.4 },
    { text: '计算栗子味粽子', rate: 0.05, dmg: 1.1 },
    { text: '计算火眼与粽子叠加', rate: 0.2, dmg: 1.5 },
    { text: '计算珠光药丸', rate: 0.1, dmg: 1.2 }
  ];

  // I15 是否计算回声 → I16 攻击力增加
  var ECHO_OPTIONS = [
    { text: '不计算', bonus: 0 },
    { text: '计算', bonus: 0.04 }
  ];

  // X23 克制百分比
  var RELATION_OPTIONS = [
    { text: '克制', pct: 1.5 },
    { text: '非克制', pct: 1 },
    { text: '抗性', pct: 0.5 }
  ];

  var HINT_WEAPON = '选好武器类型后，下面的优势 / 劣势主属性系数、熟练度、优劣比例会按该武器自动带出，仍可手动改。' +
    '单手剑 1 / 1、60%、100% / 0%；双手剑 1.15 / 1.15、60%、100% / 0%；' +
    '单手钝器 1.1 / 0.8、75%、60% / 40%；双手钝器 1.2 / 0.85、75%、60% / 40%。';

  var FIELDS = [
    { key: 'weapon',   label: '武器类型',       kind: 'weapon',   noSync: true, hint: HINT_WEAPON },
    { key: 'str',      label: '力量',           kind: 'num', step: 1,  hint: '角色力量总和，包括技能、装备及家族等。' },
    { key: 'dex',      label: '敏捷',           kind: 'num', step: 1,  hint: '角色敏捷总和，包括技能、装备及家族等。' },
    { key: 'atk',      label: '攻击力总和',     kind: 'num', step: 1,  hint: '武器 / 手套 / 盾牌 / 护肩等攻击力合计。' },
    { key: 'spd',      label: '攻速',           kind: 'speed',        hint: '填上快速、急速、粽子等 buff 后的攻速档位。2 档最快（630ms），7 档最慢（900ms）。' },
    { key: 'mastery',  label: '熟练度 %',       kind: 'num', step: 1,  hint: '精通点满：剑 60%，钝器 75%。原表 0.6 / 0.75，本页按百分比填写。' },
    { key: 'advCoef',  label: '优势主属性系数', kind: 'num', step: 0.01, hint: '武器系数，切换武器类型时自动带出，一般无需改动。' },
    { key: 'disCoef',  label: '劣势主属性系数', kind: 'num', step: 0.01, hint: '武器系数，切换武器类型时自动带出，一般无需改动。' },
    { key: 'subCoef',  label: '副属性系数',     kind: 'num', step: 0.01, hint: '武器系数，默认 1，一般无需改动。' },
    { key: 'advRatio', label: '优势比例 %',     kind: 'num', step: 1,  hint: '由劈刺比决定：环破 3:2，剑不分优劣填 100%，钝器填 60%。' },
    { key: 'disRatio', label: '劣势比例 %',     kind: 'num', step: 1,  hint: '剑填 0%，钝器填 40%。优势比例 + 劣势比例应为 100%。' },
    { key: 'def',      label: '怪物物理防御',   kind: 'num', step: 1,  hint: '怪物固有属性，物理防御力数值；打木桩填 0。' },
    { key: 'skill',    label: '技能百分比 %',   kind: 'num', step: 10, hint: '环破点满技改 580%，原表写 5.8。' },
    { key: 'buff',     label: '增益百分比 %',   kind: 'num', step: 10, hint: '技改 130% / 圣 buff 140% / 不开 buff 100%。' },
    { key: 'relation', label: '是否克制',       kind: 'relation',      hint: '默认属性 buff 点满：克制 150% / 非克制 100% / 抗性 50%。' },
    { key: 'crit',     label: '暴击增益',       kind: 'crit',          hint: 'I12：暴击增益的方式，暴击率与暴击伤害自动跟随（火眼晶晶 15% / 140%，栗子味粽子 5% / 110%，叠加 20% / 150%，珠光药丸 10% / 120%）。' },
    { key: 'echo',     label: '英雄回声',       kind: 'echo',          hint: 'I15：选择计算后，攻击力按 INT(攻击力 × 1.04) 参与面板计算。' }
  ];

  var RESULTS = [
    { key: 'advHi', label: '优势面板上限',         dec: 2 },
    { key: 'advLo', label: '优势面板下限',         dec: 2 },
    { key: 'disHi', label: '劣势面板上限',         dec: 2 },
    { key: 'disLo', label: '劣势面板下限',         dec: 2 },
    { key: 'panel', label: '面板期望值',           dec: 2, strong: true },
    { key: 'hits',  label: '每分钟攻击次数',       dec: 2 },
    { key: 'hit',   label: '单次伤害期望',         dec: 2 },
    { key: 'dpm',   label: '每分钟造成伤害（万）', dec: 2, strong: true }
  ];

  var CONVERT_FIELDS = [
    { key: 'advCoef', label: '优势武器系数', kind: 'num', step: 0.01, hint: '武器系数，无需改动。' },
    { key: 'disCoef', label: '劣势武器系数', kind: 'num', step: 0.01, hint: '武器系数，无需改动。' },
    { key: 'str',     label: '力量',         kind: 'num', step: 1,    hint: '角色力量总和。' },
    { key: 'dex',     label: '敏捷',         kind: 'num', step: 1,    hint: '角色敏捷总和。' },
    { key: 'atk',     label: '总攻击力',     kind: 'num', step: 1,    hint: '角色攻击力总和。' },
    { key: 'mastery', label: '熟练度 %',     kind: 'num', step: 1,    hint: '剑 60%，钝器 75%。原表 0.6 / 0.75。' }
  ];

  var CONVERT_RESULTS = [
    { key: 'atkPerStr', label: '攻击 / 力量收益', dec: 4, strong: true },
    { key: 'strPerDex', label: '力量 / 敏捷收益', dec: 4 },
    { key: 'atkPerDex', label: '攻击 / 敏捷收益', dec: 4, strong: true }
  ];

  // 默认四列（原表剑表默认值；钝器用「武器类型」下拉切换）
  var BASE = [
    { name: '武器1', weapon: '单手剑', str: 1200, dex: 120, atk: 200, spd: 2, mastery: 60, advCoef: 1,    disCoef: 1,    subCoef: 1, advRatio: 100, disRatio: 0, def: 0, skill: 580, buff: 130, relation: '克制', crit: '不计算', echo: '不计算' },
    { name: '武器2', weapon: '单手剑', str: 1200, dex: 120, atk: 200, spd: 2, mastery: 60, advCoef: 1,    disCoef: 1,    subCoef: 1, advRatio: 100, disRatio: 0, def: 0, skill: 580, buff: 130, relation: '克制', crit: '不计算', echo: '不计算' },
    { name: '武器3', weapon: '双手剑', str: 1400, dex: 144, atk: 250, spd: 2, mastery: 60, advCoef: 1.15, disCoef: 1.15, subCoef: 1, advRatio: 100, disRatio: 0, def: 0, skill: 580, buff: 130, relation: '克制', crit: '不计算', echo: '不计算' },
    { name: '武器4', weapon: '双手剑', str: 1400, dex: 144, atk: 200, spd: 2, mastery: 60, advCoef: 1.15, disCoef: 1.15, subCoef: 1, advRatio: 100, disRatio: 0, def: 0, skill: 580, buff: 130, relation: '克制', crit: '不计算', echo: '不计算' }
  ];

  var CONVERT_BASE = [
    { name: '双手剑',   advCoef: 1.15, disCoef: 1.15, str: 1400, dex: 144, atk: 200, mastery: 60 },
    { name: '单手剑',   advCoef: 1,    disCoef: 1,    str: 1100, dex: 120, atk: 200, mastery: 60 },
    { name: '双手钝器', advCoef: 1.2,  disCoef: 0.85, str: 1400, dex: 144, atk: 200, mastery: 75 },
    { name: '单手钝器', advCoef: 1.1,  disCoef: 0.8,  str: 1000, dex: 100, atk: 200, mastery: 75 }
  ];

  var DEFAULTS = {
    count: 4,
    columns: JSON.parse(JSON.stringify(BASE)),
    convert: JSON.parse(JSON.stringify(CONVERT_BASE))
  };

  /* ------------------------- 计算 ------------------------- */

  function optBy(list, text, fallback) {
    for (var i = 0; i < list.length; i++) if (list[i].text === text) return list[i];
    return list[fallback || 0];
  }

  function weaponBy(text) { return optBy(WEAPONS, text, 0); }

  function delayOf(level) {
    for (var i = 0; i < SPEED_TABLE.length; i++) if (SPEED_TABLE[i].level === level) return SPEED_TABLE[i].delay;
    return 0;
  }

  // 单方案计算：X14..X25（暴击 / 回声按列取，未显式传参时用该列自己的设置）
  function computeColumn(c, critLabel, echoLabel) {
    var crit = optBy(CRIT_OPTIONS, typeof critLabel === 'string' ? critLabel : c.crit);
    var echo = optBy(ECHO_OPTIONS, typeof echoLabel === 'string' ? echoLabel : c.echo);
    var rel = optBy(RELATION_OPTIONS, c.relation);

    // X4 经回声加成后取整，再除以 100（原表 INT(X4*(1+I16))/100）
    var atkEff = Math.trunc(c.atk * (1 + echo.bonus)) / 100;
    var mastery = c.mastery / 100;

    var advHi = (c.str * c.advCoef * 4 + c.dex * c.subCoef) * atkEff;
    var advLo = (c.str * c.advCoef * 4 * 0.9 * mastery + c.dex * c.subCoef) * atkEff;
    var disHi = (c.str * c.disCoef * 4 + c.dex * c.subCoef) * atkEff;
    var disLo = (c.str * c.disCoef * 4 * 0.9 * mastery + c.dex * c.subCoef) * atkEff;

    var panel = (advHi + advLo) / 2 * (c.advRatio / 100) + (disHi + disLo) / 2 * (c.disRatio / 100);

    // 防御按 0.55 折算（原表 X19*0.55）
    var core = panel * rel.pct * (c.buff / 100) - c.def * 0.55;
    var skill = c.skill / 100;
    var hit = (skill + crit.dmg) * core * crit.rate + skill * core * (1 - crit.rate);

    var delay = delayOf(c.spd);
    var hits = delay > 0 ? 60000 / delay : 0;

    return {
      atkEff: atkEff,
      advHi: advHi, advLo: advLo, disHi: disHi, disLo: disLo,
      panel: panel, relation: rel.pct, core: core,
      delay: delay, hits: hits, hit: hit,
      dpm: hit * hits / 10000,
      ratioSum: c.advRatio + c.disRatio
    };
  }

  // 攻力敏换算：X11 / X12 / X13
  function computeConvert(c) {
    var mastery = c.mastery / 100;
    var k = 0.9 * mastery + 1;
    var den = 12 * c.atk * c.advCoef * k + 8 * c.atk * c.disCoef * k;
    var num = (c.atk + 1) * (12 * c.str * c.advCoef * k + 8 * c.str * c.disCoef * k + 10 * c.dex) - 10 * c.atk * c.dex;
    var atkPerStr = (den === 0 ? NaN : num / den) - c.str;
    var strPerDex = (1 + 0.9 * mastery) * (1.2 * c.advCoef + 0.8 * c.disCoef);
    return { atkPerStr: atkPerStr, strPerDex: strPerDex, atkPerDex: atkPerStr * strPerDex };
  }

  /* ------------------------- 格式化 ------------------------- */

  function groupDigits(s) {
    var neg = s.charAt(0) === '-';
    if (neg) s = s.slice(1);
    s = s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (neg ? '-' : '') + s;
  }

  function fmtNum(v, dec) {
    if (typeof v !== 'number' || !isFinite(v)) return '—';
    var p = Math.pow(10, dec);
    var s = (Math.round(v * p) / p).toFixed(dec);
    if (s.indexOf('.') >= 0) s = s.replace(/0+$/, '').replace(/\.$/, '');
    return groupDigits(s);
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /* ------------------------- 状态 ------------------------- */

  var STORAGE_KEY = 'pldc-state-v2';

  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }

  function num(v, d) {
    if (v === '' || v === null || typeof v === 'undefined') return 0;
    var n = parseFloat(v);
    return isFinite(n) ? n : d;
  }

  function pickSpd(v, d) {
    var n = Math.round(num(v, d));
    for (var i = 0; i < SPEED_TABLE.length; i++) if (SPEED_TABLE[i].level === n) return n;
    return d;
  }

  function pickRelation(v, d) {
    for (var i = 0; i < RELATION_OPTIONS.length; i++) if (RELATION_OPTIONS[i].text === v) return v;
    return d;
  }

  function pickWeapon(v, d) {
    for (var i = 0; i < WEAPONS.length; i++) if (WEAPONS[i].text === v) return v;
    return d;
  }

  function pickText(v, list, d) {
    for (var i = 0; i < list.length; i++) if (list[i].text === v) return v;
    return d;
  }

  function pickCount(v, d) {
    var n = Math.round(num(v, d));
    return n >= 1 && n <= 4 ? n : d;
  }

  function normalizeSlot(src, base, isConvert) {
    var o = clone(base);
    if (!src || typeof src !== 'object') return o;
    if (typeof src.name === 'string' && src.name.trim()) o.name = src.name.trim().slice(0, 12);
    if (isConvert) {
      for (var kc in base) {
        if (!base.hasOwnProperty(kc) || kc === 'name') continue;
        o[kc] = num(src[kc], base[kc]);
      }
      return o;
    }
    o.weapon = pickWeapon(src.weapon, base.weapon);
    o.relation = pickRelation(src.relation, base.relation);
    o.spd = pickSpd(src.spd, base.spd);
    o.crit = pickText(src.crit, CRIT_OPTIONS, base.crit);
    o.echo = pickText(src.echo, ECHO_OPTIONS, base.echo);
    for (var k in base) {
      if (!base.hasOwnProperty(k)) continue;
      if (k === 'name' || k === 'weapon' || k === 'relation' || k === 'spd' || k === 'crit' || k === 'echo') continue;
      o[k] = num(src[k], base[k]);
    }
    return o;
  }

  function normalize(raw) {
    var out = {
      count: pickCount(raw && raw.count, DEFAULTS.count),
      columns: [],
      convert: []
    };
    for (var i = 0; i < 4; i++) {
      out.columns.push(normalizeSlot(raw && raw.columns && raw.columns[i], BASE[i], false));
      out.convert.push(normalizeSlot(raw && raw.convert && raw.convert[i], CONVERT_BASE[i], true));
    }
    return out;
  }

  /* ------------------- 分享链接：紧凑编码 ------------------- */

  // 与默认值不同的字段才写入；系数按 ×100 存整数，其余取整
  var ENC_FIELDS = ['str', 'dex', 'atk', 'spd', 'mastery', 'advCoef', 'disCoef', 'subCoef', 'advRatio', 'disRatio', 'def', 'skill', 'buff'];
  var COEF_KEYS = { advCoef: 1, disCoef: 1, subCoef: 1 };

  function scaleVal(k, v) {
    return Math.round((COEF_KEYS[k] ? v * 100 : v));
  }

  function critIndex(t) {
    for (var i = 0; i < CRIT_OPTIONS.length; i++) if (CRIT_OPTIONS[i].text === t) return i;
    return 0;
  }

  function echoIndex(t) {
    for (var i = 0; i < ECHO_OPTIONS.length; i++) if (ECHO_OPTIONS[i].text === t) return i;
    return 0;
  }

  function weaponIndex(t) {
    for (var i = 0; i < WEAPONS.length; i++) if (WEAPONS[i].text === t) return i;
    return 0;
  }

  function relationIndex(t) {
    for (var i = 0; i < RELATION_OPTIONS.length; i++) if (RELATION_OPTIONS[i].text === t) return i;
    return 0;
  }

  function pushVarint(arr, v) {
    v = Math.max(0, Math.round(v));
    while (v >= 0x80) {
      arr.push((v & 0x7f) | 0x80);
      v = Math.floor(v / 128);
    }
    arr.push(v);
  }

  function readVarint(bytes, st) {
    var shift = 0, val = 0, b;
    do {
      if (st.i >= bytes.length) throw new Error('数据不完整');
      b = bytes[st.i++];
      val += (b & 0x7f) * Math.pow(2, shift);
      shift += 7;
    } while (b & 0x80);
    return val;
  }

  function pushZig(arr, v) { pushVarint(arr, v >= 0 ? v * 2 : -v * 2 - 1); }
  function readZig(bytes, st) { var z = readVarint(bytes, st); return z % 2 ? -(z + 1) / 2 : z / 2; }

  function utf8Bytes(str) {
    var out = [];
    for (var i = 0; i < str.length; i++) {
      var code = str.charCodeAt(i);
      if (code < 0x80) out.push(code);
      else if (code < 0x800) out.push(0xc0 | (code >> 6), 0x80 | (code & 63));
      else if (code >= 0xd800 && code <= 0xdbff && i + 1 < str.length) {
        var cp = 0x10000 + ((code - 0xd800) << 10) + (str.charCodeAt(i + 1) - 0xdc00);
        i++;
        out.push(0xf0 | (cp >> 18), 0x80 | ((cp >> 12) & 63), 0x80 | ((cp >> 6) & 63), 0x80 | (cp & 63));
      } else out.push(0xe0 | (code >> 12), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63));
    }
    return out;
  }

  function utf8Read(bytes, st, len) {
    var end = st.i + len, out = '';
    while (st.i < end) {
      var b = bytes[st.i++], cp, extra;
      if (b < 0x80) { cp = b; extra = 0; }
      else if (b < 0xe0) { cp = b & 0x1f; extra = 1; }
      else if (b < 0xf0) { cp = b & 0x0f; extra = 2; }
      else { cp = b & 0x07; extra = 3; }
      for (var e = 0; e < extra; e++) cp = (cp << 6) | (bytes[st.i++] & 0x3f);
      if (cp > 0xffff) {
        cp -= 0x10000;
        out += String.fromCharCode(0xd800 + (cp >> 10), 0xdc00 + (cp & 0x3ff));
      } else out += String.fromCharCode(cp);
    }
    return out;
  }

  function toB64(bytes) {
    var bin = '';
    for (var i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  function fromB64(str) {
    var b = String(str).replace(/-/g, '+').replace(/_/g, '/');
    while (b.length % 4) b += '=';
    var bin = atob(b), out = [];
    for (var i = 0; i < bin.length; i++) out.push(bin.charCodeAt(i));
    return out;
  }

  function encodeState(s) {
    var count = pickCount(s.count, 4);
    var bytes = [2, count];
    for (var i = 0; i < count; i++) {
      var c = s.columns[i], b = BASE[i];
      var mask = 0, payload = [];
      if (c.weapon !== b.weapon) { mask |= 1; payload.push(weaponIndex(c.weapon)); }
      for (var f = 0; f < ENC_FIELDS.length; f++) {
        var k = ENC_FIELDS[f];
        var cv = scaleVal(k, c[k]), bv = scaleVal(k, b[k]);
        if (cv !== bv) { mask |= (1 << (f + 1)); pushZig(payload, cv); }
      }
      if (c.relation !== b.relation) { mask |= (1 << 14); payload.push(relationIndex(c.relation)); }
      if (c.name !== b.name) { mask |= (1 << 15); var nb = utf8Bytes(c.name); pushVarint(payload, nb.length); for (var n = 0; n < nb.length; n++) payload.push(nb[n]); }
      if (c.crit !== b.crit) { mask |= (1 << 16); payload.push(critIndex(c.crit)); }
      if (c.echo !== b.echo) { mask |= (1 << 17); payload.push(echoIndex(c.echo)); }
      bytes.push(mask & 0xff, (mask >> 8) & 0xff, (mask >> 16) & 0xff);
      for (var p = 0; p < payload.length; p++) bytes.push(payload[p]);
    }
    // 攻力敏换算表：四列 ×(4 系数/熟练度 + 力量/敏捷/攻击力)
    var cfields = ['advCoef', 'disCoef', 'str', 'dex', 'atk', 'mastery'];
    for (var j = 0; j < 4; j++) {
      var cc = s.convert[j], cb = CONVERT_BASE[j];
      var cmask = 0, cpay = [];
      for (var q = 0; q < cfields.length; q++) {
        var ck = cfields[q];
        var a = scaleVal(ck, cc[ck]), d = scaleVal(ck, cb[ck]);
        if (a !== d) { cmask |= (1 << q); pushZig(cpay, a); }
      }
      if (cc.name !== cb.name) { cmask |= (1 << 6); var cb2 = utf8Bytes(cc.name); pushVarint(cpay, cb2.length); for (var m = 0; m < cb2.length; m++) cpay.push(cb2[m]); }
      bytes.push(cmask & 0xff);
      for (var w = 0; w < cpay.length; w++) bytes.push(cpay[w]);
    }
    return toB64(bytes);
  }

  var CONVERT_ENC_FIELDS = ['advCoef', 'disCoef', 'str', 'dex', 'atk', 'mastery'];

  function decodeState(str) {
    var bytes = fromB64(str);
    if (!bytes.length) throw new Error('空数据');
    if (bytes[0] !== 1 && bytes[0] !== 2) {
      // 兼容旧版 base64(JSON) 链接
      var txt = utf8Read(bytes, { i: 0 }, bytes.length);
      if (txt.charAt(0) === '{') return JSON.parse(txt);
      throw new Error('未知格式');
    }

    // v1：全局暴击 / 回声，固定 4 列、2 字节掩码；v2：武器数量 + 每列暴击 / 回声、3 字节掩码
    var ver = bytes[0];
    var st = { i: 1 };
    var count = 4, crit0 = CRIT_OPTIONS[0], echo0 = ECHO_OPTIONS[0];
    if (ver === 1) {
      crit0 = CRIT_OPTIONS[readVarint(bytes, st)] || CRIT_OPTIONS[0];
      echo0 = ECHO_OPTIONS[readVarint(bytes, st)] || ECHO_OPTIONS[0];
    } else {
      count = Math.min(4, Math.max(1, readVarint(bytes, st)));
    }
    var out = { count: count, columns: [], convert: [] };

    for (var i = 0; i < count; i++) {
      var b = BASE[i];
      var slot = { name: b.name, crit: crit0.text, echo: echo0.text };
      var mask = bytes[st.i] | (bytes[st.i + 1] << 8) | ((ver === 2 ? bytes[st.i + 2] : 0) << 16);
      st.i += ver === 2 ? 3 : 2;
      slot.weapon = (mask & 1) ? (WEAPONS[bytes[st.i++]] || WEAPONS[0]).text : b.weapon;
      for (var f = 0; f < ENC_FIELDS.length; f++) {
        var k = ENC_FIELDS[f];
        if (mask & (1 << (f + 1))) {
          var v = readZig(bytes, st);
          slot[k] = COEF_KEYS[k] ? v / 100 : v;
        } else slot[k] = b[k];
      }
      slot.relation = (mask & (1 << 14)) ? (RELATION_OPTIONS[bytes[st.i++]] || RELATION_OPTIONS[0]).text : b.relation;
      if (mask & (1 << 15)) {
        var len = readVarint(bytes, st);
        slot.name = utf8Read(bytes, st, len);
      }
      if (ver === 2 && (mask & (1 << 16))) slot.crit = (CRIT_OPTIONS[bytes[st.i++]] || CRIT_OPTIONS[0]).text;
      if (ver === 2 && (mask & (1 << 17))) slot.echo = (ECHO_OPTIONS[bytes[st.i++]] || ECHO_OPTIONS[0]).text;
      out.columns.push(slot);
    }

    for (var j = 0; j < 4; j++) {
      var cb = CONVERT_BASE[j];
      var cslot = { name: cb.name };
      var cmask = bytes[st.i++];
      for (var q = 0; q < CONVERT_ENC_FIELDS.length; q++) {
        var ck = CONVERT_ENC_FIELDS[q];
        if (cmask & (1 << q)) {
          var cv = readZig(bytes, st);
          cslot[ck] = COEF_KEYS[ck] ? cv / 100 : cv;
        } else cslot[ck] = cb[ck];
      }
      if (cmask & (1 << 6)) {
        var clen = readVarint(bytes, st);
        cslot.name = utf8Read(bytes, st, clen);
      }
      out.convert.push(cslot);
    }
    return out;
  }

  /* ------------------------- 视图 ------------------------- */

  var S = null;
  var root = null;

  function speedOptions(cur) {
    var out = '';
    for (var i = 0; i < SPEED_TABLE.length; i++) {
      var s = SPEED_TABLE[i];
      out += '<option value="' + s.level + '"' + (s.level === cur ? ' selected' : '') + '>' +
        s.level + ' 档 · ' + s.delay + 'ms</option>';
    }
    return out;
  }

  function relationOptions(cur) {
    var out = '';
    for (var i = 0; i < RELATION_OPTIONS.length; i++) {
      var r = RELATION_OPTIONS[i];
      out += '<option value="' + r.text + '"' + (r.text === cur ? ' selected' : '') + '>' +
        r.text + '（' + Math.round(r.pct * 100) + '%）</option>';
    }
    return out;
  }

  function weaponOptions(cur) {
    var out = '';
    for (var i = 0; i < WEAPONS.length; i++) {
      out += '<option value="' + WEAPONS[i].text + '"' + (WEAPONS[i].text === cur ? ' selected' : '') + '>' +
        WEAPONS[i].text + '</option>';
    }
    return out;
  }

  // 单元格里省掉「计算」前缀，避免下拉被截断
  function critOptions(cur) {
    var out = '';
    for (var i = 0; i < CRIT_OPTIONS.length; i++) {
      var t = CRIT_OPTIONS[i].text;
      out += '<option value="' + t + '"' + (t === cur ? ' selected' : '') + '>' +
        t.replace(/^计算/, '') + '</option>';
    }
    return out;
  }

  function echoOptions(cur) {
    var out = '';
    for (var i = 0; i < ECHO_OPTIONS.length; i++) {
      var t = ECHO_OPTIONS[i].text;
      out += '<option value="' + t + '"' + (t === cur ? ' selected' : '') + '>' + t + '</option>';
    }
    return out;
  }

  function cellControl(mode, col, f, slot) {
    var base = ' data-mode="' + mode + '" data-col="' + col + '" data-key="' + f.key + '"';
    if (f.kind === 'speed') return '<select class="pldc-input pldc-select"' + base + '>' + speedOptions(slot.spd) + '</select>';
    if (f.kind === 'relation') return '<select class="pldc-input pldc-select"' + base + '>' + relationOptions(slot.relation) + '</select>';
    if (f.kind === 'weapon') return '<select class="pldc-input pldc-select"' + base + '>' + weaponOptions(slot.weapon) + '</select>';
    if (f.kind === 'crit') return '<select class="pldc-input pldc-select"' + base + '>' + critOptions(slot.crit) + '</select>';
    if (f.kind === 'echo') return '<select class="pldc-input pldc-select"' + base + '>' + echoOptions(slot.echo) + '</select>';
    var attrs = ' type="number" inputmode="decimal" step="' + (f.step || 1) + '"' +
      (typeof f.min === 'number' ? ' min="' + f.min + '"' : '') +
      (typeof f.max === 'number' ? ' max="' + f.max + '"' : '');
    return '<input class="pldc-input"' + base + attrs + ' value="' + slot[f.key] + '">';
  }

  function hintHtml(f) {
    return '<span class="pldc-hint" data-hint="' + esc(f.hint) + '" tabindex="0" role="note" aria-label="说明">?</span>';
  }

  function calcPanel() {
    var N = S.count;
    var html = '<section class="pldc-panel" data-panel="calc">' +
      '<div class="pldc-cards pldc-n' + N + '"></div>' +
      '<p class="pldc-warn" data-warn></p>' +
      '<div class="pldc-tablewrap"><table class="pldc-table pldc-n' + N + '">' +
      '<thead><tr><th class="pldc-rowhead" scope="col">参数</th>';

    for (var i = 0; i < N; i++) {
      html += '<th scope="col" class="pldc-colhead">' +
        '<input class="pldc-name" data-mode="calc" data-col="' + i + '" data-key="name" ' +
        'type="text" maxlength="12" spellcheck="false" value="' + esc(S.columns[i].name) + '" title="点击修改方案名称">' +
        '</th>';
    }
    html += '</tr></thead><tbody>';

    for (var f = 0; f < FIELDS.length; f++) {
      var fld = FIELDS[f];
      html += '<tr' + (fld.key === 'weapon' ? ' class="pldc-row-weapon"' : '') + '>' +
        '<th scope="row" class="pldc-rowhead">' + esc(fld.label) + hintHtml(fld) + '</th>';
      for (var c = 0; c < N; c++) {
        html += '<td class="pldc-cell">' + cellControl('calc', c, fld, S.columns[c]) + '</td>';
      }
      html += '</tr>';
    }

    html += '<tr class="pldc-divider"><th colspan="' + (N + 1) + '" scope="colgroup">计算结果</th></tr>';
    for (var r = 0; r < RESULTS.length; r++) {
      var res = RESULTS[r];
      html += '<tr class="' + (res.strong ? 'pldc-row-strong' : '') + '">' +
        '<th scope="row" class="pldc-rowhead">' + esc(res.label) + '</th>';
      for (var c2 = 0; c2 < N; c2++) {
        html += '<td class="pldc-cell pldc-out" data-out="calc" data-col="' + c2 + '" data-key="' + res.key + '"></td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table></div>' +
      '<p class="pldc-note">每列可单独选武器类型；切换武器会按该武器重置主属性系数、熟练度与优劣比例，之后仍可手动改。' +
      '暴击增益与英雄回声也按列单独选。比例按 <b>%</b> 填写（原表 0.6 = 本页 60%），技能 / 增益百分比同理（原表 5.8 = 本页 580）。' +
      '鼠标移到参数名旁的 <b>?</b> 上看原表说明。</p>' +
      '</section>';
    return html;
  }

  function convertPanel() {
    var html = '<section class="pldc-panel" data-panel="convert" hidden>' +
      '<p class="pldc-verdict">把四把武器的力量 / 敏捷 / 攻击力填进去，得到「1 攻击 ≈ X 点力量」的等效换算。' +
      '本表算整体输出提升，不是面板最高值；未包含英雄回声。</p>' +
      '<div class="pldc-tablewrap"><table class="pldc-table">' +
      '<thead><tr><th class="pldc-rowhead" scope="col">参数</th>';
    for (var i = 0; i < 4; i++) {
      html += '<th scope="col" class="pldc-colhead">' +
        '<input class="pldc-name" data-mode="convert" data-col="' + i + '" data-key="name" ' +
        'type="text" maxlength="12" spellcheck="false" value="' + esc(S.convert[i].name) + '" title="点击修改武器名称">' +
        '</th>';
    }
    html += '</tr></thead><tbody>';

    for (var f = 0; f < CONVERT_FIELDS.length; f++) {
      var fld = CONVERT_FIELDS[f];
      html += '<tr><th scope="row" class="pldc-rowhead">' + esc(fld.label) + hintHtml(fld) + '</th>';
      for (var c = 0; c < 4; c++) {
        html += '<td class="pldc-cell">' + cellControl('convert', c, fld, S.convert[c]) + '</td>';
      }
      html += '</tr>';
    }

    html += '<tr class="pldc-divider"><th colspan="5" scope="colgroup">换算结果</th></tr>';
    for (var r = 0; r < CONVERT_RESULTS.length; r++) {
      var res = CONVERT_RESULTS[r];
      html += '<tr class="' + (res.strong ? 'pldc-row-strong' : '') + '">' +
        '<th scope="row" class="pldc-rowhead">' + esc(res.label) + '</th>';
      for (var c2 = 0; c2 < 4; c2++) {
        html += '<td class="pldc-cell pldc-out" data-out="convert" data-col="' + c2 + '" data-key="' + res.key + '"></td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table></div>' +
      '<p class="pldc-note">读法：攻击 / 力量收益 = 7.2 表示 1 点攻击力与 7.2 点力量的输出提升相同；' +
      '攻击 / 敏捷收益 = 25.5 表示 1 点攻击与 25.5 点敏捷等价。未计算英雄回声，也未计算压制术与伤害上限。</p>' +
      '</section>';
    return html;
  }

  function toolbarHtml() {
    var cnt = '';
    for (var n = 1; n <= 4; n++) {
      cnt += '<option value="' + n + '"' + (n === S.count ? ' selected' : '') + '>对比 ' + n + ' 把武器</option>';
    }
    return '<label class="pldc-field">' +
      '<span>对比武器数量</span><span class="pldc-hint" data-hint="选择同时对比几把武器（1 ~ 4），下面的卡片与表格列数会跟着变。' +
      '隐藏起来的列数据仍保留在本机，改回 4 把就会恢复。" tabindex="0" role="note">?</span>' +
      '<select class="pldc-input pldc-select" data-global="count">' + cnt + '</select></label>' +
      '<span class="pldc-grow"></span>' +
      '<button type="button" class="pldc-btn" data-act="share">复制分享链接</button>' +
      '<button type="button" class="pldc-btn pldc-btn-ghost" data-act="reset">恢复默认</button>';
  }

  var TABS = [
    { key: 'calc',    label: '环破伤害计算' },
    { key: 'convert', label: '攻力敏换算' },
    { key: 'guide',   label: '使用说明' }
  ];

  function tabsHtml() {
    var out = '';
    for (var i = 0; i < TABS.length; i++) {
      out += '<button type="button" class="pldc-tab" data-tab="' + TABS[i].key + '">' + TABS[i].label + '</button>';
    }
    return out;
  }

  /* ------------------------- 刷新 ------------------------- */

  function resultsOf() {
    var out = [];
    for (var i = 0; i < S.count; i++) out.push(computeColumn(S.columns[i]));
    return out;
  }

  function setText(node, text) { if (node && node.textContent !== text) node.textContent = text; }

  function refreshCalc() {
    var res = resultsOf();
    var best = -Infinity, bestIdx = -1;
    for (var i = 0; i < res.length; i++) {
      if (res[i].dpm > best) { best = res[i].dpm; bestIdx = i; }
    }

    var cells = root.querySelectorAll('[data-out="calc"]');
    for (var n = 0; n < cells.length; n++) {
      var cell = cells[n];
      var col = parseInt(cell.getAttribute('data-col'), 10);
      var key = cell.getAttribute('data-key');
      var spec = null;
      for (var q = 0; q < RESULTS.length; q++) if (RESULTS[q].key === key) spec = RESULTS[q];
      if (!spec) continue;
      setText(cell, fmtNum(res[col][key], spec.dec));
      cell.classList.toggle('pldc-best', key === 'dpm' && col === bestIdx && best > 0);
    }

    var cards = [];
    for (var c = 0; c < res.length; c++) {
      var r = res[c];
      var isBest = c === bestIdx && best > 0;
      var delta = best > 0 && r.dpm > 0 && r.dpm < best ? '−' + fmtNum((1 - r.dpm / best) * 100, 1) + '%' : '';
      cards.push('<div class="pldc-card' + (isBest ? ' pldc-card-best' : '') + '">' +
        '<div class="pldc-card-top"><span class="pldc-card-name">' + esc(S.columns[c].name) + '</span>' +
        (isBest ? '<span class="pldc-pill">最高</span>' : (delta ? '<span class="pldc-pill pldc-pill-dim">' + delta + '</span>' : '')) +
        '</div>' +
        '<div class="pldc-card-main">' + fmtNum(r.dpm, 2) + '<small>万 / 分钟</small></div>' +
        '<div class="pldc-card-sub">单次 ' + fmtNum(r.hit, 2) + ' · 面板 ' + fmtNum(r.panel, 2) + '</div>' +
        '</div>');
    }
    var cardBox = root.querySelector('.pldc-cards');
    if (cardBox) cardBox.innerHTML = cards.join('');

    var warnBox = root.querySelector('[data-warn]');
    if (warnBox) {
      var msgs = [];
      for (var w = 0; w < res.length; w++) {
        if (Math.abs(res[w].ratioSum - 100) > 0.001) {
          msgs.push(esc(S.columns[w].name) + ' 的优势 ' + fmtNum(S.columns[w].advRatio, 1) + '% + 劣势 ' +
            fmtNum(S.columns[w].disRatio, 1) + '% = ' + fmtNum(res[w].ratioSum, 1) + '%');
        }
      }
      warnBox.innerHTML = msgs.length ? '注意：' + msgs.join('；') + '，合计不等于 100%，结果可能与预期不符。' : '';
      warnBox.hidden = !msgs.length;
    }
  }

  function refreshConvert() {
    var res = [];
    for (var i = 0; i < 4; i++) res.push(computeConvert(S.convert[i]));
    var cells = root.querySelectorAll('[data-out="convert"]');
    for (var n = 0; n < cells.length; n++) {
      var cell = cells[n];
      var col = parseInt(cell.getAttribute('data-col'), 10);
      var key = cell.getAttribute('data-key');
      var spec = null;
      for (var q = 0; q < CONVERT_RESULTS.length; q++) if (CONVERT_RESULTS[q].key === key) spec = CONVERT_RESULTS[q];
      if (spec) setText(cell, fmtNum(res[col][key], spec.dec));
    }
  }

  function refreshAll() {
    refreshCalc();
    refreshConvert();
  }

  /* ------------------------- 交互 ------------------------- */

  function slotsOf(mode) { return mode === 'convert' ? S.convert : S.columns; }

  function writeSlot(mode, col, key, raw) {
    var slot = slotsOf(mode)[col];
    if (key === 'name') {
      var fallback = (mode === 'convert' ? CONVERT_BASE[col] : BASE[col]).name;
      slot.name = String(raw).slice(0, 12) || fallback;
      return;
    }
    if (key === 'relation') { slot.relation = pickRelation(raw, slot.relation); return; }
    if (key === 'crit') { slot.crit = pickText(raw, CRIT_OPTIONS, slot.crit); return; }
    if (key === 'echo') { slot.echo = pickText(raw, ECHO_OPTIONS, slot.echo); return; }
    if (key === 'spd') { slot.spd = pickSpd(raw, slot.spd); return; }
    if (key === 'weapon') {
      var w = weaponBy(pickWeapon(raw, slot.weapon));
      slot.weapon = w.text;
      slot.advCoef = w.advCoef;
      slot.disCoef = w.disCoef;
      slot.mastery = w.mastery;
      slot.advRatio = w.advRatio;
      slot.disRatio = w.disRatio;
      return;
    }
    slot[key] = num(raw, 0);
  }

  function syncDerived(mode, col) {
    var slot = slotsOf(mode)[col];
    var keys = ['advCoef', 'disCoef', 'mastery', 'advRatio', 'disRatio'];
    for (var i = 0; i < keys.length; i++) {
      var input = root.querySelector('[data-mode="' + mode + '"][data-col="' + col + '"][data-key="' + keys[i] + '"]');
      if (input) input.value = String(slot[keys[i]]);
    }
  }

  function onInput(e) {
    var t = e.target;
    var g = t.getAttribute && t.getAttribute('data-global');
    if (g) {
      if (g === 'count') {
        var next = pickCount(t.value, S.count);
        if (next === S.count) return;
        S.count = next;
        rebuild();
        setTab('calc');
        saveState();
        return;
      }
      refreshAll();
      saveState();
      return;
    }
    var mode = t.getAttribute && t.getAttribute('data-mode');
    var col = t.getAttribute && t.getAttribute('data-col');
    var key = t.getAttribute && t.getAttribute('data-key');
    if (!mode || col === null || !key) return;
    writeSlot(mode, parseInt(col, 10), key, t.value);
    if (key === 'weapon') syncDerived(mode, parseInt(col, 10));
    if (mode === 'convert') refreshConvert(); else refreshCalc();
    saveState();
  }

  function onBlur(e) {
    var t = e.target;
    if (!t.getAttribute || !t.getAttribute('data-key')) return;
    var key = t.getAttribute('data-key');
    if (key === 'name' || t.tagName === 'SELECT') return;
    var mode = t.getAttribute('data-mode');
    var col = parseInt(t.getAttribute('data-col'), 10);
    if (!mode || isNaN(col)) return;
    var val = slotsOf(mode)[col][key];
    if (t.value !== String(val)) t.value = String(val);
  }

  /* --- 说明气泡 --- */

  var tipEl = null, tipTimer = null;

  function ensureTip() {
    if (!tipEl) {
      tipEl = document.createElement('div');
      tipEl.className = 'pldc-tooltip';
      document.body.appendChild(tipEl);
    }
    return tipEl;
  }

  function placeTip(target, el) {
    var r = target.getBoundingClientRect();
    var w = el.offsetWidth, h = el.offsetHeight;
    var left = r.left + r.width / 2 - w / 2;
    left = Math.max(8, Math.min(left, (window.innerWidth || 0) - w - 8));
    var top = r.top - h - 10;
    if (top < 8) top = r.bottom + 10;
    el.style.left = Math.round(left) + 'px';
    el.style.top = Math.round(top) + 'px';
  }

  function showTip(target) {
    var text = target.getAttribute('data-hint');
    if (!text) return;
    var el = ensureTip();
    el.textContent = text;
    el.classList.add('is-on');
    placeTip(target, el);
    if (tipTimer) clearTimeout(tipTimer);
    tipTimer = setTimeout(hideTip, 9000);
  }

  function hideTip() {
    if (tipEl) tipEl.classList.remove('is-on');
    if (tipTimer) { clearTimeout(tipTimer); tipTimer = null; }
  }

  function closestHint(node) {
    while (node && node !== document) {
      if (node.getAttribute && node.getAttribute('data-hint')) return node;
      node = node.parentNode;
    }
    return null;
  }

  var toastTimer = null;
  function toast(msg) {
    var box = document.getElementById('pldc-toast');
    if (!box) {
      box = document.createElement('div');
      box.id = 'pldc-toast';
      box.className = 'pldc-toast';
      document.body.appendChild(box);
    }
    box.textContent = msg;
    box.classList.add('is-on');
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { box.classList.remove('is-on'); }, 2200);
  }

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { toast('短链接已复制（' + text.length + ' 字符）'); }, function () { toast('复制失败，请手动复制地址栏'); });
      return;
    }
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', 'readonly');
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); toast('短链接已复制（' + text.length + ' 字符）'); } catch (e) { toast('复制失败，请手动复制地址栏'); }
    document.body.removeChild(ta);
  }

  function setTab(key) {
    var tabs = root.querySelectorAll('[data-tab]');
    for (var i = 0; i < tabs.length; i++) {
      tabs[i].classList.toggle('is-active', tabs[i].getAttribute('data-tab') === key);
    }
    var panels = root.querySelectorAll('[data-panel]');
    for (var p = 0; p < panels.length; p++) {
      panels[p].hidden = panels[p].getAttribute('data-panel') !== key;
    }
    var bar = document.getElementById('pldc-toolbar');
    if (bar) bar.hidden = key !== 'calc';
  }

  function onClick(e) {
    var btn = e.target.closest ? e.target.closest('[data-act],[data-tab]') : null;
    if (!btn) return;

    var tab = btn.getAttribute('data-tab');
    if (tab) { setTab(tab); return; }

    var act = btn.getAttribute('data-act');
    if (act === 'share') {
      var hash = '#d=' + encodeState(S);
      var url = location.origin + location.pathname + hash;
      if (history.replaceState) history.replaceState(null, '', hash);
      copyText(url);
      return;
    }
    if (act === 'reset') {
      if (!confirm('恢复默认数据？当前填写的所有参数会被清除。')) return;
      S = normalize(null);
      try { localStorage.removeItem(STORAGE_KEY); } catch (err) { /* 忽略 */ }
      rebuild();
      toast('已恢复默认');
      return;
    }
  }

  function build() {
    var bar = document.getElementById('pldc-toolbar');
    if (bar) bar.innerHTML = toolbarHtml();
    var tabs = document.getElementById('pldc-tabs');
    if (tabs) tabs.innerHTML = tabsHtml();
    var panels = document.getElementById('pldc-panels');
    if (panels) panels.innerHTML = calcPanel() + convertPanel();
  }

  function rebuild() {
    build();
    refreshAll();
  }

  var saveTimer = null;
  function saveState() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(S)); } catch (e) { /* 忽略配额错误 */ }
    }, 300);
  }

  function readHashState() {
    var h = location.hash || '';
    var i = h.indexOf('d=');
    if (i < 0) return null;
    try { return normalize(decodeState(h.slice(i + 2))); } catch (e) { return null; }
  }

  function loadState() {
    var fromHash = readHashState();
    if (fromHash) return fromHash;
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return normalize(JSON.parse(raw));
    } catch (e) { /* localStorage 不可用则忽略 */ }
    return normalize(null);
  }

  // 本页（且仅本页）把文章信息里的「发表于」改成「更新于」
  function relabelPostDate() {
    var labels = document.querySelectorAll('#post-meta .post-meta-label');
    for (var i = 0; i < labels.length; i++) {
      if (/发表于/.test(labels[i].textContent)) {
        labels[i].textContent = labels[i].textContent.replace('发表于', '更新于');
      }
    }
  }

  function init() {
    root = document.getElementById('pldc');
    if (!root) return;
    relabelPostDate();
    S = loadState();
    build();
    root.addEventListener('input', onInput);
    root.addEventListener('change', onInput);
    root.addEventListener('blur', onBlur, true);
    root.addEventListener('click', onClick);

    // 悬停 / 触摸显示参数说明
    root.addEventListener('pointerover', function (e) {
      var h = closestHint(e.target);
      if (h) showTip(h);
    });
    root.addEventListener('pointerout', function (e) {
      var h = closestHint(e.target);
      if (h && !(h.contains && h.contains(e.relatedTarget))) hideTip();
    });
    root.addEventListener('focusin', function (e) {
      var h = closestHint(e.target);
      if (h) showTip(h);
    });
    root.addEventListener('focusout', hideTip);
    root.addEventListener('pointerdown', function (e) {
      if (e.pointerType !== 'touch') return;
      var h = closestHint(e.target);
      if (h) { e.preventDefault(); showTip(h); } else hideTip();
    });
    document.addEventListener('pointerdown', function (e) {
      if (!closestHint(e.target) && e.pointerType !== 'touch') hideTip();
    }, true);
    window.addEventListener('scroll', hideTip, true);
    window.addEventListener('resize', hideTip);

    refreshAll();
    setTab('calc');
    root.setAttribute('data-ready', '1');
  }

  /* ------------------------- 导出 ------------------------- */

  var API = {
    computeColumn: computeColumn,
    computeConvert: computeConvert,
    normalize: normalize,
    encodeState: encodeState,
    decodeState: decodeState,
    DEFAULTS: DEFAULTS,
    BASE: BASE,
    CONVERT_BASE: CONVERT_BASE,
    WEAPONS: WEAPONS,
    SPEED_TABLE: SPEED_TABLE,
    CRIT_OPTIONS: CRIT_OPTIONS,
    ECHO_OPTIONS: ECHO_OPTIONS,
    RELATION_OPTIONS: RELATION_OPTIONS
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  if (typeof window !== 'undefined') window.PLDC = API;

  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
  }
})();
